// Method:  FRG_MCE_ImportWorkbook
// Type:    server Method, C#
// Purpose: the server side of importing the workbook of a Mass Create & Edit job. The
//          browser reads the workbook, shows the preview and drives the import, one
//          request per batch, so importing takes several calls:
//
//   step = upload    attaches the uploaded workbook and its validation result (two File
//                    items the browser stored) to the job and moves the job to Validated.
//                    The Files of an earlier upload that it replaces are deleted (see
//                    ReplacedUploadFiles); the job Form asks the user first.
//   step = begin     checks that the import may start, and says how big a batch may be and
//                    how many batches the browser may send at once.
//   step = files     adds File items the browser stored in the vault for the import (the
//                    files that file cells name) to the job's list of stored files
//                    (staged_files). Each must be a File the caller stored, with the
//                    checksum the vault set. The browser stores a batch's files before
//                    the batch.
//   step = batch     checks the rows of one batch again and writes them, as the caller.
//                    Batches that don't depend on each other may run at the same time (see
//                    "Batches at the same time" below).
//   step = complete  records the rows that failed, sets the counts and moves the job to
//                    Imported. The stored Files that no row linked are deleted.
//
// Input (properties of the context item):
//   job_id   id of the job
//   step     "upload", "begin", "files", "batch" or "complete"
//   token    (upload, begin, files, complete) the workbook token the workbook carries in its meta
//            (meta.job.token). The job accepts only the workbook it generated last.
//   upload_file_id   (upload) id of the File item that holds the uploaded .xlsx
//   results_file_id  (upload) id of the File item that holds the validation result: the
//                    JSON FRG_MCE_ValidateImport answered for that workbook
//   workbook (batch) a slice of the workbook, in the shape FRG_MCE_ValidateImport reads:
//            { meta, sheets: [ { sheetName, rows: [ { rowNumber, cells, modifiedOn, digest } ] } ] }
//            meta is the meta of the workbook as it was read. sheets lists every sheet of
//            the workbook in its order, each with the rows of this batch only (none is an
//            empty array). Each row is a row as the workbook was read, plus digest: the
//            digest the validation result has for that row, which is what the user reviewed.
//            A job with several targets has a main sheet for each, first (see
//            FRG_MCE_ValidateImport): "main-sheet row" below means a row of any of them, and the
//            parent of a relationship row is on the main sheet its sheet names.
//            What a batch holds:
//              - up to batchSize (see begin) groups. A group is a main-sheet row to write
//                with its relationship rows, or the relationship rows of one existing
//                parent whose main-sheet row is not written (it has no row, or its row is
//                unchanged). A parent with all its relationship rows is always in one batch.
//              - a relationship row whose related item a new main-sheet row creates goes in
//                the batch of that row, or in a later one. In a later batch the item exists:
//                the digest still names it by its row, through the job's Created records.
//                New items that point at each other (two alternates of each other) go in one
//                batch.
//              - rows to write only: no error rows and no unchanged rows. An unchanged
//                main-sheet row may still be sent with its relationship rows; it is not
//                written itself.
//   failed   (complete) JSON, the rows that were not imported, for the job's records:
//            [ { "sheetName": "...", "sheet": 0, "row": 7, "key": "text that names the row", "message": "..." } ]
//            sheet, the index of the sheet in the workbook, may be left out.
//   files    (batch) JSON, the files the user picked, as FRG_MCE_ValidateImport takes them.
//            The batch hands them on to the check, so file cells resolve as in the preview.
//   file_ids (files) JSON, the ids of File items the browser has just stored in the vault for
//            the import: [ "<id>", ... ]
//   counts   (complete) JSON, what the browser counted over the whole workbook:
//            { "total": rows on all sheets, "unchanged": rows with nothing to change,
//              "skipped": rows not sent because they have an error }
//
// Output: a result whose text is JSON. job is { id, number, state } after the step.
//   upload:   { job, uploadFileId, resultsFileId }
//   begin:    { job, batchSize, parallelBatches, uploadFileId, resultsFileId,
//               records: { created, updated, deleted, failed }, files: [ { id, md5, name, linked } ] }
//             parallelBatches is how many batches the browser may send at once (1 to 8).
//             records counts the records the job has already: an import that was broken
//             off has written some rows, and running it again skips them (see batch).
//             files lists the Files stored for the import so far (see files), so that a run
//             that was broken off does not store them again.
//   files:    { job, files: [ { id, md5, name, linked } ] }
//             files is the job's whole list of stored Files, the new ones included.
//   batch:    { job, rows: [ ... ], counts: { created, updated, deleted, unchanged, failed } }
//             one entry per row of the slice, sheet by sheet in the order of the slice:
//             { sheet, row, status, id?, keyedName?, reason?, message? }
//               sheet   the index of the sheet in the workbook (the main sheets come first)
//               row     the row number in the Excel sheet
//               status  "created", "updated" or "deleted": the row was written. id and keyedName
//                       name the item as it is now; for a row of a relationship sheet
//                       that is its parent (the parent may have a new generation).
//                       "unchanged": the row changes nothing and was not written.
//                       "failed": the row was not written; reason and message say why.
//               reason  "changed": the row's digest is not the one sent. Aras changed since
//                         the preview, or the row was imported already. Upload the
//                         workbook again to review it.
//                       "error": the row has an error now; message holds the messages.
//                       "parent": the row's parent row (on its main sheet) failed.
//                       "related": the main-sheet row that creates the row's related item
//                         is not written in this batch.
//                       "file": a file the row links was not stored for the import, or the
//                         stored File is not the file the preview showed.
//   complete: { job, counts: { total, created, updated, deleted, unchanged, skipped, failed } }
//             created, updated and deleted count the job's Created, Updated and Deleted
//             records, so they include rows an earlier, broken-off run of the import wrote.
//
// Errors. A step that refuses the call answers with an ordinary error (fault code 1, the
// reason in faultstring) and has written nothing. A batch whose write fails answers with
// a fault that names the rows of the failed write:
//   faultcode    MCE_ROW_FAILED
//   faultstring  "Row 7 of sheet 'Part' could not be imported: <the message of Aras>"
//   detail       JSON: { "sheet": 0, "sheetName": "Part", "row": 7, "message": "<the message of Aras>",
//                        "rows": [ { "sheet": 0, "row": 7 }, { "sheet": 1, "row": 3 }, ... ] }
//                rows lists every row of the failed write, the row named first: a main-sheet
//                row with the relationship rows nested in its write, or the relationship rows
//                of one existing parent. Any of them may be the one that could not be written.
// Nothing of a batch that fails is kept: the request is one transaction, and the error
// rolls it back, records included. Send the batch again without the rows of the failed
// write; those rows go into the failed list of complete.
// A batch that the database chose as the victim of a deadlock with another request (another
// batch, say) answers with faultcode MCE_RETRY instead, naming no row: nothing of it is kept,
// and it can be sent again as it was.
//
// Batches at the same time. The browser may send batches that don't depend on each other at
// the same time, up to the number begin gives. A batch depends on an earlier one when one of
// its rows points at an item that the earlier batch creates; it is sent once that batch has
// finished. What batches share is written so that they don't get in each other's way: each
// record gets its own sort order (sheet and row), and the job's list of stored Files is read
// again and changed while the job is held for the change (UpdateStaged).
//
// What batch does, in this one request:
//   1. checks the job as begin does, and the size of the batch;
//   2. checks the slice with FRG_MCE_ValidateImport, which must answer "validated";
//   3. writes each row whose action is new or update and whose digest is the one sent.
//      A main-sheet row is one write: add for a new item, edit for an existing one, with
//      its relationship rows nested in it. The relationship rows of a parent whose row is
//      not written are an edit of that parent with only those rows nested. An item that
//      one write of the batch gave a new generation is named by that generation in the
//      later writes. Each new item gets its id before the first write, so that
//      relationship rows can point at it; the new items that relationship rows of the
//      batch point at are written first.
//      A file cell links the File stored for the import with the MD5 and the name the cell
//      was checked with; a row with the delete marker deletes the relationship row it matches.
//   4. adds a Created, Updated or Deleted record for each row it wrote, and marks the stored
//      Files it linked. A record's sort order is its sheet and row (RecordOrder), so records
//      list in the order of the workbook.
//
// All steps refuse the call unless
//   - the caller created the job (only its creator runs a job),
//   - the job is Generated or Validated (upload) or Validated (the other steps),
//     and an upload was attached to it (all steps but upload),
//   - the job is not locked (it is not being edited),
//   - every target ItemType of the job is on the allow-list of the configuration item, for an
//     identity of the caller (or for everyone),
//   - the token is the job's current workbook token.
//
// The items of the workbook are read and written with the caller's own permissions. The
// job's own properties, its state and its records are changed with the identity
// "FRG Mass Create & Edit Steps", which only the Methods of Mass Create & Edit take on,
// for those writes only: nothing else may change them. The Files of an upload that a new
// upload replaces, and the stored Files that no row linked, are deleted with the Super User
// identity, for that delete only. The
// Method runs no SQL.

Innovator inn = this.getInnovator();

const string JobType = "FRG_MassCreateEdit";
const string RecordType = "FRG_MassCreateEdit_Record";
// The identity "FRG Mass Create & Edit Steps".
const string StepsIdentityId = "B56E4D1B3F4DDC99AECEC5AFB250BF63";
// Taken on only to delete the Files of an upload that a new upload replaced, and the Files
// stored for an import that no row linked (DeleteReplacedFiles).
const string SuperUserIdentityId = "6B14D33C4A7D41C188CCF2BC15BD01A3";
const string RowFailedCode = "MCE_ROW_FAILED";
const string RetryCode = "MCE_RETRY";
const int RecordChunk = 200;
const int MaxJsonDepth = 16;

Item rowFailure = null;

try
{
	string step = Input("step");
	if (step == "upload")
	{
		return inn.newResult(Upload());
	}
	if (step == "begin")
	{
		return inn.newResult(Begin());
	}
	if (step == "batch")
	{
		string answer = Batch();
		return rowFailure ?? inn.newResult(answer);
	}
	if (step == "files")
	{
		return inn.newResult(StoreFiles());
	}
	if (step == "complete")
	{
		return inn.newResult(Complete());
	}
	throw new ApplicationException("FRG_MCE_ImportWorkbook: the 'step' property must be 'upload', 'begin', 'files', 'batch' or 'complete'.");
}
catch (ApplicationException ex)
{
	return IsDeadlock(ex.Message) ? RetryError(ex.Message) : inn.newError(ex.Message);
}
catch (Exception ex) when (IsDeadlock(ex.Message))
{
	return RetryError(ex.Message);
}

// ---------------------------------------------------------------------------------
// The four steps
// ---------------------------------------------------------------------------------

string Upload()
{
	XmlElement job = ReadJob(new string[] { "Generated", "Validated" });
	string jobId = job.GetAttribute("id");
	CheckAllowed(job, ReadConfiguration());
	CheckToken(job, Input("token"));
	string uploadId = FileOfCaller("upload_file_id", "uploaded workbook");
	string resultsId = FileOfCaller("results_file_id", "validation result");
	if (uploadId == resultsId)
	{
		throw new ApplicationException("FRG_MCE_ImportWorkbook: the workbook and the validation result must be two File items.");
	}
	List<string> replaced = ReplacedUploadFiles(job, uploadId, resultsId);
	AsSteps(() =>
	{
		Apply(
			"<Item type=\"" + JobType + "\" action=\"edit\" id=\"" + jobId + "\">" +
			"<upload_file>" + uploadId + "</upload_file>" +
			"<results_file>" + resultsId + "</results_file>" +
			"</Item>");
		if (Text(job, "state") == "Generated")
		{
			Promote(jobId, "Validated");
		}
	});
	DeleteReplacedFiles(replaced);

	Newtonsoft.Json.Linq.JObject answer = new Newtonsoft.Json.Linq.JObject();
	answer.Add("job", JobJson(ReadJob(null)));
	answer.Add("uploadFileId", uploadId);
	answer.Add("resultsFileId", resultsId);
	return answer.ToString(Newtonsoft.Json.Formatting.None);
}

string Begin()
{
	XmlElement job = ReadJob(new string[] { "Validated" });
	XmlElement config = ReadConfiguration();
	CheckAllowed(job, config);
	CheckToken(job, Input("token"));
	CheckUploaded(job);

	Dictionary<string, int> records = RecordCounts(job.GetAttribute("id"));
	Newtonsoft.Json.Linq.JObject recordJson = new Newtonsoft.Json.Linq.JObject();
	recordJson.Add("created", records["Created"]);
	recordJson.Add("updated", records["Updated"]);
	recordJson.Add("deleted", records["Deleted"]);
	recordJson.Add("failed", records["Failed"]);

	Newtonsoft.Json.Linq.JObject answer = new Newtonsoft.Json.Linq.JObject();
	answer.Add("job", JobJson(job));
	answer.Add("batchSize", BatchSize(config));
	answer.Add("parallelBatches", ParallelBatches(config));
	answer.Add("uploadFileId", Text(job, "upload_file"));
	answer.Add("resultsFileId", Text(job, "results_file"));
	answer.Add("records", recordJson);
	answer.Add("files", ReadStaged(job));
	return answer.ToString(Newtonsoft.Json.Formatting.None);
}

// Adds Files the browser stored for the import to the job's list. The browser stores them in
// the vault as the caller, as it stores the workbook, and only then can the job name them.
string StoreFiles()
{
	XmlElement job = ReadJob(new string[] { "Validated" });
	string jobId = job.GetAttribute("id");
	CheckAllowed(job, ReadConfiguration());
	CheckToken(job, Input("token"));
	CheckUploaded(job);

	List<string> ids = new List<string>();
	try
	{
		foreach (Newtonsoft.Json.Linq.JToken id in (Newtonsoft.Json.Linq.JArray)ParseJson(Input("file_ids")))
		{
			if (!IsId((string)id))
			{
				throw new FormatException("id");
			}
			ids.Add(((string)id).ToUpperInvariant());
		}
	}
	catch (Exception)
	{
		throw new ApplicationException("FRG_MCE_ImportWorkbook: 'file_ids' must be a JSON list of File ids.");
	}
	ids = ids.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
	List<Newtonsoft.Json.Linq.JObject> added = new List<Newtonsoft.Json.Linq.JObject>();
	if (ids.Count > 0)
	{
		List<XmlElement> files = Query("<Item type=\"File\" action=\"get\" select=\"created_by_id,filename,checksum\" idlist=\"" + string.Join(",", ids) + "\"/>");
		foreach (string id in ids)
		{
			XmlElement file = files.FirstOrDefault(f => string.Equals(f.GetAttribute("id"), id, StringComparison.OrdinalIgnoreCase));
			if (file == null || Text(file, "created_by_id") != inn.getUserID())
			{
				throw new ApplicationException("A stored file was not found, or was not stored by you (" + id + ").");
			}
			if (!System.Text.RegularExpressions.Regex.IsMatch(Text(file, "checksum"), "^[0-9A-Fa-f]{32}$"))
			{
				throw new ApplicationException("The file '" + Text(file, "filename") + "' has no checksum. Store it again.");
			}
			Newtonsoft.Json.Linq.JObject entry = new Newtonsoft.Json.Linq.JObject();
			entry.Add("id", file.GetAttribute("id"));
			entry.Add("md5", Text(file, "checksum").ToUpperInvariant());
			entry.Add("name", Text(file, "filename"));
			entry.Add("linked", false);
			added.Add(entry);
		}
	}
	// Batches of the import may be adding Files and marking them at the same time: the list is
	// read again and changed while the job is held (UpdateStaged).
	Newtonsoft.Json.Linq.JArray staged = UpdateStaged(jobId, list =>
	{
		HashSet<string> known = new HashSet<string>(list.Select(entry => (string)entry["id"]), StringComparer.OrdinalIgnoreCase);
		bool changed = false;
		foreach (Newtonsoft.Json.Linq.JObject entry in added.Where(entry => !known.Contains((string)entry["id"])))
		{
			list.Add(entry);
			changed = true;
		}
		return changed;
	});
	Newtonsoft.Json.Linq.JObject answer = new Newtonsoft.Json.Linq.JObject();
	answer.Add("job", JobJson(job));
	answer.Add("files", staged);
	return answer.ToString(Newtonsoft.Json.Formatting.None);
}

string Batch()
{
	XmlElement job = ReadJob(new string[] { "Validated" });
	string jobId = job.GetAttribute("id");
	XmlElement config = ReadConfiguration();
	CheckAllowed(job, config);
	CheckUploaded(job);

	// The slice, and the digest sent with each of its rows.
	string workbookText = this.getProperty("workbook", "") ?? "";
	Newtonsoft.Json.Linq.JObject slice;
	try
	{
		slice = ParseJson(workbookText) as Newtonsoft.Json.Linq.JObject;
	}
	catch (Exception)
	{
		slice = null;
	}
	Newtonsoft.Json.Linq.JArray sliceSheets = slice == null ? null : slice["sheets"] as Newtonsoft.Json.Linq.JArray;
	if (sliceSheets == null || !(slice["meta"] is Newtonsoft.Json.Linq.JObject))
	{
		throw new ApplicationException("The batch could not be read: 'workbook' must be a slice of the workbook as JSON.");
	}
	Newtonsoft.Json.Linq.JObject metaJob = slice["meta"]["job"] as Newtonsoft.Json.Linq.JObject;
	CheckToken(job, metaJob == null ? "" : Str(metaJob, "token") ?? "");
	Dictionary<string, string> sentDigests = new Dictionary<string, string>(StringComparer.Ordinal);
	for (int sheet = 0; sheet < sliceSheets.Count; sheet++)
	{
		foreach (Newtonsoft.Json.Linq.JToken row in sliceSheets[sheet]["rows"] as Newtonsoft.Json.Linq.JArray ?? new Newtonsoft.Json.Linq.JArray())
		{
			string rowNumber = Str(row, "rowNumber");
			if (rowNumber == null || sentDigests.ContainsKey(sheet + ":" + rowNumber))
			{
				throw new ApplicationException("The batch could not be read: every row needs a row number, and a row can be in a batch once.");
			}
			sentDigests[sheet + ":" + rowNumber] = Str(row, "digest") ?? "";
		}
	}

	// The rows as they are now.
	Item validation = inn.applyMethod("FRG_MCE_ValidateImport",
		"<job_id>" + jobId + "</job_id><workbook>" + Xml(workbookText) + "</workbook><files>" + Xml(this.getProperty("files", "") ?? "") + "</files>");
	if (validation.isError())
	{
		throw new ApplicationException("The batch could not be checked: " + validation.getErrorString());
	}
	Newtonsoft.Json.Linq.JObject result = ParseJson(validation.getResult()) as Newtonsoft.Json.Linq.JObject;
	if (result == null || (string)result["status"] != "validated")
	{
		string reasons = result == null ? "" : string.Join(" ", ((Newtonsoft.Json.Linq.JArray)result["issues"])
			.Where(i => (string)i["severity"] == "error").Select(i => (string)i["message"]));
		throw new ApplicationException("The batch was refused: " + (reasons.Length > 0 ? reasons : "the workbook could not be checked."));
	}
	Newtonsoft.Json.Linq.JArray sheets = (Newtonsoft.Json.Linq.JArray)result["sheets"];
	if (sheets.Count != sliceSheets.Count)
	{
		throw new ApplicationException("The batch could not be read: it does not list every sheet of the workbook.");
	}

	// The Files stored for the import, by the file token a cell names them with
	// (file:<MD5>:<name>). Each is read again: its checksum, which the vault set, and its name
	// must be the ones of the token, so a row links only the file the preview showed.
	Newtonsoft.Json.Linq.JArray staged = ReadStaged(job);
	Dictionary<string, string> storedFiles = StoredFileIds(staged);

	// The main sheets come first, one for each target. A relationship sheet names the main
	// sheet of its parents (parentSheet), and each of its rows names the main-sheet row of its
	// parent and of a related item that a new row creates, by sheet and row.
	int mainCount = 0;
	while (mainCount < sheets.Count && (string)sheets[mainCount]["kind"] == "main")
	{
		mainCount++;
	}
	if (mainCount == 0)
	{
		throw new ApplicationException("The batch could not be read: the workbook has no main sheet.");
	}

	// What happens to each row: written, unchanged or failed. A relationship row whose
	// parent row fails fails with it.
	Dictionary<string, Newtonsoft.Json.Linq.JObject> outcome = new Dictionary<string, Newtonsoft.Json.Linq.JObject>(StringComparer.Ordinal);
	Newtonsoft.Json.Linq.JArray rowsOut = new Newtonsoft.Json.Linq.JArray();
	for (int sheet = 0; sheet < sheets.Count; sheet++)
	{
		bool isMain = sheet < mainCount;
		foreach (Newtonsoft.Json.Linq.JObject row in sheets[sheet]["rows"])
		{
			row["sheetIndex"] = sheet;
			string key = sheet + ":" + (int)row["row"];
			string action = (string)row["action"];
			string sent;
			sentDigests.TryGetValue(key, out sent);
			Newtonsoft.Json.Linq.JObject entry = new Newtonsoft.Json.Linq.JObject();
			entry.Add("sheet", sheet);
			entry.Add("row", (int)row["row"]);
			if (action == "error")
			{
				Failed(entry, "error", "The row has an error now: " + string.Join(" ",
					((Newtonsoft.Json.Linq.JArray)row["issues"]).Where(i => (string)i["severity"] == "error").Select(i => (string)i["message"])));
			}
			else if (string.IsNullOrEmpty(sent) || sent != (string)row["digest"])
			{
				Failed(entry, "changed", "The row is not what the preview showed: the data in Aras changed since, or the row was imported already. Upload the workbook again to review it.");
			}
			else if (action == "none")
			{
				entry.Add("status", "unchanged");
			}
			else
			{
				entry.Add("status", "write");
			}
			Newtonsoft.Json.Linq.JObject parent = row["parent"] as Newtonsoft.Json.Linq.JObject;
			if (!isMain && (string)entry["status"] != "failed" && parent != null && parent["row"] != null && parent["row"].Type == Newtonsoft.Json.Linq.JTokenType.Integer)
			{
				int parentSheet = ParentSheetOf(sheets, sheet);
				Newtonsoft.Json.Linq.JObject parentOutcome;
				if (outcome.TryGetValue(parentSheet + ":" + (int)parent["row"], out parentOutcome) && (string)parentOutcome["status"] == "failed")
				{
					entry.Remove("status");
					Failed(entry, "parent", "Its parent, row " + (int)parent["row"] + " of sheet '" + (string)sheets[parentSheet]["sheetName"] + "', was not imported.");
				}
			}
			// A related item that a main-sheet row of the batch creates exists only once that row is written.
			int[] relatedRow = NewRelatedRow(row);
			if (!isMain && (string)entry["status"] != "failed" && relatedRow != null)
			{
				Newtonsoft.Json.Linq.JObject relatedOutcome;
				if (!outcome.TryGetValue(relatedRow[0] + ":" + relatedRow[1], out relatedOutcome) || (string)relatedOutcome["status"] != "write")
				{
					entry.Remove("status");
					Failed(entry, "related", "Its related item, row " + relatedRow[1] + " of sheet '" + (string)sheets[relatedRow[0]]["sheetName"] + "', was not imported.");
				}
			}
			string missingFile = (string)entry["status"] == "write" ? FileTokensOf(row).FirstOrDefault(token => !storedFiles.ContainsKey(token)) : null;
			if (missingFile != null)
			{
				entry.Remove("status");
				Failed(entry, "file", "The file '" + TokenName(missingFile) + "' was not stored for the import, or is not the file the preview showed. Import again to store it.");
			}
			outcome[key] = entry;
			rowsOut.Add(entry);
		}
	}

	// The writes, each a group: a main-sheet row with the relationship rows nested in it, or
	// the relationship rows of one existing parent. A group's key is the name of its main-sheet
	// row (NewRowName) or "item:<id>" of its parent; groupSheet holds the main sheet of each.
	List<Tuple<string, Newtonsoft.Json.Linq.JObject, string>> order = new List<Tuple<string, Newtonsoft.Json.Linq.JObject, string>>();
	Dictionary<string, List<Newtonsoft.Json.Linq.JObject>> nested = new Dictionary<string, List<Newtonsoft.Json.Linq.JObject>>(StringComparer.Ordinal);
	Dictionary<string, int> groupSheet = new Dictionary<string, int>(StringComparer.Ordinal);
	for (int sheet = 0; sheet < mainCount; sheet++)
	{
		foreach (Newtonsoft.Json.Linq.JObject row in sheets[sheet]["rows"])
		{
			if ((string)outcome[sheet + ":" + (int)row["row"]]["status"] == "write")
			{
				string groupKey = NewRowName(sheet, (int)row["row"]);
				order.Add(Tuple.Create(groupKey, row, (string)null));
				nested[groupKey] = new List<Newtonsoft.Json.Linq.JObject>();
				groupSheet[groupKey] = sheet;
			}
		}
	}
	for (int sheet = mainCount; sheet < sheets.Count; sheet++)
	{
		foreach (Newtonsoft.Json.Linq.JObject row in sheets[sheet]["rows"])
		{
			if ((string)outcome[sheet + ":" + (int)row["row"]]["status"] != "write")
			{
				continue;
			}
			Newtonsoft.Json.Linq.JObject parent = (Newtonsoft.Json.Linq.JObject)row["parent"];
			int parentSheet = ParentSheetOf(sheets, sheet);
			string rowGroup = parent["row"] != null && parent["row"].Type == Newtonsoft.Json.Linq.JTokenType.Integer ? NewRowName(parentSheet, (int)parent["row"]) : null;
			string groupKey = rowGroup != null && nested.ContainsKey(rowGroup) ? rowGroup : "item:" + (string)parent["id"];
			if (!nested.ContainsKey(groupKey))
			{
				if (parent["id"] == null || parent["id"].Type != Newtonsoft.Json.Linq.JTokenType.String)
				{
					throw new ApplicationException("Row " + (int)row["row"] + " of sheet '" + (string)sheets[sheet]["sheetName"] + "' has no parent to be written with.");
				}
				order.Add(Tuple.Create(groupKey, (Newtonsoft.Json.Linq.JObject)null, (string)parent["id"]));
				nested[groupKey] = new List<Newtonsoft.Json.Linq.JObject>();
				groupSheet[groupKey] = parentSheet;
			}
			nested[groupKey].Add(row);
		}
	}
	HashSet<string> mainIds = new HashSet<string>(Enumerable.Range(0, mainCount).SelectMany(sheet => sheets[sheet]["rows"])
		.Select(r => (string)r["id"]).Where(id => id != null), StringComparer.Ordinal);
	int groups = Enumerable.Range(0, mainCount).Sum(sheet => sheets[sheet]["rows"].Count()) + order.Count(g => g.Item2 == null && !mainIds.Contains(g.Item3));
	int batchSize = BatchSize(config);
	if (groups > batchSize)
	{
		throw new ApplicationException("The batch holds " + groups + " rows and parents, and a batch takes at most " + batchSize + ".");
	}

	// A relationship row can point at an item that a main-sheet row of this batch creates.
	// Each new item gets its id before anything is written, and the items such rows point at
	// are written first. When new items point at each other, a row whose related item is not
	// written yet waits, and is added on its own once every group is written. Adding it does
	// not give its parent a new generation.
	Dictionary<string, string> newIds = new Dictionary<string, string>(StringComparer.Ordinal);
	for (int sheet = 0; sheet < mainCount; sheet++)
	{
		foreach (Newtonsoft.Json.Linq.JObject row in sheets[sheet]["rows"])
		{
			if ((string)row["action"] == "new" && (string)outcome[sheet + ":" + (int)row["row"]]["status"] == "write")
			{
				newIds[NewRowName(sheet, (int)row["row"])] = inn.getNewID();
			}
		}
	}
	// A file token stands for the id of its stored File, as a row name does for a new item.
	foreach (KeyValuePair<string, string> stored in storedFiles)
	{
		newIds[stored.Key] = stored.Value;
	}
	HashSet<string> pointedAt = new HashSet<string>(order.SelectMany(group => nested[group.Item1]).Select(NewRelatedName).Where(name => name != null), StringComparer.Ordinal);
	order = order.Where(group => group.Item2 != null && pointedAt.Contains(group.Item1))
		.Concat(order.Where(group => group.Item2 == null || !pointedAt.Contains(group.Item1))).ToList();

	// Write, as the caller. The first write that fails ends the batch, and its error
	// rolls back everything this request wrote.
	List<string[]> records = new List<string[]>();
	HashSet<string> writtenRows = new HashSet<string>(StringComparer.Ordinal);
	List<Tuple<Newtonsoft.Json.Linq.JObject, string[]>> waiting = new List<Tuple<Newtonsoft.Json.Linq.JObject, string[]>>();
	// The answer when a write fails: the error that ends the batch, naming the rows it held.
	Func<Newtonsoft.Json.Linq.JObject, IEnumerable<Newtonsoft.Json.Linq.JObject>, bool, string, Item> writeFailed = (named, rows, withMain, failure) =>
	{
		int namedSheet = (int)named["sheetIndex"];
		Newtonsoft.Json.Linq.JArray failedRows = new Newtonsoft.Json.Linq.JArray();
		if (withMain)
		{
			failedRows.Add(RowRef(namedSheet, (int)named["row"]));
		}
		foreach (Newtonsoft.Json.Linq.JObject child in rows)
		{
			failedRows.Add(RowRef((int)child["sheetIndex"], (int)child["row"]));
		}
		string sheetName = (string)sheets[namedSheet]["sheetName"];
		Newtonsoft.Json.Linq.JObject detail = new Newtonsoft.Json.Linq.JObject();
		detail.Add("sheet", namedSheet);
		detail.Add("sheetName", sheetName);
		detail.Add("row", (int)named["row"]);
		detail.Add("message", failure);
		detail.Add("rows", failedRows);
		Item error = inn.newError("Row " + (int)named["row"] + " of sheet '" + sheetName + "' could not be imported: " + failure);
		error.setErrorCode(RowFailedCode);
		error.setErrorDetail(detail.ToString(Newtonsoft.Json.Formatting.None));
		return error;
	};
	// Applies AML; the error text when it fails, else null.
	Func<string, Tuple<Item, string>> write = (aml) =>
	{
		try
		{
			Item answer = inn.applyAML("<AML>" + aml + "</AML>");
			return Tuple.Create(answer, answer.isError() ? answer.getErrorString() : null);
		}
		catch (Exception ex)
		{
			return Tuple.Create((Item)null, FaultText(ex.Message));
		}
	};
	foreach (Tuple<string, Newtonsoft.Json.Linq.JObject, string> group in order)
	{
		Newtonsoft.Json.Linq.JObject mainSheet = (Newtonsoft.Json.Linq.JObject)sheets[groupSheet[group.Item1]];
		string mainType = (string)mainSheet["itemType"]["name"];
		List<Newtonsoft.Json.Linq.JObject> children = nested[group.Item1];
		// Only a new item can be written before the item its row points at: those come first.
		List<Newtonsoft.Json.Linq.JObject> waits = children.Where(child => group.Item2 != null && (string)child["action"] == "new" && NewRelatedName(child) != null &&
			newIds.ContainsKey(NewRelatedName(child)) && !writtenRows.Contains(NewRelatedName(child))).ToList();
		children = children.Where(child => !waits.Contains(child)).ToList();
		StringBuilder relationships = new StringBuilder();
		foreach (Newtonsoft.Json.Linq.JObject child in children)
		{
			Newtonsoft.Json.Linq.JObject childSheet = (Newtonsoft.Json.Linq.JObject)sheets[(int)child["sheetIndex"]];
			string type = (string)childSheet["itemType"]["name"];
			if ((string)child["action"] == "delete")
			{
				relationships.Append("<Item type=\"").Append(Xml(type)).Append("\" action=\"delete\" id=\"").Append(NewestId(newIds, (string)child["id"])).Append("\"/>");
				continue;
			}
			if ((string)child["action"] == "new")
			{
				relationships.Append("<Item type=\"").Append(Xml(type)).Append("\" action=\"add\">");
				Newtonsoft.Json.Linq.JObject related = child["related"] as Newtonsoft.Json.Linq.JObject;
				if (related != null)
				{
					string relatedRow = NewRelatedName(child);
					string relatedFile = related["file"] != null && related["file"].Type == Newtonsoft.Json.Linq.JTokenType.String ? (string)related["file"] : null;
					relationships.Append("<related_id>")
						.Append(NewestId(newIds, relatedRow ?? relatedFile ?? (string)related["id"]))
						.Append("</related_id>");
				}
			}
			else
			{
				relationships.Append("<Item type=\"").Append(Xml(type)).Append("\" action=\"edit\" id=\"").Append(NewestId(newIds, (string)child["id"])).Append("\">");
			}
			relationships.Append(Properties(childSheet, (Newtonsoft.Json.Linq.JObject)child["values"], newIds));
			relationships.Append("</Item>");
		}
		string relationshipsXml = relationships.Length == 0 ? "" : "<Relationships>" + relationships + "</Relationships>";

		string aml;
		Newtonsoft.Json.Linq.JObject named;
		if (group.Item2 == null)
		{
			aml = "<Item type=\"" + Xml(mainType) + "\" action=\"edit\" id=\"" + NewestId(newIds, group.Item3) + "\">" + relationshipsXml + "</Item>";
			named = (Newtonsoft.Json.Linq.JObject)(children.Count > 0 ? children[0] : waits[0]);
		}
		else if ((string)group.Item2["action"] == "new")
		{
			aml = "<Item type=\"" + Xml(mainType) + "\" action=\"add\" id=\"" + newIds[group.Item1] + "\">" +
				Properties(mainSheet, (Newtonsoft.Json.Linq.JObject)group.Item2["values"], newIds) + relationshipsXml + "</Item>";
			named = group.Item2;
		}
		else
		{
			aml = "<Item type=\"" + Xml(mainType) + "\" action=\"edit\" id=\"" + NewestId(newIds, (string)group.Item2["id"]) + "\">" +
				Properties(mainSheet, (Newtonsoft.Json.Linq.JObject)group.Item2["values"], newIds) + relationshipsXml + "</Item>";
			named = group.Item2;
		}

		Tuple<Item, string> applied = write(aml);
		if (applied.Item2 != null)
		{
			// The rows that wait fail with their parent. A deadlock is no fault of the rows: the
			// whole batch is sent again.
			rowFailure = IsDeadlock(applied.Item2) ? RetryError(applied.Item2) : writeFailed(named, children.Concat(waits), group.Item2 != null, applied.Item2);
			return null;
		}

		// The item as it is now, in the new generation when the write gave it one.
		Item written = applied.Item1;
		string newId = written.getID();
		string keyedName = written.getProperty("keyed_name", "");
		string configId = written.getProperty("config_id", newId);
		string oldId = group.Item2 == null ? group.Item3 : (string)group.Item2["id"];
		if (!string.IsNullOrEmpty(oldId) && oldId != newId)
		{
			newIds[oldId] = newId;
		}
		if (group.Item2 != null)
		{
			records.Add(Record(outcome[groupSheet[group.Item1] + ":" + (int)group.Item2["row"]], (string)mainSheet["sheetName"], group.Item2, mainType, newId, configId, keyedName));
		}
		foreach (Newtonsoft.Json.Linq.JObject child in children)
		{
			int sheet = (int)child["sheetIndex"];
			records.Add(Record(outcome[sheet + ":" + (int)child["row"]], (string)sheets[sheet]["sheetName"], child, mainType, newId, configId, keyedName));
		}
		if (group.Item2 != null && (string)group.Item2["action"] == "new")
		{
			writtenRows.Add(group.Item1);
		}
		foreach (Newtonsoft.Json.Linq.JObject child in waits)
		{
			waiting.Add(Tuple.Create(child, new string[] { newId, configId, keyedName, mainType }));
		}
	}

	// The rows that waited for their related item: each added on its own, under its parent.
	foreach (Tuple<Newtonsoft.Json.Linq.JObject, string[]> entry in waiting)
	{
		Newtonsoft.Json.Linq.JObject child = entry.Item1;
		int sheet = (int)child["sheetIndex"];
		Newtonsoft.Json.Linq.JObject childSheet = (Newtonsoft.Json.Linq.JObject)sheets[sheet];
		string aml = "<Item type=\"" + Xml((string)childSheet["itemType"]["name"]) + "\" action=\"add\">" +
			"<source_id>" + entry.Item2[0] + "</source_id>" +
			"<related_id>" + NewestId(newIds, NewRelatedName(child)) + "</related_id>" +
			Properties(childSheet, (Newtonsoft.Json.Linq.JObject)child["values"], newIds) + "</Item>";
		Tuple<Item, string> applied = write(aml);
		if (applied.Item2 != null)
		{
			rowFailure = IsDeadlock(applied.Item2) ? RetryError(applied.Item2) : writeFailed(child, new Newtonsoft.Json.Linq.JObject[] { child }, false, applied.Item2);
			return null;
		}
		records.Add(Record(outcome[sheet + ":" + (int)child["row"]], (string)childSheet["sheetName"], child, entry.Item2[3], entry.Item2[0], entry.Item2[1], entry.Item2[2]));
	}

	// The records of the rows written, and the stored Files they linked, in the same
	// transaction as the rows.
	HashSet<string> linked = new HashSet<string>(sheets.SelectMany((sheet, index) => sheet["rows"].Cast<Newtonsoft.Json.Linq.JObject>()
		.Where(row => new string[] { "created", "updated" }.Contains((string)outcome[index + ":" + (int)row["row"]]["status"]))
		.SelectMany(FileTokensOf)).Select(token => storedFiles[token]), StringComparer.OrdinalIgnoreCase);
	AsSteps(() => AddRecords(jobId, records));
	if (staged.Any(entry => linked.Contains((string)entry["id"]) && !(bool)entry["linked"]))
	{
		// Other batches may be marking their Files at the same time: see UpdateStaged.
		UpdateStaged(jobId, list =>
		{
			bool changed = false;
			foreach (Newtonsoft.Json.Linq.JObject entry in list)
			{
				if (linked.Contains((string)entry["id"]) && !(bool)entry["linked"])
				{
					entry["linked"] = true;
					changed = true;
				}
			}
			return changed;
		});
	}

	Newtonsoft.Json.Linq.JObject counts = new Newtonsoft.Json.Linq.JObject();
	foreach (string status in new string[] { "created", "updated", "deleted", "unchanged", "failed" })
	{
		counts.Add(status, rowsOut.Count(r => (string)r["status"] == status));
	}
	Newtonsoft.Json.Linq.JObject answer = new Newtonsoft.Json.Linq.JObject();
	answer.Add("job", JobJson(job));
	answer.Add("rows", rowsOut);
	answer.Add("counts", counts);
	return answer.ToString(Newtonsoft.Json.Formatting.None);
}

string Complete()
{
	XmlElement job = ReadJob(new string[] { "Validated" });
	string jobId = job.GetAttribute("id");
	CheckAllowed(job, ReadConfiguration());
	CheckToken(job, Input("token"));
	CheckUploaded(job);

	List<string[]> failed = new List<string[]>();
	Newtonsoft.Json.Linq.JObject counted;
	try
	{
		Newtonsoft.Json.Linq.JArray list = ParseJson(Input("failed").Length == 0 ? "[]" : Input("failed")) as Newtonsoft.Json.Linq.JArray;
		foreach (Newtonsoft.Json.Linq.JToken entry in list)
		{
			int row;
			if (!int.TryParse(Str(entry, "row"), NumberStyles.Integer, CultureInfo.InvariantCulture, out row) || row < 1)
			{
				throw new FormatException("row");
			}
			int sheet;
			string order = int.TryParse(Str(entry, "sheet"), NumberStyles.Integer, CultureInfo.InvariantCulture, out sheet) && sheet >= 0 ? RecordOrder(sheet, row) : "";
			failed.Add(new string[] { "Failed", Clip(Str(entry, "sheetName") ?? "", 64), row.ToString(CultureInfo.InvariantCulture), "", "", "", Clip(Str(entry, "key") ?? "", 128), Str(entry, "message") ?? "", order });
		}
		counted = ParseJson(Input("counts")) as Newtonsoft.Json.Linq.JObject;
		if (counted == null)
		{
			throw new FormatException("counts");
		}
	}
	catch (Exception)
	{
		throw new ApplicationException("FRG_MCE_ImportWorkbook: 'failed' must be a JSON list of { sheetName, row, key, message } and 'counts' a JSON object { total, unchanged, skipped }.");
	}
	int total = Count(counted, "total");
	int unchanged = Count(counted, "unchanged");
	int skipped = Count(counted, "skipped");

	// The Files stored for the import that no row linked: the files of rows that failed.
	List<string> unlinked = ReadStaged(job).Where(entry => !(bool)entry["linked"]).Select(entry => (string)entry["id"]).ToList();
	Dictionary<string, int> records = null;
	AsSteps(() =>
	{
		// The records first: once the job is Imported, only administrators may change it.
		AddRecords(jobId, failed);
		records = RecordCounts(jobId);
		Apply(
			"<Item type=\"" + JobType + "\" action=\"edit\" id=\"" + jobId + "\">" +
			"<staged_files is_null=\"1\"/>" +
			"<qty_total>" + total + "</qty_total>" +
			"<qty_created>" + records["Created"] + "</qty_created>" +
			"<qty_updated>" + records["Updated"] + "</qty_updated>" +
			"<qty_unchanged>" + unchanged + "</qty_unchanged>" +
			"<qty_skipped>" + skipped + "</qty_skipped>" +
			"<qty_failed>" + failed.Count + "</qty_failed>" +
			"</Item>");
		Promote(jobId, "Imported");
	});
	DeleteReplacedFiles(unlinked, true);

	Newtonsoft.Json.Linq.JObject counts = new Newtonsoft.Json.Linq.JObject();
	counts.Add("total", total);
	counts.Add("created", records["Created"]);
	counts.Add("updated", records["Updated"]);
	counts.Add("deleted", records["Deleted"]);
	counts.Add("unchanged", unchanged);
	counts.Add("skipped", skipped);
	counts.Add("failed", failed.Count);
	Newtonsoft.Json.Linq.JObject answer = new Newtonsoft.Json.Linq.JObject();
	answer.Add("job", JobJson(ReadJob(null)));
	answer.Add("counts", counts);
	return answer.ToString(Newtonsoft.Json.Formatting.None);
}

// ---------------------------------------------------------------------------------
// Writing rows
// ---------------------------------------------------------------------------------

// The properties of a row as AML: each value under the property of its column. An empty
// value empties the property. A reference to an item that this batch gave a new
// generation names that generation.
string Properties(Newtonsoft.Json.Linq.JObject sheet, Newtonsoft.Json.Linq.JObject values, Dictionary<string, string> newIds)
{
	StringBuilder xml = new StringBuilder();
	if (values == null)
	{
		return "";
	}
	Dictionary<string, Newtonsoft.Json.Linq.JToken> columns = sheet["columns"].ToDictionary(c => (string)c["id"], c => c, StringComparer.Ordinal);
	foreach (Newtonsoft.Json.Linq.JProperty value in values.Properties())
	{
		string property;
		bool isItem;
		if (value.Name == "related_id")
		{
			property = "related_id";
			isItem = true;
		}
		else
		{
			Newtonsoft.Json.Linq.JToken column;
			if (!columns.TryGetValue(value.Name, out column) || (string)column["property"] == null)
			{
				continue;
			}
			property = (string)column["property"];
			isItem = (string)column["dataType"] == "item" || (string)column["dataType"] == "file";
		}
		if (!System.Text.RegularExpressions.Regex.IsMatch(property, "^[A-Za-z0-9_]+$"))
		{
			continue;
		}
		string text = (string)value.Value ?? "";
		if (text.Length == 0)
		{
			xml.Append('<').Append(property).Append(" is_null=\"1\"/>");
		}
		else
		{
			xml.Append('<').Append(property).Append('>').Append(Xml(isItem ? NewestId(newIds, text) : text)).Append("</").Append(property).Append('>');
		}
	}
	return xml.ToString();
}

// The message of a failed write: the text of the fault, when the message holds the fault
// as XML, which is what the user needs.
string FaultText(string message)
{
	System.Text.RegularExpressions.Match fault = System.Text.RegularExpressions.Regex.Match(
		message ?? "", "<faultstring>(.*?)</faultstring>", System.Text.RegularExpressions.RegexOptions.Singleline);
	if (!fault.Success)
	{
		return message ?? "";
	}
	XmlDocument text = inn.newXMLDocument();
	try
	{
		text.LoadXml("<t>" + fault.Groups[1].Value + "</t>");
		return text.DocumentElement.InnerText.Trim();
	}
	catch (XmlException)
	{
		return fault.Groups[1].Value.Trim();
	}
}

// How FRG_MCE_ValidateImport names an item that a main-sheet row creates, before it has an
// id: in related_id values, and as the key of the id the batch gives the item in newIds.
// "row:<number>" on the first main sheet, and "row:<sheet>:<number>" on the others.
string NewRowName(int sheet, int rowNumber)
{
	return sheet == 0
		? "row:" + rowNumber.ToString(CultureInfo.InvariantCulture)
		: "row:" + sheet.ToString(CultureInfo.InvariantCulture) + ":" + rowNumber.ToString(CultureInfo.InvariantCulture);
}

// The main-sheet row that creates the related item of a relationship row, { sheet, row }, or
// null when the related item exists (or the row has none).
int[] NewRelatedRow(Newtonsoft.Json.Linq.JObject row)
{
	Newtonsoft.Json.Linq.JObject related = row["related"] as Newtonsoft.Json.Linq.JObject;
	if (related == null || related["id"] == null || related["id"].Type != Newtonsoft.Json.Linq.JTokenType.Null ||
		related["row"] == null || related["row"].Type != Newtonsoft.Json.Linq.JTokenType.Integer)
	{
		return null;
	}
	return new int[] { (int)related["sheet"], (int)related["row"] };
}

// The name of the new item a relationship row points at (NewRowName), or null.
string NewRelatedName(Newtonsoft.Json.Linq.JObject row)
{
	int[] related = NewRelatedRow(row);
	return related == null ? null : NewRowName(related[0], related[1]);
}

// The main sheet of the parent of a relationship row: the one its sheet names.
int ParentSheetOf(Newtonsoft.Json.Linq.JArray sheets, int sheet)
{
	return (int)sheets[sheet]["parentSheet"];
}

// The file tokens a row of the check names: in its values (file cells, and related_id when
// it gets another file), and as its related file.
IEnumerable<string> FileTokensOf(Newtonsoft.Json.Linq.JObject row)
{
	List<string> tokens = new List<string>();
	Newtonsoft.Json.Linq.JObject values = row["values"] as Newtonsoft.Json.Linq.JObject;
	foreach (Newtonsoft.Json.Linq.JProperty value in values == null ? new List<Newtonsoft.Json.Linq.JProperty>() : values.Properties().ToList())
	{
		if (value.Value.Type == Newtonsoft.Json.Linq.JTokenType.String && IsFileToken((string)value.Value))
		{
			tokens.Add((string)value.Value);
		}
	}
	Newtonsoft.Json.Linq.JObject related = row["related"] as Newtonsoft.Json.Linq.JObject;
	if (related != null && related["file"] != null && related["file"].Type == Newtonsoft.Json.Linq.JTokenType.String && IsFileToken((string)related["file"]))
	{
		tokens.Add((string)related["file"]);
	}
	return tokens.Distinct(StringComparer.Ordinal);
}

// file:<MD5>:<file name>, as FRG_MCE_ValidateImport writes it for a picked file.
bool IsFileToken(string value)
{
	return System.Text.RegularExpressions.Regex.IsMatch(value ?? "", "^file:[0-9A-F]{32}:.+$");
}

string TokenName(string token)
{
	return IsFileToken(token) ? token.Substring(38) : token;
}

// The job's list of the Files stored for its import: [ { id, md5, name, linked } ].
Newtonsoft.Json.Linq.JArray ReadStaged(XmlElement job)
{
	Newtonsoft.Json.Linq.JArray staged = new Newtonsoft.Json.Linq.JArray();
	string text = Text(job, "staged_files").Trim();
	Newtonsoft.Json.Linq.JArray list = null;
	try
	{
		list = text.Length == 0 ? null : ParseJson(text) as Newtonsoft.Json.Linq.JArray;
	}
	catch (Exception)
	{
		list = null;
	}
	foreach (Newtonsoft.Json.Linq.JToken entry in list ?? new Newtonsoft.Json.Linq.JArray())
	{
		if (entry is Newtonsoft.Json.Linq.JObject && Str(entry, "id") != null && IsId(Str(entry, "id")) && Str(entry, "md5") != null && Str(entry, "name") != null)
		{
			Newtonsoft.Json.Linq.JObject clean = new Newtonsoft.Json.Linq.JObject();
			clean.Add("id", Str(entry, "id"));
			clean.Add("md5", Str(entry, "md5").ToUpperInvariant());
			clean.Add("name", Str(entry, "name"));
			clean.Add("linked", entry["linked"] != null && entry["linked"].Type == Newtonsoft.Json.Linq.JTokenType.Boolean && (bool)entry["linked"]);
			staged.Add(clean);
		}
	}
	return staged;
}

// The stored Files by file token, read again as the caller: a File whose checksum or name is
// not the one listed for it is left out.
Dictionary<string, string> StoredFileIds(Newtonsoft.Json.Linq.JArray staged)
{
	Dictionary<string, string> ids = new Dictionary<string, string>(StringComparer.Ordinal);
	if (staged.Count == 0)
	{
		return ids;
	}
	foreach (XmlElement file in Query("<Item type=\"File\" action=\"get\" select=\"filename,checksum\" idlist=\"" +
		string.Join(",", staged.Select(entry => (string)entry["id"])) + "\"/>"))
	{
		string token = "file:" + Text(file, "checksum").ToUpperInvariant() + ":" + Text(file, "filename");
		if (IsFileToken(token) && !ids.ContainsKey(token))
		{
			ids[token] = file.GetAttribute("id");
		}
	}
	return ids;
}

// The id an item has now: the newest generation a write of this batch gave it, or the id
// a new item gets, for its row name (NewRowName).
string NewestId(Dictionary<string, string> newIds, string id)
{
	string newer;
	while (id != null && newIds.TryGetValue(id, out newer))
	{
		id = newer;
	}
	return id;
}

// A record of a row that was written: { action, sheet, row, type, id, config id, keyed name, message }.
string[] Record(Newtonsoft.Json.Linq.JObject entry, string sheetName, Newtonsoft.Json.Linq.JObject row, string targetType, string id, string configId, string keyedName)
{
	string status = (string)row["action"] == "new" ? "created" : (string)row["action"] == "delete" ? "deleted" : "updated";
	entry["status"] = status;
	entry.Add("id", id);
	entry.Add("keyedName", keyedName);
	return new string[]
	{
		status == "created" ? "Created" : status == "deleted" ? "Deleted" : "Updated", Clip(sheetName, 64), ((int)row["row"]).ToString(CultureInfo.InvariantCulture),
		targetType, id, configId, Clip(keyedName, 128), "", RecordOrder((int)entry["sheet"], (int)row["row"])
	};
}

// Adds the records: { action, sheet, row, type, id, config id, keyed name, message, sort order }.
// Nothing reads the added records back, so the server does not return them (doGetItem="0");
// their server events run as for any other add. Each record brings its own sort order
// (RecordOrder), so that batches at the same time don't wait on each other's records.
void AddRecords(string jobId, List<string[]> records)
{
	for (int start = 0; start < records.Count; start += RecordChunk)
	{
		StringBuilder aml = new StringBuilder();
		foreach (string[] record in records.Skip(start).Take(RecordChunk))
		{
			aml.Append("<Item type=\"").Append(RecordType).Append("\" action=\"add\" doGetItem=\"0\">")
				.Append("<source_id>").Append(jobId).Append("</source_id>")
				.Append("<action>").Append(record[0]).Append("</action>")
				.Append("<sheet_name>").Append(Xml(record[1])).Append("</sheet_name>")
				.Append("<row_number>").Append(record[2]).Append("</row_number>");
			if (record.Length > 8 && record[8].Length > 0)
			{
				aml.Append("<sort_order>").Append(record[8]).Append("</sort_order>");
			}
			string[] names = new string[] { "target_type", "target_id", "target_config_id", "target_keyed_name", "message" };
			for (int i = 0; i < names.Length; i++)
			{
				if (record[i + 3].Length > 0)
				{
					aml.Append('<').Append(names[i]).Append('>').Append(Xml(record[i + 3])).Append("</").Append(names[i]).Append('>');
				}
			}
			aml.Append("</Item>");
		}
		Apply(aml.ToString());
	}
}

// The job's records by action: Created, Updated and Failed.
Dictionary<string, int> RecordCounts(string jobId)
{
	Dictionary<string, int> counts = new Dictionary<string, int>(StringComparer.Ordinal) { { "Created", 0 }, { "Updated", 0 }, { "Deleted", 0 }, { "Failed", 0 } };
	foreach (XmlElement record in Query("<Item type=\"" + RecordType + "\" action=\"get\" select=\"action\"><source_id>" + jobId + "</source_id></Item>"))
	{
		string action = Text(record, "action");
		if (counts.ContainsKey(action))
		{
			counts[action]++;
		}
	}
	return counts;
}

void Failed(Newtonsoft.Json.Linq.JObject entry, string reason, string message)
{
	entry.Add("status", "failed");
	entry.Add("reason", reason);
	entry.Add("message", message);
}

Newtonsoft.Json.Linq.JObject RowRef(int sheet, int row)
{
	Newtonsoft.Json.Linq.JObject reference = new Newtonsoft.Json.Linq.JObject();
	reference.Add("sheet", sheet);
	reference.Add("row", row);
	return reference;
}

// ---------------------------------------------------------------------------------
// The job and the checks
// ---------------------------------------------------------------------------------

// Reads the job and refuses the call unless its creator is calling, it is in one of the
// given states (null: any), and nobody is editing it.
XmlElement ReadJob(string[] states)
{
	string jobId = Input("job_id");
	if (!IsId(jobId))
	{
		throw new ApplicationException("FRG_MCE_ImportWorkbook: the 'job_id' property must be the id of a job.");
	}
	List<XmlElement> jobs = Query(
		"<Item type=\"" + JobType + "\" action=\"get\" id=\"" + jobId + "\" " +
		"select=\"item_number,state,created_by_id,locked_by_id,target_itemtype,targets,workbook_token,upload_file,results_file,staged_files\"/>");
	if (jobs.Count != 1)
	{
		throw new ApplicationException("Mass Create & Edit job '" + jobId + "' was not found, or you have no access to it.");
	}
	XmlElement job = jobs[0];
	if (states == null)
	{
		return job;
	}
	string number = Text(job, "item_number");
	string userId = inn.getUserID();
	if (Text(job, "created_by_id") != userId)
	{
		throw new ApplicationException(
			"Only the user who created a job can run it. Job " + number + " was created by " + Attribute(job, "created_by_id", "keyed_name") + ".");
	}
	string state = Text(job, "state");
	if (!states.Contains(state))
	{
		throw new ApplicationException(state == "Imported"
			? "Job " + number + " has been imported and cannot be used again. Start a new job."
			: "Job " + number + " is " + state + ". This step needs a job that is " + string.Join(" or ", states) + ".");
	}
	string lockedBy = Text(job, "locked_by_id");
	if (lockedBy == userId)
	{
		throw new ApplicationException("Job " + number + " is being edited. Save it and finish editing first.");
	}
	if (lockedBy.Length > 0)
	{
		throw new ApplicationException("Job " + number + " is locked by " + Attribute(job, "locked_by_id", "keyed_name") + ".");
	}
	if (Text(job, "target_itemtype").Length == 0)
	{
		throw new ApplicationException("Job " + number + " has no target ItemType.");
	}
	if (JobTargetIds(job) == null)
	{
		throw new ApplicationException("The targets of job " + number + " can't be read. Open the job, check its targets and generate the workbook again.");
	}
	return job;
}

// The target ItemTypes of a job, in the order of their main sheets, as FRG_MCE_ValidateImport
// reads them. The job's targets property lists them as JSON, [ { "itemTypeId": "...", "prefillIds":
// [ ... ] }, ... ], when the job has several; a job with one has its target_itemtype. Null when the
// list can't be read, names an ItemType twice, or does not start with target_itemtype.
List<string> JobTargetIds(XmlElement job)
{
	string first = Text(job, "target_itemtype").ToUpperInvariant();
	string text = Text(job, "targets").Trim();
	if (text.Length == 0)
	{
		return new List<string> { first };
	}
	Newtonsoft.Json.Linq.JArray entries;
	try
	{
		entries = ParseJson(text) as Newtonsoft.Json.Linq.JArray;
	}
	catch (Exception)
	{
		return null;
	}
	if (entries == null)
	{
		return null;
	}
	List<string> ids = new List<string>();
	foreach (Newtonsoft.Json.Linq.JToken entry in entries)
	{
		string id = (Str(entry, "itemTypeId") ?? "").Trim().ToUpperInvariant();
		if (!IsId(id) || ids.Contains(id))
		{
			return null;
		}
		ids.Add(id);
	}
	return ids.Count > 0 && ids[0] == first ? ids : null;
}

void CheckToken(XmlElement job, string token)
{
	if (token.Length == 0 || token != Text(job, "workbook_token"))
	{
		throw new ApplicationException(
			"The workbook is not the one job " + Text(job, "item_number") + " generated last. Fill in the latest workbook and upload it again.");
	}
}

void CheckUploaded(XmlElement job)
{
	if (Text(job, "upload_file").Length == 0)
	{
		throw new ApplicationException("Job " + Text(job, "item_number") + " has no uploaded workbook. Upload the workbook first.");
	}
}

// The configuration item: the one named Default, or else the oldest one.
XmlElement ReadConfiguration()
{
	List<XmlElement> configs = Query(
		"<Item type=\"FRG_MassCreateEdit_Config\" action=\"get\" select=\"name,batch_size,parallel_batches\" orderBy=\"created_on\"/>");
	if (configs.Count == 0)
	{
		throw new ApplicationException("Mass Create & Edit has no configuration item, or you have no access to it. Ask an administrator.");
	}
	return configs.FirstOrDefault(c => Text(c, "name") == "Default") ?? configs[0];
}

int BatchSize(XmlElement config)
{
	int size;
	return int.TryParse(Text(config, "batch_size"), NumberStyles.Integer, CultureInfo.InvariantCulture, out size) && size > 0 ? size : 50;
}

// How many batches the browser may send at once: 4 when the configuration item leaves it
// empty, at most 8.
int ParallelBatches(XmlElement config)
{
	int count;
	return int.TryParse(Text(config, "parallel_batches"), NumberStyles.Integer, CultureInfo.InvariantCulture, out count) && count > 0 ? Math.Min(count, 8) : 4;
}

// The allow-list is checked on the server as well, because a caller can skip the Form. A row
// with an identity allows the ItemType only to members of that identity, direct or through
// other identities; a row without one allows it to everyone who can use the tool. Every
// target of the job must be allowed.
void CheckAllowed(XmlElement job, XmlElement config)
{
	List<string> targetIds = JobTargetIds(job) ?? new List<string> { Text(job, "target_itemtype").ToUpperInvariant() };
	List<XmlElement> allowed = Query(
		"<Item type=\"FRG_MassCreateEdit_AllowedType\" action=\"get\" select=\"related_id,identity_id\" related_expand=\"0\">" +
		"<source_id>" + config.GetAttribute("id") + "</source_id></Item>");
	HashSet<string> mine = new HashSet<string>(
		Aras.Server.Security.Permissions.Current.IdentitiesList.Split(',').Select(id => id.Trim().Trim('\'')), StringComparer.OrdinalIgnoreCase);
	foreach (string targetId in targetIds)
	{
		List<XmlElement> rows = allowed.Where(row => string.Equals(Text(row, "related_id"), targetId, StringComparison.OrdinalIgnoreCase)).ToList();
		string targetName = targetId == Text(job, "target_itemtype").ToUpperInvariant()
			? Attribute(job, "target_itemtype", "keyed_name")
			: rows.Count > 0 ? Attribute(rows[0], "related_id", "keyed_name") : ItemTypeName(targetId);
		if (rows.Count == 0)
		{
			throw new ApplicationException(
				"ItemType '" + targetName + "' is not on the allow-list of Mass Create & Edit. " +
				"An administrator can add it on the configuration item.");
		}
		if (!rows.Any(row => Text(row, "identity_id").Length == 0 || mine.Contains(Text(row, "identity_id"))))
		{
			throw new ApplicationException(
				"Mass Create & Edit on ItemType '" + targetName + "' is limited to the members of identity '" +
				Attribute(rows[0], "identity_id", "keyed_name") + "'. Ask an administrator.");
		}
	}
}

// The name of an ItemType, for a message; its id when it can't be read.
string ItemTypeName(string itemTypeId)
{
	List<XmlElement> types = Query("<Item type=\"ItemType\" action=\"get\" select=\"name\" id=\"" + itemTypeId + "\"/>");
	return types.Count == 1 ? Text(types[0], "name") : itemTypeId;
}

// The id of a File item the caller stored, given in an input property.
string FileOfCaller(string inputName, string what)
{
	string fileId = Input(inputName);
	if (!IsId(fileId))
	{
		throw new ApplicationException("FRG_MCE_ImportWorkbook: the '" + inputName + "' property must be the id of a File item.");
	}
	List<XmlElement> files = Query("<Item type=\"File\" action=\"get\" select=\"created_by_id\" id=\"" + fileId + "\"/>");
	if (files.Count != 1)
	{
		throw new ApplicationException("The file of the " + what + " was not found, or you have no access to it.");
	}
	if (Text(files[0], "created_by_id") != inn.getUserID())
	{
		throw new ApplicationException("The file of the " + what + " was not stored by you.");
	}
	return fileId;
}

// The Files of the job's earlier upload that this upload replaces: its uploaded workbook
// (upload_file) and its validation result (results_file), each when the caller stored it
// and it is not one of the new upload's Files. Only upload sets these properties, and it
// takes only Files the caller stored, so these are the Files of an upload the tool
// attached to the job. Asked while the job still points to the Files.
List<string> ReplacedUploadFiles(XmlElement job, string uploadId, string resultsId)
{
	List<string> fileIds = new List<string>();
	foreach (string name in new string[] { "upload_file", "results_file" })
	{
		string fileId = Text(job, name);
		if (!IsId(fileId) || fileIds.Contains(fileId, StringComparer.OrdinalIgnoreCase) ||
			string.Equals(fileId, uploadId, StringComparison.OrdinalIgnoreCase) || string.Equals(fileId, resultsId, StringComparison.OrdinalIgnoreCase))
		{
			continue;
		}
		List<XmlElement> files = Query("<Item type=\"File\" action=\"get\" select=\"created_by_id\" id=\"" + fileId + "\"/>");
		if (files.Count == 1 && Text(files[0], "created_by_id") == inn.getUserID())
		{
			fileIds.Add(fileId);
		}
	}
	return fileIds;
}

// Deletes the Files that upload has just detached from the job. Its creator may not change
// a job past Draft, so these Files are deleted with the Super User identity, each only
// when no job points to it any more. The copies the user downloaded are not affected. With keepLinked, it also
// deletes the Files an import stored and no row linked; a File that cannot be deleted, as
// something else links it now, is kept.
void DeleteReplacedFiles(List<string> fileIds, bool keepLinked = false)
{
	if (fileIds.Count == 0)
	{
		return;
	}
	Aras.Server.Security.Identity superUser = Aras.Server.Security.Identity.GetById(SuperUserIdentityId);
	bool granted = Aras.Server.Security.Permissions.GrantIdentity(superUser);
	try
	{
		foreach (string fileId in fileIds)
		{
			List<XmlElement> users = Query(
				"<Item type=\"" + JobType + "\" action=\"get\" select=\"id\"><OR>" +
				"<export_file>" + fileId + "</export_file><upload_file>" + fileId + "</upload_file><results_file>" + fileId + "</results_file>" +
				"</OR></Item>");
			if (users.Count > 0)
			{
				continue;
			}
			Item deleted = inn.applyAML("<AML><Item type=\"File\" action=\"delete\" id=\"" + fileId + "\"/></AML>");
			if (deleted.isError() && !deleted.isEmpty() && !keepLinked)
			{
				throw new ApplicationException("A file of the job's earlier upload could not be deleted: " + deleted.getErrorString());
			}
		}
	}
	finally
	{
		if (granted)
		{
			Aras.Server.Security.Permissions.RevokeIdentity(superUser);
		}
	}
}

// Changes the job's list of stored Files (staged_files). Batches of one import may do this at
// the same time, so the job is held first: an edit that changes nothing makes the other
// requests that change the job wait until this one is done. Only then is the list read, and
// change gets it to change; it answers whether it changed anything. Answers the list as it is
// now.
Newtonsoft.Json.Linq.JArray UpdateStaged(string jobId, Func<Newtonsoft.Json.Linq.JArray, bool> change)
{
	Newtonsoft.Json.Linq.JArray staged = null;
	AsSteps(() =>
	{
		Apply("<Item type=\"" + JobType + "\" action=\"edit\" id=\"" + jobId + "\"/>");
		List<XmlElement> jobs = Query("<Item type=\"" + JobType + "\" action=\"get\" id=\"" + jobId + "\" select=\"staged_files\"/>");
		if (jobs.Count != 1)
		{
			throw new ApplicationException("Mass Create & Edit job '" + jobId + "' was not found.");
		}
		staged = ReadStaged(jobs[0]);
		if (change(staged))
		{
			Apply(
				"<Item type=\"" + JobType + "\" action=\"edit\" id=\"" + jobId + "\">" +
				"<staged_files>" + Xml(staged.ToString(Newtonsoft.Json.Formatting.None)) + "</staged_files></Item>");
		}
	});
	return staged;
}

// The sort order of the record of a row: its sheet, then its row. A row number is below
// 2,000,000 (Excel has 1,048,576 rows).
string RecordOrder(int sheet, int row)
{
	return ((long)sheet * 2000000L + row).ToString(CultureInfo.InvariantCulture);
}

// Whether an error is the database choosing this request as the victim of a deadlock.
bool IsDeadlock(string message)
{
	return (message ?? "").IndexOf("deadlock", StringComparison.OrdinalIgnoreCase) >= 0;
}

// The answer of a batch that met a deadlock: nothing of it is kept, and it can be sent again.
Item RetryError(string message)
{
	Item error = inn.newError("The batch met another request in the database and was rolled back. It can be sent again. (" + FaultText(message) + ")");
	error.setErrorCode(RetryCode);
	return error;
}

// Runs the given writes with the identity of the steps of Mass Create & Edit, the only
// identity that may change a job's own properties, state and records.
void AsSteps(Action writes)
{
	Aras.Server.Security.Identity steps;
	try
	{
		steps = Aras.Server.Security.Identity.GetById(StepsIdentityId);
	}
	catch (Exception)
	{
		throw new ApplicationException("The identity 'FRG Mass Create & Edit Steps' is missing. Mass Create & Edit is not installed completely; ask an administrator.");
	}
	bool granted = Aras.Server.Security.Permissions.GrantIdentity(steps);
	try
	{
		writes();
	}
	finally
	{
		if (granted)
		{
			Aras.Server.Security.Permissions.RevokeIdentity(steps);
		}
	}
}

void Promote(string jobId, string state)
{
	Apply("<Item type=\"" + JobType + "\" action=\"promoteItem\" id=\"" + jobId + "\"><state>" + state + "</state></Item>");
}

Newtonsoft.Json.Linq.JObject JobJson(XmlElement job)
{
	Newtonsoft.Json.Linq.JObject json = new Newtonsoft.Json.Linq.JObject();
	json.Add("id", job.GetAttribute("id"));
	json.Add("number", Text(job, "item_number"));
	json.Add("state", Text(job, "state"));
	return json;
}

// ---------------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------------

// An input property of the Method, trimmed; empty when it was not given.
string Input(string name)
{
	return (this.getProperty(name, "") ?? "").Trim();
}

bool IsId(string value)
{
	return value.Length == 32 && value.All(c => (c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f'));
}

int Count(Newtonsoft.Json.Linq.JObject counted, string name)
{
	int number;
	if (!int.TryParse(Str(counted, name) ?? "", NumberStyles.Integer, CultureInfo.InvariantCulture, out number) || number < 0)
	{
		throw new ApplicationException("FRG_MCE_ImportWorkbook: 'counts' must give '" + name + "' as a number of rows.");
	}
	return number;
}

// Reads JSON text. Text that looks like a date stays text, and anything after the
// first value is refused.
Newtonsoft.Json.Linq.JToken ParseJson(string text)
{
	using (System.IO.StringReader source = new System.IO.StringReader(text))
	using (Newtonsoft.Json.JsonTextReader reader = new Newtonsoft.Json.JsonTextReader(source))
	{
		reader.DateParseHandling = Newtonsoft.Json.DateParseHandling.None;
		reader.MaxDepth = MaxJsonDepth;
		Newtonsoft.Json.Linq.JToken token = Newtonsoft.Json.Linq.JToken.ReadFrom(reader);
		while (reader.Read())
		{
			if (reader.TokenType != Newtonsoft.Json.JsonToken.Comment)
			{
				throw new FormatException("Unexpected text after the JSON value.");
			}
		}
		return token;
	}
}

// A property of a JSON object as text; null when it is not there or not a single value.
string Str(Newtonsoft.Json.Linq.JToken parent, string name)
{
	Newtonsoft.Json.Linq.JObject owner = parent as Newtonsoft.Json.Linq.JObject;
	Newtonsoft.Json.Linq.JValue value = owner == null ? null : owner[name] as Newtonsoft.Json.Linq.JValue;
	return value == null || value.Value == null ? null : Convert.ToString(value.Value, CultureInfo.InvariantCulture);
}

string Clip(string value, int length)
{
	return value.Length > length ? value.Substring(0, length - 3) + "..." : value;
}

// Runs one AML query as the caller and returns the items found. Finding nothing is an
// empty result here.
List<XmlElement> Query(string aml)
{
	Item response = inn.applyAML("<AML>" + aml + "</AML>");
	List<XmlElement> items = new List<XmlElement>();
	if (response.isError())
	{
		if (response.isEmpty())
		{
			return items;
		}
		throw new ApplicationException(response.getErrorString());
	}
	foreach (XmlNode item in response.dom.SelectNodes("//Result/Item"))
	{
		items.Add((XmlElement)item);
	}
	return items;
}

// Runs one AML request that writes. A failure ends the Method.
void Apply(string aml)
{
	Item response = inn.applyAML("<AML>" + aml + "</AML>");
	if (response.isError())
	{
		throw new ApplicationException(response.getErrorString());
	}
}

// The value of a property; for an item property that came back as a nested item, its id.
string Text(XmlElement item, string name)
{
	XmlElement value = item[name];
	if (value == null)
	{
		return "";
	}
	XmlElement nested = value["Item"];
	return nested == null ? value.InnerText : nested.GetAttribute("id");
}

string Attribute(XmlElement item, string name, string attribute)
{
	XmlElement value = item[name];
	return value == null ? "" : value.GetAttribute(attribute);
}

string Xml(string value)
{
	return value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;");
}
