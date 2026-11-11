/**
 * Tests for the import logic (mce-import.js): key texts, the preview filters and
 * cells, the import plan with its units and batches, the slice of the workbook a
 * batch sends, and what happens to a batch when a row fails. Run from the
 * repository root with `node --test`.
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const {
	ACTION_LABELS, DEFAULT_FILTER, actionCounts, addKeyTexts, batchRows, cellOf, dropFailedWrite, dropRowsPointingAt, failedRowEntry, filterRows, findRow, importCounts, jsonObjectsIn,
	keyHeader, keyText, makeBatches, planImport, progressText, resultsFileName, rowMessages, rowName, rowPasses, showValue, sliceWorkbook,
	readBatchAnswer, skippedRecords, sortRecords, summaryLine, unitRowCount, valueColumns, workbookMatchesResult, writeFailureOf,
	failedRowLeavesNoItem, isMainSheet, mainSheetIndexes, recordCountsByTarget, resultTargets, sheetGroups, sheetOrder, sheetTargets, targetCountsLine, newRelatedItem, oversizedBlocks, oversizedWarning, parentSheetOf, rowKey,
	batchDependencies, isRetryFault, retryDelay, runScheduled
} = await import(new URL('../../1.0.0/Innovator/Client/customer/masscreateedit/mce-import.js', import.meta.url).href);

const fixture = () => JSON.parse(readFileSync(new URL('./fixtures/validate-import.result.json', import.meta.url), 'utf8'));

/** The workbook the fixture was validated from, as parseWorkbook would give it. */
function workbookOf(result) {
	const cellsFor = (sheet, row) => {
		const cells = {};
		if (sheet.kind === 'main') {
			const numbers = { 2: 'TA-001', 3: 'TP-001', 4: 'P-9001', 5: 'P-9002' };
			if (numbers[row.row]) {
				cells.item_number = numbers[row.row];
			}
			cells.name = 'name ' + row.row;
		} else {
			cells['parent.item_number'] = { 2: 'TA-001', 4: 'P-9001', 5: 'P-9002' }[row.parent.row] || 'TA-001';
			cells['related.item_number'] = row.related ? row.related.keyedName : 'NOPE';
			cells.sort_order = row.row * 10;
		}
		return cells;
	};
	return {
		meta: { formatVersion: 1, job: { id: result.job.id, token: 'T' }, sheets: result.sheets.map((sheet) => ({ sheetName: sheet.sheetName })) },
		sheets: result.sheets.map((sheet) => ({
			sheetName: sheet.sheetName,
			columns: sheet.columns.map((column) => column.id),
			rows: sheet.rows.map((row) => ({ rowNumber: row.row, cells: cellsFor(sheet, row), modifiedOn: null }))
		}))
	};
}

/** Gives every row of a result a digest. */
function withDigests(result) {
	result.sheets.forEach((sheet, s) => sheet.rows.forEach((row) => {
		row.digest = row.digest || 'd' + s + '-' + row.row;
	}));
	return result;
}

// ---------------------------------------------------------------------------
// Keys
// ---------------------------------------------------------------------------

test('keyText joins the key columns of a main-sheet row', () => {
	const sheet = { kind: 'main', keyColumns: ['item_number', 'classification'], columns: [] };
	assert.equal(keyText(sheet, { item_number: 'P-1', classification: 'Component', name: 'x' }), 'P-1, Component');
	assert.equal(keyText(sheet, { name: 'x' }), '');
	assert.equal(keyText({ kind: 'main', keyColumns: ['flag'] }, { flag: true }), 'Yes');
});

test('keyText names parent, related item and other match-key properties on a relationship sheet', () => {
	const result = fixture();
	const bom = result.sheets[1];
	const cells = { 'parent.item_number': 'TA-001', 'related.item_number': 'TP-002', sort_order: 10, quantity: 2 };
	assert.equal(keyText(bom, cells), 'TA-001 / TP-002 (Sequence 10)');
	assert.equal(keyText(bom, { 'parent.item_number': 'TA-001' }), 'TA-001');
	assert.equal(keyText({ ...bom, keyColumns: ['parent.item_number'], matchKeyColumns: ['parent.item_number'] }, cells), 'TA-001');
});

test('addKeyTexts writes the key text into every row, and rowName falls back on the matched item', () => {
	const result = fixture();
	addKeyTexts(result, workbookOf(result));
	assert.equal(result.sheets[0].rows[0].keyText, 'TA-001');
	assert.equal(result.sheets[0].rows[4].keyText, '', 'a row with a blank key has no key text');
	assert.equal(result.sheets[1].rows[1].keyText, 'TA-001 / TP-002 (Sequence 30)');
	assert.equal(rowName({ keyedName: 'X-1' }), 'X-1');
	assert.equal(rowName({ keyText: 'K', keyedName: 'X-1' }), 'K');
});

test('addKeyTexts leaves a row without key text when the workbook does not line up', () => {
	const result = fixture();
	const workbook = workbookOf(result);
	workbook.sheets[0].rows[1].rowNumber = 99;
	addKeyTexts(result, workbook);
	assert.equal(result.sheets[0].rows[1].keyText, '');
	assert.equal(result.sheets[0].rows[2].keyText, 'P-9001');
});

// ---------------------------------------------------------------------------
// The preview grid
// ---------------------------------------------------------------------------

test('the filters combine: hide unchanged, created or modified only, problems only', () => {
	const main = fixture().sheets[0];
	const actions = (filter) => filterRows(main, { ...DEFAULT_FILTER, ...filter }).map((index) => main.rows[index].row);
	assert.deepEqual(actions({}), [2, 3, 4, 5, 6]);
	assert.deepEqual(actions({ hideUnchanged: true }), [3, 4, 5, 6]);
	assert.deepEqual(actions({ show: 'new' }), [4]);
	assert.deepEqual(actions({ show: 'update' }), [3]);
	assert.deepEqual(actions({ problemsOnly: true }), [2, 3, 5, 6]);
	assert.deepEqual(actions({ problemsOnly: true, hideUnchanged: true }), [3, 5, 6]);
	assert.deepEqual(actions({ problemsOnly: true, show: 'new' }), []);
	assert.equal(rowPasses({ action: 'none' }, DEFAULT_FILTER), true);
});

test('filterRows stays fast on a sheet of 5,000 rows', () => {
	const rows = Array.from({ length: 5000 }, (_, i) => ({ row: i + 2, action: ['new', 'update', 'none', 'error'][i % 4], issues: i % 7 === 0 ? [{ severity: 'warning' }] : [] }));
	const started = process.hrtime.bigint();
	for (let i = 0; i < 20; i++) {
		filterRows({ rows }, { hideUnchanged: true, show: 'all', problemsOnly: i % 2 === 0 });
	}
	const ms = Number(process.hrtime.bigint() - started) / 1e6;
	assert.ok(ms < 200, 'twenty filter runs took ' + ms + ' ms');
});

test('valueColumns lists only the columns some row writes, without key and ignored columns', () => {
	const result = fixture();
	assert.deepEqual(valueColumns(result.sheets[0]).map((column) => column.id),
		['owned_by_id', 'name', 'classification', 'unit', 'make_buy', 'effective_date']);
	assert.deepEqual(valueColumns(result.sheets[1]).map((column) => column.id), ['quantity', 'reference_designator', 'sort_order']);
	const changedRelated = { ...result.sheets[1], rows: [{ action: 'update', values: { related_id: 'A' } }] };
	assert.deepEqual(valueColumns(changedRelated), [{ id: 'related_id', label: 'Related Part', dataType: 'item', pattern: null }]);
	const ignored = { ...result.sheets[0], rows: [{ action: 'new', values: { legacy_code: 'x', name: 'n' } }] };
	assert.deepEqual(valueColumns(ignored).map((column) => column.id), ['name']);
});

test('keyHeader names the key column', () => {
	const result = fixture();
	assert.equal(keyHeader(result.sheets[0]), 'Part Number');
	assert.equal(keyHeader(result.sheets[1]), 'Parent / Part');
	assert.equal(keyHeader({ kind: 'main', createOnly: true, keyColumns: [] }), 'Item');
});

test('showValue shows labels, empty values, booleans and dates', () => {
	assert.equal(showValue('ID1', { dataType: 'item' }, 'Administrators'), 'Administrators');
	assert.equal(showValue('', { dataType: 'string' }), '(empty)');
	assert.equal(showValue('1', { dataType: 'boolean' }), 'Yes');
	assert.equal(showValue('0', { dataType: 'boolean' }), 'No');
	assert.equal(showValue('2026-12-01T00:00:00', { dataType: 'date', pattern: 'short_date_time' }), '2026-12-01');
	assert.equal(showValue('2026-12-01T14:30:00', { dataType: 'date', pattern: 'short_date_time' }), '2026-12-01 14:30');
	assert.equal(showValue('2026-12-01T14:30:00', { dataType: 'date', pattern: 'short_date' }), '2026-12-01');
	assert.equal(showValue('5', { dataType: 'float' }), '5');
});

test('cellOf gives new and old values of an update row, and nothing for a column the row leaves alone', () => {
	const main = fixture().sheets[0];
	const update = main.rows[1];
	const column = (id) => main.columns.find((candidate) => candidate.id === id);
	assert.deepEqual(cellOf(update, column('owned_by_id')), { text: 'Administrators', old: '(empty)' });
	assert.deepEqual(cellOf(update, column('make_buy')), { text: 'Buy', old: 'Make' });
	assert.equal(cellOf(update, column('unit')), null);
	assert.deepEqual(cellOf(main.rows[2], column('unit')), { text: 'EA', old: null });
	assert.equal(cellOf(main.rows[3], column('name')), null, 'an error row writes nothing');
});

test('rowMessages puts errors before warnings', () => {
	const row = { issues: [{ severity: 'warning', message: 'w' }, { severity: 'error', message: 'e' }] };
	assert.deepEqual(rowMessages(row).map((issue) => issue.message), ['e', 'w']);
	assert.deepEqual(rowMessages({}), []);
});

test('actionCounts and importCounts count the rows of each action', () => {
	const result = fixture();
	assert.deepEqual(actionCounts(result.sheets[0]), { rows: 5, new: 1, update: 1, delete: 0, none: 1, error: 2, warnings: 2 });
	const counts = importCounts(result);
	assert.equal(counts.total, 12);
	assert.equal(counts.create, 4);
	assert.equal(counts.update, 2);
	assert.equal(counts.unchanged, 2);
	assert.equal(counts.skipped, 4);
	assert.deepEqual(counts.sheets.map((sheet) => sheet.sheetName), ['Part', 'BOM']);
	assert.match(summaryLine(counts), /4 rows create .*2 rows change .*2 rows are unchanged.*4 rows have errors and are skipped/);
	assert.doesNotMatch(summaryLine({ ...counts, skipped: 0 }), /errors/);
	assert.equal(ACTION_LABELS.none, 'No change');
});

// ---------------------------------------------------------------------------
// The import plan
// ---------------------------------------------------------------------------

test('planImport keeps relationship rows with their main-sheet parent and sends an unchanged parent along', () => {
	const { units } = planImport(fixture());
	// Main rows: 2 none (parent of BOM rows 3, 4, 5), 3 update, 4 new (parent of BOM row 6).
	assert.deepEqual(units.map((unit) => ({ main: unit.main, writesMain: unit.writesMain, rels: unit.rels })), [
		{ main: 0, writesMain: false, rels: [[1, 1], [1, 2], [1, 3]] },
		{ main: 1, writesMain: true, rels: [] },
		{ main: 2, writesMain: true, rels: [[1, 4]] }
	]);
	assert.equal(units.reduce((sum, unit) => sum + unitRowCount(unit), 0), 6, 'every new and update row is written once');
});

test('planImport groups relationship rows of parents that are not on the main sheet by parent', () => {
	const result = {
		sheets: [
			{ sheetName: 'Part', kind: 'main', rows: [{ row: 2, action: 'new' }] },
			{ sheetName: 'BOM', kind: 'relationship', parentSheet: 0, rows: [
				{ row: 2, action: 'new', parent: { row: null, id: 'P1' } },
				{ row: 3, action: 'update', parent: { row: null, id: 'P2' } },
				{ row: 4, action: 'none', parent: { row: null, id: 'P3' } },
				{ row: 5, action: 'new', parent: { row: null, id: 'P1' } },
				{ row: 6, action: 'error', parent: { row: null, id: 'P1' } }
			] },
			{ sheetName: 'Docs', kind: 'relationship', parentSheet: 0, rows: [{ row: 2, action: 'new', parent: { row: null, id: 'P2' } }] }
		]
	};
	const { units } = planImport(result);
	assert.deepEqual(units.map((unit) => [unit.main, unit.parentId, unit.rels]), [
		[0, null, []],
		[null, 'P1', [[1, 0], [1, 3]]],
		[null, 'P2', [[1, 1], [2, 0]]]
	]);
});

test('makeBatches puts at most batch-size units into a batch', () => {
	const units = Array.from({ length: 1201 }, (_, i) => ({ main: i, writesMain: true, parentId: null, rels: i % 2 ? [[1, i]] : [] }));
	const batches = makeBatches(units, 500);
	assert.deepEqual(batches.map((batch) => batch.length), [500, 500, 201]);
	assert.equal(makeBatches(units)[0].length, 50, 'a batch size that is not set means 50');
	batches[0][1].rels.push([1, 9999]);
	assert.equal(units[1].rels.length, 1, 'a batch has copies of the units');
	assert.deepEqual(makeBatches(units.slice(0, 3), 0).map((batch) => batch.length), [3], 'no batch size means the default');
});

test('sliceWorkbook sends every sheet with only the rows of the batch, each with its reviewed digest', () => {
	const result = withDigests(fixture());
	const workbook = workbookOf(result);
	const { units } = planImport(result);
	const slice = sliceWorkbook(workbook, result, [units[0], units[2]]);
	assert.equal(slice.meta, workbook.meta);
	assert.deepEqual(slice.sheets.map((sheet) => sheet.sheetName), ['Part', 'BOM']);
	assert.deepEqual(slice.sheets[0].rows.map((row) => [row.rowNumber, row.digest]), [[2, result.sheets[0].rows[0].digest], [4, result.sheets[0].rows[2].digest]]);
	assert.match(slice.sheets[0].rows[0].digest, /^[0-9a-f]{32}$/);
	assert.deepEqual(slice.sheets[1].rows.map((row) => row.rowNumber), [3, 4, 5, 6]);
	assert.equal(slice.sheets[1].rows[0].cells['parent.item_number'], 'TA-001');
	assert.deepEqual(slice.sheets[0].columns, workbook.sheets[0].columns);
	assert.deepEqual(batchRows([units[2]], 2), [[2], [4]]);
});

test('sliceWorkbook refuses a workbook whose rows moved since the preview', () => {
	const result = withDigests(fixture());
	const workbook = workbookOf(result);
	workbook.sheets[0].rows[2].rowNumber = 40;
	assert.throws(() => sliceWorkbook(workbook, result, planImport(result).units), /Upload the workbook again/);
	assert.match(workbookMatchesResult(workbook, result), /does not match/);
	assert.equal(workbookMatchesResult(workbookOf(result), result), null);
	assert.match(workbookMatchesResult({ sheets: [] }, result), /other sheets/);
});

test('dropFailedWrite takes a failed parent out with its relationship rows', () => {
	const result = fixture();
	const batch = makeBatches(planImport(result).units, 500)[0];
	const failed = dropFailedWrite(batch, result, { named: { sheetIndex: 0, rowIndex: 2 }, rows: [{ sheetIndex: 0, rowIndex: 2 }, { sheetIndex: 1, rowIndex: 4 }], message: 'Unique violation' });
	assert.deepEqual(failed.map((entry) => [entry.sheetIndex, entry.rowIndex]), [[0, 2], [1, 4]]);
	assert.equal(failed[0].message, 'Unique violation');
	assert.match(failed[1].message, /its parent \(row 4 of Part\) failed: Unique violation/);
	assert.deepEqual(batch.map((unit) => unit.main), [0, 1]);
});

test('dropFailedWrite takes out every row of a failed write of an unchanged parent', () => {
	const result = fixture();
	const batch = makeBatches(planImport(result).units, 500)[0];
	// Main row 2 changes nothing; BOM rows 3, 4 and 5 are one edit of it.
	const failed = dropFailedWrite(batch, result, {
		named: { sheetIndex: 1, rowIndex: 1 }, rows: [{ sheetIndex: 1, rowIndex: 1 }, { sheetIndex: 1, rowIndex: 2 }, { sheetIndex: 1, rowIndex: 3 }], message: 'locked'
	});
	assert.deepEqual(failed.map((entry) => [entry.sheetIndex, entry.rowIndex]), [[1, 1], [1, 2], [1, 3]]);
	assert.equal(failed[0].message, 'locked');
	assert.match(failed[1].message, /written together with row 3 of BOM, which failed: locked/);
	assert.deepEqual(batch.map((unit) => unit.main), [1, 2], 'the unchanged parent is no longer sent');
});

test('dropFailedWrite takes out a single relationship row, and gives null for rows the batch does not hold', () => {
	const result = fixture();
	const batch = makeBatches(planImport(result).units, 500)[0];
	assert.deepEqual(dropFailedWrite(batch, result, { named: { sheetIndex: 1, rowIndex: 1 }, rows: [], message: 'm' }), [{ sheetIndex: 1, rowIndex: 1, message: 'm' }]);
	assert.deepEqual(batch[0].rels, [[1, 2], [1, 3]]);
	assert.equal(dropFailedWrite(batch, result, { named: { sheetIndex: 1, rowIndex: 1 }, rows: [], message: 'm' }), null, 'a row the batch no longer holds');
	assert.equal(dropFailedWrite(batch, result, { named: { sheetIndex: 0, rowIndex: 4 }, rows: [], message: 'm' }), null, 'an error row is never in a batch');
	assert.equal(dropFailedWrite(batch, result, { named: { sheetIndex: 0, rowIndex: 0 }, rows: [], message: 'm' }), null, 'an unchanged parent is not written itself');
});

// New Parts with relationship rows between them: ASM (row 2) uses C1, C2 and C3 (rows 3 to 5),
// C1 and C2 are each other's alternates, row 6 changes an existing Part, and an existing
// parent that is not on the main sheet gets C3 as well.
function newRelatedResult() {
	const existing = 'A'.repeat(32);
	const newRow = (row) => ({ row, action: 'new' });
	const rel = (row, parent, related) => ({ row, action: 'new', parent: typeof parent === 'number' ? { sheet: 0, row: parent, id: null } : { sheet: 0, row: null, id: parent }, related: { sheet: 0, row: related, id: null } });
	return {
		sheets: [
			{ sheetName: 'Part', kind: 'main', rows: [newRow(2), newRow(3), newRow(4), newRow(5), { row: 6, action: 'update', id: existing }] },
			{ sheetName: 'BOM', kind: 'relationship', parentSheet: 0, relatedSheet: 0, rows: [rel(2, 2, 3), rel(3, 2, 4), rel(4, 2, 5), rel(5, 'B'.repeat(32), 5)] },
			{ sheetName: 'Alternates', kind: 'relationship', parentSheet: 0, relatedSheet: 0, rows: [rel(2, 3, 4), rel(3, 4, 3)] }
		]
	};
}
const unitName = (result, unit) => (unit.main === null ? 'parent' : 'row ' + result.sheets[0].rows[unit.main].row);

test('planImport puts the new items that relationship rows point at first, and keeps items that point at each other in one block', () => {
	const result = newRelatedResult();
	const { units } = planImport(result);
	assert.deepEqual(units.map((unit) => unitName(result, unit)), ['row 3', 'row 4', 'row 5', 'row 2', 'row 6', 'parent']);
	assert.deepEqual(units.map((unit) => unit.block), [0, 0, 1, 2, 3, 4]);
	assert.deepEqual(units[0].rels, [[2, 0]], 'C1 keeps its alternate row');
});

test('makeBatches keeps a block in one batch, and splits only a block larger than a batch', () => {
	const result = newRelatedResult();
	const { units } = planImport(result);
	const names = (batches) => batches.map((batch) => batch.map((unit) => unitName(result, unit)));
	assert.deepEqual(names(makeBatches(units, 3)), [['row 3', 'row 4', 'row 5'], ['row 2', 'row 6', 'parent']]);
	assert.deepEqual(names(makeBatches(units.slice(2, 3).concat(units.slice(0, 2)), 2)), [['row 5'], ['row 3', 'row 4']], 'the block moves to the next batch whole');
	assert.deepEqual(names(makeBatches(units, 1)).slice(0, 2), [['row 3'], ['row 4']], 'a block larger than a batch is split');
});

test('dropRowsPointingAt takes out the rows whose related item a failed row was to create', () => {
	const result = newRelatedResult();
	const batch = makeBatches(planImport(result).units, 500)[0];
	assert.deepEqual(dropRowsPointingAt(batch, result, []), []);
	const dropped = dropRowsPointingAt(batch, result, [rowKey(0, 3)]);
	assert.deepEqual(dropped.map((entry) => [entry.sheetIndex, entry.rowIndex]), [[2, 1], [1, 0]]);
	assert.equal(dropped[0].message, 'Not imported, because its related item (row 3 of Part) was not imported.');
	assert.deepEqual(batch.find((unit) => unit.main === 0).rels, [[1, 1], [1, 2]], 'the assembly keeps its other BOM rows');
});

test('dropFailedWrite also takes out the rows that point at the item of a failed new row', () => {
	const result = newRelatedResult();
	const batch = makeBatches(planImport(result).units, 500)[0];
	const failed = dropFailedWrite(batch, result, { named: { sheetIndex: 0, rowIndex: 1 }, rows: [{ sheetIndex: 0, rowIndex: 1 }, { sheetIndex: 2, rowIndex: 0 }], message: 'Unique violation' });
	assert.deepEqual(failed.map((entry) => [entry.sheetIndex, entry.rowIndex]).sort(), [[0, 1], [1, 0], [2, 0], [2, 1]]);
	assert.match(failed.find((entry) => entry.sheetIndex === 1).message, /related item \(row 3 of Part\)/);
});

// Two targets: Part (sheet 0) and CAD Document (sheet 1). Part CAD (sheet 2) links Parts to CAD
// Documents, Part BOM (sheet 3) links Parts, CAD Structure (sheet 4) links CAD Documents.
function twoTargetResult() {
	const newRow = (row) => ({ row, action: 'new' });
	return {
		sheets: [
			{ sheetName: 'Part', kind: 'main', rows: [newRow(2), newRow(3)] },
			{ sheetName: 'CAD Document', kind: 'main', rows: [newRow(2), newRow(3), { row: 4, action: 'error' }] },
			{ sheetName: 'Part CAD', kind: 'relationship', parentSheet: 0, relatedSheet: 1, rows: [
				{ row: 2, action: 'new', parent: { sheet: 0, row: 2, id: null }, related: { sheet: 1, row: 3, id: null } },
				{ row: 3, action: 'new', parent: { sheet: 0, row: 3, id: null }, related: { sheet: 1, row: 2, id: null } }
			] },
			{ sheetName: 'Part BOM', kind: 'relationship', parentSheet: 0, relatedSheet: 0, rows: [
				{ row: 2, action: 'new', parent: { sheet: 0, row: 2, id: null }, related: { sheet: 0, row: 3, id: null } }
			] },
			{ sheetName: 'CAD Structure', kind: 'relationship', parentSheet: 1, relatedSheet: 1, rows: [
				// Row 2 of the CAD sheet, not row 2 of the Part sheet.
				{ row: 2, action: 'new', parent: { sheet: 1, row: 2, id: null }, related: { sheet: 1, row: 3, id: null } },
				{ row: 3, action: 'new', parent: { sheet: 1, row: null, id: 'C'.repeat(32) }, related: { sheet: 1, row: 2, id: null } }
			] }
		]
	};
}
const sheetUnitName = (result, unit) => (unit.main === null ? 'parent' : result.sheets[unit.mainSheet].sheetName + ' ' + result.sheets[unit.mainSheet].rows[unit.main].row);

test('mainSheetIndexes, parentSheetOf and newRelatedItem read the main sheets a result names', () => {
	const result = twoTargetResult();
	assert.deepEqual(mainSheetIndexes(result), [0, 1]);
	assert.equal(isMainSheet(result, 1), true);
	assert.equal(isMainSheet(result, 2), false);
	assert.equal(parentSheetOf(result, 4), 1);
	assert.equal(parentSheetOf(result, 2), 0);
	assert.deepEqual(newRelatedItem(result.sheets[2].rows[0]), { sheet: 1, row: 3 });
	assert.equal(newRelatedItem({ related: { row: 5, id: 'X' } }), null);
});

test('planImport with two main sheets: units by sheet and row, and the CAD Documents a Part points at go first', () => {
	const result = twoTargetResult();
	const { units } = planImport(result);
	assert.deepEqual(units.map((unit) => sheetUnitName(result, unit)), ['CAD Document 3', 'CAD Document 2', 'Part 3', 'Part 2', 'parent']);
	const cad2 = units.find((unit) => unit.mainSheet === 1 && unit.main === 0);
	assert.deepEqual(cad2.rels, [[4, 0]], 'the CAD Structure row of CAD row 2 is with it, not with Part row 2');
	const part2 = units.find((unit) => unit.mainSheet === 0 && unit.main === 0);
	assert.deepEqual(part2.rels, [[2, 0], [3, 0]]);
	assert.equal(units.find((unit) => unit.main === null).mainSheet, 1);
	assert.deepEqual(batchRows([cad2, part2], 5), [[0], [0], [0], [0], [0]]);
	assert.deepEqual(batchRows([units[0]], 5), [[], [1], [], [], []]);
});

test('dropFailedWrite and dropRowsPointingAt with two main sheets', () => {
	const result = twoTargetResult();
	const batch = makeBatches(planImport(result).units, 500)[0];
	// CAD row 3 fails: Part CAD row 2 and CAD Structure row 2 point at it.
	const failed = dropFailedWrite(batch, result, { named: { sheetIndex: 1, rowIndex: 1 }, rows: [], message: 'Unique violation' });
	assert.deepEqual(failed.map((entry) => [entry.sheetIndex, entry.rowIndex]).sort(), [[1, 1], [2, 0], [4, 0]]);
	assert.match(failed.find((entry) => entry.sheetIndex === 2).message, /related item \(row 3 of CAD Document\)/);
	assert.equal(batch.some((unit) => unit.mainSheet === 1 && unit.main === 1), false);
	assert.equal(dropFailedWrite(batch, result, { named: { sheetIndex: 0, rowIndex: 1 }, rows: [], message: 'm' })[0].sheetIndex, 0, 'a Part row with the same index');
	const again = makeBatches(planImport(result).units, 500)[0];
	assert.deepEqual(dropRowsPointingAt(again, result, [rowKey(0, 2)]), [], 'row 2 of the Part sheet creates no CAD Document');
	assert.deepEqual(dropRowsPointingAt(again, result, [rowKey(1, 2)]).map((entry) => [entry.sheetIndex, entry.rowIndex]).sort(), [[2, 1], [4, 1]]);
});

test('batchDependencies: a batch waits for the earlier batches that create the items its rows point at', () => {
	const two = twoTargetResult();
	const twoBatches = makeBatches(planImport(two).units, 1);
	assert.deepEqual(twoBatches.map((batch) => sheetUnitName(two, batch[0])), ['CAD Document 3', 'CAD Document 2', 'Part 3', 'Part 2', 'parent']);
	assert.deepEqual(batchDependencies(twoBatches, two), [[], [0], [1], [0, 2], [1]]);
	assert.deepEqual(batchDependencies(makeBatches(planImport(two).units, 500), two), [[]], 'one batch waits for none');
	const one = newRelatedResult();
	// Rows 3 and 4 are each other's alternates; with a batch of one item, row 3 points at an item a later batch creates.
	assert.deepEqual(batchDependencies(makeBatches(planImport(one).units, 1), one), [[], [0], [], [0, 1, 2], [], [2]]);
});

test('batchDependencies: batches that touch the same existing item or file go one after the other', () => {
	const item = (letter) => letter.repeat(32);
	const token = 'file:' + 'F'.repeat(32) + ':spec.pdf';
	const link = (row, parent, related) => ({ row, action: 'new', parent: { sheet: 0, row: null, id: item(parent) }, related });
	const result = {
		sheets: [
			{ sheetName: 'Document', kind: 'main', rows: [
				{ row: 2, action: 'update', id: item('A') },
				{ row: 3, action: 'update', id: item('B') },
				{ row: 4, action: 'new', values: { attachment: token } }
			] },
			{ sheetName: 'Links', kind: 'relationship', parentSheet: 0, rows: [
				link(2, 'C', { id: item('A') }),
				link(3, 'D', { id: item('A') }),
				link(4, 'E', { id: null, file: token }),
				link(5, '1', { id: item('2') }),
				link(6, 'B', { id: item('3') })
			] }
		]
	};
	const batches = makeBatches(planImport(result).units, 1);
	assert.deepEqual(batchDependencies(batches, result), [
		[], // changes A
		[], // changes B
		[], // stores the file
		[0], // links to A, which batch 0 changes
		[3], // links to A too: after the last batch to touch it
		[2], // links the file of batch 2
		[], // touches nothing another batch does
		[1] // has B as the parent of its row, which batch 1 changes
	]);
	assert.deepEqual(batchDependencies(makeBatches(planImport(result).units, 500), result), [[]], 'one batch waits for none');
});

test('runScheduled runs tasks at most limit at once, each after the tasks it waits for', async () => {
	const log = [];
	let running = 0;
	let most = 0;
	const run = async (index) => {
		running++;
		most = Math.max(most, running);
		log.push('start ' + index);
		await new Promise((resolve) => setTimeout(resolve, index === 0 ? 30 : 5));
		log.push('end ' + index);
		running--;
	};
	await runScheduled(5, [[], [], [0], [], [2]], 2, run);
	assert.equal(most, 2);
	assert.ok(log.indexOf('start 2') > log.indexOf('end 0'), 'task 2 waits for task 0');
	assert.ok(log.indexOf('start 4') > log.indexOf('end 2'));
	assert.ok(log.indexOf('start 3') < log.indexOf('end 0'), 'task 3 does not wait for task 0');
	await runScheduled(0, [], 2, run);
	const order = [];
	await runScheduled(3, [[], [], []], 1, async (index) => { order.push(index); });
	assert.deepEqual(order, [0, 1, 2], 'one at a time: in order');
});

test('runScheduled starts no task after one fails, and fails with the first error once the running ones are done', async () => {
	const started = [];
	const ended = [];
	const run = async (index) => {
		started.push(index);
		await new Promise((resolve) => setTimeout(resolve, index === 1 ? 20 : 5));
		if (index === 0) {
			throw new Error('task 0 failed');
		}
		ended.push(index);
	};
	await assert.rejects(runScheduled(4, [[], [], [], []], 2, run), /task 0 failed/);
	assert.deepEqual(started, [0, 1]);
	assert.deepEqual(ended, [1], 'the running task is done before the run fails');
	await assert.rejects(runScheduled(1, [[0]], 2, run), /wait for each other/);
});

test('isRetryFault and retryDelay', () => {
	assert.equal(isRetryFault({ code: 'MCE_RETRY', message: 'The batch met another request' }), true);
	assert.equal(isRetryFault({ code: 'SOAP-ENV:Server', message: 'Transaction (Process ID 61) was deadlocked on lock resources' }), true);
	assert.equal(isRetryFault({ code: 'MCE_ROW_FAILED', message: 'Unique violation' }), false);
	assert.equal(isRetryFault(null), false);
	assert.equal(retryDelay(1, () => 0.5), 500);
	assert.equal(retryDelay(3, () => 0.5), 2000);
	assert.equal(retryDelay(1, () => 0), 375);
});

test('a main-sheet row that failed as changed may have been imported already: rows that point at it are kept', () => {
	const result = twoTargetResult();
	assert.equal(failedRowLeavesNoItem(result, { sheetIndex: 1, rowIndex: 0, reason: 'changed' }), false);
	assert.equal(failedRowLeavesNoItem(result, { sheetIndex: 1, rowIndex: 0, reason: 'error' }), true);
	assert.equal(failedRowLeavesNoItem(result, { sheetIndex: 0, rowIndex: 0 }), true, 'a failed write has no reason');
	assert.equal(failedRowLeavesNoItem(result, { sheetIndex: 2, rowIndex: 0, reason: 'error' }), false, 'a relationship row creates no item');
	const read = readBatchAnswer({ rows: [{ sheet: 1, row: 2, status: 'failed', reason: 'changed' }] }, result);
	assert.deepEqual(read.failed.map((entry) => [entry.sheetIndex, entry.rowIndex, entry.reason]), [[1, 0, 'changed']]);
});

test('resultTargets and recordCountsByTarget count the records of each target, its relationship sheets included', () => {
	const result = twoTargetResult();
	const targets = resultTargets(result);
	assert.deepEqual(targets, [
		{ label: 'Part', sheets: ['Part', 'Part CAD', 'Part BOM'] },
		{ label: 'CAD Document', sheets: ['CAD Document', 'CAD Structure'] }
	]);
	const records = [
		{ sheet: 'Part', action: 'Created' }, { sheet: 'Part CAD', action: 'Created' }, { sheet: 'Part BOM', action: 'Failed' },
		{ sheet: 'CAD Document', action: 'Created' }, { sheet: 'CAD Document', action: 'Skipped' }, { sheet: 'CAD Structure', action: 'Updated' }
	];
	const counts = recordCountsByTarget(records, targets);
	assert.deepEqual(counts.map((c) => [c.label, c.rows, c.Created, c.Updated, c.Failed, c.Skipped]), [['Part', 3, 2, 0, 1, 0], ['CAD Document', 3, 1, 1, 0, 1]]);
	assert.equal(targetCountsLine(counts[0]), 'Part: 2 created, 0 updated, 1 failed');
	assert.equal(targetCountsLine(counts[1]), 'CAD Document: 1 created, 1 updated, 0 failed, 1 skipped because of errors');
});

test('sheetGroups, sheetOrder and sheetTargets give the order of the workbook tabs: each main sheet, then its relationship sheets', () => {
	const result = twoTargetResult();
	assert.deepEqual(sheetGroups(result), [[0, 2, 3], [1, 4]]);
	assert.deepEqual(sheetOrder(result), [0, 2, 3, 1, 4]);
	assert.deepEqual(sheetTargets(result), [0, 1, 0, 0, 1]);
	assert.deepEqual(sheetGroups({ sheets: [] }), []);
	const counts = importCounts(result);
	assert.deepEqual(counts.sheets.map((sheet) => [sheet.label, sheet.target]),
		[['Part', 0], ['Part CAD', 0], ['Part BOM', 0], ['CAD Document', 1], ['CAD Structure', 1]]);
	assert.equal(counts.create, 9);
});

test('oversizedBlocks and oversizedWarning find groups that point at each other and are larger than a batch', () => {
	const result = newRelatedResult();
	const { units } = planImport(result);
	assert.deepEqual(oversizedBlocks(units, 500), []);
	assert.equal(oversizedWarning(result, 500), null);
	assert.deepEqual(oversizedBlocks(units, 1), [{ units: 2, rows: 4, first: { sheetIndex: 0, rowIndex: 1 } }]);
	const warning = oversizedWarning(result, 1);
	assert.match(warning, /^A group of 2 items \(from row 3 of Part\) that point at each other's new items is larger than one batch of 1 items\./);
	assert.match(warning, /later batch creates fails on import/);
});

test('jsonObjectsIn finds the JSON in a text', () => {
	assert.deepEqual(jsonObjectsIn('{"a":1}'), [{ a: 1 }]);
	assert.deepEqual(jsonObjectsIn('Error {not json} then {"sheet":0,"message":"a } in text","rows":[{"row":2}]} end'),
		[{ sheet: 0, message: 'a } in text', rows: [{ row: 2 }] }]);
	assert.deepEqual(jsonObjectsIn('no json'), []);
	assert.deepEqual(jsonObjectsIn(undefined), []);
	assert.deepEqual(jsonObjectsIn('{"open": '), []);
});

test('writeFailureOf reads the failed write a refused batch names', () => {
	const result = fixture();
	const detail = JSON.stringify({ sheet: 0, sheetName: 'Part', row: 4, message: 'Aras says no', rows: [{ sheet: 0, row: 4 }, { sheet: 1, row: 6 }, { sheet: 1, row: 99 }] });
	assert.deepEqual(writeFailureOf({ code: 'MCE_ROW_FAILED', message: 'Row 4 of sheet \'Part\' could not be imported: Aras says no', detail }, result), {
		named: { sheetIndex: 0, rowIndex: 2 }, rows: [{ sheetIndex: 0, rowIndex: 2 }, { sheetIndex: 1, rowIndex: 4 }], message: 'Aras says no'
	});
	const byName = writeFailureOf({ code: 'x', message: 'm', detail: 'before ' + JSON.stringify({ sheetName: 'BOM', row: 3 }) }, result);
	assert.deepEqual(byName, { named: { sheetIndex: 1, rowIndex: 1 }, rows: [], message: 'm' });
	assert.equal(writeFailureOf({ code: '1', message: 'The job is locked.', detail: '' }, result), null);
	assert.equal(writeFailureOf({ code: 'x', message: 'm', detail: JSON.stringify({ sheet: 0, sheetName: 'Other', row: 4 }) }, result), null, 'index and name disagree');
	assert.equal(writeFailureOf(null, result), null);
});

test('readBatchAnswer counts written rows and names the failed ones', () => {
	const result = fixture();
	const read = readBatchAnswer({ rows: [
		{ sheet: 0, row: 3, status: 'updated', id: 'X' },
		{ sheet: 0, row: 4, status: 'created' },
		{ sheet: 0, row: 2, status: 'unchanged' },
		{ sheet: 1, row: 3, status: 'failed', reason: 'changed', message: 'Changed since the preview.' },
		{ sheet: 1, row: 4, status: 'failed', reason: 'parent' },
		{ sheet: 1, row: 77, status: 'failed', reason: 'error', message: 'gone' }
	] }, result);
	assert.equal(read.created, 1);
	assert.equal(read.updated, 1);
	assert.equal(read.unchanged, 1);
	assert.deepEqual(read.failed, [
		{ sheetIndex: 1, rowIndex: 1, message: 'Changed since the preview.', reason: 'changed' },
		{ sheetIndex: 1, rowIndex: 2, message: 'Not imported, because its parent row failed.', reason: 'parent' }
	]);
	assert.deepEqual(read.unknown, [{ sheet: 1, row: 77, message: 'gone' }]);
	assert.deepEqual(readBatchAnswer({}, result), { created: 0, updated: 0, deleted: 0, unchanged: 0, failed: [], unknown: [] });
});

test('findRow and failedRowEntry name a row by sheet and Excel row number', () => {
	const result = addKeyTexts(fixture(), workbookOf(fixture()));
	assert.deepEqual(findRow(result, 'BOM', 6), { sheetIndex: 1, rowIndex: 4 });
	assert.deepEqual(findRow(result, 1, '6'), { sheetIndex: 1, rowIndex: 4 });
	assert.equal(findRow(result, 'BOM', 60), null);
	assert.equal(findRow(result, 'Nope', 2), null);
	assert.equal(findRow(result, 5, 2), null);
	assert.deepEqual(failedRowEntry(result, { sheetIndex: 0, rowIndex: 1, message: 'm' }), { sheet: 0, sheetName: 'Part', row: 3, key: 'TP-001', message: 'm' });
});

test('progressText and resultsFileName', () => {
	assert.equal(progressText({ batch: 2, batches: 3, written: 700, failed: 0, toWrite: 1200 }), 'Batch 2 of 3: 700 of 1200 rows.');
	assert.equal(progressText({ batch: 4, batches: 3, written: 1190, failed: 10, toWrite: 1200 }), 'Batch 3 of 3: 1200 of 1200 rows, 10 failed.');
	assert.equal(progressText({ batch: 3, running: [3], batches: 5, written: 10, failed: 0, toWrite: 50 }), 'Batch 3 of 5: 10 of 50 rows.');
	assert.equal(progressText({ batch: 4, running: [4, 2, 3], batches: 5, written: 10, failed: 0, toWrite: 50 }), 'Batches 2, 3 and 4 of 5: 10 of 50 rows.');
	assert.equal(resultsFileName('MCE-000001'), 'MCE-000001 preview.json');
	assert.equal(resultsFileName('A/B'), 'A_B preview.json');
});

test('sortRecords puts the sheets in the order they were first written in, then the rows', () => {
	const records = [
		{ sheet: 'BOM', row: 4, order: 640 }, { sheet: 'Part', row: 4, order: 384 }, { sheet: 'BOM', row: 2, order: 256 },
		{ sheet: 'Part', row: 2, order: 128 }, { sheet: 'Alternates', row: 2, order: 700 }
	];
	assert.deepEqual(sortRecords(records).map((record) => record.sheet + record.row), ['Part2', 'Part4', 'BOM2', 'BOM4', 'Alternates2']);
	assert.equal(records[0].sheet, 'BOM');
});

test('sortRecords keeps the sheet order of the workbook when it is given, also for a sheet without written rows', () => {
	const records = [
		{ sheet: 'BOM', row: 3, order: 0, action: 'Skipped' }, { sheet: 'Part', row: 4, order: 384 },
		{ sheet: 'Alternates', row: 2, order: 0, action: 'Skipped' }, { sheet: 'Part', row: 2, order: 128 }, { sheet: 'Extra', row: 2, order: 900 }
	];
	assert.deepEqual(sortRecords(records, ['Part', 'BOM', 'Alternates']).map((record) => record.sheet + record.row),
		['Part2', 'Part4', 'BOM3', 'Alternates2', 'Extra2']);
});

test('skippedRecords lists the error rows of the validation result, with their error messages', () => {
	const result = fixture();
	const skipped = skippedRecords(result);
	const errorRows = result.sheets.flatMap((sheet) => sheet.rows.filter((row) => row.action === 'error').map((row) => sheet.sheetName + row.row));
	assert.ok(errorRows.length > 0, 'the fixture has error rows');
	assert.deepEqual(skipped.map((record) => record.sheet + record.row), errorRows);
	for (const record of skipped) {
		assert.equal(record.action, 'Skipped');
		assert.equal(record.targetId, '');
		assert.ok(record.message.length > 0, record.sheet + record.row + ' has a message');
	}
	assert.deepEqual(skippedRecords(null), []);
});

// ---------------------------------------------------------------------------
// Files and delete rows
// ---------------------------------------------------------------------------

const fileLogic = await import(new URL('../../1.0.0/Innovator/Client/customer/masscreateedit/mce-import.js', import.meta.url).href);

test('files: a picked file is named by a token of its MD5 and the name at the end of its path', () => {
	const md5 = 'a'.repeat(32);
	assert.equal(fileLogic.fileToken('./Docs\sub/Spec 1.pdf', md5), 'file:' + 'A'.repeat(32) + ':Spec 1.pdf');
	assert.equal(fileLogic.normalPath('/folder/a.pdf'), 'folder/a.pdf');
	assert.equal(fileLogic.fileNameOf('folder/a.pdf'), 'a.pdf');
	assert.ok(fileLogic.isFileToken('file:' + 'B'.repeat(32) + ':x:y.txt'));
	assert.ok(!fileLogic.isFileToken('file:abc:x.txt'));
	assert.equal(fileLogic.tokenName('file:' + 'B'.repeat(32) + ':x:y.txt'), 'x:y.txt');
	assert.equal(showValue('file:' + 'C'.repeat(32) + ':drawing.dwg', { dataType: 'file' }), 'drawing.dwg');
});

test('files: a batch stores the files of the rows it writes, a delete row is written, and the counts name deletes', () => {
	const token = (letter, name) => 'file:' + letter.repeat(32) + ':' + name;
	const result = {
		sheets: [
			{ sheetName: 'CAD', kind: 'main', rows: [
				{ row: 2, action: 'update', id: 'C1', values: { native_file: token('A', 'a.dwg') } },
				{ row: 3, action: 'error', values: { native_file: token('B', 'b.dwg') } },
				{ row: 4, action: 'none' }
			] },
			{ sheetName: 'Files', kind: 'relationship', parentSheet: 0, rows: [
				{ row: 2, action: 'new', parent: { row: 4, id: 'C3' }, related: { id: null, keyedName: 'c.pdf', file: token('C', 'c.pdf') }, values: {} },
				{ row: 3, action: 'delete', id: 'R9', parent: { row: 4, id: 'C3' }, related: { id: 'F1', keyedName: 'old.pdf' } },
				{ row: 4, action: 'update', id: 'R8', parent: { row: 2, id: 'C1' }, related: { id: 'F2', keyedName: 'd.pdf' }, values: { related_id: token('D', 'd.pdf') } }
			] }
		]
	};
	assert.deepEqual(fileLogic.neededFileTokens(result).sort(), [token('A', 'a.dwg'), token('C', 'c.pdf'), token('D', 'd.pdf')]);
	const plan = planImport(result);
	assert.equal(plan.units.reduce((sum, unit) => sum + unitRowCount(unit), 0), 4);
	const batches = makeBatches(plan.units, 1);
	assert.deepEqual(batches.map((batch) => fileLogic.batchFileTokens(batch, result)), [[token('A', 'a.dwg'), token('D', 'd.pdf')], [token('C', 'c.pdf')]]);
	const counts = importCounts(result);
	assert.equal(counts.delete, 1);
	assert.equal(counts.sheets[1].delete, 1);
	assert.match(summaryLine(counts), /1 row deletes a relationship row/);
	assert.ok(rowPasses(result.sheets[1].rows[1], { ...DEFAULT_FILTER, show: 'delete' }));
	assert.ok(!rowPasses(result.sheets[1].rows[0], { ...DEFAULT_FILTER, show: 'delete' }));
	assert.equal(ACTION_LABELS.delete, 'Delete');
	const read = readBatchAnswer({ rows: [{ sheet: 1, row: 3, status: 'deleted' }, { sheet: 1, row: 2, status: 'failed', reason: 'file' }] }, result);
	assert.equal(read.deleted, 1);
	assert.match(read.failed[0].message, /file it links was not stored/);
});
