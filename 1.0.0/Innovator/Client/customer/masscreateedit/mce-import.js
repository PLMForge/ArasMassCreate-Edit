/**
 * Mass Create & Edit - the import step, without the page.
 *
 * Everything the job Form works out for the import, kept free of the browser
 * and of the Aras client so that it can be tested on its own:
 *
 *   - the text of a row's key, read from the workbook, so the preview can name
 *     every row, also the rows that have errors;
 *   - the preview grid: which rows a filter shows, which columns a sheet needs,
 *     and the text of each cell;
 *   - the import plan: which rows are sent, grouped into units that stay
 *     together (a parent with its relationship rows) and into batches;
 *   - the slice of the workbook that one batch sends, and what happens to a
 *     batch when the server names a row that could not be written;
 *   - the counts the confirmation, the progress and the result show;
 *   - the files the user picked with the workbook: their names in the
 *     validation, and which of them each batch stores.
 *
 * The validation result is the JSON of the server Method FRG_MCE_ValidateImport.
 * Its sheets and rows are in the order of the workbook that was validated, so
 * row i of result sheet s is row i of workbook sheet s.
 */

/** The actions of a validated row, as the preview names them. */
export const ACTION_LABELS = { new: 'Create', update: 'Update', delete: 'Delete', none: 'No change', error: 'Error' };

/** The filter the preview starts with. */
export const DEFAULT_FILTER = { hideUnchanged: false, show: 'all', problemsOnly: false };

const isWritten = (row) => !!row && (row.action === 'new' || row.action === 'update' || row.action === 'delete');
const labelOf = (thing) => (thing && (thing.label || thing.name)) || '';
const plural = (count, one, many) => count + ' ' + (count === 1 ? one : many);

// ---------------------------------------------------------------------------
// Keys
// ---------------------------------------------------------------------------

/** A cell value of the workbook as text. */
function cellText(value) {
	if (value === undefined || value === null) {
		return '';
	}
	if (value === true) {
		return 'Yes';
	}
	if (value === false) {
		return 'No';
	}
	return String(value).trim();
}

/**
 * The text that names a row in the preview and in the import records: the
 * values of its key columns. On a relationship sheet it is the parent's key,
 * then the related item's key, then any other match-key property, for example
 * "P-100 / P-200 (Sequence 10)".
 *
 * @param {object} sheet A sheet of the validation result (keyColumns, matchKeyColumns, columns, kind).
 * @param {object} cells The cells of the workbook row, by column id.
 * @returns {string} '' when the row has no key value at all.
 */
export function keyText(sheet, cells) {
	const values = (ids) => ids.map((id) => cellText(cells ? cells[id] : undefined)).filter((text) => text !== '').join(', ');
	const keyColumns = sheet.keyColumns || [];
	if (sheet.kind !== 'relationship') {
		return values(keyColumns);
	}
	const parent = values(keyColumns.filter((id) => id.startsWith('parent.')));
	const related = values(keyColumns.filter((id) => id.startsWith('related.')));
	let text = [parent, related].filter((part) => part !== '').join(' / ');
	const extra = (sheet.matchKeyColumns || [])
		.filter((id) => !keyColumns.includes(id))
		.map((id) => {
			const value = cellText(cells ? cells[id] : undefined);
			if (value === '') {
				return '';
			}
			const column = (sheet.columns || []).find((candidate) => candidate.id === id);
			return (column ? column.label || column.property || id : id) + ' ' + value;
		})
		.filter((part) => part !== '');
	if (extra.length > 0) {
		text += (text === '' ? '' : ' ') + '(' + extra.join(', ') + ')';
	}
	return text;
}

/**
 * Writes the key text of every row into the validation result, as `keyText`, so
 * that the stored result names its rows without the workbook.
 *
 * @param {object} result The validation result; changed in place.
 * @param {object} workbook The workbook that was validated, as parseWorkbook returns it.
 * @returns {object} The result.
 */
export function addKeyTexts(result, workbook) {
	(result.sheets || []).forEach((sheet, sheetIndex) => {
		const rows = workbook && workbook.sheets[sheetIndex] ? workbook.sheets[sheetIndex].rows : [];
		(sheet.rows || []).forEach((row, rowIndex) => {
			const source = rows[rowIndex];
			row.keyText = source && source.rowNumber === row.row ? keyText(sheet, source.cells) : '';
		});
	});
	return result;
}

/**
 * Adds what only the browser knows about an indented sheet to the validation result, so that the stored
 * result has it: the rows that can't be placed in the structure and the notes on the others, the lines a
 * parent has in Aras that the sheet does not list, and an error on every row below a row with an error.
 * A row used in several places fails only when it has failed in all of them.
 *
 * @param {object} result The validation result; changed in place.
 * @param {object} workbook The workbook that was validated, as parseWorkbook returns it.
 * @returns {object} The result.
 */
export function addIndentedResult(result, workbook) {
	(result.sheets || []).forEach((sheet, sheetIndex) => {
		const info = workbook && workbook.sheets[sheetIndex] ? workbook.sheets[sheetIndex].indented : null;
		if (!info) {
			return;
		}
		const rows = sheet.rows || [];
		const warn = (row, code, message) => {
			row.issues = (row.issues || []).concat({ severity: 'warning', code, column: null, message });
		};
		const refuse = (row, code, message) => {
			row.action = 'error';
			row.issues = [{ severity: 'error', code, column: null, message }];
		};
		const byNumber = new Map(rows.map((row) => [row.row, row]));
		if (info.comparisons && info.comparisons.length > 0) {
			sheet.comparisons = info.comparisons;
		}
		for (const row of rows) {
			if (info.problems[row.row]) {
				refuse(row, info.problems[row.row].code, info.problems[row.row].message);
				if (info.problems[row.row].comparison !== undefined) {
					row.comparison = info.problems[row.row].comparison;
				}
			}
			(info.notes[row.row] || []).forEach((message) => warn(row, 'BOM_SAME_LINE', message));
		}
		for (const parent of sheet.parents || []) {
			const first = parent.untouched > 0 ? rows.find((row) => row.parent && row.action !== 'error' &&
				(parent.id ? row.parent.id === parent.id : row.parent.row === parent.row)) : null;
			if (first) {
				warn(first, 'BOM_LINES_NOT_ON_SHEET', 'This parent has ' + parent.untouched + ' BOM line' + (parent.untouched === 1 ? '' : 's') +
					' in Aras that the sheet does not list. They stay as they are.');
			}
		}
		for (let changed = true; changed;) {
			changed = false;
			for (const row of rows) {
				const above = info.parents[row.row];
				if (row.action === 'error' || !above || above.length === 0 || above.some((number) => number === null)) {
					continue;
				}
				if (above.every((number) => byNumber.has(number) && byNumber.get(number).action === 'error')) {
					refuse(row, 'BOM_PARENT_LINE_FAILED', 'The line above it (row ' + above.join(', ') + ') has an error, so this row is not imported.');
					changed = true;
				}
			}
		}
	});
	return result;
}

/** What names a row in the grid: its key text, or else the item it matched. */
export function rowName(row) {
	if (row.keyText) {
		return row.keyText;
	}
	return row.keyedName || '';
}

// ---------------------------------------------------------------------------
// The preview grid
// ---------------------------------------------------------------------------

/**
 * Whether a row passes the filter of the preview.
 *
 * @param {object} row A row of the validation result.
 * @param {{hideUnchanged: boolean, show: 'all'|'new'|'update'|'delete', problemsOnly: boolean}} filter
 */
export function rowPasses(row, filter) {
	if (filter.hideUnchanged && row.action === 'none') {
		return false;
	}
	if (filter.show === 'new' && row.action !== 'new') {
		return false;
	}
	if (filter.show === 'update' && row.action !== 'update') {
		return false;
	}
	if (filter.show === 'delete' && row.action !== 'delete') {
		return false;
	}
	if (filter.problemsOnly && !(row.issues && row.issues.length > 0)) {
		return false;
	}
	return true;
}

/** The indexes of the rows of a sheet that pass the filter, in sheet order. */
export function filterRows(sheet, filter) {
	const shown = [];
	const rows = sheet.rows || [];
	for (let index = 0; index < rows.length; index++) {
		if (rowPasses(rows[index], filter)) {
			shown.push(index);
		}
	}
	return shown;
}

/**
 * The value columns the grid shows for a sheet: the columns that hold a value
 * to write in at least one row. Key columns are left out (the key has a column
 * of its own), and so are columns whose values are ignored.
 *
 * @returns {Array<{id: string, label: string, dataType: string, pattern: (string|null)}>}
 */
export function valueColumns(sheet) {
	const used = new Set();
	for (const row of sheet.rows || []) {
		if (row.values) {
			for (const id of Object.keys(row.values)) {
				used.add(id);
			}
		}
	}
	const keyColumns = new Set(sheet.keyColumns || []);
	const columns = (sheet.columns || [])
		.filter((column) => used.has(column.id) && !keyColumns.has(column.id) && column.status !== 'unknown')
		.map((column) => ({ id: column.id, label: column.label || column.property || column.id, dataType: column.dataType, pattern: column.pattern || null }));
	if (used.has('related_id') && !columns.some((column) => column.id === 'related_id')) {
		const related = sheet.relationshipType && sheet.relationshipType.relatedItemType;
		columns.push({ id: 'related_id', label: 'Related ' + (labelOf(related) || 'item'), dataType: 'item', pattern: null });
	}
	return columns;
}

/** The header of the key column of a sheet. */
export function keyHeader(sheet) {
	if (sheet.createOnly && sheet.kind !== 'relationship') {
		return 'Item';
	}
	if (sheet.kind === 'relationship') {
		const related = sheet.relationshipType && sheet.relationshipType.relatedItemType;
		return related ? 'Parent / ' + labelOf(related) : 'Parent';
	}
	const labels = (sheet.keyColumns || []).map((id) => {
		const column = (sheet.columns || []).find((candidate) => candidate.id === id);
		return column ? column.label || column.property : id;
	});
	return labels.join(', ') || 'Key';
}

const DATE_ONLY_PATTERNS = ['short_date', 'long_date'];

/**
 * A stored value as the grid shows it.
 *
 * @param {string} value The stored value, as the validation result has it.
 * @param {{dataType: string, pattern: (string|null)}} column
 * @param {string} [display] The display text the result gives for the value, if any.
 */
export function showValue(value, column, display) {
	if (display !== undefined && display !== null && display !== '') {
		return String(display);
	}
	if (value === undefined || value === null || value === '') {
		return '(empty)';
	}
	const text = String(value);
	if (isFileToken(text)) {
		return tokenName(text);
	}
	if (column && column.dataType === 'boolean') {
		return text === '1' ? 'Yes' : text === '0' ? 'No' : text;
	}
	if (column && column.dataType === 'date') {
		const match = /^(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2})(:\d{2})?$/.exec(text);
		if (match) {
			if (DATE_ONLY_PATTERNS.includes(column.pattern) || (match[2] === '00:00' && (match[3] || ':00') === ':00')) {
				return match[1];
			}
			return match[1] + ' ' + match[2];
		}
	}
	return text;
}

/**
 * What one cell of the grid shows.
 *
 * @returns {{text: string, old: (string|null)}|null} null when the row writes
 *   nothing to the column; `old` is the current Aras value of an update row.
 */
export function cellOf(row, column) {
	if (!row.values || !Object.prototype.hasOwnProperty.call(row.values, column.id)) {
		return null;
	}
	const text = showValue(row.values[column.id], column, row.display ? row.display[column.id] : undefined);
	if (row.action !== 'update' || !row.old || !Object.prototype.hasOwnProperty.call(row.old, column.id)) {
		return { text, old: null };
	}
	return { text, old: showValue(row.old[column.id], column, row.oldDisplay ? row.oldDisplay[column.id] : undefined) };
}

/** The messages of a row, errors first. */
export function rowMessages(row) {
	const issues = row.issues || [];
	return issues.filter((issue) => issue.severity === 'error').concat(issues.filter((issue) => issue.severity !== 'error'));
}

/** How many rows of a sheet have each action. */
export function actionCounts(sheet) {
	const counts = { rows: 0, new: 0, update: 0, delete: 0, none: 0, error: 0, warnings: 0 };
	for (const row of sheet.rows || []) {
		counts.rows++;
		if (Object.prototype.hasOwnProperty.call(counts, row.action)) {
			counts[row.action]++;
		}
		if ((row.issues || []).some((issue) => issue.severity === 'warning')) {
			counts.warnings++;
		}
	}
	return counts;
}

// ---------------------------------------------------------------------------
// The import plan
// ---------------------------------------------------------------------------

/**
 * Groups the rows to import into units. A unit is what must go into one request:
 *
 *   - a main-sheet row that creates or changes an item, with the relationship
 *     rows under it that create or change a row;
 *   - a main-sheet row that changes nothing, when relationship rows under it do:
 *     the row is sent along so that the server finds the parent in the slice as
 *     it did in the preview, and it is not written;
 *   - the relationship rows of an existing parent that is not on the main sheet,
 *     grouped by that parent.
 *
 * Rows with action "none" or "error" are never written. A relationship
 * row whose parent row has an error is an error row itself.
 *
 * A workbook has a main sheet for each target; a relationship sheet names the
 * main sheet of its parents, and so does each of its rows (parent.sheet). A result
 * that names none has one main sheet, the first.
 *
 * @param {object} result The validation result.
 * @returns {{units: Array<{main: (number|null), mainSheet: number, writesMain: boolean, parentId: (string|null),
 *   rels: Array<[number, number]>}>, counts: object}}
 *   `main` is the index of the row on main sheet `mainSheet` (an index in result.sheets),
 *   `rels` holds [sheet index, row index] pairs.
 */
export function planImport(result) {
	const sheets = result.sheets || [];
	const mains = mainSheetIndexes(result);
	const unitOfMainRow = new Map();
	const rowIndexOfNumber = new Map();
	for (const mainSheet of mains) {
		(sheets[mainSheet].rows || []).forEach((row, index) => {
			rowIndexOfNumber.set(rowKey(mainSheet, row.row), index);
			if (isWritten(row)) {
				unitOfMainRow.set(rowKey(mainSheet, index), { main: index, mainSheet, writesMain: true, parentId: row.id || null, rels: [] });
			}
		});
	}
	const byParentId = new Map();
	sheets.forEach((sheet, sheetIndex) => {
		if (mains.includes(sheetIndex)) {
			return;
		}
		(sheet.rows || []).forEach((row, rowIndex) => {
			if (!isWritten(row)) {
				return;
			}
			const parent = row.parent || {};
			const mainSheet = parentSheetOf(result, sheetIndex);
			const parentKey = rowKey(mainSheet, parent.row);
			if (parent.row !== null && parent.row !== undefined && rowIndexOfNumber.has(parentKey)) {
				const mainIndex = rowIndexOfNumber.get(parentKey);
				let unit = unitOfMainRow.get(rowKey(mainSheet, mainIndex));
				if (!unit) {
					unit = { main: mainIndex, mainSheet, writesMain: false, parentId: sheets[mainSheet].rows[mainIndex].id || null, rels: [] };
					unitOfMainRow.set(rowKey(mainSheet, mainIndex), unit);
				}
				unit.rels.push([sheetIndex, rowIndex]);
				return;
			}
			const parentId = parent.id || '';
			const idKey = mainSheet + '|' + parentId;
			let unit = byParentId.get(idKey);
			if (!unit) {
				unit = { main: null, mainSheet, writesMain: false, parentId: parentId || null, rels: [] };
				byParentId.set(idKey, unit);
			}
			unit.rels.push([sheetIndex, rowIndex]);
		});
	});
	// Units keep the order of the workbook: main-sheet rows first, in sheet order.
	const ordered = Array.from(unitOfMainRow.values()).sort((a, b) => a.mainSheet - b.mainSheet || a.main - b.main)
		.concat(Array.from(byParentId.values()));
	for (const unit of ordered) {
		unit.rels.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
	}
	return { units: orderByRelated(ordered, result), counts: importCounts(result) };
}

/** The indexes in result.sheets of the main sheets. */
export function mainSheetIndexes(result) {
	const indexes = [];
	(result && result.sheets || []).forEach((sheet, index) => {
		if (sheet && sheet.kind === 'main') {
			indexes.push(index);
		}
	});
	return indexes;
}

/** Whether a sheet of the result is a main sheet. */
export function isMainSheet(result, sheetIndex) {
	return mainSheetIndexes(result).includes(sheetIndex);
}

/** A key for a row of a sheet: "sheet:row". */
export function rowKey(sheetIndex, row) {
	return sheetIndex + ':' + row;
}

/** The main sheet of the parent of a relationship row: the one its sheet names. */
export function parentSheetOf(result, sheetIndex) {
	return result.sheets[sheetIndex].parentSheet;
}

/**
 * The main-sheet row that creates the related item of a relationship row, { sheet, row }
 * (sheet: index in result.sheets, row: row number), or null when the related item exists
 * or the row names none (see FRG_MCE_ValidateImport).
 */
export function newRelatedItem(row) {
	const related = row && row.related;
	if (!related || (related.id !== null && related.id !== undefined) || !Number.isInteger(related.row)) {
		return null;
	}
	return { sheet: related.sheet, row: related.row };
}

/**
 * Puts the units in an order the import can write them in. A relationship row can point
 * at an item that a new main-sheet row creates, so the unit of that row goes first: in the
 * same batch or an earlier one. Units that point at each other in a ring (two new items
 * that are each other's alternates, say) get the same `block`, and makeBatches keeps a
 * block in one batch, where the server gives every new item its id before writing.
 * Otherwise the order of the workbook is kept.
 *
 * @param {Array<object>} units The units of planImport, in workbook order.
 * @param {object} result The validation result.
 * @returns {Array<object>} The units, each with a `block` number; units of one block are next to each other.
 */
export function orderByRelated(units, result) {
	const sheets = result.sheets || [];
	const unitOfRow = new Map();
	units.forEach((unit, index) => {
		const mainSheet = unit.mainSheet;
		const mainRows = (sheets[mainSheet] && sheets[mainSheet].rows) || [];
		if (unit.main !== null && unit.writesMain && mainRows[unit.main] && mainRows[unit.main].action === 'new') {
			unitOfRow.set(rowKey(mainSheet, mainRows[unit.main].row), index);
		}
	});
	// needs[i]: the units that must be written no later than unit i.
	const needs = units.map((unit, index) => {
		const targets = new Set();
		for (const [sheetIndex, rowIndex] of unit.rels) {
			const made = newRelatedItem(sheets[sheetIndex].rows[rowIndex]);
			const target = made === null ? undefined : unitOfRow.get(rowKey(made.sheet, made.row));
			if (target !== undefined && target !== index) {
				targets.add(target);
			}
		}
		return Array.from(targets);
	});

	// The rings: strongly connected units (Tarjan, without recursion).
	const blockOf = new Array(units.length).fill(-1);
	const low = new Array(units.length).fill(0);
	const seen = new Array(units.length).fill(-1);
	const onStack = new Array(units.length).fill(false);
	const stack = [];
	let counter = 0;
	let blocks = 0;
	for (let start = 0; start < units.length; start++) {
		if (seen[start] >= 0) {
			continue;
		}
		const path = [[start, 0]];
		seen[start] = low[start] = counter++;
		stack.push(start);
		onStack[start] = true;
		while (path.length > 0) {
			const frame = path[path.length - 1];
			const [node, next] = frame;
			if (next < needs[node].length) {
				frame[1]++;
				const target = needs[node][next];
				if (seen[target] < 0) {
					seen[target] = low[target] = counter++;
					stack.push(target);
					onStack[target] = true;
					path.push([target, 0]);
				} else if (onStack[target]) {
					low[node] = Math.min(low[node], seen[target]);
				}
				continue;
			}
			path.pop();
			if (path.length > 0) {
				const parent = path[path.length - 1][0];
				low[parent] = Math.min(low[parent], low[node]);
			}
			if (low[node] === seen[node]) {
				let member;
				do {
					member = stack.pop();
					onStack[member] = false;
					blockOf[member] = blocks;
				} while (member !== node);
				blocks++;
			}
		}
	}

	// The blocks in an order where every block comes after the blocks it needs, and else
	// in the order of their first unit in the workbook.
	const members = Array.from({ length: blocks }, () => []);
	units.forEach((unit, index) => members[blockOf[index]].push(index));
	const waitingFor = members.map(() => new Set());
	const neededBy = members.map(() => new Set());
	needs.forEach((targets, index) => {
		for (const target of targets) {
			if (blockOf[target] !== blockOf[index]) {
				waitingFor[blockOf[index]].add(blockOf[target]);
				neededBy[blockOf[target]].add(blockOf[index]);
			}
		}
	});
	const ready = members.map((list, block) => block).filter((block) => waitingFor[block].size === 0);
	const firstOf = (block) => members[block][0];
	const ordered = [];
	let number = 0;
	while (ready.length > 0) {
		ready.sort((a, b) => firstOf(a) - firstOf(b));
		const block = ready.shift();
		for (const index of members[block]) {
			ordered.push({ ...units[index], block: number });
		}
		number++;
		for (const other of neededBy[block]) {
			waitingFor[other].delete(block);
			if (waitingFor[other].size === 0) {
				ready.push(other);
			}
		}
	}
	return ordered;
}

/**
 * What an import of the result does, counted over every sheet. The sheets are in the order the
 * workbook shows them (sheetOrder); target is the position of the sheet's target (sheetTargets).
 *
 * @returns {{total: number, create: number, update: number, delete: number, unchanged: number, skipped: number,
 *   sheets: Array<{sheetName: string, label: string, target: number, create: number, update: number, delete: number, unchanged: number, skipped: number}>}}
 */
export function importCounts(result) {
	const counts = { total: 0, create: 0, update: 0, delete: 0, unchanged: 0, skipped: 0, sheets: [] };
	const targets = sheetTargets(result);
	for (const index of sheetOrder(result)) {
		const sheet = result.sheets[index];
		const own = actionCounts(sheet);
		counts.sheets.push({
			sheetName: sheet.sheetName,
			label: sheet.sheetName,
			target: targets[index],
			total: own.rows,
			create: own.new,
			update: own.update,
			delete: own.delete,
			unchanged: own.none,
			skipped: own.error
		});
		counts.total += own.rows;
		counts.create += own.new;
		counts.update += own.update;
		counts.delete += own.delete;
		counts.unchanged += own.none;
		counts.skipped += own.error;
	}
	return counts;
}

/** How many rows a unit writes. */
export function unitRowCount(unit) {
	return (unit.writesMain ? 1 : 0) + unit.rels.length;
}

/**
 * Splits the units into batches of at most `batchSize` units: a main-sheet
 * row with its relationship rows counts as one, and so does a parent that is not
 * on the main sheet with its relationship rows. The units of one block (see
 * orderByRelated) go into one batch; only a block larger than a batch is split,
 * and its rows that point at an item a later batch creates then fail on import.
 */
export function makeBatches(units, batchSize) {
	const size = Math.max(1, Math.floor(Number(batchSize) || 0) || 50);
	const batches = [];
	let current = [];
	const close = () => {
		if (current.length > 0) {
			batches.push(current);
			current = [];
		}
	};
	for (let start = 0; start < units.length;) {
		let end = start + 1;
		while (end < units.length && units[end].block !== undefined && units[end].block === units[start].block) {
			end++;
		}
		const block = units.slice(start, end).map((unit) => ({ ...unit, rels: unit.rels.slice() }));
		if (current.length + block.length > size) {
			close();
		}
		for (const unit of block) {
			if (current.length === size) {
				close();
			}
			current.push(unit);
		}
		start = end;
	}
	close();
	return batches;
}

/**
 * The groups of items that point at each other (blocks of orderByRelated) with more units than a
 * batch holds. makeBatches splits such a group, and its rows that point at an item a later
 * batch creates fail on import.
 *
 * @param {Array<object>} units The units of planImport.
 * @param {number} batchSize Units in a batch.
 * @returns {Array<{units: number, rows: number, first: {sheetIndex: number, rowIndex: number}|null}>}
 *   For each such group: its units, the rows it writes, and its first main-sheet row.
 */
export function oversizedBlocks(units, batchSize) {
	const size = Math.max(1, Math.floor(Number(batchSize) || 0) || 50);
	const found = [];
	for (let start = 0; start < units.length;) {
		let end = start + 1;
		while (end < units.length && units[end].block !== undefined && units[end].block === units[start].block) {
			end++;
		}
		if (end - start > size) {
			const block = units.slice(start, end);
			const firstMain = block.find((unit) => unit.main !== null);
			found.push({
				units: block.length,
				rows: block.reduce((sum, unit) => sum + unitRowCount(unit), 0),
				first: firstMain ? { sheetIndex: firstMain.mainSheet, rowIndex: firstMain.main } : null
			});
		}
		start = end;
	}
	return found;
}

/**
 * The warning of the preview about groups of items that point at each other and do not fit in
 * one batch, or null when there is none.
 *
 * @param {object} result The validation result.
 * @param {number} batchSize Units in a batch.
 */
export function oversizedWarning(result, batchSize) {
	const blocks = oversizedBlocks(planImport(result).units, batchSize);
	if (blocks.length === 0) {
		return null;
	}
	const size = Math.max(1, Math.floor(Number(batchSize) || 0) || 50);
	const where = (block) => {
		const sheet = block.first && result.sheets[block.first.sheetIndex];
		const row = sheet && sheet.rows[block.first.rowIndex];
		return row ? ' (from row ' + row.row + ' of ' + sheet.sheetName + ')' : '';
	};
	const groups = blocks.map((block) => block.units + ' items' + where(block)).join(', ');
	return (blocks.length === 1 ? 'A group of ' + groups + ' that point at each other\'s new items is' : blocks.length +
		' groups of items that point at each other\'s new items are') +
		' larger than one batch of ' + size + ' items' + (blocks.length === 1 ? '' : ': ' + groups) +
		'. A group is written over several batches then, and a row that points at an item a later batch creates fails on import. ' +
		'An administrator can raise the batch size to avoid this.';
}

/**
 * The rows of a batch, per sheet: [sheet index, row index] pairs in workbook order.
 * A main-sheet row that is only sent along is included.
 */
export function batchRows(batch, sheetCount) {
	const perSheet = Array.from({ length: sheetCount }, () => []);
	for (const unit of batch) {
		if (unit.main !== null) {
			perSheet[unit.mainSheet].push(unit.main);
		}
		for (const [sheetIndex, rowIndex] of unit.rels) {
			perSheet[sheetIndex].push(rowIndex);
		}
	}
	return perSheet.map((indexes) => Array.from(new Set(indexes)).sort((a, b) => a - b));
}

/**
 * The slice of the workbook one batch sends: the same JSON shape as the workbook,
 * with every sheet and only the rows of the batch. Each row carries the digest
 * the user reviewed in the preview.
 *
 * @param {object} workbook The workbook as parseWorkbook returns it.
 * @param {object} result The validation result the user reviewed.
 * @param {Array<object>} batch The units of the batch.
 * @returns {object} { meta, sheets: [ { sheetName, columns, rows: [ { rowNumber, cells, modifiedOn, digest } ] } ] }
 */
export function sliceWorkbook(workbook, result, batch) {
	const picked = batchRows(batch, workbook.sheets.length);
	return {
		meta: workbook.meta,
		sheets: workbook.sheets.map((sheet, sheetIndex) => {
			const resultRows = (result.sheets[sheetIndex] && result.sheets[sheetIndex].rows) || [];
			const slice = { sheetName: sheet.sheetName, columns: sheet.columns, rows: [] };
			for (const rowIndex of picked[sheetIndex]) {
				const source = sheet.rows[rowIndex];
				const reviewed = resultRows[rowIndex];
				if (!source || !reviewed || source.rowNumber !== reviewed.row) {
					throw new Error('The workbook does not match the preview: row ' + (reviewed ? reviewed.row : rowIndex + 2) +
						' of ' + sheet.sheetName + ' is not where it was. Upload the workbook again.');
				}
				slice.rows.push({ rowNumber: source.rowNumber, cells: source.cells, modifiedOn: source.modifiedOn, digest: reviewed.digest || '', sequenceIfNew: source.sequenceIfNew });
			}
			return slice;
		})
	};
}

/**
 * Checks that the workbook read again for the import is the one the preview was
 * made from: the same sheets, and every row where the preview has it.
 *
 * @returns {string|null} Why it is not, or null.
 */
export function workbookMatchesResult(workbook, result) {
	const sheets = result.sheets || [];
	if (!workbook || !Array.isArray(workbook.sheets) || workbook.sheets.length !== sheets.length) {
		return 'The uploaded workbook has other sheets than the preview.';
	}
	for (let index = 0; index < sheets.length; index++) {
		const rows = workbook.sheets[index].rows || [];
		const reviewed = sheets[index].rows || [];
		if (workbook.sheets[index].sheetName !== sheets[index].sheetName || rows.length !== reviewed.length) {
			return 'Sheet ' + sheets[index].sheetName + ' of the uploaded workbook does not match the preview.';
		}
		for (let row = 0; row < rows.length; row++) {
			if (rows[row].rowNumber !== reviewed[row].row) {
				return 'Sheet ' + sheets[index].sheetName + ' of the uploaded workbook does not match the preview.';
			}
		}
	}
	return null;
}

/**
 * Finds a row of the result by its sheet (index in the workbook, or sheet name) and
 * its row number in Excel.
 *
 * @returns {{sheetIndex: number, rowIndex: number}|null}
 */
export function findRow(result, sheet, rowNumber) {
	const sheets = result.sheets || [];
	const sheetIndex = typeof sheet === 'number' ? (sheet >= 0 && sheet < sheets.length ? sheet : -1) : sheets.findIndex((candidate) => candidate.sheetName === sheet);
	if (sheetIndex < 0) {
		return null;
	}
	const rowIndex = (sheets[sheetIndex].rows || []).findIndex((row) => row.row === Number(rowNumber));
	return rowIndex < 0 ? null : { sheetIndex, rowIndex };
}

/** The JSON objects written into a text, for example into the detail of a server fault. */
export function jsonObjectsIn(text) {
	const found = [];
	if (typeof text !== 'string') {
		return found;
	}
	let start = text.indexOf('{');
	while (start >= 0 && found.length < 10) {
		let depth = 0;
		let inString = false;
		let escaped = false;
		let end = -1;
		for (let at = start; at < text.length; at++) {
			const char = text[at];
			if (inString) {
				if (escaped) {
					escaped = false;
				} else if (char === '\\') {
					escaped = true;
				} else if (char === '"') {
					inString = false;
				}
			} else if (char === '"') {
				inString = true;
			} else if (char === '{') {
				depth++;
			} else if (char === '}' && --depth === 0) {
				end = at;
				break;
			}
		}
		if (end < 0) {
			break;
		}
		try {
			found.push(JSON.parse(text.slice(start, end + 1)));
			start = text.indexOf('{', end + 1);
		} catch (error) {
			start = text.indexOf('{', start + 1);
		}
	}
	return found;
}

/**
 * The failed write a refused batch names. When a write fails, the server rolls
 * the whole request back and answers with a fault whose detail is JSON:
 * { sheet, sheetName, row, message, rows: [ { sheet, row } ] }, where `rows` lists
 * every row of the failed write, the named row first.
 *
 * @param {{code: string, message: string, detail: string}|null} fault The fault of the answer.
 * @param {object} result The validation result.
 * @returns {{named: {sheetIndex: number, rowIndex: number}, rows: Array<{sheetIndex: number, rowIndex: number}>,
 *   message: string}|null} null when the fault names no row of the result.
 */
export function writeFailureOf(fault, result) {
	if (!fault) {
		return null;
	}
	const locate = (entry) => {
		if (!entry || !(Number(entry.row) > 0)) {
			return null;
		}
		const sheets = result.sheets || [];
		const byIndex = Number.isInteger(entry.sheet) && sheets[entry.sheet] &&
			(typeof entry.sheetName !== 'string' || sheets[entry.sheet].sheetName === entry.sheetName);
		if (byIndex) {
			return findRow(result, entry.sheet, entry.row);
		}
		if (typeof entry.sheetName === 'string') {
			return findRow(result, entry.sheetName, entry.row);
		}
		return typeof entry.sheet === 'string' ? findRow(result, entry.sheet, entry.row) : null;
	};
	for (const text of [fault.detail, fault.message]) {
		for (const object of jsonObjectsIn(text)) {
			const named = locate(object);
			if (!named) {
				continue;
			}
			const rows = (Array.isArray(object.rows) ? object.rows : []).map(locate).filter((where) => where !== null);
			return { named, rows, message: String(object.message || fault.message || 'The row could not be written.') };
		}
	}
	return null;
}

/**
 * Takes the rows of a failed write out of a batch, so that the batch can be sent
 * again without them. A main-sheet row that was to be written takes its
 * relationship rows with it. A unit left with nothing to write is taken out too.
 *
 * @param {Array<object>} batch The units of the batch; changed in place.
 * @param {object} result The validation result.
 * @param {{named: {sheetIndex: number, rowIndex: number}, rows: Array<{sheetIndex: number, rowIndex: number}>, message: string}} failure
 * @returns {Array<{sheetIndex: number, rowIndex: number, message: string}>|null} The rows that
 *   failed, each with its message, or null when the batch holds none of the rows.
 */
export function dropFailedWrite(batch, result, failure) {
	const { named, message } = failure;
	const nameOf = (where) => 'row ' + result.sheets[where.sheetIndex].rows[where.rowIndex].row + ' of ' + result.sheets[where.sheetIndex].sheetName;
	const failed = [];
	const done = new Set();
	const fail = (sheetIndex, rowIndex, text) => {
		const key = sheetIndex + ':' + rowIndex;
		if (!done.has(key)) {
			done.add(key);
			failed.push({ sheetIndex, rowIndex, message: text });
		}
	};
	const own = (where) => (where.sheetIndex === named.sheetIndex && where.rowIndex === named.rowIndex
		? message
		: 'Not imported, because it was written together with ' + nameOf(named) + ', which failed: ' + message);
	const mains = mainSheetIndexes(result);
	const mainSheetOfUnit = (unit) => unit.mainSheet;
	for (const where of [named].concat(failure.rows || [])) {
		if (mains.includes(where.sheetIndex)) {
			const unit = batch.find((candidate) => candidate.main === where.rowIndex && mainSheetOfUnit(candidate) === where.sheetIndex);
			if (!unit || !unit.writesMain) {
				continue;
			}
			fail(where.sheetIndex, where.rowIndex, own(where));
			for (const [relSheet, relRow] of unit.rels) {
				fail(relSheet, relRow, 'Not imported, because its parent (' + nameOf(where) + ') failed: ' + message);
			}
			unit.writesMain = false;
			unit.rels = [];
			continue;
		}
		for (const unit of batch) {
			const at = unit.rels.findIndex(([relSheet, relRow]) => relSheet === where.sheetIndex && relRow === where.rowIndex);
			if (at >= 0) {
				unit.rels.splice(at, 1);
				fail(where.sheetIndex, where.rowIndex, own(where));
				break;
			}
		}
	}
	// Rows that point at an item a failed main-sheet row was to create fail with it.
	const failedMain = failed.filter((entry) => mains.includes(entry.sheetIndex))
		.map((entry) => rowKey(entry.sheetIndex, result.sheets[entry.sheetIndex].rows[entry.rowIndex].row));
	for (const entry of dropRowsPointingAt(batch, result, failedMain)) {
		fail(entry.sheetIndex, entry.rowIndex, entry.message);
	}
	for (let position = batch.length - 1; position >= 0; position--) {
		if (!batch[position].writesMain && batch[position].rels.length === 0) {
			batch.splice(position, 1);
		}
	}
	return failed.length > 0 ? failed : null;
}

/**
 * Whether a failed main-sheet row means its item was not created, so that the rows of later
 * batches that point at it fail too. A row that failed as "changed" is left out: it may have
 * been imported already, by an earlier run that was broken off, and then its item exists. When
 * its item does not exist, the server's check of a later batch fails the rows that point at it.
 */
export function failedRowLeavesNoItem(result, failure) {
	return isMainSheet(result, failure.sheetIndex) && failure.reason !== 'changed';
}

/**
 * Takes the relationship rows out of a batch whose related item is one that a main-sheet
 * row was to create, for main-sheet rows that were not imported: those items don't exist.
 *
 * @param {Array<object>} batch The units of a batch; changed in place. Units left empty stay.
 * @param {object} result The validation result.
 * @param {Iterable<string>} rowKeys The main-sheet rows that were not imported, as rowKey(sheet index, row number).
 * @returns {Array<{sheetIndex: number, rowIndex: number, message: string}>} The rows taken out.
 */
export function dropRowsPointingAt(batch, result, rowKeys) {
	const gone = new Set(rowKeys);
	const dropped = [];
	if (gone.size === 0) {
		return dropped;
	}
	for (const unit of batch) {
		unit.rels = unit.rels.filter(([sheetIndex, rowIndex]) => {
			const made = newRelatedItem(result.sheets[sheetIndex].rows[rowIndex]);
			if (made === null || !gone.has(rowKey(made.sheet, made.row))) {
				return true;
			}
			const mainName = result.sheets[made.sheet] ? result.sheets[made.sheet].sheetName : '';
			dropped.push({ sheetIndex, rowIndex, message: 'Not imported, because its related item (row ' + made.row + ' of ' + mainName + ') was not imported.' });
			return false;
		});
	}
	return dropped;
}

/**
 * Takes the rows of an indented sheet out of a batch when every line above them failed on import.
 * A row that has already been written stays written: batches go at the same time.
 *
 * @param {Array<object>} batch The units of a batch; changed in place. Units left empty stay.
 * @param {object} result The validation result.
 * @param {object} workbook The workbook, as parseWorkbook returns it.
 * @param {Set<string>} failed The rows that were not imported, as rowKey(sheet index, row number).
 * @returns {Array<{sheetIndex: number, rowIndex: number, message: string}>} The rows taken out.
 */
export function dropRowsBelowFailed(batch, result, workbook, failed) {
	const dropped = [];
	if (failed.size === 0) {
		return dropped;
	}
	for (const unit of batch) {
		unit.rels = unit.rels.filter(([sheetIndex, rowIndex]) => {
			const info = workbook.sheets[sheetIndex].indented;
			const row = result.sheets[sheetIndex].rows[rowIndex];
			const above = info ? info.parents[row.row] : null;
			if (!above || above.length === 0 || above.some((number) => number === null || !failed.has(rowKey(sheetIndex, number)))) {
				return true;
			}
			dropped.push({ sheetIndex, rowIndex, message: 'Not imported, because the line above it (row ' + above.join(', ') + ') was not imported.' });
			return false;
		});
	}
	return dropped;
}

/**
 * Reads the answer to a batch that went through.
 *
 * @returns {{created: number, updated: number, deleted: number, unchanged: number,
 *   failed: Array<{sheetIndex: number, rowIndex: number, message: string, reason: (string|null)}>,
 *   unknown: Array<{sheet: *, row: number, message: string}>}} `unknown` holds failed rows the result does not have;
 *   `reason` is the server's reason for a failed row (see FRG_MCE_ImportWorkbook).
 */
export function readBatchAnswer(answer, result) {
	const read = { created: 0, updated: 0, deleted: 0, unchanged: 0, failed: [], unknown: [] };
	const reasons = {
		changed: 'Changed since the preview, or imported already. Upload the workbook again to review it.',
		error: 'The row has an error now.',
		parent: 'Not imported, because its parent row failed.',
		related: 'Not imported, because the row that creates its related item failed.',
		file: 'Not imported, because a file it links was not stored for the import.'
	};
	for (const row of (answer && Array.isArray(answer.rows)) ? answer.rows : []) {
		if (row.status === 'created' || row.status === 'updated' || row.status === 'deleted' || row.status === 'unchanged') {
			read[row.status]++;
			continue;
		}
		if (row.status !== 'failed') {
			continue;
		}
		const message = row.message || reasons[row.reason] || 'The row was not imported.';
		const where = findRow(result, row.sheet, row.row);
		if (where) {
			read.failed.push({ ...where, message, reason: row.reason || null });
		} else {
			read.unknown.push({ sheet: row.sheet, row: Number(row.row), message });
		}
	}
	return read;
}

/** A failed row as the `complete` step takes it. */
export function failedRowEntry(result, failure) {
	const sheet = result.sheets[failure.sheetIndex];
	const row = sheet.rows[failure.rowIndex];
	return { sheet: failure.sheetIndex, sheetName: sheet.sheetName, row: row.row, key: rowName(row), message: failure.message };
}

// ---------------------------------------------------------------------------
// Batches at the same time
// ---------------------------------------------------------------------------
//
// An import may send several batches at once. A batch whose relationship rows
// point at items that an earlier batch creates waits until that batch is done,
// so the items exist when its rows are written. Batches that touch the same
// existing item or the same file go one after the other: at the same time they
// would hold each other up in the database, or deadlock.

/** The existing items and the files a batch writes or points at. */
function batchTouches(batch, result) {
	const touches = new Set(batchFileTokens(batch, result));
	for (const unit of batch) {
		if (unit.parentId) {
			touches.add(unit.parentId);
		}
		for (const [sheetIndex, rowIndex] of unit.rels) {
			const related = result.sheets[sheetIndex].rows[rowIndex].related;
			if (related && related.id) {
				touches.add(related.id);
			}
		}
	}
	return touches;
}

/**
 * For each batch, the earlier batches it waits for: those with a main-sheet row that creates
 * an item one of its relationship rows points at, and the last one before it to touch the same
 * existing item or file (an item it changes, is the parent of its rows, or is their related
 * item). That one waits in turn for the one before it, so of the batches touching an item or
 * file only one runs at a time. A row that points at an item a later batch creates (a group of
 * items larger than a batch, see makeBatches) does not make it wait.
 *
 * @param {Array<Array<object>>} batches The batches of makeBatches.
 * @param {object} result The validation result.
 * @returns {number[][]} For each batch, the indexes of the batches it waits for, ascending.
 */
export function batchDependencies(batches, result) {
	const sheets = result.sheets || [];
	const madeIn = new Map();
	batches.forEach((batch, index) => {
		for (const unit of batch) {
			const mainSheet = unit.mainSheet;
			const row = unit.main !== null && unit.writesMain && sheets[mainSheet] ? sheets[mainSheet].rows[unit.main] : null;
			if (row && row.action === 'new') {
				madeIn.set(rowKey(mainSheet, row.row), index);
			}
		}
	});
	const lastToTouch = new Map();
	return batches.map((batch, index) => {
		const waits = new Set();
		for (const unit of batch) {
			for (const [sheetIndex, rowIndex] of unit.rels) {
				const made = newRelatedItem(sheets[sheetIndex].rows[rowIndex]);
				const maker = made === null ? undefined : madeIn.get(rowKey(made.sheet, made.row));
				if (maker !== undefined && maker < index) {
					waits.add(maker);
				}
			}
		}
		for (const touched of batchTouches(batch, result)) {
			if (lastToTouch.has(touched)) {
				waits.add(lastToTouch.get(touched));
			}
			lastToTouch.set(touched, index);
		}
		return Array.from(waits).sort((a, b) => a - b);
	});
}

/**
 * Runs tasks 0 to count - 1, at most `limit` at once. A task starts once the tasks it waits for
 * are done; of the tasks that may start, the one with the lowest index goes first. After a task
 * fails no other task starts, and once the running ones are done the run fails with the error
 * of the first that failed.
 *
 * @param {number} count How many tasks.
 * @param {number[][]} waitsFor For each task, the tasks it waits for (lower indexes).
 * @param {number} limit How many tasks may run at once.
 * @param {function(number): Promise<*>} run Runs a task.
 * @returns {Promise<void>}
 */
export function runScheduled(count, waitsFor, limit, run) {
	const most = Math.max(1, Math.floor(Number(limit) || 0) || 1);
	const state = new Array(count).fill('waiting');
	let running = 0;
	let failure = null;
	return new Promise((resolve, reject) => {
		const pump = () => {
			for (let index = 0; failure === null && index < count && running < most; index++) {
				if (state[index] !== 'waiting' || !(waitsFor[index] || []).every((other) => state[other] === 'done')) {
					continue;
				}
				state[index] = 'running';
				running++;
				Promise.resolve()
					.then(() => run(index))
					.then(() => {
						state[index] = 'done';
					}, (error) => {
						state[index] = 'failed';
						if (failure === null) {
							failure = { error };
						}
					})
					.then(() => {
						running--;
						pump();
					});
			}
			if (running > 0) {
				return;
			}
			if (failure !== null) {
				reject(failure.error);
			} else if (state.every((value) => value === 'done')) {
				resolve();
			} else {
				reject(new Error('Some batches wait for each other and cannot be sent.'));
			}
		};
		pump();
	});
}

/** The code of the server's answer to a batch that met another request in the database and can be sent again. */
export const RETRY_CODE = 'MCE_RETRY';

/**
 * Whether a refused request can be sent again unchanged: the database chose it as the victim
 * of a deadlock with another request, and rolled it back.
 *
 * @param {{code: string, message: string}|null} fault The fault of the answer.
 */
export function isRetryFault(fault) {
	if (!fault) {
		return false;
	}
	return String(fault.code || '').includes(RETRY_CODE) || /deadlock/i.test(String(fault.message || ''));
}

/** How long to wait before sending a request again, in milliseconds: longer after each try, and a bit random. */
export function retryDelay(attempt, random = Math.random) {
	return Math.round(500 * Math.pow(2, Math.max(0, attempt - 1)) * (0.75 + random() / 2));
}

// ---------------------------------------------------------------------------
// Files
// ---------------------------------------------------------------------------
//
// A file cell names a file the user picked with the workbook. The validation
// gives it as a file token, "file:<MD5>:<file name>": the stored value of a file
// cell, the related_id of a row that gets another file, and related.file on a
// sheet whose related item is a File. Before a batch, the browser stores the
// files its rows name, as File items with that name; the server links each to
// the File whose checksum and name are those of the token (see FRG_MCE_ImportWorkbook).

const FILE_TOKEN = /^file:([0-9A-F]{32}):(.+)$/;

/** Whether a value is a file token. */
export function isFileToken(value) {
	return typeof value === 'string' && FILE_TOKEN.test(value);
}

/** The file name of a file token. */
export function tokenName(token) {
	const match = FILE_TOKEN.exec(String(token));
	return match ? match[2] : String(token);
}

/** The file token of a picked file: its MD5 and the name at the end of its path. */
export function fileToken(path, md5) {
	return 'file:' + String(md5).toUpperCase() + ':' + fileNameOf(path);
}

/** The path of a picked file as the validation reads it: forward slashes, no leading ./ or /. */
export function normalPath(path) {
	let text = String(path || '').trim().replace(/\\/g, '/');
	while (text.startsWith('./') || text.startsWith('/')) {
		text = text.slice(text.startsWith('./') ? 2 : 1);
	}
	return text;
}

/** The file name at the end of a path. */
export function fileNameOf(path) {
	const text = normalPath(path);
	return text.slice(text.lastIndexOf('/') + 1);
}

/** The file tokens a row of the validation result names. */
export function rowFileTokens(row) {
	const tokens = new Set();
	for (const value of Object.values((row && row.values) || {})) {
		if (isFileToken(value)) {
			tokens.add(value);
		}
	}
	if (row && row.related && isFileToken(row.related.file)) {
		tokens.add(row.related.file);
	}
	return Array.from(tokens);
}

/** The file tokens of the rows a batch writes, in the order of the rows. */
export function batchFileTokens(batch, result) {
	const tokens = new Set();
	for (const unit of batch) {
		const rows = (unit.writesMain ? [result.sheets[unit.mainSheet].rows[unit.main]] : [])
			.concat(unit.rels.map(([sheetIndex, rowIndex]) => result.sheets[sheetIndex].rows[rowIndex]));
		for (const row of rows) {
			for (const token of rowFileTokens(row)) {
				tokens.add(token);
			}
		}
	}
	return Array.from(tokens);
}

/** The file tokens of every row an import of the result writes. */
export function neededFileTokens(result) {
	const tokens = new Set();
	for (const sheet of (result && result.sheets) || []) {
		for (const row of sheet.rows || []) {
			if (isWritten(row)) {
				for (const token of rowFileTokens(row)) {
					tokens.add(token);
				}
			}
		}
	}
	return Array.from(tokens);
}

// ---------------------------------------------------------------------------
// Texts
// ---------------------------------------------------------------------------

/** One line that says what an import of the result does. */
export function summaryLine(counts) {
	const parts = [];
	parts.push(plural(counts.create, 'row creates', 'rows create') + ' an item or relationship row');
	parts.push(plural(counts.update, 'row changes', 'rows change') + ' one');
	if (counts.delete > 0) {
		parts.push(plural(counts.delete, 'row deletes', 'rows delete') + ' a relationship row');
	}
	parts.push(plural(counts.unchanged, 'row is', 'rows are') + ' unchanged');
	if (counts.skipped > 0) {
		parts.push(plural(counts.skipped, 'row has', 'rows have') + ' errors and ' + (counts.skipped === 1 ? 'is' : 'are') + ' skipped');
	}
	return parts.join(', ') + '.';
}

/**
 * The progress of an import, as text. `running` holds the numbers of the batches being sent; without
 * it, `batch` is the number of the batch being sent.
 */
export function progressText(progress) {
	const done = progress.written + progress.failed;
	const running = Array.isArray(progress.running) ? progress.running.slice().sort((a, b) => a - b) : [];
	const which = running.length > 1
		? 'Batches ' + running.slice(0, -1).join(', ') + ' and ' + running[running.length - 1]
		: 'Batch ' + Math.min(running.length === 1 ? running[0] : progress.batch, progress.batches);
	return which + ' of ' + progress.batches + ': ' +
		done + ' of ' + progress.toWrite + ' rows' + (progress.failed > 0 ? ', ' + progress.failed + ' failed' : '') + '.';
}

/** The file name of the stored validation result. */
export function resultsFileName(jobNumber) {
	const safe = String(jobNumber || '').replace(/[\\/:*?"<>|]+/g, '_').trim();
	return [safe, 'preview'].filter((part) => part !== '').join(' ') + '.json';
}

/** The action of a row the import skipped because it had errors. */
export const SKIPPED_ACTION = 'Skipped';

/**
 * The rows an import skipped because they had errors, as records to list beside the stored ones.
 * The import stores records only for the rows it wrote or tried to write, so the skipped rows are
 * read from the validation result the import ran from.
 *
 * @param {object} result The validation result.
 * @returns {object[]} Records { sheet, row, action: 'Skipped', name, message }, without a target.
 */
export function skippedRecords(result) {
	const records = [];
	for (const sheet of (result && result.sheets) || []) {
		for (const row of sheet.rows || []) {
			if (row.action === 'error') {
				const errors = rowMessages(row).filter((issue) => issue.severity === 'error');
				records.push({
					sheet: sheet.sheetName, row: row.row, action: SKIPPED_ACTION, targetType: '', targetId: '', name: rowName(row),
					message: errors.map((issue) => issue.message).join('  ')
				});
			}
		}
	}
	return records;
}

/**
 * The import records in the order of the workbook: sheet by sheet, row by row. When the sheet
 * names of the workbook are given, the sheets keep that order. Otherwise, and for a sheet not
 * named there, a sheet's place is where its first record was written (records carry their write
 * order), so the main sheet, whose rows are written before their relationship rows, comes first.
 *
 * @param {object[]} records
 * @param {string[]} [sheetOrder] The sheet names of the workbook, in order.
 */
export function sortRecords(records, sheetOrder = []) {
	const written = records.slice().sort((a, b) => (a.order || 0) - (b.order || 0));
	const rank = new Map();
	for (const name of sheetOrder) {
		if (!rank.has(name)) {
			rank.set(name, rank.size);
		}
	}
	for (const record of written) {
		if (!rank.has(record.sheet)) {
			rank.set(record.sheet, rank.size);
		}
	}
	return records.slice().sort((a, b) => rank.get(a.sheet) - rank.get(b.sheet) || a.row - b.row);
}

/**
 * The targets of a validation result with their sheets: each main sheet, with the relationship
 * sheets whose parents are on it.
 *
 * @param {object} result The validation result.
 * @returns {Array<{label: string, sheets: string[]}>} In the order of the main sheets; label is the main sheet's name.
 */
export function resultTargets(result) {
	const sheets = (result && result.sheets) || [];
	return sheetGroups(result).map((group) => ({ label: sheets[group[0]].sheetName, sheets: group.map((index) => sheets[index].sheetName) }));
}

/**
 * The sheets of a validation result by target, as indexes in result.sheets: each main sheet,
 * then the relationship sheets whose parents are on it, in the order of the main sheets. This is
 * the order the workbook shows its tabs in.
 *
 * @param {object} result The validation result.
 * @returns {number[][]} One array per target; its first index is the target's main sheet.
 */
export function sheetGroups(result) {
	const sheets = (result && result.sheets) || [];
	const mains = mainSheetIndexes(result);
	const groups = mains.map((index) => [index]);
	sheets.forEach((sheet, index) => {
		if (!mains.includes(index)) {
			groups[mains.indexOf(sheet.parentSheet)].push(index);
		}
	});
	return groups;
}

/** The indexes of the sheets of a validation result in the order the workbook shows them (sheetGroups). */
export function sheetOrder(result) {
	return sheetGroups(result).flat();
}

/**
 * The target of each sheet of a validation result, by index in result.sheets: the position of its
 * group in sheetGroups. The sheets of one target share a colour, as the workbook's tabs do.
 */
export function sheetTargets(result) {
	const targets = [];
	sheetGroups(result).forEach((group, at) => {
		for (const index of group) {
			targets[index] = at;
		}
	});
	return targets;
}

/**
 * What an import did for each target: the records of the target's sheets, counted by action.
 *
 * @param {object[]} records Records { sheet, action }, as the job's result lists them.
 * @param {Array<{label: string, sheets: string[]}>} targets As resultTargets gives them.
 * @returns {Array<{label: string, rows: number, Created: number, Updated: number, Deleted: number, Failed: number, Skipped: number}>}
 */
export function recordCountsByTarget(records, targets) {
	return targets.map((target) => {
		const counts = { label: target.label, rows: 0, Created: 0, Updated: 0, Deleted: 0, Failed: 0, Skipped: 0 };
		for (const record of records) {
			if (target.sheets.includes(record.sheet)) {
				counts.rows++;
				if (Object.prototype.hasOwnProperty.call(counts, record.action) && record.action !== 'label' && record.action !== 'rows') {
					counts[record.action]++;
				}
			}
		}
		return counts;
	});
}

/** The line of one target in the result: "Part: 3 created, 1 updated, 0 failed". */
export function targetCountsLine(counts) {
	const parts = [counts.Created + ' created', counts.Updated + ' updated'];
	if (counts.Deleted > 0) {
		parts.push(counts.Deleted + ' deleted');
	}
	parts.push(counts.Failed + ' failed');
	if (counts.Skipped > 0) {
		parts.push(counts.Skipped + ' skipped because of errors');
	}
	return counts.label + ': ' + parts.join(', ');
}
