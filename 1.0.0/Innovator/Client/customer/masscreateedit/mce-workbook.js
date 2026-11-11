/**
 * Mass Create & Edit - the Excel workbook module.
 *
 * buildWorkbook(schema, options) turns a template schema (the sheets and
 * columns of an ItemType, as the server describes them) into an .xlsx with
 * Excel data validation. parseWorkbook(arrayBuffer) reads a filled-in workbook
 * back into plain rows. Nothing here talks to the server and nothing here
 * judges the data: Excel validation only guides the user (paste goes around
 * it, and a blank cell is never checked), so the server validates every cell
 * again.
 *
 * Workbook layout:
 *   - one visible, protected sheet per schema sheet: row 1 is the locked
 *     header, the data cells below it are unlocked. A job with several targets
 *     has a main sheet for each. The schema lists the main sheets first, and
 *     _meta keeps those indexes; the tabs show each main sheet followed by its
 *     relationship sheets, and the tabs of a target share a colour. A main
 *     sheet is named after its ItemType, a relationship sheet after its
 *     RelationshipType ("Part CAD");
 *   - `_lists` (veryHidden): the drop-down values, two columns per list
 *     (the text Excel shows, and the value Aras stores);
 *   - `_meta` (veryHidden): what the workbook was generated from and which
 *     column holds which property. A file without it is refused.
 *
 * The .xlsx file itself is written by mce-xlsx-write.js and read by
 * mce-xlsx-read.js. Both are part of this tool and need no library: the zip
 * is packed and unpacked with the browser's own CompressionStream and
 * DecompressionStream. This module loads them when it first needs them, with
 * its own cache tag, so a new version of the tool brings all three files.
 *
 * Rules this module follows:
 *   - dates are built with Date.UTC and read through the UTC fields, because
 *     the writer turns a Date into a serial number from its UTC fields. A cell
 *     therefore holds the wall-clock value, whatever time zone the browser is in;
 *   - one validation entry per column range, never one per cell.
 */

const FORMAT_VERSION = 1;
const INDENTED_LAYOUT = 'indented';
const META_SHEET = '_meta';
const LISTS_SHEET = '_lists';
const META_MARKER = 'mce.meta';
const META_LAYOUT = 1;
const CLEAR_TOKEN = '#CLEAR';
const DELETE_TOKEN = '#DELETE';
const DEFAULT_MAX_ROWS = 5000;
const EXCEL_LAST_ROW = 1048576;
const CELL_TEXT_LIMIT = 32767;
const INT32_MIN = -2147483648;
const INT32_MAX = 2147483647;

// Excel's limits for the texts of a validation entry.
const TITLE_LIMIT = 32;
const MESSAGE_LIMIT = 255;
// A text literal inside a formula can't be longer than this.
const FORMULA_TEXT_LIMIT = 255;

// Serial numbers of the first and the last day Excel can show as a date.
const FIRST_DATE = new Date(Date.UTC(1899, 11, 31));
const LAST_DATE = new Date(Date.UTC(9999, 11, 31));
const LAST_DATE_TIME = new Date(Date.UTC(9999, 11, 31, 23, 59, 59));
const LAST_DATE_SERIAL = 2958465;

const DATE_ONLY_PATTERNS = ['short_date', 'long_date'];
const DATE_FORMAT = 'yyyy-mm-dd';
const DATE_TIME_FORMAT = 'yyyy-mm-dd hh:mm:ss';
const TEXT_FORMAT = '@';

const DATA_TYPES = [
	'string', 'text', 'formatted text', 'ml_string', 'integer', 'decimal', 'float', 'date', 'boolean',
	'list', 'filter list', 'mv_list', 'item', 'color', 'color list',
	// A cell that names a file, and the column of the delete marker.
	'file', 'delete'
];

const HEADER_FILL = 'FFDDEBF7';
const KEY_HEADER_FILL = 'FFFFE699';
const NOT_APPLICABLE_FILL = 'FFD9D9D9';
const NOT_APPLICABLE_FONT = 'FF808080';
// The tab colour of each target's sheets, in the order of the targets, when there are several.
// The Import tab of the job uses the same colours (mce-job-form.css).
const TARGET_TAB_COLORS = ['FF4472C4', 'FFED7D31', 'FF70AD47', 'FF9B59B6', 'FFFFC000', 'FF1ABC9C'];

// The positional fields of a `sheet` and a `column` record on the _meta sheet.
// Every field has a cell of its own, so no cell can outgrow Excel's limit of
// 32,767 characters, however many columns an ItemType has.
const META_SHEET_FIELDS = [
	'index', 'sheetName', 'kind', 'itemTypeId', 'itemTypeName', 'relationshipTypeId', 'relationshipTypeName',
	'relatedItemTypeId', 'relatedItemTypeName', 'createOnly', 'modifiedOnColumn', 'parentSheet', 'relatedSheet', 'layout'
];
const META_COLUMN_FIELDS = [
	'sheetIndex', 'position', 'id', 'role', 'property', 'dataType', 'isKey', 'isMatchKey', 'listColumn'
];
const META_VALUE_KEYS = ['formatVersion', 'database', 'generatedOn', 'language', 'definitionHash', 'jobId', 'jobToken'];

function fail(code, message) {
	const error = new Error(message);
	error.code = code;
	return error;
}

// ---------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------

const isDate = (value) => Object.prototype.toString.call(value) === '[object Date]';
const isFiniteNumber = (value) => typeof value === 'number' && Number.isFinite(value);
const pad = (number, length) => String(number).padStart(length, '0');
const clip = (text, limit) => (text.length > limit ? text.slice(0, limit - 3) + '...' : text);
const foldCase = (text) => String(text).trim().toLowerCase();

function columnLetter(position) {
	let letters = '';
	for (let n = position; n > 0; n = Math.floor((n - 1) / 26)) {
		letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
	}
	return letters;
}

/** A neutral date string (or the start of one) as a Date whose UTC fields hold that wall-clock value. */
function parseNeutralDate(text) {
	const match = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?/.exec(String(text).trim());
	if (!match) {
		return null;
	}
	const [year, month, day, hour, minute, second] = match.slice(1).map((part) => Number(part || 0));
	const date = new Date(Date.UTC(year, month - 1, day, hour, minute, second));
	return Number.isNaN(date.getTime()) ? null : date;
}

/** yyyy-MM-ddTHH:mm:ss from the UTC fields, rounded to the second (a serial number is a fraction of a day). */
function formatNeutralDate(date) {
	const rounded = new Date(Math.round(date.getTime() / 1000) * 1000);
	return (
		pad(rounded.getUTCFullYear(), 4) + '-' + pad(rounded.getUTCMonth() + 1, 2) + '-' + pad(rounded.getUTCDate(), 2) +
		'T' + pad(rounded.getUTCHours(), 2) + ':' + pad(rounded.getUTCMinutes(), 2) + ':' + pad(rounded.getUTCSeconds(), 2)
	);
}

function serialToDate(serial, date1904) {
	return new Date(Math.round((serial - 25569 + (date1904 ? 1462 : 0)) * 86400) * 1000);
}

/** A number as Excel reads it inside a formula ("1E-7", not "1e-7"). */
const formulaNumber = (number) => String(number).toUpperCase();

/** A text as a formula literal. */
const formulaText = (text) => '"' + String(text).replace(/"/g, '""') + '"';

function toNumberOrNull(value) {
	if (value === null || value === undefined || value === '') {
		return null;
	}
	const number = Number(value);
	return Number.isFinite(number) ? number : null;
}

// ---------------------------------------------------------------------------
// Planning: sheet names, column kinds, drop-down lists
// ---------------------------------------------------------------------------

function checkSchema(schema) {
	const bad = (message) => fail('MCE_SCHEMA_INVALID', 'The template schema is not valid: ' + message);
	if (!schema || typeof schema !== 'object') {
		throw bad('it is missing.');
	}
	if (schema.formatVersion !== FORMAT_VERSION) {
		throw bad('format version ' + schema.formatVersion + ' is not supported (expected ' + FORMAT_VERSION + ').');
	}
	if (!Array.isArray(schema.sheets) || schema.sheets.length === 0) {
		throw bad('it has no sheets.');
	}
	schema.sheets.forEach((sheet, sheetIndex) => {
		const where = 'sheet ' + (sheetIndex + 1);
		if (!sheet || !sheet.itemType || !Array.isArray(sheet.columns)) {
			throw bad(where + ' has no itemType or no columns.');
		}
		const isMain = (index) => Number.isInteger(index) && schema.sheets[index] && schema.sheets[index].kind !== 'relationship';
		if (sheet.kind === 'relationship' && (!isMain(sheet.parentSheet) || (sheet.relatedSheet !== null && !isMain(sheet.relatedSheet)))) {
			throw bad(where + ' does not name the main sheets of its parent and its related item.');
		}
		const seen = new Set();
		for (const column of sheet.columns) {
			if (!column || typeof column.id !== 'string' || column.id === '') {
				throw bad(where + ' has a column without an id.');
			}
			if (seen.has(column.id)) {
				throw bad(where + ' has two columns with the id "' + column.id + '".');
			}
			seen.add(column.id);
			if (!DATA_TYPES.includes(column.dataType)) {
				throw bad('column "' + column.id + '" on ' + where + ' has the unknown data type "' + column.dataType + '".');
			}
			const isListType = column.dataType === 'list' || column.dataType === 'filter list';
			if (isListType && !(schema.lists && schema.lists[column.listId])) {
				throw bad('column "' + column.id + '" on ' + where + ' refers to a list that is not in the schema.');
			}
		}
	});
}

/** A sheet name Excel accepts: at most 31 characters, none of []:*?/\, not quoted, not reserved, not taken. */
function safeSheetName(wanted, taken) {
	const tidy = (text) => text.replace(/\s+/g, ' ').replace(/^[' ]+|[' ]+$/g, '');
	const cut = (text, length) => {
		// Don't cut a surrogate pair in half: a lone half is not valid in the file's XML.
		const end = /[\uD800-\uDBFF]/.test(text.charAt(length - 1)) ? length - 1 : length;
		return tidy(text.slice(0, end));
	};
	let base = cut(tidy(String(wanted || '').replace(/[[\]:*?/\\\u0000-\u001F]/g, ' ')), 31) || 'Sheet';
	if (base.toLowerCase() === 'history') {
		base = 'History_';
	}
	let name = base;
	for (let n = 2; taken.has(name.toLowerCase()); n++) {
		const suffix = ' (' + n + ')';
		name = cut(base, 31 - suffix.length) + suffix;
	}
	taken.add(name.toLowerCase());
	return name;
}

/**
 * What a drop-down shows for each value of a list.
 *
 * The label is shown. The value is shown instead when the label is empty or
 * when another value of the same list has the same label. A value shown that
 * way can still equal the label of another entry; that entry then shows its
 * value too, so every text in the drop-down stands for exactly one value.
 * Texts are compared without regard to case or surrounding spaces, the way
 * Excel compares typed text with a drop-down range.
 */
function displayedEntries(values) {
	const entries = (values || [])
		.map((entry) => ({
			value: entry.value === null || entry.value === undefined ? '' : String(entry.value),
			label: entry.label === null || entry.label === undefined ? '' : String(entry.label)
		}))
		.filter((entry) => entry.value.trim() !== '');
	const tally = (texts) => {
		const counts = new Map();
		for (const text of texts) {
			counts.set(text, (counts.get(text) || 0) + 1);
		}
		return counts;
	};
	const labels = tally(entries.map((entry) => foldCase(entry.label)));
	for (const entry of entries) {
		entry.showLabel = foldCase(entry.label) !== '' && labels.get(foldCase(entry.label)) === 1;
	}
	for (;;) {
		const shown = tally(entries.map((entry) => foldCase(entry.showLabel ? entry.label : entry.value)));
		const clashing = entries.filter((entry) => entry.showLabel && shown.get(foldCase(entry.label)) > 1);
		if (clashing.length === 0) {
			break;
		}
		for (const entry of clashing) {
			entry.showLabel = false;
		}
	}
	return entries.map((entry) => ({ display: entry.showLabel ? entry.label : entry.value, value: entry.value }));
}

/** Finds the entry a typed or stored text stands for: an exact match first, then one that differs only in case or spaces. */
function findEntry(entries, field, text) {
	const exact = entries.find((entry) => entry[field] === text);
	if (exact) {
		return exact;
	}
	const folded = foldCase(text);
	const loose = entries.filter((entry) => foldCase(entry[field]) === folded);
	return loose.length === 1 ? loose[0] : null;
}

/** The drop-down lists a workbook needs, each of which becomes a block of two columns on the _lists sheet. */
function createListPlanner(schema) {
	const blocks = [];
	const byKey = new Map();
	const block = (key, title, makeEntries) => {
		if (!byKey.has(key)) {
			const entries = makeEntries();
			const created = {
				title,
				entries,
				// Left column: the text shown. Right column: the stored value.
				column: blocks.length * 2 + 1,
				name: 'mce_list_' + (blocks.length + 1),
				usedPlain: false,
				usedWithClear: false
			};
			blocks.push(created);
			byKey.set(key, created);
		}
		return byKey.get(key);
	};
	return {
		blocks,
		forList: (listId) => block('list:' + listId, 'list ' + listId, () => displayedEntries(schema.lists[listId].values)),
		forClassPaths: (sheetIndex, classPaths) =>
			block('class:' + sheetIndex, 'class paths of sheet ' + (sheetIndex + 1), () =>
				classPaths.map((path) => ({ display: String(path), value: String(path) })))
	};
}

/** Works out, for one schema column, how its cells behave in Excel. */
function planColumn(column, position, sheet, sheetIndex, lists) {
	const plan = {
		def: column,
		position,
		letter: columnLetter(position),
		// #CLEAR empties a value. It makes no sense where a value is mandatory, so those columns reject it.
		clearable: !column.required && !column.isKey && !column.isMatchKey,
		kind: 'text',
		format: TEXT_FORMAT,
		width: 20,
		block: null
	};
	const hasClassPaths = Array.isArray(sheet.classPaths) && sheet.classPaths.length > 0;
	if (column.isClassification && hasClassPaths) {
		plan.kind = 'choice';
		plan.block = lists.forClassPaths(sheetIndex, sheet.classPaths);
		plan.width = 28;
		return plan;
	}
	switch (column.dataType) {
		case 'string':
		case 'ml_string':
			plan.kind = 'string';
			plan.width = Math.min(Math.max(toNumberOrNull(column.storedLength) || 20, 16), 40);
			break;
		case 'text':
		case 'formatted text':
			plan.kind = 'longText';
			plan.width = 40;
			break;
		case 'integer':
			plan.kind = 'integer';
			plan.format = 'General';
			plan.width = 16;
			break;
		case 'decimal':
		case 'float':
			plan.kind = 'number';
			plan.format = 'General';
			plan.width = 16;
			break;
		case 'date':
			plan.kind = DATE_ONLY_PATTERNS.includes(column.pattern) ? 'date' : 'dateTime';
			plan.format = plan.kind === 'date' ? DATE_FORMAT : DATE_TIME_FORMAT;
			plan.width = plan.kind === 'date' ? 16 : 22;
			break;
		case 'boolean':
			plan.kind = 'boolean';
			plan.width = 16;
			break;
		case 'list':
		case 'filter list':
			plan.kind = 'choice';
			plan.block = lists.forList(column.listId);
			plan.width = 24;
			break;
		case 'mv_list':
			plan.kind = 'multiValue';
			plan.width = 28;
			break;
		case 'item':
			plan.kind = 'item';
			plan.width = 28;
			break;
		case 'file':
			plan.kind = 'file';
			plan.width = 32;
			break;
		case 'delete':
			plan.kind = 'delete';
			plan.clearable = false;
			plan.width = 12;
			break;
		case 'color':
		case 'color list':
			plan.kind = 'color';
			plan.width = 16;
			break;
		default:
			break;
	}
	return plan;
}

/**
 * The main sheet a relationship sheet belongs to, and the main sheet of its related ItemType (null
 * when that is not a target of the job), by index, as the schema gives them.
 */
export function parentSheetOf(schema, sheet) {
	if (!sheet || sheet.kind !== 'relationship') {
		return null;
	}
	return sheet.parentSheet;
}

export function relatedSheetOf(schema, sheet) {
	if (!sheet || sheet.kind !== 'relationship') {
		return null;
	}
	return Number.isInteger(sheet.relatedSheet) ? sheet.relatedSheet : null;
}

/**
 * The sheets of a schema by target, as indexes in schema.sheets: each main sheet, then the
 * relationship sheets whose parents are on it, in the order of the main sheets. The workbook shows
 * its tabs in this order; the indexes, which _meta and the server use, stay those of the schema.
 */
export function sheetGroups(schema) {
	const mains = [];
	schema.sheets.forEach((sheet, index) => {
		if (sheet.kind !== 'relationship') {
			mains.push(index);
		}
	});
	const groups = mains.map((index) => [index]);
	schema.sheets.forEach((sheet, index) => {
		if (sheet.kind === 'relationship') {
			groups[mains.indexOf(parentSheetOf(schema, sheet))].push(index);
		}
	});
	return groups;
}

/**
 * The name a sheet wants: the label of a main sheet's ItemType, and the name of a relationship
 * sheet's RelationshipType ("Part CAD", not its label "CAD Documents"). The name says which
 * ItemType the rows belong to, and no two RelationshipTypes have the same one.
 */
const nameOfSheet = (sheet) => {
	const relationship = sheet.relationshipType || null;
	return (relationship && (relationship.name || relationship.label)) || sheet.itemType.label || sheet.itemType.name;
};

function planSheets(schema, lists) {
	const taken = new Set([META_SHEET, LISTS_SHEET]);
	return schema.sheets.map((sheet, sheetIndex) => {
		const label = nameOfSheet(sheet);
		const columns = sheet.columns.map((column, index) => planColumn(column, index + 1, sheet, sheetIndex, lists));
		for (const column of columns) {
			column.label = column.def.label || column.def.property || column.def.id;
			column.width = Math.min(Math.max(column.width, column.label.length + 4), 50);
		}
		return {
			def: sheet,
			index: sheetIndex + 1,
			name: safeSheetName(label, taken),
			columns,
			// Two empty columns sit between the data and the hidden column: a header note opens over the two
			// columns to the right of its cell, and it needs them to have a width.
			modifiedOnColumn: columns.length + 3,
			classification: columns.find((column) => column.def.isClassification && column.def.role === 'property') || null
		};
	});
}

// ---------------------------------------------------------------------------
// Validation, notes and conditional formatting
// ---------------------------------------------------------------------------

function integerBounds(column) {
	const min = toNumberOrNull(column.minRange);
	const max = toNumberOrNull(column.maxRange);
	let low = INT32_MIN;
	let high = INT32_MAX;
	if (min !== null) {
		low = Math.max(low, column.rangeInclusive ? Math.ceil(min) : Math.floor(min) + 1);
	}
	if (max !== null) {
		high = Math.min(high, column.rangeInclusive ? Math.floor(max) : Math.ceil(max) - 1);
	}
	return { low, high, bounded: min !== null || max !== null };
}

function rangeWords(column) {
	const min = toNumberOrNull(column.minRange);
	const max = toNumberOrNull(column.maxRange);
	const inclusive = !!column.rangeInclusive;
	if (min !== null && max !== null) {
		return inclusive ? ' from ' + min + ' to ' + max : ' greater than ' + min + ' and less than ' + max;
	}
	if (min !== null) {
		return inclusive ? ' of ' + min + ' or more' : ' greater than ' + min;
	}
	if (max !== null) {
		return inclusive ? ' of ' + max + ' or less' : ' less than ' + max;
	}
	return '';
}

/**
 * The validation entry of a column, or null when it has none.
 *
 * `cell` is the address of the first data cell; Excel shifts it for the rows
 * below. A column that accepts #CLEAR can't use Excel's whole-number, decimal
 * or date validation, which reject any text, so it gets a formula that accepts
 * either the token or a valid value. Those formulas choose with IF, never with
 * OR or AND across the two cases: Excel evaluates every argument of OR and
 * AND, and an error in the numeric test would reject the token as well.
 */
function validationFor(plan, cell) {
	const column = plan.def;
	const isClear = 'EXACT(' + cell + ',' + formulaText(CLEAR_TOKEN) + ')';
	const orClear = plan.clearable ? ' To empty the value, type ' + CLEAR_TOKEN + '.' : '';
	const entry = (rule, message) => ({
		allowBlank: true,
		showErrorMessage: true,
		errorStyle: 'stop',
		errorTitle: clip(plan.label, TITLE_LIMIT),
		error: clip(message + orClear, MESSAGE_LIMIT),
		...rule
	});
	const custom = (formula, message) => entry({ type: 'custom', formulae: [formula] }, message);

	switch (plan.kind) {
		case 'string': {
			const length = toNumberOrNull(column.storedLength);
			if (length === null || length <= 0) {
				return null;
			}
			const message = plan.label + ' allows at most ' + length + ' characters.';
			if (plan.clearable && length < CLEAR_TOKEN.length) {
				return custom('OR(LEN(' + cell + ')<=' + length + ',' + isClear + ')', message);
			}
			return entry({ type: 'textLength', operator: 'lessThanOrEqual', formulae: [length] }, message);
		}
		case 'integer': {
			const { low, high, bounded } = integerBounds(column);
			const message = bounded ? 'Enter a whole number from ' + low + ' to ' + high + '.' : 'Enter a whole number.';
			if (plan.clearable) {
				return custom(
					'IF(ISNUMBER(' + cell + '),AND(' + cell + '=INT(' + cell + '),' + cell + '>=' + low + ',' + cell + '<=' + high + '),' + isClear + ')',
					message
				);
			}
			return entry({ type: 'whole', operator: 'between', formulae: [low, high] }, message);
		}
		case 'number': {
			const min = toNumberOrNull(column.minRange);
			const max = toNumberOrNull(column.maxRange);
			const inclusive = !!column.rangeInclusive;
			const message = 'Enter a number' + rangeWords(column) + '.';
			const tests = [];
			if (min !== null) {
				tests.push(cell + (inclusive ? '>=' : '>') + formulaNumber(min));
			}
			if (max !== null) {
				tests.push(cell + (inclusive ? '<=' : '<') + formulaNumber(max));
			}
			const inRange = tests.length > 1 ? 'AND(' + tests.join(',') + ')' : tests[0];
			if (plan.clearable) {
				return custom(
					tests.length ? 'IF(ISNUMBER(' + cell + '),' + inRange + ',' + isClear + ')' : 'OR(ISNUMBER(' + cell + '),' + isClear + ')',
					message
				);
			}
			if (min !== null && max !== null) {
				// Excel's "between" includes both ends; it has no operator that excludes them.
				return inclusive
					? entry({ type: 'decimal', operator: 'between', formulae: [formulaNumber(min), formulaNumber(max)] }, message)
					: custom('AND(ISNUMBER(' + cell + '),' + tests.join(',') + ')', message);
			}
			if (min !== null) {
				return entry({ type: 'decimal', operator: inclusive ? 'greaterThanOrEqual' : 'greaterThan', formulae: [formulaNumber(min)] }, message);
			}
			if (max !== null) {
				return entry({ type: 'decimal', operator: inclusive ? 'lessThanOrEqual' : 'lessThan', formulae: [formulaNumber(max)] }, message);
			}
			return custom('ISNUMBER(' + cell + ')', message);
		}
		case 'date':
		case 'dateTime': {
			const message = plan.kind === 'date' ? 'Enter a date.' : 'Enter a date and time.';
			if (plan.clearable) {
				return custom('IF(ISNUMBER(' + cell + '),AND(' + cell + '>=1,' + cell + '<' + (LAST_DATE_SERIAL + 1) + '),' + isClear + ')', message);
			}
			return entry({ type: 'date', operator: 'between', formulae: [FIRST_DATE, plan.kind === 'date' ? LAST_DATE : LAST_DATE_TIME] }, message);
		}
		case 'boolean':
			return entry(
				{ type: 'list', formulae: [formulaText(plan.clearable ? 'Yes,No,' + CLEAR_TOKEN : 'Yes,No')] },
				'Pick Yes or No.'
			);
		case 'choice': {
			// A range has to hold at least one cell. A list without values and without the token gets no validation.
			if (plan.block.entries.length === 0 && !plan.clearable) {
				return null;
			}
			if (plan.clearable) {
				plan.block.usedWithClear = true;
			} else {
				plan.block.usedPlain = true;
			}
			return entry({ type: 'list', formulae: [plan.block.name + (plan.clearable ? '_clear' : '')] }, 'Pick a value from the drop-down.');
		}
		case 'item':
			return {
				type: 'any',
				showInputMessage: true,
				promptTitle: clip(plan.label, TITLE_LIMIT),
				prompt: clip(itemReferenceText(column), MESSAGE_LIMIT)
			};
		case 'file':
			return {
				type: 'any',
				showInputMessage: true,
				promptTitle: clip(plan.label, TITLE_LIMIT),
				prompt: clip(FILE_TEXT, MESSAGE_LIMIT)
			};
		case 'delete':
			return entry({ type: 'list', formulae: [formulaText(DELETE_TOKEN)] }, deleteText(plan.def));
		case 'color': {
			// HEX2DEC fails on anything but hexadecimal digits, and ISNUMBER turns that failure into FALSE.
			const isColor = 'AND(LEN(' + cell + ')=7,LEFT(' + cell + ',1)="#",ISNUMBER(HEX2DEC(MID(' + cell + ',2,6))))';
			return custom(plan.clearable ? 'OR(' + isColor + ',' + isClear + ')' : isColor, 'Enter a colour as #RRGGBB, for example #1A2B3C.');
		}
		default:
			return null;
	}
}

const FILE_TEXT = 'File: the name of a file you pick with the workbook when you upload it, ' +
	'or a path inside the picked folder when two files have the same name. The current file name keeps the file.';
const deleteText = (column) => (column && column.indentedOnly
	? DELETE_TOKEN + ' deletes this BOM line. The part stays in Aras, and so do the lines below it.'
	: DELETE_TOKEN + ' deletes the row this row matches. The file itself stays in Aras.');

function itemReferenceText(column) {
	const target = column.itemTarget && (column.itemTarget.label || column.itemTarget.name);
	return (
		'Item reference: the text must match the keyed name of ' + (target ? 'the ' + target : 'the referenced item') +
		' exactly. Upper and lower case are not distinguished.'
	);
}

/** The note on a header cell: what the column takes, in a few short sentences. */
function headerNote(plan, sheetPlan, schema, sheetPlans) {
	const column = plan.def;
	const sheet = sheetPlan.def;
	const lines = [];
	const several = schema.sheets.filter((candidate) => candidate.kind === 'main').length > 1;
	if (column.role === 'level') {
		lines.push('Where the row sits in the structure: 0 is a top assembly, 1 a line directly under it, 2 a line under that, and so on. ' +
			'A row\'s parent is the nearest row above it one level up. Only the part number of a level 0 row is read. Required.');
	} else if (column.role === 'parentKey') {
		const parentIndex = parentSheetOf(schema, sheet);
		const parent = schema.sheets[parentIndex].itemType;
		lines.push('Identifies the parent ' + (parent.label || parent.name) + ': ' +
			(several ? 'a row on sheet \'' + sheetPlans[parentIndex].name + '\'' : 'a row on the first sheet') + ' or an existing item. Required.');
	} else if (column.role === 'relatedKey') {
		const related = sheet.relationshipType && sheet.relationshipType.relatedItemType;
		const relatedIndex = relatedSheetOf(schema, sheet);
		lines.push('Identifies the related ' + ((related && (related.label || related.name)) || 'item') +
			(relatedIndex === null || !sheetPlans[relatedIndex] ? '' : ': an existing item, or a row on sheet \'' + sheetPlans[relatedIndex].name + '\', a new one too') + '. Required.');
	} else if (column.isKey) {
		lines.push('Key column. A row whose key matches an existing item updates that item; any other row creates one. Required.');
	} else if (column.required) {
		lines.push('Required.');
	}
	if (column.isMatchKey && !column.isKey) {
		lines.push('Part of the match key that identifies an existing row, so an import cannot change it.');
	}
	switch (plan.kind) {
		case 'string':
			if (toNumberOrNull(column.storedLength)) {
				lines.push('Text, at most ' + column.storedLength + ' characters.');
			}
			break;
		case 'integer':
			lines.push('Whole number' + rangeWords(column) + '.');
			break;
		case 'number':
			lines.push('Number' + rangeWords(column) + '.');
			break;
		case 'date':
			lines.push('Date.');
			break;
		case 'dateTime':
			lines.push('Date and time.');
			break;
		case 'boolean':
			lines.push('Yes or No.');
			break;
		case 'choice':
			lines.push('Pick a value from the drop-down.');
			break;
		case 'multiValue':
			lines.push('Takes the stored values of the list, separated by commas.');
			break;
		case 'item':
			lines.push(itemReferenceText(column));
			break;
		case 'file':
			lines.push(FILE_TEXT);
			break;
		case 'delete':
			lines.push(deleteText(column));
			break;
		case 'color':
			lines.push('Colour as #RRGGBB.');
			break;
		default:
			break;
	}
	if (column.classPath) {
		lines.push('Applies to class ' + column.classPath + ' only; greyed out on rows of another class.');
	}
	if (plan.clearable) {
		lines.push('A blank cell leaves the value unchanged; ' + CLEAR_TOKEN + ' empties it.');
	}
	return lines.join('\n');
}

/**
 * The formula that greys out a class-specific cell: the row has a
 * classification, and it is neither the column's class nor one of its
 * subclasses. Null when the class path is too long for a formula literal.
 */
function notApplicableFormula(classPath, classificationLetter) {
	const classCell = '$' + classificationLetter + '2';
	const subclassPrefix = classPath + '/';
	if (formulaText(subclassPrefix).length - 2 > FORMULA_TEXT_LIMIT) {
		return null;
	}
	return (
		'AND(' + classCell + '<>"",' + classCell + '<>' + formulaText(classPath) +
		',LEFT(' + classCell + ',' + subclassPrefix.length + ')<>' + formulaText(subclassPrefix) + ')'
	);
}

// ---------------------------------------------------------------------------
// Writing
// ---------------------------------------------------------------------------

// The workbook is described as plain data and written by mce-xlsx-write.js. Each style below is one
// object, shared by every cell that has it.

const HEADER_FONT = { name: 'Calibri', size: 11, bold: true };
const HEADER_ALIGNMENT = { vertical: 'top', wrapText: true };
const headerStyle = (fill) => ({ numFmt: TEXT_FORMAT, locked: true, font: HEADER_FONT, alignment: HEADER_ALIGNMENT, fill: { color: fill } });
const HEADER_STYLE = headerStyle(HEADER_FILL);
const KEY_HEADER_STYLE = headerStyle(KEY_HEADER_FILL);
const MODIFIED_ON_STYLE = { numFmt: TEXT_FORMAT, locked: true };
const LIST_STYLE = { numFmt: TEXT_FORMAT };
const NOT_APPLICABLE_STYLE = { fill: { color: NOT_APPLICABLE_FILL }, font: { color: NOT_APPLICABLE_FONT } };

let writerPromise = null;

/**
 * Loads mce-xlsx-write.js once, with the cache tag this module was loaded with
 * (see mce-job-form.js), so that a new version of the tool brings both files.
 */
function loadWriter() {
	if (!writerPromise) {
		const self = new URL(import.meta.url);
		writerPromise = import(new URL('./mce-xlsx-write.js' + self.search, self).href);
		// A failed load (a network error, say) must not block a later attempt.
		writerPromise.catch(() => {
			writerPromise = null;
		});
	}
	return writerPromise;
}

/** The rows of a sheet as the writer takes them, filled in cell by cell in any order. */
function createRows() {
	const byIndex = new Map();
	const row = (index) => {
		if (!byIndex.has(index)) {
			byIndex.set(index, { index, cells: [] });
		}
		return byIndex.get(index);
	};
	return {
		row,
		cell: (index, column, cell) => row(index).cells.push({ column, ...cell }),
		list: () => Array.from(byIndex.values())
	};
}

/** The _lists sheet, and the names of the ranges the validation entries use. */
function listsSheet(lists) {
	const rows = createRows();
	const columns = [];
	const definedNames = [];
	const text = (rowNumber, column, value) => rows.cell(rowNumber, column, { value, style: LIST_STYLE });
	for (const block of lists.blocks) {
		columns.push({ index: block.column, style: LIST_STYLE }, { index: block.column + 1, style: LIST_STYLE });
		text(1, block.column, block.title);
		text(1, block.column + 1, 'stored value');
		block.entries.forEach((entry, index) => {
			text(index + 2, block.column, entry.display);
			text(index + 2, block.column + 1, entry.value);
		});
		// The token is the last cell of the block. A column that accepts it uses the range that includes this
		// cell, and a column that doesn't uses the range that stops one row earlier.
		const clearRow = block.entries.length + 2;
		text(clearRow, block.column, CLEAR_TOKEN);
		text(clearRow, block.column + 1, CLEAR_TOKEN);
		const letter = columnLetter(block.column);
		const range = (lastRow) => '$' + letter + '$2:$' + letter + '$' + lastRow;
		if (block.usedPlain) {
			definedNames.push({ name: block.name, sheet: LISTS_SHEET, range: range(clearRow - 1) });
		}
		if (block.usedWithClear) {
			definedNames.push({ name: block.name + '_clear', sheet: LISTS_SHEET, range: range(clearRow) });
		}
	}
	return { sheet: { name: LISTS_SHEET, state: 'veryHidden', columns, rows: rows.list() }, definedNames };
}

/** A pre-fill value as the cell content for its column: { value } or { note } when Excel can't hold it. */
function cellContent(plan, value) {
	if (value === null || value === undefined || value === '') {
		return null;
	}
	let content = value;
	switch (plan.kind) {
		case 'date':
		case 'dateTime':
			// The time is kept even in a date-only column, so a value exported and read back is unchanged.
			content = isDate(value) ? value : parseNeutralDate(value) || String(value);
			break;
		case 'integer':
		case 'number':
			if (typeof value === 'string') {
				// A double holds 15 significant digits. A longer decimal stays text so that it comes back exactly.
				const text = value.trim();
				const digits = text.replace(/[eE].*$/, '').replace(/[^0-9]/g, '').replace(/^0+/, '');
				const isNumeric = /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(text);
				content = isNumeric && digits.length <= 15 ? Number(text) : value;
			}
			break;
		case 'boolean': {
			const text = foldCase(value);
			content = ['1', 'true', 'yes'].includes(text) ? 'Yes' : ['0', 'false', 'no'].includes(text) ? 'No' : String(value);
			break;
		}
		case 'choice': {
			const entry = findEntry(plan.block.entries, 'value', String(value));
			content = entry ? entry.display : String(value);
			break;
		}
		default:
			content = typeof value === 'string' ? value : String(value);
			break;
	}
	if (typeof content === 'string' && content.length > CELL_TEXT_LIMIT) {
		return {
			note: 'The value has ' + content.length + ' characters, more than an Excel cell can hold. ' +
				'The cell is left blank, so an import leaves the value unchanged.'
		};
	}
	return { value: content };
}

/** A data sheet: the header, a cell for every data cell up to lastRow, validation, formatting and protection. */
function dataSheet(sheetPlan, schema, rows, lastRow) {
	const columns = sheetPlan.columns;
	const sheetRows = createRows();
	// The style of a column also covers the rows below the data, so every cell a user can reach in a data
	// column is unlocked and has the right number format. Everything else on the sheet stays locked.
	const cellStyles = columns.map((plan) => ({ numFmt: plan.format, locked: false }));
	const sheetColumns = columns.map((plan, index) => ({ index: plan.position, width: plan.width, style: cellStyles[index] }));
	// Two empty columns, so that a header note has room to open (see planSheets).
	sheetColumns.push(
		{ index: columns.length + 1, width: 24 },
		{ index: columns.length + 2, width: 24 },
		{ index: sheetPlan.modifiedOnColumn, width: 22, hidden: true, style: MODIFIED_ON_STYLE }
	);

	const header = sheetRows.row(1);
	header.height = 30;
	for (const plan of columns) {
		const note = headerNote(plan, sheetPlan, schema, sheetPlan.all);
		sheetRows.cell(1, plan.position, {
			value: plan.label + (plan.def.required || plan.def.isKey ? ' *' : ''),
			style: plan.def.isKey ? KEY_HEADER_STYLE : HEADER_STYLE,
			note: note || null
		});
	}
	sheetRows.cell(1, sheetPlan.modifiedOnColumn, { value: 'modified_on', style: MODIFIED_ON_STYLE });

	// Every data cell of every row up to lastRow is written, empty or not, with the unlocked style of its
	// column. Excel shows a missing cell as unlocked, but the cell it creates when a value is typed there
	// is locked, so the value could not be changed again. That holds for the rows below the data too.
	for (let rowNumber = 2; rowNumber <= lastRow; rowNumber++) {
		const row = rows[rowNumber - 2];
		const values = (row && row.values) || {};
		columns.forEach((plan, index) => {
			const content = row ? cellContent(plan, values[plan.def.id]) : null;
			sheetRows.cell(rowNumber, plan.position, {
				value: content && !content.note ? content.value : null,
				style: cellStyles[index],
				note: content && content.note ? content.note : null
			});
		});
		if (row && row.modifiedOn !== null && row.modifiedOn !== undefined && row.modifiedOn !== '') {
			// Kept as text, so the server gets back exactly the value it handed out.
			sheetRows.cell(rowNumber, sheetPlan.modifiedOnColumn, {
				value: isDate(row.modifiedOn) ? formatNeutralDate(row.modifiedOn) : String(row.modifiedOn),
				style: MODIFIED_ON_STYLE
			});
		}
	}

	const dataValidations = [];
	const conditionalFormats = [];
	let priority = 1;
	for (const plan of columns) {
		const range = plan.letter + '2:' + plan.letter + lastRow;
		const validation = validationFor(plan, plan.letter + '2');
		if (validation) {
			dataValidations.push({ sqref: range, ...validation });
		}
		const classification = sheetPlan.classification;
		if (plan.def.classPath && classification && classification !== plan) {
			const formula = notApplicableFormula(String(plan.def.classPath), classification.letter);
			if (formula) {
				conditionalFormats.push({
					sqref: range,
					rules: [{ type: 'expression', priority: priority++, formulae: [formula], style: NOT_APPLICABLE_STYLE }]
				});
			}
		}
	}

	return {
		name: sheetPlan.name,
		frozenRows: 1,
		columns: sheetColumns,
		rows: sheetRows.list(),
		dataValidations,
		conditionalFormats,
		// No password: the protection guards against accidents (a renamed or moved column), not against intent.
		// Rows can't be sorted, inserted or deleted, because each row is tied to its cell in the hidden, locked
		// column. A user adds rows at the bottom and empties a row to drop it. On an indented sheet a row's place
		// is its meaning, so rows can be inserted there. A flag that is true protects; what is not listed is
		// protected too, apart from selecting cells.
		protection: {
			sheet: true,
			// Objects include the header notes, which should not be edited or removed either.
			objects: true,
			scenarios: true,
			formatColumns: false,
			formatRows: false,
			...(sheetPlan.def.layout === INDENTED_LAYOUT ? { insertRows: false } : {})
		}
	};
}

function metaSheet(schema, sheetPlans, job) {
	const present = (value) => value !== null && value !== undefined && value !== '';
	const flag = (value) => (value ? 1 : 0);
	const rows = [];
	const add = (cells) => {
		const row = { index: rows.length + 1, cells: [] };
		cells.forEach((value, index) => {
			if (present(value)) {
				row.cells.push({ column: index + 1, value: typeof value === 'number' ? value : String(value) });
			}
		});
		rows.push(row);
	};
	add([META_MARKER, META_LAYOUT]);
	const values = {
		formatVersion: schema.formatVersion,
		database: schema.database,
		generatedOn: schema.generatedOn,
		language: schema.language,
		definitionHash: schema.definitionHash,
		jobId: job && job.id,
		jobToken: job && job.token
	};
	for (const key of META_VALUE_KEYS) {
		if (present(values[key])) {
			add([key, values[key]]);
		}
	}
	for (const sheetPlan of sheetPlans) {
		const def = sheetPlan.def;
		const relationship = def.relationshipType || null;
		const related = (relationship && relationship.relatedItemType) || null;
		const record = {
			index: sheetPlan.index,
			sheetName: sheetPlan.name,
			kind: def.kind,
			itemTypeId: def.itemType.id,
			itemTypeName: def.itemType.name,
			relationshipTypeId: relationship && relationship.id,
			relationshipTypeName: relationship && relationship.name,
			relatedItemTypeId: related && related.id,
			relatedItemTypeName: related && related.name,
			createOnly: flag(def.createOnly),
			modifiedOnColumn: sheetPlan.modifiedOnColumn,
			// The positions of the sheets, as index is.
			parentSheet: def.kind === 'relationship' ? parentSheetOf(schema, def) + 1 : null,
			relatedSheet: def.kind === 'relationship' && relatedSheetOf(schema, def) !== null ? relatedSheetOf(schema, def) + 1 : null,
			layout: def.layout || null
		};
		add(['sheet', ...META_SHEET_FIELDS.map((field) => record[field])]);
	}
	for (const sheetPlan of sheetPlans) {
		for (const plan of sheetPlan.columns) {
			const record = {
				sheetIndex: sheetPlan.index,
				position: plan.position,
				id: plan.def.id,
				role: plan.def.role,
				property: plan.def.property,
				dataType: plan.def.dataType,
				isKey: flag(plan.def.isKey),
				isMatchKey: flag(plan.def.isMatchKey),
				listColumn: plan.block ? plan.block.column : null
			};
			add(['column', ...META_COLUMN_FIELDS.map((field) => record[field])]);
		}
	}
	return { name: META_SHEET, state: 'veryHidden', rows };
}

/**
 * Builds the workbook for a template schema.
 *
 * @param {object} schema The template schema, format version 1.
 * @param {object} [options]
 * @param {{id: string, token: string}|null} [options.job] The job the workbook belongs to; written to _meta.
 * @param {Array<Array<{values: object, modifiedOn: (string|Date|null)}>>} [options.rows] Pre-fill rows, one
 *   array per schema sheet, in the order of schema.sheets. `values` maps a column id to the stored value:
 *   a date as yyyy-MM-ddTHH:mm:ss, a list value as the stored value, a boolean as 1/0 or true/false.
 * @param {number} [options.maxRows=5000] How many data rows get validation.
 * @returns {Promise<ArrayBuffer>} The .xlsx file.
 */
export async function buildWorkbook(schema, options = {}) {
	checkSchema(schema);
	const { writeXlsx } = await loadWriter();
	const generatedOn = parseNeutralDate(schema.generatedOn || '') || new Date();

	const lists = createListPlanner(schema);
	const sheetPlans = planSheets(schema, lists);
	// A header note names other sheets.
	for (const sheetPlan of sheetPlans) {
		sheetPlan.all = sheetPlans;
	}
	const rowsBySheet = Array.isArray(options.rows) ? options.rows : [];
	const wantedRows = Math.floor(toNumberOrNull(options.maxRows) || DEFAULT_MAX_ROWS);
	const maxRows = Math.min(Math.max(wantedRows, 1), EXCEL_LAST_ROW - 1);

	const dataSheets = sheetPlans.map((sheetPlan, index) => {
		const rows = (Array.isArray(rowsBySheet[index]) ? rowsBySheet[index] : []).slice(0, EXCEL_LAST_ROW - 1);
		return dataSheet(sheetPlan, schema, rows, 1 + Math.max(maxRows, rows.length));
	});
	// After the data sheets: only now is it known which ranges the validation entries use.
	const listed = listsSheet(lists);

	// The visible sheets come first, so the workbook opens on the main sheet. Each target's main
	// sheet is followed by its relationship sheets; with several targets, a target's tabs share a colour.
	const groups = sheetGroups(schema);
	const sheets = [];
	groups.forEach((group, at) => {
		for (const index of group) {
			sheets.push(groups.length > 1 ? { ...dataSheets[index], tabColor: TARGET_TAB_COLORS[at % TARGET_TAB_COLORS.length] } : dataSheets[index]);
		}
	});
	sheets.push(listed.sheet, metaSheet(schema, sheetPlans, options.job || null));

	return writeXlsx({
		creator: 'Mass Create & Edit',
		created: generatedOn,
		modified: generatedOn,
		sheets,
		definedNames: listed.definedNames
	});
}

// ---------------------------------------------------------------------------
// Reading
// ---------------------------------------------------------------------------

let readerPromise = null;

/**
 * Loads the workbook reader (mce-xlsx-read.js) once, with the cache tag this
 * file was loaded with, so that a new version of the tool brings it along.
 * Cells come from it as plain values: a string, a number, a boolean, a Date,
 * or the text of an error; rich text, formulas and links are already their
 * text or result.
 */
function loadReader() {
	if (!readerPromise) {
		const here = new URL(import.meta.url);
		readerPromise = import(new URL('./mce-xlsx-read.js', here).href.split('?')[0] + here.search).catch((error) => {
			// A failed load (a network error, say) must not block a later attempt.
			readerPromise = null;
			throw fail('MCE_LIBRARY', 'The workbook reader (mce-xlsx-read.js) could not be loaded: ' + error.message);
		});
	}
	return readerPromise;
}

/** A failure of the reader as the error parseWorkbook gives. */
function notAWorkbook(error) {
	const result = fail(
		'MCE_NOT_A_WORKBOOK',
		error && error.code === 'XLSX_ENCRYPTED' ? error.message : 'The file is not an Excel workbook (.xlsx) or can\'t be read.'
	);
	result.cause = error;
	return result;
}

/** A sheet of the workbook, or null when the file has no such sheet. */
async function readSheet(workbook, name) {
	try {
		return await workbook.readSheet(name);
	} catch (error) {
		throw notAWorkbook(error);
	}
}

const isBlank = (value) => value === null || value === undefined || (typeof value === 'string' && value.trim() === '');

function readMeta(sheet) {
	const invalid = (message) => fail('MCE_META_INVALID', 'The workbook\'s ' + META_SHEET + ' sheet is damaged: ' + message);
	const cellAt = (rowNumber, position) => sheet.value(rowNumber, position);
	if (cellAt(1, 1) !== META_MARKER) {
		throw invalid('its first cell is not "' + META_MARKER + '".');
	}
	if (cellAt(1, 2) !== META_LAYOUT) {
		throw invalid('layout ' + cellAt(1, 2) + ' is not supported (expected ' + META_LAYOUT + ').');
	}
	const record = (rowNumber, fields) => {
		const result = {};
		fields.forEach((field, index) => {
			const value = cellAt(rowNumber, index + 2);
			result[field] = isBlank(value) ? null : value;
		});
		return result;
	};
	const values = {};
	const sheets = [];
	const columns = [];
	for (let rowNumber = 2; rowNumber <= sheet.lastRow; rowNumber++) {
		const type = cellAt(rowNumber, 1);
		if (type === 'sheet') {
			sheets.push(record(rowNumber, META_SHEET_FIELDS));
		} else if (type === 'column') {
			columns.push(record(rowNumber, META_COLUMN_FIELDS));
		} else if (META_VALUE_KEYS.includes(type)) {
			values[type] = cellAt(rowNumber, 2);
		}
	}
	if (values.formatVersion !== FORMAT_VERSION) {
		throw fail(
			'MCE_META_INVALID',
			'The workbook has format version ' + values.formatVersion + ', and this version of Mass Create & Edit reads version ' +
				FORMAT_VERSION + '. Generate the workbook again.'
		);
	}
	if (sheets.length === 0) {
		throw invalid('it lists no sheets.');
	}
	const isPosition = (value) => Number.isInteger(value) && value >= 1;
	const text = (value) => (value === null || value === undefined ? null : String(value));
	const metaSheets = sheets.map((entry) => {
		if (!isPosition(entry.index) || isBlank(entry.sheetName) || isBlank(entry.kind) || isBlank(entry.itemTypeName)) {
			throw invalid('a sheet record is incomplete.');
		}
		return {
			index: entry.index,
			sheetName: String(entry.sheetName),
			kind: String(entry.kind),
			itemType: { id: text(entry.itemTypeId), name: String(entry.itemTypeName) },
			relationshipType: isBlank(entry.relationshipTypeName) && isBlank(entry.relationshipTypeId) ? null : {
				id: text(entry.relationshipTypeId),
				name: text(entry.relationshipTypeName),
				relatedItemType: isBlank(entry.relatedItemTypeName) && isBlank(entry.relatedItemTypeId)
					? null
					: { id: text(entry.relatedItemTypeId), name: text(entry.relatedItemTypeName) }
			},
			createOnly: entry.createOnly === 1,
			modifiedOnColumn: isPosition(entry.modifiedOnColumn) ? entry.modifiedOnColumn : null,
			layout: entry.layout === INDENTED_LAYOUT ? INDENTED_LAYOUT : null,
			columns: []
		};
	});
	for (const entry of columns) {
		const owner = metaSheets.find((candidate) => candidate.index === entry.sheetIndex);
		if (!owner || !isPosition(entry.position) || isBlank(entry.id) || isBlank(entry.dataType)) {
			throw invalid('a column record is incomplete.');
		}
		if (owner.columns.some((column) => column.id === String(entry.id) || column.position === entry.position)) {
			throw invalid('column "' + entry.id + '" is listed twice.');
		}
		owner.columns.push({
			id: String(entry.id),
			role: text(entry.role),
			property: text(entry.property),
			dataType: String(entry.dataType),
			isKey: entry.isKey === 1,
			isMatchKey: entry.isMatchKey === 1,
			position: entry.position,
			listColumn: isPosition(entry.listColumn) ? entry.listColumn : null
		});
	}
	for (const metaSheet of metaSheets) {
		metaSheet.columns.sort((a, b) => a.position - b.position);
	}
	// The main sheet of each relationship sheet, and that of its related ItemType, by index in sheets.
	const indexAt = (position) => {
		const found = metaSheets.findIndex((candidate) => candidate.index === position);
		if (found < 0) {
			throw invalid('a sheet record names a sheet that is not listed.');
		}
		return found;
	};
	sheets.forEach((entry, at) => {
		const metaSheet = metaSheets[at];
		const relationship = metaSheet.kind === 'relationship';
		if (relationship && !isPosition(entry.parentSheet)) {
			throw invalid('a relationship sheet record names no parent sheet.');
		}
		metaSheet.parentSheet = relationship ? indexAt(entry.parentSheet) : null;
		metaSheet.relatedSheet = relationship && isPosition(entry.relatedSheet) ? indexAt(entry.relatedSheet) : null;
	});
	return {
		formatVersion: values.formatVersion,
		database: text(values.database),
		generatedOn: text(values.generatedOn),
		language: text(values.language),
		definitionHash: text(values.definitionHash),
		job: isBlank(values.jobId) && isBlank(values.jobToken) ? null : { id: text(values.jobId), token: text(values.jobToken) },
		sheets: metaSheets
	};
}

/** The entries of one block of the _lists sheet: the text shown and the stored value, per row. */
function readListBlock(listsSheet, column) {
	const entries = [];
	if (!listsSheet) {
		return entries;
	}
	for (let rowNumber = 2; rowNumber <= listsSheet.lastRow; rowNumber++) {
		const display = listsSheet.value(rowNumber, column);
		const value = listsSheet.value(rowNumber, column + 1);
		if (!isBlank(display) && !isBlank(value)) {
			entries.push({ display: String(display), value: String(value) });
		}
	}
	return entries;
}

/** The value of a data cell as the server gets it, or undefined for a blank cell. */
function readCell(value, column, context) {
	if (isDate(value)) {
		return formatNeutralDate(value);
	}
	if (isBlank(value)) {
		return undefined;
	}
	if (column.listColumn !== null) {
		// The sheet holds the text the drop-down shows; the server gets the stored value.
		const entry = findEntry(context.listEntries(column.listColumn), 'display', String(value));
		return entry ? entry.value : value;
	}
	if (column.dataType === 'date' && isFiniteNumber(value)) {
		// A date whose cell lost its date format (after a paste, say) arrives as a bare serial number.
		return formatNeutralDate(serialToDate(value, context.date1904));
	}
	if (column.dataType === 'boolean' && typeof value === 'string') {
		const text = foldCase(value);
		return text === 'yes' ? true : text === 'no' ? false : value;
	}
	return value;
}

/** The columns an indented sheet goes on with: its parent key in place of Level, named like its related key. */
function flatColumns(columns) {
	const parent = columns.filter((column) => column.role === 'relatedKey').map((column) =>
		({ ...column, id: parentCellOf(column.id), role: 'parentKey', isKey: true, isMatchKey: true, listColumn: null }));
	return [...parent, ...columns.filter((column) => column.role !== 'level')];
}

/** The cell of a row's parent that holds the given related-key cell of the parent's own row. */
const parentCellOf = (relatedId) => relatedId.replace(/^related\./, 'parent.');

/** The level of a row: a whole number from 0, or null when the cell holds anything else. */
function levelOf(value) {
	const text = typeof value === 'string' ? value.trim() : value;
	const level = typeof text === 'number' ? text : /^\d+$/.test(String(text)) ? Number(text) : NaN;
	return Number.isInteger(level) && level >= 0 ? level : null;
}

const SAME_LINE = '\u001F';
const SEQUENCE_COLUMN = 'sort_order';

/** What a line says, apart from where it sits: every cell but the level, in a fixed order. */
const lineText = (cells) => Object.keys(cells).filter((id) => id !== 'level').sort()
	.map((id) => id + '=' + foldCase(cells[id])).join(SAME_LINE);

/**
 * Turns the rows of an indented sheet into ordinary rows, each naming its parent and its related item.
 *
 * A row's parent is the nearest row above it one level up. A level 0 row is a top assembly: it names a
 * parent and is no line of its own, so it does not come back. A subassembly can stand in several places: its
 * lines are written once, and a later place that lists none refers to the first. Places that list different
 * lines can't be told apart, so those lines are errors. Rows that can't be placed come back without a parent,
 * so that the server refuses them; `problems` has what to tell the user instead (addIndentedResult). A row
 * below a row with a problem has a problem too. Row numbers stay those of the sheet.
 *
 * @param {Array<{rowNumber: number, cells: object, modifiedOn: (string|null)}>} rows The rows, in sheet order.
 * @param {string[]} relatedIds The columns of the related key.
 * @param {Object<string, string>} [labels] The heading of each column, for the messages.
 * @param {string[]} [matchIds] The columns that, with the parent, tell the lines apart: the match key without
 *   the parent. Without it, the related key.
 * @returns {{rows: Array<object>, indented: {problems: object, notes: object, parents: object, comparisons: Array<object>}}}
 *   Each row may have `sequenceIfNew`: its place among the lines of its parent, for a new line with no Sequence.
 *   By row number: problems { code, message, comparison }, notes the warnings to show, and parents the rows
 *   of the lines above a line, with null for a top assembly. comparisons holds, for each part that stands in
 *   places with different lines, { part, places: [{ row, under }], lines: [{ part, differs, cells }] }: a line
 *   has a cell per place, null where the place does not list it, else { row, values: [{ label, value }] }.
 */
export function flattenIndented(rows, relatedIds, labels = {}, matchIds = relatedIds) {
	const keyOf = (cells) => relatedIds.map((id) => (cells[id] === undefined ? '' : foldCase(cells[id]))).join(SAME_LINE);
	const partOf = (node) => relatedIds.map((id) => node.row.cells[id]).filter((value) => value !== undefined).join(' ');
	const lineKeyOf = (cells) => matchIds.map((id) => (cells[id] === undefined ? '' : foldCase(cells[id]))).join(SAME_LINE);
	const labelOf = (id) => labels[id] || id;
	const isDelete = (cells) => typeof cells.delete === 'string' && cells.delete.trim().toUpperCase() === DELETE_TOKEN;
	const problems = {};
	const reject = (node, code, message) => {
		if (!node.failed) {
			node.failed = true;
			problems[node.row.rowNumber] = { code, message };
		}
	};
	const nodes = [];
	const places = new Map();
	let path = [];
	// After a row whose level can't be read, the rows below this level have no certain parent.
	let unknownBelow = null;
	for (const row of rows) {
		const node = { row, key: keyOf(row.cells), lineKey: lineKeyOf(row.cells), parent: null, lines: [], failed: false };
		nodes.push(node);
		const level = levelOf(row.cells.level);
		if (level === null) {
			reject(node, 'LEVEL_INVALID', 'The level must be a whole number, 0 or more.');
			unknownBelow = unknownBelow === null ? path.length - 1 : unknownBelow;
			continue;
		}
		if (level === 0 || (unknownBelow !== null && level <= unknownBelow)) {
			unknownBelow = null;
		}
		const parent = level > 0 ? path[level - 1] : null;
		node.parent = parent || null;
		if (node.key.split(SAME_LINE).every((part) => part === '')) {
			reject(node, 'LEVEL_NO_PART', 'The row names no part.');
		} else if (unknownBelow !== null) {
			reject(node, 'LEVEL_NO_PARENT', 'A row above has a level that can\'t be read, so the parent of this row is not known.');
		} else if (level > 0 && !parent) {
			reject(node, 'LEVEL_JUMP', path.length === 0
				? 'The first row must be level 0, a top assembly.'
				: 'Level ' + level + ' can\'t follow level ' + (path.length - 1) + ': no row above is its parent.');
		} else if (parent && parent.failed) {
			reject(node, 'LEVEL_PARENT_FAILED', 'Row ' + parent.row.rowNumber + ', its parent, has a problem.');
		} else if (parent && matchIds.includes(SEQUENCE_COLUMN) && row.cells[SEQUENCE_COLUMN] === undefined && !isDelete(row.cells)) {
			reject(node, 'BOM_SEQUENCE_REQUIRED', (labels[SEQUENCE_COLUMN] || 'Sequence') + ' is one of the columns that match this line, so it can\'t be blank.');
		} else if (!isDelete(row.cells)) {
			for (let above = parent; above; above = above.parent) {
				if (above.key === node.key) {
					const between = [];
					for (let step = parent; step !== above; step = step.parent) {
						between.unshift(step);
					}
					reject(node, 'BOM_CYCLE', 'Rows ' + above.row.rowNumber + ', ' + node.row.rowNumber + ' are the same part, one inside the other: '
						+ [above, ...between, node].map(partOf).join(' > ') + '.');
					break;
				}
			}
		}
		path.length = Math.min(path.length, level);
		path[level] = node;
		if (parent) {
			parent.lines.push(node);
		}
		if (!node.failed) {
			places.set(node.key, (places.get(node.key) || []).concat(node));
		}
	}
	const goodLines = (node) => node.lines.filter((line) => !line.failed);
	const failBelow = () => {
		for (const node of nodes) {
			if (node.parent && node.parent.failed) {
				reject(node, 'LEVEL_PARENT_FAILED', 'Row ' + node.parent.row.rowNumber + ', its parent, has a problem.');
			}
		}
	};

	// The same child twice directly under one place: nothing tells the two lines apart.
	for (const node of nodes) {
		const byChild = new Map();
		for (const line of goodLines(node)) {
			byChild.set(line.lineKey, (byChild.get(line.lineKey) || []).concat(line));
		}
		for (const same of byChild.values()) {
			if (same.length > 1) {
				const numbers = same.map((line) => line.row.rowNumber).join(', ');
				const matchedOnPart = matchIds.length === relatedIds.length && matchIds.every((id) => relatedIds.includes(id));
				const text = matchedOnPart
					? 'Rows ' + numbers + ' are the same part under the same parent. Put them on one row with the total quantity.'
					: 'Rows ' + numbers + ' are the same line under the same parent, because they match on ' + (matchIds.map(labelOf).join(', ') || 'the parent alone') + '. Change one of them, or put them on one row.';
				same.forEach((line) => reject(line, 'BOM_SAME_CHILD', text));
			}
		}
	}

	// Parts that contain each other across places: what is left after taking off leaves and tops.
	const edges = [];
	for (const node of nodes) {
		if (node.parent && !node.failed && !isDelete(node.row.cells)) {
			edges.push({ from: node.parent.key, to: node.key, node });
		}
	}
	let inRing = edges;
	for (let trimmed = true; trimmed;) {
		const froms = new Set(inRing.map((edge) => edge.from));
		const tos = new Set(inRing.map((edge) => edge.to));
		const left = inRing.filter((edge) => froms.has(edge.to) && tos.has(edge.from));
		trimmed = left.length !== inRing.length;
		inRing = left;
	}
	const edgesFrom = new Map();
	inRing.forEach((edge) => edgesFrom.set(edge.from, (edgesFrom.get(edge.from) || []).concat(edge)));
	// The shortest way round through an edge, from its lowest row; null when the edge only joins two rings.
	const ringThrough = (start) => {
		const cameBy = new Map([[start.to, start]]);
		const queue = [start.to];
		for (let next = 0; next < queue.length && !cameBy.has(start.from); next++) {
			for (const edge of edgesFrom.get(queue[next]) || []) {
				if (!cameBy.has(edge.to)) {
					cameBy.set(edge.to, edge);
					queue.push(edge.to);
				}
			}
		}
		if (!cameBy.has(start.from)) {
			return null;
		}
		const ring = [];
		for (let edge = cameBy.get(start.from); edge !== start; edge = cameBy.get(edge.from)) {
			ring.unshift(edge);
		}
		ring.unshift(start);
		const rowOf = (edge) => edge.node.row.rowNumber;
		const lowest = ring.reduce((best, edge, index) => (rowOf(edge) < rowOf(ring[best]) ? index : best), 0);
		return [...ring.slice(lowest), ...ring.slice(0, lowest)];
	};
	const ringOf = new Map(inRing.map((edge) => [edge, ringThrough(edge)]));
	for (const [edge, ring] of ringOf) {
		if (ring) {
			const numbers = ring.map((step) => step.node.row.rowNumber).sort((a, b) => a - b).join(', ');
			const names = [ring[0].node.parent, ...ring.map((step) => step.node)].map(partOf).join(' > ');
			reject(edge.node, 'BOM_CYCLE', 'Rows ' + numbers + ' make parts contain each other: ' + names + '.');
		}
	}
	failBelow();

	// A part in several places: the places that list lines must list the same ones.
	const comparisons = [];
	// The components below each place, side by side: in each line, only the values that differ between the places.
	const comparePlaces = (listing) => {
		const samples = new Map();
		listing.forEach((place) => goodLines(place).forEach((line) => {
			if (!samples.has(line.lineKey)) {
				samples.set(line.lineKey, line);
			}
		}));
		const textOf = (line, id) => (line.row.cells[id] === undefined ? '' : foldCase(line.row.cells[id]));
		// What the line is known by: its part when the part is matched on, and any other columns matched on.
		const lineName = (line) => {
			const extras = matchIds.filter((id) => !relatedIds.includes(id)).map((id) => labelOf(id) + ': ' + (line.row.cells[id] === undefined ? 'blank' : line.row.cells[id]));
			return [...(relatedIds.every((id) => matchIds.includes(id)) || matchIds.length === 0 ? [partOf(line)] : []), ...extras].join(', ');
		};
		const lines = Array.from(samples, ([key, sample]) => {
			const found = listing.map((place) => goodLines(place).find((line) => line.lineKey === key) || null);
			const present = found.filter((line) => line);
			const changed = Array.from(new Set(present.flatMap((line) => Object.keys(line.row.cells))))
				.filter((id) => id !== 'level' && !matchIds.includes(id) && new Set(present.map((line) => textOf(line, id))).size > 1);
			return {
				part: lineName(sample),
				differs: present.length < listing.length || changed.length > 0,
				cells: found.map((line) => line && {
					row: line.row.rowNumber,
					values: changed.map((id) => ({ label: labelOf(id), value: line.row.cells[id] === undefined ? 'blank' : String(line.row.cells[id]) }))
				})
			};
		});
		return { part: partOf(listing[0]), places: listing.map((place) => ({ row: place.row.rowNumber, under: place.parent ? partOf(place.parent) : null })), lines };
	};
	const firstOf = new Map();
	for (const [key, list] of places) {
		const listing = list.filter((node) => !node.failed && goodLines(node).length > 0);
		if (listing.length === 0) {
			continue;
		}
		const textOf = (node) => goodLines(node).map((line) => lineText(line.row.cells)).sort().join('\u001E');
		if (listing.some((node) => textOf(node) !== textOf(listing[0]))) {
			const numbers = listing.map((node) => node.row.rowNumber).join(', ');
			const message = 'Part ' + partOf(listing[0]) + ' is on rows ' + numbers + ' with different components below it, but a part can have only one BOM. '
				+ 'Make the rows below them match, or list the components below only one of them.';
			comparisons.push(comparePlaces(listing));
			for (const place of listing) {
				for (const line of goodLines(place)) {
					reject(line, 'BOM_DIFFERENT_PLACES', message);
					problems[line.row.rowNumber].comparison = comparisons.length - 1;
				}
			}
		} else {
			firstOf.set(key, listing[0]);
		}
	}
	failBelow();

	// The line of the first place stands for the same line of the other places.
	const keptRow = new Map();
	const also = new Map();
	const keptLines = new Set();
	for (const [key, first] of firstOf) {
		const lines = new Map(goodLines(first).map((line) => [line.lineKey, line]));
		lines.forEach((line) => keptLines.add(line));
		for (const place of places.get(key)) {
			if (place === first || place.failed) {
				continue;
			}
			for (const line of goodLines(place)) {
				const kept = lines.get(line.lineKey);
				keptRow.set(line.row.rowNumber, kept.row.rowNumber);
				also.set(kept, (also.get(kept) || []).concat(line.row.rowNumber));
			}
		}
	}

	const flat = [];
	const notes = {};
	const parents = {};
	for (const node of nodes) {
		const { level, ...cells } = node.row.cells;
		if (node.failed) {
			flat.push({ rowNumber: node.row.rowNumber, cells, modifiedOn: node.row.modifiedOn });
			continue;
		}
		if (!keptLines.has(node)) {
			continue;
		}
		const first = firstOf.get(node.parent.key);
		const out = { rowNumber: node.row.rowNumber, cells: { ...cells }, modifiedOn: node.row.modifiedOn };
		relatedIds.forEach((id) => {
			if (node.parent.row.cells[id] !== undefined) {
				out.cells[parentCellOf(id)] = node.parent.row.cells[id];
			}
		});
		if (cells.sort_order === undefined) {
			out.sequenceIfNew = (goodLines(first).indexOf(node) + 1) * 10;
		}
		flat.push(out);
		parents[node.row.rowNumber] = Array.from(new Set(places.get(node.parent.key).filter((place) => !place.failed)
			.map((place) => (place.parent ? keptRow.get(place.row.rowNumber) || place.row.rowNumber : null))));
		if (also.has(node)) {
			notes[node.row.rowNumber] = ['The same line is also on row' + (also.get(node).length > 1 ? 's ' : ' ') + also.get(node).join(', ') + '. It is written once.'];
		}
	}
	return { rows: flat, indented: { problems, notes, parents, comparisons } };
}

/**
 * Reads a workbook made by buildWorkbook back into plain rows.
 *
 * Values are strings, numbers and booleans. A date is the string
 * yyyy-MM-ddTHH:mm:ss, a drop-down cell is the stored value of the list, and
 * Yes/No is true/false. Blank cells are left out of `cells`, and rows without
 * any value are skipped. Nothing is checked against the Aras definitions.
 *
 * @param {ArrayBuffer|Uint8Array} arrayBuffer The .xlsx file. Its bytes may belong to another window (a frame).
 * @returns {Promise<{meta: object, sheets: Array<{sheetName: string, columns: string[],
 *   rows: Array<{rowNumber: number, cells: object, modifiedOn: (string|null)}>}>}>}
 *   A sheet that _meta lists but the file no longer has comes back without rows and with `missing: true`.
 * @throws {Error} With code MCE_NOT_A_WORKBOOK, MCE_META_MISSING or MCE_META_INVALID.
 */
export async function parseWorkbook(arrayBuffer) {
	const { openXlsx } = await loadReader();
	let workbook;
	try {
		workbook = await openXlsx(arrayBuffer);
	} catch (error) {
		throw notAWorkbook(error);
	}
	// Decided from the list of sheets alone, before any sheet is unpacked: a file picked as an
	// attachment is turned away quickly, however large it is.
	if (!workbook.hasSheet(META_SHEET)) {
		throw fail(
			'MCE_META_MISSING',
			'This workbook was not generated by Mass Create & Edit: it has no ' + META_SHEET + ' sheet. ' +
				'Generate a workbook from the job and fill that one in.'
		);
	}
	const meta = readMeta(await readSheet(workbook, META_SHEET));
	const needsLists = meta.sheets.some((entry) => entry.columns.some((column) => column.listColumn !== null));
	const listsSheet = needsLists ? await readSheet(workbook, LISTS_SHEET) : null;
	const dataSheets = new Map();
	for (const entry of meta.sheets) {
		if (!dataSheets.has(entry.sheetName)) {
			dataSheets.set(entry.sheetName, await readSheet(workbook, entry.sheetName));
		}
	}
	const listBlocks = new Map();
	const context = {
		date1904: workbook.date1904,
		listEntries: (column) => {
			if (!listBlocks.has(column)) {
				listBlocks.set(column, readListBlock(listsSheet, column));
			}
			return listBlocks.get(column);
		}
	};

	const sheets = meta.sheets.map((metaSheetEntry) => {
		// An indented sheet is read as it stands, and goes on as an ordinary sheet of parent and related rows.
		const sheetColumns = metaSheetEntry.columns;
		const indented = metaSheetEntry.layout === INDENTED_LAYOUT;
		if (indented) {
			metaSheetEntry.columns = flatColumns(sheetColumns);
		}
		const result = { sheetName: metaSheetEntry.sheetName, columns: metaSheetEntry.columns.map((column) => column.id), rows: [] };
		const sheet = dataSheets.get(metaSheetEntry.sheetName);
		if (!sheet) {
			result.missing = true;
			return result;
		}
		for (const rowNumber of sheet.rowNumbers()) {
			if (rowNumber === 1) {
				continue;
			}
			const cells = {};
			let hasValue = false;
			for (const column of sheetColumns) {
				const value = readCell(sheet.value(rowNumber, column.position), column, context);
				if (value !== undefined) {
					cells[column.id] = value;
					hasValue = true;
				}
			}
			if (!hasValue) {
				continue;
			}
			let modifiedOn = metaSheetEntry.modifiedOnColumn ? sheet.value(rowNumber, metaSheetEntry.modifiedOnColumn) : null;
			if (isDate(modifiedOn)) {
				modifiedOn = formatNeutralDate(modifiedOn);
			}
			result.rows.push({ rowNumber, cells, modifiedOn: isBlank(modifiedOn) ? null : String(modifiedOn) });
		}
		if (indented) {
			const related = sheetColumns.filter((column) => column.role === 'relatedKey').map((column) => column.id);
			const labels = {};
			for (const column of sheetColumns) {
				const header = sheet.value(1, column.position);
				labels[column.id] = typeof header === 'string' ? header.replace(/\s*\*\s*$/, '') : column.id;
			}
			const matched = sheetColumns.filter((column) => column.isMatchKey && column.role !== 'level').map((column) => column.id);
			Object.assign(result, flattenIndented(result.rows, related, labels, matched));
		}
		return result;
	});

	return {
		meta: {
			formatVersion: meta.formatVersion,
			database: meta.database,
			generatedOn: meta.generatedOn,
			language: meta.language,
			definitionHash: meta.definitionHash,
			job: meta.job,
			sheets: meta.sheets.map((metaSheetEntry) => ({
				sheetName: metaSheetEntry.sheetName,
				kind: metaSheetEntry.kind,
				itemType: metaSheetEntry.itemType,
				relationshipType: metaSheetEntry.relationshipType,
				createOnly: metaSheetEntry.createOnly,
				parentSheet: metaSheetEntry.parentSheet,
				relatedSheet: metaSheetEntry.relatedSheet,
				columns: metaSheetEntry.columns.map(({ id, role, property, dataType, isKey, isMatchKey }) =>
					({ id, role, property, dataType, isKey, isMatchKey })),
				keyColumns: metaSheetEntry.columns.filter((column) => column.isKey).map((column) => column.id),
				matchKeyColumns: metaSheetEntry.columns.filter((column) => column.isMatchKey).map((column) => column.id)
			}))
		},
		sheets
	};
}
