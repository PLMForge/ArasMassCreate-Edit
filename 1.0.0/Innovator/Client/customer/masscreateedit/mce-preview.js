/**
 * Mass Create & Edit - the item preview of the import step.
 *
 * Beside the grid of an uploaded workbook, the preview pane shows the item a
 * row stands for:
 *
 *   - the item's own Form, read-only, filled with the row's values. A new row
 *     shows a new item; an update row shows the item as it is in Aras now,
 *     read as the user, with the row's changes applied; a row that changes
 *     nothing shows the item as it is;
 *   - below the Form, as the relationships of an item are, a tab for each
 *     relationship sheet of the item's main sheet (a workbook has a main
 *     sheet for each target ItemType) on which the import changes the item's
 *     rows. A tab lists the item's relationship rows as they will be after
 *     the import: the existing rows the import leaves alone, the existing
 *     rows it changes and the new rows, each marked as such. The tabs fold
 *     away, and stay folded until they are opened again.
 *
 * A row of the main sheet with an error has no preview: the pane shows the
 * row's messages. A row of a relationship sheet shows its parent item, with
 * that sheet's tab open and the row marked; a relationship row with an error
 * shows its parent as well, with the row's messages, and is not on the tab,
 * since it is not imported. Only when the parent can't be shown (it has
 * errors itself, or is not known) does the pane show messages alone.
 *
 * The Form marks the fields an update row changes with a subtle red edge,
 * and every field of a new item is new, so none is marked.
 *
 * Above the Form, the pane says only what the grid, the Form and the tabs
 * don't: the errors of a relationship row while its tab is folded away, the
 * components of a subassembly in each place it is used, and changes to
 * properties the Form has no field for. Most rows have none of these, and
 * then nothing is shown above the Form.
 *
 * An imported job keeps its preview, and the pane shows it as it was before
 * the import (context.imported). The relationship rows in Aras then include
 * the rows the import wrote, so they are not read: a tab lists the rows of the
 * workbook only. An update row's Form is built on the item read by the id the
 * preview found, which for a versioned item is the generation before the
 * import.
 *
 * Showing an item has no side effects. Nothing is locked or saved, the item
 * opens in no tab or window, and what the preview reads stays in the pane: the
 * Form is drawn into a frame of its own, where opening items, recording
 * history and opening dialogs do nothing.
 *
 * The pure functions (exported for tests) work on the result of the
 * validation Method FRG_MCE_ValidateImport; the header comment of that Method
 * documents its shape. createPreviewPane() is the part that needs a browser
 * and the Aras client.
 */

// ---------------------------------------------------------------------------
// The result, without the page
// ---------------------------------------------------------------------------

const sheetsOf = (result) => (result && Array.isArray(result.sheets) ? result.sheets : []);

const isRelationshipKey = (columnId) => /^(parent|related)\./.test(String(columnId));

const messagesOf = (row, severity) =>
	((row && Array.isArray(row.issues)) ? row.issues : []).filter((issue) => !severity || issue.severity === severity).map((issue) => issue.message);

/** The relationship sheets of a result, with their index in result.sheets. */
export function comparisonOf(result, sheetIndex, row) {
	const list = (sheetsOf(result)[sheetIndex] || {}).comparisons;
	return row && Array.isArray(list) && list[row.comparison] ? list[row.comparison] : null;
}

export function relationshipSheets(result) {
	return sheetsOf(result)
		.map((sheet, index) => ({ sheet, index }))
		.filter((entry) => entry.sheet && entry.sheet.kind === 'relationship');
}

/** The relationship sheets whose parents are items of the given main sheet. */
export function relationshipSheetsOf(result, mainSheetIndex) {
	return relationshipSheets(result).filter((entry) => entry.sheet.parentSheet === mainSheetIndex);
}

/**
 * What the pane shows for a row of the result.
 *
 * @param {object} result The validation result.
 * @param {number} sheetIndex Index of the sheet in result.sheets.
 * @param {number} rowIndex Index of the row in that sheet's rows (not its row number in Excel).
 * @returns {object} One of
 *   { kind: 'missing' }: there is no such row;
 *   { kind: 'error', sheetIndex, row, messages }: the row has errors and no preview;
 *   { kind: 'item', sheetIndex, row, item, tabSheetIndex, focusRow, warnings, rowErrors }: the item to show.
 *   rowErrors are the errors of a relationship row whose parent is shown although the row itself is not
 *   imported; focusRow is null for such a row, since it is not among the rows of the tab.
 *   item is { action: 'new' | 'update' | 'none', id, keyedName, mainSheet, mainRow, values, display, old, oldDisplay }:
 *   mainSheet is the index in result.sheets of the main sheet the item belongs to (a workbook has one
 *   for each target); mainRow is the row number of the item on that sheet, or null when the
 *   item is not on it (the parent of a relationship row, which exists in Aras).
 *   tabSheetIndex is the relationship sheet whose tab opens first, focusRow the
 *   row number to mark on it; both are null for a main-sheet row.
 */
export function previewTarget(result, sheetIndex, rowIndex) {
	const sheets = sheetsOf(result);
	const sheet = sheets[sheetIndex];
	const row = sheet && Array.isArray(sheet.rows) ? sheet.rows[rowIndex] : undefined;
	if (!row) {
		return { kind: 'missing' };
	}
	const failed = row.action === 'error';
	if (failed && sheet.kind !== 'relationship') {
		return { kind: 'error', sheetIndex, row, messages: messagesOf(row, 'error') };
	}
	const warnings = messagesOf(row, 'warning');
	if (sheet.kind !== 'relationship') {
		return { kind: 'item', sheetIndex, row, item: itemOfMainRow(row, sheetIndex), tabSheetIndex: null, focusRow: null, warnings, rowErrors: [] };
	}

	// A relationship row: its parent is shown. A row with errors is not imported, so it is not marked on
	// the tab; its messages go with the parent.
	const parent = row.parent || {};
	const mainIndex = sheet.parentSheet;
	const mainSheet = sheets[mainIndex];
	const rowErrors = failed ? messagesOf(row, 'error') : [];
	const tab = { tabSheetIndex: sheetIndex, focusRow: failed ? null : row.row, warnings, rowErrors };
	if (parent.row !== null && parent.row !== undefined) {
		const mainRow = (mainSheet && Array.isArray(mainSheet.rows) ? mainSheet.rows : []).find((candidate) => candidate.row === parent.row);
		if (mainRow && mainRow.action !== 'error') {
			return { kind: 'item', sheetIndex, row, item: itemOfMainRow(mainRow, mainIndex), ...tab };
		}
		if (mainRow) {
			return { kind: 'error', sheetIndex, row, messages: failed ? rowErrors : messagesOf(mainRow, 'error') };
		}
	}
	if (parent.id) {
		const item = { action: 'none', id: parent.id, keyedName: parent.keyedName || '', mainSheet: mainIndex, mainRow: null, values: {}, display: {}, old: {}, oldDisplay: {} };
		return { kind: 'item', sheetIndex, row, item, ...tab };
	}
	return { kind: 'error', sheetIndex, row, messages: failed ? rowErrors : ['The parent of this row is not known, so it can\'t be shown.'] };
}

function itemOfMainRow(row, mainSheet) {
	return {
		action: row.action === 'new' || row.action === 'update' ? row.action : 'none',
		id: row.id || null,
		keyedName: row.keyedName || '',
		mainSheet,
		mainRow: row.row,
		values: row.values || {},
		display: row.display || {},
		old: row.old || {},
		oldDisplay: row.oldDisplay || {}
	};
}

/** The rows of a relationship sheet that belong to the shown item. */
export function rowsForParent(sheet, item) {
	if (!sheet || sheet.parentSheet !== item.mainSheet) {
		return [];
	}
	return (Array.isArray(sheet.rows) ? sheet.rows : []).filter((row) => {
		const parent = row.parent || {};
		if (item.mainRow !== null && item.mainRow !== undefined) {
			return parent.row === item.mainRow;
		}
		return (parent.row === null || parent.row === undefined) && !!item.id && parent.id === item.id;
	});
}

/**
 * How many relationship rows of the sheet the item has in Aras now: 0 for a
 * new item, the count the validation found when it lists the item, and null
 * when it does not (no row of the sheet names the item), so the rows have to
 * be read to know.
 */
export function existingCount(sheet, item) {
	if (item.action === 'new' || !item.id || !sheet || sheet.parentSheet !== item.mainSheet) {
		return 0;
	}
	const parents = sheet && Array.isArray(sheet.parents) ? sheet.parents : [];
	const entry = parents.find((parent) =>
		(item.mainRow !== null && item.mainRow !== undefined && parent.row === item.mainRow) ||
		((parent.row === null || parent.row === undefined) && parent.id === item.id));
	return entry && Number.isInteger(entry.existing) ? entry.existing : null;
}

/** The columns a relationship tab shows: the relationship's own properties, not the parent and related key columns. */
export function relationshipColumns(sheet) {
	return (sheet && Array.isArray(sheet.columns) ? sheet.columns : [])
		.filter((column) => !isRelationshipKey(column.id) && column.property && column.status !== 'unknown');
}

/**
 * An existing relationship row, as read from Aras, in the shape mergeRelationshipRows takes.
 *
 * @param {object[]} columns The columns of the tab (relationshipColumns).
 * @param {{ id: string, props: object }} item The row as read: props maps a property name to { value, keyedName }.
 * @returns {{ id, related, values, display }} values and display by column id.
 */
export function existingRow(columns, item) {
	const props = item.props || {};
	const related = props.related_id && props.related_id.value
		? { id: props.related_id.value, keyedName: props.related_id.keyedName || props.related_id.value }
		: null;
	const values = {};
	const display = {};
	for (const column of columns) {
		const prop = props[column.property];
		values[column.id] = prop ? String(prop.value || '') : '';
		if (prop && prop.keyedName && prop.value) {
			display[column.id] = prop.keyedName;
		}
	}
	return { id: item.id, related, values, display };
}

const RELATED = 'related_id';

/**
 * The relationship rows of an item as they will be after the import.
 *
 * @param {object} sheet The relationship sheet of the result (its columns are used).
 * @param {object[]} sheetRows The rows of the sheet for this item (rowsForParent).
 * @param {object[]} existing The item's rows in Aras now (existingRow), in the order read.
 * @param {boolean} [workbookOnly] The existing rows were not read, as after the import: the rows of
 *   the sheet are listed as the preview showed them, in the order of the sheet, and none counts as missing.
 * @returns {{ rows: object[], skipped: object[], missing: number }}
 *   rows: { status: 'new' | 'changed' | 'deleted' | 'unchanged', id, sheetRow, related, values, display,
 *           old, oldDisplay, changed: column ids (with 'related_id' when the related item changes),
 *           missing: true for a matched row that was not found in Aras any more, warnings }
 *   in the order of the Sequence column when the sheet has one, else existing rows first;
 *   skipped: the rows of the sheet with errors, { row, messages }: they are not imported;
 *   missing: how many matched rows were not found among the existing ones.
 */
export function mergeRelationshipRows(sheet, sheetRows, existing, workbookOnly = false) {
	const columns = relationshipColumns(sheet);
	const byId = new Map();
	for (const row of sheetRows) {
		if (row.id && row.action !== 'new') {
			byId.set(row.id, row);
		}
	}
	const used = new Set();
	const rows = [];
	const skipped = [];
	let missing = 0;

	for (const current of existing) {
		const row = byId.get(current.id);
		if (!row) {
			rows.push(finalRow('unchanged', current, null));
			continue;
		}
		used.add(row);
		if (row.action === 'update') {
			rows.push(finalRow('changed', current, row));
		} else if (row.action === 'delete') {
			rows.push(finalRow('deleted', current, row));
		} else {
			rows.push(finalRow('unchanged', current, row));
		}
	}

	for (const row of sheetRows) {
		if (row.action === 'error') {
			skipped.push({ row: row.row, messages: messagesOf(row, 'error') });
			continue;
		}
		if (row.action === 'new') {
			rows.push(finalRow('new', { id: null, related: row.related || null, values: {}, display: {} }, row));
			continue;
		}
		if (!used.has(row)) {
			// Matched by the validation, but no longer among the rows read now; or none were read.
			const shown = finalRow(row.action === 'update' ? 'changed' : row.action === 'delete' ? 'deleted' : 'unchanged', { id: row.id || null, related: row.related || null, values: {}, display: {} }, row);
			if (!workbookOnly) {
				missing++;
				shown.missing = true;
			}
			rows.push(shown);
		}
	}

	const sequence = columns.find((column) => column.property === 'sort_order');
	if (workbookOnly) {
		// Rows the import left alone have no Sequence of their own here: the rows keep the order of the sheet.
		rows.sort((a, b) => a.sheetRow - b.sheetRow);
	} else if (sequence) {
		const key = (row) => {
			const text = String(row.values[sequence.id] === undefined ? '' : row.values[sequence.id]).trim();
			const number = Number(text);
			return text === '' || !Number.isFinite(number) ? Infinity : number;
		};
		// Array.prototype.sort is stable, so rows with the same Sequence keep their order.
		rows.sort((a, b) => key(a) - key(b));
	}
	return { rows, skipped, missing };
}

/** One row of a relationship tab: the existing values with the sheet row's values applied. */
function finalRow(status, current, row) {
	const values = { ...current.values };
	const display = { ...current.display };
	const changed = [];
	let related = current.related;
	// Only a row that is imported changes anything: one with errors leaves the existing row as it is.
	const rowValues = row && (row.action === 'new' || row.action === 'update') ? row.values || {} : {};
	const rowDisplay = row ? row.display || {} : {};
	for (const [columnId, value] of Object.entries(rowValues)) {
		if (columnId === RELATED) {
			related = row.related || null;
			changed.push(RELATED);
			continue;
		}
		values[columnId] = value === null || value === undefined ? '' : String(value);
		if (rowDisplay[columnId] !== undefined && rowDisplay[columnId] !== null) {
			display[columnId] = rowDisplay[columnId];
		} else {
			delete display[columnId];
		}
		if (status !== 'new') {
			changed.push(columnId);
		}
	}
	if (status === 'new' && row) {
		related = row.related || null;
	}
	return {
		status,
		id: current.id || (row && row.id) || null,
		sheetRow: row ? row.row : null,
		related,
		values,
		display,
		old: status === 'changed' && row ? row.old || {} : {},
		oldDisplay: status === 'changed' && row ? row.oldDisplay || {} : {},
		changed,
		missing: false,
		warnings: row ? messagesOf(row, 'warning') : []
	};
}

/** What the import does to an item's rows on one sheet, for the label of its tab: { new, changed, deleted, errors }. */
export function tabSummary(sheetRows) {
	const summary = { new: 0, changed: 0, deleted: 0, errors: 0 };
	for (const row of sheetRows) {
		if (row.action === 'new') {
			summary.new++;
		} else if (row.action === 'update') {
			summary.changed++;
		} else if (row.action === 'delete') {
			summary.deleted++;
		} else if (row.action === 'error') {
			summary.errors++;
		}
	}
	return summary;
}

/**
 * The relationship sheets that have a tab for the item shown: those on which the import adds, changes
 * or deletes rows of the item, or has rows of it with errors, and the sheet of the row chosen, whatever
 * that row does.
 *
 * @param {object} result The validation result.
 * @param {object} item The item shown (previewTarget).
 * @param {number|null} [focusSheetIndex] The sheet of the relationship row chosen (tabSheetIndex), if any.
 * @returns {{ sheet, index, summary }[]} In the order of the result; summary as tabSummary gives it.
 */
export function changedRelationshipSheets(result, item, focusSheetIndex = null) {
	return relationshipSheets(result)
		.map((entry) => ({ ...entry, summary: tabSummary(rowsForParent(entry.sheet, item)) }))
		.filter(({ index, summary }) => index === focusSheetIndex || summary.new + summary.changed + summary.deleted + summary.errors > 0);
}

/**
 * The properties of the item a row sets on the Form: [{ name, value, keyedName, isItem }].
 * keyedName is the text an item property shows; it is null for other properties.
 */
export function formProperties(sheet, values, display) {
	const columns = new Map((sheet && Array.isArray(sheet.columns) ? sheet.columns : []).map((column) => [column.id, column]));
	const properties = [];
	for (const [columnId, value] of Object.entries(values || {})) {
		const column = columns.get(columnId);
		if (!column || !column.property || column.status === 'unknown' || column.status === 'typeChanged' || isRelationshipKey(columnId)) {
			continue;
		}
		// A file cell holds the File's id, or a picked file's token; display has its name.
		const isItem = column.dataType === 'item' || column.dataType === 'file';
		const text = value === null || value === undefined ? '' : String(value);
		properties.push({
			name: column.property,
			value: text,
			keyedName: isItem ? (display && display[columnId] ? String(display[columnId]) : '') : null,
			isItem
		});
	}
	return properties;
}

/** The labels of the columns an update row changes, in sheet order. */
export function changedLabels(sheet, values) {
	const keys = new Set(Object.keys(values || {}));
	return (sheet && Array.isArray(sheet.columns) ? sheet.columns : [])
		.filter((column) => keys.has(column.id))
		.map((column) => column.label || column.property || column.id);
}

/**
 * The properties an update row changes, for the marks on the Form: [{ property, label, was }].
 * was is the value in Aras now as the user reads it (its display text when there is one), '' when empty.
 * A new item or a row that changes nothing has none.
 */
export function changedProperties(sheet, item) {
	if (!item || item.action !== 'update') {
		return [];
	}
	const values = item.values || {};
	const old = item.old || {};
	const oldDisplay = item.oldDisplay || {};
	return (sheet && Array.isArray(sheet.columns) ? sheet.columns : [])
		.filter((column) => Object.prototype.hasOwnProperty.call(values, column.id) && column.property && isPropertyName(column.property) &&
			column.status !== 'unknown' && column.status !== 'typeChanged' && !isRelationshipKey(column.id))
		.map((column) => {
			const was = oldDisplay[column.id] !== undefined && oldDisplay[column.id] !== null ? oldDisplay[column.id] : old[column.id];
			return { property: column.property, label: column.label || column.property, was: was === undefined || was === null ? '' : String(was) };
		});
}

/** A name for an item that has none yet: the values of its key columns. */
export function newItemName(sheet, values) {
	const keys = sheet && Array.isArray(sheet.keyColumns) ? sheet.keyColumns : [];
	return keys.map((key) => (values && values[key] !== undefined ? String(values[key]) : '')).filter((text) => text !== '').join(' / ');
}

const escapeXml = (text) =>
	String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');

const isPropertyName = (name) => /^[A-Za-z_][A-Za-z0-9_]*$/.test(name);

/**
 * The XML of a new item for the preview. It is marked temporary (isTemp), as
 * an item that was never saved.
 *
 * @param {{ name: string, id: string }} itemType
 * @param {string} id The id the preview gives the item.
 * @param {object[]} properties As formProperties gives them, each optionally with `type` (the ItemType an item property points to).
 */
export function newItemXml(itemType, id, properties) {
	let body = '';
	for (const property of properties) {
		if (!isPropertyName(property.name)) {
			continue;
		}
		let attributes = '';
		if (property.isItem && property.value !== '') {
			attributes += ' keyed_name="' + escapeXml(property.keyedName || '') + '"';
			if (property.type) {
				attributes += ' type="' + escapeXml(property.type) + '"';
			}
		}
		body += '<' + property.name + attributes + '>' + escapeXml(property.value) + '</' + property.name + '>';
	}
	return '<Item type="' + escapeXml(itemType.name) + '"' + (itemType.id ? ' typeId="' + escapeXml(itemType.id) + '"' : '') +
		' id="' + escapeXml(id) + '" action="add" isTemp="1" isDirty="1">' + body + '</Item>';
}

/**
 * The default values of an ItemType's properties that a new row leaves empty, the
 * way the server fills them in. Item properties are left out: their keyed name
 * is not known without asking the server.
 *
 * @param {object[]} definitions [{ name, dataType, defaultValue }]
 * @param {Set<string>|string[]} filled The property names the row sets.
 */
export function defaultProperties(definitions, filled) {
	const set = filled instanceof Set ? filled : new Set(filled || []);
	const skip = new Set(['item', 'sequence', 'foreign', 'image', 'md5', 'federated']);
	return definitions
		.filter((definition) => definition.defaultValue !== null && definition.defaultValue !== undefined && String(definition.defaultValue) !== '')
		.filter((definition) => !set.has(definition.name) && !skip.has(definition.dataType) && isPropertyName(definition.name))
		.map((definition) => ({ name: definition.name, value: String(definition.defaultValue), keyedName: null, isItem: false }));
}

// ---------------------------------------------------------------------------
// The pane
// ---------------------------------------------------------------------------

const STYLE_ID = 'mce-preview-style';
const FORM_HOST = '/Modules/formPreview/formPreview.html';
const KEPT_READS = 40;
const SHOWN_ROWS = 2000;
const MARK_WAIT = 5000;

// The marks on the changed fields of the Form. The Form is a page of its own, so the rule goes into it.
// A thin red frame around the field and a faint red tint, readable on light and dark themes. The frame goes
// on the field's own box, label included, so it fits every kind of field (text, list, item, check box).
const REDLINE_STYLE_ID = 'mce-pv-redline-style';
const REDLINE_CLASS = 'mce-pv-redline';
const REDLINE_CSS =
	'.sys_f_container.' + REDLINE_CLASS + ' { outline: 1px solid rgba(224, 69, 63, 0.7); outline-offset: 2px;' +
	' background-color: rgba(224, 69, 63, 0.08); border-radius: 2px; }';

// The style sheet is loaded with the cache tag this file was loaded with.
const cacheTag = new URL(import.meta.url).search;
const sibling = (name) => new URL(name, import.meta.url).href.split('?')[0] + cacheTag;

const STATUS_TEXT = { new: 'New', changed: 'Changed', deleted: 'Deleted', unchanged: 'Unchanged' };
// The words the grid of the upload uses for what a row does.
const ACTION_TEXT = { new: 'Create', update: 'Update', none: 'No change', error: 'Error' };

/**
 * Creates the preview pane in a container element.
 *
 * @param {HTMLElement} container The element the pane fills. The pane owns everything inside it.
 * @param {{ top: Window, job?: object, result: object, imported?: boolean }} context
 *   top: the main window of the client; result: the validation result the grid shows; imported: the
 *   job is imported, and the pane shows the preview as it was before the import (see the header).
 * @returns {{ show(where: { sheetIndex: number, rowIndex: number }): void, clear(): void, dispose(): void }}
 *   show() takes the index of a sheet in result.sheets and of a row in that sheet's rows.
 */
export function createPreviewPane(container, context) {
	const top = context.top;
	const result = context.result;
	const imported = context.imported === true;
	const doc = container.ownerDocument;
	const make = (tag, attributes, ...children) => element(doc, tag, attributes, ...children);

	if (!doc.getElementById(STYLE_ID)) {
		doc.head.append(make('link', { id: STYLE_ID, rel: 'stylesheet', href: sibling('./mce-preview.css') }));
	}

	const pane = {
		disposed: false,
		// Every show(), clear() and dispose() takes a new turn; work of an older turn is dropped when it comes back.
		turn: 0,
		target: null,
		// The relationship tab open now, and the one the user chose last: an item that has that tab opens on it.
		tab: null,
		chosen: null,
		// The relationships are folded away; the pane keeps them so for every item it shows next.
		collapsed: false,
		// The labels of the properties the row changes that have no field on the Form shown.
		unmarked: [],
		host: null,
		reads: new Map(),
		types: new Map(),
		lists: new Map(),
		newIds: new Map()
	};
	/** The main sheet of the item shown: a workbook has one for each target. */
	const mainSheetOf = (item) => sheetsOf(result)[item.mainSheet];
	// The tabs of the item shown: the relationship sheets on which the import changes its rows.
	let shownTabs = [];

	// --- The frame of the pane ---
	//
	//   div.mce-pv
	//     div.mce-pv-head          notes the grid, the Form and the tabs don't show; hidden when there are none
	//     div.mce-pv-form          the Form, in a frame, taking the room that is left
	//     div.mce-pv-acc           the relationships: a bar of tabs with a fold button, and the rows of the tab open
	//     p.mce-pv-acc-none        instead of the tabs, when the import changes no relationship of the item
	//     div.mce-pv-empty         instead of all of the above, when there is no item to show
	const notes = make('div', { class: 'mce-pv-notes' });
	const head = make('div', { class: 'mce-pv-head', hidden: 'hidden' }, notes);
	const formStatus = make('div', { class: 'mce-pv-status', role: 'status', hidden: 'hidden' });
	const formPanel = make('div', { class: 'mce-pv-form' }, formStatus);
	const tabList = make('ul', { class: 'aras-tabs__list', role: 'tablist', 'aria-label': 'Relationships' });
	const toggle = make('button', { type: 'button', class: 'mce-pv-acc-toggle', 'aria-controls': 'mce-pv-acc-body' }, chevron(doc));
	const accordionHead = make('div', { class: 'mce-pv-acc-head' },
		make('div', { class: 'aras-tabs' }, make('div', { class: 'aras-tabs__list-container' }, tabList)), toggle);
	const listPanel = make('div', { class: 'mce-pv-acc-body', id: 'mce-pv-acc-body', role: 'tabpanel', tabindex: '0' });
	const accordion = make('div', { class: 'mce-pv-acc', hidden: 'hidden' }, accordionHead, listPanel);
	const noTabs = make('p', { class: 'mce-pv-acc-none mce-pv-muted', hidden: 'hidden' }, 'This import changes no relationships of this item.');
	const empty = make('div', { class: 'mce-pv-empty' });
	const root = make('div', { class: 'mce-pv' }, head, formPanel, accordion, noTabs, empty);
	container.textContent = '';
	container.append(root);

	const tabButtons = new Map();
	toggle.addEventListener('click', () => {
		setCollapsed(!pane.collapsed);
		loadShownTab(pane.turn);
	});
	setCollapsed(false);

	setEmpty('Choose a row to preview it.');

	/** The tabs of the item shown, each named as the sheet of the workbook is, after the RelationshipType ("Part CAD"). */
	function drawTabs(target) {
		shownTabs = changedRelationshipSheets(result, target.item, target.tabSheetIndex);
		accordion.hidden = shownTabs.length === 0;
		// A workbook without relationship sheets for the item's target has nothing to say about them.
		noTabs.hidden = shownTabs.length > 0 || relationshipSheetsOf(result, target.item.mainSheet).length === 0;
		tabList.textContent = '';
		tabButtons.clear();
		for (const { sheet, index, summary } of shownTabs) {
			const key = String(index);
			const parts = [];
			if (summary.new > 0) {
				parts.push(summary.new + ' new');
			}
			if (summary.changed > 0) {
				parts.push(summary.changed + ' changed');
			}
			if (summary.deleted > 0) {
				parts.push(summary.deleted + ' deleted');
			}
			if (summary.errors > 0) {
				parts.push(summary.errors + ' with errors');
			}
			const tab = make('li', {
				class: 'aras-tabs__tab', role: 'tab', id: 'mce-pv-tab-' + key, 'aria-selected': 'false', 'aria-controls': 'mce-pv-acc-body', tabindex: '-1',
				title: parts.length > 0 ? 'This import: ' + parts.join(', ') + '.' : 'This import changes no rows here.'
			},
			make('span', { class: 'mce-pv-tab-label' }, sheetLabel(sheet)),
			parts.length > 0 ? make('span', { class: 'mce-pv-tab-count' }, '(' + parts.join(', ') + ')') : null);
			// Choosing a tab opens the relationships when they are folded away.
			const choose = () => {
				pane.chosen = key;
				setCollapsed(false);
				selectTab(key);
			};
			tab.addEventListener('click', choose);
			tab.addEventListener('keydown', (event) => {
				const keys = Array.from(tabButtons.keys());
				const at = keys.indexOf(key);
				const next = event.key === 'ArrowRight' ? keys[(at + 1) % keys.length]
					: event.key === 'ArrowLeft' ? keys[(at - 1 + keys.length) % keys.length]
						: event.key === 'Home' ? keys[0]
							: event.key === 'End' ? keys[keys.length - 1] : null;
				if (next !== null) {
					event.preventDefault();
					pane.chosen = next;
					selectTab(next);
					tabButtons.get(next).focus();
				} else if (event.key === 'Enter' || event.key === ' ') {
					event.preventDefault();
					choose();
				}
			});
			tabButtons.set(key, tab);
			tabList.append(tab);
		}
	}

	function setCollapsed(collapsed) {
		pane.collapsed = collapsed;
		accordion.classList.toggle('mce-pv-acc-collapsed', collapsed);
		toggle.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
		toggle.title = collapsed ? 'Show the relationships' : 'Hide the relationships';
		toggle.setAttribute('aria-label', toggle.title);
		// The notes say a row's errors only while its tab is folded away.
		if (pane.target && pane.target.kind === 'item') {
			drawNotes(pane.target);
		}
	}

	// --- Showing a row ---

	function show(where) {
		if (pane.disposed) {
			return;
		}
		const turn = ++pane.turn;
		const target = previewTarget(result, where ? where.sheetIndex : -1, where ? where.rowIndex : -1);
		pane.target = target;
		if (target.kind === 'missing') {
			setEmpty('This row is not part of the upload.');
			return;
		}
		if (target.kind === 'error') {
			showErrors(target);
			return;
		}
		root.classList.remove('mce-pv-is-empty', 'mce-pv-is-error');
		pane.unmarked = [];
		drawTabs(target);
		drawNotes(target);
		// A relationship row opens on its own sheet's tab; a main-sheet row on the tab chosen last, when the
		// item has it, else on its first tab.
		if (target.tabSheetIndex !== null) {
			pane.chosen = String(target.tabSheetIndex);
		}
		selectTab(tabButtons.has(pane.chosen) ? pane.chosen : shownTabs.length > 0 ? String(shownTabs[0].index) : null, turn);
		loadForm(target, turn);
	}

	function clear() {
		if (pane.disposed) {
			return;
		}
		pane.turn++;
		pane.target = null;
		setEmpty('Choose a row to preview it.');
		blankForm();
	}

	function dispose() {
		if (pane.disposed) {
			return;
		}
		pane.disposed = true;
		pane.turn++;
		pane.target = null;
		if (pane.host && pane.host.frame) {
			try {
				// The Form's own scripts are unloaded before the frame goes.
				pane.host.frame.src = 'about:blank';
			} catch (error) {
				// The frame is gone already.
			}
		}
		pane.host = null;
		pane.reads.clear();
		root.remove();
	}

	const current = (turn) => !pane.disposed && pane.turn === turn;

	function setEmpty(message) {
		root.classList.add('mce-pv-is-empty');
		root.classList.remove('mce-pv-is-error');
		notes.textContent = '';
		head.hidden = true;
		empty.textContent = '';
		empty.append(make('p', { class: 'mce-pv-muted' }, message));
	}

	function showErrors(target) {
		root.classList.add('mce-pv-is-empty', 'mce-pv-is-error');
		notes.textContent = '';
		head.hidden = true;
		empty.textContent = '';
		const sheet = sheetsOf(result)[target.sheetIndex] || {};
		empty.append(make('p', { class: 'mce-pv-heading' }, make('span', { class: 'mce-pv-badge mce-pv-badge-error' }, ACTION_TEXT.error),
			sheetLabel(sheet) + ', row ' + target.row.row + ': not imported'));
		empty.append(make('p', { class: 'mce-pv-muted' }, 'A row with errors has no preview. Correct the workbook and upload it again.'));
		const messages = target.messages.length > 0 ? target.messages : ['The row has an error.'];
		empty.append(make('ul', { class: 'mce-pv-errors' }, ...messages.map((message) => make('li', null, message))));
		appendComparison(empty, target);
		const warnings = messagesOf(target.row, 'warning');
		if (warnings.length > 0) {
			empty.append(make('ul', { class: 'mce-pv-warnings' }, ...warnings.map((message) => make('li', null, message))));
		}
		blankForm();
	}

	/** A part used in several places with different components: the components of each place, side by side. */
	function appendComparison(parent, target) {
		const comparison = comparisonOf(result, target.sheetIndex, target.row);
		if (!comparison) {
			return;
		}
		const head = make('tr', null, make('th', { class: 'aras-grid-head-cell', scope: 'col' }, 'Component'), ...comparison.places.map((place) =>
			make('th', { class: 'aras-grid-head-cell', scope: 'col' }, 'Row ' + place.row + (place.under ? ', below ' + place.under : ', top level'))));
		const body = make('tbody', null, ...comparison.lines.map((line) => make('tr', { class: line.differs ? 'mce-pv-compare-differs' : null },
			make('td', null, line.part),
			...line.cells.map((cell) => (cell
				? make('td', { class: cell.row === target.row.row ? 'mce-pv-compare-this' : null }, 'Row ' + cell.row,
					...cell.values.map((value) => make('div', null, value.label + ': ' + value.value)))
				: make('td', { class: 'mce-pv-compare-missing' }, 'Not listed'))))));
		parent.append(
			make('p', { class: 'mce-pv-heading' }, 'Components below ' + comparison.part + ', in each place it is used'),
			make('div', { class: 'mce-pv-compare' }, make('table', { class: 'mce-pv-compare-table' }, make('thead', null, head), body)),
			make('p', { class: 'mce-pv-muted' }, 'Highlighted lines are the ones that differ. The row you chose is in bold.'));
	}

	/**
	 * The notes above the Form: only what the grid, the Form and the tabs don't show. The grid shows the
	 * row's action, name and messages; the Form marks the fields the row changes; a tab lists the rows of
	 * its relationship, with the errors of those not imported. So the notes hold the errors of a
	 * relationship row while its tab is folded away, the components of a subassembly in each place it is
	 * used, and the changes to properties the Form has no field for. Without notes there is no header.
	 */
	function drawNotes(target) {
		notes.textContent = '';
		if (target.tabSheetIndex !== null && target.rowErrors.length > 0) {
			const onTab = !pane.collapsed && shownTabs.some((tab) => tab.index === target.tabSheetIndex);
			if (!onTab) {
				const sheet = sheetsOf(result)[target.tabSheetIndex] || {};
				notes.append(make('p', { class: 'mce-pv-heading' }, sheetLabel(sheet) + ', row ' + target.row.row + ': not imported'));
				notes.append(make('ul', { class: 'mce-pv-errors' }, ...target.rowErrors.map((message) => make('li', null, message))));
			}
			appendComparison(notes, target);
		}
		if (pane.unmarked.length > 0) {
			notes.append(make('p', null, make('strong', null, 'Also changed, with no field on this Form: '), pane.unmarked.join(', ')));
		}
		head.hidden = notes.childElementCount === 0;
	}

	/** Opens a relationship tab of the item shown; null when the item has none. */
	function selectTab(key, turn) {
		if (!tabButtons.has(key)) {
			key = null;
		}
		pane.tab = key;
		for (const [name, tab] of tabButtons) {
			const on = name === key;
			tab.setAttribute('aria-selected', on ? 'true' : 'false');
			tab.tabIndex = on ? 0 : -1;
			tab.classList.toggle('aras-tabs__tab_active', on);
		}
		if (key === null) {
			listPanel.removeAttribute('aria-labelledby');
			listPanel.textContent = '';
			return;
		}
		listPanel.setAttribute('aria-labelledby', 'mce-pv-tab-' + key);
		loadShownTab(turn === undefined ? pane.turn : turn);
	}

	/** Lists the rows of the tab open now, unless the relationships are folded away: they are listed when opened. */
	function loadShownTab(turn) {
		if (!pane.collapsed && pane.tab !== null && pane.target && pane.target.kind === 'item') {
			loadRelationships(pane.target, Number(pane.tab), turn);
		}
	}

	// --- The Form ---

	async function loadForm(target, turn) {
		formStatus.textContent = 'Loading the Form...';
		formStatus.hidden = false;
		try {
			const node = await itemNode(target.item, turn);
			if (!current(turn)) {
				return;
			}
			if (!node) {
				formStatus.textContent = 'The item can\'t be read any more: it was deleted, or you no longer have access to it.';
				blankForm();
				return;
			}
			const host = await formHost();
			if (!current(turn)) {
				return;
			}
			host.show(node);
			formStatus.hidden = true;
			markChanges(host, target, turn);
		} catch (error) {
			if (current(turn)) {
				formStatus.textContent = 'The Form could not be shown: ' + (error && error.message ? error.message : String(error));
				formStatus.hidden = false;
				blankForm();
			}
		}
	}

	/** The item node the Form shows: built for a new row, read and changed for an existing item. */
	async function itemNode(item, turn) {
		const mainSheet = mainSheetOf(item);
		const typeName = mainSheet.itemType.name;
		const aras = top.aras;
		const xmlDoc = aras.createXMLDocument();
		if (item.action === 'new') {
			const newKey = item.mainSheet + ':' + item.mainRow;
			if (!pane.newIds.has(newKey)) {
				pane.newIds.set(newKey, aras.generateNewGUID());
			}
			const properties = withReferenceTypes(typeName, formProperties(mainSheet, item.values, item.display));
			const defaults = defaultProperties(typeOf(typeName).definitions, new Set(properties.map((property) => property.name)));
			xmlDoc.loadXML(newItemXml(mainSheet.itemType, pane.newIds.get(newKey), properties.concat(defaults)));
			return xmlDoc.documentElement;
		}
		const xml = await readItem(typeName, item.id);
		if (!current(turn) || xml === null) {
			return null;
		}
		xmlDoc.loadXML(xml);
		const node = xmlDoc.documentElement;
		for (const property of withReferenceTypes(typeName, formProperties(mainSheet, item.values, item.display))) {
			setProperty(node, property);
		}
		// Marked as never saved, so the Form shows the node as it is built here.
		node.setAttribute('isTemp', '1');
		return node;
	}

	/** Reads an item as the user. The answer is kept for the life of the pane. */
	async function readItem(typeName, id) {
		const key = 'item|' + typeName + '|' + id;
		if (!pane.reads.has(key)) {
			const reading = query(top, '<Item type="' + escapeXml(typeName) + '" action="get" id="' + escapeXml(id) + '"/>')
				.then((nodes) => (nodes.length > 0 ? new top.XMLSerializer().serializeToString(nodes[0]) : null));
			keep(key, reading);
		}
		return pane.reads.get(key);
	}

	function keep(key, promise) {
		pane.reads.set(key, promise);
		promise.catch(() => pane.reads.delete(key));
		while (pane.reads.size > KEPT_READS) {
			pane.reads.delete(pane.reads.keys().next().value);
		}
	}

	function setProperty(node, property) {
		let child = Array.from(node.children).find((candidate) => candidate.localName === property.name && !candidate.prefix);
		if (!child) {
			child = node.ownerDocument.createElement(property.name);
			node.append(child);
		}
		child.textContent = property.value;
		child.removeAttribute('is_null');
		if (property.isItem && property.value !== '') {
			child.setAttribute('keyed_name', property.keyedName || '');
			if (property.type) {
				child.setAttribute('type', property.type);
			}
		} else {
			child.removeAttribute('keyed_name');
		}
	}

	/** The ItemType definition, kept for the life of the pane: property data types, list ids, defaults. */
	function typeOf(typeName) {
		if (!pane.types.has(typeName)) {
			const info = { definitions: [], byName: new Map() };
			try {
				const node = top.aras.getItemTypeForClient(typeName, 'name').node;
				const properties = node ? Array.from(node.selectNodes("Relationships/Item[@type='Property']")) : [];
				for (const property of properties) {
					const text = (name) => top.aras.getItemProperty(property, name) || '';
					const definition = { name: text('name'), dataType: text('data_type'), dataSource: text('data_source'), defaultValue: text('default_value'), pattern: text('pattern') };
					info.definitions.push(definition);
					info.byName.set(definition.name, definition);
				}
			} catch (error) {
				// Without the definition, the preview shows the values as they are.
			}
			pane.types.set(typeName, info);
		}
		return pane.types.get(typeName);
	}

	/** Adds to item properties the name of the ItemType they point to, which the Form needs to show them as links. */
	function withReferenceTypes(typeName, properties) {
		const info = typeOf(typeName);
		return properties.map((property) => {
			const definition = info.byName.get(property.name);
			if (!property.isItem || !definition || !definition.dataSource) {
				return property;
			}
			let type = '';
			try {
				type = top.aras.getItemTypeName(definition.dataSource) || '';
			} catch (error) {
				type = '';
			}
			return type ? { ...property, type } : property;
		});
	}

	/**
	 * The frame the Form is drawn in, made once per pane. It shows the Form
	 * preview page of Innovator, set up for the pane:
	 *   - what the Form reads stays in the pane;
	 *   - opening an item, recording it in the history and opening dialogs do nothing;
	 *   - an error of the Form shows a message in the pane instead of a dialog.
	 */
	function formHost() {
		if (pane.host) {
			return pane.host.ready;
		}
		const frame = make('iframe', { class: 'mce-pv-frame', title: 'Form preview' });
		const host = { frame, ready: null };
		pane.host = host;
		host.ready = new Promise((resolve, reject) => {
			frame.addEventListener('load', () => {
				try {
					resolve(prepareHost(frame.contentWindow));
				} catch (error) {
					reject(error);
				}
			}, { once: true });
		});
		host.ready.catch(() => {
			if (pane.host === host) {
				pane.host = null;
				frame.remove();
			}
		});
		frame.src = top.aras.getBaseURL() + FORM_HOST;
		formPanel.append(frame);
		return host.ready;
	}

	function prepareHost(win) {
		const parentAras = top.aras;
		if (!win || typeof win.Aras !== 'function' || !win.document.getElementById('preview-iframe')) {
			throw new Error('the Form preview page of the client could not be loaded.');
		}
		win.store = win.store || top.store;
		const own = new win.Aras(parentAras);
		own.commonProperties = Object.assign({}, own.commonProperties);
		own.commonProperties.mainWindow = win;
		own.getMostTopWindowWithAras = () => win;
		own.newIOMInnovator = () => parentAras.newIOMInnovator();
		// A cache of its own: what the Form reads stays in the pane.
		if (typeof win.ClientCache === 'function') {
			own.itemsCache = new win.ClientCache(own);
		}
		const nothing = () => false;
		own.uiShowItemEx = nothing;
		own.uiShowItem = nothing;
		own.uiShowItemByID = nothing;
		own.uiOpenWindowEx = nothing;
		own.saveUICommandHistoryIfNeed = () => undefined;
		own.AlertError = (message) => formFailed(message);
		own.AlertWarning = () => undefined;
		own.AlertSuccess = () => undefined;
		own.AlertAlways = () => undefined;
		own.confirm = () => false;
		own.prompt = () => null;
		// The dialogs are turned off only when they belong to the preview page, so that dialogs of
		// the main window keep working.
		if (win.ArasModules && win.ArasModules !== top.ArasModules && win.ArasModules.Dialog !== (top.ArasModules && top.ArasModules.Dialog)) {
			const refuse = () => {
				const refused = Promise.reject(new Error('Dialogs are not available in the preview.'));
				refused.catch(() => undefined);
				return { promise: refused };
			};
			const quiet = () => Promise.resolve(false);
			if (win.ArasModules.Dialog) {
				win.ArasModules.Dialog.show = refuse;
				for (const name of ['alert', 'confirm', 'prompt']) {
					if (typeof win.ArasModules.Dialog[name] === 'function') {
						win.ArasModules.Dialog[name] = quiet;
					}
				}
			}
			for (const name of ['alert', 'confirm', 'prompt']) {
				if (win.ArasModules[name]) {
					win.ArasModules[name] = quiet;
				}
			}
		}
		win.alert = () => undefined;
		win.confirm = () => false;
		win.prompt = () => null;
		win.aras = own;
		if (win.formPreview) {
			win.formPreview.errorMessage = 'The Form could not be shown.';
			win.formPreview.defaultMessage = '';
		}
		const formFrame = win.document.getElementById('preview-iframe');
		const cover = win.document.getElementById('form-preview-disable');
		// The changes marked on the Form, applied again when a Form loads after they were set.
		let marked = [];
		formFrame.addEventListener('load', () => {
			// Dialogs of the Form's own window stay closed as well.
			try {
				const formWindow = formFrame.contentWindow;
				formWindow.alert = () => undefined;
				formWindow.confirm = () => false;
				formWindow.prompt = () => null;
			} catch (error) {
				// A frame that is not the Form's.
			}
			applyMarks(marked);
		});

		function applyMarks(changes) {
			let formDoc;
			try {
				formDoc = formFrame.contentWindow.document;
			} catch (error) {
				return null;
			}
			if (!formDoc || !formDoc.getElementById('MainDataForm')) {
				return null;
			}
			if (!formDoc.getElementById(REDLINE_STYLE_ID)) {
				const style = formDoc.createElement('style');
				style.id = REDLINE_STYLE_ID;
				style.textContent = REDLINE_CSS;
				(formDoc.head || formDoc.documentElement).append(style);
			}
			for (const field of Array.from(formDoc.querySelectorAll('.' + REDLINE_CLASS))) {
				field.classList.remove(REDLINE_CLASS);
				if (field.dataset.mceTitle !== undefined) {
					field.title = field.dataset.mceTitle;
					delete field.dataset.mceTitle;
				}
			}
			const missing = [];
			for (const change of changes) {
				const fields = Array.from(formDoc.querySelectorAll('.sys_f_container.sys_pn_' + change.property));
				for (const field of fields) {
					field.classList.add(REDLINE_CLASS);
					field.dataset.mceTitle = field.title || '';
					field.title = 'Changed by this row. In Aras now: ' + (change.was === '' ? '(empty)' : change.was);
				}
				if (fields.length === 0) {
					missing.push(change);
				}
			}
			return missing;
		}

		return {
			show(node) {
				const shown = own.uiShowItemInFrameEx(formFrame.contentWindow, node, 'view', 0);
				if (shown === false) {
					throw new Error('there is no Form for this item that you may see.');
				}
				// Hides the preview page's cover, so that the Form can be scrolled and clicked. A blank
				// Form (blank()) shows the cover again.
				if (cover) {
					cover.style.display = 'none';
				}
			},
			blank() {
				marked = [];
				if (win.formPreview) {
					win.formPreview.clearForm('');
				}
			},
			/**
			 * Marks the fields of the given properties on the Form shown now, and unmarks all others.
			 * A Form that loads later gets the same marks. Returns the changes that have no field on
			 * the Form, or null when no Form is there yet.
			 */
			mark(changes) {
				marked = changes;
				return applyMarks(changes);
			}
		};
	}
	/**
	 * Marks the changed fields once the Form has drawn them. The Form is drawn after show() returns,
	 * so the marks are tried again for a while. The changes left without a field are then named in
	 * the notes, as nothing else shows them.
	 */
	function markChanges(host, target, turn) {
		const changes = changedProperties(mainSheetOf(target.item), target.item);
		const started = Date.now();
		const attempt = () => {
			if (!current(turn)) {
				return;
			}
			const missing = host.mark(changes);
			if ((missing === null || missing.length > 0) && Date.now() - started < MARK_WAIT) {
				setTimeout(attempt, 150);
				return;
			}
			pane.unmarked = (missing === null ? changes : missing).map((change) => change.label);
			drawNotes(target);
		};
		attempt();
	}

	/** An error the Form reported, shown in the pane instead of a dialog. */
	function formFailed(message) {
		const text = message && typeof message === 'object' && message.message ? message.message : message;
		formStatus.textContent = 'The Form reported a problem' + (typeof text === 'string' && text.trim() !== '' ? ': ' + text.trim() : '.');
		formStatus.hidden = false;
	}

	function blankForm() {
		if (pane.host && pane.host.ready) {
			pane.host.ready.then((host) => host.blank()).catch(() => undefined);
		}
	}

	// --- Relationship tabs ---

	async function loadRelationships(target, sheetIndex, turn) {
		const sheet = sheetsOf(result)[sheetIndex];
		const { item } = target;
		listPanel.textContent = '';
		listPanel.append(make('p', { class: 'mce-pv-muted mce-pv-pad' }, 'Loading...'));
		try {
			const sheetRows = rowsForParent(sheet, item);
			const columns = relationshipColumns(sheet);
			let existing = [];
			// After the import, the rows in Aras include those it wrote: they are not read.
			if (!imported && existingCount(sheet, item) !== 0) {
				existing = await readRelationships(sheet, columns, item.id);
			}
			if (!current(turn) || pane.tab !== String(sheetIndex) || pane.target !== target) {
				return;
			}
			drawRelationships(sheet, columns, item, mergeRelationshipRows(sheet, sheetRows, existing, imported), target.tabSheetIndex === sheetIndex ? target.focusRow : null);
		} catch (error) {
			if (current(turn) && pane.tab === String(sheetIndex)) {
				listPanel.textContent = '';
				listPanel.append(make('p', { class: 'mce-pv-error mce-pv-pad' },
					'The rows could not be read: ' + (error && error.message ? error.message : String(error))));
			}
		}
	}

	/** The item's rows of a relationship, read as the user. Kept for the life of the pane. */
	function readRelationships(sheet, columns, parentId) {
		const key = 'rel|' + sheet.itemType.name + '|' + parentId;
		if (!pane.reads.has(key)) {
			const hasRelated = !!(sheet.relationshipType && sheet.relationshipType.relatedItemType);
			const select = Array.from(new Set(columns.map((column) => column.property).concat(hasRelated ? [RELATED] : []).concat(['sort_order'])))
				.filter(isPropertyName).join(',');
			const aml = '<Item type="' + escapeXml(sheet.itemType.name) + '" action="get" select="' + select + '" orderBy="sort_order">' +
				'<source_id>' + escapeXml(parentId) + '</source_id></Item>';
			keep(key, query(top, aml).then((nodes) => nodes.map((node) => existingRow(columns, toItem(node)))));
		}
		return pane.reads.get(key);
	}

	function drawRelationships(sheet, columns, item, merged, focusRow) {
		listPanel.textContent = '';
		const relationshipType = sheet.relationshipType || {};
		const relatedType = relationshipType.relatedItemType || null;
		const counts = { new: 0, changed: 0, deleted: 0, unchanged: 0 };
		for (const row of merged.rows) {
			counts[row.status]++;
		}
		const tally = [counts.new + ' new', counts.changed + ' changed', counts.unchanged + ' unchanged'].join(', ') + '.' +
			(counts.deleted > 0 ? ' ' + counts.deleted + ' deleted.' : '');
		const summary = merged.rows.length === 0
			? 'No rows' + (item.action === 'new' ? ': the item is new, and no row of this sheet adds one.' : imported ? ' of this sheet.' : '.')
			: imported
				? (merged.rows.length === 1 ? '1 row' : merged.rows.length + ' rows') + ' of this sheet: ' + tally
				: (merged.rows.length - counts.deleted) + (merged.rows.length - counts.deleted === 1 ? ' row' : ' rows') + ' after the import: ' + tally;
		const head = make('div', { class: 'mce-pv-pad' }, make('p', null, summary));
		if (imported) {
			head.append(make('p', { class: 'mce-pv-muted' }, 'Only the rows of the workbook are listed, as the preview showed them. Open the item to see all its rows as they are now.'));
		} else if (item.action === 'new') {
			head.append(make('p', { class: 'mce-pv-muted' }, 'A new item has no rows yet: only the rows this import adds are listed.'));
		}
		if (merged.missing > 0) {
			head.append(make('p', { class: 'mce-pv-muted' }, merged.missing === 1
				? 'One row the upload matched was not found in Aras now: it may have been deleted since the check.'
				: merged.missing + ' rows the upload matched were not found in Aras now: they may have been deleted since the check.'));
		}
		if (merged.skipped.length > 0) {
			const list = make('ul', { class: 'mce-pv-errors' });
			for (const skipped of merged.skipped) {
				list.append(make('li', null, 'Row ' + skipped.row + ' is not imported: ' + (skipped.messages.join(' ') || 'it has an error.')));
			}
			head.append(list);
		}
		listPanel.append(head);
		if (merged.rows.length === 0) {
			return;
		}

		const th = (text, attributes) => make('th', { class: 'aras-grid-head-cell', scope: 'col', ...attributes }, text);
		const headRow = make('tr', null, th('Status'), th('Row', { title: 'The row number on the sheet' }));
		if (relatedType) {
			headRow.append(th(relatedType.label || relatedType.name));
		}
		for (const column of columns) {
			headRow.append(th(column.label || column.property));
		}
		const tbody = make('tbody');
		let focused = null;
		for (const row of merged.rows.slice(0, SHOWN_ROWS)) {
			const tr = make('tr', { class: 'mce-pv-row-' + row.status + (row.sheetRow !== null && row.sheetRow === focusRow ? ' mce-pv-row-focus' : '') },
				make('td', { class: 'aras-grid-row-cell' }, make('span', { class: 'mce-pv-badge mce-pv-badge-' + row.status }, STATUS_TEXT[row.status] + (row.missing ? ' (not found)' : ''))),
				make('td', { class: 'aras-grid-row-cell' }, make('span', { class: 'mce-pv-muted' }, row.sheetRow === null ? '' : String(row.sheetRow))));
			if (relatedType) {
				const changedRelated = row.changed.includes(RELATED);
				tr.append(cell(row.related ? row.related.keyedName || row.related.id : '', changedRelated,
					changedRelated ? (row.oldDisplay[RELATED] || row.old[RELATED] || '') : ''));
			}
			for (const column of columns) {
				const changed = row.changed.includes(column.id);
				const text = formatValue(sheet, column, row.values[column.id], row.display[column.id]);
				const was = changed ? formatValue(sheet, column, row.old[column.id], row.oldDisplay[column.id]) : '';
				tr.append(cell(text, changed, was));
			}
			if (row.warnings.length > 0) {
				tr.title = row.warnings.join(' ');
			}
			if (row.sheetRow !== null && row.sheetRow === focusRow) {
				focused = tr;
			}
			tbody.append(tr);
		}
		listPanel.append(make('table', { class: 'mce-pv-grid' }, make('thead', null, headRow), tbody));
		if (merged.rows.length > SHOWN_ROWS) {
			listPanel.append(make('p', { class: 'mce-pv-muted mce-pv-pad' }, 'The first ' + SHOWN_ROWS + ' rows are shown.'));
		}
		if (focused) {
			focused.scrollIntoView({ block: 'nearest' });
		}
	}

	/** A cell of a relationship row. A cell shows one line, so its title holds the whole value. */
	function cell(text, changed, was) {
		const td = make('td', { class: 'aras-grid-row-cell' + (changed ? ' mce-pv-changed' : '') });
		if (changed) {
			const before = was === '' ? '(empty)' : was;
			td.append(make('span', { class: 'mce-pv-was' }, before), ' ');
			td.title = (imported ? 'Before the import: ' : 'In Aras now: ') + before + '\nAfter the import: ' + (text === '' ? '(empty)' : text);
		} else if (text !== '') {
			td.title = text;
		}
		td.append(make('span', { class: 'mce-pv-value' }, text));
		return td;
	}

	/** A value as the user reads it: the label of a list value, a date in the user's format, Yes or No. */
	function formatValue(sheet, column, value, display) {
		if (display !== undefined && display !== null && display !== '') {
			return String(display);
		}
		const text = value === undefined || value === null ? '' : String(value);
		if (text === '') {
			return '';
		}
		try {
			switch (column.dataType) {
				case 'boolean':
					return text === '1' ? 'Yes' : text === '0' ? 'No' : text;
				case 'date':
					return top.aras.convertFromNeutral(text, 'date', top.aras.getDotNetDatePattern(column.pattern || '')) || text;
				case 'list':
					return listLabel(sheet, column, text);
				default:
					return text;
			}
		} catch (error) {
			return text;
		}
	}

	function listLabel(sheet, column, value) {
		const definition = typeOf(sheet.itemType.name).byName.get(column.property);
		if (!definition || !definition.dataSource) {
			return value;
		}
		if (!pane.lists.has(definition.dataSource)) {
			const labels = new Map();
			try {
				for (const node of Array.from(top.aras.getListValues(definition.dataSource) || [])) {
					labels.set(top.aras.getItemProperty(node, 'value'), top.aras.getItemProperty(node, 'label'));
				}
			} catch (error) {
				// Without the list, the stored value is shown.
			}
			pane.lists.set(definition.dataSource, labels);
		}
		return pane.lists.get(definition.dataSource).get(value) || value;
	}

	return { show, clear, dispose };
}

// ---------------------------------------------------------------------------
// Helpers of the page
// ---------------------------------------------------------------------------

/** A sheet as the workbook names it; a relationship sheet is named after its RelationshipType. */
const sheetLabel = (sheet) =>
	sheet.sheetName || (sheet.relationshipType && (sheet.relationshipType.name || sheet.relationshipType.label)) ||
	(sheet.itemType && (sheet.itemType.label || sheet.itemType.name)) || 'Sheet';

/** Creates an element. A null child is skipped. */
function element(doc, tag, attributes, ...children) {
	const node = doc.createElement(tag);
	for (const [name, value] of Object.entries(attributes || {})) {
		if (value !== false && value !== null && value !== undefined) {
			node.setAttribute(name, String(value));
		}
	}
	for (const child of children) {
		if (child !== null && child !== undefined && child !== false) {
			node.append(child);
		}
	}
	return node;
}

/** The chevron of the button that folds the relationships away: it points up while they are open. */
function chevron(doc) {
	const ns = 'http://www.w3.org/2000/svg';
	const svg = doc.createElementNS(ns, 'svg');
	svg.setAttribute('class', 'mce-pv-chevron');
	svg.setAttribute('viewBox', '0 0 24 24');
	svg.setAttribute('aria-hidden', 'true');
	svg.setAttribute('focusable', 'false');
	const path = doc.createElementNS(ns, 'path');
	path.setAttribute('d', 'M6 15l6-6 6 6');
	svg.append(path);
	return svg;
}

/** Runs a query as the user and returns the Item elements of the answer. Finding nothing gives an empty list. */
async function query(top, aml) {
	let answer;
	try {
		answer = await top.ArasModules.soap('<AML>' + aml + '</AML>', { async: true, method: 'ApplyAML' });
	} catch (failure) {
		const text = failure && typeof failure.responseText === 'string' ? failure.responseText : '';
		const fault = text ? new DOMParser().parseFromString(text, 'text/xml') : null;
		const pick = (name) => {
			const node = fault ? fault.getElementsByTagName(name)[0] : null;
			return node ? node.textContent.trim() : '';
		};
		// The server reports "no items found" as a fault with code 0.
		if (pick('faultcode') === '0') {
			return [];
		}
		throw new Error(pick('faultstring') || (failure && failure.message) || 'The server could not be reached.');
	}
	return Array.from(answer.children).filter((node) => node.localName === 'Item');
}

/** An item of an AML answer as a plain object: { id, props: { name: { value, keyedName } } }. */
function toItem(node) {
	const item = { id: node.getAttribute('id') || '', props: {} };
	for (const child of Array.from(node.children)) {
		const nested = Array.from(child.children).find((candidate) => candidate.localName === 'Item') || null;
		item.props[child.localName] = {
			value: nested ? nested.getAttribute('id') || '' : child.textContent,
			keyedName: child.getAttribute('keyed_name') || (nested ? nested.getAttribute('keyed_name') || '' : '')
		};
	}
	return item;
}
