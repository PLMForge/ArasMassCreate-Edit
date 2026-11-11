/**
 * Tests for the logic of the item preview (mce-preview.js): which item a row of
 * the validation result shows, which relationship rows belong to it, and how
 * the existing rows and the rows of the upload merge into the final state that
 * a relationship tab lists. Run from the repository root with `node --test`.
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const {
	changedLabels, changedProperties, changedRelationshipSheets, defaultProperties, existingCount, existingRow, formProperties, mergeRelationshipRows, newItemName, newItemXml,
	previewTarget, relationshipColumns, relationshipSheets, relationshipSheetsOf, rowsForParent, tabSummary
} = await import(new URL('../../1.0.0/Innovator/Client/customer/masscreateedit/mce-preview.js', import.meta.url).href);

const fixture = () => JSON.parse(readFileSync(new URL('./fixtures/validate-import.result.json', import.meta.url), 'utf8'));

const TA = '10C9E2CDE0416CA92132DF8805841B00';
const BOM_ROW_1 = 'CC039BC9DB5661565B0E6306B413B9DB';
const BOM_ROW_2 = '87E8CACE93AE3AD01D706A888E9328D3';
const BOM_ROW_3 = '0D2C3A7E0F1B4C8F9E6A5B4C3D2E1F00';
const TP_001 = 'D4E06766E751A61F6BCDEC13D149FD15';
const TP_002 = 'E7D599F81E3D0DC277C10ED107C173EE';
const TP_003 = '5B1A0E2D3C4B5A69788796A5B4C3D2E1';

/** The rows of TA-001's BOM in Aras, as the pane reads them: the three the fixture counts. */
function existingBom(result) {
	const columns = relationshipColumns(result.sheets[1]);
	const read = [
		{ id: BOM_ROW_1, props: { related_id: { value: TP_001, keyedName: 'TP-001' }, quantity: { value: '1' }, sort_order: { value: '1' } } },
		{ id: BOM_ROW_2, props: { related_id: { value: TP_002, keyedName: 'TP-002' }, quantity: { value: '3' }, sort_order: { value: '2' } } },
		{ id: BOM_ROW_3, props: { related_id: { value: TP_003, keyedName: 'TP-003' }, quantity: { value: '7' }, reference_designator: { value: 'C9' }, sort_order: { value: '3' } } }
	];
	return read.map((item) => existingRow(columns, item));
}

// ---------------------------------------------------------------------------
// Which item a row shows
// ---------------------------------------------------------------------------

test('a main-sheet row shows its own item', () => {
	const result = fixture();
	const update = previewTarget(result, 0, 1);
	assert.equal(update.kind, 'item');
	assert.equal(update.item.action, 'update');
	assert.equal(update.item.id, 'D4E06766E751A61F6BCDEC13D149FD15');
	assert.equal(update.item.mainRow, 3);
	assert.equal(update.tabSheetIndex, null);
	assert.deepEqual(update.warnings, [result.sheets[0].rows[1].issues[0].message]);

	const created = previewTarget(result, 0, 2);
	assert.equal(created.item.action, 'new');
	assert.equal(created.item.id, null);
	assert.equal(created.item.values.item_number, 'P-9001');

	const unchanged = previewTarget(result, 0, 0);
	assert.equal(unchanged.item.action, 'none');
	assert.equal(unchanged.item.keyedName, 'TA-001');
});

test('a row with errors has no preview, only its messages', () => {
	const result = fixture();
	const target = previewTarget(result, 0, 3);
	assert.equal(target.kind, 'error');
	assert.equal(target.messages.length, 2);
	assert.match(target.messages[0], /70 characters/);

	const relationship = previewTarget(result, 1, 5);
	assert.equal(relationship.kind, 'error');
	assert.match(relationship.messages[0], /parent/);
});

test('a relationship row with errors shows its parent with the messages, and is not marked on the tab', () => {
	const result = fixture();
	const row = result.sheets[1].rows[1];
	row.action = 'error';
	row.issues = [{ severity: 'error', code: 'TOO_LONG', message: 'Reference Designator: too long.' }, { severity: 'warning', code: 'W', message: 'A warning.' }];
	const target = previewTarget(result, 1, 1);
	assert.equal(target.kind, 'item');
	assert.equal(target.item.mainRow, 2);
	assert.equal(target.item.id, TA);
	assert.equal(target.tabSheetIndex, 1);
	assert.equal(target.focusRow, null);
	assert.deepEqual(target.rowErrors, ['Reference Designator: too long.']);
	assert.deepEqual(target.warnings, ['A warning.']);
	// On the tab, the row's change is not shown: the existing row it matched stays as it is in Aras.
	const merged = mergeRelationshipRows(result.sheets[1], rowsForParent(result.sheets[1], target.item), existingBom(result));
	assert.deepEqual(merged.rows.filter((shown) => shown.sheetRow === row.row).map((shown) => [shown.status, shown.changed]), [['unchanged', []]]);
	assert.ok(merged.skipped.some((skipped) => skipped.row === row.row));

	// A parent known only by its id is shown the same way.
	row.parent = { row: null, id: 'AAAABBBBCCCCDDDDEEEEFFFF00001111', keyedName: 'TA-777' };
	const byId = previewTarget(result, 1, 1);
	assert.deepEqual([byId.kind, byId.item.id, byId.focusRow], ['item', 'AAAABBBBCCCCDDDDEEEEFFFF00001111', null]);

	// Without a parent to show, only the row's own messages are shown.
	row.parent = null;
	const alone = previewTarget(result, 1, 1);
	assert.equal(alone.kind, 'error');
	assert.deepEqual(alone.messages, ['Reference Designator: too long.']);

	// A main-sheet row with errors still has no preview, and other rows say nothing of row errors.
	assert.equal(previewTarget(result, 0, 3).kind, 'error');
	assert.deepEqual(previewTarget(result, 0, 1).rowErrors, []);
	assert.deepEqual(previewTarget(result, 1, 2).rowErrors, []);
});

test('a row that is not there is missing', () => {
	const result = fixture();
	assert.equal(previewTarget(result, 0, 99).kind, 'missing');
	assert.equal(previewTarget(result, 7, 0).kind, 'missing');
	assert.equal(previewTarget(null, 0, 0).kind, 'missing');
});

test('a relationship row shows its parent, with its own tab open and the row marked', () => {
	const result = fixture();
	const target = previewTarget(result, 1, 1);
	assert.equal(target.kind, 'item');
	assert.equal(target.item.mainRow, 2);
	assert.equal(target.item.id, TA);
	assert.equal(target.item.action, 'none');
	assert.equal(target.tabSheetIndex, 1);
	assert.equal(target.focusRow, 3);

	// The parent of row 6 is the new main row 4.
	const underNew = previewTarget(result, 1, 4);
	assert.equal(underNew.item.action, 'new');
	assert.equal(underNew.item.mainRow, 4);
	assert.equal(underNew.focusRow, 6);
});

test('a relationship row whose parent is not on the main sheet shows the existing parent as it is', () => {
	const result = fixture();
	result.sheets[1].rows[2].parent = { row: null, id: 'AAAABBBBCCCCDDDDEEEEFFFF00001111', keyedName: 'TA-777' };
	const target = previewTarget(result, 1, 2);
	assert.equal(target.kind, 'item');
	assert.deepEqual([target.item.action, target.item.id, target.item.keyedName, target.item.mainRow], ['none', 'AAAABBBBCCCCDDDDEEEEFFFF00001111', 'TA-777', null]);
	assert.deepEqual(rowsForParent(result.sheets[1], target.item).map((row) => row.row), [4]);
});

test('the relationship sheets of a result are listed with their index', () => {
	const result = fixture();
	assert.deepEqual(relationshipSheets(result).map((entry) => [entry.index, entry.sheet.sheetName]), [[1, 'BOM']]);
	assert.deepEqual(relationshipSheets({ sheets: [result.sheets[0]] }), []);
});

// Two targets: Part (sheet 0) and CAD Document (sheet 1), Part CAD (sheet 2) under the Parts and
// CAD Structure (sheet 3) under the CAD Documents. Row 2 is on both main sheets.
function twoTargets() {
	return {
		sheets: [
			{ sheetName: 'Part', kind: 'main', itemType: { name: 'Part' }, rows: [{ row: 2, action: 'new', values: { item_number: 'P1' } }] },
			{ sheetName: 'CAD Document', kind: 'main', itemType: { name: 'CAD' }, rows: [
				{ row: 2, action: 'new', values: { item_number: 'C1' } },
				{ row: 3, action: 'update', id: 'E'.repeat(32), keyedName: 'C2', values: {} }
			] },
			{ sheetName: 'Part CAD', kind: 'relationship', parentSheet: 0, relatedSheet: 1, rows: [
				{ row: 2, action: 'new', parent: { sheet: 0, row: 2, id: null }, related: { sheet: 1, row: 2, id: null, keyedName: 'C1' } }
			], parents: [] },
			{ sheetName: 'CAD Structure', kind: 'relationship', parentSheet: 1, relatedSheet: 1, rows: [
				{ row: 2, action: 'new', parent: { sheet: 1, row: 2, id: null }, related: { sheet: 1, row: 3, id: null, keyedName: 'C2' } },
				{ row: 3, action: 'new', parent: { sheet: 1, row: 3, id: 'E'.repeat(32) }, related: null }
			], parents: [{ row: 3, id: 'E'.repeat(32), existing: 4 }] }
		]
	};
}

test('with two main sheets, an item knows its main sheet and has only the tabs and rows of its own', () => {
	const result = twoTargets();
	const part = previewTarget(result, 0, 0).item;
	const cad = previewTarget(result, 1, 0).item;
	assert.deepEqual([part.mainSheet, part.mainRow, cad.mainSheet, cad.mainRow], [0, 2, 1, 2]);
	assert.deepEqual(relationshipSheetsOf(result, 0).map((entry) => entry.index), [2]);
	assert.deepEqual(relationshipSheetsOf(result, 1).map((entry) => entry.index), [3]);
	// Row 2 of the Part sheet is not the parent of the CAD Structure row under CAD row 2.
	assert.deepEqual(rowsForParent(result.sheets[3], part), []);
	assert.deepEqual(rowsForParent(result.sheets[3], cad).map((row) => row.row), [2]);
	assert.deepEqual(rowsForParent(result.sheets[2], part).map((row) => row.row), [2]);
	assert.deepEqual(rowsForParent(result.sheets[2], cad), []);
	const existing = previewTarget(result, 1, 1).item;
	assert.equal(existingCount(result.sheets[3], existing), 4);
	assert.equal(existingCount(result.sheets[2], existing), 0, 'a sheet of the other target');

	// A relationship row shows the parent from its own main sheet.
	const structure = previewTarget(result, 3, 0);
	assert.deepEqual([structure.kind, structure.item.mainSheet, structure.item.values.item_number, structure.tabSheetIndex], ['item', 1, 'C1', 3]);
	const partCad = previewTarget(result, 2, 0);
	assert.deepEqual([partCad.item.mainSheet, partCad.item.values.item_number], [0, 'P1']);
	// A sheet that names its main sheet, with a row that does not.
	delete result.sheets[3].rows[0].parent.sheet;
	assert.equal(previewTarget(result, 3, 0).item.mainSheet, 1);
	// A parent known by its id belongs to the main sheet of the relationship sheet.
	result.sheets[3].rows[0].parent = { row: null, id: 'F'.repeat(32) };
	assert.equal(previewTarget(result, 3, 0).item.mainSheet, 1);
});

test('a result of one main sheet has every tab on it', () => {
	const result = fixture();
	const item = previewTarget(result, 0, 0).item;
	assert.equal(item.mainSheet, 0);
	assert.deepEqual(relationshipSheetsOf(result, item.mainSheet).map((entry) => entry.index), [1]);
	assert.equal(previewTarget(result, 1, 1).item.mainSheet, 0);
});

// ---------------------------------------------------------------------------
// The rows of a relationship tab
// ---------------------------------------------------------------------------

test('the rows of a sheet are found by the parent row, or by the parent id when the parent is not on the main sheet', () => {
	const result = fixture();
	const bom = result.sheets[1];
	const ta = previewTarget(result, 0, 0).item;
	assert.deepEqual(rowsForParent(bom, ta).map((row) => row.row), [2, 3, 4, 5, 8]);
	const created = previewTarget(result, 0, 2).item;
	assert.deepEqual(rowsForParent(bom, created).map((row) => row.row), [6]);
	assert.deepEqual(rowsForParent(bom, { action: 'none', id: 'X', mainRow: null }), []);
});

test('how many rows exist is known for a new item, and for a parent the result lists', () => {
	const result = fixture();
	const bom = result.sheets[1];
	assert.equal(existingCount(bom, previewTarget(result, 0, 0).item), 3);
	assert.equal(existingCount(bom, previewTarget(result, 0, 2).item), 0);
	// TP-001 (main row 3) is named by no BOM row: its rows have to be read.
	assert.equal(existingCount(bom, previewTarget(result, 0, 1).item), null);
});

test('a tab lists the relationship properties, not the parent and related key columns', () => {
	const result = fixture();
	assert.deepEqual(relationshipColumns(result.sheets[1]).map((column) => column.id), ['quantity', 'reference_designator', 'sort_order']);
});

test('an existing row is read into values by column id, with keyed names for item properties', () => {
	const columns = [
		{ id: 'quantity', property: 'quantity', dataType: 'float' },
		{ id: 'maker', property: 'maker_id', dataType: 'item' }
	];
	const row = existingRow(columns, { id: 'R1', props: { related_id: { value: 'P1', keyedName: 'P-1' }, quantity: { value: '2' }, maker_id: { value: 'M1', keyedName: 'Acme' } } });
	assert.deepEqual(row, { id: 'R1', related: { id: 'P1', keyedName: 'P-1' }, values: { quantity: '2', maker: 'M1' }, display: { maker: 'Acme' } });
	const bare = existingRow(columns, { id: 'R2', props: {} });
	assert.deepEqual(bare, { id: 'R2', related: null, values: { quantity: '', maker: '' }, display: {} });
});

test('the final state lists untouched, changed and new rows in Sequence order', () => {
	const result = fixture();
	const bom = result.sheets[1];
	const ta = previewTarget(result, 0, 0).item;
	const merged = mergeRelationshipRows(bom, rowsForParent(bom, ta), existingBom(result));

	assert.deepEqual(merged.rows.map((row) => [row.status, row.sheetRow, row.related && row.related.keyedName, row.values.sort_order]), [
		['unchanged', 2, 'TP-001', '1'],
		['changed', 3, 'TP-002', '2'],
		['unchanged', null, 'TP-003', '3'],
		['new', 4, 'TP-004', '4'],
		['new', 5, 'TP-001', '5']
	]);
	const changed = merged.rows[1];
	assert.equal(changed.id, BOM_ROW_2);
	assert.deepEqual(changed.values, { quantity: '5', reference_designator: 'R1', sort_order: '2' });
	assert.deepEqual(changed.changed, ['quantity', 'reference_designator']);
	assert.deepEqual(changed.old, { quantity: '3', reference_designator: '' });

	// The untouched row keeps every value it has in Aras.
	assert.deepEqual(merged.rows[2].values, { quantity: '7', reference_designator: 'C9', sort_order: '3' });
	assert.deepEqual(merged.rows[2].changed, []);
	// A new row has its values, and none of them counts as a change.
	assert.deepEqual(merged.rows[3].values, { quantity: '2', sort_order: '4' });
	assert.deepEqual(merged.rows[3].changed, []);
	assert.equal(merged.rows[3].id, null);
	// The repeated child keeps its warning.
	assert.equal(merged.rows[4].warnings.length, 1);

	// The error row is not imported, and says why.
	assert.deepEqual(merged.skipped.map((row) => row.row), [8]);
	assert.equal(merged.skipped[0].messages.length, 2);
	assert.equal(merged.missing, 0);
});

test('every existing row is listed even when the import only adds one', () => {
	const result = fixture();
	const bom = result.sheets[1];
	const ta = previewTarget(result, 0, 0).item;
	const onlyNew = rowsForParent(bom, ta).filter((row) => row.row === 4);
	const merged = mergeRelationshipRows(bom, onlyNew, existingBom(result));
	assert.deepEqual(merged.rows.map((row) => row.status), ['unchanged', 'unchanged', 'unchanged', 'new']);
});

test('a new parent shows only the new rows', () => {
	const result = fixture();
	const bom = result.sheets[1];
	const created = previewTarget(result, 0, 2).item;
	const merged = mergeRelationshipRows(bom, rowsForParent(bom, created), []);
	assert.deepEqual(merged.rows.map((row) => [row.status, row.sheetRow, row.related.keyedName]), [['new', 6, 'TP-001']]);
});

test('a matched row that is no longer in Aras is still listed, and counted as missing', () => {
	const result = fixture();
	const bom = result.sheets[1];
	const ta = previewTarget(result, 0, 0).item;
	const existing = existingBom(result).filter((row) => row.id !== BOM_ROW_2);
	const merged = mergeRelationshipRows(bom, rowsForParent(bom, ta), existing);
	assert.equal(merged.missing, 1);
	const row = merged.rows.find((candidate) => candidate.sheetRow === 3);
	assert.equal(row.status, 'changed');
	assert.equal(row.missing, true);
	assert.equal(row.values.quantity, '5');
});

test('after the import, only the rows of the workbook are listed, and none counts as missing', () => {
	const result = fixture();
	const bom = result.sheets[1];
	const ta = previewTarget(result, 0, 0).item;
	const merged = mergeRelationshipRows(bom, rowsForParent(bom, ta), [], true);
	assert.equal(merged.missing, 0);
	assert.deepEqual(merged.rows.map((row) => [row.status, row.sheetRow, row.missing]), [
		['unchanged', 2, false], ['changed', 3, false], ['new', 4, false], ['new', 5, false]
	]);
	const changed = merged.rows[1];
	assert.deepEqual(changed.changed, ['quantity', 'reference_designator']);
	assert.deepEqual(changed.old, { quantity: '3', reference_designator: '' });
	assert.deepEqual(merged.skipped.map((row) => row.row), [8]);
});

test('a row that gets another related item shows the new one and marks the change', () => {
	const sheet = {
		kind: 'relationship',
		columns: [{ id: 'parent.code', property: 'code' }, { id: 'related.code', property: 'code' }, { id: 'note', property: 'note', dataType: 'string' }],
		rows: [{
			row: 2, action: 'update', id: 'R1', values: { related_id: 'NEW', note: 'moved' }, old: { related_id: 'OLD', note: '' },
			oldDisplay: { related_id: 'Old item' }, related: { id: 'NEW', keyedName: 'New item' }, parent: { row: 2, id: 'P', keyedName: 'P' }
		}]
	};
	const existing = [{ id: 'R1', related: { id: 'OLD', keyedName: 'Old item' }, values: { note: '' }, display: {} }];
	const [row] = mergeRelationshipRows(sheet, sheet.rows, existing).rows;
	assert.equal(row.status, 'changed');
	assert.deepEqual(row.related, { id: 'NEW', keyedName: 'New item' });
	assert.deepEqual(row.changed, ['related_id', 'note']);
	assert.equal(row.oldDisplay.related_id, 'Old item');
	assert.equal(row.values.related_id, undefined);
});

test('an item value that an update empties loses its keyed name', () => {
	const sheet = { kind: 'relationship', columns: [{ id: 'maker', property: 'maker_id', dataType: 'item' }], rows: [] };
	const rows = [{ row: 3, action: 'update', id: 'R1', values: { maker: '' }, old: { maker: 'M1' }, oldDisplay: { maker: 'Acme' }, parent: { row: 2 } }];
	const existing = [{ id: 'R1', related: null, values: { maker: 'M1' }, display: { maker: 'Acme' } }];
	const [row] = mergeRelationshipRows(sheet, rows, existing).rows;
	assert.equal(row.values.maker, '');
	assert.equal(row.display.maker, undefined);
	assert.equal(row.oldDisplay.maker, 'Acme');
});

test('without a Sequence column, existing rows come first in the order read, then the new rows', () => {
	const sheet = { kind: 'relationship', columns: [{ id: 'note', property: 'note' }], rows: [] };
	const rows = [{ row: 2, action: 'new', values: { note: 'n' }, related: null, parent: { row: 2 } }];
	const existing = [
		{ id: 'B', related: null, values: { note: 'b' }, display: {} },
		{ id: 'A', related: null, values: { note: 'a' }, display: {} }
	];
	assert.deepEqual(mergeRelationshipRows(sheet, rows, existing).rows.map((row) => row.values.note), ['b', 'a', 'n']);
});

test('rows without a Sequence value go last, and equal Sequences keep their order', () => {
	const sheet = { kind: 'relationship', columns: [{ id: 'sort_order', property: 'sort_order' }], rows: [] };
	const existing = [
		{ id: 'A', related: null, values: { sort_order: '' }, display: {} },
		{ id: 'B', related: null, values: { sort_order: '20' }, display: {} },
		{ id: 'C', related: null, values: { sort_order: '10' }, display: {} },
		{ id: 'D', related: null, values: { sort_order: '10' }, display: {} }
	];
	assert.deepEqual(mergeRelationshipRows(sheet, [], existing).rows.map((row) => row.id), ['C', 'D', 'B', 'A']);
});

test('an existing row matched by an error row stays unchanged, and the error is reported', () => {
	const sheet = { kind: 'relationship', columns: [{ id: 'note', property: 'note' }], rows: [] };
	const rows = [{ row: 4, action: 'error', id: 'A', issues: [{ severity: 'error', code: 'TOO_LONG', message: 'Note: too long.' }], parent: { row: 2 } }];
	const existing = [{ id: 'A', related: null, values: { note: 'a' }, display: {} }];
	const merged = mergeRelationshipRows(sheet, rows, existing);
	assert.deepEqual(merged.rows.map((row) => [row.status, row.sheetRow, row.values.note]), [['unchanged', 4, 'a']]);
	assert.deepEqual(merged.skipped, [{ row: 4, messages: ['Note: too long.'] }]);
});

test('the label of a tab counts what the import does to the item rows', () => {
	const result = fixture();
	const bom = result.sheets[1];
	assert.deepEqual(tabSummary(rowsForParent(bom, previewTarget(result, 0, 0).item)), { new: 2, changed: 1, deleted: 0, errors: 1 });
	assert.deepEqual(tabSummary([]), { new: 0, changed: 0, deleted: 0, errors: 0 });
});

test('only the relationships the import changes for the item have a tab, and the sheet of the row chosen', () => {
	const result = fixture();
	const tabsOf = (target) => changedRelationshipSheets(result, target.item, target.tabSheetIndex).map((entry) => [entry.index, entry.summary]);
	// TA-001 gets new and changed BOM rows, and one with errors.
	assert.deepEqual(tabsOf(previewTarget(result, 0, 0)), [[1, { new: 2, changed: 1, deleted: 0, errors: 1 }]]);
	// The new item gets a new BOM row.
	assert.deepEqual(tabsOf(previewTarget(result, 0, 2)).map(([index]) => index), [1]);
	// No BOM row names TP-001: it has no tab, unless a row of the sheet is the one chosen.
	const tp = previewTarget(result, 0, 1).item;
	assert.deepEqual(changedRelationshipSheets(result, tp), []);
	assert.deepEqual(changedRelationshipSheets(result, tp, 1).map((entry) => [entry.index, entry.summary]), [[1, { new: 0, changed: 0, deleted: 0, errors: 0 }]]);

	// A row the import leaves alone keeps its tab when it is the row chosen, and also when it is in error.
	const quiet = { sheets: [
		{ sheetName: 'Part', kind: 'main', itemType: { name: 'Part' }, rows: [{ row: 2, action: 'none', id: 'A'.repeat(32), keyedName: 'P1' }] },
		{ sheetName: 'BOM', kind: 'relationship', parentSheet: 0, rows: [{ row: 2, action: 'none', id: 'R1', parent: { sheet: 0, row: 2, id: 'A'.repeat(32) } }] },
		{ sheetName: 'Part CAD', kind: 'relationship', parentSheet: 0, rows: [{ row: 2, action: 'delete', id: 'R2', parent: { sheet: 0, row: 2, id: 'A'.repeat(32) } }] }
	] };
	const fromMain = previewTarget(quiet, 0, 0);
	assert.deepEqual(changedRelationshipSheets(quiet, fromMain.item, fromMain.tabSheetIndex).map((entry) => entry.index), [2]);
	const fromBom = previewTarget(quiet, 1, 0);
	assert.deepEqual(changedRelationshipSheets(quiet, fromBom.item, fromBom.tabSheetIndex).map((entry) => entry.index), [1, 2]);
	quiet.sheets[1].rows[0].action = 'error';
	quiet.sheets[1].rows[0].issues = [{ severity: 'error', message: 'Quantity: not a number.' }];
	const failed = previewTarget(quiet, 1, 0);
	assert.deepEqual([failed.kind, failed.tabSheetIndex], ['item', 1]);
	assert.deepEqual(changedRelationshipSheets(quiet, failed.item, failed.tabSheetIndex).map((entry) => [entry.index, entry.summary.errors]), [[1, 1], [2, 0]]);

	// With two main sheets, an item has only tabs of its own target.
	const two = twoTargets();
	assert.deepEqual(changedRelationshipSheets(two, previewTarget(two, 0, 0).item).map((entry) => entry.index), [2]);
	assert.deepEqual(changedRelationshipSheets(two, previewTarget(two, 1, 1).item).map((entry) => entry.index), [3]);
});

// ---------------------------------------------------------------------------
// The item on the Form
// ---------------------------------------------------------------------------

test('the Form gets the row values by property, with keyed names for item properties', () => {
	const result = fixture();
	const main = result.sheets[0];
	const update = previewTarget(result, 0, 1).item;
	assert.deepEqual(formProperties(main, update.values, update.display), [
		{ name: 'owned_by_id', value: 'D307AF07B9B308341F0ECB1513548651', keyedName: 'Administrators', isItem: true },
		{ name: 'name', value: 'Test Part 1, revised', keyedName: null, isItem: false },
		{ name: 'make_buy', value: 'Buy', keyedName: null, isItem: false }
	]);
	// Columns the template no longer has are left out.
	assert.deepEqual(formProperties(main, { legacy_code: 'L-1', name: 'x' }, {}).map((property) => property.name), ['name']);
	// An emptied item property has no keyed name.
	assert.deepEqual(formProperties(main, { owned_by_id: '' }, {}), [{ name: 'owned_by_id', value: '', keyedName: '', isItem: true }]);
});

test('an update row names the properties it changes, in sheet order', () => {
	const result = fixture();
	assert.deepEqual(changedLabels(result.sheets[0], result.sheets[0].rows[1].values), ['Assigned Creator', 'Name', 'Make / Buy']);
});

test('the Form marks the properties an update row changes, with the value in Aras now', () => {
	const result = fixture();
	const main = result.sheets[0];
	const update = previewTarget(result, 0, 1).item;
	const changes = changedProperties(main, update);
	assert.deepEqual(changes.map((change) => change.property), ['owned_by_id', 'name', 'make_buy']);
	for (const change of changes) {
		const column = main.columns.find((candidate) => candidate.property === change.property);
		const expected = update.oldDisplay[column.id] !== undefined ? update.oldDisplay[column.id] : update.old[column.id];
		assert.equal(change.was, expected === undefined || expected === null ? '' : String(expected));
		assert.equal(change.label, column.label);
	}
	// A new item and a row that changes nothing have no marks.
	assert.deepEqual(changedProperties(main, previewTarget(result, 0, 2).item), []);
	assert.deepEqual(changedProperties(main, previewTarget(result, 0, 0).item), []);
	// A column the template no longer has, or a name that is no property name, is not marked.
	const odd = { action: 'update', values: { legacy_code: 'x', 'name"]': 'y' }, old: {} };
	assert.deepEqual(changedProperties({ columns: [{ id: 'legacy_code', property: 'legacy_code', status: 'unknown' }, { id: 'name"]', property: 'name"]' }] }, odd), []);
});

test('a new item is named by its key values', () => {
	const result = fixture();
	assert.equal(newItemName(result.sheets[0], result.sheets[0].rows[2].values), 'P-9001');
	assert.equal(newItemName({ keyColumns: ['a', 'b'] }, { a: 'X', b: 2 }), 'X / 2');
	assert.equal(newItemName({ keyColumns: [] }, { a: 'X' }), '');
});

test('the XML of a new item is a temporary item with escaped values', () => {
	const xml = newItemXml({ name: 'Part', id: 'T1' }, 'ID1', [
		{ name: 'name', value: 'A & B <c>', keyedName: null, isItem: false },
		{ name: 'owned_by_id', value: 'I1', keyedName: 'R&D "team"', isItem: true, type: 'Identity' },
		{ name: 'managed_by_id', value: '', keyedName: '', isItem: true },
		{ name: 'bad name', value: 'x', keyedName: null, isItem: false }
	]);
	assert.equal(xml,
		'<Item type="Part" typeId="T1" id="ID1" action="add" isTemp="1" isDirty="1">' +
		'<name>A &amp; B &lt;c&gt;</name>' +
		'<owned_by_id keyed_name="R&amp;D &quot;team&quot;" type="Identity">I1</owned_by_id>' +
		'<managed_by_id></managed_by_id>' +
		'</Item>');
});

test('defaults fill only the properties the row leaves empty, and never item properties', () => {
	const definitions = [
		{ name: 'unit', dataType: 'list', defaultValue: 'EA' },
		{ name: 'make_buy', dataType: 'list', defaultValue: 'Make' },
		{ name: 'owned_by_id', dataType: 'item', defaultValue: 'I1' },
		{ name: 'item_number', dataType: 'sequence', defaultValue: 'x' },
		{ name: 'description', dataType: 'string', defaultValue: '' },
		{ name: 'is_template', dataType: 'boolean', defaultValue: '0' }
	];
	assert.deepEqual(defaultProperties(definitions, new Set(['make_buy'])).map((property) => [property.name, property.value]), [['unit', 'EA'], ['is_template', '0']]);
	assert.deepEqual(defaultProperties(definitions, ['unit', 'make_buy', 'is_template']), []);
});
