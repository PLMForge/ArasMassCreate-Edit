/**
 * Tests for the export logic (mce-export.js): the export options of a job,
 * what the sheet picker offers, the requests that read the items to pre-fill,
 * and the rows made from their answers. Run from the repository root with
 * `node --test`.
 */

import test from 'node:test';
import assert from 'node:assert/strict';

import { buildWorkbook, parseWorkbook } from './helpers/workbook-tools.mjs';

const {
	IDS_PER_REQUEST, allowedForUser, chunk, columnChoices, configIdsAml, lockedReason, narrowSchema, defaultMatchKey, describeKey, editSelection, parseIdList, selectionNamesAml, mainItemsAml, mainRows, mainSelect, matchKeyConsequences, matchKeyFor,
	orderMainItems, parseOptions, parseTargets, relationshipRows, relationshipRowsAml, selectedGroups, serializeOptions, serializeTargets, sheetChoices, tidyOptions, workbookFileName
} = await import(new URL('../../1.0.0/Innovator/Client/customer/masscreateedit/mce-export.js', import.meta.url).href);

// ---------------------------------------------------------------------------
// A small schema, shaped like the answer of the schema Method
// ---------------------------------------------------------------------------

const PART = { id: '4F1AC04A2B484F3ABA4E20DB63808A88', name: 'Part', label: 'Part', isVersionable: true };
const MAKER = { id: 'FDE9DA6F52C642EA807287C340EE0B72', name: 'Manufacturer Part', label: 'Manufacturer Part', isVersionable: false };
const BOM = '159C6D88795B4A86864420863466F728';
const AML = '56595520EC8A4EAB9B6D80A617F5D6AA';
const GOAL = '024A379F7F5C466C8AE23042E2CB60AB';
const TREE = '567E4149FBF74DACA0B0C4C9B1E79A3B';

const column = (id, extra = {}) => ({
	id, role: 'property', property: id.replace(/^(parent|related)\./, ''), label: id, dataType: 'string', isKey: false, isMatchKey: false, ...extra
});

function schema() {
	return {
		sheets: [
			{
				kind: 'main', itemType: PART, relationshipType: null, createOnly: false,
				columns: [
					column('item_number', { label: 'Part Number', isKey: true }),
					column('name', { label: 'Name' }),
					column('cost', { label: 'Cost', dataType: 'decimal' }),
					column('owned_by_id', { label: 'Owner', dataType: 'item' })
				]
			},
			{
				kind: 'relationship', itemType: { id: 'A', name: 'Part BOM', label: 'Part BOM', isVersionable: false },
				relationshipType: { id: BOM, name: 'Part BOM', label: 'BOM', relatedItemType: PART }, createOnly: false,
				columns: [
					column('parent.item_number', { role: 'parentKey', label: 'Part Number', isKey: true, isMatchKey: true }),
					column('related.item_number', { role: 'relatedKey', label: 'Part Number', isKey: true, isMatchKey: true }),
					column('quantity', { label: 'Quantity', dataType: 'float' }),
					column('sort_order', { label: 'Sequence', dataType: 'integer', isMatchKey: true })
				]
			},
			{
				kind: 'relationship', itemType: { id: 'B', name: 'Part AML', label: 'Part AML', isVersionable: false },
				relationshipType: { id: AML, name: 'Part AML', label: 'AML', relatedItemType: MAKER }, createOnly: true,
				columns: [column('parent.item_number', { role: 'parentKey', label: 'Part Number', isKey: true, isMatchKey: true })]
			},
			{
				kind: 'relationship', itemType: { id: 'C', name: 'Part Goal', label: 'Part Goal', isVersionable: false },
				relationshipType: { id: GOAL, name: 'Part Goal', label: 'Goals', relatedItemType: null }, createOnly: false,
				columns: [
					column('parent.item_number', { role: 'parentKey', label: 'Part Number', isKey: true, isMatchKey: true }),
					column('goal', { label: 'Goal' }),
					column('sort_order', { label: 'Sequence', dataType: 'integer', isMatchKey: true })
				]
			}
		],
		relationships: [
			{ id: BOM, name: 'Part BOM', label: 'BOM', relatedItemType: PART, relatedRequired: true, canBeSheet: true, reason: null },
			{ id: TREE, name: 'Part MultiLevel BOM', label: 'BOM Structure', relatedItemType: null, relatedRequired: false, canBeSheet: false, reason: 'Shown by a page of its own.' },
			{ id: AML, name: 'Part AML', label: 'AML', relatedItemType: MAKER, relatedRequired: true, canBeSheet: true, reason: null },
			{ id: GOAL, name: 'Part Goal', label: 'Goals', relatedItemType: null, relatedRequired: false, canBeSheet: true, reason: null }
		]
	};
}

const choiceOf = (id) => sheetChoices(schema()).find((choice) => choice.relationshipTypeId === id);

/** An item as the Form hands it over: props maps a name to { value, keyedName, item, restricted }. */
const item = (id, props) => ({
	id,
	props: Object.fromEntries(Object.entries(props).map(([name, value]) => [
		name,
		value !== null && typeof value === 'object' ? { value: '', keyedName: '', item: null, restricted: false, ...value } : { value, keyedName: '', item: null, restricted: false }
	]))
});

const id32 = (digit) => String(digit).repeat(32);

// ---------------------------------------------------------------------------
// Export options
// ---------------------------------------------------------------------------

test('export options survive being written and read', () => {
	const options = { sheets: [{ relationshipTypeId: BOM, name: 'Part BOM', matchKey: ['parent.item_number', 'sort_order'] }] };
	assert.deepEqual(parseOptions(serializeOptions(options)), options);
});

test('options without sheets are written out, so that "no sheets" differs from "never chosen"', () => {
	assert.equal(serializeOptions({ sheets: [] }), '{"version":1,"sheets":[]}');
	assert.equal(serializeOptions(null), '{"version":1,"sheets":[]}');
});

test('options that are empty, damaged or of the wrong shape read as no sheets', () => {
	for (const text of [null, undefined, '', '   ', '{', '[]', '{"sheets":7}', 'null']) {
		assert.deepEqual(parseOptions(text), { sheets: [] });
	}
});

test('reading options drops sheets without a RelationshipType id, repeated sheets and empty match-key entries', () => {
	const read = parseOptions(JSON.stringify({
		sheets: [
			{ relationshipTypeId: BOM, name: 'Part BOM', matchKey: ['parent.item_number', '', 5, 'sort_order'] },
			{ relationshipTypeId: BOM, name: 'again', matchKey: [] },
			{ relationshipTypeId: 'not an id', matchKey: [] },
			null,
			{ relationshipTypeId: GOAL }
		]
	}));
	assert.deepEqual(read, {
		sheets: [
			{ relationshipTypeId: BOM, name: 'Part BOM', matchKey: ['parent.item_number', 'sort_order'] },
			{ relationshipTypeId: GOAL, name: '', matchKey: [] }
		]
	});
});

// ---------------------------------------------------------------------------
// What the Form shows
// ---------------------------------------------------------------------------

test('the key of the main sheet is described by its key columns', () => {
	assert.deepEqual(describeKey(schema()), { itemType: 'Part', createOnly: false, columns: [{ id: 'item_number', label: 'Part Number' }] });
	const keyless = schema();
	keyless.sheets[0].createOnly = true;
	keyless.sheets[0].columns.forEach((entry) => { entry.isKey = false; });
	assert.deepEqual(describeKey(keyless), { itemType: 'Part', createOnly: true, columns: [] });
});

test('the picker offers one choice per relationship, in the order of the schema', () => {
	const choices = sheetChoices(schema());
	assert.deepEqual(choices.map((choice) => choice.label), ['BOM', 'BOM Structure', 'AML', 'Goals']);
	assert.deepEqual(choices.map((choice) => choice.selectable), [true, false, false, true]);
});

test('a relationship the schema refuses is offered with the reason, and without a match key', () => {
	const choice = choiceOf(TREE);
	assert.equal(choice.reason, 'Shown by a page of its own.');
	assert.deepEqual(choice.groups, []);
});

test('a relationship that needs a related item which no row could name can\'t be picked', () => {
	const choice = choiceOf(AML);
	assert.equal(choice.selectable, false);
	assert.match(choice.reason, /needs a related Manufacturer Part/);
});

test('a create-only relationship whose related item is optional can be picked, with the parent as its whole key', () => {
	const optional = schema();
	optional.relationships[2].relatedRequired = false;
	const choice = sheetChoices(optional).find((candidate) => candidate.relationshipTypeId === AML);
	assert.equal(choice.selectable, true);
	assert.equal(choice.createOnly, true);
	assert.deepEqual(choice.groups.map((group) => group.key), ['parent']);
	assert.match(matchKeyConsequences(choice, ['parent'])[0], /^Create only/);
});

test('the default match key is the parent key, the related key and the sequence', () => {
	const choice = choiceOf(BOM);
	assert.deepEqual(choice.groups.map((group) => [group.key, group.fixed, group.byDefault]), [
		['parent', true, true], ['related', false, true], ['quantity', false, false], ['sort_order', false, true]
	]);
	assert.deepEqual(defaultMatchKey(choice), ['parent.item_number', 'related.item_number', 'sort_order']);
});

test('a relationship without a related ItemType has no related group', () => {
	const choice = choiceOf(GOAL);
	assert.deepEqual(choice.groups.map((group) => group.key), ['parent', 'goal', 'sort_order']);
	assert.deepEqual(defaultMatchKey(choice), ['parent.item_number', 'sort_order']);
});

test('the parent key can\'t be taken out of a match key', () => {
	const choice = choiceOf(BOM);
	assert.deepEqual(matchKeyFor(choice, []), ['parent.item_number']);
	assert.deepEqual(selectedGroups(choice, []), ['parent']);
	assert.deepEqual(selectedGroups(choice, ['related.item_number']), ['parent', 'related']);
});

test('a match key comes out in the order of the sheet\'s columns, whatever the order of ticking', () => {
	const choice = choiceOf(BOM);
	assert.deepEqual(matchKeyFor(choice, ['sort_order', 'quantity']), ['parent.item_number', 'quantity', 'sort_order']);
});

test('a key of several properties is ticked as a whole', () => {
	const twoKeys = schema();
	twoKeys.sheets[1].columns.splice(2, 0, column('related.major_rev', { role: 'relatedKey', label: 'Revision', isKey: true, isMatchKey: true }));
	const choice = sheetChoices(twoKeys)[0];
	assert.deepEqual(choice.groups[1].ids, ['related.item_number', 'related.major_rev']);
	assert.deepEqual(selectedGroups(choice, ['parent.item_number', 'related.item_number']), ['parent']);
	assert.deepEqual(matchKeyFor(choice, ['related']), ['parent.item_number', 'related.item_number', 'related.major_rev']);
});

test('tidying options drops what the ItemType does not offer and orders the sheets like the picker', () => {
	const choices = sheetChoices(schema());
	const tidy = tidyOptions({
		sheets: [
			{ relationshipTypeId: GOAL, name: 'old name', matchKey: ['parent.item_number', 'gone_property'] },
			{ relationshipTypeId: TREE, name: 'x', matchKey: [] },
			{ relationshipTypeId: id32(9), name: 'deleted relationship', matchKey: [] },
			{ relationshipTypeId: BOM, name: 'Part BOM', matchKey: ['quantity'] }
		]
	}, choices);
	assert.deepEqual(tidy, {
		sheets: [
			{ relationshipTypeId: BOM, name: 'Part BOM', matchKey: ['parent.item_number', 'quantity'] },
			{ relationshipTypeId: GOAL, name: 'Part Goal', matchKey: ['parent.item_number'] }
		]
	});
});

test('a sheet stored without a match key gets the default one', () => {
	const tidy = tidyOptions({ sheets: [{ relationshipTypeId: BOM, name: '', matchKey: [] }] }, sheetChoices(schema()));
	assert.deepEqual(tidy.sheets[0].matchKey, ['parent.item_number', 'related.item_number', 'sort_order']);
});

test('the consequences of a match key name what the user gives up', () => {
	const choice = choiceOf(BOM);
	const byDefault = matchKeyConsequences(choice, ['related', 'sort_order']);
	assert.equal(byDefault.length, 2);
	assert.match(byDefault[1], /Sequence.*adds a new row/);

	const withoutRelated = matchKeyConsequences(choice, ['sort_order']);
	assert.ok(withoutRelated.some((line) => /related item is not ticked/.test(line)));

	const parentOnly = matchKeyConsequences(choice, []);
	assert.ok(parentOnly.some((line) => /Only the parent is ticked/.test(line)));
	assert.ok(!parentOnly.some((line) => /renumber/.test(line)));
});

test('an optional related item in the match key is called out', () => {
	const optional = schema();
	optional.relationships[0].relatedRequired = false;
	const choice = sheetChoices(optional)[0];
	assert.ok(matchKeyConsequences(choice, ['related']).some((line) => /without a related item can't be matched/.test(line)));
	assert.ok(!matchKeyConsequences(choice, ['sort_order']).some((line) => /without a related item can't be matched/.test(line)));
});

// ---------------------------------------------------------------------------
// Requests
// ---------------------------------------------------------------------------

test('a list is cut into requests of at most 250 ids', () => {
	const ids = Array.from({ length: 501 }, (unused, index) => index);
	assert.equal(IDS_PER_REQUEST, 250);
	assert.deepEqual(chunk(ids).map((part) => part.length), [250, 250, 1]);
	assert.deepEqual(chunk([]), []);
});

test('the request for config_ids takes ids only', () => {
	assert.equal(
		configIdsAml('Part', [id32(1), 'x" action="delete', id32(2)]),
		'<Item type="Part" action="get" select="config_id" idlist="' + id32(1) + ',' + id32(2) + '"/>'
	);
});

test('the selected items: read from the job, removed and added, each once and in order', () => {
	const [a, b, c, d] = [id32(1), id32(2), id32(3), id32(4)];
	assert.deepEqual(parseIdList(' ' + a + ', ' + b + ',,' + a + ',not-an-id,' + c), [a, b, c]);
	assert.deepEqual(parseIdList(''), []);
	assert.deepEqual(parseIdList(null), []);
	assert.deepEqual(editSelection([a, b, c], { remove: [b] }), [a, c]);
	assert.deepEqual(editSelection([a, b], { add: [c, a, d, 'x'] }), [a, b, c, d]);
	assert.deepEqual(editSelection([a, b], { remove: [a, b] }), []);
	assert.deepEqual(editSelection([a], {}), [a]);
	assert.equal(selectionNamesAml('Part', [a, '" action="delete', b]),
		'<Item type="Part" action="get" select="keyed_name" idlist="' + a + ',' + b + '"/>');
});

test('a versionable ItemType is read by config_id in its current generation', () => {
	const main = schema().sheets[0];
	assert.equal(mainSelect(main), 'config_id,modified_on,item_number,name,cost,owned_by_id');
	assert.equal(
		mainItemsAml(main, [id32(1), id32(2)]),
		'<Item type="Part" action="get" select="config_id,modified_on,item_number,name,cost,owned_by_id">' +
		'<config_id condition="in">\'' + id32(1) + '\',\'' + id32(2) + '\'</config_id><is_current>1</is_current></Item>'
	);
});

test('an ItemType without versions is read by id', () => {
	const main = schema().sheets[0];
	main.itemType = { id: 'X', name: 'Customer & Co', label: 'Customer', isVersionable: false };
	assert.equal(
		mainItemsAml(main, [id32(1)]),
		'<Item type="Customer &amp; Co" action="get" select="config_id,modified_on,item_number,name,cost,owned_by_id" idlist="' + id32(1) + '"/>'
	);
});

test('relationship rows are asked for with the related item\'s key', () => {
	const sheets = schema().sheets;
	assert.equal(
		relationshipRowsAml(sheets[1], [id32(1), 'nonsense']),
		'<Item type="Part BOM" action="get" select="source_id,related_id(item_number),sort_order,modified_on,quantity">' +
		'<source_id condition="in">\'' + id32(1) + '\'</source_id></Item>'
	);
	assert.match(relationshipRowsAml(sheets[3], [id32(1)]), /select="source_id,related_id,sort_order,modified_on,goal"/);
});

// ---------------------------------------------------------------------------
// Rows
// ---------------------------------------------------------------------------

test('items come out in the order they were selected in, each once, in the generation that was read', () => {
	const configIdOf = new Map([[id32(1), id32('A')], [id32(2), id32('B')], [id32(3), id32('A')], [id32(4), id32('D')]]);
	const current = [
		item(id32(7), { config_id: id32('B'), item_number: 'P-2' }),
		item(id32(8), { config_id: id32('A'), item_number: 'P-1' })
	];
	// Selected: an old generation of A, then B, then another generation of A, one that is gone and one that was never found.
	const ordered = orderMainItems([id32(1), id32(2), id32(3), id32(4), id32(5)], configIdOf, current);
	assert.deepEqual(ordered.items.map((entry) => entry.props.item_number.value), ['P-1', 'P-2']);
	assert.equal(ordered.missing, 2);
});

test('without versions, items are matched by id', () => {
	const ordered = orderMainItems([id32(2), id32(1), id32(3)], null, [item(id32(1), {}), item(id32(2), {})]);
	assert.deepEqual(ordered.items.map((entry) => entry.id), [id32(2), id32(1)]);
	assert.equal(ordered.missing, 1);
});

test('a main row takes the keyed name of an item property and leaves empty values blank', () => {
	const main = schema().sheets[0];
	const rows = mainRows(main, [
		item(id32(1), {
			item_number: 'P-1', name: '', cost: '12.5000', modified_on: '2026-03-01T10:00:00',
			owned_by_id: { value: id32(5), keyedName: 'Engineering' }
		}),
		item(id32(2), { item_number: 'P-2', owned_by_id: { value: '', restricted: true } })
	]);
	assert.deepEqual(rows, [
		{ values: { item_number: 'P-1', name: '', cost: '12.5000', owned_by_id: 'Engineering' }, modifiedOn: '2026-03-01T10:00:00' },
		{ values: { item_number: 'P-2', name: '', cost: '', owned_by_id: '' }, modifiedOn: null }
	]);
});

test('relationship rows follow their parents, and the sequence within a parent', () => {
	const sheet = schema().sheets[1];
	const parents = [item(id32(1), { item_number: 'A-1' }), item(id32(2), { item_number: 'A-2' })];
	const related = (number) => ({ value: id32(9), keyedName: number, item: item(id32(9), { item_number: number }) });
	const read = relationshipRows(sheet, parents, [
		item(id32('c'), { source_id: id32(2), related_id: related('P-9'), sort_order: '10', quantity: '1' }),
		item(id32('b'), { source_id: id32(1), related_id: related('P-2'), sort_order: '20', quantity: '2', modified_on: '2026-03-01T10:00:00' }),
		item(id32('a'), { source_id: id32(1), related_id: related('P-3'), sort_order: '', quantity: '3' }),
		item(id32('d'), { source_id: id32(1), related_id: related('P-1'), sort_order: '3', quantity: '4' }),
		item(id32('e'), { source_id: id32(8), related_id: related('P-0'), sort_order: '1', quantity: '5' })
	]);
	assert.deepEqual(read.rows.map((row) => [row.values['parent.item_number'], row.values['related.item_number'], row.values.sort_order]), [
		['A-1', 'P-1', '3'], ['A-1', 'P-2', '20'], ['A-1', 'P-3', ''], ['A-2', 'P-9', '10']
	]);
	assert.equal(read.rows[1].modifiedOn, '2026-03-01T10:00:00');
	assert.equal(read.rows[0].modifiedOn, null);
	assert.deepEqual([read.unreadable, read.withoutRelated], [0, 0]);
});

test('a row whose related item can\'t be read is left out and counted', () => {
	const sheet = schema().sheets[1];
	const parents = [item(id32(1), { item_number: 'A-1' })];
	const read = relationshipRows(sheet, parents, [
		item(id32('a'), { source_id: id32(1), related_id: { value: '', restricted: true }, sort_order: '1' }),
		item(id32('b'), { source_id: id32(1), related_id: { value: id32(9), keyedName: 'P-9' }, sort_order: '2' })
	]);
	assert.equal(read.rows.length, 0);
	assert.equal(read.unreadable, 2);
});

test('a row without a related item is exported only when the related key is not in the match key', () => {
	const parents = [item(id32(1), { item_number: 'A-1' })];
	const rows = [item(id32('a'), { source_id: id32(1), related_id: '', sort_order: '1', quantity: '2' })];

	const inKey = relationshipRows(schema().sheets[1], parents, rows);
	assert.deepEqual([inKey.rows.length, inKey.withoutRelated], [0, 1]);

	const notInKey = schema().sheets[1];
	notInKey.columns[1].isMatchKey = false;
	const read = relationshipRows(notInKey, parents, rows);
	assert.deepEqual(read.rows.map((row) => row.values), [{ 'parent.item_number': 'A-1', 'related.item_number': '', quantity: '2', sort_order: '1' }]);
	assert.equal(read.withoutRelated, 0);
});

test('a relationship without a related ItemType exports all of its rows', () => {
	const sheet = schema().sheets[3];
	const read = relationshipRows(sheet, [item(id32(1), { item_number: 'A-1' })], [
		item(id32('a'), { source_id: id32(1), goal: 'Lighter', sort_order: '1' })
	]);
	assert.deepEqual(read.rows.map((row) => row.values), [{ 'parent.item_number': 'A-1', goal: 'Lighter', sort_order: '1' }]);
});

test('the file name of a workbook is the job number, without letters a file system refuses', () => {
	assert.equal(workbookFileName('MCE-000012'), 'MCE-000012.xlsx');
	assert.equal(workbookFileName('In/Out: "Test"?'), 'In_Out_ _Test__.xlsx');
	assert.equal(workbookFileName(''), 'workbook.xlsx');
});

// ---------------------------------------------------------------------------
// Together with the workbook module
// ---------------------------------------------------------------------------

test('rows made here come back out of a generated workbook', async () => {
	const { readFixture } = await import('./helpers/workbook-tools.mjs');
	const partSchema = readFixture('part-with-bom.schema.json');
	const [main, bom] = partSchema.sheets;
	const parents = [
		item(id32(1), { item_number: 'A-1', name: 'Frame', weight: '2.5', modified_on: '2026-03-01T10:00:00', owned_by_id: { value: id32(5), keyedName: 'Engineering' } }),
		item(id32(2), { item_number: 'A-2', name: 'Wheel' })
	];
	const related = (number) => ({ value: id32(9), keyedName: number, item: item(id32(9), { item_number: number }) });
	const bomRows = relationshipRows(bom, parents, [
		item(id32('b'), { source_id: id32(2), related_id: related('P-7'), sort_order: '10', quantity: '4', modified_on: '2026-03-02T08:30:00' }),
		item(id32('a'), { source_id: id32(1), related_id: related('P-3'), sort_order: '20', quantity: '1.5' })
	]);
	const job = { id: id32('F'), token: 'token-1' };
	const file = await buildWorkbook(partSchema, { job, rows: [mainRows(main, parents), bomRows.rows], maxRows: 10 });
	const parsed = await parseWorkbook(file);

	assert.deepEqual(parsed.meta.job, job);
	const [mainOut, bomOut] = parsed.sheets;
	assert.deepEqual(mainOut.rows.map((row) => [row.cells.item_number, row.cells.name, row.cells.weight, row.cells.owned_by_id, row.modifiedOn]), [
		['A-1', 'Frame', 2.5, 'Engineering', '2026-03-01T10:00:00'],
		// A value that is empty in Aras leaves no cell behind.
		['A-2', 'Wheel', undefined, undefined, null]
	]);
	assert.deepEqual(bomOut.rows.map((row) => [row.cells['parent.item_number'], row.cells['related.item_number'], row.cells.sort_order, row.cells.quantity, row.modifiedOn]), [
		['A-1', 'P-3', 20, 1.5, null],
		['A-2', 'P-7', 10, 4, '2026-03-02T08:30:00']
	]);
});

test('a workbook made from a narrowed schema has only the columns kept, and its meta lists only those', async () => {
	const { readFixture } = await import('./helpers/workbook-tools.mjs');
	const partSchema = readFixture('part-with-bom.schema.json');
	const options = {
		omit: ['description', 'weight', 'item_number'],
		sheets: [{ relationshipTypeId: partSchema.sheets[1].relationshipType.id, name: 'Part BOM', matchKey: [], omit: ['reference_designator', 'sort_order'] }]
	};
	const narrowed = narrowSchema(partSchema, options);
	const parents = [item(id32(1), { item_number: 'A-1', name: 'Frame', description: 'not written', weight: '2.5' })];
	const file = await buildWorkbook(narrowed, { job: { id: id32('F'), token: 't' }, rows: [mainRows(narrowed.sheets[0], parents), []], maxRows: 5 });
	const parsed = await parseWorkbook(file);

	// The key stays although omit names it, and so does Sequence, which is part of the BOM's match key.
	assert.deepEqual(parsed.meta.sheets[0].columns.map((column) => column.id),
		['item_number', 'name', 'classification', 'unit', 'make_buy', 'effective_date', 'owned_by_id']);
	assert.deepEqual(parsed.meta.sheets[1].columns.map((column) => column.id), ['parent.item_number', 'related.item_number', 'sort_order', 'quantity']);
	assert.deepEqual(parsed.meta.sheets[1].matchKeyColumns, ['parent.item_number', 'related.item_number', 'sort_order']);
	assert.deepEqual(parsed.sheets[0].rows[0].cells, { item_number: 'A-1', name: 'Frame' });
	assert.equal(parsed.meta.definitionHash, partSchema.definitionHash);
});

// ---------------------------------------------------------------------------
// The columns of a sheet
// ---------------------------------------------------------------------------

test('options keep the columns left out, per sheet, and write them only when there are any', () => {
	const options = {
		omit: ['cost', 'name'],
		sheets: [
			{ relationshipTypeId: BOM, name: 'Part BOM', matchKey: ['parent.item_number'], omit: ['quantity'] },
			{ relationshipTypeId: GOAL, name: 'Part Goal', matchKey: ['parent.item_number'] }
		]
	};
	assert.deepEqual(parseOptions(serializeOptions(options)), options);
	assert.doesNotMatch(serializeOptions({ omit: [], sheets: [{ relationshipTypeId: BOM, name: '', matchKey: [], omit: [] }] }), /omit/);
	// Repeated and empty ids, and ids that are not text, are dropped when reading.
	assert.deepEqual(parseOptions(JSON.stringify({ omit: ['cost', 'cost', '', 4], sheets: [] })), { omit: ['cost'], sheets: [] });
});

test('tidying options drops left-out columns a sheet no longer has, and keeps the others through every change', () => {
	const tidy = tidyOptions({
		omit: ['cost', 'gone'],
		sheets: [{ relationshipTypeId: BOM, name: '', matchKey: [], omit: ['quantity', 'gone_too'] }]
	}, sheetChoices(schema()), schema().sheets[0]);
	assert.deepEqual(tidy.omit, ['cost']);
	assert.deepEqual(tidy.sheets[0].omit, ['quantity']);
	// Without the main sheet, its list is kept as it is.
	assert.deepEqual(tidyOptions({ omit: ['gone'], sheets: [] }, sheetChoices(schema())).omit, ['gone']);
	assert.equal('omit' in tidyOptions({ sheets: [] }, sheetChoices(schema()), schema().sheets[0]), false);
});

test('the key, the match key, the parent and related columns and required properties without a default are locked', () => {
	const [main, bom] = schema().sheets;
	assert.equal(lockedReason(main.columns[0]), 'key');
	assert.equal(lockedReason(main.columns[1]), null);
	assert.equal(lockedReason(column('state', { required: true, defaultValue: null })), 'required');
	assert.equal(lockedReason(column('unit', { required: true, defaultValue: 'EA' })), null);
	assert.equal(lockedReason(bom.columns[0]), 'parent');
	assert.equal(lockedReason(bom.columns[1]), 'related item');
	// The match key chosen on the Form decides, not the one the schema was read with.
	assert.equal(lockedReason(bom.columns[3]), 'match key');
	assert.equal(lockedReason(bom.columns[3], ['parent.item_number']), null);
	assert.equal(lockedReason(bom.columns[2], ['parent.item_number', 'quantity']), 'match key');
});

test('the column picker lists every column, ticked unless left out, and a locked column is always ticked', () => {
	const [main] = schema().sheets;
	assert.deepEqual(columnChoices(main, ['item_number', 'cost']), [
		{ id: 'item_number', label: 'Part Number', locked: 'key', included: true },
		{ id: 'name', label: 'Name', locked: null, included: true },
		{ id: 'cost', label: 'Cost', locked: null, included: false },
		{ id: 'owned_by_id', label: 'Owner', locked: null, included: true }
	]);
});

test('narrowing a schema takes the left-out columns off their sheets and leaves the rest of the schema alone', () => {
	const whole = schema();
	const narrowed = narrowSchema(whole, { omit: ['cost', 'item_number'], sheets: [{ relationshipTypeId: GOAL, name: '', matchKey: [], omit: ['goal', 'sort_order'] }] });
	assert.deepEqual(narrowed.sheets[0].columns.map((c) => c.id), ['item_number', 'name', 'owned_by_id']);
	assert.equal(narrowed.sheets[1], whole.sheets[1]);
	// Sequence is in the Goal sheet's match key in the schema, so it stays.
	assert.deepEqual(narrowed.sheets[3].columns.map((c) => c.id), ['parent.item_number', 'sort_order']);
	assert.equal(narrowed.relationships, whole.relationships);
	assert.deepEqual(whole.sheets[0].columns.length, 4);
	assert.equal(narrowSchema(whole, { sheets: [] }).sheets[0], whole.sheets[0]);
});

// ---------------------------------------------------------------------------
// The allow-list and the user's identities
// ---------------------------------------------------------------------------

const ALLOWED_PART = { id: 'A'.repeat(32), name: 'Part', label: 'Part' };
const ALLOWED_DOCUMENT = { id: 'B'.repeat(32), name: 'Document', label: 'Document' };
const ALLOWED_CAD = { id: 'C'.repeat(32), name: 'CAD', label: 'CAD Document' };
const ENGINEERING_ID = 'E'.repeat(32);
const QUALITY_ID = 'F'.repeat(32);

test('a row without an identity allows its ItemType to everyone', () => {
	const { allowed, restricted } = allowedForUser([{ itemType: ALLOWED_PART, identityId: '', identityName: '' }], []);
	assert.deepEqual(allowed, [ALLOWED_PART]);
	assert.deepEqual(restricted, []);
});

test('a row with an identity allows its ItemType to members of that identity only', () => {
	const rows = [
		{ itemType: ALLOWED_PART, identityId: ENGINEERING_ID, identityName: 'Engineering' },
		{ itemType: ALLOWED_DOCUMENT, identityId: QUALITY_ID, identityName: 'Quality' }
	];
	// Ids may be listed in upper case; the comparison does not depend on it.
	const { allowed, restricted } = allowedForUser(rows, [ENGINEERING_ID.toLowerCase(), ' A73B655731924CD0B027E4F4D5FCC0A9']);
	assert.deepEqual(allowed, [ALLOWED_PART]);
	assert.deepEqual(restricted, [{ ...ALLOWED_DOCUMENT, identityName: 'Quality' }]);
});

test('an ItemType on several rows is allowed when any of them allows it', () => {
	const rows = [
		{ itemType: ALLOWED_PART, identityId: QUALITY_ID, identityName: 'Quality' },
		{ itemType: ALLOWED_PART, identityId: ENGINEERING_ID, identityName: 'Engineering' },
		{ itemType: ALLOWED_CAD, identityId: QUALITY_ID, identityName: 'Quality' },
		{ itemType: ALLOWED_CAD, identityId: '', identityName: '' }
	];
	const { allowed, restricted } = allowedForUser(rows, [ENGINEERING_ID]);
	assert.deepEqual(allowed.map((itemType) => itemType.name), ['CAD', 'Part']);
	assert.deepEqual(restricted, []);
});

test('rows without an ItemType are left out, and the lists are sorted by label', () => {
	const rows = [
		{ itemType: ALLOWED_PART, identityId: '' },
		null,
		{ itemType: null, identityId: '' },
		{ itemType: ALLOWED_DOCUMENT, identityId: '' },
		{ itemType: ALLOWED_CAD, identityId: QUALITY_ID, identityName: '' }
	];
	const { allowed, restricted } = allowedForUser(rows, []);
	assert.deepEqual(allowed.map((itemType) => itemType.label), ['Document', 'Part']);
	// An identity without a name is named by its id.
	assert.deepEqual(restricted.map((itemType) => itemType.identityName), [QUALITY_ID]);
});

// ---------------------------------------------------------------------------
// Several targets
// ---------------------------------------------------------------------------

const CAD = { id: 'C'.repeat(32), name: 'CAD', label: 'CAD Document', isVersionable: true };

test('options of a job with several targets keep the columns left out of each main sheet', () => {
	const options = {
		omit: ['cost'],
		sheets: [{ relationshipTypeId: BOM, name: 'Part BOM', matchKey: ['parent.item_number'] }],
		targets: [{ itemTypeId: CAD.id, omit: ['description'] }]
	};
	const text = serializeOptions(options);
	assert.equal(JSON.parse(text).version, 2);
	assert.deepEqual(parseOptions(text), options);
	// One target: written as version 1.
	assert.equal(JSON.parse(serializeOptions({ sheets: [] })).version, 1);
	// A target named twice, or without an id, is dropped when reading.
	assert.deepEqual(parseOptions(JSON.stringify({ sheets: [], targets: [{ itemTypeId: CAD.id }, { itemTypeId: CAD.id, omit: ['x'] }, { omit: ['y'] }] })).targets,
		[{ itemTypeId: CAD.id }]);
});

test('tidying options of several targets keeps the sheets of each target, and an entry per added target', () => {
	const cadSchema = {
		sheets: [{ kind: 'main', itemType: CAD, relationshipType: null, createOnly: false, columns: [column('item_number', { isKey: true }), column('description')] }],
		relationships: []
	};
	const tidy = tidyOptions({
		sheets: [{ relationshipTypeId: BOM, name: '', matchKey: [] }],
		targets: [{ itemTypeId: CAD.id, omit: ['description', 'gone'] }, { itemTypeId: MAKER.id, omit: ['x'] }]
	}, sheetChoices(schema()), schema().sheets[0], [{ itemTypeId: CAD.id, choices: sheetChoices(cadSchema), mainSheet: cadSchema.sheets[0] }]);
	assert.deepEqual(tidy.sheets.map((sheet) => sheet.relationshipTypeId), [BOM]);
	// The entries follow the targets given, not the stored ones.
	assert.deepEqual(tidy.targets, [{ itemTypeId: CAD.id, omit: ['description'] }]);
});

test('narrowing a schema with several main sheets takes each target\'s columns off its own main sheet', () => {
	const whole = schema();
	whole.sheets.splice(1, 0, { kind: 'main', itemType: CAD, relationshipType: null, createOnly: false, columns: [column('item_number', { isKey: true }), column('description')] });
	const narrowed = narrowSchema(whole, { omit: ['name'], sheets: [], targets: [{ itemTypeId: CAD.id, omit: ['description', 'item_number'] }] });
	assert.deepEqual(narrowed.sheets[0].columns.map((c) => c.id), ['item_number', 'cost', 'owned_by_id']);
	assert.deepEqual(narrowed.sheets[1].columns.map((c) => c.id), ['item_number']);
});

test('the targets of a job come from its targets property, or else from its target ItemType and selection', () => {
	const first = PART.id;
	const ids = ['A'.repeat(32), 'B'.repeat(32)];
	// One target: target_itemtype and prefill_idlist.
	assert.deepEqual(parseTargets('', first, ids.join(', ')), [{ itemTypeId: first, prefillIds: ids }]);
	assert.deepEqual(parseTargets('', '', ''), []);
	const two = [{ itemTypeId: first, prefillIds: ids }, { itemTypeId: CAD.id, prefillIds: [] }];
	const text = serializeTargets(two);
	assert.deepEqual(parseTargets(text, first.toLowerCase(), ''), two);
	// A job with one target keeps the property empty.
	assert.equal(serializeTargets(two.slice(0, 1)), '');
	// Repeated targets and ids, and entries that are no ids, are dropped.
	assert.deepEqual(parseTargets(JSON.stringify([
		{ itemTypeId: first, prefillIds: [ids[0], ids[0], 'x', 7] }, { itemTypeId: first.toLowerCase() }, { itemTypeId: 'nope' }, { itemTypeId: CAD.id.toLowerCase() }
	]), first, ''), [{ itemTypeId: first, prefillIds: [ids[0]] }, { itemTypeId: CAD.id, prefillIds: [] }]);
	// A list that can't be read, or that does not start with the job's target ItemType, counts as not there.
	assert.deepEqual(parseTargets('{oops', first, ids[1]), [{ itemTypeId: first, prefillIds: [ids[1]] }]);
	assert.deepEqual(parseTargets(JSON.stringify([{ itemTypeId: CAD.id }, { itemTypeId: first }]), first, ''), [{ itemTypeId: first, prefillIds: [] }]);
});

test('a choice names the ItemType of its related items, so the Form can tell when the workbook creates them', () => {
	const choices = sheetChoices(schema());
	const bom = choices.find((choice) => choice.relationshipTypeId === BOM);
	assert.equal(bom.relatedItemTypeId, PART.id);
});
