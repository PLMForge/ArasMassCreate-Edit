/**
 * Tests for the indented Part BOM: the option on the job, the workbook it gives, how an indented
 * sheet becomes ordinary rows, and what the preview and the import do with them. Run from the
 * repository root with `node --test`.
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const client = (name) => new URL('../../1.0.0/Innovator/Client/customer/masscreateedit/' + name, import.meta.url).href;
const { INDENTED_LAYOUT, indentedMatchKey, indentedRows, narrowSchema, readIndentedStructure, parseOptions, serializeOptions, sheetChoices, tidyOptions } = await import(client('mce-export.js'));
const { flattenIndented } = await import(client('mce-workbook.js'));
const { buildWorkbook, openWorkbook, parseWorkbook } = await import('./helpers/workbook-tools.mjs');
const { addIndentedResult, dropRowsBelowFailed, rowKey, sliceWorkbook } = await import(client('mce-import.js'));

const PART_BOM = '159C6D88795B4A86864420863466F728';
const PART_DOCUMENT = 'A'.repeat(32);
const RELATED = 'related.item_number';

/** The rows of an indented sheet: [level, part, more cells], from row 2. */
const sheetRows = (lines) => lines.map(([level, part, more], index) => ({
	rowNumber: index + 2,
	modifiedOn: null,
	cells: { ...(level === undefined ? {} : { level }), ...(part === undefined ? {} : { [RELATED]: part }), ...more }
}));
const flatten = (lines) => flattenIndented(sheetRows(lines), [RELATED]);
const pairs = (flat) => flat.rows.filter((row) => row.cells['parent.item_number']).map((row) => row.cells['parent.item_number'] + '>' + row.cells[RELATED]);
const codes = (flat) => Object.fromEntries(Object.entries(flat.indented.problems).map(([row, problem]) => [row, problem.code]));

// ---------------------------------------------------------------------------
// From levels to parents
// ---------------------------------------------------------------------------

test('a row\'s parent is the nearest row above it one level up, and level 0 writes no line', () => {
	const flat = flatten([[0, 'TOP'], [1, 'A'], [2, 'A1'], [2, 'A2'], [1, 'B'], [0, 'TOP2'], [1, 'C']]);
	assert.deepEqual(pairs(flat), ['TOP>A', 'A>A1', 'A>A2', 'TOP>B', 'TOP2>C']);
	assert.deepEqual(flat.rows.map((row) => row.rowNumber), [3, 4, 5, 6, 8], 'the rows keep the numbers of the sheet');
	assert.deepEqual(codes(flat), {});
	assert.equal('level' in flat.rows[0].cells, false, 'the level is not sent on');
});

test('a level can be a number or a text of digits', () => {
	assert.deepEqual(pairs(flatten([['0', 'TOP'], [' 1 ', 'A']])), ['TOP>A']);
});

test('the same subassembly under two parents, with matching levels below it, is no error and is written once', () => {
	const flat = flatten([
		[0, 'TOP'],
		[1, 'SUB-A'],
		[2, 'SHARED', { quantity: 1 }],
		[3, 'BOLT', { quantity: 4 }],
		[4, 'THREAD', { quantity: 1 }],
		[1, 'SUB-B'],
		[2, 'SHARED', { quantity: 2 }],
		[3, 'BOLT', { quantity: 4 }],
		[4, 'THREAD', { quantity: 1 }]
	]);
	assert.deepEqual(codes(flat), {});
	assert.deepEqual(pairs(flat), ['TOP>SUB-A', 'SUB-A>SHARED', 'SHARED>BOLT', 'BOLT>THREAD', 'TOP>SUB-B', 'SUB-B>SHARED']);
	assert.deepEqual(flat.rows.find((row) => row.cells['parent.item_number'] === 'SUB-B').cells.quantity, 2, 'each parent keeps its own quantity');
	assert.deepEqual(flat.indented.notes, { 5: ['The same line is also on row 9. It is written once.'], 6: ['The same line is also on row 10. It is written once.'] });
	assert.deepEqual(flat.indented.parents[5], [4, 8], 'BOLT is below SHARED in both places');
});

test('a later place that lists no lines refers to the first, and upper and lower case are the same part', () => {
	const flat = flatten([[0, 'TOP'], [1, 'S'], [2, 'X'], [1, 'T'], [2, 's']]);
	assert.deepEqual(codes(flat), {});
	assert.deepEqual(pairs(flat), ['TOP>S', 'S>X', 'TOP>T', 'T>s']);
	assert.deepEqual(flat.indented.parents[4], [3, 6]);
});

test('places that list different lines are errors, with everything below them', () => {
	const flat = flatten([[0, 'TOP'], [1, 'S'], [2, 'X', { quantity: 1 }], [3, 'Y'], [1, 'T'], [2, 'S'], [3, 'X', { quantity: 2 }]]);
	assert.equal(codes(flat)[4], 'BOM_DIFFERENT_PLACES');
	assert.equal(codes(flat)[8], 'BOM_DIFFERENT_PLACES');
	assert.equal(codes(flat)[5], 'LEVEL_PARENT_FAILED');
	assert.match(flat.indented.problems[4].message, /rows 3, 7/);
	assert.deepEqual(pairs(flat), ['TOP>S', 'TOP>T', 'T>S'], 'the other lines stay');
});

test('places that list different lines say why and what to do, on every row below them', () => {
	const flat = flatten([[0, 'TOP'], [1, 'S'], [2, 'X'], [2, 'Y'], [1, 'T'], [2, 'S'], [3, 'X']]);
	const message = 'Part S is on rows 3, 7 with different components below it, but a part can have only one BOM. Make the rows below them match, or list the components below only one of them.';
	assert.deepEqual([4, 5, 8].map((row) => flat.indented.problems[row].message), [message, message, message]);
	assert.deepEqual([4, 5, 8].map((row) => flat.indented.problems[row].comparison), [0, 0, 0]);
});

test('the places of a part are compared side by side: which lines each lists, and only the values that differ', () => {
	const flat = flattenIndented(sheetRows([
		[0, 'TOP'], [1, 'S'], [2, 'X', { quantity: 1, reference_designator: 'R1' }], [2, 'Y', { quantity: 3 }],
		[1, 'T'], [2, 'S'], [3, 'X', { quantity: 2, reference_designator: 'R1' }], [3, 'Z', { quantity: 5 }]
	]), [RELATED], { quantity: 'Quantity' });
	assert.deepEqual(flat.indented.comparisons, [{
		part: 'S',
		places: [{ row: 3, under: 'TOP' }, { row: 7, under: 'T' }],
		lines: [
			{ part: 'X', differs: true, cells: [{ row: 4, values: [{ label: 'Quantity', value: '1' }] }, { row: 8, values: [{ label: 'Quantity', value: '2' }] }] },
			{ part: 'Y', differs: true, cells: [{ row: 5, values: [] }, null] },
			{ part: 'Z', differs: true, cells: [null, { row: 9, values: [] }] }
		]
	}]);
});

test('a line the same in each place is shown as not differing, and a blank value is named', () => {
	const flat = flattenIndented(sheetRows([
		[0, 'TOP'], [1, 'S'], [2, 'X', { quantity: 1 }], [2, 'Y', { quantity: 1 }],
		[1, 'T'], [2, 'S'], [3, 'X', { quantity: 1 }], [3, 'Y']
	]), [RELATED], { quantity: 'Quantity' });
	const [same, other] = flat.indented.comparisons[0].lines;
	assert.equal(same.differs, false);
	assert.deepEqual(same.cells.map((cell) => cell.values), [[], []]);
	assert.deepEqual(other.cells.map((cell) => cell.values), [[{ label: 'Quantity', value: '1' }], [{ label: 'Quantity', value: 'blank' }]]);
	assert.equal(flat.indented.comparisons[0].places[0].under, 'TOP');
});

test('a part at the top and lower down is compared too, and the preview keeps the comparison with its rows', async () => {
	const flat = flatten([[0, 'S'], [1, 'X'], [0, 'TOP'], [1, 'S'], [2, 'Y']]);
	assert.deepEqual(flat.indented.comparisons[0].places, [{ row: 2, under: null }, { row: 5, under: 'TOP' }]);
	const result = { sheets: [{ rows: [{ row: 3, action: 'new', issues: [] }, { row: 5, action: 'new', issues: [] }] }] };
	addIndentedResult(result, { sheets: [{ indented: flat.indented }] });
	assert.equal(result.sheets[0].rows[0].comparison, 0);
	assert.equal(result.sheets[0].rows[1].comparison, undefined);
	const { comparisonOf } = await import(client('mce-preview.js'));
	assert.equal(comparisonOf(result, 0, result.sheets[0].rows[0]).part, 'S');
	assert.equal(comparisonOf(result, 0, result.sheets[0].rows[1]), null);
});

test('the same child twice directly under one parent is an error, not merged', () => {
	const flat = flatten([[0, 'TOP'], [1, 'X', { quantity: 1 }], [1, 'X', { quantity: 1 }], [1, 'Y']]);
	assert.deepEqual(codes(flat), { 3: 'BOM_SAME_CHILD', 4: 'BOM_SAME_CHILD' });
	assert.deepEqual(pairs(flat), ['TOP>Y']);
});

test('a broken structure is an error on the row and on the rows below it, and nothing is guessed', () => {
	assert.deepEqual(codes(flatten([[1, 'A'], [2, 'B']])), { 2: 'LEVEL_JUMP', 3: 'LEVEL_PARENT_FAILED' });
	assert.match(flatten([[1, 'A']]).indented.problems[2].message, /first row must be level 0/);
	const jump = flatten([[0, 'TOP'], [2, 'B'], [3, 'C'], [1, 'D']]);
	assert.deepEqual(codes(jump), { 3: 'LEVEL_JUMP', 4: 'LEVEL_PARENT_FAILED' });
	assert.deepEqual(pairs(jump), ['TOP>D']);
	for (const bad of [undefined, -1, 1.5, 'x']) {
		const flat = flatten([[0, 'TOP'], [1, 'A'], [bad, 'X'], [2, 'B'], [1, 'C']]);
		assert.deepEqual(codes(flat), { 4: 'LEVEL_INVALID', 5: 'LEVEL_NO_PARENT' }, String(bad));
		assert.deepEqual(pairs(flat), ['TOP>A', 'TOP>C'], String(bad));
	}
	const blank = flatten([[0, 'TOP'], [1, undefined], [2, 'B'], [1, 'C']]);
	assert.deepEqual(codes(blank), { 3: 'LEVEL_NO_PART', 4: 'LEVEL_PARENT_FAILED' });
	assert.deepEqual(codes(flatten([[0, 'TOP']])), {}, 'a top assembly without lines writes nothing, and is no error');
});

test('a row with a problem comes back without a parent, so that the server refuses it too', () => {
	const flat = flatten([[0, 'TOP'], [1, 'A'], [3, 'B']]);
	assert.deepEqual(flat.rows[1], { rowNumber: 4, cells: { [RELATED]: 'B' }, modifiedOn: null });
	assert.deepEqual(codes(flatten([[0, 'TOP'], [1, undefined]])), { 3: 'LEVEL_NO_PART' }, 'the top assembly has a line, even though the line has a problem');
});

test('parts that contain each other are errors, in one branch and across branches', () => {
	assert.equal(codes(flatten([[0, 'A'], [1, 'B'], [2, 'A']]))[4], 'BOM_CYCLE');
	assert.equal(codes(flatten([[0, 'A'], [1, 'A']]))[3], 'BOM_CYCLE');
	const across = flatten([[0, 'A'], [1, 'B'], [0, 'B'], [1, 'C'], [0, 'C'], [1, 'A'], [0, 'D'], [1, 'E']]);
	assert.deepEqual(codes(across), { 3: 'BOM_CYCLE', 5: 'BOM_CYCLE', 7: 'BOM_CYCLE' });
	assert.deepEqual(pairs(across), ['D>E'], 'lines outside the ring stay');
});

const messages = (flat) => Object.fromEntries(Object.entries(flat.indented.problems).map(([row, problem]) => [row, problem.message]));

test('a part inside itself names both rows and the way from one to the other', () => {
	assert.deepEqual(messages(flatten([[0, 'A'], [1, 'B'], [2, 'A']])), { 4: 'Rows 2, 4 are the same part, one inside the other: A > B > A.' });
	assert.deepEqual(messages(flatten([[0, 'TOP'], [1, 'A'], [2, 'A']])), { 4: 'Rows 3, 4 are the same part, one inside the other: A > A.' });
});

test('parts that contain each other across branches name the rows of the ring, and every row says the same', () => {
	const two = 'Rows 3, 5 make parts contain each other: A > B > A.';
	assert.deepEqual(messages(flatten([[0, 'A'], [1, 'B'], [0, 'B'], [1, 'A']])), { 3: two, 5: two });
	const three = 'Rows 3, 5, 7 make parts contain each other: A > B > C > A.';
	assert.deepEqual(messages(flatten([[0, 'A'], [1, 'B'], [0, 'B'], [1, 'C'], [0, 'C'], [1, 'A']])), { 3: three, 5: three, 7: three });
});

test('a line that only joins two rings is no ring itself', () => {
	const flat = flatten([[0, 'A'], [1, 'B'], [0, 'B'], [1, 'A'], [1, 'C'], [0, 'C'], [1, 'D'], [0, 'D'], [1, 'C']]);
	assert.deepEqual(codes(flat), { 3: 'BOM_CYCLE', 5: 'BOM_CYCLE', 8: 'BOM_CYCLE', 10: 'BOM_CYCLE' });
	assert.deepEqual(pairs(flat), ['B>C']);
});

test('a new line gets its place under its parent, unless it has a Sequence', () => {
	const flat = flatten([[0, 'TOP'], [1, 'A'], [1, 'B', { sort_order: 5 }], [1, 'C'], [2, 'D']]);
	assert.deepEqual(flat.rows.map((row) => row.sequenceIfNew), [10, undefined, 30, 10]);
});

// ---------------------------------------------------------------------------
// The option on the job
// ---------------------------------------------------------------------------

const schemaWithRelationships = () => {
	const schema = JSON.parse(readFileSync(new URL('./fixtures/part-with-bom.schema.json', import.meta.url), 'utf8'));
	schema.sheets[1].createOnly = false;
	const related = schema.sheets[1].relationshipType.relatedItemType;
	schema.relationships = [
		{ id: PART_BOM, name: 'Part BOM', label: 'BOM', relatedItemType: related, canBeSheet: true, canBeIndented: true, reason: null },
		{ id: PART_DOCUMENT, name: 'Part Document', label: 'Documents', relatedItemType: { id: 'B'.repeat(32), name: 'Document' }, canBeSheet: true, canBeIndented: false, reason: null }
	];
	schema.sheets.push({ ...schema.sheets[1], relationshipType: { ...schema.sheets[1].relationshipType, id: PART_DOCUMENT, name: 'Part Document' } });
	return schema;
};

test('only a relationship the server marks can keep the indented layout', () => {
	const schema = schemaWithRelationships();
	const choices = sheetChoices(schema);
	assert.deepEqual(choices.map((choice) => choice.canBeIndented), [true, false]);
	const stored = parseOptions(JSON.stringify({ sheets: [
		{ relationshipTypeId: PART_BOM, name: 'Part BOM', matchKey: ['parent.item_number', 'related.item_number', 'sort_order'], layout: 'indented' },
		{ relationshipTypeId: PART_DOCUMENT, name: 'Part Document', matchKey: [], layout: 'indented' }
	] }));
	const tidy = tidyOptions(stored, choices, schema.sheets[0]);
	assert.equal(tidy.sheets[0].layout, INDENTED_LAYOUT);
	assert.deepEqual(tidy.sheets[0].matchKey, ['parent.item_number', 'related.item_number', 'sort_order'], 'the match key chosen is kept');
	assert.equal('layout' in tidy.sheets[1], false);
	const fresh = tidyOptions(parseOptions(JSON.stringify({ sheets: [{ relationshipTypeId: PART_BOM, name: 'Part BOM', matchKey: [], layout: 'indented' }] })), choices, schema.sheets[0]);
	assert.deepEqual(fresh.sheets[0].matchKey, ['parent.item_number', 'related.item_number'], 'without a choice: parent and child, without Sequence');
	assert.deepEqual(indentedMatchKey(choices[0]), fresh.sheets[0].matchKey);
});

// ---------------------------------------------------------------------------
// Deleting lines, and an existing structure in the workbook
// ---------------------------------------------------------------------------

test('a line marked #DELETE goes on to the server, and is not part of a loop or asked for a Sequence', () => {
	const flat = flatten([[0, 'A'], [1, 'B', { delete: '#DELETE' }], [2, 'A', { delete: '#delete ' }]]);
	assert.deepEqual(codes(flat), {});
	assert.deepEqual(pairs(flat), ['A>B', 'B>A']);
	assert.deepEqual(codes(flattenIndented(sheetRows([[0, 'TOP'], [1, 'X', { delete: '#DELETE' }]]), [RELATED], {}, [RELATED, 'sort_order'])), {});
	const across = flatten([[0, 'A'], [1, 'B'], [0, 'B'], [1, 'A', { delete: '#DELETE' }]]);
	assert.deepEqual(codes(across), {}, 'deleting one line of a ring opens it');
});

test('a line deleted in only one place of a part is a difference between its places', () => {
	const flat = flattenIndented(sheetRows([[0, 'TOP'], [1, 'S'], [2, 'X'], [1, 'T'], [2, 'S'], [3, 'X', { delete: '#DELETE' }]]), [RELATED], { delete: 'Delete' });
	assert.equal(flat.indented.problems[4].code, 'BOM_DIFFERENT_PLACES');
	assert.deepEqual(flat.indented.comparisons[0].lines[0].cells.map((cell) => cell.values), [[{ label: 'Delete', value: 'blank' }], [{ label: 'Delete', value: '#DELETE' }]]);
});

const bomSheet = (match = [RELATED]) => ({
	relationshipType: { id: PART_BOM }, layout: INDENTED_LAYOUT, itemType: { name: 'Part BOM' },
	columns: [
		{ id: 'level', role: 'level' },
		{ id: RELATED, role: 'relatedKey', property: 'item_number', dataType: 'string', isMatchKey: match.includes(RELATED) },
		{ id: 'sort_order', role: 'property', property: 'sort_order', dataType: 'integer', isMatchKey: match.includes('sort_order') },
		{ id: 'quantity', role: 'property', property: 'quantity', dataType: 'decimal', isMatchKey: false }
	]
});
const prop = (value) => ({ value, keyedName: '', item: null, restricted: false });
const part = (id, number, config = 'C' + id) => ({ id, props: { item_number: prop(number), config_id: prop(config) } });
let lineIds = 0;
const bomLine = (source, child, sequence, quantity) => ({
	id: 'L' + (++lineIds), props: {
		source_id: prop(source), sort_order: prop(String(sequence)), quantity: prop(String(quantity)), modified_on: prop('2026-01-02T03:04:05'),
		related_id: { value: child ? child.id : '', keyedName: '', item: child, restricted: false }
	}
});
/** Answers the two queries of readIndentedStructure from a list of lines and the current generations. */
const fakeQuery = (lines, current) => async (aml) => {
	const ids = /condition="in">([^<]*)</.exec(aml)[1].split(',').map((id) => id.replace(/'/g, ''));
	return aml.includes('type="Part BOM"')
		? lines.filter((line) => ids.includes(line.props.source_id.value))
		: current.filter((item) => ids.includes(item.props.config_id.value));
};
const shape = (read) => read.rows.map((row) => [row.values.level, row.values[RELATED], row.values.quantity].filter((value) => value !== undefined));

const ID = (letter) => letter.repeat(32);
const TOP = part(ID('A'), 'TOP', ID('1'));
const SUBA = part(ID('B'), 'SUB-A', ID('2'));
const SUBB = part(ID('C'), 'SUB-B', ID('3'));
const SHARED = part(ID('D'), 'SHARED', ID('4'));
const BOLT = part(ID('E'), 'BOLT', ID('5'));
const everything = [TOP, SUBA, SUBB, SHARED, BOLT];

test('an existing structure is read level by level and listed in Sequence order, in full in every place', async () => {
	const lines = [
		bomLine(TOP.id, SUBB, 20, 1), bomLine(TOP.id, SUBA, 10, 1),
		bomLine(SUBA.id, SHARED, 10, 1), bomLine(SUBB.id, SHARED, 10, 2),
		bomLine(SHARED.id, BOLT, 10, 4)
	];
	const sheet = bomSheet();
	const read = indentedRows(sheet, [TOP], await readIndentedStructure(fakeQuery(lines, everything), sheet, [TOP], 'Part'));
	assert.deepEqual(shape(read), [[0, 'TOP'], [1, 'SUB-A', '1'], [2, 'SHARED', '1'], [3, 'BOLT', '4'], [1, 'SUB-B', '1'], [2, 'SHARED', '2'], [3, 'BOLT', '4']]);
	assert.equal(read.rows[0].modifiedOn, null, 'a top assembly is no line');
	assert.equal(read.rows[1].modifiedOn, '2026-01-02T03:04:05');
	assert.deepEqual([read.loops, read.sameLine, read.unreadable], [0, 0, 0]);
});

test('such a structure goes through the check with no errors, and repeated lines are written once', async () => {
	const lines = [bomLine(TOP.id, SUBA, 10, 1), bomLine(TOP.id, SUBB, 20, 1), bomLine(SUBA.id, SHARED, 10, 1), bomLine(SUBB.id, SHARED, 10, 2), bomLine(SHARED.id, BOLT, 10, 4)];
	const sheet = bomSheet();
	const read = indentedRows(sheet, [TOP], await readIndentedStructure(fakeQuery(lines, everything), sheet, [TOP], 'Part'));
	const flat = flattenIndented(read.rows.map((row, index) => ({ rowNumber: index + 2, cells: row.values, modifiedOn: row.modifiedOn })), [RELATED]);
	assert.deepEqual(codes(flat), {});
	assert.deepEqual(pairs(flat), ['TOP>SUB-A', 'SUB-A>SHARED', 'SHARED>BOLT', 'TOP>SUB-B', 'SUB-B>SHARED']);
});

test('a line follows the current generation of its part', async () => {
	const old = part(ID('F'), 'SUB-A', SUBA.props.config_id.value);
	const lines = [bomLine(TOP.id, old, 10, 1), bomLine(old.id, BOLT, 10, 9), bomLine(SUBA.id, BOLT, 10, 2)];
	const sheet = bomSheet();
	const read = indentedRows(sheet, [TOP], await readIndentedStructure(fakeQuery(lines, everything), sheet, [TOP], 'Part'));
	assert.deepEqual(shape(read), [[0, 'TOP'], [1, 'SUB-A', '1'], [2, 'BOLT', '2']]);
});

test('a line that leads back to a part above it is listed and not followed, and a top with no lines stays', async () => {
	const lines = [bomLine(TOP.id, SUBA, 10, 1), bomLine(SUBA.id, TOP, 10, 1)];
	const sheet = bomSheet();
	const read = indentedRows(sheet, [TOP, BOLT], await readIndentedStructure(fakeQuery(lines, everything), sheet, [TOP, BOLT], 'Part'));
	assert.deepEqual(shape(read), [[0, 'TOP'], [1, 'SUB-A', '1'], [2, 'TOP', '1'], [0, 'BOLT']]);
	assert.equal(read.loops, 1);
});

test('two lines the match columns cannot tell apart are counted, and Sequence tells them apart', async () => {
	const lines = [bomLine(TOP.id, BOLT, 10, 1), bomLine(TOP.id, BOLT, 20, 1)];
	const byPart = bomSheet();
	assert.equal(indentedRows(byPart, [TOP], await readIndentedStructure(fakeQuery(lines, everything), byPart, [TOP], 'Part')).sameLine, 1);
	const bySequence = bomSheet([RELATED, 'sort_order']);
	assert.equal(indentedRows(bySequence, [TOP], await readIndentedStructure(fakeQuery(lines, everything), bySequence, [TOP], 'Part')).sameLine, 0);
});

test('an unreadable part is left out and counted once, and too many rows stop the walk', async () => {
	const hidden = { id: 'X', props: { source_id: prop(TOP.id), sort_order: prop('10'), quantity: prop('1'), related_id: { value: ID('9'), keyedName: '', item: null, restricted: false } } };
	const sheet = bomSheet();
	const lines = [hidden, bomLine(TOP.id, BOLT, 20, 1)];
	const structure = await readIndentedStructure(fakeQuery(lines, everything), sheet, [TOP], 'Part');
	const read = indentedRows(sheet, [TOP], structure);
	assert.deepEqual(shape(read), [[0, 'TOP'], [1, 'BOLT', '1']]);
	assert.equal(read.unreadable, 1);
	assert.equal(indentedRows(sheet, [TOP], structure, 1).tooMany, true);
});

test('only an indented sheet is given the column that deletes a line', () => {
	const schema = JSON.parse(readFileSync(new URL('./fixtures/part-with-bom.schema.json', import.meta.url), 'utf8'));
	schema.sheets[1].columns.push({ id: 'delete', role: 'delete', dataType: 'delete', indentedOnly: true, label: 'Delete' });
	const flat = narrowSchema(schema, { sheets: [{ relationshipTypeId: PART_BOM, matchKey: [] }] });
	assert.equal(flat.sheets[1].columns.some((column) => column.id === 'delete'), false);
	const indented = narrowSchema(schema, { sheets: [{ relationshipTypeId: PART_BOM, matchKey: [], layout: INDENTED_LAYOUT }] });
	assert.equal(indented.sheets[1].columns.at(-1).id, 'delete');
});

const SEQUENCE_KEY = [RELATED, 'sort_order'];
const flattenBy = (matchIds, lines) => flattenIndented(sheetRows(lines), [RELATED], { [RELATED]: 'Part Number', sort_order: 'Sequence', quantity: 'Quantity' }, matchIds);

test('with Sequence in the match key, the same part can be listed twice under a parent', () => {
	const flat = flattenBy(SEQUENCE_KEY, [[0, 'TOP'], [1, 'X', { sort_order: 10, quantity: 1 }], [1, 'X', { sort_order: 20, quantity: 2 }]]);
	assert.deepEqual(codes(flat), {});
	assert.deepEqual(flat.rows.map((row) => [row.cells.sort_order, row.cells.quantity]), [[10, 1], [20, 2]]);
});

test('with Sequence in the match key, a line needs a Sequence, and a top assembly does not', () => {
	const flat = flattenBy(SEQUENCE_KEY, [[0, 'TOP'], [1, 'X', { sort_order: 10 }], [1, 'Y'], [2, 'Z', { sort_order: 10 }]]);
	assert.deepEqual(codes(flat), { 4: 'BOM_SEQUENCE_REQUIRED', 5: 'LEVEL_PARENT_FAILED' });
	assert.equal(flat.indented.problems[4].message, "Sequence is one of the columns that match this line, so it can't be blank.");
});

test('the same match key twice under a parent says which columns make the rows the same', () => {
	const flat = flattenBy(SEQUENCE_KEY, [[0, 'TOP'], [1, 'X', { sort_order: 10 }], [1, 'X', { sort_order: 10, quantity: 3 }]]);
	assert.deepEqual(codes(flat), { 3: 'BOM_SAME_CHILD', 4: 'BOM_SAME_CHILD' });
	assert.equal(flat.indented.problems[3].message, 'Rows 3, 4 are the same line under the same parent, because they match on Part Number, Sequence. Change one of them, or put them on one row.');
});

test('matched on parent and Sequence alone, the part is a value of the line', () => {
	const key = ['sort_order'];
	const twice = flattenBy(key, [[0, 'TOP'], [1, 'X', { sort_order: 10 }], [1, 'Y', { sort_order: 20 }]]);
	assert.deepEqual(codes(twice), {}, 'different parts, different Sequence');
	const clash = flattenBy(key, [[0, 'TOP'], [1, 'X', { sort_order: 10 }], [1, 'Y', { sort_order: 10 }]]);
	assert.deepEqual(codes(clash), { 3: 'BOM_SAME_CHILD', 4: 'BOM_SAME_CHILD' });
	const places = flattenBy(key, [[0, 'TOP'], [1, 'S', { sort_order: 10 }], [2, 'X', { sort_order: 10 }], [1, 'T', { sort_order: 20 }], [2, 'S', { sort_order: 10 }], [3, 'Y', { sort_order: 10 }]]);
	const line = places.indented.comparisons[0].lines[0];
	assert.equal(line.part, 'Sequence: 10');
	assert.deepEqual(line.cells.map((cell) => cell.values), [[{ label: 'Part Number', value: 'X' }], [{ label: 'Part Number', value: 'Y' }]]);
});

test('with Sequence in the match key, a repeated subassembly must repeat its Sequences, and the table says so', () => {
	const same = flattenBy(SEQUENCE_KEY, [[0, 'TOP'], [1, 'S', { sort_order: 10 }], [2, 'X', { sort_order: 10 }], [1, 'T', { sort_order: 20 }], [2, 'S', { sort_order: 10 }], [3, 'X', { sort_order: 10 }]]);
	assert.deepEqual(codes(same), {});
	assert.deepEqual(same.indented.notes[4], ['The same line is also on row 7. It is written once.']);
	const other = flattenBy(SEQUENCE_KEY, [[0, 'TOP'], [1, 'S', { sort_order: 10 }], [2, 'X', { sort_order: 10 }], [1, 'T', { sort_order: 20 }], [2, 'S', { sort_order: 10 }], [3, 'X', { sort_order: 20 }]]);
	assert.deepEqual(other.indented.comparisons[0].lines.map((line) => [line.part, line.cells.map((cell) => cell && cell.row)]),
		[['X, Sequence: 10', [4, null]], ['X, Sequence: 20', [null, 7]]]);
});

test('the workbook tells the check which columns match the lines', async () => {
	const schema = JSON.parse(readFileSync(new URL('./fixtures/part-with-bom.schema.json', import.meta.url), 'utf8'));
	const narrowed = narrowSchema(schema, { sheets: [{ relationshipTypeId: PART_BOM, name: 'Part BOM', matchKey: ['parent.item_number', RELATED, 'sort_order'], layout: INDENTED_LAYOUT }] });
	assert.equal(narrowed.sheets[1].columns.find((column) => column.id === 'sort_order').isMatchKey, true, 'the fixture matches on Sequence');
	const rows = [[], [value(0, 'TOP'), value(1, 'X', { sort_order: 10 }), value(1, 'X', { sort_order: 20 }), value(1, 'Y')]];
	const parsed = await parseWorkbook(await buildWorkbook(narrowed, { job: { id: 'J'.repeat(32), token: 'T' }, rows, maxRows: 10 }));
	assert.deepEqual(Object.fromEntries(Object.entries(parsed.sheets[1].indented.problems).map(([row, problem]) => [row, problem.code])), { 5: 'BOM_SEQUENCE_REQUIRED' });
});

test('the layout is stored with the options, and any other value is dropped', () => {
	const text = serializeOptions({ sheets: [{ relationshipTypeId: PART_BOM, name: 'Part BOM', matchKey: ['x'], layout: INDENTED_LAYOUT }] });
	assert.equal(parseOptions(text).sheets[0].layout, INDENTED_LAYOUT);
	assert.equal('layout' in parseOptions(JSON.stringify({ sheets: [{ relationshipTypeId: PART_BOM, layout: 'sideways' }] })).sheets[0], false);
});

const indentedOptions = { sheets: [{ relationshipTypeId: PART_BOM, name: 'Part BOM', matchKey: ['parent.item_number', RELATED], layout: INDENTED_LAYOUT }] };
const indentedSchema = () => {
	const schema = JSON.parse(readFileSync(new URL('./fixtures/part-with-bom.schema.json', import.meta.url), 'utf8'));
	schema.sheets[1].columns.find((column) => column.id === 'sort_order').isMatchKey = false;
	return narrowSchema(schema, indentedOptions);
};

test('the schema of an indented sheet has a Level column in place of the parent', () => {
	const sheet = indentedSchema().sheets[1];
	assert.equal(sheet.layout, INDENTED_LAYOUT);
	assert.deepEqual(sheet.columns.map((column) => column.id), ['level', RELATED, 'sort_order', 'quantity', 'reference_designator']);
});

// ---------------------------------------------------------------------------
// The workbook
// ---------------------------------------------------------------------------

const value = (level, part, more) => ({ values: { level, [RELATED]: part, ...more }, modifiedOn: null });

test('an indented sheet comes back from the workbook as ordinary rows the server knows', async () => {
	const schema = indentedSchema();
	const rows = [[], [value(0, 'TOP'), value(1, 'A', { quantity: 2 }), value(2, 'B')]];
	const parsed = await parseWorkbook(await buildWorkbook(schema, { job: { id: 'J'.repeat(32), token: 'T' }, rows, maxRows: 10 }));
	assert.equal(parsed.meta.formatVersion, 1);
	const sheet = parsed.sheets[1];
	assert.deepEqual(sheet.columns, ['parent.item_number', RELATED, 'sort_order', 'quantity', 'reference_designator']);
	assert.deepEqual(sheet.rows.map((row) => [row.rowNumber, row.cells['parent.item_number'], row.cells[RELATED], row.cells.quantity, row.sequenceIfNew]),
		[[3, 'TOP', 'A', 2, 10], [4, 'A', 'B', undefined, 10]]);
	assert.deepEqual(parsed.meta.sheets[1].columns.map((column) => [column.id, column.role, column.isKey]).slice(0, 2),
		[['parent.item_number', 'parentKey', true], [RELATED, 'relatedKey', true]]);
	assert.deepEqual(parsed.meta.sheets[1].keyColumns.includes('parent.item_number'), true);
	assert.equal(parsed.meta.sheets[1].columns.some((column) => column.id === 'level'), false);
});

test('a message about a difference uses the heading of the column in the workbook', async () => {
	const rows = [[], [value(0, 'TOP'), value(1, 'S'), value(2, 'X', { quantity: 1 }), value(1, 'T'), value(2, 'S'), value(3, 'X', { quantity: 2 })]];
	const parsed = await parseWorkbook(await buildWorkbook(indentedSchema(), { job: { id: 'J'.repeat(32), token: 'T' }, rows, maxRows: 10 }));
	assert.deepEqual(parsed.sheets[1].indented.comparisons[0].lines[0].cells[0].values, [{ label: 'Quantity', value: '1' }]);
});

test('rows can be inserted only on an indented sheet', async () => {
	const indented = openWorkbook(await buildWorkbook(indentedSchema(), { job: { id: 'J'.repeat(32), token: 'T' }, maxRows: 10 }));
	assert.equal(indented.sheet('_meta').cell('B2').value, 1);
	const [mainSheet, bomSheet] = indented.sheets;
	assert.match(String(bomSheet.cell('A1').value), /^Level/);
	assert.equal(bomSheet.protection.insertRows, false, 'rows may be inserted');
	assert.notEqual(mainSheet.protection.insertRows, false);

	const plainFile = await buildWorkbook(JSON.parse(readFileSync(new URL('./fixtures/part-with-bom.schema.json', import.meta.url), 'utf8')), { maxRows: 10 });
	const plain = openWorkbook(plainFile);
	assert.notEqual(plain.sheets[1].protection.insertRows, false);
	assert.equal((await parseWorkbook(plainFile)).sheets[1].indented, undefined);
});

// ---------------------------------------------------------------------------
// The preview and the import
// ---------------------------------------------------------------------------

/** A validation result for the flat rows, every row a new line unless it has a problem. */
function resultFor(flat, parents = []) {
	return {
		sheets: [
			{ sheetName: 'Part', kind: 'main', rows: [] },
			{
				sheetName: 'Part BOM', kind: 'relationship', parents,
				rows: flat.rows.map((row) => flat.indented.problems[row.rowNumber]
					? { row: row.rowNumber, action: 'error', issues: [{ severity: 'error', code: 'KEY_BLANK', message: 'server' }], parent: { id: null } }
					: { row: row.rowNumber, action: 'new', issues: [], parent: { id: row.cells['parent.item_number'] } })
			}
		]
	};
}
const workbookOf = (flat) => ({ meta: {}, sheets: [{ sheetName: 'Part', columns: [], rows: [] }, { sheetName: 'Part BOM', columns: [], ...flat }] });

test('the preview shows the problem of the structure instead of the server\'s, and the notes as warnings', () => {
	const flat = flatten([[0, 'TOP'], [1, 'S'], [2, 'X'], [0, 'T2'], [1, 'S'], [2, 'X'], [3, 'Y', { quantity: 1 }], [1, 'R'], [3, 'Z']]);
	const result = addIndentedResult(resultFor(flat), workbookOf(flat));
	const row = (number) => result.sheets[1].rows.find((candidate) => candidate.row === number);
	assert.equal(row(10).action, 'error');
	assert.deepEqual(row(10).issues.map((issue) => issue.code), ['LEVEL_JUMP']);
	assert.deepEqual(row(4).issues.map((issue) => [issue.severity, issue.code]), [['warning', 'BOM_SAME_LINE']]);
});

test('a line with an error takes the rows below it along, unless the subassembly stands elsewhere without one', () => {
	const flat = flatten([[0, 'TOP'], [1, 'S'], [2, 'X'], [3, 'Y'], [1, 'T'], [2, 'U'], [0, 'T2'], [1, 'U'], [2, 'V']]);
	const failing = resultFor(flat);
	const row = (result, number) => result.sheets[1].rows.find((candidate) => candidate.row === number);
	row(failing, 3).action = 'error';
	row(failing, 7).action = 'error';
	addIndentedResult(failing, workbookOf(flat));
	assert.equal(row(failing, 4).action, 'error', 'below the failed line');
	assert.equal(row(failing, 5).action, 'error', 'two levels below');
	assert.match(row(failing, 5).issues[0].message, /row 4/);
	assert.equal(row(failing, 10).action, 'new', 'U is also a line of T2, which did not fail');
});

test('lines a parent has in Aras that the sheet does not list stay, with a warning on the parent\'s first row', () => {
	const flat = flatten([[0, 'TOP'], [1, 'A'], [1, 'B']]);
	const result = addIndentedResult(resultFor(flat, [{ id: 'TOP', existing: 3, untouched: 2 }, { id: 'OTHER', existing: 1, untouched: 0 }]), workbookOf(flat));
	assert.deepEqual(result.sheets[1].rows[0].issues.map((issue) => issue.code), ['BOM_LINES_NOT_ON_SHEET']);
	assert.match(result.sheets[1].rows[0].issues[0].message, /2 BOM lines/);
	assert.deepEqual(result.sheets[1].rows[1].issues, []);
});

test('a sheet that is not indented is left alone', () => {
	const result = { sheets: [{ rows: [{ row: 2, action: 'new', issues: [] }] }] };
	const before = JSON.stringify(result);
	assert.equal(JSON.stringify(addIndentedResult(result, { sheets: [{ rows: [] }] })), before);
});

test('on import, rows below a line that failed are left out, in every place it stands', () => {
	const flat = flatten([[0, 'TOP'], [1, 'S'], [2, 'X'], [1, 'T'], [2, 'S'], [3, 'X'], [0, 'T2'], [1, 'Y'], [2, 'Z']]);
	const result = resultFor(flat);
	const workbook = workbookOf(flat);
	const indexOf = (number) => result.sheets[1].rows.findIndex((row) => row.row === number);
	const batch = [{ main: null, mainSheet: 0, writesMain: false, parentId: 'x', rels: [[1, indexOf(4)], [1, indexOf(10)]] }];
	assert.deepEqual(dropRowsBelowFailed(batch, result, workbook, new Set()), []);
	assert.deepEqual(dropRowsBelowFailed(batch, result, workbook, new Set([rowKey(1, 3)])), [], 'S also stands under T');
	const dropped = dropRowsBelowFailed(batch, result, workbook, new Set([rowKey(1, 3), rowKey(1, 6)]));
	assert.deepEqual(dropped.map((entry) => entry.rowIndex), [indexOf(4)]);
	assert.deepEqual(batch[0].rels, [[1, indexOf(10)]]);
});

test('a batch carries the place of a new line to the server', () => {
	const flat = flatten([[0, 'TOP'], [1, 'A']]);
	const result = resultFor(flat);
	result.sheets[1].rows[0].digest = 'd';
	const slice = sliceWorkbook(workbookOf(flat), result, [{ main: null, mainSheet: 0, writesMain: false, parentId: 'x', rels: [[1, 0]] }]);
	assert.equal(slice.sheets[1].rows[0].sequenceIfNew, 10);
});
