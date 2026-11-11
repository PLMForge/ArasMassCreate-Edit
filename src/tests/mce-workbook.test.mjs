/**
 * Tests for the workbook module (mce-workbook.js). Run from the repository
 * root with `node --test`; nothing has to be installed.
 *
 * A generated workbook is checked three ways: by reading it back with the
 * module itself, by opening it with the test tools' own reader, which shares no
 * code with the module (validation, notes, protection), and by looking at the
 * XML parts of the file.
 */

import test from 'node:test';
import assert from 'node:assert/strict';

import { buildWorkbook, editWorkbook, openWorkbook, parseWorkbook, plainWorkbook, readFixture, readParts } from './helpers/workbook-tools.mjs';
import { CLASS_PATHS, allTypesSchema, columnAt } from './fixtures/all-types.schema.mjs';

const partSchema = () => readFixture('part-with-bom.schema.json');
const JOB = { id: '0123456789ABCDEF0123456789ABCDEF', token: 'c2FtcGxlLXRva2Vu' };

/** Builds the all-types workbook with ten validated rows and opens it for inspection. */
async function openAllTypes(options = {}) {
	const schema = allTypesSchema();
	const file = await buildWorkbook(schema, { maxRows: 10, ...options });
	const workbook = openWorkbook(file);
	const sheet = workbook.sheet('Sample Item');
	const letter = (id) => columnAt(schema, id).letter;
	return { schema, file, workbook, sheet, letter, validation: (id) => sheet.cell(letter(id) + '2').validation };
}

const cellOf = (id, row = 2) => (letter) => letter(id) + row;

test('Part with BOM: sheets, header row, protection and hidden sheets', async () => {
	const file = await buildWorkbook(partSchema(), { job: JOB, maxRows: 10 });
	const workbook = openWorkbook(file);

	assert.deepEqual(workbook.sheets.map((sheet) => [sheet.name, sheet.state]), [
		['Part', 'visible'], ['Part BOM', 'visible'], ['_lists', 'veryHidden'], ['_meta', 'veryHidden']
	]);

	const part = workbook.sheet('Part');
	const headers = [];
	for (let position = 1; position <= 9; position++) {
		headers.push(part.cell(1, position).value);
	}
	// Required and key columns carry an asterisk.
	assert.deepEqual(headers, [
		'Part Number *', 'Name', 'Type', 'Description', 'Unit', 'Make / Buy *', 'Effective Date', 'Weight', 'Assigned Creator'
	]);
	// The key column also has a fill of its own.
	assert.notEqual(part.cell('A1').fill, part.cell('B1').fill);
	assert.equal(part.cell('A1').bold, true);
	assert.match(part.cell('A1').note, /^Key column\./);
	assert.match(part.cell('F1').note, /^Required\./);

	// Frozen header.
	assert.deepEqual(part.frozen, { rows: 1, topLeftCell: 'A2' });

	// Protected without a password; the header is locked and every data cell is unlocked, also below the data.
	assert.equal(part.protection.sheet, true);
	assert.equal(part.protection.objects, true, 'the header notes are protected too');
	for (const action of ['insertRows', 'deleteRows', 'insertColumns', 'deleteColumns', 'sort', 'formatCells']) {
		assert.notEqual(part.protection[action], false, action + ' is not allowed');
	}
	assert.equal(part.protection.algorithmName, undefined);
	assert.equal(part.protection.hashValue, undefined);
	assert.equal(part.cell('A1').locked, true);
	for (const address of ['A2', 'I2', 'A11', 'A5000', 'I100000']) {
		assert.equal(part.cell(address).locked, false, address + ' is unlocked');
	}
	// A column that is not part of the template stays locked.
	assert.equal(part.cell('J2').locked, true);

	// The modified-on column is hidden and locked.
	assert.equal(part.column(12).hidden, true);
	assert.equal(part.cell(1, 12).value, 'modified_on');
	assert.equal(part.cell(2, 12).locked, true);

	const bom = workbook.sheet('Part BOM');
	assert.match(bom.cell('A1').note, /^Identifies the parent Part/);
	assert.match(bom.cell('B1').note, /^Identifies the related Part/);
	assert.match(bom.cell('C1').note, /match key/);
});

test('validation: text length for string and ml_string, none for text and formatted text', async () => {
	const { validation } = await openAllTypes();

	assert.deepEqual(
		[validation('code').type, validation('code').operator, validation('code').formulae],
		['textLength', 'lessThanOrEqual', [32]]
	);
	assert.deepEqual(
		[validation('title').type, validation('title').operator, validation('title').formulae],
		['textLength', 'lessThanOrEqual', [64]]
	);
	assert.equal(validation('notes'), undefined);
	assert.equal(validation('body'), undefined);
});

test('validation: whole number for integer, with the range and its ends', async () => {
	const { validation, letter } = await openAllTypes();

	assert.deepEqual(
		[validation('count_required').type, validation('count_required').operator, validation('count_required').formulae],
		['whole', 'between', [1, 999]]
	);
	// An optional column accepts #CLEAR as well. Its range excludes the ends: more than 0, less than 10.
	const optional = letter('count_optional') + '2';
	assert.equal(validation('count_optional').type, 'custom');
	assert.deepEqual(validation('count_optional').formulae, [
		`IF(ISNUMBER(${optional}),AND(${optional}=INT(${optional}),${optional}>=1,${optional}<=9),EXACT(${optional},"#CLEAR"))`
	]);
	// Without a range, the limits are those of a 32-bit integer.
	assert.match(validation('count_free').formulae[0], />=-2147483648,[A-Z]+2<=2147483647\)/);
});

test('validation: decimal for decimal and float, with minimum and maximum when set', async () => {
	const { validation, letter } = await openAllTypes();
	const summary = (id) => [validation(id).type, validation(id).operator, validation(id).formulae];

	assert.deepEqual(summary('price_required'), ['decimal', 'between', [0, 99999.99]]);
	assert.deepEqual(summary('ratio_min'), ['decimal', 'greaterThan', [0]]);
	assert.deepEqual(summary('ratio_max'), ['decimal', 'lessThanOrEqual', [100]]);
	// "Between" includes its ends, so a range that excludes both needs a formula.
	const exclusive = letter('price_exclusive') + '2';
	assert.deepEqual(summary('price_exclusive'), ['custom', undefined, [`AND(ISNUMBER(${exclusive}),${exclusive}>0,${exclusive}<1)`]]);
	assert.deepEqual(summary('ratio_any'), ['custom', undefined, [`ISNUMBER(${letter('ratio_any')}2)`]]);

	const optional = letter('price_optional') + '2';
	assert.deepEqual(summary('price_optional'), [
		'custom', undefined, [`IF(ISNUMBER(${optional}),AND(${optional}>=0,${optional}<=99999.99),EXACT(${optional},"#CLEAR"))`]
	]);
	const free = letter('ratio_optional') + '2';
	assert.deepEqual(summary('ratio_optional'), ['custom', undefined, [`OR(ISNUMBER(${free}),EXACT(${free},"#CLEAR"))`]]);
});

test('dates: date-only for short_date and long_date, date-time otherwise, with date validation', async () => {
	const { sheet, validation, letter } = await openAllTypes();

	assert.equal(sheet.cell(letter('start_date') + '2').numFmt, 'yyyy-mm-dd');
	assert.equal(sheet.cell(letter('end_date') + '2').numFmt, 'yyyy-mm-dd');
	assert.equal(sheet.cell(letter('released_on') + '2').numFmt, 'yyyy-mm-dd hh:mm:ss');
	assert.equal(sheet.cell(letter('checked_on') + '2').numFmt, 'yyyy-mm-dd hh:mm:ss');
	// The format also covers the rows below the data, where a user types new dates.
	assert.equal(sheet.cell(letter('start_date') + '4000').numFmt, 'yyyy-mm-dd');

	assert.equal(validation('start_date').type, 'date');
	assert.equal(validation('start_date').operator, 'between');
	assert.equal(validation('released_on').type, 'date');
	const optional = letter('end_date') + '2';
	assert.deepEqual(validation('end_date').formulae, [
		`IF(ISNUMBER(${optional}),AND(${optional}>=1,${optional}<2958466),EXACT(${optional},"#CLEAR"))`
	]);
});

test('dates are written and read through UTC, so the wall-clock value survives any time zone', async () => {
	const rows = [[{
		values: {
			code: 'A-1',
			start_date: '2026-03-01T00:00:00',
			end_date: '2026-03-29',
			released_on: '2026-12-31T23:59:59',
			checked_on: new Date(Date.UTC(2026, 5, 15, 8, 30, 0))
		},
		modifiedOn: '2026-10-31T08:15:42'
	}]];
	const { file, sheet, letter } = await openAllTypes({ rows });

	const start = sheet.cell(letter('start_date') + '2').value;
	assert.ok(start instanceof Date, 'a real Excel date cell');
	assert.equal(start.toISOString(), '2026-03-01T00:00:00.000Z');
	assert.equal(sheet.cell(letter('released_on') + '2').value.toISOString(), '2026-12-31T23:59:59.000Z');

	const parsed = await parseWorkbook(file);
	assert.deepEqual(parsed.sheets[0].rows, [{
		rowNumber: 2,
		cells: {
			code: 'A-1',
			start_date: '2026-03-01T00:00:00',
			end_date: '2026-03-29T00:00:00',
			released_on: '2026-12-31T23:59:59',
			checked_on: '2026-06-15T08:30:00'
		},
		modifiedOn: '2026-10-31T08:15:42'
	}]);
});

test('booleans: a Yes,No drop-down, read back as true and false', async () => {
	const rows = [[{ values: { code: 'A-1', is_active: '1', is_template: false }, modifiedOn: null }]];
	const { file, sheet, validation, letter } = await openAllTypes({ rows });

	assert.deepEqual([validation('is_active').type, validation('is_active').formulae], ['list', ['"Yes,No"']]);
	assert.deepEqual([validation('is_template').type, validation('is_template').formulae], ['list', ['"Yes,No,#CLEAR"']]);
	assert.equal(sheet.cell(letter('is_active') + '2').value, 'Yes');
	assert.equal(sheet.cell(letter('is_template') + '2').value, 'No');

	const edited = await editWorkbook(file, (book) => {
		book.set('Sample Item', 'A3', 'A-2');
		book.set('Sample Item', letter('is_active') + '3', 'no');
		book.set('Sample Item', letter('is_template') + '3', true);
		book.set('Sample Item', 'A4', 'A-3');
		book.set('Sample Item', letter('is_active') + '4', 'Maybe');
	});
	const cells = (await parseWorkbook(edited)).sheets[0].rows.map((row) => row.cells);
	assert.deepEqual(cells, [
		{ code: 'A-1', is_active: true, is_template: false },
		{ code: 'A-2', is_active: false, is_template: true },
		// Anything else is passed on as it is; the server reports it.
		{ code: 'A-3', is_active: 'Maybe' }
	]);
});

test('lists: a drop-down from a named range on _lists, never an inline list', async () => {
	const { workbook, validation } = await openAllTypes();
	const lists = workbook.sheet('_lists');
	const rangeOf = (name) => workbook.definedName(name);
	const column = (letter, from, to) => {
		const texts = [];
		for (let row = from; row <= to; row++) {
			texts.push(lists.cell(letter + row).value);
		}
		return texts;
	};

	// A required column: the plain range, which stops before the token.
	assert.deepEqual([validation('unit').type, validation('unit').formulae], ['list', ['mce_list_2']]);
	assert.equal(rangeOf('mce_list_2'), "'_lists'!$C$2:$C$8");
	// The label is shown. The value is shown when the label is empty ("M, long"), when two values share a
	// label (L and GAL, both "Volume"), and when a label equals a value that is shown (PK, labelled "Pack").
	assert.deepEqual(column('C', 2, 8), ['Each', 'Kilogram, net', 'M, long', 'L', 'GAL', 'PK', 'Pack']);
	// The column to the right maps each text back to its stored value.
	assert.deepEqual(column('D', 2, 8), ['EA', 'KG', 'M, long', 'L', 'GAL', 'PK', 'Pack']);

	// An optional column: the same list with #CLEAR as its last entry.
	assert.deepEqual(validation('unit_optional').formulae, ['mce_list_2_clear']);
	assert.equal(rangeOf('mce_list_2_clear'), "'_lists'!$C$2:$C$9");
	assert.equal(lists.cell('C9').value, '#CLEAR');

	// A filter list gets its whole list.
	assert.deepEqual(validation('region').formulae, ['mce_list_3_clear']);
	assert.deepEqual(column('E', 2, 4), ['North', 'South', '#CLEAR']);

	// A required list without values has nothing to offer, so it has no validation.
	assert.equal(validation('retired_choice'), undefined);
});

test('lists: the sheet shows the label, parseWorkbook returns the stored value', async () => {
	const rows = [[
		{ values: { code: 'A-1', unit: 'KG', unit_optional: 'GAL', region: 'south' }, modifiedOn: null },
		{ values: { code: 'A-2', unit: 'M, long', unit_optional: 'PK' }, modifiedOn: null },
		// A value that is no longer in the list is shown as it is stored.
		{ values: { code: 'A-3', unit: 'OLD' }, modifiedOn: null }
	]];
	const { file, sheet, letter } = await openAllTypes({ rows });

	// A list value that contains a comma is one entry, on the sheet and in the drop-down range.
	assert.equal(sheet.cell(letter('unit') + '2').value, 'Kilogram, net');
	assert.equal(sheet.cell(letter('unit_optional') + '2').value, 'GAL');
	assert.equal(sheet.cell(letter('region') + '2').value, 'South');
	assert.equal(sheet.cell(letter('unit') + '3').value, 'M, long');

	const edited = await editWorkbook(file, (book) => {
		book.set('Sample Item', 'A5', 'A-4');
		// Excel accepts typed text in any case when the drop-down comes from a range.
		book.set('Sample Item', letter('unit') + '5', 'kilogram, NET');
		book.set('Sample Item', letter('unit_optional') + '5', 'Pack');
	});
	const cells = (await parseWorkbook(edited)).sheets[0].rows.map((row) => row.cells);
	assert.deepEqual(cells, [
		{ code: 'A-1', unit: 'KG', unit_optional: 'GAL', region: 'south' },
		{ code: 'A-2', unit: 'M, long', unit_optional: 'PK' },
		{ code: 'A-3', unit: 'OLD' },
		{ code: 'A-4', unit: 'KG', unit_optional: 'Pack' }
	]);
});

test('mv_list: plain text without validation, with a header note', async () => {
	const rows = [[{ values: { code: 'A-1', tags: 'EA,KG' }, modifiedOn: null }]];
	const { file, sheet, validation, letter } = await openAllTypes({ rows });

	assert.equal(validation('tags'), undefined);
	assert.match(sheet.cell(letter('tags') + '1').note, /stored values of the list, separated by commas/);
	assert.equal((await parseWorkbook(file)).sheets[0].rows[0].cells.tags, 'EA,KG');
});

test('classification: a drop-down of the class paths, whose names may contain commas', async () => {
	const rows = [[{ values: { code: 'A-1', classification: 'Hardware/Bracket, Angle' }, modifiedOn: null }]];
	const { file, workbook, validation } = await openAllTypes({ rows });
	const lists = workbook.sheet('_lists');

	assert.deepEqual([validation('classification').type, validation('classification').formulae], ['list', ['mce_list_1_clear']]);
	assert.equal(workbook.definedName('mce_list_1_clear'), "'_lists'!$A$2:$A$6");
	assert.deepEqual([2, 3, 4, 5, 6].map((row) => lists.cell('A' + row).value), [...CLASS_PATHS, '#CLEAR']);

	const edited = await editWorkbook(file, (book) => {
		book.set('Sample Item', 'A3', 'A-2');
		book.set('Sample Item', 'F3', 'monthly, numbered day');
	});
	const cells = (await parseWorkbook(edited)).sheets[0].rows.map((row) => row.cells.classification);
	// A path typed in another case comes back as the class structure spells it.
	assert.deepEqual(cells, ['Hardware/Bracket, Angle', 'Monthly, Numbered Day']);
});

test('classification without a class structure is a plain text column', async () => {
	const schema = allTypesSchema();
	schema.sheets[0].classPaths = null;
	const workbook = openWorkbook(await buildWorkbook(schema, { maxRows: 10 }));
	const { validation } = workbook.sheet('Sample Item').cell('F2');

	assert.deepEqual([validation.type, validation.formulae], ['textLength', [512]]);
});

test('item references: free text with an input message and a header note', async () => {
	const { sheet, validation, letter } = await openAllTypes();
	const expected = /must match the keyed name of the Identity exactly\. Upper and lower case are not distinguished\./;

	assert.equal(validation('owner').type, 'any');
	assert.equal(validation('owner').showInputMessage, true);
	assert.match(validation('owner').prompt, expected);
	assert.match(sheet.cell(letter('owner') + '1').note, expected);
});

test('colours: #RRGGBB checked with a formula', async () => {
	const { validation, letter } = await openAllTypes();
	const required = letter('colour') + '2';
	const optional = letter('accent') + '2';
	const isColour = (cell) => `AND(LEN(${cell})=7,LEFT(${cell},1)="#",ISNUMBER(HEX2DEC(MID(${cell},2,6))))`;

	assert.deepEqual([validation('colour').type, validation('colour').formulae], ['custom', [isColour(required)]]);
	assert.deepEqual(validation('accent').formulae, [`OR(${isColour(optional)},EXACT(${optional},"#CLEAR"))`]);
});

test('class-specific columns are greyed out on rows of another class', async () => {
	const { sheet, letter } = await openAllTypes();
	const column = letter('thread_size');

	assert.equal(sheet.conditionalFormats.length, 1);
	const [formatting] = sheet.conditionalFormats;
	assert.equal(formatting.sqref, `${column}2:${column}11`);
	// Greyed when the classification is filled in and is neither the class nor one of its subclasses.
	assert.deepEqual(formatting.rules[0].formulae, ['AND($F2<>"",$F2<>"Hardware",LEFT($F2,9)<>"Hardware/")']);
	assert.match(formatting.rules[0].style.fillColor, /^FF[0-9A-F]{6}$/, 'a solid fill');
	assert.match(sheet.cell(column + '1').note, /Applies to class Hardware only/);
});

test('#CLEAR: accepted in every column that is not required, a key or a match key', async () => {
	const { schema, file, workbook, validation, letter } = await openAllTypes();
	const columns = schema.sheets[0].columns;
	const clearable = columns.filter((column) => !column.required && !column.isKey && !column.isMatchKey);
	const mandatory = columns.filter((column) => !clearable.includes(column));

	// The validation of a clearable column lets the token through, each kind in its own way.
	for (const column of clearable) {
		const entry = validation(column.id);
		const where = column.id + ' (' + column.dataType + ')';
		if (!entry || entry.type === 'any') {
			continue; // no validation at all: any text is accepted
		}
		if (entry.type === 'textLength') {
			assert.ok(entry.formulae[0] >= '#CLEAR'.length, where + ' is long enough for the token');
		} else if (entry.type === 'list') {
			assert.match(entry.formulae[0], /#CLEAR"$|_clear$/, where + ' offers the token');
		} else {
			assert.equal(entry.type, 'custom', where);
			assert.match(entry.formulae[0], /EXACT\([A-Z]+2,"#CLEAR"\)/, where + ' tests for the token');
		}
	}
	// A text column shorter than the token needs a formula to accept it.
	assert.deepEqual(validation('short_code').formulae, ['OR(LEN(C2)<=4,EXACT(C2,"#CLEAR"))']);
	// Every range that offers the token ends on it.
	const lists = workbook.sheet('_lists');
	for (const name of ['mce_list_1_clear', 'mce_list_2_clear', 'mce_list_3_clear']) {
		const lastCell = workbook.definedName(name).split(':')[1].replace(/\$/g, '');
		assert.equal(lists.cell(lastCell).value, '#CLEAR', name);
	}

	// Required and key columns keep plain validation, without the token.
	for (const column of mandatory) {
		const entry = validation(column.id);
		if (entry) {
			assert.doesNotMatch(JSON.stringify(entry), /#CLEAR/, column.id);
		}
	}
	assert.doesNotMatch(workbook.sheet('Sample Item').cell('A1').note, /#CLEAR/);

	// The token typed into every clearable column comes back as the token.
	const edited = await editWorkbook(file, (book) => {
		book.set('Sample Item', 'A2', 'A-1');
		for (const column of clearable) {
			book.set('Sample Item', letter(column.id) + '2', '#CLEAR');
		}
		book.set('Sample Item', 'A3', 'A-2');
		book.set('Sample Item', letter('unit_optional') + '3', '#clear');
	});
	const [first, second] = (await parseWorkbook(edited)).sheets[0].rows;
	assert.deepEqual(first.cells, Object.fromEntries([['code', 'A-1'], ...clearable.map((column) => [column.id, '#CLEAR'])]));
	assert.equal(second.cells.unit_optional, '#CLEAR');
});

test('one validation entry per column range', async () => {
	const { schema, file } = await openAllTypes();
	const entries = readParts(file)['xl/worksheets/sheet1.xml'].match(/<dataValidation [^>]*>/g);
	const withoutValidation = ['notes', 'body', 'tags', 'retired_choice'];

	assert.equal(entries.length, schema.sheets[0].columns.length - withoutValidation.length);
	for (const entry of entries) {
		assert.match(entry, / sqref="([A-Z]+)2:\1 ?11"/);
	}
});

test('maxRows: 5000 validated rows by default, more when the pre-filled rows need it', async () => {
	const sheetXml = async (options) => readParts(await buildWorkbook(partSchema(), options))['xl/worksheets/sheet1.xml'];

	assert.match(await sheetXml({}), / sqref="A2:A5001"/);
	assert.match(await sheetXml({ maxRows: 50 }), / sqref="A2:A51"/);
	const rows = [Array.from({ length: 5 }, (unused, index) => ({ values: { item_number: 'P-' + index }, modifiedOn: null }))];
	const xml = await sheetXml({ maxRows: 2, rows });
	assert.match(xml, / sqref="A2:A6"/);
	assert.match(xml, /<conditionalFormatting sqref="H2:H6">/);
});

test('round trip: pre-filled Part and BOM rows come back as they went in', async () => {
	const rows = [
		[
			{
				values: {
					item_number: 'P-0001', name: 'Bracket, left', classification: 'Component', description: 'Line 1\nLine 2 <b>&</b> "quoted"',
					unit: 'KG', make_buy: 'Buy', effective_date: '2026-03-01T00:00:00', weight: '1.25', owned_by_id: 'Innovator Admin'
				},
				modifiedOn: '2026-10-31T08:15:42'
			},
			{
				// Empty values stay blank cells, so an import leaves them unchanged.
				values: { item_number: '000123', name: '', classification: null, unit: 'IN', make_buy: 'Make', weight: 0 },
				modifiedOn: '2026-10-30T10:00:00'
			}
		],
		[
			{
				values: { 'parent.item_number': '000123', 'related.item_number': 'P-0001', sort_order: '10', quantity: '2.5', reference_designator: 'R1,R2' },
				modifiedOn: '2026-10-30T10:00:01'
			},
			{
				// More digits than a number cell can hold: kept as text, so nothing is rounded.
				values: { 'parent.item_number': '000123', 'related.item_number': 'P-0001', sort_order: 20, quantity: '12345678901234567890.12' },
				modifiedOn: '2026-10-30T10:00:02'
			}
		]
	];
	const file = await buildWorkbook(partSchema(), { job: JOB, rows });
	assert.ok(file instanceof ArrayBuffer);
	const parsed = await parseWorkbook(file);

	assert.deepEqual(parsed.sheets, [
		{
			sheetName: 'Part',
			columns: ['item_number', 'name', 'classification', 'description', 'unit', 'make_buy', 'effective_date', 'weight', 'owned_by_id'],
			rows: [
				{
					rowNumber: 2,
					cells: {
						item_number: 'P-0001', name: 'Bracket, left', classification: 'Component', description: 'Line 1\nLine 2 <b>&</b> "quoted"',
						unit: 'KG', make_buy: 'Buy', effective_date: '2026-03-01T00:00:00', weight: 1.25, owned_by_id: 'Innovator Admin'
					},
					modifiedOn: '2026-10-31T08:15:42'
				},
				{
					rowNumber: 3,
					cells: { item_number: '000123', unit: 'IN', make_buy: 'Make', weight: 0 },
					modifiedOn: '2026-10-30T10:00:00'
				}
			]
		},
		{
			sheetName: 'Part BOM',
			columns: ['parent.item_number', 'related.item_number', 'sort_order', 'quantity', 'reference_designator'],
			rows: [
				{
					rowNumber: 2,
					cells: { 'parent.item_number': '000123', 'related.item_number': 'P-0001', sort_order: 10, quantity: 2.5, reference_designator: 'R1,R2' },
					modifiedOn: '2026-10-30T10:00:01'
				},
				{
					rowNumber: 3,
					cells: { 'parent.item_number': '000123', 'related.item_number': 'P-0001', sort_order: 20, quantity: '12345678901234567890.12' },
					modifiedOn: '2026-10-30T10:00:02'
				}
			]
		}
	]);
});

test('every data cell is written and unlocked, in the pre-filled rows and in the empty rows below them', async () => {
	const rows = [[
		{ values: { item_number: 'P-0001', unit: 'EA' }, modifiedOn: '2026-10-31T08:15:42' },
		{ values: { item_number: 'P-0002' }, modifiedOn: '2026-10-31T08:15:43' }
	]];
	const parts = readParts(await buildWorkbook(partSchema(), { job: JOB, rows, maxRows: 10 }));
	const styles = parts['xl/styles.xml'].match(/<cellXfs[^>]*>(.*?)<\/cellXfs>/)[1].split('<xf ').slice(1);
	const unlocked = (index) => /<protection locked="0"\/>/.test(styles[index]);
	const sheetXml = parts['xl/worksheets/sheet1.xml'];
	// Rows 2 and 3 hold data; rows 4 to 11 are the empty rows a user fills in (maxRows 10).
	for (const rowNumber of [2, 3, 4, 11]) {
		const rowXml = new RegExp('<row r="' + rowNumber + '"[^>]*>(.*?)</row>').exec(sheetXml)[1];
		// Excel locks a cell it has to create when a value is typed, so each of the nine data cells must be there.
		for (const letter of 'ABCDEFGHI') {
			const cell = new RegExp('<c r="' + letter + rowNumber + '" s="(\\d+)"').exec(rowXml);
			assert.ok(cell, letter + rowNumber + ' is written');
			assert.ok(unlocked(Number(cell[1])), letter + rowNumber + ' is unlocked');
		}
	}
	// Each data cell keeps the number format of its column.
	const workbook = openWorkbook(await buildWorkbook(partSchema(), { job: JOB, rows, maxRows: 10 }));
	const part = workbook.sheet('Part');
	for (let position = 1; position <= 9; position++) {
		assert.equal(part.cell(3, position).numFmt, part.column(position).numFmt, 'column ' + position);
	}
});

test('_meta: read back exactly, with the job, the keys and the match keys', async () => {
	const schema = partSchema();
	const parsed = await parseWorkbook(await buildWorkbook(schema, { job: JOB }));
	const columnsOf = (sheet) => sheet.columns.map(({ id, role, property, dataType, isKey, isMatchKey }) =>
		({ id, role, property, dataType, isKey, isMatchKey }));

	assert.deepEqual(parsed.meta, {
		formatVersion: 1,
		database: 'SampleDatabase',
		generatedOn: '2026-11-05T09:30:00',
		language: 'en',
		definitionHash: schema.definitionHash,
		job: JOB,
		sheets: [
			{
				sheetName: 'Part',
				kind: 'main',
				itemType: { id: '4F1AC04A2B484F3ABA4E20DB63808A88', name: 'Part' },
				relationshipType: null,
				createOnly: false,
				parentSheet: null,
				relatedSheet: null,
				columns: columnsOf(schema.sheets[0]),
				keyColumns: ['item_number'],
				matchKeyColumns: []
			},
			{
				sheetName: 'Part BOM',
				kind: 'relationship',
				itemType: { id: '5E9C5A12CC58413A8670CF4003C57848', name: 'Part BOM' },
				relationshipType: {
					id: '159C6D88795B4A86864420863466F728',
					name: 'Part BOM',
					relatedItemType: { id: '4F1AC04A2B484F3ABA4E20DB63808A88', name: 'Part' }
				},
				createOnly: true,
				parentSheet: 0,
				relatedSheet: 0,
				columns: columnsOf(schema.sheets[1]),
				keyColumns: ['parent.item_number', 'related.item_number'],
				matchKeyColumns: ['parent.item_number', 'related.item_number', 'sort_order']
			}
		]
	});

	// A blank template has no job.
	assert.equal((await parseWorkbook(await buildWorkbook(schema))).meta.job, null);
});

test('_meta: a wide ItemType stays far below the 32,767 characters of a cell', async () => {
	const schema = partSchema();
	const [template] = schema.sheets[0].columns;
	const wide = Array.from({ length: 400 }, (unused, index) => ({
		...template,
		id: 'property_' + String(index).padStart(3, '0') + '_' + 'x'.repeat(100),
		property: 'property_' + String(index).padStart(3, '0') + '_' + 'x'.repeat(100),
		label: 'Property ' + index,
		isKey: index === 0,
		required: index === 0
	}));
	schema.sheets[0].columns = wide;
	schema.sheets.length = 1;
	const file = await buildWorkbook(schema, { maxRows: 10 });

	const workbook = openWorkbook(file);
	let longest = 0;
	for (const cell of workbook.sheet('_meta').cells()) {
		longest = Math.max(longest, String(cell.value).length);
	}
	assert.ok(longest < 200, 'the longest _meta cell has ' + longest + ' characters');

	const parsed = await parseWorkbook(file);
	assert.deepEqual(parsed.sheets[0].columns, wide.map((column) => column.id));
	assert.deepEqual(parsed.meta.sheets[0].columns.map((column) => column.property), wide.map((column) => column.property));
	assert.deepEqual(parsed.meta.sheets[0].keyColumns, [wide[0].id]);
});

test('parseWorkbook refuses a file without a valid _meta sheet', async () => {
	// An ordinary workbook that this tool did not make.
	const foreign = plainWorkbook([{ name: 'Part', rows: [['Part Number', 'Name']] }]);
	await assert.rejects(parseWorkbook(foreign), (error) => {
		assert.equal(error.code, 'MCE_META_MISSING');
		assert.match(error.message, /not generated by Mass Create & Edit/);
		return true;
	});

	// Not a workbook at all.
	await assert.rejects(parseWorkbook(new TextEncoder().encode('item_number,name\nP-1,Bolt\n').buffer), { code: 'MCE_NOT_A_WORKBOOK' });

	// A _meta sheet that is not ours.
	const impostor = plainWorkbook([{ name: '_meta', rows: [['something', 'else']] }]);
	await assert.rejects(parseWorkbook(impostor), { code: 'MCE_META_INVALID' });

	// A workbook from a newer format.
	const newer = await editWorkbook(await buildWorkbook(partSchema()), (book) => {
		book.set('_meta', 'B2', 3);
	});
	await assert.rejects(parseWorkbook(newer), (error) => {
		assert.equal(error.code, 'MCE_META_INVALID');
		assert.match(error.message, /format version 3/);
		return true;
	});

	// A _meta sheet whose column records were removed.
	const gutted = await editWorkbook(await buildWorkbook(partSchema()), (book) => {
		for (const { ref, value } of book.workbook.sheet('_meta').cells()) {
			if (/^A\d+$/.test(ref) && value === 'sheet') {
				book.set('_meta', ref, null);
			}
		}
	});
	await assert.rejects(parseWorkbook(gutted), { code: 'MCE_META_INVALID' });
});

test('parseWorkbook flattens what a user can put in a cell and skips empty rows', async () => {
	const rows = [[{ values: { item_number: 'P-0001', make_buy: 'Make' }, modifiedOn: '2026-10-31T08:15:42' }]];
	const edited = await editWorkbook(await buildWorkbook(partSchema(), { rows }), (book) => {
		// Row 3 is left empty. Row 4: rich text, a formula, a link, a typed date, a formula with a number.
		book.set('Part', 'A4', { richText: [{ text: 'P-' }, { text: '0004', bold: true }] });
		book.set('Part', 'B4', { formula: '"Na"&"me"', result: 'Name' });
		book.link('Part', 'I4', 'Innovator Admin', 'mailto:admin@example.com');
		book.set('Part', 'G4', new Date(Date.UTC(2027, 0, 15)));
		book.set('Part', 'H4', { formula: '1+1.5', result: 2.5 });
		// Row 5: a date that lost its date format arrives as a serial number; an error value; padded text.
		book.set('Part', 'A5', '  P-0005  ');
		book.set('Part', 'G5', 46082, { numFmt: 'General' });
		book.set('Part', 'B5', { error: '#N/A' });
		// Row 6 holds only spaces, row 7 only a formula without a result: both count as empty.
		book.set('Part', 'B6', '   ');
		book.set('Part', 'B7', { formula: 'A99' });
		// Row 20: a row far below the others keeps its Excel row number.
		book.set('Part', 'A20', 'P-0020');
		book.set('Part', 'H20', 0);
	});
	const parsed = await parseWorkbook(edited);

	assert.deepEqual(parsed.sheets[0].rows, [
		{ rowNumber: 2, cells: { item_number: 'P-0001', make_buy: 'Make' }, modifiedOn: '2026-10-31T08:15:42' },
		{
			rowNumber: 4,
			cells: { item_number: 'P-0004', name: 'Name', effective_date: '2027-01-15T00:00:00', weight: 2.5, owned_by_id: 'Innovator Admin' },
			modifiedOn: null
		},
		// Text is not trimmed here; the server decides what spaces mean.
		{ rowNumber: 5, cells: { item_number: '  P-0005  ', name: '#N/A', effective_date: '2026-03-01T00:00:00' }, modifiedOn: null },
		{ rowNumber: 20, cells: { item_number: 'P-0020', weight: 0 }, modifiedOn: null }
	]);
	assert.deepEqual(parsed.sheets[1].rows, []);
});

test('a sheet the user deleted comes back without rows and marked as missing', async () => {
	const edited = await editWorkbook(await buildWorkbook(partSchema()), (book) => {
		book.removeSheet('Part BOM');
	});
	const parsed = await parseWorkbook(edited);

	assert.equal(parsed.sheets[0].missing, undefined);
	assert.deepEqual(parsed.sheets[1], {
		sheetName: 'Part BOM',
		columns: ['parent.item_number', 'related.item_number', 'sort_order', 'quantity', 'reference_designator'],
		rows: [],
		missing: true
	});
});

test('a value too long for a cell is left blank, with a note that says so', async () => {
	const rows = [[{ values: { item_number: 'P-0001', description: 'x'.repeat(40000) }, modifiedOn: null }]];
	const file = await buildWorkbook(partSchema(), { rows, maxRows: 10 });
	const part = (openWorkbook(file)).sheet('Part');

	assert.equal(part.cell('D2').value, null);
	assert.match(part.cell('D2').note, /40000 characters, more than an Excel cell can hold/);
	assert.deepEqual((await parseWorkbook(file)).sheets[0].rows[0].cells, { item_number: 'P-0001' });
});

test('sheet names are made safe for Excel: 31 characters, no reserved characters, unique', async () => {
	const schema = partSchema();
	const relationship = schema.sheets[1];
	// A relationship sheet is named after its RelationshipType: its name, or its label when it has none.
	const withLabel = (name) => ({ ...relationship, relationshipType: { ...relationship.relationshipType, name, label: 'BOM' } });
	schema.sheets = [
		schema.sheets[0],
		withLabel('Part/BOM: [draft]*?'),
		withLabel('A relationship label that is far too long for a sheet'),
		withLabel('part'),
		withLabel('History'),
		withLabel("'Quoted'"),
		withLabel('_meta'),
		withLabel('')
	];
	const parsed = await parseWorkbook(await buildWorkbook(schema, { maxRows: 10 }));
	const names = parsed.sheets.map((sheet) => sheet.sheetName);

	assert.deepEqual(names, [
		'Part', 'Part BOM draft', 'A relationship label that is fa', 'part (2)', 'History_', 'Quoted', '_meta (2)', 'BOM'
	]);
	assert.ok(names.every((name) => name.length <= 31));
});

test('buildWorkbook refuses a schema it can not build from', async () => {
	const broken = (change) => {
		const schema = partSchema();
		change(schema);
		return buildWorkbook(schema);
	};

	await assert.rejects(broken((schema) => { schema.formatVersion = 2; }), { code: 'MCE_SCHEMA_INVALID' });
	await assert.rejects(broken((schema) => { schema.sheets = []; }), { code: 'MCE_SCHEMA_INVALID' });
	await assert.rejects(broken((schema) => { schema.sheets[0].columns[1].id = 'item_number'; }), /two columns with the id "item_number"/);
	await assert.rejects(broken((schema) => { schema.sheets[0].columns[1].dataType = 'image'; }), /unknown data type "image"/);
	await assert.rejects(broken((schema) => { schema.lists = {}; }), /refers to a list that is not in the schema/);
});

test('the styles part lists font properties in the order the Open XML schema requires', async () => {
	const order = ['b', 'i', 'strike', 'condense', 'extend', 'outline', 'shadow', 'u', 'vertAlign', 'sz', 'color', 'name', 'family', 'charset', 'scheme'];
	const fonts = readParts(await buildWorkbook(allTypesSchema(), { maxRows: 10 }))['xl/styles.xml'].match(/<font>.*?<\/font>/g);

	assert.ok(fonts.length >= 2);
	for (const font of fonts) {
		const ranks = [...font.matchAll(/<([a-zA-Z]+)[^<>]*\/>/g)].map((match) => order.indexOf(match[1]));
		assert.deepEqual(ranks, [...ranks].sort((a, b) => a - b), font);
	}
});

// ---------------------------------------------------------------------------
// Several targets
// ---------------------------------------------------------------------------

/** Part and CAD Document, with Part's BOM, a Part CAD sheet and a Files sheet under each. */
function twoTargetSchema() {
	const schema = partSchema();
	const [part, bom] = schema.sheets;
	const CAD = { ...part.itemType, id: 'C'.repeat(32), name: 'CAD', label: 'CAD Document' };
	const cad = { ...structuredClone(part), itemType: CAD };
	const relationship = (sheet, id, name, label, related, parentSheet, relatedSheet) => ({
		...structuredClone(sheet), relationshipType: { ...sheet.relationshipType, id, name, label, relatedItemType: related }, parentSheet, relatedSheet
	});
	schema.sheets = [
		part, cad,
		{ ...bom, parentSheet: 0, relatedSheet: 0 },
		relationship(bom, 'D'.repeat(32), 'Part CAD', 'CAD', CAD, 0, 1),
		relationship(bom, 'E'.repeat(32), 'Part Files', 'Files', null, 0, null),
		relationship(bom, 'F'.repeat(32), 'CAD Files', 'Files', null, 1, null)
	];
	return schema;
}

test('several targets: a main sheet each, and relationship sheets named after their RelationshipType', async () => {
	const workbook = openWorkbook(await buildWorkbook(twoTargetSchema(), { job: JOB, maxRows: 5 }));
	// Each main sheet is followed by its own relationship sheets, and a target's tabs share a colour.
	assert.deepEqual(workbook.sheets.filter((sheet) => sheet.state === 'visible').map((sheet) => [sheet.name, sheet.tabColor]), [
		['Part', 'FF4472C4'], ['Part BOM', 'FF4472C4'], ['Part CAD', 'FF4472C4'], ['Part Files', 'FF4472C4'],
		['CAD Document', 'FFED7D31'], ['CAD Files', 'FFED7D31']
	]);
	const partCad = workbook.sheet('Part CAD');
	assert.match(partCad.cell('A1').note, /^Identifies the parent Part: a row on sheet 'Part' or an existing item/);
	assert.match(partCad.cell('B1').note, /^Identifies the related CAD Document: an existing item, or a row on sheet 'CAD Document', a new one too/);
	assert.match(workbook.sheet('CAD Files').cell('A1').note, /a row on sheet 'CAD Document'/);
});

test('several targets: the meta names the main sheet and the related main sheet of each relationship sheet', async () => {
	const parsed = await parseWorkbook(await buildWorkbook(twoTargetSchema(), { job: JOB, maxRows: 5 }));
	assert.deepEqual(parsed.meta.sheets.map((sheet) => [sheet.kind, sheet.parentSheet, sheet.relatedSheet]), [
		['main', null, null], ['main', null, null],
		['relationship', 0, 0], ['relationship', 0, 1], ['relationship', 0, null], ['relationship', 1, null]
	]);
	// The sheets come back in the order of the schema, whatever the order of the tabs.
	assert.deepEqual(parsed.sheets.map((sheet) => sheet.sheetName), ['Part', 'CAD Document', 'Part BOM', 'Part CAD', 'Part Files', 'CAD Files']);
});

test('one target: a schema without parent and related sheets gets the notes of one main sheet, and a relationship sheet has its RelationshipType name', async () => {
	const workbook = openWorkbook(await buildWorkbook(partSchema(), { job: JOB, maxRows: 5 }));
	// One target: no tab colours.
	assert.deepEqual(workbook.sheets.filter((sheet) => sheet.state === 'visible').map((sheet) => [sheet.name, !!sheet.tabColor]),
		[['Part', false], ['Part BOM', false]]);
	const bom = workbook.sheet('Part BOM');
	assert.match(bom.cell('A1').note, /a row on the first sheet or an existing item/);
	// Part BOM points at Part, the main sheet: a new Part of the workbook can be named.
	assert.match(bom.cell('B1').note, /a row on sheet 'Part', a new one too/);
	const parsed = await parseWorkbook(await buildWorkbook(partSchema(), { job: JOB, maxRows: 5 }));
	assert.deepEqual(parsed.meta.sheets.map((sheet) => [sheet.parentSheet, sheet.relatedSheet]), [[null, null], [0, 0]]);
});
