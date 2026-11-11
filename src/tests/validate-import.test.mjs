/**
 * Tests for the validation Method (FRG_MCE_ValidateImport). Run from the repository
 * root with `node --test`; nothing has to be installed.
 *
 * The Method runs on an Aras server, so what is tested here is what can be tested
 * without one:
 *   - the request: what the Method reads from an upload is what the workbook module
 *     really returns when it reads a workbook;
 *   - the result: a result recorded from the Method has the shape its header documents,
 *     and the checker used for that notices a result that has not;
 *   - the source: the issue codes it reports are the ones it documents, it asks for the
 *     template schema in one place, and it writes nothing.
 */

import test from 'node:test';
import assert from 'node:assert/strict';

import { buildWorkbook, parseWorkbook, readFixture } from './helpers/workbook-tools.mjs';
import { checkValidationResult, documentedCodes, readMethodSource, reportedCodes } from './helpers/validation-result.mjs';

const JOB = { id: '0123456789ABCDEF0123456789ABCDEF', token: 'c2FtcGxlLXRva2Vu' };
const sampleResult = () => readFixture('validate-import.result.json');

/** A Part with BOM workbook as generated for a job, with two parts and one BOM row, read back. */
async function sampleUpload() {
	const schema = readFixture('part-with-bom.schema.json');
	const file = await buildWorkbook(schema, {
		job: JOB,
		maxRows: 10,
		rows: [
			[
				{ values: { item_number: 'P-0001', name: 'Bracket', unit: 'KG', make_buy: 'Buy', effective_date: '2026-03-01T00:00:00' }, modifiedOn: '2026-10-31T08:15:42' },
				{ values: { item_number: 'P-0002', name: 'Plate' }, modifiedOn: null }
			],
			[{ values: { 'parent.item_number': 'P-0001', 'related.item_number': 'P-0002', sort_order: '10', quantity: '2' }, modifiedOn: '2026-10-31T08:15:42' }]
		]
	});
	return { schema, parsed: await parseWorkbook(file) };
}

test('request: an upload carries what the Method checks the job against', async () => {
	const { schema, parsed } = await sampleUpload();
	const meta = parsed.meta;
	assert.equal(meta.formatVersion, 1);
	assert.equal(meta.database, schema.database);
	assert.equal(meta.definitionHash, schema.definitionHash);
	assert.deepEqual(meta.job, JOB);
	assert.equal(meta.sheets.length, schema.sheets.length);
	assert.equal(parsed.sheets.length, schema.sheets.length, 'one entry of rows per sheet of the workbook');

	for (const [index, sheet] of schema.sheets.entries()) {
		const described = meta.sheets[index];
		assert.equal(described.kind, index === 0 ? 'main' : 'relationship', 'the main sheet comes first');
		assert.equal(described.itemType.id, sheet.itemType.id);
		assert.equal(described.relationshipType ? described.relationshipType.id : null, sheet.relationshipType ? sheet.relationshipType.id : null);
		assert.equal(typeof described.sheetName, 'string');
		assert.deepEqual(described.columns.map((column) => [column.id, column.dataType]), sheet.columns.map((column) => [column.id, column.dataType]),
			'every column comes with the data type it had when the workbook was generated');
		assert.deepEqual(described.keyColumns, sheet.columns.filter((column) => column.isKey).map((column) => column.id));
		assert.deepEqual(described.matchKeyColumns, sheet.columns.filter((column) => column.isMatchKey).map((column) => column.id));
	}
});

test('request: rows carry a row number, their filled cells by column id, and the export time', async () => {
	const { parsed } = await sampleUpload();
	const [main, bom] = parsed.sheets;
	assert.notEqual(main.missing, true);
	assert.deepEqual(main.rows.map((row) => row.rowNumber), [2, 3]);
	assert.deepEqual(main.rows[0].cells, { item_number: 'P-0001', name: 'Bracket', unit: 'KG', make_buy: 'Buy', effective_date: '2026-03-01T00:00:00' },
		'a list cell arrives as the stored value, a date as yyyy-MM-ddTHH:mm:ss, and a blank cell is left out');
	assert.equal(main.rows[0].modifiedOn, '2026-10-31T08:15:42', 'the export time arrives as Aras wrote it, so it can be compared as text');
	assert.equal(main.rows[1].modifiedOn, null);
	assert.deepEqual(bom.rows[0].cells, { 'parent.item_number': 'P-0001', 'related.item_number': 'P-0002', sort_order: 10, quantity: 2 },
		'a number cell arrives as a number');
});

test('result: the recorded result has the documented shape', () => {
	const result = sampleResult();
	assert.deepEqual(checkValidationResult(result), []);
	// The fixture shows every action and both severities, on both kinds of sheet.
	const rows = result.sheets.flatMap((sheet) => sheet.rows);
	assert.deepEqual(Array.from(new Set(rows.map((row) => row.action))).sort(), ['error', 'new', 'none', 'update']);
	assert.deepEqual(Array.from(new Set(rows.flatMap((row) => row.issues || []).map((issue) => issue.severity))).sort(), ['error', 'warning']);
	assert.deepEqual(result.sheets.map((sheet) => sheet.kind), ['main', 'relationship']);
});

test('result: an update row carries the changed cells only, old and new, stored and shown', () => {
	const update = sampleResult().sheets[0].rows.find((row) => row.action === 'update');
	assert.deepEqual(Object.keys(update.values).sort(), ['make_buy', 'name', 'owned_by_id']);
	assert.deepEqual(Object.keys(update.old).sort(), Object.keys(update.values).sort());
	assert.match(update.values.owned_by_id, /^[0-9A-F]{32}$/, 'an item reference is written as an id');
	assert.equal(update.display.owned_by_id, 'Administrators', 'and shown by its keyed name');
	assert.equal(update.display.name, undefined, 'a value that is shown as it is stored has no display text');
});

test('result: a relationship row names its parent, by sheet row when the parent is in the workbook', () => {
	const result = sampleResult();
	const [main, bom] = result.sheets;
	const newParentRow = main.rows.find((row) => row.action === 'new').row;
	const underNewParent = bom.rows.find((row) => row.parent && row.parent.row === newParentRow);
	assert.deepEqual(underNewParent.parent, { sheet: 0, row: newParentRow, id: null, keyedName: null }, 'a parent that is created by the same upload has no id yet');
	assert.equal(underNewParent.action, 'new');
	const underExisting = bom.rows.find((row) => row.action === 'update');
	assert.match(underExisting.parent.id, /^[0-9A-F]{32}$/);
	assert.ok(bom.parents.some((parent) => parent.id === underExisting.parent.id && parent.existing >= parent.untouched));
});

test('result: every row carries a digest of what importing it would do', () => {
	const result = sampleResult();
	const rows = result.sheets.flatMap((sheet) => sheet.rows);
	assert.ok(rows.every((row) => /^[0-9a-f]{32}$/.test(row.digest)));
	const writing = rows.filter((row) => row.action === 'new' || row.action === 'update');
	assert.equal(new Set(writing.map((row) => row.digest)).size, writing.length, 'rows that write different things have different digests');
	const errors = rows.filter((row) => row.action === 'error' && row.id === undefined && !row.parent);
	assert.ok(errors.length >= 2);
	assert.equal(new Set(errors.map((row) => row.digest)).size, 1, 'messages are not part of the digest: error rows without an item all have the same one');
});

test('result: the checker notices a result that is not as documented', () => {
	const broken = [
		(result) => { result.status = 'done'; },
		(result) => { result.totals.rows += 1; },
		(result) => { result.sheets[0].rows[0].action = 'skip'; },
		(result) => { result.sheets[0].rows.find((row) => row.action === 'error').action = 'new'; },
		(result) => { result.sheets[0].rows.find((row) => row.action === 'update').old = undefined; },
		(result) => { result.sheets[0].rows.find((row) => row.action === 'new').values.no_such_column = 'x'; },
		(result) => { result.sheets[0].rows.find((row) => row.issues).issues[0].code = 'NOT_DOCUMENTED'; },
		(result) => { result.sheets[1].rows.find((row) => row.action === 'new').parent = null; },
		(result) => { result.sheets[1].parents = undefined; },
		(result) => { delete result.sheets[0].rows[0].digest; },
		(result) => { result.sheets[1].rows[0].digest = 'not a digest'; },
		(result) => { result.status = 'rejected'; }
	];
	for (const [index, breakIt] of broken.entries()) {
		const result = sampleResult();
		breakIt(result);
		assert.notDeepEqual(checkValidationResult(JSON.parse(JSON.stringify(result))), [], 'change ' + (index + 1) + ' goes unnoticed');
	}
});

test('source: the issue codes the Method reports are the ones its header lists', () => {
	const documented = documentedCodes();
	assert.ok(documented.length > 40, 'the header lists the issue codes');
	assert.deepEqual(reportedCodes(), documented);
});

test('source: the template schema is asked for in one place', () => {
	const calls = readMethodSource().match(/applyMethod\("FRG_MCE_GetTemplateSchema"/g) || [];
	assert.equal(calls.length, 1);
});

test('source: the Method reads and never writes', () => {
	const source = readMethodSource();
	const actions = Array.from(new Set(Array.from(source.matchAll(/action=\\"(\w+)\\"/g), (match) => match[1]))).sort();
	assert.deepEqual(actions, ['get', 'getPermissions'], 'the only AML actions are reads');
	assert.doesNotMatch(source, /applySQL|\.apply\(\)|newItem\(/, 'no SQL and no items built for applying');
});
