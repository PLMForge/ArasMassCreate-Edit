/**
 * How the workbook module loads in a page: with the writer and reader beside
 * it, and nothing else. This test has a file of its own because it changes
 * globals, and `node --test` runs each file in a separate process.
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const moduleUrl = new URL('../../1.0.0/Innovator/Client/customer/masscreateedit/mce-workbook.js', import.meta.url);
const schema = JSON.parse(readFileSync(new URL('./fixtures/part-with-bom.schema.json', import.meta.url), 'utf8'));

test('the module works beside an AMD loader without registering with it and without creating globals', async () => {
	// A page may have an AMD loader. The module and its writer and reader must neither use it nor leave globals.
	const registered = [];
	globalThis.define = Object.assign((...parameters) => registered.push(parameters), { amd: {} });
	try {
		const { buildWorkbook, parseWorkbook } = await import(moduleUrl.href);
		const file = await buildWorkbook(schema, {
			rows: [[{ values: { item_number: 'P-0001', make_buy: 'Buy' }, modifiedOn: '2026-10-31T08:15:42' }]]
		});
		const parsed = await parseWorkbook(file);

		assert.deepEqual(parsed.sheets[0].rows, [
			{ rowNumber: 2, cells: { item_number: 'P-0001', make_buy: 'Buy' }, modifiedOn: '2026-10-31T08:15:42' }
		]);
		assert.equal(registered.length, 0, 'nothing was handed to the AMD loader');
		for (const name of ['module', 'exports']) {
			assert.equal(typeof globalThis[name], 'undefined', 'no global named ' + name);
		}
	} finally {
		delete globalThis.define;
	}
});

test('the module runs as a page would load it: over HTTP, with fetch, without anything from node', () => {
	const output = execFileSync(
		process.execPath,
		['--experimental-vm-modules', '--no-warnings', fileURLToPath(new URL('./helpers/browser-sandbox.mjs', import.meta.url))],
		{ encoding: 'utf8' }
	);
	const result = JSON.parse(output.trim().split('\n').pop());

	// Nothing is fetched as text and run: the writer and reader are imported as modules.
	assert.deepEqual(result.fetched, []);
	assert.equal(result.registeredWithAmd, 0);
	assert.deepEqual(result.leakedGlobals, []);
	assert.equal(result.isArrayBuffer, true);
	assert.ok(result.bytes > 5000);
	assert.equal(result.jobId, '0123456789ABCDEF0123456789ABCDEF');
	assert.deepEqual(result.rows, [{
		rowNumber: 2,
		cells: { item_number: 'P-0001', unit: 'KG', make_buy: 'Buy', effective_date: '2026-03-01T00:00:00' },
		modifiedOn: '2026-10-31T08:15:42'
	}]);
});

test('the package ships no third-party library', () => {
	assert.equal(existsSync(new URL('vendor/', moduleUrl)), false, 'the client folder has no vendor folder');
});
