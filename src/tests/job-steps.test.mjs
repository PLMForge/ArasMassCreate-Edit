/**
 * Tests for the Methods that run the steps of a job (FRG_MCE_ExportWorkbook,
 * FRG_MCE_ImportWorkbook) and for the guard that keeps everything else from changing
 * what they record (FRG_MCE_Guard). Run from the repository root with `node --test`.
 *
 * The Methods run on an Aras server; what is tested here is their source:
 *   - the steps and the guard agree on the identity the steps take on;
 *   - every property of the job that a step writes is one the guard keeps, so that it
 *     can't be changed by hand either;
 *   - the import Method's header documents the steps it answers and its fault code.
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const read = (name) => readFileSync(new URL('../methods/' + name + '.cs', import.meta.url), 'utf8');
const guard = read('FRG_MCE_Guard');
const steps = { FRG_MCE_ExportWorkbook: read('FRG_MCE_ExportWorkbook'), FRG_MCE_ImportWorkbook: read('FRG_MCE_ImportWorkbook') };
const jobItemType = readFileSync(new URL('../../1.0.0/Imports/FRG_MassCreateEdit/ItemType/FRG_MassCreateEdit.xml', import.meta.url), 'utf8');

/** The leading comment of a Method source. */
const headerOf = (source) => source.match(/^(?:\/\/.*\r?\n)+/)[0];

/** The job properties the guard keeps: the jobProperties list of its source. */
function guardedProperties() {
	const list = guard.match(/string\[\] jobProperties = new string\[\]\s*\{([^}]*)\}/);
	assert.ok(list, 'the guard lists the job properties it keeps');
	return Array.from(list[1].matchAll(/"([a-z_]+)"/g), (match) => match[1]);
}

/** The properties of the job ItemType, as the package defines them. */
const jobProperties = () => Array.from(jobItemType.matchAll(/<Item type="Property"[\s\S]*?<name>([a-z_]+)<\/name>\s*<\/Item>/g), (match) => match[1]);

test('steps and guard take on the same identity', () => {
	const ids = [guard, ...Object.values(steps)].map((source) => {
		const constant = source.match(/const string StepsIdentityId = "([0-9A-F]{32})";/);
		assert.ok(constant, 'every one of them names the identity by its id');
		return constant[1];
	});
	assert.equal(new Set(ids).size, 1);
});

test('every job property a step writes is kept by the guard', () => {
	const guarded = guardedProperties();
	// The job's state is kept by its life cycle.
	const open = new Set(['state']);
	const known = new Set([...jobProperties(), ...guarded]);
	for (const [name, source] of Object.entries(steps)) {
		const written = new Set(Array.from(source.matchAll(/<([a-z_]+)>/g), (match) => match[1]).filter((tag) => known.has(tag)));
		assert.ok(written.size > 0, name + ' writes job properties');
		for (const property of written) {
			assert.ok(guarded.includes(property) || open.has(property), name + ' writes ' + property + ', which the guard does not keep');
		}
	}
	assert.ok(guarded.includes('export_options'), 'the guard keeps export_options');
	assert.match(steps.FRG_MCE_ExportWorkbook, /<export_options>/, 'generating the workbook writes export_options');
	assert.doesNotMatch(steps.FRG_MCE_ImportWorkbook, /<export_options>/, 'importing does not');
	for (const counted of ['qty_total', 'qty_created', 'qty_updated', 'qty_unchanged', 'qty_skipped', 'qty_failed']) {
		assert.ok(guarded.includes(counted), 'the guard keeps ' + counted);
	}
});

test('import: the header documents the steps the Method answers and its fault code', () => {
	const source = steps.FRG_MCE_ImportWorkbook;
	const header = headerOf(source);
	const handled = Array.from(source.matchAll(/if \(step == "(\w+)"\)/g), (match) => match[1]);
	assert.deepEqual(handled, ['upload', 'begin', 'batch', 'files', 'complete']);
	for (const step of handled) {
		assert.match(header, new RegExp('step = ' + step + '\\b'), 'the header describes step ' + step);
		assert.match(header, new RegExp('^//\\s+' + step + ':', 'm'), 'the header gives the output of step ' + step);
	}
	const code = source.match(/const string RowFailedCode = "(\w+)";/)[1];
	assert.match(header, new RegExp('faultcode\\s+' + code));
});

test('export: begin sets or empties every job property the steps own', () => {
	// A Draft job is open to its creator, who may have set any of them by hand; nothing of
	// that may reach the workbook or its import.
	const source = steps.FRG_MCE_ExportWorkbook;
	const begin = source.slice(source.indexOf('string Begin()'), source.indexOf('string Complete()'));
	const userInput = new Set(['prefill_idlist', 'targets', 'export_options']);
	for (const property of guardedProperties().filter((name) => !userInput.has(name))) {
		assert.match(begin, new RegExp('<' + property + '[ >]|"' + property + '"'), 'begin does not write ' + property);
	}
});
