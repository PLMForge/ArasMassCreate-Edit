/**
 * Test tools for the validation Method (FRG_MCE_ValidateImport): the issue codes its
 * source documents and reports, and a checker for the shape of its result.
 *
 * The Method runs on an Aras server, so these tools work on its source text and on
 * results that were recorded from it.
 */

import { readFileSync } from 'node:fs';

export const methodUrl = new URL('../../methods/FRG_MCE_ValidateImport.cs', import.meta.url);

export const readMethodSource = () => readFileSync(methodUrl, 'utf8').replace(/^﻿/, '');

const CODE = /[A-Z][A-Z_]{3,}/g;

/** The comment block at the top of the source. */
function headerOf(source) {
	const lines = source.split(/\r?\n/);
	const end = lines.findIndex((line) => !line.startsWith('//'));
	return lines.slice(0, end < 0 ? lines.length : end).join('\n');
}

/** The issue codes the header lists, under "Issue codes". */
export function documentedCodes(source = readMethodSource()) {
	const header = headerOf(source);
	const start = header.indexOf('// Issue codes.');
	if (start < 0) {
		return [];
	}
	// The list ends at the first comment line that is empty.
	const section = header.slice(start).split('\n').slice(1);
	const end = section.findIndex((line) => line.trim() === '//');
	const text = section.slice(0, end < 0 ? section.length : end).join('\n');
	return Array.from(new Set(text.match(CODE) || [])).sort();
}

/** The issue codes the code reports: every upper-case literal such as "TOO_LONG" in the code below the header. */
export function reportedCodes(source = readMethodSource()) {
	const body = source.slice(headerOf(source).length);
	const literals = body.match(/"[A-Z][A-Z_]{3,}"/g) || [];
	const notCodes = new Set(['TRUE', 'FALSE']);
	// Fault codes of the tool's Methods start with MCE_; they are not issue codes.
	return Array.from(new Set(literals.map((literal) => literal.slice(1, -1)).filter((code) => !notCodes.has(code) && !code.startsWith('MCE_')))).sort();
}

const ACTIONS = ['new', 'update', 'delete', 'none', 'error'];
const SEVERITIES = ['error', 'warning'];
const TOTALS = ['rows', 'new', 'update', 'none', 'delete', 'error', 'errors', 'warnings'];
const COLUMN_STATUSES = ['ok', 'unknown', 'typeChanged', 'added'];

const isObject = (value) => value !== null && typeof value === 'object' && !Array.isArray(value);
const isId = (value) => typeof value === 'string' && /^[0-9A-F]{32}$/.test(value);
const isCount = (value) => Number.isInteger(value) && value >= 0;

/**
 * Checks a result of the validation Method against the shape its header documents.
 *
 * @param {object} result The parsed JSON.
 * @param {string[]} [codes] The issue codes that may occur; default: the ones the header lists.
 * @returns {string[]} What is wrong; empty when the result is as documented.
 */
export function checkValidationResult(result, codes = documentedCodes()) {
	const problems = [];
	const expect = (condition, message) => {
		if (!condition) {
			problems.push(message);
		}
	};
	const checkIssue = (issue, where, isRowIssue) => {
		expect(isObject(issue), where + ': an issue is an object');
		if (!isObject(issue)) {
			return;
		}
		expect(SEVERITIES.includes(issue.severity), where + ': severity is error or warning');
		expect(codes.includes(issue.code), where + ': code ' + issue.code + ' is a documented code');
		expect(typeof issue.message === 'string' && issue.message.trim() !== '', where + ': an issue has a message');
		if (isRowIssue) {
			expect(issue.column === null || typeof issue.column === 'string', where + ': column is a column id or null');
		}
	};
	const checkTotals = (totals, where) => {
		expect(isObject(totals) && TOTALS.every((name) => isCount(totals[name])), where + ': totals has ' + TOTALS.join(', '));
		if (isObject(totals)) {
			expect(totals.rows === totals.new + totals.update + totals.delete + totals.none + totals.error, where + ': the actions add up to the rows');
		}
	};

	expect(isObject(result), 'the result is an object');
	if (!isObject(result)) {
		return problems;
	}
	expect(result.formatVersion === 1, 'formatVersion is 1');
	expect(['validated', 'rejected', 'failed'].includes(result.status), 'status is validated, rejected or failed');
	expect(result.job === null || (isObject(result.job) && isId(result.job.id) && typeof result.job.number === 'string' && typeof result.job.state === 'string'), 'job is null or { id, number, state }');
	expect(typeof result.database === 'string', 'database is text');
	expect(result.limits === null || (isObject(result.limits) && isCount(result.limits.maxRows) && isCount(result.limits.maxFileSizeMb)), 'limits is null or { maxRows, maxFileSizeMb, ... }');
	expect(typeof result.definitionChanged === 'boolean', 'definitionChanged is true or false');
	expect(Array.isArray(result.issues), 'issues is an array');
	expect(Array.isArray(result.sheets), 'sheets is an array');
	expect(isObject(result.stats) && isCount(result.stats.elapsedMs) && isCount(result.stats.requests), 'stats is { elapsedMs, requests }');
	checkTotals(result.totals, 'result');
	if (!Array.isArray(result.issues) || !Array.isArray(result.sheets) || !isObject(result.totals)) {
		return problems;
	}
	result.issues.forEach((issue, index) => checkIssue(issue, 'issues[' + index + ']', false));
	const uploadErrors = result.issues.filter((issue) => issue.severity === 'error').length;
	if (result.status === 'validated') {
		expect(uploadErrors === 0, 'a validated result has no upload errors');
		expect(result.sheets.length > 0 && result.sheets[0].kind === 'main', 'a validated result has its sheets, a main sheet first');
	} else {
		expect(uploadErrors > 0, 'a ' + result.status + ' result says why');
		expect(result.sheets.length === 0, 'a ' + result.status + ' result has no sheets');
	}

	const sums = Object.fromEntries(TOTALS.map((name) => [name, 0]));
	// The main sheets come first, one for each target; the relationship sheets follow.
	const mainCount = result.sheets.findIndex((sheet) => !isObject(sheet) || sheet.kind !== 'main') < 0
		? result.sheets.length : result.sheets.findIndex((sheet) => !isObject(sheet) || sheet.kind !== 'main');
	const isMainIndex = (index) => Number.isInteger(index) && index >= 0 && index < mainCount;
	result.sheets.forEach((sheet, sheetIndex) => {
		const where = 'sheets[' + sheetIndex + ']';
		expect(isObject(sheet), where + ' is an object');
		if (!isObject(sheet)) {
			return;
		}
		const isRelationship = sheet.kind === 'relationship';
		expect(sheet.index === sheetIndex, where + ': index is its place');
		expect(typeof sheet.sheetName === 'string', where + ': sheetName is text');
		expect(sheet.kind === (sheetIndex < mainCount ? 'main' : 'relationship'), where + ': the main sheets come first, the relationship sheets after them');
		const parentSheet = sheet.parentSheet;
		const relatedSheet = sheet.relatedSheet;
		expect(isRelationship ? isMainIndex(parentSheet) && (relatedSheet === null || isMainIndex(relatedSheet))
			: sheet.parentSheet === undefined && sheet.relatedSheet === undefined, where + ': parentSheet and relatedSheet name main sheets, on a relationship sheet');
		expect(isObject(sheet.itemType) && isId(sheet.itemType.id) && typeof sheet.itemType.name === 'string', where + ': itemType is { id, name, ... }');
		expect(isRelationship ? isObject(sheet.relationshipType) && isId(sheet.relationshipType.id) : sheet.relationshipType === null, where + ': relationshipType fits the kind');
		expect(typeof sheet.createOnly === 'boolean', where + ': createOnly is true or false');
		expect(Array.isArray(sheet.keyColumns) && Array.isArray(sheet.matchKeyColumns), where + ': keyColumns and matchKeyColumns are arrays');
		expect(Array.isArray(sheet.columns) && sheet.columns.every((column) =>
			isObject(column) && typeof column.id === 'string' && typeof column.dataType === 'string' && COLUMN_STATUSES.includes(column.status)), where + ': columns are { id, property, label, dataType, pattern, status }');
		checkTotals(sheet.totals, where);
		expect(Array.isArray(sheet.rows), where + ': rows is an array');
		expect(isRelationship ? Array.isArray(sheet.parents) : sheet.parents === undefined, where + ': parents belongs to a relationship sheet');
		if (!Array.isArray(sheet.rows) || !isObject(sheet.totals) || !Array.isArray(sheet.columns)) {
			return;
		}
		const columnIds = new Set(sheet.columns.map((column) => column.id));
		const counted = Object.fromEntries(TOTALS.map((name) => [name, 0]));
		counted.rows = sheet.rows.length;
		sheet.rows.forEach((row) => {
			const at = where + ' row ' + (isObject(row) ? row.row : '?');
			expect(isObject(row) && Number.isInteger(row.row) && row.row >= 2, at + ': row is the number of an Excel row below the header');
			if (!isObject(row)) {
				return;
			}
			expect(ACTIONS.includes(row.action), at + ': action is new, update, delete, none or error');
			expect(typeof row.digest === 'string' && /^[0-9a-f]{32}$/.test(row.digest), at + ': digest is 32 hexadecimal digits');
			counted[row.action] = (counted[row.action] || 0) + 1;
			const issues = row.issues || [];
			expect(row.issues === undefined || (Array.isArray(row.issues) && row.issues.length > 0), at + ': issues is left out or has entries');
			issues.forEach((issue) => checkIssue(issue, at, true));
			const errors = issues.filter((issue) => issue.severity === 'error').length;
			counted.errors += errors;
			counted.warnings += issues.length - errors;
			expect((row.action === 'error') === (errors > 0), at + ': the action is error exactly when the row has an error');
			expect((row.id === undefined) === (row.keyedName === undefined) && (row.id === undefined || isId(row.id)), at + ': id and keyedName come together');
			expect(row.action !== 'update' && row.action !== 'none' || isId(row.id), at + ': an update or unchanged row names its item');
			expect(row.action !== 'new' || row.id === undefined, at + ': a new row names no item');
			const writes = row.action === 'new' || row.action === 'update';
			expect(writes ? isObject(row.values) : row.values === undefined, at + ': values comes with new and update rows only');
			expect(row.action === 'update' ? isObject(row.old) : row.old === undefined && row.oldDisplay === undefined, at + ': old comes with update rows only');
			if (isObject(row.values)) {
				const names = Object.keys(row.values);
				expect(names.every((name) => typeof row.values[name] === 'string'), at + ': values are stored values, as text');
				expect(names.every((name) => columnIds.has(name) || (isRelationship && name === 'related_id')), at + ': values are named by column id');
				expect(row.action !== 'update' || names.length > 0, at + ': an update changes something');
				expect(!isObject(row.old) || names.every((name) => typeof row.old[name] === 'string') && Object.keys(row.old).length === names.length, at + ': old has the same cells as values');
				for (const part of ['display', 'oldDisplay']) {
					expect(row[part] === undefined || (isObject(row[part]) && Object.keys(row[part]).length > 0 && Object.keys(row[part]).every((name) => names.includes(name))),
						at + ': ' + part + ' is left out or names cells of values');
				}
			}
			if (isRelationship) {
				expect(row.parent === null || (isObject(row.parent) && (row.parent.row === null || Number.isInteger(row.parent.row)) && (row.parent.id === null || isId(row.parent.id)) &&
					row.parent.sheet === parentSheet),
					at + ': parent is null or { sheet, row, id, keyedName }, on the main sheet of the parents');
				const fileToken = isObject(row.related) && typeof row.related.file === 'string' && /^file:[0-9A-F]{32}:.+$/.test(row.related.file);
				expect(row.related === null || (isObject(row.related) && (isId(row.related.id) || (row.related.id === null && (Number.isInteger(row.related.row) || fileToken)))),
					at + ': related is null, { id, keyedName }, { row, id: null, keyedName } for an item a main-sheet row creates, or { id: null, keyedName, file } for a picked file');
				const createdOn = isObject(row.related) ? row.related.sheet : null;
				expect(!isObject(row.related) || row.related.id !== null || fileToken || (createdOn === relatedSheet && isMainIndex(createdOn) &&
					Array.isArray(result.sheets[createdOn].rows) && result.sheets[createdOn].rows.some((main) => main.row === row.related.row && main.action === 'new')),
					at + ': a related item without an id is created by a new row of the main sheet of the related ItemType');
				expect(!writes || (isObject(row.parent) && (row.parent.row !== null || row.parent.id !== null)), at + ': a row that is written knows its parent');
			} else {
				expect(row.parent === undefined && row.related === undefined, at + ': a main row has no parent or related');
			}
		});
		for (const name of TOTALS) {
			expect(sheet.totals[name] === counted[name], where + ': totals.' + name + ' is ' + counted[name]);
			sums[name] += sheet.totals[name];
		}
		for (const parent of sheet.parents || []) {
			expect(isObject(parent) && (parent.row === null || Number.isInteger(parent.row)) && (parent.id === null || isId(parent.id)) && isCount(parent.existing) && isCount(parent.untouched) &&
				parent.untouched <= parent.existing, where + ': a parent is { sheet, row, id, keyedName, existing, untouched }');
		}
	});
	sums.errors += uploadErrors;
	sums.warnings += result.issues.length - uploadErrors;
	for (const name of TOTALS) {
		expect(result.totals[name] === sums[name], 'totals.' + name + ' is ' + sums[name]);
	}
	return problems;
}
