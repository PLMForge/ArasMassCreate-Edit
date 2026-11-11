/**
 * Keeps the Method sources and the package in step. Each file in src/methods/
 * is the source of one server Method, and the package ships that Method as an
 * item exported from Innovator: 1.0.0/Imports/FRG_MassCreateEdit/Method/<name>.xml,
 * with the code in its method_code property. These tests fail when the two
 * differ, so a change to one side can not be released without the other.
 *
 * What is compared, and what is evened out first:
 *
 * - Line endings. The export tool writes item files with CRLF, the source files
 *   use LF, and a checkout may convert either. CRLF and a lone CR both count
 *   as LF, on both sides.
 * - Escaping. The export tool puts the code in a CDATA section, where nothing
 *   is escaped. Code containing "]]>" is split over several sections, which are
 *   joined again here. Code exported as plain element text is read too, and its
 *   character references (&lt; &gt; &amp; &quot; &apos; and numeric ones) are
 *   decoded.
 * - A byte order mark at the start of either file is ignored.
 *
 * Everything else must be equal character for character, trailing spaces and
 * the final line break included.
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, readdirSync, readFileSync } from 'node:fs';

const sourceFolder = new URL('../methods/', import.meta.url);
const exportFolder = new URL('../../1.0.0/Imports/FRG_MassCreateEdit/Method/', import.meta.url);

const sources = readdirSync(sourceFolder).filter((name) => /\.(cs|vb|js)$/i.test(name)).sort();

const withoutBom = (text) => text.replace(/^﻿/, '');
const withLf = (text) => text.replace(/\r\n?/g, '\n');

const NAMED_REFERENCES = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };

/** Decodes the character references of XML text. */
function decodeReferences(text) {
	return text.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-z]+);/g, (reference, body) => {
		if (body.startsWith('#x')) {
			return String.fromCodePoint(parseInt(body.slice(2), 16));
		}
		if (body.startsWith('#')) {
			return String.fromCodePoint(parseInt(body.slice(1), 10));
		}
		assert.ok(body in NAMED_REFERENCES, 'unknown character reference ' + reference);
		return NAMED_REFERENCES[body];
	});
}

/** Where the tag starts in the XML, leaving out what is inside a CDATA section (the code itself holds tags). */
function tagPositions(itemXml, tag) {
	const positions = [];
	let position = 0;
	while (position < itemXml.length) {
		const section = itemXml.indexOf('<![CDATA[', position);
		const markupEnd = section === -1 ? itemXml.length : section;
		for (let at = itemXml.indexOf(tag, position); at !== -1 && at < markupEnd; at = itemXml.indexOf(tag, at + 1)) {
			positions.push(at);
		}
		if (section === -1) {
			break;
		}
		const sectionEnd = itemXml.indexOf(']]>', section);
		assert.notEqual(sectionEnd, -1, 'every CDATA section is closed');
		position = sectionEnd + 3;
	}
	return positions;
}

/** The value of the one <name> element in an exported item: CDATA sections and escaped text, joined. */
function propertyOf(itemXml, name) {
	const open = '<' + name + '>';
	const close = '</' + name + '>';
	const starts = tagPositions(itemXml, open);
	assert.equal(starts.length, 1, 'the item has exactly one ' + name + ' property');

	let value = '';
	let position = starts[0] + open.length;
	for (;;) {
		if (itemXml.startsWith('<![CDATA[', position)) {
			const end = itemXml.indexOf(']]>', position);
			assert.notEqual(end, -1, 'every CDATA section of ' + name + ' is closed');
			value += itemXml.slice(position + 9, end);
			position = end + 3;
		} else if (itemXml.startsWith(close, position)) {
			return value;
		} else {
			const next = itemXml.indexOf('<', position);
			assert.notEqual(next, -1, name + ' is closed');
			value += decodeReferences(itemXml.slice(position, next));
			position = next;
		}
	}
}

/** Where two texts first differ, as a line number and the two lines. */
function firstDifference(expected, actual) {
	const expectedLines = expected.split('\n');
	const actualLines = actual.split('\n');
	const count = Math.max(expectedLines.length, actualLines.length);
	for (let index = 0; index < count; index++) {
		if (expectedLines[index] !== actualLines[index]) {
			return 'line ' + (index + 1) + '\n  source:   ' + JSON.stringify(expectedLines[index]) +
				'\n  exported: ' + JSON.stringify(actualLines[index]);
		}
	}
	return 'no difference';
}

test('src/methods holds at least one Method source', () => {
	assert.ok(sources.length > 0);
});

for (const fileName of sources) {
	const methodName = fileName.replace(/\.[^.]+$/, '');

	test(methodName + ': the source equals the method_code of the exported Method item', () => {
		const exportUrl = new URL(encodeURIComponent(methodName) + '.xml', exportFolder);
		assert.ok(existsSync(exportUrl), 'the package holds Method/' + methodName + '.xml');

		const itemXml = withoutBom(readFileSync(exportUrl, 'utf8'));
		assert.equal(propertyOf(itemXml, 'name'), methodName, 'the item is the Method the file is named after');

		const source = withLf(withoutBom(readFileSync(new URL(fileName, sourceFolder), 'utf8')));
		const exported = withLf(propertyOf(itemXml, 'method_code'));
		assert.ok(
			source === exported,
			'src/methods/' + fileName + ' differs from the exported method_code, first at ' + firstDifference(source, exported) +
			'\nSave the source into the Method item and export it again, or bring the source in line with the item.'
		);
	});
}
