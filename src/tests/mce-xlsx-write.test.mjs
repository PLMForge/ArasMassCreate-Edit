/**
 * Tests for the workbook writer (mce-xlsx-write.js). Run from the repository
 * root with `node --test`; nothing has to be installed.
 *
 * The zip container is checked byte by byte, the XML parts by their text, and
 * every workbook is also opened with the test tools' own reader, which reads it
 * the way a spreadsheet program would and shares no code with the writer.
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import { crc32 as zlibCrc32, inflateRawSync } from 'node:zlib';

import { buildWorkbook, openWorkbook, parseWorkbook, readFixture, readParts } from './helpers/workbook-tools.mjs';

const writerUrl = new URL('../../1.0.0/Innovator/Client/customer/masscreateedit/mce-xlsx-write.js', import.meta.url);
const { columnName, crc32, dateToSerial, writeXlsx, zip } = await import(writerUrl.href);

const oneSheet = (sheet, more = {}) => ({ sheets: [{ name: 'Data', ...sheet }], ...more });
const cellsOf = (...values) => [{ index: 1, cells: values.map((value, at) => ({ column: at + 1, value })) }];

// ---------------------------------------------------------------------------
// The zip container
// ---------------------------------------------------------------------------

/** The entries of a zip file, read from its central directory and checked against the local headers. */
function readZip(bytes) {
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	const end = bytes.byteLength - 22;
	assert.equal(view.getUint32(end, true), 0x06054b50, 'the file ends with the end-of-central-directory record');
	const count = view.getUint16(end + 10, true);
	assert.equal(view.getUint16(end + 8, true), count);
	const centralSize = view.getUint32(end + 12, true);
	let offset = view.getUint32(end + 16, true);
	assert.equal(offset + centralSize, end, 'the central directory sits right before the end record');
	const entries = [];
	for (let index = 0; index < count; index++) {
		assert.equal(view.getUint32(offset, true), 0x02014b50);
		const flags = view.getUint16(offset + 8, true);
		const method = view.getUint16(offset + 10, true);
		const crc = view.getUint32(offset + 16, true);
		const compressedSize = view.getUint32(offset + 20, true);
		const size = view.getUint32(offset + 24, true);
		const nameLength = view.getUint16(offset + 28, true);
		const local = view.getUint32(offset + 42, true);
		const name = new TextDecoder().decode(bytes.subarray(offset + 46, offset + 46 + nameLength));
		assert.equal(view.getUint32(local, true), 0x04034b50, name + ': local header');
		assert.equal(view.getUint16(local + 8, true), method, name + ': same method in both headers');
		assert.equal(view.getUint32(local + 14, true), crc, name + ': same CRC in both headers');
		assert.equal(view.getUint32(local + 18, true), compressedSize);
		assert.equal(view.getUint32(local + 22, true), size);
		const start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
		const body = bytes.subarray(start, start + compressedSize);
		const data = method === 0 ? body : new Uint8Array(inflateRawSync(body));
		entries.push({ name, flags, method, crc, size, data });
		offset += 46 + nameLength + view.getUint16(offset + 30, true) + view.getUint16(offset + 32, true);
	}
	return entries;
}

test('crc32 is the CRC-32 zip records', () => {
	assert.equal(crc32(new TextEncoder().encode('123456789')), 0xcbf43926);
	assert.equal(crc32(new Uint8Array(0)), 0);
	const bytes = new Uint8Array(100000).map((unused, index) => (index * 7919) & 0xff);
	assert.equal(crc32(bytes), zlibCrc32(bytes));
});

test('zip: deflated and stored entries, with sizes, CRCs and UTF-8 names', async () => {
	const long = 'abc'.repeat(10000);
	const entries = [
		{ name: 'big.xml', data: long },
		{ name: 'tiny.txt', data: 'x' },
		{ name: 'naïve/é.bin', data: new Uint8Array([0, 1, 2, 255]) },
		{ name: 'empty', data: '' }
	];
	for (const compress of [true, false]) {
		const read = readZip(await zip(entries, { compress }));
		assert.deepEqual(read.map((entry) => entry.name), entries.map((entry) => entry.name));
		for (const [index, entry] of read.entries()) {
			const expected = typeof entries[index].data === 'string' ? new TextEncoder().encode(entries[index].data) : entries[index].data;
			assert.deepEqual(entry.data, expected, entry.name);
			assert.equal(entry.size, expected.length);
			assert.equal(entry.crc, zlibCrc32(expected));
			assert.equal(entry.flags & 0x0800, 0x0800, 'the name is flagged as UTF-8');
		}
		// A file that deflating would make larger is stored.
		assert.deepEqual(read.map((entry) => entry.method), compress ? [8, 0, 0, 0] : [0, 0, 0, 0]);
	}
});

test('zip: without CompressionStream every entry is stored', async () => {
	const saved = globalThis.CompressionStream;
	globalThis.CompressionStream = undefined;
	try {
		const read = readZip(await zip([{ name: 'a.xml', data: 'a'.repeat(5000) }]));
		assert.equal(read[0].method, 0);
		assert.equal(read[0].data.length, 5000);
	} finally {
		globalThis.CompressionStream = saved;
	}
});

test('the same workbook gives the same bytes', async () => {
	const workbook = () => oneSheet({ rows: cellsOf('a', 1, new Date(Date.UTC(2026, 0, 1))) }, { created: new Date(Date.UTC(2026, 0, 1)) });
	assert.deepEqual(new Uint8Array(await writeXlsx(workbook())), new Uint8Array(await writeXlsx(workbook())));
});

test('the package: content types, relationships and parts', async () => {
	const file = await writeXlsx({
		creator: 'Me & You',
		created: new Date(Date.UTC(2026, 10, 5, 9, 30)),
		modified: new Date(Date.UTC(2026, 10, 6)),
		sheets: [
			{ name: 'Data', rows: [{ index: 1, cells: [{ column: 1, value: 'x', note: 'A note' }] }] },
			{ name: 'Hidden', state: 'veryHidden', rows: cellsOf(1) }
		]
	});
	assert.ok(file instanceof ArrayBuffer);
	const parts = readParts(file);
	assert.deepEqual(Object.keys(parts).sort(), [
		'[Content_Types].xml', '_rels/.rels', 'docProps/app.xml', 'docProps/core.xml', 'xl/_rels/workbook.xml.rels',
		'xl/comments1.xml', 'xl/drawings/vmlDrawing1.vml', 'xl/sharedStrings.xml', 'xl/styles.xml', 'xl/workbook.xml',
		'xl/worksheets/_rels/sheet1.xml.rels', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml'
	]);
	const types = parts['[Content_Types].xml'];
	for (const part of ['/xl/workbook.xml', '/xl/styles.xml', '/xl/sharedStrings.xml', '/xl/worksheets/sheet1.xml', '/xl/worksheets/sheet2.xml', '/xl/comments1.xml', '/docProps/core.xml', '/docProps/app.xml']) {
		assert.match(types, new RegExp('<Override PartName="' + part.replace(/[.[\]]/g, '\\$&') + '"'), part);
	}
	assert.match(types, /<Default Extension="vml"/);
	assert.match(parts['xl/worksheets/_rels/sheet1.xml.rels'], /Target="\.\.\/comments1\.xml"/);
	assert.match(parts['xl/worksheets/_rels/sheet1.xml.rels'], /Id="rId2" Type="[^"]+\/vmlDrawing" Target="\.\.\/drawings\/vmlDrawing1\.vml"/);
	assert.match(parts['xl/worksheets/sheet1.xml'], /<legacyDrawing r:id="rId2"\/><\/worksheet>$/);
	assert.match(parts['docProps/core.xml'], /<dc:creator>Me &amp; You<\/dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">2026-11-05T09:30:00Z</);

	const workbook = openWorkbook(file);
	assert.equal(workbook.creator, 'Me & You');
	assert.equal(workbook.modified.toISOString(), '2026-11-06T00:00:00.000Z');
	assert.deepEqual(workbook.sheets.map((sheet) => [sheet.name, sheet.state]), [['Data', 'visible'], ['Hidden', 'veryHidden']]);
	assert.equal(workbook.sheet('Data').cell('A1').note, 'A note');
});

// ---------------------------------------------------------------------------
// Worksheets
// ---------------------------------------------------------------------------

test('the children of a worksheet come in the order the schema requires', async () => {
	const style = { numFmt: '@', locked: false };
	const xml = readParts(await writeXlsx(oneSheet({
		tabColor: 'FF4472C4',
		frozenRows: 1,
		columns: [{ index: 1, width: 20, style }],
		rows: [{ index: 1, height: 30, cells: [{ column: 1, value: 'Head', note: 'Note' }] }, { index: 2, cells: [{ column: 1, style }] }],
		protection: { sheet: true, objects: true, scenarios: true, formatColumns: false, formatRows: false },
		conditionalFormats: [{ sqref: 'A2:A9', rules: [{ type: 'expression', priority: 1, formulae: ['A2<>""'], style: { fill: { color: 'FFD9D9D9' } } }] }],
		dataValidations: [{ sqref: 'A2:A9', type: 'textLength', operator: 'lessThanOrEqual', formulae: [5], allowBlank: true }]
	})))['xl/worksheets/sheet1.xml'];
	const children = [...xml.replace(/^<\?xml[^>]*>\s*<worksheet[^>]*>/, '').matchAll(/<(\w+)[\s/>](?:[^<]*?\/>|.*?<\/\1>)/g)].map((match) => match[1]);
	assert.deepEqual(children, [
		'sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetProtection',
		'conditionalFormatting', 'dataValidations', 'pageMargins', 'legacyDrawing'
	]);
	assert.match(xml, /<dimension ref="A1:A2"\/>/);
	assert.match(xml, /<sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"\/>/);
	assert.match(xml, /<row r="1" spans="1:1" ht="30" customHeight="1">/);
});

test('rows and cells are written in ascending order, whatever order they are given in', async () => {
	const xml = readParts(await writeXlsx(oneSheet({
		rows: [
			{ index: 10, cells: [{ column: 28, value: 'AB10' }, { column: 2, value: 'B10' }] },
			{ index: 3, cells: [{ column: 1, value: 'A3' }] },
			{ index: 5, cells: [] }
		]
	})))['xl/worksheets/sheet1.xml'];
	assert.deepEqual([...xml.matchAll(/<c r="(\w+)"/g)].map((match) => match[1]), ['A3', 'B10', 'AB10']);
	assert.match(xml, /<dimension ref="A3:AB10"\/>/);
	assert.doesNotMatch(xml, /<row r="5"/, 'a row without cells or height is left out');
	assert.equal(columnName(1), 'A');
	assert.equal(columnName(26), 'Z');
	assert.equal(columnName(27), 'AA');
	assert.equal(columnName(16384), 'XFD');

	await assert.rejects(writeXlsx(oneSheet({ rows: [{ index: 1, cells: [{ column: 2, value: 1 }, { column: 2, value: 2 }] }] })), /two cells at row 1, column 2/);
	await assert.rejects(writeXlsx({ sheets: [{ name: 'A' }, { name: 'a' }] }), /needs a name of its own/);
});

test('cell values: text through shared strings, numbers, booleans and dates', async () => {
	const file = await writeXlsx(oneSheet({
		rows: [
			{ index: 1, cells: [{ column: 1, value: 'same' }, { column: 2, value: 'same' }, { column: 3, value: 42.5 }, { column: 4, value: -1.5e-7 }] },
			{ index: 2, cells: [{ column: 1, value: true }, { column: 2, value: false }, { column: 3, value: Number.NaN }, { column: 4, value: Infinity }] },
			{ index: 3, cells: [{ column: 1, value: new Date(Date.UTC(2026, 2, 1, 12)), style: { numFmt: 'yyyy-mm-dd hh:mm' } }, { column: 2, value: new Date(Number.NaN) }, { column: 3, value: '' }] }
		]
	}));
	const parts = readParts(file);
	assert.match(parts['xl/sharedStrings.xml'], /count="4" uniqueCount="3"/);
	assert.doesNotMatch(parts['xl/worksheets/sheet1.xml'], /<v>[^<]*(NaN|Infinity)/, 'no number cell is NaN or Infinity');
	assert.match(parts['xl/worksheets/sheet1.xml'], /<c r="A3" s="1"><v>46082\.5<\/v><\/c>/);
	assert.match(parts['xl/worksheets/sheet1.xml'], /<c r="B3"\/><c r="C3"\/>/, 'an invalid date and an empty text leave the cell empty');

	const sheet = openWorkbook(file).sheet('Data');
	assert.deepEqual([1, 2, 3, 4].map((column) => sheet.cell(1, column).value), ['same', 'same', 42.5, -1.5e-7]);
	assert.deepEqual([1, 2, 3, 4].map((column) => sheet.cell(2, column).value), [true, false, 'NaN', 'Infinity']);
	assert.equal(sheet.cell('A3').value.toISOString(), '2026-03-01T12:00:00.000Z');
	assert.equal(sheet.cell('A3').numFmt, 'yyyy-mm-dd hh:mm');
});

test('dates are serial numbers from the UTC fields, read back the same way', () => {
	assert.equal(dateToSerial(new Date(Date.UTC(1899, 11, 30))), 0);
	assert.equal(dateToSerial(new Date(Date.UTC(1900, 2, 1))), 61);
	assert.equal(dateToSerial(new Date(Date.UTC(2026, 2, 1))), 46082);
	assert.equal(dateToSerial(new Date(Date.UTC(2026, 2, 1, 18))), 46082.75);
	assert.equal(dateToSerial(new Date(Date.UTC(9999, 11, 31))), 2958465);
});

test('text is escaped for XML and for Excel, and comes back as it went in', async () => {
	const tricky = 'a & b < c > d "e" \'f\'\ttab\nline\r\ncrlf _x0041_ _X00e9_ \u0007bell\u0000 \uD800lone \u{1F600}';
	const kept = 'a & b < c > d "e" \'f\'\ttab\nline\r\ncrlf _x0041_ _X00e9_ bell �lone \u{1F600}';
	const message = 'Line one\nLine "two" & <three>\tend';
	const file = await writeXlsx(oneSheet({
		rows: [{ index: 1, cells: [{ column: 1, value: tricky, note: tricky }, { column: 2, value: '  padded  ' }] }],
		dataValidations: [{ sqref: 'A2:A3', type: 'any', showInputMessage: true, promptTitle: 'T', prompt: message }]
	}, { definedNames: [{ name: 'here', sheet: 'Data', range: '$A$1' }] }));
	const parts = readParts(file);
	const strings = parts['xl/sharedStrings.xml'];
	assert.ok(strings.includes('_x005F_x0041_ _x005F_X00e9_'), 'a literal _xHHHH_ is escaped');
	assert.ok(strings.includes('&amp; b &lt; c &gt; d &quot;e&quot;'));
	assert.ok(strings.includes('line&#13;\ncrlf'), 'a carriage return is kept as a reference');
	assert.ok(!/[\u0000\u0007]/.test(strings), 'characters XML does not allow are left out');
	assert.ok(strings.includes('<t xml:space="preserve">  padded  </t>'));
	assert.ok(parts['xl/worksheets/sheet1.xml'].includes('prompt="Line one&#10;Line &quot;two&quot; &amp; &lt;three&gt;&#9;end"'));

	const sheet = openWorkbook(file).sheet('Data');
	assert.equal(sheet.cell('A1').value, kept);
	assert.equal(sheet.cell('A1').note, kept);
	assert.equal(sheet.cell('B1').value, '  padded  ');
	assert.equal(sheet.cell('A2').validation.prompt, message);
});

test('a pre-filled text that looks like an Excel escape, or has a carriage return, comes back from the workbook unchanged', async () => {
	const description = 'Code _x0041_ stays\r\nas it is';
	const rows = [[{ values: { item_number: 'P-1', description }, modifiedOn: null }], []];
	const parsed = await parseWorkbook(await buildWorkbook(readFixture('part-with-bom.schema.json'), { rows, maxRows: 5 }));
	assert.equal(parsed.sheets[0].rows[0].cells.description, description);
});

test('sheet names are escaped, and quoted in a defined name', async () => {
	const file = await writeXlsx({
		sheets: [{ name: 'O\'Brien & <Co>', rows: cellsOf('x') }],
		definedNames: [{ name: 'mce_list_1', sheet: 'O\'Brien & <Co>', range: '$A$1:$A$1' }]
	});
	assert.ok(readParts(file)['xl/workbook.xml'].includes('<definedName name="mce_list_1">\'O\'\'Brien &amp; &lt;Co&gt;\'!$A$1:$A$1</definedName>'));
	const workbook = openWorkbook(file);
	assert.equal(workbook.sheets[0].name, 'O\'Brien & <Co>');
	assert.equal(workbook.definedName('mce_list_1'), '\'O\'\'Brien & <Co>\'!$A$1:$A$1');
});

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

test('styles: equal styles share an entry, number formats are numbered from 164, two fills are reserved', async () => {
	const header = () => ({ numFmt: '@', locked: true, font: { name: 'Calibri', size: 11, bold: true }, alignment: { vertical: 'top', wrapText: true }, fill: { color: 'FFDDEBF7' } });
	const file = await writeXlsx(oneSheet({
		columns: [{ index: 1, width: 12, style: { numFmt: 'yyyy-mm-dd', locked: false } }, { index: 2, width: 12, style: { numFmt: 'yyyy-mm-dd', locked: false } }],
		rows: [
			{ index: 1, cells: [{ column: 1, value: 'H1', style: header() }, { column: 2, value: 'H2', style: header() }, { column: 3, value: 'K', style: { ...header(), fill: { color: 'FFFFE699' } } }] },
			{ index: 2, cells: [
				{ column: 1, style: { numFmt: 'yyyy-mm-dd', locked: false } },
				{ column: 2, style: { numFmt: 'yyyy-mm-dd hh:mm:ss', locked: false } },
				{ column: 3, style: { numFmt: '@' } },
				{ column: 4, style: { numFmt: 'General', locked: false } },
				{ column: 5, style: { numFmt: 'yyyy-mm-dd', locked: false } }
			] }
		]
	}));
	const parts = readParts(file);
	const styles = parts['xl/styles.xml'];
	const xfs = styles.match(/<cellXfs count="(\d+)">(.*?)<\/cellXfs>/);
	const list = xfs[2].split('<xf ').slice(1);
	assert.equal(Number(xfs[1]), list.length, 'the count matches');
	assert.equal(list.length, 7, 'default, two headers, three formats unlocked, text');
	assert.match(styles, /<numFmts count="2"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"\/><numFmt numFmtId="165" formatCode="yyyy-mm-dd hh:mm:ss"\/><\/numFmts>/);
	assert.match(styles, /<fills count="4"><fill><patternFill patternType="none"\/><\/fill><fill><patternFill patternType="gray125"\/><\/fill>/);
	assert.match(styles, /<fonts count="2"><font><sz val="11"\/><name val="Calibri"\/><family val="2"\/><\/font><font><b\/><sz val="11"\/><name val="Calibri"\/><\/font><\/fonts>/);

	const sheetXml = parts['xl/worksheets/sheet1.xml'];
	const s = (ref) => Number(new RegExp('<c r="' + ref + '" s="(\\d+)"').exec(sheetXml)[1]);
	assert.equal(s('A1'), s('B1'), 'equal styles, given as different objects, share one entry');
	assert.notEqual(s('A1'), s('C1'));
	assert.equal(s('A2'), s('E2'));
	assert.match(list[s('A2')], /numFmtId="164".*<protection locked="0"\/>/);
	assert.match(list[s('C2')], /numFmtId="49"/);
	assert.match(list[s('D2')], /^numFmtId="0" .*applyProtection="1"><protection locked="0"\/>/);
	assert.match(list[s('A1')], /applyProtection="1"><alignment vertical="top" wrapText="1"\/><\/xf>/, 'a locked cell has no protection element');
	assert.match(sheetXml, /<cols><col min="1" max="2" width="12" style="\d+" customWidth="1"\/><\/cols>/, 'equal neighbouring columns share a <col>');

	const sheet = openWorkbook(file).sheet('Data');
	assert.equal(sheet.cell('A1').bold, true);
	assert.equal(sheet.cell('C1').fill, 'FFFFE699');
	assert.equal(sheet.cell('B2').numFmt, 'yyyy-mm-dd hh:mm:ss');
	assert.equal(sheet.cell('A2').locked, false);
	assert.equal(sheet.cell('A500').locked, false, 'the column style covers the rows below');
	assert.equal(sheet.cell('A1').locked, true);
});

// ---------------------------------------------------------------------------
// Validation, formatting, protection
// ---------------------------------------------------------------------------

test('data validation, conditional formatting and protection read back as written', async () => {
	const file = await writeXlsx(oneSheet({
		rows: cellsOf('x'),
		protection: { sheet: true, objects: true, scenarios: true, formatColumns: false, formatRows: false, insertRows: false },
		conditionalFormats: [{ sqref: 'B2:B9', rules: [{ type: 'expression', priority: 1, formulae: ['AND($A2<>"",$A2<>"x")'], style: { fill: { color: 'FFD9D9D9' }, font: { color: 'FF808080' } } }] }],
		dataValidations: [
			{ sqref: 'A2:A9', type: 'list', formulae: ['"Yes,No"'], allowBlank: true, showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Pick', error: 'Pick Yes or No.' },
			{ sqref: 'B2:B9', type: 'whole', operator: 'between', formulae: [1, 999], allowBlank: true },
			{ sqref: 'C2:C9', type: 'date', operator: 'between', formulae: [new Date(Date.UTC(1899, 11, 31)), new Date(Date.UTC(9999, 11, 31, 23, 59, 59))] },
			{ sqref: 'D2:D9', type: 'decimal', operator: 'greaterThan', formulae: ['1E-7'] },
			{ sqref: 'E2:E9', type: 'any', showInputMessage: true, promptTitle: 'Item', prompt: 'Type a name.' }
		]
	}));
	const xml = readParts(file)['xl/worksheets/sheet1.xml'];
	assert.match(xml, /<sheetProtection sheet="1" objects="1" scenarios="1" formatColumns="0" formatRows="0" insertRows="0"\/>/);
	assert.match(xml, /<dataValidations count="5">/);
	assert.match(xml, /<dataValidation type="date" sqref="C2:C9"><formula1>1<\/formula1><formula2>2958465\.999988426<\/formula2><\/dataValidation>/);
	assert.match(xml, /<dataValidation showInputMessage="1" promptTitle="Item" prompt="Type a name\." sqref="E2:E9"\/>/);
	assert.match(readParts(file)['xl/styles.xml'], /<dxfs count="1"><dxf><font><color rgb="FF808080"\/><\/font><fill><patternFill patternType="solid"><bgColor rgb="FFD9D9D9"\/><\/patternFill><\/fill><\/dxf><\/dxfs>/);

	const sheet = openWorkbook(file).sheet('Data');
	assert.deepEqual(sheet.cell('A5').validation, {
		type: 'list', formulae: ['"Yes,No"'], allowBlank: true, showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Pick', error: 'Pick Yes or No.'
	});
	assert.deepEqual([sheet.cell('B2').validation.operator, sheet.cell('B2').validation.formulae], ['between', [1, 999]]);
	assert.deepEqual([sheet.cell('D2').validation.operator, sheet.cell('D2').validation.formulae], ['greaterThan', [1e-7]]);
	assert.equal(sheet.cell('E9').validation.prompt, 'Type a name.');
	assert.equal(sheet.cell('F2').validation, undefined);
	assert.equal(sheet.conditionalFormats[0].sqref, 'B2:B9');
	assert.deepEqual(sheet.conditionalFormats[0].rules[0].style, { fontColor: 'FF808080', fillColor: 'FFD9D9D9' });
	assert.equal(sheet.protection.sheet, true);
	assert.equal(sheet.protection.objects, true, 'objects are protected');
	assert.equal(sheet.protection.formatRows, false, 'rows may be formatted');
	assert.equal(sheet.protection.insertRows, false, 'rows may be inserted');
	assert.notEqual(sheet.protection.deleteRows, false);

	await assert.rejects(writeXlsx(oneSheet({ protection: { sheet: true, everything: false } })), /no flag "everything"/);
});

test('notes: a comment and a hidden note box per note, anchored next to the cell', async () => {
	const file = await writeXlsx(oneSheet({ rows: [
		{ index: 1, cells: [{ column: 1, value: 'H', note: 'First\nnote' }] },
		{ index: 3, cells: [{ column: 4, note: 'On an empty cell', style: { numFmt: '@' } }] }
	] }));
	const parts = readParts(file);
	assert.match(parts['xl/comments1.xml'], /<comment ref="A1" authorId="0"><text><r><t xml:space="preserve">First\nnote<\/t><\/r><\/text><\/comment><comment ref="D3"/);
	const vml = parts['xl/drawings/vmlDrawing1.vml'];
	assert.deepEqual([...vml.matchAll(/<x:Anchor>([^<]*)<\/x:Anchor>/g)].map((match) => match[1]), ['1, 6, 0, 14, 3, 2, 4, 16', '4, 6, 1, 14, 6, 2, 5, 16']);
	assert.deepEqual([...vml.matchAll(/<x:Row>(\d+)<\/x:Row><x:Column>(\d+)<\/x:Column>/g)].map((match) => match[1] + ',' + match[2]), ['0,0', '2,3']);
	assert.deepEqual([...vml.matchAll(/<v:shape id="(\w+)"/g)].map((match) => match[1]), ['_x0000_s1025', '_x0000_s1026']);

	const sheet = openWorkbook(file).sheet('Data');
	assert.equal(sheet.cell('A1').note, 'First\nnote');
	assert.equal(sheet.cell('D3').note, 'On an empty cell');
	assert.equal(sheet.cell('D3').value, null);
});

test('workbook: the first visible sheet is the active one; hidden sheets, tab colours and frozen rows', async () => {
	const file = await writeXlsx({
		sheets: [
			{ name: 'Lists', state: 'veryHidden', rows: cellsOf('x') },
			{ name: 'Main', tabColor: 'FFED7D31', frozenRows: 2, rows: cellsOf('y') },
			{ name: 'Other', state: 'hidden' }
		]
	});
	const parts = readParts(file);
	assert.match(parts['xl/workbook.xml'], /<bookViews><workbookView activeTab="1"\/><\/bookViews>/);
	assert.match(parts['xl/worksheets/sheet2.xml'], /tabSelected="1"/);
	assert.doesNotMatch(parts['xl/worksheets/sheet1.xml'], /tabSelected/);
	assert.doesNotMatch(parts['[Content_Types].xml'], /sharedStrings\.xml"[^>]*>.*sharedStrings\.xml"/, 'listed once');
	const workbook = openWorkbook(file);
	assert.deepEqual(workbook.sheets.map((sheet) => sheet.state), ['veryHidden', 'visible', 'hidden']);
	const main = workbook.sheet('Main');
	assert.equal(main.tabColor, 'FFED7D31');
	assert.deepEqual(main.frozen, { rows: 2, topLeftCell: 'A3' });
});
