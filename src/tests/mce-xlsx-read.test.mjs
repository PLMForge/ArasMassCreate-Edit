/**
 * Tests for the workbook reader (mce-xlsx-read.js). Every file here is made
 * by hand in the test: a small zip of the XML parts, written the way one
 * spreadsheet program or another writes them.
 */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import { crc32, deflateRawSync } from 'node:zlib';

const { decodeEntities, isDateFormat, openXlsx, readZip, scanXml, unescapeExcel } =
	await import(new URL('../../1.0.0/Innovator/Client/customer/masscreateedit/mce-xlsx-read.js', import.meta.url).href);
const { buildWorkbook, editWorkbook, parseWorkbook, readFixture } = await import('./helpers/workbook-tools.mjs');

// ---------------------------------------------------------------------------
// Making files
// ---------------------------------------------------------------------------

/**
 * A zip file of the given entries. An entry is { name, data } and may say
 * how it is stored: method (8 deflate, 0 stored, or any other number),
 * flags, a size to claim in place of the real one, and descriptor (sizes
 * left out of the local header, as a streaming writer does).
 */
function zip(entries, { comment = '' } = {}) {
	const locals = [];
	const central = [];
	let offset = 0;
	for (const entry of entries) {
		const data = typeof entry.data === 'string' ? Buffer.from(entry.data, 'utf8') : Buffer.from(entry.data);
		const method = entry.method === undefined ? 8 : entry.method;
		const stored = entry.stored || (method === 8 ? deflateRawSync(data) : data);
		const name = Buffer.from(entry.name, 'utf8');
		const size = entry.claimedSize === undefined ? data.length : entry.claimedSize;
		const crc = crc32(data) >>> 0;
		const local = Buffer.alloc(30);
		local.writeUInt32LE(0x04034b50, 0);
		local.writeUInt16LE(20, 4);
		local.writeUInt16LE((entry.flags || 0) | (entry.descriptor ? 8 : 0), 6);
		local.writeUInt16LE(method, 8);
		local.writeUInt32LE(entry.descriptor ? 0 : crc, 14);
		local.writeUInt32LE(entry.descriptor ? 0 : stored.length, 18);
		local.writeUInt32LE(entry.descriptor ? 0 : size, 22);
		local.writeUInt16LE(name.length, 26);
		const header = Buffer.alloc(46);
		header.writeUInt32LE(0x02014b50, 0);
		header.writeUInt16LE(20, 4);
		header.writeUInt16LE(20, 6);
		header.writeUInt16LE((entry.flags || 0) | (entry.descriptor ? 8 : 0), 8);
		header.writeUInt16LE(method, 10);
		header.writeUInt32LE(crc, 16);
		header.writeUInt32LE(entry.claimedCompressedSize === undefined ? stored.length : entry.claimedCompressedSize, 20);
		header.writeUInt32LE(size, 24);
		header.writeUInt16LE(name.length, 28);
		header.writeUInt32LE(offset, 42);
		const descriptor = entry.descriptor ? Buffer.alloc(16) : Buffer.alloc(0);
		if (entry.descriptor) {
			descriptor.writeUInt32LE(0x08074b50, 0);
			descriptor.writeUInt32LE(crc, 4);
			descriptor.writeUInt32LE(stored.length, 8);
			descriptor.writeUInt32LE(size, 12);
		}
		locals.push(local, name, stored, descriptor);
		central.push(header, name);
		offset += local.length + name.length + stored.length + descriptor.length;
	}
	const directory = Buffer.concat(central);
	const commentBytes = Buffer.from(comment, 'utf8');
	const end = Buffer.alloc(22);
	end.writeUInt32LE(0x06054b50, 0);
	end.writeUInt16LE(entries.length, 8);
	end.writeUInt16LE(entries.length, 10);
	end.writeUInt32LE(directory.length, 12);
	end.writeUInt32LE(offset, 16);
	end.writeUInt16LE(commentBytes.length, 20);
	const file = Buffer.concat([...locals, directory, end, commentBytes]);
	return file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength);
}

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const RELATIONSHIPS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE = 'http://schemas.openxmlformats.org/package/2006/relationships';

const relationshipsXml = (list) => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
	`<Relationships xmlns="${PACKAGE}">` +
	list.map(([id, type, target]) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`).join('') +
	'</Relationships>';

const worksheet = (sheetData, after = '') => `<worksheet xmlns="${MAIN}"><sheetData>${sheetData}</sheetData>${after}</worksheet>`;

/**
 * The parts of a workbook, by name. `sheets` is a list of { name, xml },
 * each stored as xl/worksheets/sheetN.xml unless it names its own part.
 */
function workbookParts({ sheets, sharedStrings, styles, workbookPr = '', namespaces = {} } = {}) {
	const ns = { main: MAIN, relationships: RELATIONSHIPS, ...namespaces };
	const parts = {
		'[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
		'_rels/.rels': relationshipsXml([['rId1', ns.relationships + '/officeDocument', 'xl/workbook.xml']])
	};
	const relationships = [];
	const sheetElements = sheets.map((sheet, index) => {
		const id = 'rId' + (index + 1);
		const part = sheet.part || 'worksheets/sheet' + (index + 1) + '.xml';
		relationships.push([id, ns.relationships + '/' + (sheet.type || 'worksheet'), sheet.target || part]);
		if (sheet.xml !== undefined) {
			parts[sheet.zipName || 'xl/' + part] = sheet.xml;
		}
		return `<sheet name="${sheet.name}" sheetId="${index + 1}" r:id="${id}"/>`;
	});
	if (sharedStrings !== undefined) {
		relationships.push(['rIdS', ns.relationships + '/sharedStrings', 'sharedStrings.xml']);
		parts['xl/sharedStrings.xml'] = `<sst xmlns="${ns.main}">${sharedStrings}</sst>`;
	}
	if (styles !== undefined) {
		relationships.push(['rIdT', ns.relationships + '/styles', 'styles.xml']);
		parts['xl/styles.xml'] = `<styleSheet xmlns="${ns.main}">${styles}</styleSheet>`;
	}
	parts['xl/workbook.xml'] = `<workbook xmlns="${ns.main}" xmlns:r="${ns.relationships}">${workbookPr}<sheets>${sheetElements.join('')}</sheets></workbook>`;
	parts['xl/_rels/workbook.xml.rels'] = relationshipsXml(relationships);
	return parts;
}

const zipParts = (parts) => zip(Object.entries(parts).map(([name, data]) => ({ name, data })));
const xlsx = (options) => zipParts(workbookParts(options));

/** The values of a sheet as { "A1": value }, Dates as ISO text. */
async function cellsOf(file, sheetName = 'Data') {
	const sheet = await (await openXlsx(file)).readSheet(sheetName);
	const cells = {};
	for (const row of sheet.rowNumbers()) {
		for (let column = 1; column <= 26; column++) {
			const value = sheet.value(row, column);
			if (value !== null) {
				cells[String.fromCharCode(64 + column) + row] = value instanceof Date ? value.toISOString() : value;
			}
		}
	}
	return cells;
}

const rejectsWith = (promise, code) => assert.rejects(promise, (error) => {
	assert.equal(error.code, code, error.message);
	return true;
});

// ---------------------------------------------------------------------------
// The zip container
// ---------------------------------------------------------------------------

test('zip: deflated and stored parts, an archive comment, sizes only in the table of contents', async () => {
	const parts = workbookParts({ sheets: [{ name: 'Data', xml: worksheet('<row r="1"><c r="A1"><v>1</v></c></row>') }] });
	const entries = Object.entries(parts).map(([name, data], index) => ({ name, data, method: index % 2 ? 0 : 8, descriptor: index % 3 === 0 }));
	const file = zip(entries, { comment: 'PK\u0005\u0006 looks like the end, but is part of the comment' });
	assert.deepEqual(await cellsOf(file), { A1: 1 });
});

test('zip: part names are matched without regard to case or a leading slash', async () => {
	const file = xlsx({
		sheets: [{ name: 'Data', zipName: '/XL/Worksheets/SHEET1.XML', xml: worksheet('<row r="1"><c r="A1"><v>2</v></c></row>') }]
	});
	assert.deepEqual(await cellsOf(file), { A1: 2 });
	const zipped = readZip(file);
	assert.ok(zipped.entries.has('xl/worksheets/sheet1.xml'));
});

test('zip: a file that is not a zip, or is cut off, is refused', async () => {
	await rejectsWith(openXlsx(new TextEncoder().encode('item_number,name\nP-1,Bolt\n').buffer), 'XLSX_NOT_ZIP');
	await rejectsWith(openXlsx(new ArrayBuffer(0)), 'XLSX_NOT_ZIP');
	await rejectsWith(openXlsx('not bytes'), 'XLSX_NOT_ZIP');
	const file = new Uint8Array(xlsx({ sheets: [{ name: 'Data', xml: worksheet('') }] }));
	// The end record is still there, the parts it points to are not.
	const cut = new Uint8Array(file.length - 60);
	cut.set(file.subarray(0, 40));
	cut.set(file.subarray(file.length - (cut.length - 40)), 40);
	await assert.rejects(openXlsx(cut), (error) => ['XLSX_DAMAGED', 'XLSX_NOT_ZIP'].includes(error.code));
});

test('zip: an encrypted workbook (an OLE compound file) and an encrypted zip entry are refused', async () => {
	const compound = new Uint8Array(1024);
	compound.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
	await assert.rejects(openXlsx(compound.buffer), (error) => {
		assert.equal(error.code, 'XLSX_ENCRYPTED');
		assert.match(error.message, /password/);
		return true;
	});
	const parts = workbookParts({ sheets: [{ name: 'Data', xml: worksheet('') }] });
	const encrypted = zip(Object.entries(parts).map(([name, data]) => ({ name, data, flags: name.endsWith('workbook.xml') ? 1 : 0 })));
	await rejectsWith(openXlsx(encrypted), 'XLSX_ENCRYPTED');

	// parseWorkbook says what to do about it.
	await assert.rejects(parseWorkbook(compound.buffer), (error) => {
		assert.equal(error.code, 'MCE_NOT_A_WORKBOOK');
		assert.match(error.message, /password-protected/);
		return true;
	});
});

test('zip: ZIP64 and packing methods other than deflate and store are refused', async () => {
	const parts = workbookParts({ sheets: [{ name: 'Data', xml: worksheet('') }] });
	const entries = Object.entries(parts).map(([name, data]) => ({ name, data }));
	await rejectsWith(openXlsx(zip(entries.map((entry, index) => (index === 0 ? { ...entry, claimedSize: 0xffffffff } : entry)))), 'XLSX_UNSUPPORTED');
	// The end record says to look for the ZIP64 one.
	const end = new Uint8Array(zip(entries));
	new DataView(end.buffer).setUint32(end.length - 6, 0xffffffff, true);
	await rejectsWith(openXlsx(end), 'XLSX_UNSUPPORTED');
	const bzip = zip(entries.map((entry) => (entry.name === 'xl/workbook.xml' ? { ...entry, method: 12, stored: Buffer.from('BZh9') } : entry)));
	await rejectsWith(openXlsx(bzip), 'XLSX_UNSUPPORTED');
});

test('zip: a part whose packed data is damaged is refused when it is read', async () => {
	const parts = workbookParts({ sheets: [{ name: 'Data', xml: worksheet('<row r="1"><c r="A1"><v>1</v></c></row>') }] });
	const entries = Object.entries(parts).map(([name, data]) => ({ name, data }));
	const damaged = zip(entries.map((entry) => (entry.name === 'xl/worksheets/sheet1.xml'
		? { ...entry, stored: Buffer.from([0xff, 0xff, 0xff, 0xff, 0x00, 0x13, 0x37]) }
		: entry)));
	// The list of sheets does not need the sheet: it is only unpacked when it is read.
	const workbook = await openXlsx(damaged);
	assert.equal(workbook.hasSheet('Data'), true);
	await rejectsWith(workbook.readSheet('Data'), 'XLSX_DAMAGED');
});

test('zip: a part that unpacks to more than the limit is stopped (a zip bomb)', async () => {
	const zeros = new Uint8Array(8 * 1024 * 1024);
	const bomb = (claimedSize) => zip([
		...Object.entries(workbookParts({ sheets: [{ name: 'Data' }] })).map(([name, data]) => ({ name, data })),
		{ name: 'xl/worksheets/sheet1.xml', data: zeros, claimedSize }
	]);
	const limit = { maxUncompressed: 1024 * 1024 };
	// The size it claims is already too much: refused without unpacking.
	await rejectsWith((await openXlsx(bomb(), limit)).readSheet('Data'), 'XLSX_TOO_LARGE');
	// It claims to be small: unpacking stops at the limit.
	await rejectsWith((await openXlsx(bomb(100), limit)).readSheet('Data'), 'XLSX_TOO_LARGE');
	// The real limit, against a part that claims 4 GB.
	await rejectsWith((await openXlsx(bomb(0xfffffffe))).readSheet('Data'), 'XLSX_TOO_LARGE');
});

test('zip: the bytes may come from another window, or be a view into a larger buffer', async () => {
	const file = new Uint8Array(xlsx({ sheets: [{ name: 'Data', xml: worksheet('<row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row>') }] }));
	const otherWindow = vm.runInNewContext('(length) => new ArrayBuffer(length)')(file.length);
	assert.equal(otherWindow instanceof ArrayBuffer, false);
	new Uint8Array(otherWindow).set(file);
	assert.deepEqual(await cellsOf(otherWindow), { A1: 'x' });
	const OtherUint8Array = vm.runInNewContext('Uint8Array');
	assert.deepEqual(await cellsOf(new OtherUint8Array(otherWindow)), { A1: 'x' });

	const padded = new Uint8Array(file.length + 100);
	padded.set(file, 37);
	assert.deepEqual(await cellsOf(padded.subarray(37, 37 + file.length)), { A1: 'x' });
});

// ---------------------------------------------------------------------------
// XML
// ---------------------------------------------------------------------------

test('xml: quoted ">" in attributes, comments, CDATA, a byte order mark, entities in attributes', () => {
	const seen = [];
	const handler = {
		collecting: true,
		open: (name, attributes, selfClosing) => seen.push(['open', name, { ...attributes }, selfClosing]),
		close: (name) => seen.push(['close', name]),
		text: (text) => seen.push(['text', text])
	};
	scanXml('\uFEFF<?xml version="1.0"?><!DOCTYPE x [<!ENTITY y "z">]><!-- <a> -->' +
		'<x:a xmlns:x="urn:x" b="1 > 0" c=\'it&apos;s &amp; &#x1F600;\' r:id="r1" id="plain"><![CDATA[<raw> &amp;]]>&lt;&#65;<x:e/></x:a>', handler);
	assert.deepEqual(seen, [
		['open', 'a', { b: '1 > 0', c: 'it\'s & \u{1F600}', id: 'r1' }, false],
		['text', '<raw> &amp;'],
		['text', '<A'],
		['open', 'e', {}, true],
		['close', 'e'],
		['close', 'a']
	]);
	assert.throws(() => scanXml('<a b="never closed>', handler), { code: 'XLSX_DAMAGED' });
});

test('xml: entities and Excel\'s _xHHHH_ escapes', () => {
	assert.equal(decodeEntities('&lt;&gt;&amp;&quot;&apos;&#65;&#x42;&#x1f600;&unknown;&#x110000;'), '<>&"\'AB\u{1F600}&unknown;&#x110000;');
	assert.equal(unescapeExcel('a_x000D_b_x000a_c'), 'a\rb\nc');
	// _x005F_ is the escape of "_": what follows it is text, not an escape.
	assert.equal(unescapeExcel('_x005F_x0041_'), '_x0041_');
	assert.equal(unescapeExcel('_x0041'), '_x0041');
});

// ---------------------------------------------------------------------------
// Cells
// ---------------------------------------------------------------------------

test('cells: every type, and a formula as the result Excel saved with it', async () => {
	const file = xlsx({
		sharedStrings: '<si><t>shared</t></si>',
		sheets: [{
			name: 'Data',
			xml: worksheet('<row r="1">' +
				'<c r="A1" t="s"><v>0</v></c>' +
				'<c r="B1" t="str"><v>text &amp; more</v></c>' +
				'<c r="C1" t="inlineStr"><is><r><t>in</t></r><r><rPr><b/></rPr><t xml:space="preserve">line </t></r></is></c>' +
				'<c r="D1" t="b"><v>1</v></c>' +
				'<c r="E1" t="b"><v>0</v></c>' +
				'<c r="F1" t="e"><v>#N/A</v></c>' +
				'<c r="G1" t="n"><v>1.5</v></c>' +
				'<c r="H1"><v>-4.25E-3</v></c>' +
				'<c r="I1" t="d"><v>2026-03-01T10:20:30</v></c>' +
				'<c r="J1"><f>1-1</f><v>0</v></c>' +
				'<c r="K1" t="str"><f>"a"&amp;"b"</f><v>ab</v></c>' +
				'<c r="L1"><f t="shared" ref="L1:M1" si="0">A1*2</f><v>7</v></c>' +
				'<c r="M1"><f t="shared" si="0"/><v>8</v></c>' +
				'<c r="N1"><f>A99</f></c>' +
				'<c r="O1" s="0"/>' +
				'<c r="P1" t="s"><v>99</v></c>' +
				'<c r="Q1" t="e"><f>1/0</f><v>#DIV/0!</v></c>' +
				'<c r="R1" t="b"><f>TRUE()</f><v>1</v></c>' +
				'<c r="S1" t="d"><v>2026-03-01</v></c>' +
				'</row>')
		}]
	});
	assert.deepEqual(await cellsOf(file), {
		A1: 'shared', B1: 'text & more', C1: 'inline ', D1: true, E1: false, F1: '#N/A', G1: 1.5, H1: -0.00425,
		I1: '2026-03-01T10:20:30.000Z', J1: 0, K1: 'ab', L1: 7, M1: 8, Q1: '#DIV/0!', R1: true, S1: '2026-03-01T00:00:00.000Z'
	});
});

test('cells: shared strings in runs, without phonetic guides, empty, with spaces, entities and escapes', async () => {
	const file = xlsx({
		sharedStrings:
			'<si><t>plain</t></si>' +
			'<si><r><rPr><b/></rPr><t xml:space="preserve">Bracket, </t></r><r><t>left</t></r></si>' +
			'<si><t>東京</t><rPh sb="0" eb="2"><t>トウキョウ</t></rPh><phoneticPr fontId="1"/></si>' +
			'<si/>' +
			'<si><t/></si>' +
			'<si><t xml:space="preserve">  padded  </t></si>' +
			'<si><t>&lt;b&gt; &amp; &#x1F600;</t></si>' +
			'<si><t>Line 1_x000D_\nLine 2 _x005F_x0041_</t></si>',
		sheets: [{
			name: 'Data',
			xml: worksheet('<row r="1">' + [0, 1, 2, 3, 4, 5, 6, 7].map((index) =>
				`<c r="${String.fromCharCode(65 + index)}1" t="s"><v>${index}</v></c>`).join('') + '</row>')
		}]
	});
	assert.deepEqual(await cellsOf(file), {
		A1: 'plain', B1: 'Bracket, left', C1: '東京', D1: '', E1: '', F1: '  padded  ', G1: '<b> & \u{1F600}', H1: 'Line 1\r\nLine 2 _x0041_'
	});
});

test('cells: rows and cells without an r attribute are counted', async () => {
	const file = xlsx({
		sheets: [{
			name: 'Data',
			xml: worksheet('<row><c><v>1</v></c><c><v>2</v></c></row>' +
				'<row r="5"><c><v>3</v></c><c r="D5"><v>4</v></c><c><v>5</v></c></row>' +
				'<row><c t="inlineStr"><is><t>6</t></is></c></row><row r="9"/>')
		}]
	});
	const sheet = await (await openXlsx(file)).readSheet('Data');
	assert.deepEqual(await cellsOf(file), { A1: 1, B1: 2, A5: 3, D5: 4, E5: 5, A6: '6' });
	assert.deepEqual(sheet.rowNumbers(), [1, 5, 6]);
	assert.equal(sheet.lastRow, 9);
	assert.equal(sheet.value(100, 100), null);
});

test('cells: every cell of a merged range shows the value of its first cell', async () => {
	const file = xlsx({
		sheets: [{
			name: 'Data',
			xml: worksheet('<row r="2"><c r="B2"><v>1</v></c><c r="C2"><v>9</v></c></row>', '<mergeCells count="1"><mergeCell ref="B2:C3"/></mergeCells>')
		}]
	});
	assert.deepEqual(await cellsOf(file), { B2: 1, C2: 1, B3: 1, C3: 1 });
});

test('cells: a number is a date when its cell format shows one, in the 1900 or the 1904 date system', async () => {
	const formats = ['0', '14', '164', '165', '166', '167', '168', '169', '170', '22', '45', '2', '31'];
	const styles = '<numFmts count="7">' +
		'<numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd"/>' +
		'<numFmt numFmtId="165" formatCode="[h]:mm:ss"/>' +
		'<numFmt numFmtId="166" formatCode="0.00&quot; h&quot;"/>' +
		'<numFmt numFmtId="167" formatCode="#,##0\\ \\m"/>' +
		'<numFmt numFmtId="168" formatCode="[Red]0.00;[Blue]-0.00"/>' +
		'<numFmt numFmtId="169" formatCode="_(* #,##0_);_(* \\(#,##0\\);_(* &quot;-&quot;_);_(@_)"/>' +
		'<numFmt numFmtId="170" formatCode="[$-409]d\\-mmm\\-yy;@"/>' +
		'</numFmts>' +
		// The formats of cell styles (cellStyleXfs) are not what s= counts in.
		'<cellStyleXfs count="1"><xf numFmtId="14"/></cellStyleXfs>' +
		'<cellXfs count="' + formats.length + '">' + formats.map((id) => `<xf numFmtId="${id}" xfId="0"><alignment/></xf>`).join('') + '</cellXfs>';
	const row = formats.map((id, index) => `<c r="${String.fromCharCode(65 + index)}1" s="${index}"><v>46082.5</v></c>`).join('');
	const sheets = [{ name: 'Data', xml: worksheet(`<row r="1">${row}</row><row r="2"><c r="A2" s="1"><f>TODAY()</f><v>46082</v></c></row>`) }];
	const date = '2026-03-01T12:00:00.000Z';
	const expected = {
		A1: 46082.5, B1: date, C1: date, D1: date, E1: 46082.5, F1: 46082.5, G1: 46082.5, H1: 46082.5, I1: date, J1: date, K1: date,
		L1: 46082.5, M1: date, A2: '2026-03-01T00:00:00.000Z'
	};
	assert.deepEqual(await cellsOf(xlsx({ sheets, styles })), expected);

	// The 1904 date system counts from 1904: the same serial number is four years and a day later.
	for (const flag of ['1', 'true']) {
		const later = await cellsOf(xlsx({ sheets, styles, workbookPr: `<workbookPr date1904="${flag}"/>` }));
		assert.equal(later.B1, '2030-03-02T12:00:00.000Z');
		assert.equal((await openXlsx(xlsx({ sheets, styles, workbookPr: `<workbookPr date1904="${flag}"/>` }))).date1904, true);
	}
	assert.equal((await openXlsx(xlsx({ sheets, styles, workbookPr: '<workbookPr date1904="0"/>' }))).date1904, false);
});

test('number formats: what counts as a date', () => {
	for (const code of ['yyyy-mm-dd', 'm/d/yy', 'h:mm AM/PM', '[h]:mm', '[mm]:ss', '[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy', 'DD.MM.YYYY', 'yyyy"年"m"月"']) {
		assert.equal(isDateFormat(code), true, code);
	}
	for (const code of ['General', '0.00', '#,##0', '0.00" days"', '0\\h', '#,##0_m', '0*s', '[Red]0', '[$€-407]#,##0.00', '@', '0.00E+00']) {
		assert.equal(isDateFormat(code), false, code);
	}
});

// ---------------------------------------------------------------------------
// The package
// ---------------------------------------------------------------------------

test('package: prefixed element names (x:c), as some programs write them', async () => {
	const prefixed = (xml) => xml.replace(/<(\/?)(?![?!])(\w+)/g, '<$1x:$2').replace(`xmlns="${MAIN}"`, `xmlns:x="${MAIN}"`);
	const parts = workbookParts({
		sharedStrings: '<si><r><t>a</t></r><r><t>b</t></r></si>',
		styles: '<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs>',
		workbookPr: '<workbookPr date1904="1"/>',
		sheets: [{ name: 'Data', xml: worksheet('<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" s="1"><v>1</v></c></row>') }]
	});
	for (const name of ['xl/workbook.xml', 'xl/sharedStrings.xml', 'xl/styles.xml', 'xl/worksheets/sheet1.xml']) {
		parts[name] = prefixed(parts[name]);
	}
	assert.match(parts['xl/worksheets/sheet1.xml'], /<x:sheetData><x:row r="1"><x:c/);
	assert.deepEqual(await cellsOf(zipParts(parts)), { A1: 'ab', B1: '1904-01-02T00:00:00.000Z' });
});

test('package: Strict namespaces, renamed and absolute sheet parts, and sheets that have no cells', async () => {
	const file = xlsx({
		namespaces: { main: 'http://purl.oclc.org/ooxml/spreadsheetml/main', relationships: 'http://purl.oclc.org/ooxml/officeDocument/relationships' },
		sheets: [
			{ name: 'R&amp;D', part: 'data/first%20sheet.xml', zipName: 'xl/data/first sheet.xml', xml: worksheet('<row r="1"><c r="A1"><v>1</v></c></row>') },
			{ name: 'Second', target: '/xl/elsewhere.xml', zipName: 'xl/elsewhere.xml', xml: worksheet('<row r="1"><c r="A1"><v>2</v></c></row>') },
			{ name: 'Chart', type: 'chartsheet', part: 'chartsheets/sheet1.xml', xml: '<chartsheet/>' },
			{ name: 'Gone' }
		]
	});
	const workbook = await openXlsx(file);
	assert.deepEqual(workbook.sheetNames, ['R&D', 'Second', 'Chart', 'Gone']);
	assert.deepEqual(['R&D', 'Second', 'Chart', 'Gone', 'Other'].map(workbook.hasSheet), [true, true, false, false, false]);
	assert.deepEqual(await cellsOf(file, 'R&D'), { A1: 1 });
	assert.deepEqual(await cellsOf(file, 'Second'), { A1: 2 });
	assert.equal(await workbook.readSheet('Chart'), null);
	assert.equal(await workbook.readSheet('Gone'), null);
});

test('package: sheets are found through the relationships, never by guessing part names', async () => {
	const parts = workbookParts({
		sheets: [
			{ name: 'Data', part: 'worksheets/sheet2.xml', xml: worksheet('<row r="1"><c r="A1"><v>2</v></c></row>') },
			{ name: 'Other', part: 'worksheets/sheet1.xml', xml: worksheet('<row r="1"><c r="A1"><v>1</v></c></row>') }
		]
	});
	assert.deepEqual(await cellsOf(zipParts(parts), 'Data'), { A1: 2 });
	assert.deepEqual(await cellsOf(zipParts(parts), 'Other'), { A1: 1 });
});

test('package: a zip that holds no workbook is not one', async () => {
	await rejectsWith(openXlsx(zip([{ name: 'word/document.xml', data: '<document/>' }])), 'XLSX_NOT_WORKBOOK');
	await assert.rejects(parseWorkbook(zip([{ name: 'word/document.xml', data: '<document/>' }])), { code: 'MCE_NOT_A_WORKBOOK' });
});

test('package: a UTF-16 part is read', async () => {
	const parts = workbookParts({ sharedStrings: '<si><t>wide</t></si>', sheets: [{ name: 'Data', xml: worksheet('<row r="1"><c r="A1" t="s"><v>0</v></c></row>') }] });
	const entries = Object.entries(parts).map(([name, data]) => ({
		name, data: name === 'xl/sharedStrings.xml' ? Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(data, 'utf16le')]) : data
	}));
	assert.deepEqual(await cellsOf(zip(entries)), { A1: 'wide' });
});

// ---------------------------------------------------------------------------
// parseWorkbook on top of the reader
// ---------------------------------------------------------------------------

test('parseWorkbook: a formula whose saved result is 0 reads as 0', async () => {
	const rows = [[{ values: { item_number: 'P-0001' }, modifiedOn: null }]];
	const edited = await editWorkbook(await buildWorkbook(readFixture('part-with-bom.schema.json'), { rows, maxRows: 5 }), (book) => {
		book.set('Part', 'H2', { formula: '1-1', result: 0 });
	});
	assert.deepEqual((await parseWorkbook(edited)).sheets[0].rows[0].cells, { item_number: 'P-0001', weight: 0 });
});

test('parseWorkbook: a workbook without _meta is turned away before its sheets are unpacked', async () => {
	// The sheet's packed data is damaged: reading it would fail.
	const parts = workbookParts({ sheets: [{ name: 'Part', xml: worksheet('') }] });
	const file = zip(Object.entries(parts).map(([name, data]) => ({
		name, data, stored: name === 'xl/worksheets/sheet1.xml' ? Buffer.from([0xff, 0xff, 0xff]) : undefined
	})));
	await assert.rejects(parseWorkbook(file), { code: 'MCE_META_MISSING' });
});

test('parseWorkbook: bytes from another window are read as they are', async () => {
	const file = new Uint8Array(await buildWorkbook(readFixture('part-with-bom.schema.json'), {
		rows: [[{ values: { item_number: 'P-0001' }, modifiedOn: null }]], maxRows: 5
	}));
	const otherWindow = vm.runInNewContext('(length) => new ArrayBuffer(length)')(file.length);
	new Uint8Array(otherWindow).set(file);
	assert.deepEqual((await parseWorkbook(otherWindow)).sheets[0].rows[0].cells, { item_number: 'P-0001' });
});
