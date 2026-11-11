/**
 * Test tools for the workbook module: the module itself, a small reader that
 * looks inside a generated workbook, an editor that changes one the way a user
 * would in Excel, and a reader for the raw XML parts of an .xlsx.
 *
 * The reader and the editor are written for the tests only, straight on the
 * SpreadsheetML parts and node's zlib. They share no code with the tool's own
 * reader and writer, so they check what the tool writes independently.
 */

import { readFileSync } from 'node:fs';
import { crc32, deflateRawSync, inflateRawSync } from 'node:zlib';

const clientFolder = new URL('../../../1.0.0/Innovator/Client/customer/masscreateedit/', import.meta.url);

export const moduleUrl = new URL('mce-workbook.js', clientFolder);

export const { buildWorkbook, parseWorkbook } = await import(moduleUrl.href);

export function readFixture(name) {
	return JSON.parse(readFileSync(new URL('../fixtures/' + name, import.meta.url), 'utf8'));
}

// ---------------------------------------------------------------------------
// The zip container
// ---------------------------------------------------------------------------

/** The parts of an .xlsx (a zip file) as text, by name, in the order of the file. */
export function readParts(arrayBuffer) {
	const bytes = Buffer.from(arrayBuffer);
	const end = bytes.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
	const count = bytes.readUInt16LE(end + 10);
	const parts = {};
	let offset = bytes.readUInt32LE(end + 16);
	for (let index = 0; index < count; index++) {
		const method = bytes.readUInt16LE(offset + 10);
		const compressedSize = bytes.readUInt32LE(offset + 20);
		const nameLength = bytes.readUInt16LE(offset + 28);
		const extraLength = bytes.readUInt16LE(offset + 30);
		const commentLength = bytes.readUInt16LE(offset + 32);
		const localHeader = bytes.readUInt32LE(offset + 42);
		const name = bytes.toString('utf8', offset + 46, offset + 46 + nameLength);
		const dataStart = localHeader + 30 + bytes.readUInt16LE(localHeader + 26) + bytes.readUInt16LE(localHeader + 28);
		const data = bytes.subarray(dataStart, dataStart + compressedSize);
		parts[name] = (method === 0 ? data : inflateRawSync(data)).toString('utf8');
		offset += 46 + nameLength + extraLength + commentLength;
	}
	return parts;
}

/** A zip file of text parts, each deflated, as an ArrayBuffer. */
export function writeParts(parts) {
	const locals = [];
	const central = [];
	let offset = 0;
	for (const [name, text] of Object.entries(parts)) {
		const nameBytes = Buffer.from(name, 'utf8');
		const data = Buffer.from(text, 'utf8');
		const body = deflateRawSync(data);
		const local = Buffer.alloc(30);
		local.writeUInt32LE(0x04034b50, 0);
		local.writeUInt16LE(20, 4);
		local.writeUInt16LE(0x0800, 6);
		local.writeUInt16LE(8, 8);
		local.writeUInt32LE(crc32(data), 14);
		local.writeUInt32LE(body.length, 18);
		local.writeUInt32LE(data.length, 22);
		local.writeUInt16LE(nameBytes.length, 26);
		const entry = Buffer.alloc(46);
		entry.writeUInt32LE(0x02014b50, 0);
		entry.writeUInt16LE(20, 4);
		entry.writeUInt16LE(20, 6);
		entry.writeUInt16LE(0x0800, 8);
		entry.writeUInt16LE(8, 10);
		entry.writeUInt32LE(crc32(data), 16);
		entry.writeUInt32LE(body.length, 20);
		entry.writeUInt32LE(data.length, 24);
		entry.writeUInt16LE(nameBytes.length, 28);
		entry.writeUInt32LE(offset, 42);
		locals.push(local, nameBytes, body);
		central.push(entry, nameBytes);
		offset += 30 + nameBytes.length + body.length;
	}
	const centralSize = central.reduce((sum, chunk) => sum + chunk.length, 0);
	const end = Buffer.alloc(22);
	end.writeUInt32LE(0x06054b50, 0);
	end.writeUInt16LE(central.length / 2, 8);
	end.writeUInt16LE(central.length / 2, 10);
	end.writeUInt32LE(centralSize, 12);
	end.writeUInt32LE(offset, 16);
	const file = Buffer.concat([...locals, ...central, end]);
	return file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength);
}

// ---------------------------------------------------------------------------
// XML text
// ---------------------------------------------------------------------------

const ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: '\'' };

/** Text or an attribute value as it reads: entities and character references resolved. */
const decode = (text) => text.replace(/&(#x[0-9a-fA-F]+|#\d+|\w+);/g, (match, name) => (name[0] === '#'
	? String.fromCodePoint(name[1] === 'x' ? parseInt(name.slice(2), 16) : Number(name.slice(1)))
	: ENTITIES[name] ?? match));

/** Excel's own escape of a character in a text: _xHHHH_. */
const unescapeExcel = (text) => text.replace(/_x([0-9A-Fa-f]{4})_/g, (match, code) => String.fromCharCode(parseInt(code, 16)));

const escapeText = (text) => String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\r/g, '&#13;');

/** The attributes of a start tag, by name, decoded. */
function attributesOf(tag) {
	const attributes = {};
	for (const [, name, value] of tag.matchAll(/([\w:]+)="([^"]*)"/g)) {
		attributes[name] = decode(value);
	}
	return attributes;
}

/** The text of an element holding <t> runs (a shared string, an inline string, a comment). */
const runsText = (xml) => [...xml.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>|<t(?:\s[^>]*)?\/>/g)]
	.map((match) => unescapeExcel(decode(match[1] || ''))).join('');

/** Every element with the given name: its attributes and its content. */
const elements = (xml, name) => [...xml.matchAll(new RegExp('<' + name + '(\\s[^>]*?)?(?:/>|>([\\s\\S]*?)</' + name + '>)', 'g'))]
	.map((match) => ({ attributes: attributesOf(match[1] || ''), content: match[2] || '' }));

// ---------------------------------------------------------------------------
// Cell references
// ---------------------------------------------------------------------------

export const columnIndex = (letters) => [...letters].reduce((sum, letter) => sum * 26 + letter.charCodeAt(0) - 64, 0);

export function columnLetters(index) {
	let letters = '';
	for (let rest = index; rest > 0; rest = Math.floor((rest - 1) / 26)) {
		letters = String.fromCharCode(65 + ((rest - 1) % 26)) + letters;
	}
	return letters;
}

const splitRef = (ref) => {
	const match = /^\$?([A-Z]+)\$?(\d+)$/.exec(ref);
	return { column: columnIndex(match[1]), row: Number(match[2]) };
};

/** Whether a cell lies in a list of ranges such as "A2:A11 C2". */
function inRanges(sqref, ref) {
	const cell = splitRef(ref);
	return sqref.split(/\s+/).some((range) => {
		const [from, to = from] = range.split(':').map(splitRef);
		return cell.column >= from.column && cell.column <= to.column && cell.row >= from.row && cell.row <= to.row;
	});
}

// ---------------------------------------------------------------------------
// Reading a workbook
// ---------------------------------------------------------------------------

const BUILT_IN_FORMATS = { 0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 14: 'mm-dd-yy', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yy h:mm', 49: '@' };
const isDateFormat = (code) => code !== 'General' && /[ymdhs]/i.test(code.replace(/"[^"]*"|\[[^\]]*\]/g, ''));
const serialToDate = (serial) => new Date(Math.round((serial - 25569) * 86400000));

/** The cell styles of a styles part: number format, bold, fill colour and locking of each. */
function readStyles(xml) {
	const formats = { ...BUILT_IN_FORMATS };
	for (const { attributes } of elements(xml, 'numFmt')) {
		formats[attributes.numFmtId] = attributes.formatCode;
	}
	const fonts = elements(/<fonts[\s\S]*?<\/fonts>/.exec(xml)[0], 'font').map((font) => ({ bold: /<b\/>/.test(font.content) }));
	const fills = elements(/<fills[\s\S]*?<\/fills>/.exec(xml)[0], 'fill').map((fill) => (/<fgColor rgb="(\w+)"/.exec(fill.content) || [])[1] || null);
	const xfs = elements(/<cellXfs[\s\S]*?<\/cellXfs>/.exec(xml)[0], 'xf').map(({ attributes, content }) => ({
		numFmt: formats[attributes.numFmtId || 0],
		bold: fonts[attributes.fontId || 0].bold,
		fill: fills[attributes.fillId || 0],
		locked: !/<protection locked="0"\/>/.test(content)
	}));
	const dxfs = elements(/<dxfs[\s\S]*?<\/dxfs>|$/.exec(xml)[0], 'dxf').map(({ content }) => ({
		fontColor: (/<font>[\s\S]*?<color rgb="(\w+)"/.exec(content) || [])[1] || null,
		fillColor: (/<bgColor rgb="(\w+)"/.exec(content) || [])[1] || null
	}));
	return { xfs, dxfs };
}

const VALIDATION_FLAGS = ['allowBlank', 'showInputMessage', 'showErrorMessage'];
const VALIDATION_TEXTS = ['errorStyle', 'errorTitle', 'error', 'promptTitle', 'prompt'];
const OPERATOR_TYPES = ['whole', 'decimal', 'date', 'time', 'textLength'];

/** A data validation entry as it reads: type "any" when none is given, "between" as the default operator, numbers as numbers. */
function readValidation({ attributes, content }) {
	const type = attributes.type || 'any';
	const validation = { type };
	const operator = attributes.operator || (OPERATOR_TYPES.includes(type) ? 'between' : undefined);
	if (operator) {
		validation.operator = operator;
	}
	for (const flag of VALIDATION_FLAGS) {
		if (attributes[flag] === '1') {
			validation[flag] = true;
		}
	}
	for (const text of VALIDATION_TEXTS) {
		if (attributes[text] !== undefined) {
			validation[text] = attributes[text];
		}
	}
	const formulae = [...content.matchAll(/<formula[12]>([\s\S]*?)<\/formula[12]>/g)].map((match) => decode(match[1]))
		.map((formula) => (/^-?\d+(\.\d+)?(E[-+]?\d+)?$/i.test(formula) ? Number(formula) : formula));
	if (formulae.length) {
		validation.formulae = formulae;
	}
	return validation;
}

/** The value of a <c> element, as Excel shows it: a text, a number, true or false, a Date when the format is a date format, or null. */
function cellValue(attributes, content, strings, numFmt) {
	const raw = (/<v>([\s\S]*?)<\/v>/.exec(content) || [])[1];
	switch (attributes.t) {
		case 's':
			return strings[Number(raw)];
		case 'inlineStr':
			return runsText((/<is>([\s\S]*?)<\/is>/.exec(content) || [])[1] || '');
		case 'str':
		case 'e':
			return raw === undefined ? null : unescapeExcel(decode(raw));
		case 'b':
			return raw === '1';
		default:
			if (raw === undefined) {
				return null;
			}
			return isDateFormat(numFmt) ? serialToDate(Number(raw)) : Number(raw);
	}
}

/** Reads one worksheet part and the notes that go with it. */
function readSheet(name, state, xml, notesXml, strings, styles) {
	const cells = new Map();
	for (const match of xml.matchAll(/<c r="([A-Z]+\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
		cells.set(match[1], { attributes: attributesOf(match[2]), content: match[3] || '' });
	}
	const columns = elements(xml, 'col').map(({ attributes }) => ({
		min: Number(attributes.min), max: Number(attributes.max), hidden: attributes.hidden === '1', style: Number(attributes.style || 0)
	}));
	const columnOf = (index) => columns.find((column) => index >= column.min && index <= column.max);
	const notes = new Map(elements(notesXml || '', 'comment').map(({ attributes, content }) => [attributes.ref, runsText(content)]));
	const validations = elements(xml, 'dataValidation').map((entry) => ({ sqref: entry.attributes.sqref, validation: readValidation(entry) }));
	const pane = elements(xml, 'pane')[0];
	const protection = elements(xml, 'sheetProtection')[0];

	const sheet = {
		name,
		state,
		tabColor: (/<tabColor rgb="(\w+)"/.exec(xml) || [])[1] || null,
		/** The frozen rows at the top, or null. */
		frozen: pane && pane.attributes.state === 'frozen' ? { rows: Number(pane.attributes.ySplit), topLeftCell: pane.attributes.topLeftCell } : null,
		/** The attributes of <sheetProtection> as the file has them: true protects that part of the sheet, false leaves it free. */
		protection: protection ? Object.fromEntries(Object.entries(protection.attributes).map(([flag, value]) =>
			[flag, value === '1' ? true : value === '0' ? false : value])) : null,
		conditionalFormats: elements(xml, 'conditionalFormatting').map(({ attributes, content }) => ({
			sqref: attributes.sqref,
			rules: elements(content, 'cfRule').map((rule) => ({
				type: rule.attributes.type,
				formulae: [...rule.content.matchAll(/<formula>([\s\S]*?)<\/formula>/g)].map((match) => decode(match[1])),
				style: styles.dxfs[Number(rule.attributes.dxfId)]
			}))
		})),
		/** A column: whether it is hidden, and the style of its cells. */
		column(index) {
			const column = columnOf(index);
			return { hidden: !!(column && column.hidden), ...styles.xfs[column ? column.style : 0] };
		},
		/**
		 * A cell by reference ("B2") or by row and column: its value, its style (numFmt, bold, fill, locked), its note
		 * ('' for none) and its data validation (undefined for none). A cell the file leaves out has the style of its column.
		 */
		cell(refOrRow, columnNumber) {
			const ref = columnNumber === undefined ? refOrRow : columnLetters(columnNumber) + refOrRow;
			const found = cells.get(ref);
			const styleIndex = found && found.attributes.s !== undefined ? Number(found.attributes.s) : (columnOf(splitRef(ref).column) || { style: 0 }).style;
			const style = styles.xfs[styleIndex];
			const entry = validations.find((candidate) => inRanges(candidate.sqref, ref));
			return {
				...style,
				value: found ? cellValue(found.attributes, found.content, strings, style.numFmt) : null,
				note: notes.get(ref) || '',
				validation: entry ? entry.validation : undefined
			};
		},
		/** Every cell the file holds, with its reference and value, in the order of the file. */
		cells() {
			return [...cells.keys()].map((ref) => ({ ref, value: sheet.cell(ref).value }));
		}
	};
	return sheet;
}

/** Resolves a relationship target against the folder of the part that names it. */
const resolveTarget = (folder, target) => {
	const segments = (target.startsWith('/') ? target.slice(1) : folder + target).split('/');
	const resolved = [];
	for (const segment of segments) {
		if (segment === '..') {
			resolved.pop();
		} else if (segment !== '.') {
			resolved.push(segment);
		}
	}
	return resolved.join('/');
};

const relationshipsOf = (parts, partName) => {
	const slash = partName.lastIndexOf('/');
	const folder = partName.slice(0, slash + 1);
	const xml = parts[folder + '_rels/' + partName.slice(slash + 1) + '.rels'] || '';
	return elements(xml, 'Relationship').map(({ attributes }) => ({ ...attributes, part: resolveTarget(folder, attributes.Target) }));
};

/**
 * Opens a workbook for inspection.
 *
 * @returns {{sheets: object[], sheet: function(string): object, definedName: function(string): (string|undefined), creator: (string|null), modified: (Date|null)}}
 */
export function openWorkbook(arrayBuffer) {
	const parts = readParts(arrayBuffer);
	const workbookXml = parts['xl/workbook.xml'];
	const workbookRelationships = relationshipsOf(parts, 'xl/workbook.xml');
	const strings = elements(parts['xl/sharedStrings.xml'] || '', 'si').map(({ content }) => runsText(content));
	const styles = readStyles(parts['xl/styles.xml']);
	const sheets = elements(workbookXml, 'sheet').map(({ attributes }) => {
		const partName = workbookRelationships.find((relationship) => relationship.Id === attributes['r:id']).part;
		const comments = relationshipsOf(parts, partName).find((relationship) => relationship.Type.endsWith('/comments'));
		return readSheet(attributes.name, attributes.state || 'visible', parts[partName], comments ? parts[comments.part] : '', strings, styles);
	});
	const definedNames = new Map(elements(workbookXml, 'definedName').map(({ attributes, content }) => [attributes.name, decode(content)]));
	const core = parts['docProps/core.xml'] || '';
	const creator = /<dc:creator>([^<]*)</.exec(core);
	const modified = /<dcterms:modified[^>]*>([^<]*)</.exec(core);
	return {
		sheets,
		sheet: (name) => sheets.find((sheet) => sheet.name === name),
		definedName: (name) => definedNames.get(name),
		creator: creator ? decode(creator[1]) : null,
		modified: modified ? new Date(modified[1]) : null
	};
}

// ---------------------------------------------------------------------------
// Editing a workbook
// ---------------------------------------------------------------------------

/**
 * The XML of a cell with the given value, as Excel saves it. A value is a text, a number, true or
 * false, a Date (saved as its serial number, from the UTC fields), null (an empty cell), or one of
 *   { formula, result }      a formula and the result Excel saved with it (none: not calculated)
 *   { richText: [{ text, bold }] }  text in runs of their own formatting
 *   { error: '#N/A' }       an error value
 */
function cellXml(ref, styleIndex, value) {
	const head = '<c r="' + ref + '"' + (styleIndex ? ' s="' + styleIndex + '"' : '');
	if (value === null || value === undefined) {
		return head + '/>';
	}
	if (value instanceof Date) {
		return head + '><v>' + (value.getTime() / 86400000 + 25569) + '</v></c>';
	}
	if (typeof value === 'number') {
		return head + '><v>' + value + '</v></c>';
	}
	if (typeof value === 'boolean') {
		return head + ' t="b"><v>' + (value ? 1 : 0) + '</v></c>';
	}
	if (typeof value === 'string') {
		return head + ' t="inlineStr"><is><t xml:space="preserve">' + escapeText(value) + '</t></is></c>';
	}
	if (value.richText) {
		return head + ' t="inlineStr"><is>' + value.richText.map((run) =>
			'<r>' + (run.bold ? '<rPr><b/></rPr>' : '') + '<t xml:space="preserve">' + escapeText(run.text) + '</t></r>').join('') + '</is></c>';
	}
	if (value.error) {
		return head + ' t="e"><v>' + escapeText(value.error) + '</v></c>';
	}
	if (value.formula !== undefined) {
		const formula = '<f>' + escapeText(value.formula) + '</f>';
		if (value.result === undefined) {
			return head + '>' + formula + '</c>';
		}
		return typeof value.result === 'string'
			? head + ' t="str">' + formula + '<v>' + escapeText(value.result) + '</v></c>'
			: head + '>' + formula + '<v>' + value.result + '</v></c>';
	}
	throw new Error('No cell XML for ' + JSON.stringify(value));
}

/** Sets one cell in the XML of a worksheet, keeping its style, or the style given; adds the row and the cell where they are missing. */
function setCell(xml, ref, value, styleIndex) {
	const { row } = splitRef(ref);
	const columnStyle = () => {
		const column = elements(xml, 'col').find(({ attributes }) => splitRef(ref).column >= Number(attributes.min) && splitRef(ref).column <= Number(attributes.max));
		return column ? Number(column.attributes.style || 0) : 0;
	};
	const existing = new RegExp('<c r="' + ref + '"([^>]*?)(?:/>|>[\\s\\S]*?</c>)').exec(xml);
	if (existing) {
		const style = styleIndex ?? Number(attributesOf(existing[1]).s || 0);
		return xml.slice(0, existing.index) + cellXml(ref, style, value) + xml.slice(existing.index + existing[0].length);
	}
	const style = styleIndex ?? columnStyle();
	const newCell = cellXml(ref, style, value);
	const rowMatch = new RegExp('<row r="' + row + '"([^>]*?)(?:/>|>([\\s\\S]*?)</row>)').exec(xml);
	if (rowMatch) {
		// The cell goes before the first cell of a later column.
		const cellsXml = rowMatch[2] || '';
		const later = [...cellsXml.matchAll(/<c r="([A-Z]+)\d+"/g)].find((match) => columnIndex(match[1]) > splitRef(ref).column);
		const at = later ? later.index : cellsXml.length;
		const rowXml = '<row r="' + row + '">' + cellsXml.slice(0, at) + newCell + cellsXml.slice(at) + '</row>';
		return xml.slice(0, rowMatch.index) + rowXml + xml.slice(rowMatch.index + rowMatch[0].length);
	}
	// A new row goes before the first later row.
	const laterRow = [...xml.matchAll(/<row r="(\d+)"/g)].find((match) => Number(match[1]) > row);
	const rowXml = '<row r="' + row + '">' + newCell + '</row>';
	if (laterRow) {
		return xml.slice(0, laterRow.index) + rowXml + xml.slice(laterRow.index);
	}
	return xml.replace(/<sheetData\/>/, '<sheetData></sheetData>').replace('</sheetData>', rowXml + '</sheetData>');
}

/**
 * Changes a built workbook the way a user would in Excel, and returns the saved file.
 *
 * `edit` gets an editor:
 *   set(sheetName, ref, value, { numFmt: 'General' })  sets a cell (see cellXml for the values); the
 *                       option takes the cell's own format away, as choosing General in Excel does
 *   link(sheetName, ref, text, address)               a cell with a hyperlink
 *   removeSheet(sheetName)                             deletes a worksheet
 *   workbook                                           the workbook as openWorkbook reads it, before the edit
 */
export async function editWorkbook(arrayBuffer, edit) {
	const parts = readParts(arrayBuffer);
	const workbookRelationships = () => relationshipsOf(parts, 'xl/workbook.xml');
	const partOf = (sheetName) => {
		const sheet = elements(parts['xl/workbook.xml'], 'sheet').find(({ attributes }) => attributes.name === sheetName);
		if (!sheet) {
			throw new Error('No sheet ' + sheetName);
		}
		return workbookRelationships().find((relationship) => relationship.Id === sheet.attributes['r:id']);
	};
	const editor = {
		workbook: openWorkbook(arrayBuffer),
		set(sheetName, ref, value, options = {}) {
			const { part } = partOf(sheetName);
			parts[part] = setCell(parts[part], ref, value, options.numFmt === 'General' ? 0 : undefined);
		},
		link(sheetName, ref, text, address) {
			editor.set(sheetName, ref, text);
			const { part } = partOf(sheetName);
			const slash = part.lastIndexOf('/');
			const relsName = part.slice(0, slash + 1) + '_rels/' + part.slice(slash + 1) + '.rels';
			const rels = parts[relsName] || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
				'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
			const id = 'rIdLink' + ((rels.match(/<Relationship /g) || []).length + 1);
			parts[relsName] = rels.replace('</Relationships>', '<Relationship Id="' + id + '" ' +
				'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="' + escapeText(address) + '" TargetMode="External"/></Relationships>');
			// <hyperlinks> comes after the validation and before the page margins.
			parts[part] = parts[part].replace('<pageMargins', '<hyperlinks><hyperlink ref="' + ref + '" r:id="' + id + '"/></hyperlinks><pageMargins');
		},
		removeSheet(sheetName) {
			const relationship = partOf(sheetName);
			const sheetPart = relationship.part;
			const slash = sheetPart.lastIndexOf('/');
			const removed = [sheetPart, ...relationshipsOf(parts, sheetPart).filter((related) => related.TargetMode !== 'External').map((related) => related.part)];
			delete parts[sheetPart.slice(0, slash + 1) + '_rels/' + sheetPart.slice(slash + 1) + '.rels'];
			for (const part of removed) {
				delete parts[part];
				parts['[Content_Types].xml'] = parts['[Content_Types].xml'].replace('<Override PartName="/' + part + '"', '<Removed ').replace(/<Removed [^>]*\/>/, '');
			}
			parts['xl/workbook.xml'] = parts['xl/workbook.xml'].replace(new RegExp('<sheet [^>]*r:id="' + relationship.Id + '"[^>]*/>'), '');
			parts['xl/_rels/workbook.xml.rels'] = parts['xl/_rels/workbook.xml.rels'].replace(new RegExp('<Relationship Id="' + relationship.Id + '"[^>]*/>'), '');
		}
	};
	await edit(editor);
	return writeParts(parts);
}

// ---------------------------------------------------------------------------
// A workbook from elsewhere
// ---------------------------------------------------------------------------

/** A plain workbook with text cells, as another program would save it: [{ name, rows: [['A1', 'B1'], ...] }]. */
export function plainWorkbook(sheets) {
	const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
	const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
	const PACKAGE_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
	const TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.';
	const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
	const parts = {
		'[Content_Types].xml': head + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
			'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
			'<Override PartName="/xl/workbook.xml" ContentType="' + TYPE + 'sheet.main+xml"/>' +
			sheets.map((sheet, at) => '<Override PartName="/xl/worksheets/sheet' + (at + 1) + '.xml" ContentType="' + TYPE + 'worksheet+xml"/>').join('') + '</Types>',
		'_rels/.rels': head + '<Relationships xmlns="' + PACKAGE_REL + '"><Relationship Id="rId1" Type="' + REL + '/officeDocument" Target="xl/workbook.xml"/></Relationships>',
		'xl/workbook.xml': head + '<workbook xmlns="' + MAIN + '" xmlns:r="' + REL + '"><sheets>' +
			sheets.map((sheet, at) => '<sheet name="' + escapeText(sheet.name) + '" sheetId="' + (at + 1) + '" r:id="rId' + (at + 1) + '"/>').join('') + '</sheets></workbook>',
		'xl/_rels/workbook.xml.rels': head + '<Relationships xmlns="' + PACKAGE_REL + '">' +
			sheets.map((sheet, at) => '<Relationship Id="rId' + (at + 1) + '" Type="' + REL + '/worksheet" Target="worksheets/sheet' + (at + 1) + '.xml"/>').join('') + '</Relationships>'
	};
	sheets.forEach((sheet, at) => {
		const rows = sheet.rows.map((values, index) => '<row r="' + (index + 1) + '">' +
			values.map((value, column) => cellXml(columnLetters(column + 1) + (index + 1), 0, value)).join('') + '</row>').join('');
		parts['xl/worksheets/sheet' + (at + 1) + '.xml'] = head + '<worksheet xmlns="' + MAIN + '"><sheetData>' + rows + '</sheetData></worksheet>';
	});
	return writeParts(parts);
}
