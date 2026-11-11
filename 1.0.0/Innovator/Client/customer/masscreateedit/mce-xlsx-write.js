/**
 * Mass Create & Edit - a small writer for Excel workbooks (.xlsx).
 *
 * writeXlsx(workbook) turns a plain description of a workbook into the bytes
 * of an .xlsx file: a zip container of SpreadsheetML parts. It writes only what
 * the tool's workbooks use (text, number and date cells, cell styles, column
 * widths, frozen header rows, data validation, conditional formatting, sheet
 * protection without a password, notes, defined names and the document
 * properties), and it needs no library: the parts are deflated with the
 * platform's CompressionStream, and stored as they are where there is none.
 * It runs the same in a browser and under node.
 *
 * The description:
 *
 *   {
 *     creator: 'Mass Create & Edit', created: Date, modified: Date,
 *     definedNames: [{ name: 'mce_list_1', sheet: '_lists', range: '$A$2:$A$9' }],
 *     sheets: [{
 *       name: 'Part',
 *       state: 'visible' | 'hidden' | 'veryHidden',
 *       tabColor: 'FF4472C4',
 *       frozenRows: 1,
 *       columns: [{ index: 1, width: 24, hidden: false, style }],
 *       rows: [{ index: 1, height: 30, cells: [{ column: 1, value, style, note: 'text' }] }],
 *       protection: { sheet: true, objects: true, formatRows: false, ... },
 *       conditionalFormats: [{ sqref: 'H2:H11', rules: [{ type: 'expression', priority: 1, formulae: ['...'], style: dxf }] }],
 *       dataValidations: [{ sqref: 'A2:A11', type: 'list', formulae: ['mce_list_1'], allowBlank: true, ... }]
 *     }]
 *   }
 *
 * Indexes of rows and columns start at 1. A cell value is a string, a finite
 * number, a boolean or a Date; a Date is written as Excel's serial number from
 * its UTC fields, so give the cell a date format. A style is
 *   { numFmt: 'yyyy-mm-dd', font: { name, size, bold, italic, color },
 *     fill: { color }, alignment: { horizontal, vertical, wrapText }, locked }
 * where colours are ARGB ('FFDDEBF7'), a fill is solid, and `locked` sets the
 * cell's protection. Equal styles share one entry of the styles part. The
 * style of a conditional format (a dxf) takes { font: { color }, fill: { color } }.
 * The protection flags are the attributes of the file's <sheetProtection>: a
 * flag set to true protects that part of the sheet, false leaves it free.
 */

const encoder = new TextEncoder();

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.';
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

// The number formats Excel knows by number; any other format gets a number from 164 up.
const BUILT_IN_FORMATS = new Map([
	['General', 0], ['0', 1], ['0.00', 2], ['#,##0', 3], ['#,##0.00', 4], ['0%', 9], ['0.00%', 10], ['0.00E+00', 11],
	['mm-dd-yy', 14], ['d-mmm-yy', 15], ['d-mmm', 16], ['mmm-yy', 17], ['h:mm', 20], ['h:mm:ss', 21], ['@', 49]
]);
const FIRST_CUSTOM_FORMAT = 164;

const PROTECTION_FLAGS = [
	'sheet', 'objects', 'scenarios', 'formatCells', 'formatColumns', 'formatRows', 'insertColumns', 'insertRows',
	'insertHyperlinks', 'deleteColumns', 'deleteRows', 'selectLockedCells', 'sort', 'autoFilter', 'pivotTables', 'selectUnlockedCells'
];

// A column without a width of its own; a <col> without one would be shown 0 wide.
const DEFAULT_COLUMN_WIDTH = 9;

// ---------------------------------------------------------------------------
// Text
// ---------------------------------------------------------------------------

// Characters XML 1.0 does not allow at all. Tab, line feed and carriage return are allowed. A lone
// surrogate is not allowed either, but TextEncoder already writes it as U+FFFD.
const NOT_XML = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g;
const NEEDS_ESCAPING = /[&<>"\r\n\t_\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/;

/**
 * A text as XML character data, or as an attribute value when `attribute` is set.
 *
 * Characters XML can't hold are dropped. Excel reads _xHHHH_ in a text as the
 * character with that code, so a text that has such a sequence of its own gets
 * its underscore escaped as _x005F_, and comes back unchanged. Escaping one
 * that a reader would not have decoded does no harm: _x005F_ is always an
 * underscore. A carriage return is written as a reference, which keeps it: a
 * reader turns a bare one into a line feed. In an attribute, line feeds and tabs are references too,
 * because a reader turns a bare one into a space.
 */
function escapeXml(value, attribute = false) {
	const text = String(value);
	if (!NEEDS_ESCAPING.test(text)) {
		return text;
	}
	let escaped = text
		.replace(NOT_XML, '')
		.replace(/_([xX][0-9A-Fa-f]{4}_)/g, '_x005F_$1')
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/\r/g, '&#13;');
	if (attribute) {
		escaped = escaped.replace(/\n/g, '&#10;').replace(/\t/g, '&#9;');
	}
	return escaped;
}

const attr = (value) => escapeXml(value, true);

/** A <t> element; leading or trailing spaces and line breaks are kept only with xml:space="preserve". */
function textElement(text) {
	const preserve = /^\s|\s$|[\t\n\r]/.test(text) ? ' xml:space="preserve"' : '';
	return '<t' + preserve + '>' + escapeXml(text) + '</t>';
}

// ---------------------------------------------------------------------------
// Cells and dates
// ---------------------------------------------------------------------------

const isDate = (value) => Object.prototype.toString.call(value) === '[object Date]';

const columnNames = [''];

/** The letters of a column: 1 is A, 27 is AA. */
export function columnName(index) {
	if (!columnNames[index]) {
		let letters = '';
		for (let n = index; n > 0; n = Math.floor((n - 1) / 26)) {
			letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
		}
		columnNames[index] = letters;
	}
	return columnNames[index];
}

/**
 * Excel's serial number of a Date, from its UTC fields: days since 1899-12-30, the time as a fraction.
 *
 * Excel counts a 29 February 1900 that never was, so the serial of a day before 1 March 1900 is one
 * lower in Excel than here. The reader of these workbooks converts back the same way, so such a date
 * comes back unchanged; it is only shown a day late in Excel.
 */
export function dateToSerial(date) {
	return date.getTime() / 86400000 + 25569;
}

/** A number as a formula or a cell value: never NaN or Infinity, which Excel can't read. */
function numberText(number) {
	if (!Number.isFinite(number)) {
		throw new Error('The workbook can\'t hold the number ' + number + '.');
	}
	return String(number);
}

/** A formula of a validation or a conditional format; a number or a Date is written as a number. */
function formulaText(formula) {
	if (isDate(formula)) {
		return numberText(dateToSerial(formula));
	}
	return typeof formula === 'number' ? numberText(formula) : escapeXml(formula);
}

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

function fontXml(font) {
	// The children of <font> have a fixed order in the schema.
	let xml = '<font>';
	if (font.bold) {
		xml += '<b/>';
	}
	if (font.italic) {
		xml += '<i/>';
	}
	if (font.underline) {
		xml += '<u/>';
	}
	if (font.size) {
		xml += '<sz val="' + numberText(font.size) + '"/>';
	}
	if (font.color) {
		xml += '<color rgb="' + attr(font.color) + '"/>';
	}
	if (font.name) {
		xml += '<name val="' + attr(font.name) + '"/>';
	}
	if (font.family) {
		xml += '<family val="' + numberText(font.family) + '"/>';
	}
	return xml + '</font>';
}

/** The styles part, built up as cells ask for their styles. */
function createStyles() {
	const numFmts = [];
	const numFmtIds = new Map(BUILT_IN_FORMATS);
	const fonts = [fontXml({ size: 11, name: 'Calibri', family: 2 })];
	// The first two fills are reserved: Excel expects "none" and "gray125" there.
	const fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
	const xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
	const dxfs = [];
	const byObject = new Map();

	const indexIn = (list, xml) => {
		const found = list.indexOf(xml);
		return found >= 0 ? found : list.push(xml) - 1;
	};
	const numFmtId = (code) => {
		if (!numFmtIds.has(code)) {
			const id = FIRST_CUSTOM_FORMAT + numFmts.length;
			numFmts.push('<numFmt numFmtId="' + id + '" formatCode="' + attr(code) + '"/>');
			numFmtIds.set(code, id);
		}
		return numFmtIds.get(code);
	};

	const xfOf = (style) => {
		const numFmt = style.numFmt === undefined || style.numFmt === null ? 0 : numFmtId(String(style.numFmt));
		const font = style.font ? indexIn(fonts, fontXml(style.font)) : 0;
		const fill = style.fill && style.fill.color
			? indexIn(fills, '<fill><patternFill patternType="solid"><fgColor rgb="' + attr(style.fill.color) + '"/></patternFill></fill>')
			: 0;
		let xml = '<xf numFmtId="' + numFmt + '" fontId="' + font + '" fillId="' + fill + '" borderId="0" xfId="0"';
		let children = '';
		if (numFmt) {
			xml += ' applyNumberFormat="1"';
		}
		if (font) {
			xml += ' applyFont="1"';
		}
		if (fill) {
			xml += ' applyFill="1"';
		}
		const alignment = style.alignment;
		if (alignment) {
			xml += ' applyAlignment="1"';
			children += '<alignment' +
				(alignment.horizontal ? ' horizontal="' + attr(alignment.horizontal) + '"' : '') +
				(alignment.vertical ? ' vertical="' + attr(alignment.vertical) + '"' : '') +
				(alignment.wrapText ? ' wrapText="1"' : '') +
				(alignment.indent ? ' indent="' + numberText(alignment.indent) + '"' : '') + '/>';
		}
		if (style.locked !== undefined) {
			// A cell is locked unless its style says otherwise.
			xml += ' applyProtection="1"';
			if (!style.locked) {
				children += '<protection locked="0"/>';
			}
		}
		return children ? xml + '>' + children + '</xf>' : xml + '/>';
	};

	return {
		/** The index of a cell style in cellXfs; 0 for none. */
		xf(style) {
			if (!style) {
				return 0;
			}
			if (!byObject.has(style)) {
				byObject.set(style, indexIn(xfs, xfOf(style)));
			}
			return byObject.get(style);
		},
		/** The index of a differential format, the style of a conditional format. */
		dxf(style) {
			let xml = '<dxf>';
			if (style.font && style.font.color) {
				xml += '<font><color rgb="' + attr(style.font.color) + '"/></font>';
			}
			if (style.fill && style.fill.color) {
				// In a differential format, the colour of a solid fill is its background colour.
				xml += '<fill><patternFill patternType="solid"><bgColor rgb="' + attr(style.fill.color) + '"/></patternFill></fill>';
			}
			return indexIn(dxfs, xml + '</dxf>');
		},
		xml() {
			const list = (tag, items) => '<' + tag + ' count="' + items.length + '">' + items.join('') + '</' + tag + '>';
			return XML_HEAD + '<styleSheet xmlns="' + MAIN_NS + '">' +
				(numFmts.length ? list('numFmts', numFmts) : '') +
				list('fonts', fonts) +
				list('fills', fills) +
				'<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
				'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
				list('cellXfs', xfs) +
				'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
				list('dxfs', dxfs) +
				'</styleSheet>';
		}
	};
}

// ---------------------------------------------------------------------------
// Shared strings
// ---------------------------------------------------------------------------

/** The shared-string table: each distinct text is stored once, and a cell refers to it by index. */
function createSharedStrings() {
	const indexes = new Map();
	const items = [];
	let count = 0;
	return {
		index(text) {
			count++;
			let index = indexes.get(text);
			if (index === undefined) {
				index = items.push('<si>' + textElement(text) + '</si>') - 1;
				indexes.set(text, index);
			}
			return index;
		},
		get size() {
			return items.length;
		},
		xml() {
			return XML_HEAD + '<sst xmlns="' + MAIN_NS + '" count="' + count + '" uniqueCount="' + items.length + '">' + items.join('') + '</sst>';
		}
	};
}

// ---------------------------------------------------------------------------
// Worksheets
// ---------------------------------------------------------------------------

function cellXml(ref, cell, styles, strings) {
	const styleIndex = styles.xf(cell.style);
	const head = '<c r="' + ref + '"' + (styleIndex ? ' s="' + styleIndex + '"' : '');
	let value = cell.value;
	if (isDate(value)) {
		// An invalid Date has no serial number; the cell stays empty.
		value = Number.isNaN(value.getTime()) ? null : dateToSerial(value);
	}
	if (value === null || value === undefined || value === '') {
		return head + '/>';
	}
	if (typeof value === 'number' && Number.isFinite(value)) {
		return head + '><v>' + value + '</v></c>';
	}
	if (typeof value === 'boolean') {
		return head + ' t="b"><v>' + (value ? 1 : 0) + '</v></c>';
	}
	return head + ' t="s"><v>' + strings.index(String(value)) + '</v></c>';
}

/** The <cols> element: neighbouring columns that look the same share one <col>. */
function columnsXml(columns, styles) {
	const sorted = (columns || []).slice().sort((a, b) => a.index - b.index);
	const runs = [];
	for (const column of sorted) {
		const width = column.width || DEFAULT_COLUMN_WIDTH;
		const style = styles.xf(column.style);
		const hidden = !!column.hidden;
		const last = runs[runs.length - 1];
		if (last && last.max === column.index) {
			throw new Error('Column ' + column.index + ' is described twice.');
		}
		if (last && last.max === column.index - 1 && last.width === width && last.style === style && last.hidden === hidden) {
			last.max = column.index;
		} else {
			runs.push({ min: column.index, max: column.index, width, style, hidden });
		}
	}
	if (!runs.length) {
		return '';
	}
	return '<cols>' + runs.map((run) => '<col min="' + run.min + '" max="' + run.max + '" width="' + numberText(run.width) + '"' +
		(run.style ? ' style="' + run.style + '"' : '') + (run.hidden ? ' hidden="1"' : '') + ' customWidth="1"/>').join('') + '</cols>';
}

/** The XML of a worksheet, and the notes on it. */
function worksheetXml(sheet, selected, styles, strings) {
	const notes = [];
	const rows = (sheet.rows || []).slice().sort((a, b) => a.index - b.index);
	let first = null;
	let last = null;
	const rowParts = [];
	for (const row of rows) {
		const cells = (row.cells || []).slice().sort((a, b) => a.column - b.column);
		let xml = '';
		for (const [at, cell] of cells.entries()) {
			if (at > 0 && cells[at - 1].column === cell.column) {
				throw new Error('Sheet ' + sheet.name + ' has two cells at row ' + row.index + ', column ' + cell.column + '.');
			}
			const ref = columnName(cell.column) + row.index;
			xml += cellXml(ref, cell, styles, strings);
			if (cell.note) {
				notes.push({ ref, row: row.index, column: cell.column, text: String(cell.note) });
			}
		}
		if (!xml && !row.height) {
			continue;
		}
		const spans = cells.length ? ' spans="' + cells[0].column + ':' + cells[cells.length - 1].column + '"' : '';
		const height = row.height ? ' ht="' + numberText(row.height) + '" customHeight="1"' : '';
		rowParts.push('<row r="' + row.index + '"' + spans + height + '>' + xml + '</row>');
		if (cells.length) {
			first = first || { row: row.index, column: cells[0].column };
			first.column = Math.min(first.column, cells[0].column);
			last = last || { row: row.index, column: 0 };
			last.row = row.index;
			last.column = Math.max(last.column, cells[cells.length - 1].column);
		}
	}

	// The children of <worksheet> have a fixed order in the schema.
	let xml = XML_HEAD + '<worksheet xmlns="' + MAIN_NS + '" xmlns:r="' + REL_NS + '">';
	if (sheet.tabColor) {
		xml += '<sheetPr><tabColor rgb="' + attr(sheet.tabColor) + '"/></sheetPr>';
	}
	const topLeft = first ? columnName(first.column) + first.row : 'A1';
	const bottomRight = last ? columnName(last.column) + last.row : 'A1';
	xml += '<dimension ref="' + (topLeft === bottomRight ? topLeft : topLeft + ':' + bottomRight) + '"/>';
	const frozenRows = sheet.frozenRows > 0 ? Math.floor(sheet.frozenRows) : 0;
	if (selected || frozenRows) {
		xml += '<sheetViews><sheetView' + (selected ? ' tabSelected="1"' : '') + ' workbookViewId="0"';
		if (frozenRows) {
			const below = 'A' + (frozenRows + 1);
			xml += '><pane ySplit="' + frozenRows + '" topLeftCell="' + below + '" activePane="bottomLeft" state="frozen"/>' +
				'<selection pane="bottomLeft" activeCell="' + below + '" sqref="' + below + '"/></sheetView>';
		} else {
			xml += '/>';
		}
		xml += '</sheetViews>';
	}
	xml += '<sheetFormatPr defaultRowHeight="15"/>';
	xml += columnsXml(sheet.columns, styles);
	xml += '<sheetData>' + rowParts.join('') + '</sheetData>';

	if (sheet.protection) {
		xml += '<sheetProtection' + Object.entries(sheet.protection).map(([flag, on]) => {
			if (!PROTECTION_FLAGS.includes(flag)) {
				throw new Error('Sheet protection has no flag "' + flag + '".');
			}
			return ' ' + flag + '="' + (on ? 1 : 0) + '"';
		}).join('') + '/>';
	}

	for (const format of sheet.conditionalFormats || []) {
		xml += '<conditionalFormatting sqref="' + attr(format.sqref) + '">' + format.rules.map((rule) =>
			'<cfRule type="' + attr(rule.type || 'expression') + '" dxfId="' + styles.dxf(rule.style || {}) + '" priority="' + numberText(rule.priority) + '">' +
			(rule.formulae || []).map((formula) => '<formula>' + formulaText(formula) + '</formula>').join('') +
			'</cfRule>').join('') + '</conditionalFormatting>';
	}

	const validations = sheet.dataValidations || [];
	if (validations.length) {
		xml += '<dataValidations count="' + validations.length + '">' + validations.map(validationXml).join('') + '</dataValidations>';
	}

	xml += '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>';
	if (notes.length) {
		xml += '<legacyDrawing r:id="rId2"/>';
	}
	return { xml: xml + '</worksheet>', notes };
}

function validationXml(validation) {
	let xml = '<dataValidation';
	// "any" is the absence of a type: such an entry only shows its input message.
	if (validation.type && validation.type !== 'any') {
		xml += ' type="' + attr(validation.type) + '"';
	}
	// "between" is what an entry without an operator does.
	if (validation.operator && validation.operator !== 'between') {
		xml += ' operator="' + attr(validation.operator) + '"';
	}
	for (const flag of ['allowBlank', 'showInputMessage', 'showErrorMessage']) {
		if (validation[flag]) {
			xml += ' ' + flag + '="1"';
		}
	}
	for (const text of ['errorStyle', 'errorTitle', 'error', 'promptTitle', 'prompt']) {
		if (validation[text] !== undefined && validation[text] !== null) {
			xml += ' ' + text + '="' + attr(validation[text]) + '"';
		}
	}
	xml += ' sqref="' + attr(validation.sqref) + '"';
	const formulae = validation.formulae || [];
	if (!formulae.length) {
		return xml + '/>';
	}
	return xml + '>' + formulae.slice(0, 2).map((formula, index) =>
		'<formula' + (index + 1) + '>' + formulaText(formula) + '</formula' + (index + 1) + '>').join('') + '</dataValidation>';
}

/** The notes of a sheet: the comments part, which holds the texts. */
function commentsXml(notes) {
	return XML_HEAD + '<comments xmlns="' + MAIN_NS + '"><authors><author>Author</author></authors><commentList>' +
		notes.map((note) => '<comment ref="' + note.ref + '" authorId="0"><text><r>' + textElement(note.text) + '</r></text></comment>').join('') +
		'</commentList></comments>';
}

/**
 * The notes of a sheet: the legacy drawing, which holds the boxes Excel shows them in. Excel shows no
 * note without one. Each box is hidden until the cell is pointed at, and is anchored just right of its
 * cell, two columns wide and four rows high.
 */
function vmlXml(notes) {
	return XML_HEAD + '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
		'<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>' +
		'<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">' +
		'<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>' +
		notes.map((note, index) => {
			const top = Math.max(note.row - 2, 0);
			const anchor = [note.column, 6, top, 14, note.column + 2, 2, top + 4, 16].join(', ');
			return '<v:shape id="_x0000_s' + (1025 + index) + '" type="#_x0000_t202" ' +
				'style="position:absolute; margin-left:105.3pt;margin-top:10.5pt;width:97.8pt;height:59.1pt;z-index:1;visibility:hidden" ' +
				'fillcolor="infoBackground [80]" strokecolor="none [81]" o:insetmode="auto">' +
				'<v:fill color2="infoBackground [80]"/><v:shadow color="none [81]" obscured="t"/><v:path o:connecttype="none"/>' +
				'<v:textbox style="mso-direction-alt:auto" inset="1.3mm,1.3mm,2.5mm,2.5mm"><div style="text-align:left"/></v:textbox>' +
				'<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>' + anchor + '</x:Anchor>' +
				'<x:Locked>True</x:Locked><x:AutoFill>False</x:AutoFill><x:LockText>True</x:LockText>' +
				'<x:Row>' + (note.row - 1) + '</x:Row><x:Column>' + (note.column - 1) + '</x:Column></x:ClientData></v:shape>';
		}).join('') + '</xml>';
}

// ---------------------------------------------------------------------------
// The package
// ---------------------------------------------------------------------------

const relationshipsXml = (list) => XML_HEAD + '<Relationships xmlns="' + PACKAGE_REL_NS + '">' +
	list.map(([id, type, target]) => '<Relationship Id="' + id + '" Type="' + type + '" Target="' + target + '"/>').join('') +
	'</Relationships>';

/** A Date as a W3C date-time, to the second; null for an invalid one. */
function w3cDate(date) {
	return isDate(date) && !Number.isNaN(date.getTime()) ? date.toISOString().replace(/\.\d{3}Z$/, 'Z') : null;
}

function corePropertiesXml(workbook) {
	const dated = (tag, date) => (w3cDate(date) ? '<dcterms:' + tag + ' xsi:type="dcterms:W3CDTF">' + w3cDate(date) + '</dcterms:' + tag + '>' : '');
	return XML_HEAD + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
		'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
		'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
		(workbook.creator ? '<dc:creator>' + escapeXml(workbook.creator) + '</dc:creator>' : '') +
		dated('created', workbook.created) + dated('modified', workbook.modified) +
		'</cp:coreProperties>';
}

function appPropertiesXml(sheets) {
	return XML_HEAD + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' +
		'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
		'<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>' +
		'<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>' +
		'<vt:variant><vt:i4>' + sheets.length + '</vt:i4></vt:variant></vt:vector></HeadingPairs>' +
		'<TitlesOfParts><vt:vector size="' + sheets.length + '" baseType="lpstr">' +
		sheets.map((sheet) => '<vt:lpstr>' + escapeXml(sheet.name) + '</vt:lpstr>').join('') + '</vt:vector></TitlesOfParts>' +
		'<LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged>' +
		'<AppVersion>16.0300</AppVersion></Properties>';
}

/** A sheet name in a reference, always quoted: a quote inside it is doubled. */
const quoteSheetName = (name) => '\'' + String(name).replace(/'/g, '\'\'') + '\'';

/**
 * Writes a workbook.
 *
 * @param {object} workbook The workbook, as the header of this file describes it.
 * @param {object} [options]
 * @param {boolean} [options.compress=true] Deflate the parts; false stores them as they are.
 * @returns {Promise<ArrayBuffer>} The .xlsx file.
 */
export async function writeXlsx(workbook, options = {}) {
	const sheets = workbook.sheets || [];
	if (sheets.length === 0) {
		throw new Error('A workbook needs at least one sheet.');
	}
	const names = new Set();
	for (const sheet of sheets) {
		const name = String(sheet.name || '');
		if (!name || names.has(name.toLowerCase())) {
			throw new Error('Every sheet needs a name of its own; "' + name + '" is empty or taken.');
		}
		names.add(name.toLowerCase());
	}
	const activeTab = Math.max(sheets.findIndex((sheet) => !sheet.state || sheet.state === 'visible'), 0);

	const styles = createStyles();
	const strings = createSharedStrings();
	const parts = [];
	const overrides = [];
	const sheetEntries = [];
	let hasNotes = false;
	sheets.forEach((sheet, at) => {
		const number = at + 1;
		const { xml, notes } = worksheetXml(sheet, at === activeTab, styles, strings);
		sheetEntries.push(['xl/worksheets/sheet' + number + '.xml', xml]);
		overrides.push(['/xl/worksheets/sheet' + number + '.xml', CONTENT_TYPE + 'worksheet+xml']);
		if (notes.length) {
			hasNotes = true;
			sheetEntries.push(
				['xl/worksheets/_rels/sheet' + number + '.xml.rels', relationshipsXml([
					['rId1', REL_NS + '/comments', '../comments' + number + '.xml'],
					['rId2', REL_NS + '/vmlDrawing', '../drawings/vmlDrawing' + number + '.vml']
				])],
				['xl/comments' + number + '.xml', commentsXml(notes)],
				['xl/drawings/vmlDrawing' + number + '.vml', vmlXml(notes)]
			);
			overrides.push(['/xl/comments' + number + '.xml', CONTENT_TYPE + 'comments+xml']);
		}
	});

	const workbookRelationships = sheets.map((sheet, at) => ['rId' + (at + 1), REL_NS + '/worksheet', 'worksheets/sheet' + (at + 1) + '.xml']);
	workbookRelationships.push(['rId' + (sheets.length + 1), REL_NS + '/styles', 'styles.xml']);
	if (strings.size) {
		workbookRelationships.push(['rId' + (sheets.length + 2), REL_NS + '/sharedStrings', 'sharedStrings.xml']);
		overrides.push(['/xl/sharedStrings.xml', CONTENT_TYPE + 'sharedStrings+xml']);
	}
	const definedNames = workbook.definedNames || [];
	const workbookXml = XML_HEAD + '<workbook xmlns="' + MAIN_NS + '" xmlns:r="' + REL_NS + '">' +
		'<bookViews><workbookView activeTab="' + activeTab + '"/></bookViews>' +
		'<sheets>' + sheets.map((sheet, at) => '<sheet name="' + attr(sheet.name) + '" sheetId="' + (at + 1) + '"' +
			(sheet.state && sheet.state !== 'visible' ? ' state="' + attr(sheet.state) + '"' : '') + ' r:id="rId' + (at + 1) + '"/>').join('') + '</sheets>' +
		(definedNames.length
			? '<definedNames>' + definedNames.map((entry) => '<definedName name="' + attr(entry.name) + '">' +
				escapeXml(quoteSheetName(entry.sheet) + '!' + entry.range) + '</definedName>').join('') + '</definedNames>'
			: '') +
		'<calcPr calcId="171027"/></workbook>';

	overrides.unshift(
		['/xl/workbook.xml', CONTENT_TYPE + 'sheet.main+xml'],
		['/xl/styles.xml', CONTENT_TYPE + 'styles+xml'],
		['/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'],
		['/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml']
	);
	const contentTypesXml = XML_HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
		'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
		'<Default Extension="xml" ContentType="application/xml"/>' +
		(hasNotes ? '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' : '') +
		overrides.map(([name, type]) => '<Override PartName="' + name + '" ContentType="' + type + '"/>').join('') + '</Types>';

	// The worksheets come first: they fill the styles and the shared strings.
	parts.push(
		['[Content_Types].xml', contentTypesXml],
		['_rels/.rels', relationshipsXml([
			['rId1', REL_NS + '/officeDocument', 'xl/workbook.xml'],
			['rId2', PACKAGE_REL_NS + '/metadata/core-properties', 'docProps/core.xml'],
			['rId3', REL_NS + '/extended-properties', 'docProps/app.xml']
		])],
		['docProps/core.xml', corePropertiesXml(workbook)],
		['docProps/app.xml', appPropertiesXml(sheets)],
		['xl/workbook.xml', workbookXml],
		['xl/_rels/workbook.xml.rels', relationshipsXml(workbookRelationships)],
		['xl/styles.xml', styles.xml()],
		...(strings.size ? [['xl/sharedStrings.xml', strings.xml()]] : []),
		...sheetEntries
	);
	const file = await zip(parts.map(([name, data]) => ({ name, data })), options);
	return file.buffer;
}

// ---------------------------------------------------------------------------
// The zip container
// ---------------------------------------------------------------------------

const CRC_TABLE = (() => {
	const table = new Uint32Array(256);
	for (let n = 0; n < 256; n++) {
		let c = n;
		for (let k = 0; k < 8; k++) {
			c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
		}
		table[n] = c >>> 0;
	}
	return table;
})();

/** The CRC-32 of some bytes, as zip records it. */
export function crc32(bytes) {
	let crc = 0xffffffff;
	for (let i = 0; i < bytes.length; i++) {
		crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
	}
	return (crc ^ 0xffffffff) >>> 0;
}

/** Raw deflate, as zip stores it, through the platform's CompressionStream. */
async function deflateRaw(bytes) {
	const stream = new CompressionStream('deflate-raw');
	const writer = stream.writable.getWriter();
	// Not awaited before reading: the stream holds back the write until its output is read.
	const writing = writer.write(bytes).then(() => writer.close());
	const reader = stream.readable.getReader();
	const chunks = [];
	let length = 0;
	for (;;) {
		const { done, value } = await reader.read();
		if (done) {
			break;
		}
		chunks.push(value);
		length += value.length;
	}
	await writing;
	return concat(chunks, length);
}

function concat(chunks, length) {
	const out = new Uint8Array(length);
	let offset = 0;
	for (const chunk of chunks) {
		out.set(chunk, offset);
		offset += chunk.length;
	}
	return out;
}

/**
 * A zip file of the given entries, each deflated, or stored when deflating does not make it
 * smaller or the platform has no CompressionStream. Names are UTF-8, and every entry is dated
 * 1 January 1980, the earliest date zip can record, so the same input gives the same file.
 *
 * @param {Array<{name: string, data: (string|Uint8Array)}>} entries
 * @param {{compress: (boolean|undefined)}} [options]
 * @returns {Promise<Uint8Array>}
 */
export async function zip(entries, { compress = true } = {}) {
	const deflate = compress && typeof CompressionStream === 'function';
	const files = await Promise.all(entries.map(async ({ name, data }) => {
		const bytes = typeof data === 'string' ? encoder.encode(data) : data;
		const deflated = deflate ? await deflateRaw(bytes) : null;
		const stored = !deflated || deflated.length >= bytes.length;
		return { name: encoder.encode(name), size: bytes.length, crc: crc32(bytes), method: stored ? 0 : 8, body: stored ? bytes : deflated };
	}));
	if (files.length > 0xffff) {
		throw new Error('A zip file without zip64 holds at most 65,535 entries.');
	}
	const DOS_DATE = (0 << 9) | (1 << 5) | 1;
	const chunks = [];
	const central = [];
	let offset = 0;
	for (const file of files) {
		const local = new DataView(new ArrayBuffer(30));
		local.setUint32(0, 0x04034b50, true);
		local.setUint16(4, 20, true);
		local.setUint16(6, 0x0800, true);
		local.setUint16(8, file.method, true);
		local.setUint16(10, 0, true);
		local.setUint16(12, DOS_DATE, true);
		local.setUint32(14, file.crc, true);
		local.setUint32(18, file.body.length, true);
		local.setUint32(22, file.size, true);
		local.setUint16(26, file.name.length, true);
		local.setUint16(28, 0, true);
		const header = new DataView(new ArrayBuffer(46));
		header.setUint32(0, 0x02014b50, true);
		header.setUint16(4, 20, true);
		header.setUint16(6, 20, true);
		header.setUint16(8, 0x0800, true);
		header.setUint16(10, file.method, true);
		header.setUint16(12, 0, true);
		header.setUint16(14, DOS_DATE, true);
		header.setUint32(16, file.crc, true);
		header.setUint32(20, file.body.length, true);
		header.setUint32(24, file.size, true);
		header.setUint16(28, file.name.length, true);
		header.setUint32(42, offset, true);
		chunks.push(new Uint8Array(local.buffer), file.name, file.body);
		central.push(new Uint8Array(header.buffer), file.name);
		offset += 30 + file.name.length + file.body.length;
	}
	const centralSize = central.reduce((sum, chunk) => sum + chunk.length, 0);
	if (offset + centralSize > 0xffffffff) {
		throw new Error('A zip file without zip64 holds at most 4 GB.');
	}
	const end = new DataView(new ArrayBuffer(22));
	end.setUint32(0, 0x06054b50, true);
	end.setUint16(8, files.length, true);
	end.setUint16(10, files.length, true);
	end.setUint32(12, centralSize, true);
	end.setUint32(16, offset, true);
	const all = [...chunks, ...central, new Uint8Array(end.buffer)];
	return concat(all, offset + centralSize + 22);
}
