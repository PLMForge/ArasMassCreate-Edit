/**
 * Mass Create & Edit - reading an .xlsx file.
 *
 * openXlsx(bytes) opens a workbook and reads its sheets into plain cell values:
 * a string, a number, a boolean, a Date, or the text of an error ('#N/A').
 * Rich text comes back as its text, a formula as the result Excel saved with
 * it, a link as the text it shows. Nothing else of the file is read: no
 * styles beyond the number formats, no validation, no comments.
 *
 * It needs nothing but the platform: the zip container is read here, its
 * parts are inflated with DecompressionStream, and the XML is read by a small
 * tokenizer of its own (a browser has DOMParser, node has not, and one reader
 * for both keeps the two the same).
 *
 * Only what is asked for is inflated: opening a workbook reads its list of
 * sheets, and a sheet is read when it is asked for. A file that is not the
 * workbook wanted is turned away without inflating its sheets.
 *
 * Dates follow the workbook module: a Date whose UTC fields hold the
 * wall-clock value of the cell, whatever time zone the browser is in.
 */

// The uncompressed size of all the parts read from one file, at most. Workbooks compress four-
// to ninefold, so a legitimate upload of 10 MB inflates to under 100 MB; this leaves room for
// five times that, and stops a small file that inflates to gigabytes (a "zip bomb").
const MAX_UNCOMPRESSED = 512 * 1024 * 1024;

const ZIP_END = 0x06054b50;
const ZIP_CENTRAL = 0x02014b50;
const ZIP_LOCAL = 0x04034b50;
const ZIP_END_LENGTH = 22;
const ZIP_COMMENT_MAX = 0xffff;
// The first bytes of an OLE compound file: an encrypted (password-protected) workbook, or an old .xls.
const COMPOUND_FILE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

const EXCEL_LAST_ROW = 1048576;
const EXCEL_LAST_COLUMN = 16384;

// Built-in number formats that show a date or a time: 14-22 and 45-47 everywhere, 27-36 and
// 50-58 in the East Asian versions of Excel.
const BUILT_IN_DATE_FORMATS = new Set([
	14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47,
	50, 51, 52, 53, 54, 55, 56, 57, 58
]);

function xlsxError(code, message, cause) {
	const error = new Error(message);
	error.code = code;
	if (cause !== undefined) {
		error.cause = cause;
	}
	return error;
}

const damaged = (what) => xlsxError('XLSX_DAMAGED', 'The file is damaged: ' + what + '.');

// ---------------------------------------------------------------------------
// The zip container
// ---------------------------------------------------------------------------

/**
 * The bytes of an ArrayBuffer or a typed array. Neither `instanceof` nor the
 * constructor's identity is used: a buffer made in another window (a frame)
 * is just as good.
 */
function bytesOf(data) {
	if (ArrayBuffer.isView(data)) {
		return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
	}
	const tag = Object.prototype.toString.call(data);
	if (tag === '[object ArrayBuffer]' || tag === '[object SharedArrayBuffer]') {
		return new Uint8Array(data);
	}
	throw xlsxError('XLSX_NOT_ZIP', 'The file is not an Excel workbook (.xlsx).');
}

const startsWith = (bytes, signature) => signature.every((byte, index) => bytes[index] === byte);

/**
 * The entries of a zip file, by lower-case name without a leading slash.
 *
 * The central directory is the table of contents, and the only place whose
 * sizes can be trusted: a local header may say 0 and leave the sizes to a
 * descriptor after the data.
 */
export function readZip(data) {
	const bytes = bytesOf(data);
	if (startsWith(bytes, COMPOUND_FILE)) {
		throw xlsxError(
			'XLSX_ENCRYPTED',
			'The file is password-protected, or it is an old Excel 97-2003 workbook (.xls). Save it as an Excel workbook (.xlsx) without a password.'
		);
	}
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	const notZip = () => xlsxError('XLSX_NOT_ZIP', 'The file is not an Excel workbook (.xlsx).');
	if (bytes.length < ZIP_END_LENGTH) {
		throw notZip();
	}
	// The end record sits at the very end, followed only by a comment of up to 64 KB. The comment
	// may hold the record's signature too: a record counts when the table it points to is there.
	const zip64 = () => xlsxError('XLSX_UNSUPPORTED', 'The file is a ZIP64 archive, which an Excel workbook never is.');
	let end = -1;
	let found = null;
	const lowest = Math.max(0, bytes.length - ZIP_END_LENGTH - ZIP_COMMENT_MAX);
	for (let at = bytes.length - ZIP_END_LENGTH; at >= lowest && end < 0; at--) {
		if (bytes[at] !== 0x50 || view.getUint32(at, true) !== ZIP_END) {
			continue;
		}
		const size = view.getUint32(at + 12, true);
		const start = view.getUint32(at + 16, true);
		const isZip64 = view.getUint16(at + 10, true) === 0xffff || size === 0xffffffff || start === 0xffffffff;
		found = found || (isZip64 ? 'zip64' : 'record');
		if (!isZip64 && start + size <= at && (size === 0 || view.getUint32(start, true) === ZIP_CENTRAL)) {
			end = at;
		}
	}
	if (end < 0) {
		throw found === 'zip64' ? zip64() : found ? damaged('its table of contents is cut off') : notZip();
	}
	const count = view.getUint16(end + 10, true);
	const directoryOffset = view.getUint32(end + 16, true);
	const decoder = new TextDecoder('utf-8');
	const entries = new Map();
	let offset = directoryOffset;
	for (let index = 0; index < count; index++) {
		if (offset + 46 > end || view.getUint32(offset, true) !== ZIP_CENTRAL) {
			throw damaged('its table of contents is not readable');
		}
		const flags = view.getUint16(offset + 8, true);
		const method = view.getUint16(offset + 10, true);
		const compressedSize = view.getUint32(offset + 20, true);
		const size = view.getUint32(offset + 24, true);
		const nameLength = view.getUint16(offset + 28, true);
		const extraLength = view.getUint16(offset + 30, true);
		const commentLength = view.getUint16(offset + 32, true);
		const localOffset = view.getUint32(offset + 42, true);
		const name = decoder.decode(bytes.subarray(offset + 46, offset + 46 + nameLength));
		offset += 46 + nameLength + extraLength + commentLength;
		if (flags & 1) {
			throw xlsxError('XLSX_ENCRYPTED', 'The file is encrypted. Save it as an Excel workbook (.xlsx) without a password.');
		}
		if (compressedSize === 0xffffffff || size === 0xffffffff || localOffset === 0xffffffff) {
			throw zip64();
		}
		const key = name.replace(/^\/+/, '').toLowerCase();
		if (!key || key.endsWith('/') || entries.has(key)) {
			// A folder has no content; of two entries with the same name, the first one counts.
			continue;
		}
		entries.set(key, { name, method, compressedSize, size, localOffset });
	}
	return { bytes, view, entries };
}

/** The stored bytes of an entry, found through its local header. */
function storedBytes(zip, entry) {
	const { bytes, view } = zip;
	const at = entry.localOffset;
	if (at + 30 > bytes.length || view.getUint32(at, true) !== ZIP_LOCAL) {
		throw damaged('part ' + entry.name + ' is not where the table of contents says');
	}
	const start = at + 30 + view.getUint16(at + 26, true) + view.getUint16(at + 28, true);
	if (start + entry.compressedSize > bytes.length) {
		throw damaged('part ' + entry.name + ' is cut off');
	}
	return bytes.subarray(start, start + entry.compressedSize);
}

/**
 * Inflates raw deflate data into text, stopping as soon as it grows past
 * `budget` bytes: the size a zip claims is not believed.
 */
async function inflateText(data, budget, name, limit) {
	const stream = new DecompressionStream('deflate-raw');
	const writer = stream.writable.getWriter();
	// The reader below reports any failure; these only keep it from being reported twice.
	writer.write(data).catch(() => {});
	writer.close().catch(() => {});
	const reader = stream.readable.getReader();
	let decoder = null;
	let total = 0;
	let text = '';
	try {
		for (;;) {
			const { done, value } = await reader.read();
			if (done) {
				break;
			}
			total += value.length;
			if (total > budget) {
				reader.cancel().catch(() => {});
				throw tooLarge(limit);
			}
			decoder = decoder || decoderFor(value);
			text += decoder.decode(value, { stream: true });
		}
		return { text: decoder ? text + decoder.decode() : '', size: total };
	} catch (error) {
		if (error && typeof error.code === 'string' && error.code.startsWith('XLSX_')) {
			throw error;
		}
		// A string as long as the part would be is more than the engine can make.
		if (error instanceof RangeError) {
			throw tooLarge(limit);
		}
		throw xlsxError('XLSX_DAMAGED', 'The file is damaged: part ' + name + ' can\'t be unpacked.', error);
	}
}

const tooLarge = (limit) => xlsxError(
	'XLSX_TOO_LARGE',
	'The workbook unpacks to more than ' + Math.round(limit / (1024 * 1024)) + ' MB, more than a workbook of this tool can hold.'
);

/** A decoder for XML that starts with these bytes: UTF-8, unless a UTF-16 byte order mark says otherwise. */
function decoderFor(first) {
	if (first.length >= 2 && first[0] === 0xff && first[1] === 0xfe) {
		return new TextDecoder('utf-16le');
	}
	if (first.length >= 2 && first[0] === 0xfe && first[1] === 0xff) {
		return new TextDecoder('utf-16be');
	}
	return new TextDecoder('utf-8');
}

/** The parts of a zip file as text, read when asked for and counted against the size limit. */
function partReader(data, limit) {
	const zip = readZip(data);
	let budget = limit;
	const find = (name) => zip.entries.get(String(name).replace(/^\/+/, '').toLowerCase()) || null;
	return {
		has: (name) => find(name) !== null,
		async text(name) {
			const entry = find(name);
			if (!entry) {
				return null;
			}
			if (entry.size > budget) {
				throw tooLarge(limit);
			}
			const stored = storedBytes(zip, entry);
			let result;
			if (entry.method === 0) {
				result = { text: decoderFor(stored).decode(stored), size: stored.length };
			} else if (entry.method === 8) {
				result = await inflateText(stored, budget, entry.name, limit);
			} else {
				throw xlsxError('XLSX_UNSUPPORTED', 'Part ' + entry.name + ' is packed with method ' + entry.method + ', which Excel does not use.');
			}
			budget -= result.size;
			return result.text;
		}
	};
}

// ---------------------------------------------------------------------------
// XML
// ---------------------------------------------------------------------------

const NAMED_ENTITIES = { lt: '<', gt: '>', amp: '&', quot: '"', apos: '\'' };

/** Text with its character and entity references replaced. An unknown or invalid reference is kept as it is. */
export function decodeEntities(text) {
	if (text.indexOf('&') < 0) {
		return text;
	}
	return text.replace(/&(#[xX][0-9a-fA-F]+|#[0-9]+|[A-Za-z][A-Za-z0-9]*);/g, (whole, reference) => {
		if (reference[0] !== '#') {
			return Object.prototype.hasOwnProperty.call(NAMED_ENTITIES, reference) ? NAMED_ENTITIES[reference] : whole;
		}
		const code = reference[1] === 'x' || reference[1] === 'X' ? parseInt(reference.slice(2), 16) : parseInt(reference.slice(1), 10);
		return code <= 0x10ffff ? String.fromCodePoint(code) : whole;
	});
}

/**
 * Text with Excel's escapes of characters XML can't hold replaced: _x000D_
 * is a carriage return. _x005F_ is the escape of "_" itself, so "_x005F_x0041_"
 * stands for the text "_x0041_"; replacing from left to right in one pass
 * gets that right.
 */
export function unescapeExcel(text) {
	if (text.indexOf('_x') < 0) {
		return text;
	}
	return text.replace(/_x([0-9A-Fa-f]{4})_/g, (whole, hex) => String.fromCharCode(parseInt(hex, 16)));
}

const localName = (name) => {
	const colon = name.indexOf(':');
	return colon < 0 ? name : name.slice(colon + 1);
};

const isSpace = (code) => code === 32 || code === 9 || code === 10 || code === 13;
const NO_ATTRIBUTES = Object.freeze({});

/**
 * The attributes of a start tag, by local name, values decoded. Namespace
 * declarations are left out. A prefixed attribute wins over an unprefixed one
 * of the same local name (r:id over id).
 */
function readAttributes(xml, start, end) {
	const attributes = {};
	const prefixed = new Set();
	let at = start;
	while (at < end) {
		while (at < end && isSpace(xml.charCodeAt(at))) {
			at++;
		}
		if (at >= end) {
			break;
		}
		const nameStart = at;
		while (at < end) {
			const code = xml.charCodeAt(at);
			if (code === 61 /* = */ || isSpace(code)) {
				break;
			}
			at++;
		}
		const name = xml.slice(nameStart, at);
		while (at < end && isSpace(xml.charCodeAt(at))) {
			at++;
		}
		if (xml.charCodeAt(at) !== 61) {
			// An attribute without a value is not XML; skip the word.
			continue;
		}
		at++;
		while (at < end && isSpace(xml.charCodeAt(at))) {
			at++;
		}
		const quote = xml[at];
		if (quote !== '"' && quote !== '\'') {
			throw damaged('an attribute value is not quoted');
		}
		const close = xml.indexOf(quote, at + 1);
		if (close < 0 || close >= end) {
			throw damaged('an attribute value is not closed');
		}
		const value = decodeEntities(xml.slice(at + 1, close));
		at = close + 1;
		if (name === 'xmlns' || name.startsWith('xmlns:')) {
			continue;
		}
		const colon = name.indexOf(':');
		const local = colon < 0 ? name : name.slice(colon + 1);
		if (colon >= 0) {
			attributes[local] = value;
			prefixed.add(local);
		} else if (!prefixed.has(local)) {
			attributes[local] = value;
		}
	}
	return attributes;
}

/**
 * Reads an XML document from start to end and calls the handler for each
 * element and each piece of text:
 *   open(localName, attributes, selfClosing), close(localName), text(decoded).
 * Element names lose their prefix (x:c is c). Text is only decoded and handed
 * over while handler.collecting is true. Comments, processing instructions
 * and a DOCTYPE are skipped; CDATA is text. A ">" inside a quoted attribute
 * value does not end the tag.
 */
export function scanXml(xml, handler) {
	const length = xml.length;
	let at = xml.charCodeAt(0) === 0xfeff ? 1 : 0;
	while (at < length) {
		const open = xml.indexOf('<', at);
		const textEnd = open < 0 ? length : open;
		if (textEnd > at && handler.collecting) {
			handler.text(decodeEntities(xml.slice(at, textEnd)));
		}
		if (open < 0) {
			break;
		}
		const next = xml.charCodeAt(open + 1);
		if (next === 33 /* ! */) {
			if (xml.startsWith('<!--', open)) {
				const close = xml.indexOf('-->', open + 4);
				if (close < 0) {
					throw damaged('a comment is not closed');
				}
				at = close + 3;
			} else if (xml.startsWith('<![CDATA[', open)) {
				const close = xml.indexOf(']]>', open + 9);
				if (close < 0) {
					throw damaged('a CDATA section is not closed');
				}
				if (handler.collecting) {
					handler.text(xml.slice(open + 9, close));
				}
				at = close + 3;
			} else {
				at = skipDeclaration(xml, open);
			}
			continue;
		}
		if (next === 63 /* ? */) {
			const close = xml.indexOf('?>', open + 2);
			if (close < 0) {
				throw damaged('a processing instruction is not closed');
			}
			at = close + 2;
			continue;
		}
		if (next === 47 /* / */) {
			const close = xml.indexOf('>', open + 2);
			if (close < 0) {
				throw damaged('an end tag is not closed');
			}
			handler.close(localName(xml.slice(open + 2, close).trim()));
			at = close + 1;
			continue;
		}
		let index = open + 1;
		while (index < length) {
			const code = xml.charCodeAt(index);
			if (isSpace(code) || code === 47 || code === 62) {
				break;
			}
			index++;
		}
		const name = localName(xml.slice(open + 1, index));
		const attributesStart = index;
		for (; index < length; index++) {
			const code = xml.charCodeAt(index);
			if (code === 34 || code === 39) {
				const close = xml.indexOf(code === 34 ? '"' : '\'', index + 1);
				if (close < 0) {
					throw damaged('an attribute value is not closed');
				}
				index = close;
			} else if (code === 62) {
				break;
			}
		}
		if (index >= length) {
			throw damaged('a tag is not closed');
		}
		const selfClosing = xml.charCodeAt(index - 1) === 47;
		const attributesEnd = selfClosing ? index - 1 : index;
		const attributes = attributesEnd > attributesStart ? readAttributes(xml, attributesStart, attributesEnd) : NO_ATTRIBUTES;
		handler.open(name, attributes, selfClosing);
		if (selfClosing) {
			handler.close(name);
		}
		at = index + 1;
	}
}

/** The end of a <!DOCTYPE ...> or other declaration, internal subset in brackets included. */
function skipDeclaration(xml, open) {
	let depth = 0;
	for (let index = open + 2; index < xml.length; index++) {
		const code = xml.charCodeAt(index);
		if (code === 34 || code === 39) {
			const close = xml.indexOf(code === 34 ? '"' : '\'', index + 1);
			if (close < 0) {
				break;
			}
			index = close;
		} else if (code === 91 /* [ */) {
			depth++;
		} else if (code === 93 /* ] */) {
			depth--;
		} else if (code === 62 && depth <= 0) {
			return index + 1;
		}
	}
	throw damaged('a declaration is not closed');
}

/** Calls `onElement(localName, attributes)` for each start tag. */
function eachElement(xml, onElement) {
	scanXml(xml, { collecting: false, open: onElement, close() {}, text() {} });
}

// ---------------------------------------------------------------------------
// The package: which part is the workbook, which parts are its sheets
// ---------------------------------------------------------------------------

/** The relationships of a .rels part: { id, type, target } each, external targets left out. */
function readRelationships(xml) {
	const relationships = [];
	if (xml) {
		eachElement(xml, (name, attributes) => {
			if (name === 'Relationship' && attributes.Target !== undefined && attributes.TargetMode !== 'External') {
				relationships.push({ id: attributes.Id, type: attributes.Type || '', target: attributes.Target });
			}
		});
	}
	return relationships;
}

// Transitional and Strict files name the same relationship types under different
// namespaces (schemas.openxmlformats.org, purl.oclc.org); the last segment is the same.
const typeIs = (relationship, kind) => relationship.type.endsWith('/' + kind);

/** A relationship target as a part name, relative to the part `from`. */
function resolveTarget(from, target) {
	let path = target.split('#')[0];
	try {
		path = decodeURI(path);
	} catch (ignored) {
		// Not percent-encoded after all: the target names the part as it is.
	}
	const parts = path.startsWith('/') ? [] : from.split('/').slice(0, -1);
	for (const segment of path.split('/')) {
		if (segment === '..') {
			parts.pop();
		} else if (segment !== '.' && segment !== '') {
			parts.push(segment);
		}
	}
	return parts.join('/');
}

const relsPartOf = (part) => {
	const slash = part.lastIndexOf('/');
	return part.slice(0, slash + 1) + '_rels/' + part.slice(slash + 1) + '.rels';
};

/**
 * Opens an .xlsx file: reads its list of sheets, and nothing more until a
 * sheet is asked for.
 *
 * @param {ArrayBuffer|ArrayBufferView} data The file, from this window or another.
 * @param {{maxUncompressed?: number}} [options] The most bytes the parts read may unpack to, in all.
 * @returns {Promise<{date1904: boolean, sheetNames: string[], hasSheet: function(string): boolean,
 *   readSheet: function(string): Promise<?XlsxSheet>}>} hasSheet is true for a sheet that has
 *   cells (a worksheet whose part is in the file); readSheet gives null for any other name.
 * @throws {Error} With code XLSX_NOT_ZIP, XLSX_ENCRYPTED, XLSX_UNSUPPORTED, XLSX_DAMAGED,
 *   XLSX_TOO_LARGE or XLSX_NOT_WORKBOOK. readSheet throws the same.
 */
export async function openXlsx(data, options = {}) {
	const parts = partReader(data, options.maxUncompressed || MAX_UNCOMPRESSED);
	const rootRelationships = readRelationships(await parts.text('_rels/.rels'));
	const document = rootRelationships.find((relationship) => typeIs(relationship, 'officeDocument'));
	const workbookPart = document ? resolveTarget('', document.target) : 'xl/workbook.xml';
	const workbookXml = await parts.text(workbookPart);
	if (workbookXml === null) {
		throw xlsxError('XLSX_NOT_WORKBOOK', 'The file is not an Excel workbook (.xlsx): it has no workbook part.');
	}

	const relationships = readRelationships(await parts.text(relsPartOf(workbookPart)));
	const byId = new Map(relationships.map((relationship) => [relationship.id, relationship]));
	const partOf = (kind, fallback) => {
		const found = relationships.find((relationship) => typeIs(relationship, kind));
		return found ? resolveTarget(workbookPart, found.target) : fallback;
	};

	let isWorkbook = false;
	let date1904 = false;
	const sheets = new Map();
	eachElement(workbookXml, (name, attributes) => {
		if (name === 'workbook') {
			isWorkbook = true;
		} else if (name === 'workbookPr') {
			date1904 = attributes.date1904 === '1' || attributes.date1904 === 'true';
		} else if (name === 'sheet' && attributes.name !== undefined && !sheets.has(attributes.name)) {
			const relationship = byId.get(attributes.id);
			// A chart sheet or a dialog sheet has no cells.
			const part = relationship && typeIs(relationship, 'worksheet') ? resolveTarget(workbookPart, relationship.target) : null;
			sheets.set(attributes.name, part && parts.has(part) ? part : null);
		}
	});
	if (!isWorkbook) {
		throw xlsxError('XLSX_NOT_WORKBOOK', 'The file is not an Excel workbook (.xlsx).');
	}

	// What all sheets share, read with the first sheet that needs it.
	let shared = null;
	const sharedParts = () => shared || (shared = (async () => ({
		strings: readSharedStrings(await parts.text(partOf('sharedStrings', 'xl/sharedStrings.xml'))),
		dateStyles: readDateStyles(await parts.text(partOf('styles', 'xl/styles.xml')))
	}))());

	return {
		date1904,
		sheetNames: Array.from(sheets.keys()),
		hasSheet: (name) => !!sheets.get(name),
		async readSheet(name) {
			const part = sheets.get(name);
			if (!part) {
				return null;
			}
			const { strings, dateStyles } = await sharedParts();
			return readSheetXml(await parts.text(part), { strings, dateStyles, date1904 });
		}
	};
}

// ---------------------------------------------------------------------------
// Shared strings and number formats
// ---------------------------------------------------------------------------

/**
 * The shared strings, in order. A string is the text of its <t> elements,
 * plain or in rich-text runs; a phonetic guide (<rPh>, the reading of
 * Japanese text) is not part of it. Spaces are kept.
 */
function readSharedStrings(xml) {
	const strings = [];
	if (!xml) {
		return strings;
	}
	let text = null;
	let piece = '';
	let phonetic = 0;
	const handler = {
		collecting: false,
		open(name) {
			if (name === 'si') {
				text = '';
			} else if (name === 'rPh') {
				phonetic++;
			} else if (name === 't' && text !== null && phonetic === 0) {
				handler.collecting = true;
				piece = '';
			}
		},
		close(name) {
			if (name === 't' && handler.collecting) {
				handler.collecting = false;
				text += unescapeExcel(piece);
			} else if (name === 'rPh') {
				phonetic--;
			} else if (name === 'si' && text !== null) {
				strings.push(text);
				text = null;
			}
		},
		text(value) {
			piece += value;
		}
	};
	scanXml(xml, handler);
	return strings;
}

/**
 * Whether a number format shows a date or a time. Quoted text, escaped
 * characters (\x), spacing (_x), fill (*x) and bracketed parts ([Red],
 * [$-409]) are not looked at, except the elapsed-time parts [h], [m], [s].
 */
export function isDateFormat(code) {
	let bare = '';
	for (let index = 0; index < code.length; index++) {
		const char = code[index];
		if (char === '"') {
			const close = code.indexOf('"', index + 1);
			index = close < 0 ? code.length : close;
		} else if (char === '\\' || char === '_' || char === '*') {
			index++;
		} else if (char === '[') {
			const close = code.indexOf(']', index + 1);
			const inside = close < 0 ? '' : code.slice(index + 1, close);
			if (/^(h+|m+|s+)$/i.test(inside)) {
				bare += inside;
			}
			index = close < 0 ? code.length : close;
		} else {
			bare += char;
		}
	}
	return /[ymdhs]/i.test(bare);
}

/** For each cell format (cellXfs, the index a cell's s attribute names), whether it shows a date. */
function readDateStyles(xml) {
	const styles = [];
	if (!xml) {
		return styles;
	}
	const custom = new Map();
	const formatIds = [];
	let inCellFormats = false;
	let depth = 0;
	scanXml(xml, {
		collecting: false,
		open(name, attributes, selfClosing) {
			if (name === 'numFmt' && attributes.numFmtId !== undefined) {
				custom.set(Number(attributes.numFmtId), attributes.formatCode || '');
			} else if (name === 'cellXfs') {
				inCellFormats = true;
				depth = 0;
				return;
			}
			if (inCellFormats) {
				if (depth === 0 && name === 'xf') {
					formatIds.push(Number(attributes.numFmtId || 0));
				}
				if (!selfClosing) {
					depth++;
				}
			}
		},
		close(name) {
			if (name === 'cellXfs' && depth === 0) {
				inCellFormats = false;
			} else if (inCellFormats && depth > 0) {
				depth--;
			}
		},
		text() {}
	});
	for (const id of formatIds) {
		styles.push(custom.has(id) ? isDateFormat(custom.get(id)) : BUILT_IN_DATE_FORMATS.has(id));
	}
	return styles;
}

// ---------------------------------------------------------------------------
// Sheets
// ---------------------------------------------------------------------------

/** The row and the column of a reference such as "AB12"; either is 0 when it is missing. */
function readReference(reference) {
	let column = 0;
	let index = 0;
	for (; index < reference.length; index++) {
		const code = reference.charCodeAt(index) & ~32;
		if (code < 65 || code > 90) {
			break;
		}
		column = column * 26 + code - 64;
	}
	const row = index < reference.length ? parseInt(reference.slice(index).replace(/^\$/, ''), 10) : 0;
	return { row: row > 0 ? row : 0, column: column <= EXCEL_LAST_COLUMN ? column : 0 };
}

/** A serial number as a Date, rounded to the millisecond (1900 or 1904 date system). */
function serialToDate(serial, date1904) {
	return new Date(Math.round((serial - 25569 + (date1904 ? 1462 : 0)) * 86400000));
}

/** An ISO 8601 date of a t="d" cell as a Date of its wall-clock value, or null. */
function isoToDate(text) {
	const match = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(\.\d+)?)?)?(?:Z|[+-]\d{2}:?\d{2})?$/.exec(text.trim());
	if (!match) {
		return null;
	}
	const [year, month, day, hour, minute, second] = match.slice(1, 7).map((part) => Number(part || 0));
	const date = new Date(Date.UTC(year, month - 1, day, hour, minute, second, Math.round(Number(match[7] || 0) * 1000)));
	return Number.isNaN(date.getTime()) ? null : date;
}

/**
 * @typedef {object} XlsxSheet
 * @property {number} lastRow The number of the last row the sheet has.
 * @property {function(number, number): *} value The value of a cell by row and column (1-based), or null.
 * @property {function(): number[]} rowNumbers The rows that hold a value, in order.
 */

function readSheetXml(xml, context) {
	const { strings, dateStyles, date1904 } = context;
	const rows = new Map();
	const merges = [];
	let lastRow = 0;

	let inData = false;
	let rowNumber = 0;
	let columnNumber = 0;
	// The cell being read.
	let cellRow = 0;
	let cellColumn = 0;
	let type = 'n';
	let style = 0;
	let value = null;
	let inline = null;
	let phonetic = 0;
	let piece = '';

	const set = (row, column, cellValue) => {
		let cells = rows.get(row);
		if (!cells) {
			cells = [];
			rows.set(row, cells);
		}
		cells[column] = cellValue;
	};

	const handler = {
		collecting: false,
		open(name, attributes) {
			if (!inData) {
				if (name === 'sheetData') {
					inData = true;
				} else if (name === 'mergeCell' && attributes.ref) {
					const [first, last = first] = attributes.ref.split(':').map(readReference);
					if (first.row && first.column && last.row && last.column) {
						merges.push({
							top: Math.min(first.row, last.row), bottom: Math.max(first.row, last.row),
							left: Math.min(first.column, last.column), right: Math.max(first.column, last.column)
						});
					}
				}
				return;
			}
			switch (name) {
				case 'row': {
					const number = attributes.r ? parseInt(attributes.r, 10) : NaN;
					rowNumber = number > 0 ? number : rowNumber + 1;
					columnNumber = 0;
					lastRow = Math.max(lastRow, rowNumber);
					break;
				}
				case 'c': {
					const reference = attributes.r ? readReference(attributes.r) : { row: 0, column: 0 };
					cellColumn = reference.column || columnNumber + 1;
					cellRow = reference.row || rowNumber;
					columnNumber = cellColumn;
					type = attributes.t || 'n';
					style = attributes.s ? parseInt(attributes.s, 10) : 0;
					value = null;
					inline = null;
					break;
				}
				case 'v':
					handler.collecting = true;
					piece = '';
					break;
				case 'is':
					inline = '';
					break;
				case 'rPh':
					phonetic++;
					break;
				case 't':
					if (inline !== null && phonetic === 0) {
						handler.collecting = true;
						piece = '';
					}
					break;
				default:
			}
		},
		close(name) {
			if (!inData) {
				return;
			}
			switch (name) {
				case 'v':
					handler.collecting = false;
					value = piece;
					break;
				case 't':
					if (handler.collecting) {
						handler.collecting = false;
						inline += unescapeExcel(piece);
					}
					break;
				case 'rPh':
					phonetic--;
					break;
				case 'c': {
					const cellValue = cellValueOf(type, value, inline, style);
					if (cellValue !== null && cellRow > 0) {
						set(cellRow, cellColumn, cellValue);
						lastRow = Math.max(lastRow, cellRow);
					}
					break;
				}
				case 'sheetData':
					inData = false;
					break;
				default:
			}
		},
		text(text) {
			piece += text;
		}
	};

	function cellValueOf(cellType, raw, inlineText, styleIndex) {
		switch (cellType) {
			case 'inlineStr':
				return inlineText;
			case 's': {
				if (raw === null) {
					return null;
				}
				const found = strings[parseInt(raw, 10)];
				return found === undefined ? null : found;
			}
			case 'str':
				return raw === null ? null : unescapeExcel(raw);
			case 'b':
				return raw === null ? null : raw.trim() === '1' || raw.trim() === 'true';
			case 'e':
				return raw === null ? null : raw;
			case 'd':
				return raw === null ? null : isoToDate(raw) || raw;
			default: {
				if (raw === null || raw.trim() === '') {
					return null;
				}
				const number = Number(raw);
				if (!Number.isFinite(number)) {
					return null;
				}
				return dateStyles[styleIndex] ? serialToDate(number, date1904) : number;
			}
		}
	}

	scanXml(xml, handler);

	// Every cell of a merged range shows the value of its top-left cell.
	const mergeAt = (row, column) => {
		for (const merge of merges) {
			if (row >= merge.top && row <= merge.bottom && column >= merge.left && column <= merge.right) {
				return merge;
			}
		}
		return null;
	};
	const ownValue = (row, column) => {
		const cells = rows.get(row);
		const found = cells ? cells[column] : undefined;
		return found === undefined ? null : found;
	};
	const valueAt = merges.length === 0 ? ownValue : (row, column) => {
		const merge = mergeAt(row, column);
		return merge ? ownValue(merge.top, merge.left) : ownValue(row, column);
	};
	for (const merge of merges) {
		lastRow = Math.max(lastRow, Math.min(merge.bottom, EXCEL_LAST_ROW));
	}

	let numbers = null;
	return {
		lastRow,
		value: valueAt,
		rowNumbers() {
			if (!numbers) {
				const found = new Set(rows.keys());
				for (const merge of merges) {
					if (ownValue(merge.top, merge.left) !== null) {
						for (let row = merge.top; row <= Math.min(merge.bottom, EXCEL_LAST_ROW); row++) {
							found.add(row);
						}
					}
				}
				numbers = Array.from(found).sort((a, b) => a - b);
			}
			return numbers;
		}
	};
}
