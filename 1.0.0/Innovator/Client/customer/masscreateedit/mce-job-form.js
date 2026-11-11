/**
 * Mass Create & Edit - the tool on the job Form.
 *
 * The Form of a job (ItemType FRG_MassCreateEdit) has one HTML field with an
 * empty element, id "mce-app". attach() builds the tool in it, as an editor
 * that fills the job's item window, the way the Query Definition and MAC
 * Policy editors fill theirs:
 *
 *   toolbar    Generate workbook, Download workbook, Upload workbook, Import,
 *              and what is running now
 *   header     the job's state, target ItemType and files
 *   messages   what the last run said
 *   tabs       Workbook: what the workbook holds (the export section)
 *              Import: the preview of an upload, or the result of the import
 *
 * The Form's own fields are hidden: the header shows what they hold. Each tab
 * scrolls on its own, so the toolbar always stays in view.
 *
 * The export section shows the target ItemType, the key items are matched by,
 * the relationship sheets with their match keys, and the selected items. The
 * toolbar's Generate command generates the workbook.
 *
 * Generating a workbook takes these steps, in this order:
 *
 *   1. A Draft job that is being edited is saved and its editing is finished,
 *      with the item window's own Done command.
 *   2. Server Method FRG_MCE_ExportWorkbook, step "begin": checks that the
 *      user may generate, stores the export options chosen here (the relationship
 *      sheets and their match keys) on the job, issues a new workbook token,
 *      records what the workbook is generated from, and returns the template schema.
 *   3. The items to pre-fill are read here, with the user's own permissions.
 *   4. mce-workbook.js builds the file in the browser.
 *   5. The file is stored in the vault as a File item.
 *   6. Server Method FRG_MCE_ExportWorkbook, step "complete": attaches the
 *      File to the job and moves the job to Generated.
 *   7. The browser downloads the file, and the item window is refreshed.
 *
 * The import section uploads the filled-in workbook, shows what an import would
 * do in a preview grid beside an item preview, runs the import in batches, and
 * shows the result. Its steps are described at "The import section" below.
 *
 * Everything that changes the job after it is saved is done by the server
 * Methods. This module never writes a job property or a state itself, except
 * the target ItemType and the selected items (prefill_idlist) of a Draft job,
 * through the Form like a field does. The export options are kept in memory
 * until generating hands them to the server.
 *
 * The module is loaded once into the main window, not into the Form's frame,
 * so that generating goes on while the job is saved or refreshed. The Form
 * hands its window to attach() each time it is shown.
 */

const APP_ID = 'mce-app';
const HOST_ID = 'mce-export';
/** The start of the key in view.openKeys of a target whose selected items are all shown; the rest is its index. */
const ITEMS_KEY = 'items:';
// The column dialog of a sheet with more columns than this gets a box to find a column by its name.
const FILTER_FROM_COLUMNS = 12;
// The least height of the tool. A smaller item window scrolls the tool as a whole.
const MIN_TOOL_HEIGHT = 460;
const EXPORT_METHOD = 'FRG_MCE_ExportWorkbook';
const SCHEMA_METHOD = 'FRG_MCE_GetTemplateSchema';
const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const OPEN_STATES = ['Draft', 'Generated', 'Validated'];
const SCHEMA_KEPT_MS = 5 * 60 * 1000;
const MESSAGE_KEPT_MS = 10 * 1000;

// The other files of the tool are loaded with the cache tag this file was loaded with,
// so that a new version of the tool brings all of its files along.
const cacheTag = new URL(import.meta.url).search;
const sibling = (name) => new URL(name, import.meta.url).href.split('?')[0] + cacheTag;

let logicLoading = null;
let workbookLoading = null;
const loadLogic = () => logicLoading || (logicLoading = import(sibling('./mce-export.js')));
const loadWorkbook = () => workbookLoading || (workbookLoading = import(sibling('./mce-workbook.js')));
let importLogicLoading = null;
const loadImportLogic = () => importLogicLoading || (importLogicLoading = import(sibling('./mce-import.js')));
let md5Loading = null;
const loadMd5 = () => md5Loading || (md5Loading = import(sibling('./mce-md5.js')));

/** The functions of mce-export.js, once attach() has loaded them. */
let logic = null;

/**
 * What is remembered per job while the client is open:
 *   view     the Form that shows the job now
 *   flow     the state of generating: { kind: 'busy' | 'error' | 'done', message, at, notes, fileId }
 *   pending  export options chosen and not generated yet: { base, text }
 *   imp      the import: see importMemoryOf()
 *   names    the names of the job's selected items, by id: null for an item that can't be read
 *   tab      the tab the user chose, for the state the job was in then: { tab, state }
 *   adding   true while the Workbook tab shows a new target whose ItemType is not chosen yet
 */
const jobs = new Map();

function memoryOf(jobId) {
	let memory = jobs.get(jobId);
	if (!memory) {
		memory = { view: null, flow: null, pending: null, imp: null, names: new Map(), tab: null, adding: false };
		jobs.set(jobId, memory);
	}
	return memory;
}

/** Template schemas by ItemType name, kept for a few minutes. */
const schemas = new Map();

// ---------------------------------------------------------------------------
// Talking to the server
// ---------------------------------------------------------------------------

const escapeXml = (text) =>
	String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');

/** What the server said when it refused a request. A request that fails is rejected with the request object. */
function faultOf(failure) {
	const answer = failure && typeof failure.responseText === 'string' ? failure.responseText : '';
	if (answer === '') {
		return { code: '', message: failure && failure.message ? failure.message : 'The server could not be reached.' };
	}
	const xml = new DOMParser().parseFromString(answer, 'text/xml');
	const pick = (name) => {
		const node = xml.getElementsByTagName(name)[0];
		return node ? node.textContent.trim() : '';
	};
	return { code: pick('faultcode'), message: pick('faultstring') || 'The server refused the request.', detail: pick('detail') };
}

/** Sends one request as the user and returns the Result element of the answer. */
async function send(top, request, action) {
	try {
		return await top.ArasModules.soap(request, { async: true, method: action });
	} catch (failure) {
		const fault = faultOf(failure);
		const error = new Error(fault.message);
		// The server reports "no items found" as a fault with code 0.
		error.noItems = fault.code === '0';
		error.fault = fault;
		throw error;
	}
}

/** An item of an AML answer as a plain object: { id, props: { name: { value, keyedName, item, restricted } } }. */
function toItem(node) {
	const item = { id: node.getAttribute('id') || '', props: {} };
	for (const child of Array.from(node.children)) {
		const nested = Array.from(child.children).find((candidate) => candidate.localName === 'Item') || null;
		item.props[child.localName] = {
			value: nested ? nested.getAttribute('id') || '' : child.textContent,
			keyedName: child.getAttribute('keyed_name') || '',
			item: nested ? toItem(nested) : null,
			// A value the user may not read comes back empty and marked as "not null".
			restricted: child.getAttribute('is_null') === '0'
		};
	}
	return item;
}

/** Runs a query as the user. Finding nothing gives an empty list. */
async function query(top, aml) {
	let result;
	try {
		result = await send(top, '<AML>' + aml + '</AML>', 'ApplyAML');
	} catch (error) {
		if (error.noItems) {
			return [];
		}
		throw error;
	}
	return Array.from(result.children).filter((node) => node.localName === 'Item').map(toItem);
}

/** Calls a server Method that answers with JSON. */
async function callMethod(top, name, inputs) {
	const body = Object.entries(inputs).map(([key, value]) => '<' + key + '>' + escapeXml(value) + '</' + key + '>').join('');
	const result = await send(top, '<Item type="Method" action="' + name + '">' + body + '</Item>', 'ApplyMethod');
	return JSON.parse(result.textContent);
}

const textOf = (item, name) => (item.props[name] ? String(item.props[name].value || '') : '');

/**
 * The configuration item (the one named Default, or else the oldest) and its allow-list:
 * allowed holds the ItemTypes the user may use, restricted those on the list for other
 * identities only.
 */
async function readConfiguration(top) {
	const exportLogic = await loadLogic();
	const configs = await query(top, '<Item type="FRG_MassCreateEdit_Config" action="get" select="name,max_rows,max_file_size_mb,batch_size,parallel_batches,max_attachment_mb,max_attachments_total_mb,max_attachments" orderBy="created_on"/>');
	if (configs.length === 0) {
		throw new Error('Mass Create & Edit has no configuration item, or you have no access to it. Ask an administrator.');
	}
	const config = configs.find((candidate) => textOf(candidate, 'name') === 'Default') || configs[0];
	const rows = await query(
		top,
		'<Item type="FRG_MassCreateEdit_AllowedType" action="get" select="related_id(name,label),identity_id"><source_id>' + config.id + '</source_id></Item>'
	);
	const listed = rows
		.filter((row) => row.props.related_id && row.props.related_id.item)
		.map((row) => {
			const itemType = row.props.related_id.item;
			return {
				itemType: { id: itemType.id, name: textOf(itemType, 'name'), label: textOf(itemType, 'label') || textOf(itemType, 'name') },
				identityId: textOf(row, 'identity_id'),
				identityName: row.props.identity_id ? row.props.identity_id.keyedName || '' : ''
			};
		});
	const limit = (name, otherwise) => (textOf(config, name) === '' ? otherwise : Number(textOf(config, name)) || 0);
	const { allowed, restricted } = exportLogic.allowedForUser(listed, String(top.aras.getIdentityList() || '').split(','));
	return {
		maxRows: Number(textOf(config, 'max_rows')) || 0,
		maxFileSizeMb: Number(textOf(config, 'max_file_size_mb')) || 0,
		batchSize: Number(textOf(config, 'batch_size')) || 0,
		parallelBatches: limit('parallel_batches', 4),
		// Empty when the configuration item does not set them: the defaults FRG_MCE_ValidateImport applies.
		maxAttachmentMb: limit('max_attachment_mb', 100),
		maxAttachmentsTotalMb: limit('max_attachments_total_mb', 2000),
		maxAttachments: limit('max_attachments', 1000),
		allowed,
		restricted
	};
}

/** The configuration, read once for each Form that is shown. */
function configOf(view) {
	if (!view.configLoading) {
		view.configLoading = readConfiguration(view.top);
		view.configLoading.catch(() => {
			view.configLoading = null;
		});
	}
	return view.configLoading;
}

/** The template schema of an ItemType with every relationship it has. */
function readSchema(top, itemTypeName) {
	const kept = schemas.get(itemTypeName);
	if (kept && Date.now() - kept.since < SCHEMA_KEPT_MS) {
		return kept.loading;
	}
	const loading = callMethod(top, SCHEMA_METHOD, { itemtype: itemTypeName, all_relationships: '1' });
	schemas.set(itemTypeName, { since: Date.now(), loading });
	loading.catch(() => schemas.delete(itemTypeName));
	return loading;
}

// ---------------------------------------------------------------------------
// The job as the Form holds it
// ---------------------------------------------------------------------------

function readJob(view) {
	const node = view.doc.item;
	const aras = view.top.aras;
	const child = (name) => Array.from(node.children).find((candidate) => candidate.localName === name) || null;
	const value = (name) => {
		const element = child(name);
		if (!element) {
			return '';
		}
		const nested = Array.from(element.children).find((candidate) => candidate.localName === 'Item');
		return nested ? nested.getAttribute('id') || '' : element.textContent.trim();
	};
	const keyedName = (name) => (child(name) ? child(name).getAttribute('keyed_name') || '' : '');
	const isNew = !!aras.isTempEx(node);
	const createdBy = value('created_by_id');
	const job = {
		id: node.getAttribute('id'),
		isNew,
		editing: !!view.doc.isEditMode,
		state: value('state') || 'Draft',
		number: value('item_number'),
		isCreator: isNew || createdBy === '' || createdBy === aras.getUserID(),
		creatorName: keyedName('created_by_id'),
		lockedByOther: !isNew && value('locked_by_id') !== '' && value('locked_by_id') !== aras.getUserID(),
		targetId: value('target_itemtype'),
		targetName: keyedName('target_itemtype'),
		// The items the workbook is pre-filled with: in order, each once.
		prefillIds: Array.from(new Set(value('prefill_idlist').split(',').map((id) => id.trim()).filter((id) => /^[0-9A-Fa-f]{32}$/.test(id)))),
		optionsText: value('export_options'),
		targetsText: value('targets'),
		exportFile: value('export_file'),
		uploadFile: value('upload_file'),
		resultsFile: value('results_file'),
		fileNames: { export: keyedName('export_file'), upload: keyedName('upload_file') },
		exportedOn: value('exported_on'),
		copiedFrom: value('copied_from') ? keyedName('copied_from') || value('copied_from') : '',
		counts: {
			total: value('qty_total'), created: value('qty_created'), updated: value('qty_updated'), unchanged: value('qty_unchanged'),
			failed: value('qty_failed'), skipped: value('qty_skipped')
		}
	};
	// The targets, in the order of their main sheets, each with the items it is pre-filled with. Until
	// mce-export.js is loaded, the job's target_itemtype and prefill_idlist stand for them.
	job.targets = logic
		? logic.parseTargets(job.targetsText, job.targetId, value('prefill_idlist'))
		: job.targetId ? [{ itemTypeId: job.targetId.toUpperCase(), prefillIds: job.prefillIds }] : [];
	job.prefillCount = job.targets.reduce((sum, target) => sum + target.prefillIds.length, 0);
	// Pending options belong to the job's first target: a target added or removed keeps them.
	job.base = job.targetId + '|' + job.optionsText;
	return job;
}

/** Changes a property of the job in the Form, the way a field of the Form does. Only for a job that is being edited. */
function setJobProperty(view, name, value, attributes) {
	const item = view.doc.item;
	const aras = view.window.aras;
	item.setAttribute('isDirty', '1');
	aras.setItemProperty(item, name, value, undefined, view.doc.itemTypeNd);
	for (const [attribute, text] of Object.entries(attributes || {})) {
		aras.setItemPropertyAttribute(item, name, attribute, text);
	}
	const action = item.getAttribute('action');
	if (action !== 'add' && action !== 'create') {
		item.setAttribute('action', 'update');
	}
	if (typeof view.doc.userChangeHandler === 'function') {
		view.doc.userChangeHandler(name, value);
	}
}

/**
 * Sets the job's targets in the Form: the first is the job's target_itemtype, with its items in
 * prefill_idlist, and the targets property lists them all when there are several. Only the
 * properties that change are set.
 */
function setTargets(view, targets) {
	const job = readJob(view);
	const first = targets[0] || null;
	const firstId = first ? first.itemTypeId : '';
	if (firstId.toUpperCase() !== job.targetId.toUpperCase()) {
		const itemType = first ? view.model.config.allowed.find((candidate) => candidate.id.toUpperCase() === firstId) : null;
		setJobProperty(view, 'target_itemtype', firstId, first ? { keyed_name: itemType ? itemType.name : '', type: 'ItemType' } : null);
	}
	const prefillText = first ? first.prefillIds.join(',') : '';
	if (prefillText !== job.prefillIds.join(',')) {
		setJobProperty(view, 'prefill_idlist', prefillText);
	}
	const targetsText = logic.serializeTargets(targets);
	if (targetsText !== job.targetsText) {
		setJobProperty(view, 'targets', targetsText);
	}
}

/** A copy of the job's targets that can be changed. */
const copyTargets = (job) => job.targets.map((target) => ({ itemTypeId: target.itemTypeId, prefillIds: target.prefillIds.slice() }));

/** What a target is called: its label, when the user may use it, else what is known of it. */
function targetLabel(model, index) {
	const section = model.sections[index] || null;
	const target = model.job.targets[index] || null;
	if (section && section.target) {
		return section.target.label;
	}
	const id = target ? target.itemTypeId : '';
	const restricted = model.config ? model.config.restricted.find((itemType) => itemType.id.toUpperCase() === id) : null;
	return restricted ? restricted.label : (index === 0 && model.job.targetName) || id;
}

/** Why the user can't generate a workbook for the job, or null when they can. */
function whyNot(model) {
	const { job } = model;
	if (!job.isCreator) {
		return 'Only the user who created a job can run it. This job was created by ' + (job.creatorName || 'another user') + '.';
	}
	if (!OPEN_STATES.includes(job.state)) {
		return 'The job is ' + job.state + '. Its workbook can no longer be generated; start a new job instead.';
	}
	if (job.lockedByOther) {
		return 'The job is locked by another user.';
	}
	return null;
}

// ---------------------------------------------------------------------------
// Drawing the section
// ---------------------------------------------------------------------------

/** Creates an element in the Form's document. A null child is skipped. */
function element(doc, tag, attributes, ...children) {
	const node = doc.createElement(tag);
	for (const [name, value] of Object.entries(attributes || {})) {
		if (value === false || value === null || value === undefined) {
			continue;
		}
		if (name === 'checked' || name === 'disabled' || name === 'selected') {
			node[name] = !!value;
		} else {
			node.setAttribute(name, String(value));
		}
	}
	for (const child of children) {
		if (child !== null && child !== undefined && child !== false) {
			node.append(child);
		}
	}
	return node;
}

// With several targets, the sheets of a target share a colour: the classes mce-target-0 to -5 in
// mce-job-form.css, the same colours as the workbook's tabs (TARGET_TAB_COLORS in mce-workbook.js).
const TARGET_COLOURS = 6;

/** The class that gives a sheet the colour of its target, at its position, or '' with a single target. */
const targetClass = (at, several) => (several && Number.isInteger(at) && at >= 0 ? ' mce-target-' + (at % TARGET_COLOURS) : '');

/** A small square in the colour of a target, put before a sheet's name; null with a single target. */
function targetMark(doc, at, several) {
	const colour = targetClass(at, several);
	return colour ? element(doc, 'span', { class: 'mce-target-mark' + colour, 'aria-hidden': 'true' }) : null;
}

/** Small line icons, drawn in the colour of the text so that they follow the client's theme. */
const ICONS = {
	generate: 'M4 4h10v16H4z M4 9h10 M4 14h10 M9 4v16 M17 12h5 M19.5 9.5 22 12l-2.5 2.5',
	download: 'M12 4v11 M7.5 10.5 12 15l4.5-4.5 M5 20h14',
	upload: 'M12 15V4 M7.5 8.5 12 4l4.5 4.5 M5 20h14',
	import: 'M4 12.5l5 5L20 6.5',
	add: 'M12 5v14 M5 12h14',
	remove: 'M5 12h14'
};

function icon(doc, name) {
	const ns = 'http://www.w3.org/2000/svg';
	const svg = doc.createElementNS(ns, 'svg');
	svg.setAttribute('viewBox', '0 0 24 24');
	svg.setAttribute('class', 'mce-icon');
	svg.setAttribute('aria-hidden', 'true');
	const path = doc.createElementNS(ns, 'path');
	path.setAttribute('d', ICONS[name]);
	svg.append(path);
	return svg;
}

/** A command of the tool's toolbar: a flat Aras button. The primary one is the job's next step. */
function toolButton(view, { id, iconName, label, primary, disabled, title, onClick }) {
	const button = element(view.doc, 'button', {
		id, type: 'button', class: 'aras-button mce-tool ' + (primary ? 'aras-button_primary' : 'aras-button_b'), disabled: !!disabled, title: title || null
	}, icon(view.doc, iconName), element(view.doc, 'span', { class: 'aras-button__text' }, label));
	if (onClick) {
		button.addEventListener('click', onClick);
	}
	return button;
}

/** An Aras button of the tool's panels, with an optional icon. */
function panelButton(view, { id, label, iconName, primary, disabled, title, onClick }) {
	const node = element(view.doc, 'button', {
		id, type: 'button', class: 'aras-button mce-button' + (primary ? ' aras-button_primary' : ''), disabled: !!disabled, title: title || null
	}, iconName ? icon(view.doc, iconName) : null, element(view.doc, 'span', { class: 'aras-button__text' }, label));
	if (onClick) {
		node.addEventListener('click', onClick);
	}
	return node;
}

/**
 * Puts text on the clipboard and says so with Aras's success message. When the browser does not
 * let the copy happen, the client's error message asks the user to press Ctrl+C.
 */
function copyText(view, text, message) {
	const aras = view.top.aras;
	const doc = view.doc;
	const focused = doc.activeElement;
	const buffer = element(doc, 'textarea', { class: 'mce-copy-buffer', readonly: 'readonly', 'aria-hidden': 'true', tabindex: '-1' });
	buffer.value = text;
	doc.body.append(buffer);
	const selection = doc.defaultView.getSelection();
	selection.removeAllRanges();
	buffer.select();
	let copied = false;
	try {
		copied = doc.execCommand('copy');
	} catch (error) {
		copied = false;
	}
	selection.removeAllRanges();
	buffer.remove();
	if (focused && typeof focused.focus === 'function') {
		focused.focus();
	}
	if (copied) {
		aras.AlertSuccess(message);
	} else {
		aras.AlertError(aras.getResource('', 'clipboardmanager.use_ctrl_c'));
	}
}

let checkboxCount = 0;

/**
 * A tick box in the markup of the check boxes of Aras's Forms: the input, and after it the label
 * that draws the box and holds the text. label is the element to place: the two together.
 *
 * field: the box shows a setting of the job, like a check box field of the Form. Read only, it
 * then looks as such a field does: plain while the job is not edited, dimmed while it is.
 */
function checkbox(doc, { id, checked, disabled, title, field: isField, onChange }, ...text) {
	const inputId = id || 'mce-check-' + ++checkboxCount;
	const input = element(doc, 'input', { type: 'checkbox', id: inputId, class: 'arasCheckboxOrRadio', checked: !!checked, disabled: !!disabled });
	if (onChange) {
		input.addEventListener('change', () => onChange(input.checked, input));
	}
	const label = element(doc, 'span', { class: 'mce-check' + (isField ? ' sys_ft_checkbox' : '') + (disabled ? ' mce-check-off' : ''), title: title || null },
		input, element(doc, 'label', { for: inputId }, ...text));
	return { label, input };
}

/** A drop-down drawn by Aras's own select classes. */
function selectBox(doc, { id, className, onChange }, ...options) {
	const select = element(doc, 'select', { id }, ...options);
	if (onChange) {
		select.addEventListener('change', () => onChange(select.value, select));
	}
	const box = element(doc, 'div', { class: 'aras-select mce-select' + (className ? ' ' + className : '') },
		element(doc, 'div', { class: 'aras-select__container' }, select, element(doc, 'span', { class: 'aras-select__button', 'aria-hidden': 'true' })));
	return { box, select };
}

/** Puts a select element in the markup of Aras's drop-downs. */
function arasSelect(doc, select) {
	return element(doc, 'div', { class: 'aras-select mce-select' },
		element(doc, 'div', { class: 'aras-select__container' }, select, element(doc, 'span', { class: 'aras-select__button', 'aria-hidden': 'true' })));
}

/** A line of a card: its label on the left, its value on the right. */
function field(doc, label, ...value) {
	return element(doc, 'div', { class: 'mce-field' }, element(doc, 'span', { class: 'mce-field-label' }, label),
		element(doc, 'div', { class: 'mce-field-value' }, ...value));
}

/**
 * A dialog of the tool, built from the markup and classes of the client's own dialogs, so that it
 * looks like them: a title bar with a close button, the content, and a bar of buttons. It is shown
 * modal in the Form's document and removed when it closes.
 *
 * @returns {{ node: HTMLDialogElement, content: Element, buttons: Element, close(): void }}
 */
function openDialog(view, { id, title, className }) {
	const { doc } = view;
	const old = doc.getElementById(id);
	if (old) {
		old.remove();
	}
	const close = element(doc, 'span', { class: 'aras-dialog__close-button aras-icon-close', role: 'button', tabindex: '0', title: 'Close', 'aria-label': 'Close' });
	const content = element(doc, 'div', { class: 'aras-dialog__content' });
	const buttons = element(doc, 'div', { class: 'aras-buttons-bar aras-buttons-bar_right' });
	const node = element(doc, 'dialog', { id, class: 'aras-dialog mce-dialog' + (className ? ' ' + className : ''), 'aria-label': title },
		element(doc, 'div', { class: 'aras-dialog__title-bar' }, close, element(doc, 'span', { class: 'aras-dialog__title' }, title)),
		content);
	const shut = () => {
		if (node.open) {
			node.close();
		}
	};
	close.addEventListener('click', shut);
	close.addEventListener('keydown', (event) => {
		if (event.key === 'Enter' || event.key === ' ') {
			event.preventDefault();
			shut();
		}
	});
	node.addEventListener('close', () => node.remove());
	doc.body.append(node);
	const dialog = { node, content, buttons, close: shut };
	// The content is filled by the caller; the bar of buttons comes last.
	queueMicrotask(() => {
		content.append(buttons);
		node.showModal();
	});
	return dialog;
}

/** A button of a dialog's bar: OK is the primary one, as in the client's own dialogs. */
function dialogButton(view, label, primary, onClick, id) {
	const node = element(view.doc, 'button', {
		type: 'button', id: id || null, class: 'aras-button aras-buttons-bar__button ' + (primary ? 'aras-button_primary' : 'aras-button_secondary-light')
	}, element(view.doc, 'span', { class: 'aras-button__text' }, label));
	node.addEventListener('click', onClick);
	return node;
}

/**
 * Builds the frame of the tool in the Form: toolbar, header, messages, the two tabs and their
 * panels. The export section draws into the Export panel and its part of the toolbar, the
 * import section into the Import panel and its part.
 */
function buildShell(doc, root) {
	const make = (tag, attributes, ...children) => element(doc, tag, attributes, ...children);
	root.textContent = '';
	root.className = 'mce-app';
	const shell = {
		root,
		exportTools: make('div', { class: 'mce-tools', role: 'group', 'aria-label': 'Workbook' }),
		importTools: make('div', { class: 'mce-tools', role: 'group', 'aria-label': 'Import' }),
		exportStatus: make('span', { class: 'mce-status', role: 'status' }),
		importStatus: make('span', { class: 'mce-status', role: 'status' }),
		head: make('div', { class: 'mce-head' }),
		exportMessages: make('div', { class: 'mce-messages' }),
		importMessages: make('div', { class: 'mce-messages' }),
		tabs: make('ul', { class: 'aras-tabs__list', role: 'tablist', 'aria-label': 'Steps' }),
		exportPanel: make('div', { class: 'mce-panel', role: 'tabpanel', id: 'mce-panel-export', 'aria-labelledby': 'mce-step-export' }),
		importPanel: make('div', { class: 'mce-panel mce-panel-import', role: 'tabpanel', id: 'mce-panel-import', 'aria-labelledby': 'mce-step-import' }),
		exportHost: make('div', { id: HOST_ID }),
		importHost: make('div', { id: IMPORT_HOST_ID })
	};
	// What a run said is shown at the top of the step it belongs to.
	shell.exportPanel.append(shell.exportMessages, shell.exportHost);
	shell.importPanel.append(shell.importMessages, shell.importHost);
	root.append(
		make('div', { class: 'mce-bar', role: 'toolbar', 'aria-label': 'Mass Create & Edit' },
			shell.exportTools, make('span', { class: 'mce-bar-separator', 'aria-hidden': 'true' }), shell.importTools,
			make('span', { class: 'mce-bar-status' }, shell.exportStatus, shell.importStatus)),
		shell.head,
		make('div', { class: 'aras-tabs mce-steps' }, make('div', { class: 'aras-tabs__list-container' }, shell.tabs)),
		make('div', { class: 'mce-panels' }, shell.exportPanel, shell.importPanel));
	return shell;
}

/**
 * Fills the tool's Form in a job Form.
 *
 * @param {Window} formWindow The window of the Form, as the Form's own events see it.
 * @returns {Promise<void>} Settles when the sections are drawn.
 */
export async function attach(formWindow) {
	const doc = formWindow.document;
	const root = doc.getElementById(APP_ID);
	if (!root || !doc.item) {
		return;
	}
	const top = formWindow.aras.getMainWindow();
	if (!doc.getElementById('mce-job-form-style')) {
		doc.head.append(element(doc, 'link', { id: 'mce-job-form-style', rel: 'stylesheet', href: sibling('./mce-job-form.css') }));
	}
	const shell = buildShell(doc, root);
	const view = {
		window: formWindow,
		doc,
		host: shell.exportHost,
		shell,
		top,
		jobId: doc.item.getAttribute('id'),
		// Saving and refreshing are commands of the item window. A Form shown elsewhere, as a preview, has none.
		itemWindow: formWindow.parent && typeof formWindow.parent.onDoneCommand === 'function' ? formWindow.parent : null,
		openKeys: new Set(),
		model: null
	};
	const memory = memoryOf(view.jobId);
	if (memory.view && memory.view.doc === doc) {
		view.openKeys = memory.view.openKeys;
	}
	// What the last run said is shown once, on the Form as it comes back from that run. A Form opened later starts clean.
	if (memory.flow && memory.flow.kind !== 'busy' && Date.now() - memory.flow.at > MESSAGE_KEPT_MS) {
		memory.flow = null;
	}
	if (memory.view && memory.view !== view) {
		leaveView(memory.view);
	}
	memory.view = view;
	if (view.itemWindow) {
		askBeforeClosingNewJob(view.itemWindow);
	}
	await Promise.all([render(view), renderImport(view)]);
}

/** Reads what the section shows and draws it. */
async function render(view) {
	const memory = memoryOf(view.jobId);
	// sections: one for each target of the job, in its order (see loadSection).
	const model = { job: readJob(view), loading: true, problem: null, config: null, sections: [], options: { sheets: [] } };
	view.model = model;
	paint(view);
	try {
		logic = logic || (await loadLogic());
		model.job = readJob(view);
		model.config = await configOf(view);
		model.sections = await Promise.all(model.job.targets.map((target) => loadSection(view, model.config, target.itemTypeId)));
		if (model.sections.length > 0 && model.sections[0].schema) {
			model.options = await startingOptions(view, model);
		}
	} catch (error) {
		model.problem = error.message;
	}
	model.loading = false;
	if (memory.view === view && view.model === model) {
		paint(view);
		loadSelectionNames(view, model);
	}
}

/**
 * What the Workbook tab shows of one target: the allow-listed ItemType (target, null when the
 * user may not use it), its template schema with every relationship, the key of its main sheet
 * and the relationships it offers as sheets. problem: why the schema could not be read.
 */
async function loadSection(view, config, itemTypeId) {
	const section = {
		itemTypeId, target: config.allowed.find((itemType) => itemType.id.toUpperCase() === itemTypeId) || null,
		schema: null, key: null, choices: [], problem: null
	};
	if (section.target) {
		try {
			section.schema = await readSchema(view.top, section.target.name);
			section.key = logic.describeKey(section.schema);
			section.choices = logic.sheetChoices(section.schema);
		} catch (error) {
			section.problem = error.message;
		}
	}
	return section;
}

/**
 * The export options with what the targets don't offer removed (see tidyOptions). The first
 * target is the options' own, the others are in options.targets.
 */
function tidyAll(model, options) {
	const [first, ...rest] = model.sections;
	const more = rest.filter((section) => !!section.schema)
		.map((section) => ({ itemTypeId: section.itemTypeId, choices: section.choices, mainSheet: section.schema.sheets[0] }));
	return logic.tidyOptions(options, first.choices, first.schema.sheets[0], more);
}

/** Reads the names of the selected items the Form does not know yet, then draws the section again. */
async function loadSelectionNames(view, model) {
	const memory = memoryOf(view.jobId);
	let read = false;
	for (const [index, section] of model.sections.entries()) {
		const target = model.job.targets[index];
		const ids = target ? target.prefillIds.filter((id) => !memory.names.has(id)) : [];
		if (ids.length === 0 || !section.target) {
			continue;
		}
		read = true;
		try {
			for (const part of logic.chunk(ids)) {
				const found = new Map((await query(view.top, logic.selectionNamesAml(section.target.name, part)))
					.map((item) => [item.id, textOf(item, 'keyed_name') || item.id]));
				for (const id of part) {
					memory.names.set(id, found.has(id) ? found.get(id) : null);
				}
			}
		} catch (error) {
			// Without the names, the list shows the ids.
			for (const id of ids) {
				if (!memory.names.has(id)) {
					memory.names.set(id, id);
				}
			}
		}
	}
	if (read && memory.view === view && view.model === model) {
		paint(view);
	}
}

/** The export options the section starts with. */
async function startingOptions(view, model) {
	const { job } = model;
	const memory = memoryOf(job.id);
	// Choices not generated yet. The job's export options belong to the server, which stores
	// them when a workbook is generated, so they are kept here until then.
	if (memory.pending && memory.pending.base === job.base) {
		return tidyAll(model, logic.parseOptions(memory.pending.text));
	}
	memory.pending = null;
	if (job.isNew || job.optionsText.trim() !== '') {
		return tidyAll(model, logic.parseOptions(job.optionsText));
	}
	// A job whose options were never chosen here: its Sheet rows say what it exports.
	const rows = await query(
		view.top,
		'<Item type="FRG_MassCreateEdit_Sheet" action="get" select="relationship_type,match_properties"><source_id>' + job.id + '</source_id></Item>'
	);
	// The Sheet rows of the main sheets name no relationship.
	const sheets = rows.filter((row) => textOf(row, 'relationship_type') !== '').map((row) => ({
		relationshipTypeId: textOf(row, 'relationship_type'),
		name: '',
		matchKey: textOf(row, 'match_properties').split(',').map((id) => id.trim()).filter((id) => id !== '')
	}));
	return tidyAll(model, { sheets });
}

/** Draws the export section, and the header and tabs, which depend on the job as it is now. */
function paint(view) {
	drawSection(view);
	drawChrome(view);
	fitFrame(view);
}

// The job has no relationship tabs, so the strip above the Form is hidden in the job's own
// item window, and the tool fills the window, as an editor does.
const ACCORDION_STYLE = '.aras-accordion__header { display: none !important; }';

/** Hides the strip above the Form in the job's item window. Without an open shadow root the strip stays. */
function hideFormStrip(holder) {
	const accordion = holder.closest('aras-accordion');
	const root = accordion ? accordion.shadowRoot : null;
	if (!root || root.getElementById('mce-accordion-style')) {
		return;
	}
	const style = holder.ownerDocument.createElement('style');
	style.id = 'mce-accordion-style';
	style.textContent = ACCORDION_STYLE;
	root.append(style);
}

/**
 * Makes the Form's frame fill the item window below its toolbar, so that the tool's
 * toolbar stays in view and each tab scrolls on its own. A Form shown elsewhere, as a
 * preview, keeps the height it is given, at least MIN_TOOL_HEIGHT.
 */
function fitFrame(view) {
	const frame = view.window.frameElement;
	const holder = frame ? frame.parentElement : null;
	if (!holder) {
		return;
	}
	const scroller = view.itemWindow ? holder.closest('#formTab') : null;
	if (!scroller) {
		if (/^[0-9.]+px$/.test(holder.style.height) && parseFloat(holder.style.height) < MIN_TOOL_HEIGHT) {
			holder.style.height = MIN_TOOL_HEIGHT + 'px';
		}
		return;
	}
	hideFormStrip(holder);
	// The space the item window keeps below the Form (the bottom margins, paddings and borders
	// around it) stays: the frame takes the rest of the room.
	const style = (node) => view.itemWindow.getComputedStyle(node);
	const px = (value) => parseFloat(value) || 0;
	let below = px(style(holder).marginBottom);
	const up = (node) => node.assignedSlot || node.parentElement || (node.parentNode && node.parentNode.host) || null;
	for (let node = up(holder); node && node !== scroller; node = up(node)) {
		const box = style(node);
		below += px(box.paddingBottom) + px(box.borderBottomWidth) + px(box.marginBottom);
	}
	const room = scroller.getBoundingClientRect().bottom - holder.getBoundingClientRect().top - below;
	holder.style.height = Math.max(MIN_TOOL_HEIGHT, Math.floor(room)) + 'px';
	const jobId = view.jobId;
	const again = () => {
		const current = memoryOf(jobId).view;
		if (current && !current.window.closed && current.window.frameElement && scroller.contains(current.window.frameElement)) {
			fitFrame(current);
		}
	};
	// The item window changes size with the browser window and its panes: the frame follows.
	// The accordion changes size when the strip above the Form is hidden.
	if (!scroller.mceResizeObserver && view.itemWindow.ResizeObserver) {
		scroller.mceResizeObserver = new view.itemWindow.ResizeObserver(again);
		scroller.mceResizeObserver.observe(scroller);
		const accordion = holder.closest('aras-accordion');
		if (accordion) {
			scroller.mceResizeObserver.observe(accordion);
		}
	}
	// While the item window is still being laid out, the room measured is not final yet:
	// the frame is measured again shortly after it is first filled.
	if (!holder.mceRefitted) {
		holder.mceRefitted = true;
		for (const delay of [100, 500, 1500]) {
			view.itemWindow.setTimeout(again, delay);
		}
	}
}

/** The tab a job opens on: the step that comes next. */
function defaultTab(job) {
	if (job.state === 'Validated' || job.state === 'Imported') {
		return 'import';
	}
	return job.state === 'Generated' && job.exportFile ? 'import' : 'export';
}

/** The tab shown: the user's choice while the job stays in the state it was made in. */
function currentTab(view, job) {
	const chosen = memoryOf(view.jobId).tab;
	return chosen && chosen.state === job.state ? chosen.tab : defaultTab(job);
}

function showTab(view, tab) {
	const job = readJob(view);
	memoryOf(view.jobId).tab = { tab, state: job.state };
	drawChrome(view);
	if (tab === 'import' && view.previewUi && view.previewUi.grid) {
		// A grid drawn while its tab was hidden had no height to draw rows into.
		view.previewUi.grid.draw();
	}
}

/** A date of the job, as the client shows dates. */
function shownDate(view, neutral) {
	if (!neutral) {
		return '';
	}
	try {
		return view.top.aras.convertFromNeutral(neutral, 'date', 'short_date_time') || neutral;
	} catch (error) {
		return neutral;
	}
}

/** Reads a File of the job from the vault and hands it to the browser. */
async function downloadJobFile(view, fileId, fileName, type) {
	try {
		const buffer = await readStoredFile(view.top, fileId);
		download(view.top, buffer, fileName || 'file', type);
	} catch (error) {
		view.top.aras.AlertError(error && error.message ? error.message : String(error));
	}
}

/** Draws the header (state, target, files) and the tabs, and shows the panel of the current tab. */
function drawChrome(view) {
	const { doc, shell } = view;
	if (!shell || !doc.contains(shell.root)) {
		return;
	}
	const make = (tag, attributes, ...children) => element(doc, tag, attributes, ...children);
	const job = readJob(view);
	const focusId = doc.activeElement && shell.tabs.contains(doc.activeElement) ? doc.activeElement.id : '';

	// --- Header ---
	shell.head.textContent = '';
	const field = (label, ...value) => make('span', { class: 'mce-head-field' }, make('span', { class: 'mce-head-label' }, label), ...value);
	shell.head.append(field('State', make('span', { class: 'mce-state mce-state-' + job.state.toLowerCase() }, job.isNew ? 'New' : job.state)));
	const model = view.model;
	const targetNames = model && model.sections.length > 0 ? model.sections.map((section, index) => targetLabel(model, index)) : [job.targetName].filter(Boolean);
	shell.head.append(field(targetNames.length > 1 ? 'Targets' : 'Target', make('span', { class: 'mce-head-value' }, targetNames.join(' + ') || 'None')));
	const fileLink = (id, fileId, name, type, title) => {
		const link = make('button', { type: 'button', id, class: 'mce-file-link', title }, name || 'file');
		link.addEventListener('click', () => downloadJobFile(view, fileId, name, type));
		return link;
	};
	if (job.exportFile) {
		shell.head.append(field('Workbook', fileLink('mce-head-export', job.exportFile, job.fileNames.export, XLSX_TYPE, 'Download the generated workbook'),
			job.exportedOn ? make('span', { class: 'mce-muted' }, shownDate(view, job.exportedOn)) : null));
	}
	if (job.uploadFile) {
		shell.head.append(field('Upload', fileLink('mce-head-upload', job.uploadFile, job.fileNames.upload, XLSX_TYPE, 'Download the uploaded workbook')));
	}
	// The validation result is not linked: it is the tool's own record of the preview, read back when the job
	// is opened again, and of the skipped rows after the import. The Import tab shows everything it holds.
	if (job.copiedFrom) {
		shell.head.append(field('Copied from', make('span', { class: 'mce-head-value' }, job.copiedFrom)));
	}

	// --- Tabs ---
	const tab = currentTab(view, job);
	shell.tabs.textContent = '';
	const importLabel = job.state === 'Imported' ? 'Result' : job.state === 'Validated' ? 'Import Preview' : 'Import';
	for (const [name, number, label] of [['export', '1', 'Workbook'], ['import', '2', importLabel]]) {
		const on = name === tab;
		const button = make('li', {
			role: 'tab', id: 'mce-step-' + name, class: 'aras-tabs__tab mce-step' + (on ? ' aras-tabs__tab_active mce-step-on' : ''),
			'aria-selected': on ? 'true' : 'false', 'aria-controls': 'mce-panel-' + name, tabindex: on ? '0' : '-1'
		}, make('span', { class: 'mce-step-number' }, number), make('span', { class: 'aras-tabs__label' }, label));
		button.addEventListener('click', () => showTab(view, name));
		button.addEventListener('keydown', (event) => {
			if (event.key === 'Enter' || event.key === ' ') {
				event.preventDefault();
				showTab(view, name);
			} else if (event.key === 'ArrowLeft' || event.key === 'ArrowRight') {
				event.preventDefault();
				showTab(view, name === 'export' ? 'import' : 'export');
				const other = doc.getElementById('mce-step-' + (name === 'export' ? 'import' : 'export'));
				if (other) {
					other.focus();
				}
			}
		});
		shell.tabs.append(button);
	}
	shell.exportPanel.hidden = tab !== 'export';
	shell.importPanel.hidden = tab !== 'import';
	if (focusId) {
		const focused = doc.getElementById(focusId);
		if (focused) {
			focused.focus();
		}
	}
}

function drawSection(view) {
	const { doc, host, model, shell } = view;
	const { job } = model;
	const memory = memoryOf(view.jobId);
	const flow = memory.flow;
	const busy = (!!flow && flow.kind === 'busy') || importBusy(memory);
	const make = (tag, attributes, ...children) => element(doc, tag, attributes, ...children);
	const focusId = doc.activeElement && shell.root.contains(doc.activeElement) ? doc.activeElement.id : '';

	host.textContent = '';
	host.className = 'mce-job mce-setup';
	shell.exportTools.textContent = '';
	shell.exportMessages.textContent = '';
	shell.exportStatus.textContent = '';

	const refusal = whyNot(model);
	const locked = busy || refusal !== null || !view.itemWindow;
	// Begin moves a Draft job to Generated before the workbook exists, so a Generated job may have none:
	// generating stopped half way. Generating again is the way on.
	const again = job.exportFile !== '';

	// --- Toolbar: generating, and downloading the workbook generated ---
	// What keeps the workbook from being generated, shown on the command. Set below, once known.
	let cannot = model.loading ? 'Loading...' : refusal || (!view.itemWindow ? 'Open the job to generate its workbook.' : null);
	const generateButton = toolButton(view, {
		id: 'mce-generate', iconName: 'generate', label: again ? 'Generate Again' : 'Generate Workbook',
		primary: !again && !refusal, onClick: () => generate(view.jobId)
	});
	const downloadButton = toolButton(view, {
		id: 'mce-download', iconName: 'download', label: 'Download Workbook', disabled: !again,
		title: again ? 'Download ' + (job.fileNames.export || 'the generated workbook') : 'The job has no workbook yet.',
		onClick: () => downloadJobFile(view, job.exportFile, job.fileNames.export, XLSX_TYPE)
	});
	shell.exportTools.append(generateButton, downloadButton);
	const finishToolbar = (blocked) => {
		cannot = cannot || blocked;
		generateButton.disabled = locked || model.loading || !!cannot;
		generateButton.title = cannot || (again
			? 'Generate the workbook again: it replaces the Generated Workbook of the job.'
			: 'Generate the workbook and download it.');
		if (busy && flow && flow.kind === 'busy') {
			shell.exportStatus.append(make('span', { class: 'mce-spinner', 'aria-hidden': 'true' }), flow.message);
		}
	};

	// --- Messages of the last run ---
	if (flow && flow.kind === 'error') {
		shell.exportMessages.append(make('div', { class: 'mce-message mce-message-error', role: 'alert' }, flow.message));
	}
	if (flow && flow.kind === 'done') {
		if (flow.fileId === job.exportFile) {
			const done = make('div', { class: 'mce-message mce-message-done', role: 'status' }, make('p', null, flow.message));
			if (flow.notes.length > 0) {
				done.append(make('ul', null, ...flow.notes.map((note) => make('li', null, note))));
			}
			shell.exportMessages.append(done);
		} else {
			// The message was about a workbook that is no longer the job's.
			memory.flow = null;
		}
	}

	// --- The panel: for each target, its main sheet on the left and its relationship sheets on the right ---
	if (refusal && !model.loading) {
		host.append(make('p', { class: 'mce-note' }, refusal));
	} else if (!view.itemWindow) {
		host.append(make('p', { class: 'mce-note' }, 'Open the job to generate its workbook.'));
	}
	// Targets are added, removed and changed while the job is a Draft that is being edited.
	const canEdit = !locked && job.editing && job.state === 'Draft' && !!model.config;
	// The relationship sheets and the columns can be changed while the job is being edited, in any state
	// the workbook can still be generated in.
	const canChange = !locked && job.editing;
	const state = { locked, refusal, canEdit, canChange };
	const notes = generatingNotes(view, busy, refusal, again);
	if (notes.length > 0 && !model.loading) {
		host.append(make('p', { class: 'mce-note' }, notes.join(' ')));
	}

	if (model.loading || model.problem || model.sections.length === 0) {
		const workbook = make('section', { class: 'mce-card' }, make('h3', { class: 'mce-card-title' }, 'Target'));
		host.append(make('div', { class: 'mce-columns' }, workbook));
		workbook.append(targetRow(view, 0, null, canEdit && model.sections.length === 0 && !model.loading));
		let text;
		let blocked = null;
		if (model.loading) {
			text = make('p', { class: 'mce-muted' }, 'Loading...');
		} else if (model.problem) {
			text = make('p', { class: 'mce-error', role: 'alert' }, model.problem);
			blocked = model.problem;
		} else {
			text = make('p', { class: 'mce-help' }, canEdit
				? 'Choose the ItemType whose items the workbook creates and changes.'
				: 'The job has no target ItemType.' + (refusal ? '' : ' Edit the job to choose one.'));
			blocked = 'Choose the target ItemType first.';
		}
		workbook.append(text);
		finishToolbar(blocked);
		restoreFocus(doc, focusId);
		return;
	}

	let blocked = null;
	model.sections.forEach((section, index) => {
		if (index > 0) {
			host.append(make('hr', { class: 'mce-target-divider' }));
		}
		blocked = drawTarget(view, section, index, state) || blocked;
	});

	// --- A target not chosen yet, and the command that adds one ---
	const unused = model.config ? model.config.allowed.filter((itemType) => !job.targets.some((target) => target.itemTypeId === itemType.id.toUpperCase())) : [];
	if (memory.adding && canEdit && unused.length > 0) {
		host.append(make('hr', { class: 'mce-target-divider' }));
		const index = model.sections.length;
		const workbook = make('section', { class: 'mce-card' }, make('h3', { class: 'mce-card-title' }, 'Target ' + (index + 1)));
		host.append(make('div', { class: 'mce-columns' }, workbook));
		workbook.append(targetRow(view, index, null, true));
		workbook.append(make('p', { class: 'mce-help' },
			'Choose another ItemType whose items the workbook creates and changes. It gets a main sheet of its own, after those of the targets above.'));
	} else if (canEdit) {
		const add = panelButton(view, {
			id: 'mce-target-add', label: 'Add Target', iconName: 'add', disabled: unused.length === 0,
			title: unused.length === 0 ? 'Every ItemType you may use is a target of the job already.' : null
		});
		add.addEventListener('click', () => {
			memory.adding = true;
			paint(view);
			const select = doc.getElementById('mce-target-' + (model.sections.length + 1));
			if (select) {
				select.focus();
			}
		});
		host.append(make('div', { class: 'mce-target-add' }, add,
			make('span', { class: 'mce-muted' }, 'A workbook can create and change the items of several ItemTypes, each on a main sheet of its own.')));
	}

	finishToolbar(blocked);
	restoreFocus(doc, focusId);
}

/** What generating does, said before it is done: a note under the first target. */
function generatingNotes(view, busy, refusal, again) {
	const { job } = view.model;
	const notes = [];
	if (!view.itemWindow || busy || refusal) {
		return notes;
	}
	if ((job.isNew || job.editing) && job.state === 'Draft') {
		notes.push('Generating saves the job and finishes editing it.');
	}
	if (!again && job.state !== 'Draft') {
		notes.push('The job has no workbook yet: generating it did not finish. Generate the workbook again.');
	}
	if (again) {
		notes.push('A job has one valid workbook: generating again replaces the Generated Workbook of the job, and an upload of an earlier workbook is refused.');
	}
	if (job.state === 'Validated') {
		notes.push('The job goes back to Generated, and the uploaded workbook and its validation result are deleted from Aras: you are asked first.');
	}
	return notes;
}

/** The id of an element of a target's section: the first target's ids are those of a job with one target. */
const sectionId = (base, index) => (index === 0 ? base : base + '-' + (index + 1));

/**
 * The ItemType of a section: a drop-down while it can be changed, with the ItemTypes no other
 * target has, and Remove Target on a target after the first.
 */
function targetRow(view, index, section, canChange) {
	const { doc, model } = view;
	const { job } = model;
	const make = (tag, attributes, ...children) => element(doc, tag, attributes, ...children);
	const id = sectionId('mce-target', index);
	const current = job.targets[index] ? job.targets[index].itemTypeId : '';
	const line = make('div', { class: 'mce-line' });
	if (canChange) {
		const options = [];
		// The first target can be cleared while it is the only one. A target after it is removed instead.
		if (current === '' || (index === 0 && job.targets.length <= 1)) {
			options.push(make('option', { value: '' }, 'Choose an ItemType'));
		}
		for (const itemType of model.config.allowed) {
			const itemTypeId = itemType.id.toUpperCase();
			if (itemTypeId === current || !job.targets.some((target) => target.itemTypeId === itemTypeId)) {
				options.push(make('option', { value: itemTypeId, selected: itemTypeId === current }, itemType.label));
			}
		}
		if (current && !(section && section.target)) {
			options.push(make('option', { value: current, selected: true }, targetLabel(model, index)));
		}
		const { box, select } = selectBox(doc, { id, className: 'mce-select-wide', onChange: (value) => changeTarget(view, index, value) }, ...options);
		select.setAttribute('aria-label', 'Target ItemType');
		line.append(box);
	} else {
		line.append(make('span', { id, class: 'mce-strong' }, current ? targetLabel(model, index) : 'None'));
	}
	const memory = memoryOf(view.jobId);
	const canRemove = index > 0 && !!model.config && job.editing && job.state === 'Draft' && !(memory.flow && memory.flow.kind === 'busy') && !importBusy(memory);
	if (canRemove) {
		line.append(panelButton(view, { id: sectionId('mce-target-remove', index), label: current ? 'Remove Target' : 'Cancel', onClick: () => removeTarget(view, index) }));
	}
	return field(doc, 'ItemType', line);
}

/**
 * Draws the section of one target: its Target card (ItemType, key, items, columns) and its
 * Relationships card. Returns what keeps the workbook from being generated, or null.
 */
function drawTarget(view, section, index, { refusal, canEdit, canChange }) {
	const { doc, host, model } = view;
	const { job } = model;
	const make = (tag, attributes, ...children) => element(doc, tag, attributes, ...children);
	const several = model.sections.length > 1;
	const card = (title, ...children) => make('section', { class: 'mce-card' }, make('h3', { class: 'mce-card-title' }, ...title), ...children);
	const workbook = card([several ? 'Target ' + (index + 1) : 'Target']);
	const columns = make('div', { class: 'mce-columns' }, workbook);
	host.append(columns);
	const prefillIds = job.targets[index] ? job.targets[index].prefillIds : [];
	const count = prefillIds.length;

	// --- Target ItemType: changed while no items are selected for it ---
	workbook.append(targetRow(view, index, section, canEdit && count === 0));
	if (section.problem) {
		workbook.append(make('p', { class: 'mce-error', role: 'alert' }, section.problem));
		return section.problem;
	}
	if (!section.target) {
		const limited = model.config.restricted.find((itemType) => itemType.id.toUpperCase() === section.itemTypeId) || null;
		const text = limited
			? 'Mass Create & Edit on ItemType \'' + limited.label + '\' is limited to the members of identity \'' + limited.identityName + '\'. Ask an administrator.'
			: 'ItemType \'' + targetLabel(model, index) + '\' is not on the allow-list of Mass Create & Edit. ' +
				'An administrator can add it on the configuration item.';
		workbook.append(make('p', { class: 'mce-error', role: 'alert' }, text));
		return text;
	}

	// --- The key of the main sheet ---
	const key = section.key;
	const keyNames = key.columns.map((column) => column.label).join(' + ');
	workbook.append(field(doc, 'Key',
		make('div', { class: 'mce-chips' }, ...(key.createOnly
			? [make('span', { class: 'mce-chip mce-chip-fixed' }, 'None (create only)')]
			: key.columns.map((column) => make('span', { class: 'mce-chip mce-chip-fixed' }, column.label)))),
		make('p', { class: 'mce-help' }, key.createOnly
			? key.itemType + ' has no Unique properties, so a row can\'t be matched with an existing item. Every row creates a new ' + key.itemType + '.'
			: 'A row whose ' + keyNames + ' exists changes that ' + key.itemType + '. Any other row creates a new one. The key is fixed: ' +
				(key.columns.length === 1 ? 'it is the Unique property' : 'they are the Unique properties') + ' of ' + key.itemType + '.')));

	// --- The items the main sheet is filled with ---
	let blocked = null;
	if (count > 0 && !refusal && key.createOnly) {
		blocked = 'Existing items of ' + key.itemType + ' can\'t be pre-filled for editing: without Unique properties the import could not find them again. ' +
			'Remove the selected items to create items instead.';
	}
	// The limit counts the rows of every main sheet together: it is said under the last target that has items.
	const lastWithItems = job.targets.reduce((last, target, at) => (target.prefillIds.length > 0 ? at : last), -1);
	if (!refusal && !blocked && index === lastWithItems && model.config.maxRows > 0 && job.prefillCount > model.config.maxRows) {
		blocked = several
			? 'The job has ' + job.prefillCount + ' items to pre-fill on its main sheets, and a workbook takes at most ' + model.config.maxRows +
				' rows on its main sheets together. Remove items from the selections, or ask an administrator to raise the limit.'
			: 'The job has ' + job.prefillCount + ' items to pre-fill, and a workbook takes at most ' + model.config.maxRows +
				' rows on its main sheet. Remove items from the selection, or ask an administrator to raise the limit.';
	}
	// The selected items can be changed while the job is a Draft that is being edited. Items of an
	// ItemType without a key can only be removed: they could not be found again on import.
	const canRemove = canEdit;
	const canAdd = canRemove && !key.createOnly;
	const items = field(doc, 'Items', ...selectionBlock(view, section, index, { canRemove, canAdd, refusal }));
	workbook.append(items);
	if (blocked) {
		items.lastChild.append(make('p', { class: 'mce-error', role: 'alert' }, blocked));
	}

	// --- The columns of the main sheet ---
	const stored = index === 0 ? model.options : ((model.options.targets || []).find((target) => target.itemTypeId === section.itemTypeId) || {});
	workbook.append(field(doc, 'Columns', columnsLine(view, {
		title: 'Select Columns: ' + section.target.label, sheet: section.schema.sheets[0], omit: stored.omit, matchKey: null, readOnly: !canChange,
		pickerKey: index === 0 ? 'main' : 'main-' + section.itemTypeId, onOmit: (omit) => changeMainColumns(view, index, omit)
	})));

	// --- Relationships ---
	const included = section.choices.filter((choice) => model.options.sheets.some((sheet) => sheet.relationshipTypeId === choice.relationshipTypeId));
	const offered = section.choices.filter((choice) => choice.selectable);
	const sheets = card(['Relationships', offered.length > 0 && !refusal
		? make('span', { class: 'mce-card-count' }, included.length + ' of ' + offered.length + ' included') : null]);
	columns.append(sheets);
	if (refusal) {
		// Nothing can be changed any more: only what the workbook holds is shown.
		if (included.length === 0) {
			sheets.append(make('p', { class: 'mce-help' }, 'None.'));
		}
		const list = make('div', { class: 'mce-sheets' });
		for (const choice of included) {
			list.append(sheetBlock(view, section, choice, true));
		}
		sheets.append(list);
	} else if (offered.length === 0) {
		// Only the relationships that can be a sheet are offered (see sheetChoices for the reasons one can't).
		sheets.append(make('p', { class: 'mce-help' }, section.choices.length === 0
			? key.itemType + ' has no relationships.'
			: key.itemType + ' has no relationships that can be a sheet. The workbook has the ' + key.itemType + ' sheet only.'));
	} else {
		sheets.append(make('p', { class: 'mce-help' },
			'Each ticked relationship gets a sheet of its own' + (count > 0 ? ', filled with the existing rows of the selected items.' : '.')));
		const list = make('div', { class: 'mce-sheets' });
		for (const choice of offered) {
			list.append(sheetBlock(view, section, choice, !canChange));
		}
		sheets.append(list);
	}
	return blocked;
}

/** Gives the focus back to the element that had it before the tool was drawn again. */
function restoreFocus(doc, focusId) {
	if (focusId) {
		const focused = doc.getElementById(focusId);
		if (focused && !focused.disabled) {
			focused.focus();
		}
	}
}

const FLAT_HELP = 'A row for each relationship, with its parent in a column of its own.';
const INDENTED_HELP = 'A Level column instead of the parent: level 0 is a top assembly, level 1 the lines under it, and so on. ' +
	'Selected items are listed with their BOMs at every level, so lines can be changed or deleted; without a selection the sheet starts empty.';

/**
 * One relationship of a target: its tick box and, when ticked, its layout, the columns its
 * existing rows are matched by, whether its rows can name the items the workbook creates, and
 * its columns. readOnly: the job is not being edited, so the choices are shown but not offered.
 */
function sheetBlock(view, section, choice, readOnly) {
	const { doc, model } = view;
	const make = (tag, attributes, ...children) => element(doc, tag, attributes, ...children);
	const stored = model.options.sheets.find((sheet) => sheet.relationshipTypeId === choice.relationshipTypeId) || null;
	const rowId = 'mce-sheet-' + choice.relationshipTypeId;

	// The sheet is named after the RelationshipType ("Part CAD"), so that name leads; the label is the
	// name of the tab on the item ("CAD Documents").
	const about = [
		choice.label && choice.label !== choice.name ? choice.label + ' tab' : '',
		choice.hasRelated ? 'related ' + choice.relatedLabel : choice.selectable ? 'no related item' : ''
	].filter((part) => part !== '').join(' · ');
	const tick = checkbox(doc, { id: rowId, checked: stored !== null, disabled: readOnly || !choice.selectable, field: true, onChange: (on) => toggleSheet(view, choice, on) },
		make('span', { class: 'mce-sheet-name' }, choice.name || choice.label), about ? make('span', { class: 'mce-muted mce-sheet-about' }, about) : null);
	const block = make('div', { class: 'mce-sheet' + (choice.selectable ? '' : ' mce-sheet-off') }, tick.label);
	if (!choice.selectable) {
		block.append(make('p', { class: 'mce-help mce-sheet-note' }, 'Can\'t be a sheet. ' + (choice.reason || '')));
		return block;
	}
	if (stored === null) {
		return block;
	}

	const ticked = logic.selectedGroups(choice, stored.matchKey);
	const detail = make('div', { class: 'mce-sheet-detail' });
	const indented = stored.layout === logic.INDENTED_LAYOUT;
	if (choice.canBeIndented) {
		detail.append(field(doc, 'Layout', layoutSwitch(view, choice, rowId, indented, readOnly), make('p', { class: 'mce-help' }, indented ? INDENTED_HELP : FLAT_HELP)));
	}
	if (choice.createOnly) {
		detail.append(make('ul', { class: 'mce-help mce-consequences' },
			...logic.matchKeyConsequences(choice, ticked, indented).map((line) => make('li', null, line))));
	} else {
		const chips = make('div', { class: 'mce-chips' },
			...choice.groups.filter((group) => ticked.includes(group.key)).map((group) => make('span', { class: 'mce-chip mce-chip-fixed' }, group.label)));
		if (!readOnly) {
			chips.append(panelButton(view, { id: rowId + '-change', label: 'Change...', onClick: () => matchKeyDialog(view, choice, stored, indented) }));
		}
		detail.append(field(doc, 'Matched by', chips));
	}
	const sheet = section.schema.sheets.find((candidate) => !!candidate.relationshipType && candidate.relationshipType.id === choice.relationshipTypeId);
	const relatedHint = sheet ? relatedTargetHint(view, choice, sheet, readOnly) : null;
	if (relatedHint) {
		detail.append(field(doc, 'Related', ...relatedHint));
	}
	if (sheet) {
		detail.append(field(doc, 'Columns', columnsLine(view, {
			title: 'Select Columns: ' + (choice.name || choice.label),
			sheet: indented ? { ...sheet, columns: sheet.columns.filter((column) => column.role !== 'parentKey') } : logic.withoutIndentedColumns(sheet),
			omit: stored.omit, matchKey: stored.matchKey, readOnly,
			pickerKey: choice.relationshipTypeId, onOmit: (omit) => changeColumns(view, choice.relationshipTypeId, omit)
		})));
	}
	block.append(detail);
	return block;
}

/** The two ways a relationship sheet can be laid out, as a switch: Flat or Indented. */
function layoutSwitch(view, choice, rowId, indented, readOnly) {
	const make = (tag, attributes, ...children) => element(view.doc, tag, attributes, ...children);
	const option = (value, label) => {
		const on = value === indented;
		const node = make('button', {
			type: 'button', id: rowId + (value ? '-indented' : '-flat'), class: 'mce-segment' + (on ? ' mce-segment-on' : ''),
			role: 'radio', 'aria-checked': on ? 'true' : 'false', disabled: readOnly
		}, label);
		node.addEventListener('click', () => {
			if (!on) {
				changeLayout(view, choice, value);
			}
		});
		return node;
	};
	return make('div', { class: 'mce-segmented', role: 'radiogroup', 'aria-label': 'Layout of ' + (choice.name || choice.label) },
		option(false, 'Flat'), option(true, 'Indented'));
}

/**
 * The dialog that changes the columns a relationship sheet's existing rows are matched by. What
 * the choice means is said as it is changed; OK keeps it.
 */
function matchKeyDialog(view, choice, stored, indented) {
	const { doc } = view;
	const make = (tag, attributes, ...children) => element(doc, tag, attributes, ...children);
	const now = new Set(logic.selectedGroups(choice, stored.matchKey));
	const keys = () => choice.groups.filter((group) => now.has(group.key)).map((group) => group.key);
	const dialog = openDialog(view, { id: 'mce-dialog-match', title: 'Match Columns: ' + (choice.name || choice.label) });
	const consequences = make('ul', { class: 'mce-help mce-consequences' });
	const explain = () => {
		consequences.textContent = '';
		consequences.append(...logic.matchKeyConsequences(choice, keys(), indented).map((line) => make('li', null, line)));
	};
	const list = make('div', { class: 'mce-dialog-list' });
	for (const group of choice.groups) {
		list.append(checkbox(doc, {
			id: 'mce-dialog-match-' + group.key, checked: now.has(group.key), disabled: group.fixed,
			onChange: (on) => {
				if (on) {
					now.add(group.key);
				} else {
					now.delete(group.key);
				}
				explain();
			}
		}, make('span', null, group.label), group.fixed ? make('span', { class: 'mce-tag' }, 'always') : null).label);
	}
	explain();
	dialog.content.append(make('p', { class: 'mce-help' }, 'Existing rows are matched and changed by the ticked columns.'), list, consequences);
	dialog.buttons.append(
		dialogButton(view, 'OK', true, () => {
			dialog.close();
			changeSheet(view, choice, logic.matchKeyFor(choice, keys()));
		}, 'mce-dialog-match-ok'),
		dialogButton(view, 'Cancel', false, dialog.close, 'mce-dialog-match-cancel'));
}

/**
 * Whether the rows of a relationship sheet can name a related item the workbook creates: they
 * can when the related ItemType is a target of the job. When it is not and could be, the
 * hint offers to add it as a target. Null when there is nothing to say; otherwise the nodes.
 */
function relatedTargetHint(view, choice, sheet, readOnly) {
	const { doc, model } = view;
	const make = (tag, attributes, ...children) => element(doc, tag, attributes, ...children);
	if (!choice.relatedItemTypeId || !sheet.columns.some((column) => column.role === 'relatedKey')) {
		return null;
	}
	const job = model.job;
	if (job.targets.some((target) => target.itemTypeId === choice.relatedItemTypeId)) {
		return [make('p', { class: 'mce-help mce-first' },
			'A row can name an existing ' + choice.relatedLabel + ', or a new one that a row of the ' + choice.relatedLabel + ' sheet creates.')];
	}
	const allowed = model.config && model.config.allowed.find((itemType) => itemType.id.toUpperCase() === choice.relatedItemTypeId);
	if (!allowed || readOnly || job.state !== 'Draft') {
		return null;
	}
	return [
		make('p', { class: 'mce-help mce-first' }, 'A row names an existing ' + choice.relatedLabel + '. To create new ones in the same workbook, add ' + allowed.label + ' as a target.'),
		make('div', { class: 'mce-line' }, panelButton(view, {
			id: 'mce-sheet-' + choice.relationshipTypeId + '-add-target', label: 'Add ' + allowed.label + ' as Target', iconName: 'add',
			onClick: () => changeTarget(view, job.targets.length, allowed.id.toUpperCase())
		}))
	];
}

/**
 * The columns of one sheet: how many of them the workbook gets and, while the job is being
 * edited (not readOnly), the button that opens the dialog to choose them.
 *
 * pickerKey names the sheet in ids ('main' for the main sheet of the first target), and onOmit
 * takes the columns left out after a change.
 */
function columnsLine(view, { title, sheet, omit, matchKey, pickerKey, readOnly, onOmit }) {
	const { doc } = view;
	const choices = logic.columnChoices(sheet, omit, matchKey);
	const included = choices.filter((choice) => choice.included).length;
	const count = included === choices.length ? 'All ' + plural(choices.length, 'column', 'columns') : included + ' of ' + choices.length + ' columns';
	const open = readOnly ? null : panelButton(view, {
		id: 'mce-columns-' + pickerKey + '-toggle', label: 'Select Columns...', onClick: () => columnsDialog(view, { title, choices, onOmit })
	});
	return element(doc, 'div', { class: 'mce-line' }, element(doc, 'span', null, count), open);
}

/**
 * The dialog that chooses the columns of a sheet: a tick box for each. A column the import needs
 * is ticked and can't be changed, with the reason. OK keeps the choice.
 */
function columnsDialog(view, { title, choices, onOmit }) {
	const { doc } = view;
	const make = (tag, attributes, ...children) => element(doc, tag, attributes, ...children);
	const included = new Set(choices.filter((choice) => choice.included).map((choice) => choice.id));
	const dialog = openDialog(view, { id: 'mce-dialog-columns', title, className: 'mce-dialog-columns' });
	const boxes = new Map();
	const entries = [];
	const list = make('div', { class: 'mce-dialog-list mce-column-list' });
	for (const choice of choices) {
		const box = checkbox(doc, {
			id: 'mce-dialog-column-' + choice.id, checked: choice.included, disabled: choice.locked !== null,
			title: choice.locked ? 'Always in the workbook: the import needs this column (' + choice.locked + ').' : null,
			onChange: (on) => (on ? included.add(choice.id) : included.delete(choice.id))
		}, make('span', null, choice.label), choice.locked ? make('span', { class: 'mce-tag' }, choice.locked) : null);
		boxes.set(choice.id, box.input);
		entries.push({ label: box.label, text: (choice.label + ' ' + choice.id).toLowerCase() });
		list.append(box.label);
	}
	const tools = make('div', { class: 'mce-line' });
	if (choices.length > FILTER_FROM_COLUMNS) {
		const filter = make('input', { type: 'search', class: 'aras-input mce-dialog-filter', placeholder: 'Find a column', 'aria-label': 'Find a column', autofocus: 'autofocus' });
		filter.addEventListener('input', () => {
			const wanted = filter.value.trim().toLowerCase();
			for (const entry of entries) {
				entry.label.hidden = wanted !== '' && !entry.text.includes(wanted);
			}
		});
		tools.append(filter);
	}
	const setAll = (on) => {
		for (const choice of choices) {
			if (choice.locked === null) {
				boxes.get(choice.id).checked = on;
				if (on) {
					included.add(choice.id);
				} else {
					included.delete(choice.id);
				}
			}
		}
	};
	tools.append(panelButton(view, { label: 'Select All', onClick: () => setAll(true) }), panelButton(view, { label: 'Required Only', onClick: () => setAll(false) }));
	dialog.content.append(tools, list,
		make('p', { class: 'mce-help' }, 'A column left out is not in the workbook, and an import leaves that property of each item as it is.'));
	dialog.buttons.append(
		dialogButton(view, 'OK', true, () => {
			dialog.close();
			onOmit(choices.filter((choice) => choice.locked === null && !included.has(choice.id)).map((choice) => choice.id));
		}, 'mce-dialog-columns-ok'),
		dialogButton(view, 'Cancel', false, dialog.close, 'mce-dialog-columns-cancel'));
}

/** How many selected items are shown by name before Show All. */
const CHIPS_SHOWN = 8;

/**
 * The items the workbook is pre-filled with: how many, the first of them by name, and Show All
 * for a list of all of them that can be filtered. While the job is a Draft that is being edited,
 * items can be removed and added. Returns the nodes of the Items line of the card.
 */
function selectionBlock(view, section, index, { canRemove, canAdd, refusal }) {
	const { doc, model } = view;
	const make = (tag, attributes, ...children) => element(doc, tag, attributes, ...children);
	const ids = model.job.targets[index] ? model.job.targets[index].prefillIds : [];
	const idOf = (base) => sectionId(base, index);
	const names = memoryOf(view.jobId).names;
	const nameOf = (id) => {
		const name = names.has(id) ? names.get(id) : undefined;
		return { text: name === undefined ? 'Loading...' : name === null ? 'Not found, or you can\'t read it' : name, missing: name === null };
	};
	const count = ids.length;
	const parts = [];
	if (count === 0) {
		parts.push(make('p', { class: 'mce-help mce-first' }, 'None: the workbook is a blank template, with column headers and drop-downs and no rows.'));
	} else {
		parts.push(make('div', { class: 'mce-line' }, make('span', { class: 'mce-strong' }, plural(count, 'item selected', 'items selected')),
			refusal ? null : make('span', { class: 'mce-muted' }, 'pre-filled as ' + (count === 1 ? 'it is' : 'they are') + ' when the workbook is generated')));
		const chips = make('div', { class: 'mce-chips', role: 'list', 'aria-label': 'Selected items of ' + section.target.label });
		ids.slice(0, CHIPS_SHOWN).forEach((id, at) => {
			const { text, missing } = nameOf(id);
			const chip = make('span', { class: 'mce-chip' + (missing ? ' mce-chip-missing' : ''), role: 'listitem', title: missing ? id : text },
				make('span', { class: 'mce-chip-text' }, text));
			if (canRemove) {
				const remove = make('button', { type: 'button', id: idOf('mce-selection-remove') + '-' + at, class: 'mce-chip-remove', title: 'Remove', 'aria-label': 'Remove ' + text }, '×');
				remove.addEventListener('click', () => changeSelection(view, index, { remove: [id] }));
				chip.append(remove);
			}
			chips.append(chip);
		});
		if (count > CHIPS_SHOWN) {
			chips.append(make('span', { class: 'mce-muted mce-chip-more' }, '+' + (count - CHIPS_SHOWN) + ' more'));
		}
		parts.push(chips);
	}
	const openKey = ITEMS_KEY + index;
	const open = count > CHIPS_SHOWN && view.openKeys.has(openKey);
	const actions = make('div', { class: 'mce-line' });
	if (canAdd) {
		actions.append(panelButton(view, { id: idOf('mce-selection-add'), label: 'Add Items...', iconName: 'add', onClick: () => addItems(view, index) }));
	}
	if (count > CHIPS_SHOWN) {
		actions.append(panelButton(view, {
			id: idOf('mce-selection-all'), label: open ? 'Show Less' : 'Show All',
			onClick: () => {
				if (open) {
					view.openKeys.delete(openKey);
				} else {
					view.openKeys.add(openKey);
				}
				paint(view);
			}
		}));
	}
	if (canRemove && count > 1) {
		actions.append(panelButton(view, { id: idOf('mce-selection-clear'), label: 'Remove All', onClick: () => changeSelection(view, index, { remove: ids }) }));
	}
	if (actions.children.length > 0) {
		parts.push(actions);
	}
	if (open) {
		parts.push(selectionTable(view, index, ids, nameOf, canRemove));
	}
	if (canAdd) {
		parts.push(make('p', { class: 'mce-help' }, count === 0
			? 'Add existing items to fill the workbook with them, so that they can be changed.'
			: 'Added items are pre-filled as they are when the workbook is generated. A versioned item is pre-filled from its current generation.'));
	}
	return parts;
}

/**
 * All the selected items of a target, in a table that scrolls under a fixed header, with a box that
 * filters them by name. While items can be removed, each has a tick box, and Remove Selected removes
 * the ticked ones.
 */
function selectionTable(view, index, ids, nameOf, canRemove) {
	const { doc } = view;
	const make = (tag, attributes, ...children) => element(doc, tag, attributes, ...children);
	const idOf = (base) => sectionId(base, index);
	view.itemFilters = view.itemFilters || new Map();
	view.itemPicks = view.itemPicks || new Map();
	const picks = view.itemPicks.get(index) || new Set();
	view.itemPicks.set(index, picks);
	for (const id of Array.from(picks)) {
		if (!ids.includes(id)) {
			picks.delete(id);
		}
	}
	const rows = [];
	const shown = make('span', { class: 'mce-muted', role: 'status' });
	const removePicked = canRemove
		? panelButton(view, { id: idOf('mce-selection-remove-picked'), label: 'Remove Selected', onClick: () => changeSelection(view, index, { remove: Array.from(picks) }) })
		: null;
	let all = null;
	const update = () => {
		const visible = rows.filter((row) => !row.tr.hidden);
		shown.textContent = visible.length === rows.length ? plural(rows.length, 'item', 'items') : visible.length + ' of ' + rows.length + ' items';
		if (removePicked) {
			removePicked.disabled = picks.size === 0;
			removePicked.querySelector('.aras-button__text').textContent = picks.size > 0 ? 'Remove ' + picks.size + ' Selected' : 'Remove Selected';
		}
		if (all) {
			all.input.checked = visible.length > 0 && visible.every((row) => picks.has(row.id));
		}
	};
	if (canRemove) {
		all = checkbox(doc, {
			id: idOf('mce-selection-pick-all'),
			onChange: (on) => {
				for (const row of rows) {
					if (!row.tr.hidden) {
						row.input.checked = on;
						if (on) {
							picks.add(row.id);
						} else {
							picks.delete(row.id);
						}
					}
				}
				update();
			}
		});
		all.label.title = 'Select all shown';
	}
	const body = make('tbody');
	ids.forEach((id, at) => {
		const { text, missing } = nameOf(id);
		const cells = [];
		let input = null;
		if (canRemove) {
			const box = checkbox(doc, {
				id: idOf('mce-selection-pick') + '-' + at, checked: picks.has(id),
				onChange: (on) => {
					if (on) {
						picks.add(id);
					} else {
						picks.delete(id);
					}
					update();
				}
			});
			box.input.setAttribute('aria-label', 'Select ' + text);
			input = box.input;
			cells.push(make('td', { class: 'aras-grid-row-cell mce-table-check' }, box.label));
		}
		cells.push(make('td', { class: 'aras-grid-row-cell' + (missing ? ' mce-cell-empty' : ''), title: missing ? id : text }, text));
		const tr = make('tr', null, ...cells);
		rows.push({ id, tr, input, text: text.toLowerCase() });
		body.append(tr);
	});
	const filter = make('input', {
		type: 'search', id: idOf('mce-selection-filter'), class: 'aras-input mce-selection-filter', placeholder: 'Filter items',
		'aria-label': 'Filter the selected items', value: view.itemFilters.get(index) || ''
	});
	// Typing filters the table in place: drawing the tool again would lose the caret.
	const apply = () => {
		const wanted = filter.value.trim().toLowerCase();
		view.itemFilters.set(index, filter.value);
		for (const row of rows) {
			row.tr.hidden = wanted !== '' && !row.text.includes(wanted);
		}
		update();
	};
	filter.addEventListener('input', apply);
	apply();
	const head = make('tr', null, all ? make('th', { class: 'aras-grid-head-cell mce-table-check' }, all.label) : null, make('th', { class: 'aras-grid-head-cell' }, 'Item'));
	return make('div', { class: 'mce-selection-all' },
		make('div', { class: 'mce-line' }, filter, shown, make('span', { class: 'mce-gap' }), removePicked),
		make('div', { class: 'mce-table-box' }, make('table', { class: 'mce-table' }, make('thead', null, head), body)));
}

// ---------------------------------------------------------------------------
// The user's changes
// ---------------------------------------------------------------------------

/** Removes or adds selected items of a target, through the Form like a field does: saving the job keeps the change. */
function changeSelection(view, index, edit) {
	const { job } = view.model;
	if (!job.editing || job.state !== 'Draft' || !job.targets[index]) {
		return;
	}
	const targets = copyTargets(job);
	targets[index].prefillIds = logic.editSelection(targets[index].prefillIds, edit);
	const memory = memoryOf(view.jobId);
	if (memory.flow && memory.flow.kind !== 'busy') {
		memory.flow = null;
	}
	setTargets(view, targets);
	view.model.job = readJob(view);
	paint(view);
	loadSelectionNames(view, view.model);
}

/** Opens the client's search dialog on a target's ItemType and adds the items picked there. */
async function addItems(view, index) {
	const { model, top } = view;
	const section = model.sections[index];
	if (!section || !section.target) {
		return;
	}
	let picked;
	try {
		const dialog = top.ArasModules.MaximazableDialog.show('iframe', {
			aras: top.aras, type: 'SearchDialog', itemtypeName: section.target.name, multiselect: true, fullMultiResponse: true
		});
		picked = await dialog.promise;
	} catch (error) {
		return;
	}
	if (!Array.isArray(picked) || picked.length === 0 || memoryOf(view.jobId).view !== view || view.model !== model) {
		return;
	}
	const names = memoryOf(view.jobId).names;
	const ids = [];
	for (const entry of picked) {
		const id = entry && (entry.itemID || (entry.item && entry.item.getAttribute && entry.item.getAttribute('id')));
		if (id) {
			ids.push(id);
			if (entry.keyed_name) {
				names.set(id, entry.keyed_name);
			}
		}
	}
	changeSelection(view, index, { add: ids });
}

/**
 * Sets the ItemType of a target: a target after the last adds one, and an empty ItemType on
 * the first, while it is the only one, leaves the job without a target.
 */
function changeTarget(view, index, itemTypeId) {
	const { job } = view.model;
	if (!job.editing || job.state !== 'Draft') {
		return;
	}
	const memory = memoryOf(view.jobId);
	memory.flow = null;
	memory.adding = false;
	const targets = copyTargets(job);
	const id = String(itemTypeId || '').toUpperCase();
	if (targets.some((target, at) => at !== index && target.itemTypeId === id) || (id === '' && (index !== 0 || targets.length > 1))) {
		return;
	}
	if (index === 0) {
		// The sheets that were picked belong to the ItemType that was chosen before. The stored
		// options of the job are not cleared here: they are the server's, and the options sent
		// when generating replace them.
		memory.pending = null;
		view.openKeys.clear();
	}
	if (id === '') {
		targets.length = 0;
	} else if (index < targets.length) {
		targets[index] = { itemTypeId: id, prefillIds: [] };
	} else {
		targets.push({ itemTypeId: id, prefillIds: [] });
	}
	setTargets(view, targets);
	render(view);
}

/**
 * Removes a target after the first, with its relationship sheets. With selected items, an Aras
 * dialog asks first. A new target whose ItemType is not chosen yet just goes.
 */
async function removeTarget(view, index) {
	const { job } = view.model;
	const memory = memoryOf(view.jobId);
	if (index >= job.targets.length) {
		memory.adding = false;
		paint(view);
		return;
	}
	if (index === 0 || !job.editing || job.state !== 'Draft') {
		return;
	}
	const count = job.targets[index].prefillIds.length;
	if (count > 0) {
		const aras = view.top.aras;
		const modules = (view.itemWindow && view.itemWindow.ArasModules) || view.top.ArasModules;
		const label = targetLabel(view.model, index);
		const answer = await modules.Dialog.confirm(
			'Target ' + label + ' has ' + plural(count, 'selected item', 'selected items') + '. Removing the target takes ' +
			(count === 1 ? 'it' : 'them') + ' and its relationship sheets out of the workbook. The items themselves are not changed.',
			{ title: 'Remove target ' + label + '?', okButtonText: 'Remove', cancelButtonText: aras.getResource('', 'common.cancel') }
		);
		if (answer !== 'ok' || memory.view !== view) {
			return;
		}
	}
	const targets = copyTargets(readJob(view));
	targets.splice(index, 1);
	memory.flow = null;
	setTargets(view, targets);
	render(view);
}

function toggleSheet(view, choice, on) {
	const sheets = view.model.options.sheets.filter((sheet) => sheet.relationshipTypeId !== choice.relationshipTypeId);
	if (on) {
		sheets.push({ relationshipTypeId: choice.relationshipTypeId, name: choice.name, matchKey: logic.defaultMatchKey(choice) });
	}
	keepOptions(view, tidyAll(view.model, { ...view.model.options, sheets }));
}

function changeSheet(view, choice, matchKey) {
	const sheets = view.model.options.sheets.map((sheet) =>
		sheet.relationshipTypeId === choice.relationshipTypeId ? { ...sheet, name: choice.name, matchKey } : sheet);
	keepOptions(view, { ...view.model.options, sheets });
}

/** Writes a relationship sheet as an indented structure, or as ordinary rows again. */
function changeLayout(view, choice, indented) {
	const sheets = view.model.options.sheets.map((sheet) => {
		if (sheet.relationshipTypeId !== choice.relationshipTypeId) {
			return sheet;
		}
		const { layout, ...rest } = sheet;
		return indented ? { ...rest, layout: logic.INDENTED_LAYOUT, matchKey: logic.indentedMatchKey(choice) } : { ...rest, matchKey: logic.defaultMatchKey(choice) };
	});
	keepOptions(view, { ...view.model.options, sheets });
}

/** A part of the options with its columns left out set: without omit when none are. */
const withColumnsLeftOut = (part, omit) => {
	const copy = { ...part };
	delete copy.omit;
	return omit.length > 0 ? { ...copy, omit } : copy;
};

/**
 * Sets the columns left out of a relationship sheet. A locked column that omit names stays
 * in the workbook anyway.
 */
function changeColumns(view, relationshipTypeId, omit) {
	const options = view.model.options;
	keepOptions(view, { ...options, sheets: options.sheets.map((sheet) => (sheet.relationshipTypeId === relationshipTypeId ? withColumnsLeftOut(sheet, omit) : sheet)) });
}

/** Sets the columns left out of the main sheet of a target: the first target's are on the options themselves. */
function changeMainColumns(view, index, omit) {
	const options = view.model.options;
	if (index === 0) {
		const changed = withColumnsLeftOut(options, omit);
		keepOptions(view, changed);
		return;
	}
	const itemTypeId = view.model.sections[index].itemTypeId;
	const targets = (options.targets || []).slice();
	const at = targets.findIndex((target) => target.itemTypeId === itemTypeId);
	const entry = withColumnsLeftOut(at >= 0 ? targets[at] : { itemTypeId }, omit);
	if (at >= 0) {
		targets[at] = entry;
	} else {
		targets.push(entry);
	}
	keepOptions(view, { ...options, targets });
}

/** Keeps changed export options in memory until generating hands them to the server. */
function keepOptions(view, options) {
	const { job } = view.model;
	const memory = memoryOf(view.jobId);
	memory.pending = { base: job.base, text: logic.serializeOptions(options) };
	if (memory.flow && memory.flow.kind !== 'busy') {
		memory.flow = null;
	}
	view.model.options = options;
	paint(view);
}

// ---------------------------------------------------------------------------
// Generating the workbook
// ---------------------------------------------------------------------------

/** Draws the job's sections again, on the Form that shows the job now. */
function repaint(jobId) {
	const view = memoryOf(jobId).view;
	if (!view || view.window.closed) {
		return;
	}
	if (view.model && view.doc.getElementById(HOST_ID) === view.host) {
		paint(view);
	}
	if (view.parts && view.importModel && view.doc.contains(view.parts.top)) {
		drawImport(view);
	}
}

const plural = (count, one, many) => count + ' ' + (count === 1 ? one : many);

/**
 * Generating again deletes the Files of the job's earlier workbook. When the job holds an
 * upload, the user's own work, an Aras dialog asks first and names what is deleted.
 * Resolves true when generating goes ahead.
 */
async function confirmDeletingUpload(view) {
	const job = readJob(view);
	if (job.state === 'Draft' || (!job.uploadFile && !job.resultsFile)) {
		return true;
	}
	const files = [];
	if (job.exportFile) {
		files.push('the earlier workbook (Generated Workbook)');
	}
	if (job.uploadFile) {
		files.push('the uploaded workbook');
	}
	if (job.resultsFile) {
		files.push('its validation result');
	}
	const listed = files.length === 1 ? files[0] : files.slice(0, -1).join(', ') + ' and ' + files[files.length - 1];
	const aras = view.top.aras;
	const modules = (view.itemWindow && view.itemWindow.ArasModules) || view.top.ArasModules;
	const answer = await modules.Dialog.confirm(
		'Generating the workbook again deletes ' + listed + ' of job ' + job.number + ' from Aras, and clears the import preview. ' +
		'Copies you downloaded are not affected. Are you sure?',
		{ title: 'Delete the uploaded workbook?', okButtonText: 'Delete and generate', cancelButtonText: aras.getResource('', 'common.cancel') }
	);
	return answer === 'ok';
}

async function generate(jobId) {
	const memory = memoryOf(jobId);
	if ((memory.flow && memory.flow.kind === 'busy') || importBusy(memory)) {
		return;
	}
	const view = memory.view;
	const top = view.top;
	const aras = top.aras;
	if (memory.askingToGenerate) {
		return;
	}
	memory.askingToGenerate = true;
	let go = false;
	try {
		go = await confirmDeletingUpload(view);
	} finally {
		memory.askingToGenerate = false;
	}
	if (!go) {
		return;
	}
	const say = (message) => {
		memory.flow = { kind: 'busy', message };
		repaint(jobId);
	};
	// A new target whose ItemType was not chosen is not part of the job.
	memory.adding = false;
	let begun = null;
	try {
		const optionsText = logic.serializeOptions(view.model.options);

		// 1. The server works on the saved job, and refuses a job that is being edited.
		await finishEditing(view, say, 'no workbook was generated');

		// 2. The server checks the job, stores the options, issues the token and says what the workbook holds.
		say('Preparing the workbook...');
		begun = await callMethod(top, EXPORT_METHOD, { job_id: jobId, step: 'begin', options: optionsText });
		// Begin cleared the upload, the preview and the counts of the job: what the import section held is gone.
		const imp = importMemoryOf(jobId);
		imp.preview = null;
		imp.flow = null;
		imp.confirming = false;
		// The workbook gets the columns the options keep; the upload is checked against the whole schema.
		const schema = logic.narrowSchema(begun.schema, view.model.options);
		// The main sheets come first, one for each target, in the order of the targets.
		const mainSheets = schema.sheets.filter((sheet) => sheet.kind === 'main');
		const targets = begun.targets;
		const several = mainSheets.length > 1;
		const workbook = await loadWorkbook();
		const notes = [];

		// 3. The rows, read as the user: the selected items of each target, and the relationship rows under them.
		const rows = [];
		const counts = [];
		const parentsOf = [];
		for (const [index, mainSheet] of mainSheets.entries()) {
			const ids = targets[index].prefillIds;
			const label = mainSheet.itemType.label || mainSheet.itemType.name;
			let parents = [];
			if (ids.length > 0) {
				say('Reading ' + plural(ids.length, 'item', 'items') + (several ? ' of ' + label : '') + '...');
				parents = await readMainItems(top, mainSheet, ids, notes, several ? label : null);
			}
			parentsOf.push(parents);
			rows.push(logic.mainRows(mainSheet, parents));
			counts.push(plural(rows[index].length, 'row', 'rows') + ' of ' + label);
		}
		for (const sheet of schema.sheets.slice(mainSheets.length)) {
			const label = sheet.relationshipType.label || sheet.relationshipType.name;
			const parents = parentsOf[workbook.parentSheetOf(schema, sheet)] || [];
			let sheetRows = [];
			if (parents.length > 0 && !sheet.createOnly) {
				say('Reading the rows of ' + label + '...');
				sheetRows = sheet.layout
					? await readIndentedRows(top, sheet, parents, schema.sheets[workbook.parentSheetOf(schema, sheet)].itemType.name, label, notes)
					: await readRelationshipRows(top, sheet, parents, label, notes);
			}
			rows.push(sheetRows);
			counts.push(plural(sheetRows.length, 'row', 'rows') + ' of ' + label);
		}

		// 4. The file.
		say('Building the workbook...');
		const buffer = await workbook.buildWorkbook(schema, { job: { id: jobId, token: begun.token }, rows, maxRows: begun.limits.maxRows || undefined });
		const megabytes = buffer.byteLength / (1024 * 1024);
		if (begun.limits.maxFileSizeMb > 0 && megabytes > begun.limits.maxFileSizeMb) {
			throw new Error('The workbook is ' + megabytes.toFixed(1) + ' MB, and an upload takes at most ' + begun.limits.maxFileSizeMb +
				' MB: it could not be uploaded again. Start the job from a smaller selection, or with fewer relationship sheets.');
		}
		const fileName = logic.workbookFileName(begun.job.number);

		// 5. and 6. The file goes into the vault, and the server attaches it to the job.
		say('Storing the workbook...');
		const fileId = await storeFile(top, fileName, buffer);
		await callMethod(top, EXPORT_METHOD, { job_id: jobId, step: 'complete', token: begun.token, file_id: fileId });

		// 7.
		download(top, buffer, fileName);
		memory.pending = null;
		// The browser may hold back a download that a page starts by itself, for instance a second one.
		// The workbook is attached to the job in any case.
		memory.flow = {
			kind: 'done', at: Date.now(), fileId, notes,
			message: 'The workbook ' + fileName + ' was generated: ' + counts.join(', ') + '. ' +
				'If your browser did not download it, open it from the Generated Workbook field of the job.'
		};
	} catch (error) {
		const message = error && error.message ? error.message : String(error);
		memory.flow = {
			kind: 'error', at: Date.now(),
			message: begun ? 'The workbook was not generated: ' + message + ' The job has no workbook now; generate the workbook again.' : message
		};
	}
	// The job changed on the server from step 2 on: the item window shows it again as it is now.
	const itemWindow = begun ? aras.uiFindWindowEx(jobId) : null;
	if (itemWindow && typeof itemWindow.onRefresh === 'function') {
		try {
			await itemWindow.onRefresh();
		} catch (error) {
			repaint(jobId);
		}
	} else {
		repaint(jobId);
	}
}

/**
 * The selected items of a main sheet, in their current generation and in the order they were selected in.
 * label names the sheet's ItemType in a note, when the workbook has several main sheets.
 */
async function readMainItems(top, mainSheet, selectedIds, notes, label = null) {
	const typeName = mainSheet.itemType.name;
	let configIdOf = null;
	let keys = selectedIds;
	if (mainSheet.itemType.isVersionable) {
		configIdOf = new Map();
		for (const ids of logic.chunk(selectedIds)) {
			for (const item of await query(top, logic.configIdsAml(typeName, ids))) {
				configIdOf.set(item.id, textOf(item, 'config_id'));
			}
		}
		keys = Array.from(new Set(configIdOf.values()));
	}
	const items = [];
	for (const part of logic.chunk(keys)) {
		items.push(...(await query(top, logic.mainItemsAml(mainSheet, part))));
	}
	const ordered = logic.orderMainItems(selectedIds, configIdOf, items);
	if (ordered.missing > 0) {
		notes.push((label ? label + ': ' : '') + plural(ordered.missing, 'selected item was', 'selected items were') +
			' not pre-filled: deleted since the job was started, or not readable for you.');
	}
	return ordered.items;
}

/** The existing rows of a relationship sheet under the pre-filled items. */
async function readRelationshipRows(top, sheet, parents, label, notes) {
	const relationships = [];
	for (const part of logic.chunk(parents.map((parent) => parent.id))) {
		relationships.push(...(await query(top, logic.relationshipRowsAml(sheet, part))));
	}
	const read = logic.relationshipRows(sheet, parents, relationships);
	if (read.rows.length > logic.SHEET_ROW_LIMIT) {
		throw new Error(label + ' has ' + read.rows.length + ' rows under the selected items, more than a sheet can hold. Start the job from a smaller selection.');
	}
	if (read.unreadable > 0) {
		notes.push(label + ': ' + plural(read.unreadable, 'row was', 'rows were') + ' left out, because you can\'t read the related item.');
	}
	if (read.withoutRelated > 0) {
		notes.push(label + ': ' + plural(read.withoutRelated, 'row', 'rows') +
			' without a related item ' + (read.withoutRelated === 1 ? 'was' : 'were') + ' left out. The related item is part of the match key, so an import could not find such a row again.');
	}
	return read.rows;
}

/** The rows of an indented sheet: the selected items, with their lines at every level. */
async function readIndentedRows(top, sheet, parents, relatedTypeName, label, notes) {
	const structure = await logic.readIndentedStructure((aml) => query(top, aml), sheet, parents, relatedTypeName);
	const read = logic.indentedRows(sheet, parents, structure);
	if (read.tooMany) {
		throw new Error(label + ' has more rows at all levels under the selected items than a sheet can hold. Start the job from a smaller selection.');
	}
	if (read.unreadable > 0) {
		notes.push(label + ': ' + plural(read.unreadable, 'line was', 'lines were') + ' left out, because you can\'t read the part.');
	}
	if (read.withoutRelated > 0) {
		notes.push(label + ': ' + plural(read.withoutRelated, 'line', 'lines') + ' without a part ' + (read.withoutRelated === 1 ? 'was' : 'were') + ' left out.');
	}
	if (read.loops > 0) {
		notes.push(label + ': ' + plural(read.loops, 'line leads', 'lines lead') + ' back to a part above it in Aras. The structure stops there, and an upload refuses such a line unless it is marked #DELETE.');
	}
	if (read.sameLine > 0) {
		notes.push(label + ': ' + plural(read.sameLine, 'parent lists', 'parents list') + ' a line the match columns can\'t tell apart from another, so an upload refuses those rows. Add Sequence to the columns that match, and generate again.');
	}
	return read.rows;
}

/** Stores a file (the workbook, unless a type is given) in the vault as a new File item and returns the File's id. */
async function storeFile(top, fileName, buffer, type = XLSX_TYPE) {
	const aras = top.aras;
	const fileId = aras.generateNewGUID();
	const file = new top.File([buffer], fileName, { type });
	const xml = aras.createXMLDocument();
	xml.loadXML(
		'<Item type="File" action="add" id="' + fileId + '">' +
		'<filename>' + escapeXml(fileName) + '</filename>' +
		'<file_size>' + file.size + '</file_size>' +
		'<Relationships><Item type="Located" action="add"><related_id>' + aras.getVaultServerID() + '</related_id></Item></Relationships>' +
		'</Item>'
	);
	aras.vault.addFileToList(fileId, file);
	const stored = await aras.sendFilesWithVaultAppletAsync(xml.documentElement, '');
	if (!stored) {
		throw new Error('The file ' + fileName + ' could not be stored in the vault.');
	}
	return fileId;
}

/** Hands a file (the workbook, unless a type is given) to the browser as a download. */
function download(top, buffer, fileName, type = XLSX_TYPE) {
	const doc = top.document;
	const url = top.URL.createObjectURL(new top.Blob([buffer], { type }));
	const link = doc.createElement('a');
	link.href = url;
	link.download = fileName;
	link.style.display = 'none';
	doc.body.append(link);
	link.click();
	link.remove();
	top.setTimeout(() => top.URL.revokeObjectURL(url), 60000);
}

// ---------------------------------------------------------------------------
// The import section
// ---------------------------------------------------------------------------
//
// Uploading a workbook (job Generated or Validated, run by its creator):
//   1. The job is not saved: after Draft it is changed by the server Methods only.
//      A job still locked for editing is refused, with the reason.
//   2. The user picks the workbook together with the files its file cells name,
//      or a folder that holds them all, or drops them on the section. The
//      workbook is the one .xlsx file, or the one generated for this job. It is
//      checked against the upload size limit and read here with parseWorkbook.
//      The MD5 of each other file is computed here (mce-md5.js); the files are
//      not stored yet.
//   3. Server Method FRG_MCE_ValidateImport checks every row, with the list of
//      the picked files (path, size, MD5). An upload it refuses changes nothing;
//      its reasons are shown. The list is kept in the validation result.
//   4. The workbook and the validation result are stored in the vault as two
//      File items, and server Method FRG_MCE_ImportWorkbook, step "upload",
//      attaches them to the job and moves the job to Validated.
//   5. The preview shows the result: a grid per sheet beside the item preview.
//      A Validated job that is opened again shows the stored result.
//
// Importing, after the user confirms:
//   1. FRG_MCE_ImportWorkbook, step "begin": checks the job, gives the batch size.
//   2. The uploaded workbook is read again, unless it is still in memory. The
//      picked files are kept in memory only: a job opened again asks for them
//      again, and takes only files with the MD5 the preview was made with.
//   3. Before each batch, the files its rows name that are not stored yet are
//      stored in the vault as File items, and step "files" adds them to the
//      job. The server links a file cell to the File with the MD5 and name the
//      preview showed, and deletes the stored Files no row linked at the end.
//      The rows to write are sent in batches, one request each (step "batch"),
//      with the digest of every row as the preview showed it. The server checks
//      the rows again and writes those that are still as reviewed. When a write
//      fails, the request is rolled back and the server names the row: that row
//      is marked failed, and the batch is sent again without it.
//   4. Step "complete" records the failed rows and the counts, and moves the job
//      to Imported. An import that stops half way leaves the job Validated.

const IMPORT_HOST_ID = 'mce-import';
const VALIDATE_METHOD = 'FRG_MCE_ValidateImport';
const IMPORT_METHOD = 'FRG_MCE_ImportWorkbook';
const JSON_TYPE = 'application/json';
const ROW_HEIGHT = 26;
const RECORD_ACTIONS = { Created: 'Created', Updated: 'Updated', Deleted: 'Deleted', Failed: 'Failed', Skipped: 'Skipped' };
/** The choices of the result's Show filter: the value, its label, and the action of the rows it shows. */
const RECORD_FILTERS = [
	{ value: 'all', label: 'All rows', action: null },
	{ value: 'created', label: 'Created', action: RECORD_ACTIONS.Created },
	{ value: 'updated', label: 'Updated', action: RECORD_ACTIONS.Updated },
	{ value: 'deleted', label: 'Deleted', action: RECORD_ACTIONS.Deleted },
	{ value: 'failed', label: 'Failed', action: RECORD_ACTIONS.Failed },
	{ value: 'skipped', label: 'Skipped for errors', action: RECORD_ACTIONS.Skipped }
];

let previewLoading = null;
/** The item preview pane. A missing or broken module gives null, and the import works without the pane. */
const loadPreviewPane = () =>
	previewLoading ||
	(previewLoading = import(sibling('./mce-preview.js')).catch(() => {
		previewLoading = null;
		return null;
	}));

/** The functions of mce-import.js, once loaded. */
let imports = null;

/**
 * What is remembered about a job's import while the client is open:
 *   flow        the state of an upload or import: { kind: 'busy' | 'error' | 'rejected' | 'uploaded' | 'done', step, message, at, ... }
 *   preview     the validation result shown: { resultsFileId, uploadFileId, result, workbook, workbookSize, sheetIndex,
 *               filter, selected, picks, scrollTops, gridWidth, columns }
 *   picked      the files picked for the file cells, kept until the import is finished: [ { file, path, md5 } ].
 *               md5 is read when the workbook is checked with them. The files are not stored before the import.
 *   pickedOpen  whether the list of the picked files is open
 *   confirming  the user asked to import and is shown what will happen
 *   resultShow  which rows the result of an imported job shows: all, or one action (RECORD_FILTERS)
 *   resultSheet the sheet whose rows the result shows, or '' for every sheet
 *   resultView  what the Result tab of an imported job shows: 'records', the rows the import wrote, or
 *               'preview', the preview as it was before the import
 */
function importMemoryOf(jobId) {
	const memory = memoryOf(jobId);
	if (!memory.imp) {
		memory.imp = { flow: null, preview: null, picked: [], pickedOpen: null, confirming: false, resultShow: 'all', resultSheet: '', resultView: 'records' };
	}
	return memory.imp;
}

const importBusy = (memory) => !!memory.imp && !!memory.imp.flow && memory.imp.flow.kind === 'busy';

// --- The import Method. What depends on its exact requests and answers is kept here. ---

const importStep = {
	upload: (top, jobId, token, uploadFileId, resultsFileId) =>
		callMethod(top, IMPORT_METHOD, { job_id: jobId, step: 'upload', token, upload_file_id: uploadFileId, results_file_id: resultsFileId }),

	/**
	 * Answers { batchSize, parallelBatches, records: { created, updated, deleted, failed }, files: [ { id, md5, name, linked } ] }:
	 * parallelBatches is how many batches may be sent at once, the records count rows an earlier run wrote, and files
	 * lists the Files stored for the import so far.
	 */
	begin: (top, jobId, token) => callMethod(top, IMPORT_METHOD, { job_id: jobId, step: 'begin', token }),

	/** Adds Files just stored in the vault to the job's stored files. Answers { job, files } as begin does. */
	files: (top, jobId, token, fileIds) => callMethod(top, IMPORT_METHOD, { job_id: jobId, step: 'files', token, file_ids: JSON.stringify(fileIds) }),

	/**
	 * Sends one batch, with the list of the picked files the preview was made with. A failed write rejects with the
	 * error of the request, whose `fault` names the rows. A batch that met another request in the database rejects
	 * with a fault that isRetryFault knows; nothing of it was kept.
	 */
	batch: (top, jobId, slice, files) =>
		callMethod(top, IMPORT_METHOD, { job_id: jobId, step: 'batch', workbook: JSON.stringify(slice), files: JSON.stringify(files || []) }),

	/** Answers { job, counts: { total, created, updated, deleted, unchanged, skipped, failed } }. */
	complete: (top, jobId, token, failedRows, counts) =>
		callMethod(top, IMPORT_METHOD, {
			job_id: jobId, step: 'complete', token,
			failed: JSON.stringify(failedRows),
			counts: JSON.stringify({ total: counts.total, unchanged: counts.unchanged, skipped: counts.skipped })
		})
};

/** How often a request that met another request in the database is sent again. */
const RETRIES = 4;

/** Sends a request, and sends it again after a pause when the database rolled it back for meeting another request. */
async function withRetry(sendRequest) {
	for (let attempt = 1; ; attempt++) {
		try {
			return await sendRequest();
		} catch (error) {
			if (attempt > RETRIES || !imports.isRetryFault(error && error.fault)) {
				throw error;
			}
			await new Promise((resolve) => setTimeout(resolve, imports.retryDelay(attempt)));
		}
	}
}

// --- Reading files of the vault ---

async function readStoredFile(top, fileId) {
	const aras = top.aras;
	const url = aras.IomInnovator.getFileUrl(fileId, aras.Enums.UrlType.SecurityToken);
	const response = await top.fetch(url, { credentials: 'same-origin' });
	if (!response.ok) {
		throw new Error('A file of the job could not be read from the vault (' + response.status + ').');
	}
	return response.arrayBuffer();
}

// --- The section ---

/** Why the user can't upload to or import the job, or null when they can. */
function whyNotImport(view, job) {
	if (!job.isCreator) {
		return 'Only the user who created a job can run it. This job was created by ' + (job.creatorName || 'another user') + '.';
	}
	if (job.lockedByOther) {
		return 'The job is locked by another user.';
	}
	if (!view.itemWindow) {
		return 'Open the job to upload a workbook or to import it.';
	}
	return null;
}

/** Reads what the import section shows and draws it. */
async function renderImport(view) {
	const host = view.doc.getElementById(IMPORT_HOST_ID);
	if (!host) {
		return;
	}
	const memory = memoryOf(view.jobId);
	const imp = importMemoryOf(view.jobId);
	if (imp.flow && imp.flow.kind !== 'busy' && Date.now() - imp.flow.at > MESSAGE_KEPT_MS) {
		imp.flow = null;
	}
	disposePreviewUi(view);
	host.textContent = '';
	host.className = 'mce-job mce-import';
	const make = (tag, attributes, ...children) => element(view.doc, tag, attributes, ...children);
	// top: the Upload card and notes; card: the preview or the result, with head: its title, counts and notes,
	// bottom: the confirmation or progress of an import, and preview: the grid and the item preview.
	view.parts = { top: make('div', { class: 'mce-import-top' }), head: make('div', { class: 'mce-preview-head' }), bottom: make('div'), preview: make('div', { class: 'mce-preview-part' }) };
	view.parts.card = make('section', { class: 'mce-card mce-preview-card' }, view.parts.head, view.parts.bottom, view.parts.preview);
	host.append(view.parts.top, view.parts.card);
	takeDrops(view, host);
	view.importModel = { loading: true, problem: null, config: null, records: null, targets: [] };
	drawImport(view);
	try {
		imports = imports || (await loadImportLogic());
		view.importModel.config = await configOf(view);
		const job = readJob(view);
		if (job.state === 'Validated' && job.resultsFile) {
			await loadPreview(view, job);
		}
		if (job.state === 'Imported') {
			// The validation result stays on the job: it lists the rows skipped for errors, and the
			// Result tab can show the preview as it was before the import. Without it, the result
			// shows the rows written, and no preview.
			if (job.resultsFile) {
				try {
					await loadPreview(view, job);
				} catch (ignored) {
					// Shown without the skipped rows and the preview.
				}
			}
			const preview = importedPreview(view, job);
			const read = resultRecords(await readRecords(view.top, job.id), preview ? preview.result : null);
			view.importModel.records = read.records;
			view.importModel.targets = read.targets;
		}
	} catch (error) {
		view.importModel.problem = error.message;
	}
	view.importModel.loading = false;
	if (memory.view === view) {
		drawImport(view);
	}
}

/** The stored validation result of a Validated or Imported job, kept in memory once read. */
async function loadPreview(view, job) {
	const imp = importMemoryOf(job.id);
	if (imp.preview && imp.preview.resultsFileId === job.resultsFile) {
		return;
	}
	const buffer = await readStoredFile(view.top, job.resultsFile);
	let result;
	try {
		result = JSON.parse(new TextDecoder().decode(buffer));
	} catch (error) {
		throw new Error('The preview stored on the job can\'t be read. Upload the workbook again.');
	}
	imp.preview = newPreview(job.resultsFile, job.uploadFile, result, null, null);
}

function newPreview(resultsFileId, uploadFileId, result, workbook, workbookSize) {
	return {
		resultsFileId, uploadFileId, result, workbook, workbookSize,
		sheetIndex: 0,
		filter: { ...imports.DEFAULT_FILTER },
		selected: null,
		picks: [],
		scrollTops: [],
		gridWidth: null,
		columns: result.sheets.map((sheet) => imports.valueColumns(sheet))
	};
}

/** The validation result of an imported job, as loadPreview read it, or null. */
function importedPreview(view, job) {
	const imp = importMemoryOf(view.jobId);
	return job.resultsFile && imp.preview && imp.preview.resultsFileId === job.resultsFile ? imp.preview : null;
}

/**
 * The rows of an imported job's result: the stored records of the rows written or failed, and
 * the rows skipped because they had errors, which are read from the job's validation result.
 * Without the validation result (null), the skipped rows are left out; the counts still name them.
 * Answers { records, targets }: targets are the job's targets with their sheets (resultTargets),
 * empty without the validation result.
 */
function resultRecords(stored, result) {
	if (result) {
		try {
			const sheetOrder = imports.sheetOrder(result).map((index) => result.sheets[index].sheetName);
			return { records: imports.sortRecords(stored.concat(imports.skippedRecords(result)), sheetOrder), targets: imports.resultTargets(result) };
		} catch (ignored) {
			// A result that can't be read: the written rows are shown without the skipped ones.
		}
	}
	return { records: imports.sortRecords(stored, []), targets: [] };
}

async function readRecords(top, jobId) {
	const records = await query(top,
		'<Item type="FRG_MassCreateEdit_Record" action="get" select="sheet_name,row_number,action,target_type,target_id,target_keyed_name,message,sort_order">' +
		'<source_id>' + jobId + '</source_id></Item>');
	return records.map((record) => ({
		sheet: textOf(record, 'sheet_name'),
		row: Number(textOf(record, 'row_number')) || 0,
		action: textOf(record, 'action'),
		targetType: textOf(record, 'target_type'),
		targetId: textOf(record, 'target_id'),
		name: textOf(record, 'target_keyed_name'),
		message: textOf(record, 'message'),
		order: Number(textOf(record, 'sort_order')) || 0
	}));
}

/** Draws the sections again: the export section locks while an import runs, and the other way round. */
const repaintImport = (jobId) => repaint(jobId);

function drawImport(view) {
	const { doc, shell } = view;
	const make = (tag, attributes, ...children) => element(doc, tag, attributes, ...children);
	const memory = memoryOf(view.jobId);
	const imp = importMemoryOf(view.jobId);
	const flow = imp.flow;
	const busy = importBusy(memory) || (!!memory.flow && memory.flow.kind === 'busy');
	const job = readJob(view);
	const model = view.importModel;
	const top = view.parts.top;
	const bottom = view.parts.bottom;
	const focusId = doc.activeElement && shell.root.contains(doc.activeElement) ? doc.activeElement.id : '';
	top.textContent = '';
	bottom.textContent = '';
	shell.importTools.textContent = '';
	shell.importMessages.textContent = '';
	shell.importStatus.textContent = '';

	const ready = !model.loading && !model.problem;
	const refusal = ready ? whyNotImport(view, job) : null;
	const validated = job.state === 'Validated';
	const preview = ready && validated && imp.preview && imp.preview.resultsFileId === job.resultsFile ? imp.preview : null;
	const counts = preview ? imports.importCounts(preview.result) : null;
	const toWrite = counts ? counts.create + counts.update + counts.delete : 0;
	const uploadable = ready && !job.isNew && ((job.state === 'Generated' && job.exportFile !== '') || validated);

	// --- Toolbar: uploading and importing ---
	// The workbook, the files its file cells name, or both; or a folder. Picked files add to those picked before.
	const input = filePicker(view, 'mce-upload-file', false, busy, true);
	const folderInput = filePicker(view, 'mce-upload-folder', true, busy, true);
	const whyNoUpload = model.loading ? 'Loading...' : !uploadable
		? (job.state === 'Imported' ? 'The job is Imported: it can\'t be uploaded again.' : 'Generate the workbook first.')
		: refusal;
	shell.importTools.append(toolButton(view, {
		id: 'mce-upload', iconName: 'upload', label: validated ? 'Upload Again...' : 'Upload Workbook...',
		primary: uploadable && !validated && !refusal, disabled: busy || !!whyNoUpload,
		title: whyNoUpload || 'Pick the filled-in workbook, the files its file cells name, or both, to see what an import would do. Nothing is written yet.',
		onClick: () => input.click()
	}), input, folderInput);
	const whyNoImport = !preview ? 'Upload the workbook first: the import writes the rows its preview shows.'
		: refusal || (toWrite === 0 ? 'Nothing to import: every row is unchanged or has errors.' : null);
	shell.importTools.append(toolButton(view, {
		id: 'mce-import-start', iconName: 'import', label: 'Import...', primary: !!preview && !whyNoImport,
		disabled: busy || !!whyNoImport || imp.confirming,
		title: whyNoImport || 'See what the import will do, then confirm it.',
		onClick: () => {
			imp.confirming = true;
			if (imp.flow && imp.flow.kind !== 'busy') {
				imp.flow = null;
			}
			memory.tab = { tab: 'import', state: job.state };
			drawImport(view);
			drawChrome(view);
		}
	}));
	if (flow && flow.kind === 'busy') {
		shell.importStatus.append(make('span', { class: 'mce-spinner', 'aria-hidden': 'true' }), flow.message);
	}

	// --- Messages of the last run ---
	if (flow && flow.kind === 'error') {
		shell.importMessages.append(make('div', { class: 'mce-message mce-message-error', role: 'alert' }, flow.message));
	}
	if (flow && flow.kind === 'rejected') {
		shell.importMessages.append(issueBlock(view, flow.message, flow.issues, true));
	}
	if (flow && flow.kind === 'done') {
		shell.importMessages.append(make('div', { class: 'mce-message mce-message-done', role: 'status' }, flow.message));
	}

	const head = view.parts.head;
	head.textContent = '';
	const finish = () => {
		// The card of the preview or the result is shown when it has something to show.
		view.parts.card.hidden = head.childElementCount === 0 && view.parts.preview.childElementCount === 0 && bottom.childElementCount === 0;
		restoreFocus(doc, focusId);
		fitFrame(view);
	};
	if (model.loading) {
		top.append(make('p', { class: 'mce-muted' }, 'Loading...'));
		finish();
		return;
	}
	if (model.problem) {
		top.append(make('p', { class: 'mce-error', role: 'alert' }, model.problem));
		clearPreviewPart(view);
		finish();
		return;
	}
	if (job.isNew || job.state === 'Draft' || (job.state === 'Generated' && job.exportFile === '')) {
		top.append(make('p', { class: 'mce-note' },
			'Generate the workbook on the Workbook step first. Fill it in, then upload it here to see what an import would do.'));
		clearPreviewPart(view);
		finish();
		return;
	}
	if (job.state === 'Imported') {
		head.append(make('h3', { class: 'mce-card-title' }, 'Import Result'));
		drawImported(view, job);
		finish();
		return;
	}

	// --- Upload ---
	if (refusal) {
		top.append(make('p', { class: 'mce-note' }, refusal));
	}
	// The limits of the configuration item: rows of the main sheet (the items to create or edit) and file size.
	const limits = [];
	if (model.config.maxRows > 0) {
		limits.push(model.config.maxRows.toLocaleString() + (job.targets.length > 1 ? ' items (rows of the main sheets together)' : ' items (rows of the main sheet)'));
	}
	if (model.config.maxFileSizeMb > 0) {
		limits.push(model.config.maxFileSizeMb.toLocaleString() + ' MB');
	}
	const uploadNote = 'The workbook generated for this job, filled in, and the files its file cells name, or a folder that holds them. ' +
		(limits.length > 0 ? 'One upload takes at most ' + limits.join(' and ') + '. ' : '') +
		'Nothing is written before you confirm the import.';
	const picked = pickedPanel(view, preview, busy || !!refusal);
	if (!refusal || preview) {
		top.append(uploadCard(view, { job, preview, input, folderInput, uploadNote, picked, disabled: busy || !!refusal, validated }));
	}

	// --- Preview ---
	if (!preview) {
		clearPreviewPart(view);
		finish();
		return;
	}
	head.append(make('h3', { class: 'mce-card-title' }, 'Import Preview'), countsRow(view, counts));
	if (!refusal) {
		head.append(make('p', { class: 'mce-help' },
			'Nothing is written before you confirm the import. Upload again to replace this preview: the earlier upload is then deleted from Aras, and you are asked first.'));
	}
	if (counts.skipped > 0) {
		head.append(make('div', { class: 'mce-message mce-message-warning', role: 'note' },
			make('strong', null, plural(counts.skipped, 'row has errors', 'rows have errors')), ' and ' + (counts.skipped === 1 ? 'is' : 'are') +
			' left out of the import. Correct ' + (counts.skipped === 1 ? 'it' : 'them') + ' in the workbook and upload it again, or import the other rows.'));
	}
	const uploadIssues = preview.result.issues || [];
	if (uploadIssues.length > 0) {
		head.append(issueBlock(view, 'About the workbook as a whole:', uploadIssues, false));
	}
	if (preview.oversized === undefined) {
		// Worked out once per preview: it plans the whole import.
		preview.oversized = imports.oversizedWarning(preview.result, model.config.batchSize);
	}
	if (preview.oversized) {
		head.append(make('div', { class: 'mce-message mce-message-warning', role: 'note' }, preview.oversized));
	}
	drawPreviewPart(view, preview);

	// --- Import ---
	if (!refusal) {
		drawImportActions(view, preview, counts, busy);
	}
	finish();
}

/** What the import would do, counted: a box for each action, in the action's colour. */
function countsRow(view, counts) {
	const make = (tag, attributes, ...children) => element(view.doc, tag, attributes, ...children);
	const box = (action, count) => make('div', { class: 'mce-count mce-count-' + action },
		make('span', { class: 'mce-count-number' }, count.toLocaleString()), make('span', { class: 'mce-count-label' }, imports.ACTION_LABELS[action]));
	return make('div', { class: 'mce-counts-row', role: 'group', 'aria-label': imports.summaryLine(counts) },
		box('new', counts.create), box('update', counts.update), box('delete', counts.delete), box('none', counts.unchanged), box('error', counts.skipped));
}

/**
 * The Upload card. Until a workbook is checked it holds the drop zone, with the commands that pick
 * the workbook, its files or a folder, and below it the files picked so far. Once a preview is
 * shown it is one line: the uploaded workbook, the files picked with it, and Upload Again.
 */
function uploadCard(view, { job, preview, input, folderInput, uploadNote, picked, disabled, validated }) {
	const { doc } = view;
	const make = (tag, attributes, ...children) => element(doc, tag, attributes, ...children);
	const imp = importMemoryOf(view.jobId);
	const card = make('section', { class: 'mce-card mce-upload-card' },
		make('h3', { class: 'mce-card-title' }, 'Upload', imp.picked.length > 0 && !preview ? make('span', { class: 'mce-card-count' }, plural(imp.picked.length, 'file', 'files')) : null));
	if (preview) {
		const line = make('div', { class: 'mce-line' },
			make('span', { class: 'mce-strong' }, job.fileNames.upload || 'The uploaded workbook'),
			picked ? make('span', { class: picked.alert ? 'mce-error-text' : 'mce-muted' }, picked.facts) : null,
			make('span', { class: 'mce-gap' }));
		if (picked) {
			line.append(panelButton(view, {
				id: 'mce-picked-toggle', label: picked.open ? 'Hide Files' : 'Show Files',
				onClick: () => {
					imp.pickedOpen = !picked.open;
					drawImport(view);
				}
			}));
		}
		line.append(panelButton(view, { id: 'mce-upload-again', label: 'Upload Again...', iconName: 'upload', disabled, onClick: () => input.click() }));
		card.append(line);
		if (picked && picked.open) {
			card.append(picked.body);
		}
		return card;
	}
	if (validated) {
		card.append(make('p', { class: 'mce-note' }, 'The job has no preview. Upload the workbook again.'));
	}
	const zoneIcon = icon(doc, 'upload');
	zoneIcon.setAttribute('class', 'mce-icon mce-drop-icon');
	const zone = make('div', { class: 'mce-drop' }, zoneIcon, make('p', { class: 'mce-drop-title' }, 'Drop the filled-in workbook and its files here'));
	if (!disabled) {
		zone.append(make('div', { class: 'mce-line mce-line-center' },
			panelButton(view, { id: 'mce-upload-empty', label: 'Select Files...', primary: true, onClick: () => input.click() }),
			panelButton(view, { id: 'mce-upload-empty-folder', label: 'Select Folder...', onClick: () => folderInput.click() })));
	}
	zone.append(make('p', { class: 'mce-help' }, uploadNote));
	card.append(zone);
	if (imp.picked.length > 0) {
		card.append(make('div', { class: 'mce-message mce-message-warning', role: 'status' }, make('strong', null, 'Now add the workbook. '),
			plural(imp.picked.length, 'file is', 'files are') + ' picked for the file cells; they are checked with the workbook once it is added.'));
	}
	if (picked) {
		card.append(picked.body);
	}
	return card;
}

/** The files of a file input, each with its path: its name, or its path inside a picked folder. */
function pickedFiles(fileList) {
	return Array.from(fileList || []).map((file) => ({ file, path: file.webkitRelativePath || file.name }));
}

/**
 * A hidden file input for the import section: of files, or of a folder. What is picked goes to
 * takePicked(); asWorkbook says whether a lone .xlsx file is taken as the workbook (see pickedWorkbook).
 */
function filePicker(view, id, folder, disabled, asWorkbook) {
	const picker = element(view.doc, 'input', folder
		? { type: 'file', id, class: 'mce-file', webkitdirectory: true, disabled, tabindex: '-1' }
		: { type: 'file', id, class: 'mce-file', multiple: true, disabled, tabindex: '-1' });
	picker.addEventListener('change', () => {
		const picked = pickedFiles(picker.files);
		picker.value = '';
		takePicked(view.jobId, picked, asWorkbook);
	});
	return picker;
}

/**
 * Files and folders dropped on the import section are uploaded as if they were picked, when an upload
 * can start: the Upload command is enabled.
 */
function takeDrops(view, host) {
	const canUpload = () => {
		const button = view.doc.getElementById('mce-upload');
		return !!button && !button.disabled;
	};
	const hasFiles = (event) => !!event.dataTransfer && Array.from(event.dataTransfer.types || []).includes('Files');
	host.addEventListener('dragover', (event) => {
		if (hasFiles(event) && canUpload()) {
			event.preventDefault();
			event.dataTransfer.dropEffect = 'copy';
			host.classList.add('mce-drop-target');
		}
	});
	host.addEventListener('dragleave', (event) => {
		if (!host.contains(event.relatedTarget)) {
			host.classList.remove('mce-drop-target');
		}
	});
	host.addEventListener('drop', (event) => {
		host.classList.remove('mce-drop-target');
		if (!hasFiles(event)) {
			return;
		}
		event.preventDefault();
		if (!canUpload()) {
			return;
		}
		// The entries are taken while the event lasts; their files are read afterwards.
		const items = Array.from(event.dataTransfer.items || []).filter((item) => item.kind === 'file');
		const entries = items.map((item) => (typeof item.webkitGetAsEntry === 'function' ? item.webkitGetAsEntry() : null));
		const plain = pickedFiles(event.dataTransfer.files);
		droppedFiles(entries, plain).then((picked) => takePicked(view.jobId, picked, true), (error) => {
			const imp = importMemoryOf(view.jobId);
			imp.flow = { kind: 'error', step: 'upload', at: Date.now(), message: 'The dropped files could not be read: ' + (error && error.message ? error.message : String(error)) };
			repaint(view.jobId);
		});
	});
}

/** The files of dropped entries, folders read through, each with its path. Without entries, the dropped files as they are. */
async function droppedFiles(entries, plain) {
	if (entries.length === 0 || entries.some((entry) => !entry)) {
		return plain;
	}
	const picked = [];
	const walk = async (entry) => {
		if (entry.isFile) {
			const file = await new Promise((resolve, reject) => entry.file(resolve, reject));
			picked.push({ file, path: String(entry.fullPath || file.name).replace(/^\/+/, '') });
		} else if (entry.isDirectory) {
			const reader = entry.createReader();
			for (;;) {
				const children = await new Promise((resolve, reject) => reader.readEntries(resolve, reject));
				if (children.length === 0) {
					break;
				}
				for (const child of children) {
					await walk(child);
				}
			}
		}
	};
	for (const entry of entries) {
		await walk(entry);
	}
	return picked;
}

/** A list of issues: the reasons an upload was refused, or warnings about it. */
function issueBlock(view, title, issues, refused) {
	const make = (tag, attributes, ...children) => element(view.doc, tag, attributes, ...children);
	const block = make('div', { class: refused ? 'mce-issues mce-issues-error' : 'mce-issues', role: refused ? 'alert' : null },
		make('p', { class: refused ? 'mce-error' : 'mce-label-inline' }, title));
	const list = make('ul');
	for (const issue of issues || []) {
		list.append(make('li', { class: issue.severity === 'error' ? 'mce-error-text' : null }, issue.message));
	}
	block.append(list);
	return block;
}

/** Above the preview: the confirmation of an import the user asked for, or the progress of one that runs. */
function drawImportActions(view, preview, counts, busy) {
	const make = (tag, attributes, ...children) => element(view.doc, tag, attributes, ...children);
	const imp = importMemoryOf(view.jobId);
	const flow = imp.flow;
	const bottom = view.parts.bottom;
	const toWrite = counts.create + counts.update + counts.delete;

	if (flow && flow.kind === 'busy' && flow.step === 'import') {
		const progress = flow.progress || { written: 0, failed: 0, toWrite: 0 };
		const share = progress.toWrite > 0 ? Math.round(((progress.written + progress.failed) / progress.toWrite) * 100) : 0;
		bottom.append(make('div', { class: 'mce-progress-block', role: 'status' },
			make('p', { class: 'mce-status' }, flow.message),
			make('progress', { class: 'mce-progress', max: '100', value: String(share) }),
			make('p', { class: 'mce-muted' }, 'The import goes on when you close the job window, as long as this browser tab stays open.')));
		return;
	}
	if (!imp.confirming || busy) {
		return;
	}
	const table = make('table', { class: 'mce-counts' },
		make('thead', null, make('tr', null, ...['Sheet', 'Created', 'Updated', 'Deleted', 'Skipped (errors)', 'Unchanged'].map((text) => make('th', { class: 'aras-grid-head-cell' }, text)))));
	const body = make('tbody');
	const several = counts.sheets.some((sheet) => sheet.target > 0);
	for (const sheet of counts.sheets) {
		body.append(make('tr', null, make('td', null, targetMark(view.doc, sheet.target, several), sheet.label),
			...[sheet.create, sheet.update, sheet.delete, sheet.skipped, sheet.unchanged].map((count) => make('td', { class: 'mce-number' }, String(count)))));
	}
	if (counts.sheets.length > 1) {
		body.append(make('tr', { class: 'mce-total' }, make('td', null, 'All sheets'),
			...[counts.create, counts.update, counts.delete, counts.skipped, counts.unchanged].map((count) => make('td', { class: 'mce-number' }, String(count)))));
	}
	table.append(body);
	const go = make('button', { id: 'mce-import-confirm', type: 'button', class: 'aras-button aras-button_primary mce-button' },
		make('span', { class: 'aras-button__text' }, 'Import ' + plural(toWrite, 'Row', 'Rows')));
	go.addEventListener('click', () => runImport(view.jobId));
	// The files the rows name are kept in memory only: a job opened again asks for them again.
	const needed = imports.neededFileTokens(preview.result);
	const missing = missingFileTokens(imp, preview);
	go.disabled = missing.length > 0;
	let filesBlock = null;
	if (missing.length > 0) {
		const names = missing.slice(0, 10).map(imports.tokenName);
		filesBlock = make('div', { class: 'mce-message mce-message-info', role: 'status' },
			make('p', null, 'The rows to import name ' + plural(missing.length, 'file', 'files') + ' that ' + (missing.length === 1 ? 'is' : 'are') +
				' not picked in this window: ' + names.join(', ') + (missing.length > names.length ? ', ...' : '') +
				'. Add them under Files above, the same files as for the upload: the workbook is then checked again with them.'));
	}
	const cancel = make('button', { id: 'mce-import-cancel', type: 'button', class: 'aras-button mce-button' },
		make('span', { class: 'aras-button__text' }, 'Cancel'));
	cancel.addEventListener('click', () => {
		imp.confirming = false;
		drawImport(view);
	});
	bottom.append(make('div', { class: 'mce-confirm', role: 'group', 'aria-label': 'Confirm the import' },
		make('div', { class: 'mce-confirm-text' },
			make('p', { class: 'mce-label-inline' }, 'The import will do this:'),
			table,
			make('ul', { class: 'mce-muted mce-consequences' },
				make('li', null, 'Items are written with your own permissions. An update is a normal edit: Aras versions the item as it does for an edit by hand.'),
				make('li', null, 'Rows with errors and unchanged rows are not written. A row that changed in Aras since this preview is not written either; it is reported as failed.'),
				needed.length > 0 ? make('li', null, 'The ' + plural(needed.length, 'file', 'files') + ' the rows name ' + (needed.length === 1 ? 'is' : 'are') +
					' stored in Aras, each before the batch that links it. A stored file that ends up linked by no row is deleted again.') : null,
				counts.delete > 0 ? make('li', null, 'A row marked #DELETE deletes the relationship row it matches. The item or file that row points at stays in Aras.') : null,
				make('li', null, 'The job moves to Imported. It can\'t be uploaded or imported again.'))),
		filesBlock,
		make('div', { class: 'mce-actions' }, go, cancel)));
	if (!go.disabled) {
		go.focus();
	}
}

// --- The picked files ---

/** The picked files whose MD5 is read, by their file token (see mce-import.js). */
function pickedByToken(imp) {
	const files = new Map();
	for (const entry of imp.picked) {
		if (entry.md5) {
			files.set(imports.fileToken(entry.path, entry.md5), entry.file);
		}
	}
	return files;
}

/** The file tokens of the rows to import that no picked file has: a job opened again has no picked files. */
function missingFileTokens(imp, preview) {
	const files = pickedByToken(imp);
	return imports.neededFileTokens(preview.result).filter((token) => !files.has(token));
}

/**
 * Adds files to the picked files. A file with the path of one picked before takes its place, and so
 * does the same file picked again under another path: the same name, size and time of change.
 */
function addPicked(imp, picked) {
	for (const entry of picked) {
		const path = imports.normalPath(entry.path);
		const name = imports.fileNameOf(path).toLowerCase();
		const same = (other) => other.path.toLowerCase() === path.toLowerCase() ||
			(imports.fileNameOf(other.path).toLowerCase() === name && other.file.size === entry.file.size && other.file.lastModified === entry.file.lastModified);
		imp.picked = imp.picked.filter((other) => !same(other));
		imp.picked.push({ file: entry.file, path, md5: null });
	}
}

function sizeText(bytes) {
	if (bytes < 1024) {
		return bytes + ' B';
	}
	if (bytes < 1024 * 1024) {
		return Math.ceil(bytes / 1024) + ' KB';
	}
	return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

/** The preview the import section shows, read; null when there is none. */
function shownPreview(view) {
	const job = readJob(view);
	const imp = importMemoryOf(view.jobId);
	return job.state === 'Validated' && imp.preview && imp.preview.resultsFileId === job.resultsFile ? imp.preview : null;
}

/**
 * The files picked for the file cells: what the preview does with each, the files the rows name that are
 * not picked, and commands that add and remove files. Adding or removing checks the workbook again.
 * Shown when files are picked, or when the workbook has file cells; null otherwise.
 *
 * @returns {{ facts: string, alert: boolean, open: boolean, body: Element } | null} facts: one line about
 *   the files; alert: whether a file is missing; open: whether the list is shown under a preview; body: the list.
 */
function pickedPanel(view, preview, disabled) {
	const make = (tag, attributes, ...children) => element(view.doc, tag, attributes, ...children);
	const imp = importMemoryOf(view.jobId);
	const hasFileCells = !!preview && preview.result.sheets.some((sheet) => (sheet.columns || []).some((column) => column.dataType === 'file'));
	if (imp.picked.length === 0 && !hasFileCells) {
		return null;
	}
	// How many rows of the preview name each file, and how many have an error about a file:
	// one that is not picked, two that fit, or one over the limits.
	const named = new Map();
	let notPicked = 0;
	let fileErrors = 0;
	for (const sheet of preview ? preview.result.sheets : []) {
		for (const row of sheet.rows || []) {
			for (const token of imports.rowFileTokens(row)) {
				named.set(token, (named.get(token) || 0) + 1);
			}
			const codes = (row.issues || []).filter((issue) => issue.severity === 'error').map((issue) => String(issue.code || ''));
			if (codes.includes('FILE_NOT_PICKED')) {
				notPicked++;
			}
			if (codes.some((code) => code.startsWith('FILE_'))) {
				fileErrors++;
			}
		}
	}
	const tokenOf = (entry) => (entry.md5 ? imports.fileToken(entry.path, entry.md5) : null);
	const missing = preview ? missingFileTokens(imp, preview) : [];
	const unused = preview ? imp.picked.filter((entry) => tokenOf(entry) && !named.has(tokenOf(entry))).length : 0;
	const facts = [imp.picked.length === 0 ? 'no files picked'
		: plural(imp.picked.length, 'file', 'files') + ' picked, ' + sizeText(imp.picked.reduce((sum, entry) => sum + entry.file.size, 0))];
	if (missing.length > 0) {
		facts.push(plural(missing.length, 'file', 'files') + ' to pick again');
	}
	if (notPicked > 0) {
		facts.push(plural(notPicked, 'row names a file', 'rows name files') + ' not picked');
	}
	if (unused > 0) {
		facts.push(unused + ' named by no row');
	}
	// Under a preview, the list is open while files are missing or a row has an error about a file.
	// Opening or closing it by hand holds until the next preview (upload() resets it).
	const open = imp.pickedOpen !== null ? imp.pickedOpen : missing.length > 0 || fileErrors > 0;

	const add = filePicker(view, 'mce-picked-add-file', false, disabled, false);
	const addFolder = filePicker(view, 'mce-picked-add-folder', true, disabled, false);
	const body = make('div', { class: 'mce-picked', id: 'mce-picked' });
	body.append(make('div', { class: 'mce-line' },
		make('span', { class: 'mce-strong' }, 'Files'), make('span', { class: missing.length > 0 || notPicked > 0 ? 'mce-error-text' : 'mce-muted' }, facts.join(', ')),
		make('span', { class: 'mce-gap' }),
		preview ? panelButton(view, { id: 'mce-picked-add', label: 'Add Files...', disabled, onClick: () => add.click() }) : null,
		preview ? panelButton(view, { id: 'mce-picked-add-dir', label: 'Add Folder...', disabled, onClick: () => addFolder.click() }) : null,
		panelButton(view, {
			id: 'mce-picked-copy', label: 'Copy All', disabled: imp.picked.length === 0,
			title: 'Copy the File column, one file to a line, to paste it down a column of file cells',
			onClick: () => copyText(view, imp.picked.map((entry) => entry.path).join('\r\n'),
				plural(imp.picked.length, 'file name', 'file names') + ' copied.')
		}),
		panelButton(view, { id: 'mce-picked-clear', label: 'Remove All', disabled: disabled || imp.picked.length === 0, onClick: () => removePicked(view.jobId, imp.picked.slice()) }),
		add, addFolder));
	body.append(make('p', { class: 'mce-help' },
		'A file cell names a picked file by its name, or by its path inside a picked folder. The files are read in this window only: ' +
		'the import stores in Aras those its rows name.' + (preview ? ' Adding or removing files checks the workbook again.' : '')));

	if (imp.picked.length === 0 && missing.length === 0) {
		return { facts: facts.join(', '), alert: false, open, body };
	}
	const rows = make('tbody');
	for (const entry of imp.picked) {
		const token = tokenOf(entry);
		const count = token ? named.get(token) || 0 : 0;
		const copy = panelButton(view, { label: 'Copy', title: 'Copy ' + entry.path + ' to paste it into a file cell', onClick: () => copyText(view, entry.path, entry.path + ' copied.') });
		const remove = panelButton(view, { label: 'Remove', disabled, title: 'Remove ' + entry.path + ' from the picked files', onClick: () => removePicked(view.jobId, [entry]) });
		rows.append(make('tr', null,
			make('td', { class: 'aras-grid-row-cell mce-picked-path', title: entry.path }, entry.path),
			make('td', { class: 'aras-grid-row-cell mce-number' }, sizeText(entry.file.size)),
			make('td', { class: 'aras-grid-row-cell' + (count === 0 && preview && token ? ' mce-muted' : '') },
				!preview || !token ? 'Not checked yet' : count > 0 ? 'Named by ' + plural(count, 'row', 'rows') : 'Named by no row'),
			make('td', { class: 'aras-grid-row-cell mce-table-action' }, copy, remove)));
	}
	for (const token of missing) {
		rows.append(make('tr', null,
			make('td', { class: 'aras-grid-row-cell mce-picked-path' }, imports.tokenName(token)),
			make('td', { class: 'aras-grid-row-cell mce-number' }, ''),
			make('td', { class: 'aras-grid-row-cell mce-error-text', title: 'Not picked in this window: add it again to import its rows' }, 'Not picked in this window: add it again to import its rows'),
			make('td', { class: 'aras-grid-row-cell' })));
	}
	const th = (text, className) => make('th', { class: 'aras-grid-head-cell' + (className ? ' ' + className : '') }, text);
	body.append(make('div', { class: 'mce-table-box' }, make('table', { class: 'mce-table mce-picked-table' },
		make('thead', null, make('tr', null, th('File'), th('Size', 'mce-number mce-col-size'), th(preview ? 'In the preview' : ''), th('', 'mce-col-action'))),
		rows)));
	return { facts: facts.join(', '), alert: missing.length > 0 || notPicked > 0, open, body };
}

/**
 * Takes files picked or dropped on the import section. The workbook generated for this job, when it is
 * among them, is uploaded and checked with every file picked so far; the other files are added to the
 * picked files. Without a workbook, the files are added, and the preview, when there is one, is made again
 * with them; with no preview, they wait for the workbook.
 */
async function takePicked(jobId, picked, asWorkbook) {
	const memory = memoryOf(jobId);
	const imp = importMemoryOf(jobId);
	if (picked.length === 0 || !memory.view || importBusy(memory) || (memory.flow && memory.flow.kind === 'busy')) {
		return;
	}
	imp.confirming = false;
	let book;
	try {
		book = await pickedWorkbook(memory.view, picked, asWorkbook, (message) => {
			imp.flow = { kind: 'busy', step: 'upload', message };
			repaintImport(jobId);
		});
	} catch (error) {
		imp.flow = { kind: 'error', step: 'upload', at: Date.now(), message: error && error.message ? error.message : String(error) };
		repaintImport(jobId);
		return;
	}
	addPicked(imp, picked.filter((entry) => !book || entry !== book.entry));
	imp.flow = null;
	if (book || shownPreview(memory.view)) {
		await upload(jobId, book);
	} else {
		repaintImport(jobId);
	}
}

/** Removes picked files. The preview, when there is one, is made again without them. */
async function removePicked(jobId, entries) {
	const memory = memoryOf(jobId);
	const imp = importMemoryOf(jobId);
	if (!memory.view || importBusy(memory) || (memory.flow && memory.flow.kind === 'busy')) {
		return;
	}
	imp.picked = imp.picked.filter((entry) => !entries.includes(entry));
	imp.confirming = false;
	if (imp.flow && imp.flow.kind !== 'busy') {
		imp.flow = null;
	}
	if (shownPreview(memory.view)) {
		await upload(jobId, null);
	} else {
		repaintImport(jobId);
	}
}

/**
 * The workbook generated for this job among picked files, read: { entry, bytes, workbook }, or null when
 * there is none; any other .xlsx file is a file for the file cells. With asWorkbook, a lone .xlsx file is
 * taken as the workbook whatever it holds, and the upload says what is wrong with it.
 */
async function pickedWorkbook(view, picked, asWorkbook, say) {
	const books = picked.filter((entry) => /\.xlsx$/i.test(entry.file.name));
	if (books.length === 0) {
		return null;
	}
	const lone = asWorkbook && picked.length === 1;
	const config = await configOf(view);
	const workbookLogic = await loadWorkbook();
	for (const entry of books) {
		const megabytes = entry.file.size / (1024 * 1024);
		if (config.maxFileSizeMb > 0 && megabytes > config.maxFileSizeMb) {
			if (lone) {
				throw new Error('The workbook is ' + megabytes.toFixed(1) + ' MB, and an upload takes at most ' + config.maxFileSizeMb + ' MB.');
			}
			continue;
		}
		say('Reading ' + entry.file.name + '...');
		// The file comes from the Form's frame, so its bytes belong to that window. The workbook reader takes
		// bytes from any window.
		const bytes = await entry.file.arrayBuffer();
		let workbook;
		try {
			workbook = await workbookLogic.parseWorkbook(bytes);
		} catch (error) {
			if (lone) {
				throw error;
			}
			continue;
		}
		const generatedFor = workbook.meta && workbook.meta.job ? String(workbook.meta.job.id || '').toUpperCase() : '';
		if (lone || generatedFor === String(view.jobId).toUpperCase()) {
			return { entry, bytes, workbook };
		}
	}
	return null;
}

// --- The preview grid and the item preview ---

function clearPreviewPart(view) {
	disposePreviewUi(view);
	view.parts.preview.textContent = '';
}

function disposePreviewUi(view) {
	const ui = view.previewUi;
	view.previewUi = null;
	if (ui && ui.pane && typeof ui.pane.dispose === 'function') {
		try {
			ui.pane.dispose();
		} catch (ignored) {
			// The pane's frame may be gone already.
		}
	}
}

/** Called when another Form takes over the job: the item preview of the old one is let go. */
function leaveView(view) {
	disposePreviewUi(view);
}

/**
 * Draws the preview: the sheet tabs and filters, the grid, and beside it the row's
 * messages and the item preview pane. It is built once for a result and then only
 * updated, so that the pane is not loaded again each time the section is drawn.
 */
function drawPreviewPart(view, preview) {
	if (view.previewUi && view.previewUi.preview === preview) {
		return;
	}
	clearPreviewPart(view);
	const { doc } = view;
	const make = (tag, attributes, ...children) => element(doc, tag, attributes, ...children);
	const toolbar = make('div', { class: 'mce-toolbar' });
	const detail = make('div', { class: 'mce-row-detail', 'aria-live': 'polite' });
	const paneHost = make('div', { class: 'mce-pane-host' });
	const side = make('div', { class: 'mce-side' }, detail, paneHost);
	const divider = make('div', { class: 'mce-divider', role: 'separator', 'aria-orientation': 'vertical', 'aria-label': 'Resize the grid', tabindex: '0', title: 'Drag to resize' });
	const gridBox = make('div', { class: 'mce-grid-box' });
	const split = make('div', { class: 'mce-split' }, gridBox, divider, side);
	view.parts.preview.append(toolbar, split);

	const ui = { preview, toolbar, detail, paneHost, gridBox, split, pane: null, grid: null, shown: [] };
	view.previewUi = ui;
	ui.grid = createGrid(doc, { onSelect: (position) => selectPosition(view, position), onOpen: null });
	gridBox.append(ui.grid.element);
	if (preview.gridWidth) {
		gridBox.style.flexBasis = preview.gridWidth + 'px';
	}
	makeResizable(view, divider, gridBox, split, (width) => {
		preview.gridWidth = width;
	});

	paneHost.append(make('p', { class: 'mce-muted mce-pane-note' }, 'Loading the item preview...'));
	loadPreviewPane().then((module) => {
		if (view.previewUi !== ui) {
			return;
		}
		paneHost.textContent = '';
		if (!module || typeof module.createPreviewPane !== 'function') {
			paneHost.append(make('p', { class: 'mce-muted mce-pane-note' }, 'The item preview is not available.'));
			return;
		}
		if (!doc.getElementById('mce-preview-style')) {
			doc.head.append(element(doc, 'link', { id: 'mce-preview-style', rel: 'stylesheet', href: sibling('./mce-preview.css') }));
		}
		try {
			const job = readJob(view);
			// After the import, Aras holds what it wrote: the pane then shows the preview's rows only.
			ui.pane = module.createPreviewPane(paneHost, { top: view.top, job, result: preview.result, imported: job.state === 'Imported' });
		} catch (error) {
			paneHost.textContent = '';
			paneHost.append(make('p', { class: 'mce-error mce-pane-note' }, 'The item preview could not be shown: ' + error.message));
			return;
		}
		showSelected(view);
	});

	drawToolbar(view);
	fillGrid(view, true);
	showSelected(view);
}

function drawToolbar(view) {
	const ui = view.previewUi;
	const preview = ui.preview;
	const { doc } = view;
	const make = (tag, attributes, ...children) => element(doc, tag, attributes, ...children);
	ui.toolbar.textContent = '';
	const tabs = make('div', { class: 'mce-tabs', role: 'tablist', 'aria-label': 'Sheets' });
	// In the order of the workbook's tabs: each target's main sheet, then its relationship sheets.
	// With several targets, the tabs of a target share its colour, and a gap sets them apart.
	const targets = imports.sheetTargets(preview.result);
	const several = imports.sheetGroups(preview.result).length > 1;
	imports.sheetOrder(preview.result).forEach((index, place, order) => {
		const sheet = preview.result.sheets[index];
		const own = imports.actionCounts(sheet);
		const on = index === preview.sheetIndex;
		const startsTarget = several && place > 0 && targets[order[place - 1]] !== targets[index];
		const tab = make('button', {
			type: 'button', role: 'tab', id: 'mce-sheet-tab-' + index, 'aria-selected': on ? 'true' : 'false',
			class: 'mce-tab' + (on ? ' mce-tab-on' : '') + targetClass(targets[index], several) + (startsTarget ? ' mce-tab-next-target' : '')
		},
			sheet.sheetName,
			make('span', { class: 'mce-tab-count', title: plural(own.rows, 'row', 'rows') }, String(own.rows)),
			own.error > 0 ? make('span', { class: 'mce-tab-errors', title: plural(own.error, 'row has', 'rows have') + ' errors' }, String(own.error)) : null);
		tab.addEventListener('click', () => {
			if (preview.sheetIndex !== index) {
				preview.scrollTops[preview.sheetIndex] = ui.grid.element.scrollTop;
				preview.picks[preview.sheetIndex] = preview.selected;
				preview.sheetIndex = index;
				// Each sheet keeps its own selected row, so the preview beside the grid always shows a row of the sheet on view.
				preview.selected = preview.picks[index] || null;
				drawToolbar(view);
				fillGrid(view, true);
				showSelected(view);
			}
		});
		tabs.append(tab);
	});

	const filter = preview.filter;
	const change = () => {
		drawToolbar(view);
		fillGrid(view, false);
		const focused = doc.getElementById(focusId);
		if (focused) {
			focused.focus();
		}
	};
	let focusId = '';
	const hide = checkbox(doc, {
		id: 'mce-filter-unchanged', checked: filter.hideUnchanged,
		onChange: (on, input) => {
			filter.hideUnchanged = on;
			focusId = input.id;
			change();
		}
	}, 'Hide No change');
	const show = selectBox(doc, {
		id: 'mce-filter-show',
		onChange: (value, select) => {
			filter.show = value;
			focusId = select.id;
			change();
		}
	},
	make('option', { value: 'all', selected: filter.show === 'all' }, 'All rows'),
	make('option', { value: 'new', selected: filter.show === 'new' }, 'Create only'),
	make('option', { value: 'update', selected: filter.show === 'update' }, 'Update only'),
	make('option', { value: 'delete', selected: filter.show === 'delete' }, 'Delete only'));
	const problems = checkbox(doc, {
		id: 'mce-filter-problems', checked: filter.problemsOnly,
		onChange: (on, input) => {
			filter.problemsOnly = on;
			focusId = input.id;
			change();
		}
	}, 'Errors and warnings only');
	const sheet = preview.result.sheets[preview.sheetIndex];
	const shown = imports.filterRows(sheet, filter).length;
	ui.toolbar.append(tabs, make('div', { class: 'mce-filters' },
		make('label', { class: 'mce-filter', for: 'mce-filter-show' }, 'Show'), show.box,
		hide.label,
		problems.label,
		make('span', { class: 'mce-muted mce-shown', role: 'status' }, shown === sheet.rows.length ? plural(shown, 'row', 'rows') : shown + ' of ' + sheet.rows.length + ' rows')));
}

/** Fills the grid with the rows of the current sheet that pass the filter. */
function fillGrid(view, sheetChanged) {
	const ui = view.previewUi;
	const preview = ui.preview;
	const sheetIndex = preview.sheetIndex;
	const sheet = preview.result.sheets[sheetIndex];
	const valueColumns = preview.columns[sheetIndex];
	const doc = view.doc;
	ui.shown = imports.filterRows(sheet, preview.filter);
	ui.grid.setColumns([
		{ label: 'Row', width: 56, className: 'mce-cell-number' },
		{ label: 'Action', width: 92 },
		{ label: imports.keyHeader(sheet), width: 210 },
		...valueColumns.map((column) => ({ label: column.label, width: 170 })),
		{ label: 'Messages', width: 380 }
	]);
	ui.grid.setRows(ui.shown.length, (rowElement, position) => {
		const row = sheet.rows[ui.shown[position]];
		const cell = (text, className, title) => {
			const node = doc.createElement('div');
			node.className = 'mce-cell' + (className ? ' ' + className : '');
			node.textContent = text;
			if (title) {
				node.title = title;
			}
			return node;
		};
		rowElement.append(cell(String(row.row), 'mce-cell-number'));
		const action = cell('', 'mce-cell-action');
		action.append(element(doc, 'span', { class: 'mce-action mce-action-' + row.action }, imports.ACTION_LABELS[row.action] || row.action));
		rowElement.append(action);
		const name = imports.rowName(row);
		rowElement.append(cell(name || '(no key)', name ? '' : 'mce-cell-empty', name));
		for (const column of valueColumns) {
			const value = imports.cellOf(row, column);
			if (!value) {
				rowElement.append(cell('', ''));
			} else if (value.old === null) {
				rowElement.append(cell(value.text, '', value.text));
			} else {
				const node = cell('', 'mce-cell-changed', value.old + ' → ' + value.text);
				node.append(element(doc, 'span', { class: 'mce-old' }, value.old), ' → ', element(doc, 'span', { class: 'mce-new' }, value.text));
				rowElement.append(node);
			}
		}
		const messages = imports.rowMessages(row);
		const first = messages[0];
		rowElement.append(cell(messages.map((issue) => issue.message).join('  '),
			first ? (first.severity === 'error' ? 'mce-cell-error' : 'mce-cell-warning') : '', messages.map((issue) => issue.message).join('\n')));
	});
	if (sheetChanged) {
		ui.grid.element.scrollTop = preview.scrollTops[sheetIndex] || 0;
	}
	const selected = preview.selected && preview.selected.sheetIndex === sheetIndex ? ui.shown.indexOf(preview.selected.rowIndex) : -1;
	ui.grid.select(selected, false);
}

function selectPosition(view, position) {
	const ui = view.previewUi;
	if (!ui || position < 0 || position >= ui.shown.length) {
		return;
	}
	ui.preview.selected = { sheetIndex: ui.preview.sheetIndex, rowIndex: ui.shown[position] };
	ui.grid.select(position, true);
	showSelected(view);
}

/**
 * Shows the selected row beside the grid. The item preview pane shows the item,
 * or the messages of a row with errors, which has no preview. Without the
 * pane, the row's messages are shown here.
 */
function showSelected(view) {
	const ui = view.previewUi;
	const { doc } = view;
	const make = (tag, attributes, ...children) => element(doc, tag, attributes, ...children);
	const selected = ui.preview.selected;
	ui.detail.textContent = '';
	if (ui.pane) {
		ui.detail.hidden = true;
		if (selected) {
			ui.pane.show({ sheetIndex: selected.sheetIndex, rowIndex: selected.rowIndex });
		} else {
			ui.pane.clear();
		}
		return;
	}
	ui.detail.hidden = false;
	if (!selected) {
		ui.detail.append(make('p', { class: 'mce-muted' }, 'Choose a row to see its messages.'));
		return;
	}
	const sheet = ui.preview.result.sheets[selected.sheetIndex];
	const row = sheet.rows[selected.rowIndex];
	const name = imports.rowName(row);
	ui.detail.append(make('p', { class: 'mce-detail-title' },
		make('span', { class: 'mce-action mce-action-' + row.action }, imports.ACTION_LABELS[row.action] || row.action),
		' Row ' + row.row + ' of ' + sheet.sheetName + (name ? ': ' + name : '')));
	const messages = imports.rowMessages(row);
	if (messages.length > 0) {
		ui.detail.append(make('ul', { class: 'mce-detail-messages' },
			...messages.map((issue) => make('li', { class: issue.severity === 'error' ? 'mce-error-text' : 'mce-warning-text' }, issue.message))));
	}
	if (row.action === 'error') {
		ui.detail.append(make('p', { class: 'mce-muted' }, 'A row with errors is not imported. Correct it in the workbook and upload the workbook again.'));
	}
}

/** The divider between the grid and the side: dragging it, or the arrow keys, change the grid's width. */
function makeResizable(view, divider, gridBox, split, remember) {
	const setWidth = (width) => {
		const total = split.getBoundingClientRect().width;
		const clamped = Math.round(Math.max(280, Math.min(total - 280, width)));
		gridBox.style.flexBasis = clamped + 'px';
		remember(clamped);
	};
	divider.addEventListener('pointerdown', (event) => {
		event.preventDefault();
		divider.setPointerCapture(event.pointerId);
		split.classList.add('mce-resizing');
		const left = split.getBoundingClientRect().left;
		const move = (moved) => setWidth(moved.clientX - left);
		const stop = () => {
			split.classList.remove('mce-resizing');
			divider.removeEventListener('pointermove', move);
			divider.removeEventListener('pointerup', stop);
			divider.removeEventListener('pointercancel', stop);
			if (view.previewUi) {
				view.previewUi.grid.draw();
			}
		};
		divider.addEventListener('pointermove', move);
		divider.addEventListener('pointerup', stop);
		divider.addEventListener('pointercancel', stop);
	});
	divider.addEventListener('keydown', (event) => {
		if (event.key === 'ArrowLeft' || event.key === 'ArrowRight') {
			event.preventDefault();
			setWidth(gridBox.getBoundingClientRect().width + (event.key === 'ArrowLeft' ? -40 : 40));
		}
	});
}

/**
 * A grid that draws only the rows in view, so that it stays fast with many
 * thousands of rows. Every row has the same height.
 *
 * @param {Document} doc
 * @param {{onSelect: function(number), onOpen: (function(number)|null)}} options
 *   Called with the position of a row among the rows shown.
 */
function createGrid(doc, options) {
	const win = doc.defaultView;
	const scroller = doc.createElement('div');
	scroller.className = 'mce-grid';
	scroller.tabIndex = 0;
	scroller.setAttribute('role', 'grid');
	const head = doc.createElement('div');
	head.className = 'mce-grid-head';
	head.setAttribute('role', 'row');
	const body = doc.createElement('div');
	body.className = 'mce-grid-body';
	body.setAttribute('role', 'rowgroup');
	scroller.append(head, body);
	// drawn: the rows from first to last are in the page for the rows set now.
	// widths: the widths the user gave columns, by label, kept when the columns are set again.
	const state = { count: 0, renderRow: null, selected: -1, frame: 0, drawn: false, first: 0, last: 0, columns: [], widths: new Map() };

	/**
	 * Lays the columns out. The last column takes at least its width and also the room left
	 * over, so the grid always fills its box; wider than the box, the grid scrolls sideways.
	 */
	function layColumns() {
		const widths = state.columns.map((column) => state.widths.get(column.label) || column.width);
		const total = widths.reduce((sum, width) => sum + width, 0);
		scroller.style.setProperty('--mce-grid-columns', widths
			.map((width, index) => (index === widths.length - 1 ? 'minmax(' + width + 'px, 1fr)' : width + 'px')).join(' '));
		head.style.minWidth = total + 'px';
		body.style.minWidth = total + 'px';
	}

	/** Dragging the right edge of a column's header changes its width; a double-click gives the width back. */
	function resizer(column, cell) {
		const handle = doc.createElement('span');
		handle.className = 'mce-col-handle';
		handle.setAttribute('aria-hidden', 'true');
		handle.title = 'Drag to resize the column. Double-click to reset it.';
		handle.addEventListener('pointerdown', (event) => {
			event.preventDefault();
			event.stopPropagation();
			handle.setPointerCapture(event.pointerId);
			scroller.classList.add('mce-resizing');
			const start = event.clientX;
			const from = cell.getBoundingClientRect().width;
			const move = (moved) => {
				state.widths.set(column.label, Math.round(Math.max(40, from + moved.clientX - start)));
				layColumns();
			};
			const stop = () => {
				scroller.classList.remove('mce-resizing');
				handle.removeEventListener('pointermove', move);
				handle.removeEventListener('pointerup', stop);
				handle.removeEventListener('pointercancel', stop);
			};
			handle.addEventListener('pointermove', move);
			handle.addEventListener('pointerup', stop);
			handle.addEventListener('pointercancel', stop);
		});
		handle.addEventListener('dblclick', (event) => {
			event.stopPropagation();
			state.widths.delete(column.label);
			layColumns();
		});
		return handle;
	}

	/**
	 * Draws the rows in view. Rows that are drawn already stay as they are unless `again` is
	 * given: a row that is replaced between the two clicks of a double-click gets no dblclick.
	 */
	function draw(again = true) {
		state.frame = 0;
		const offset = head.offsetHeight;
		const top = Math.max(0, scroller.scrollTop - offset);
		const height = scroller.clientHeight || 600;
		const first = Math.max(0, Math.floor(top / ROW_HEIGHT) - 6);
		const last = Math.min(state.count, Math.ceil((top + height) / ROW_HEIGHT) + 6);
		if (!again && state.drawn && first === state.first && last === state.last) {
			return;
		}
		state.drawn = true;
		state.first = first;
		state.last = last;
		body.textContent = '';
		for (let position = first; position < last; position++) {
			const row = doc.createElement('div');
			row.className = 'mce-grid-row' + (position === state.selected ? ' mce-selected' : '');
			row.setAttribute('role', 'row');
			row.setAttribute('aria-rowindex', String(position + 2));
			row.setAttribute('aria-selected', position === state.selected ? 'true' : 'false');
			row.dataset.position = String(position);
			row.style.top = position * ROW_HEIGHT + 'px';
			state.renderRow(row, position);
			body.append(row);
		}
		if (state.count === 0) {
			const empty = doc.createElement('p');
			empty.className = 'mce-muted mce-grid-empty';
			empty.textContent = 'No rows to show.';
			body.append(empty);
		}
	}

	function reveal(position) {
		const offset = head.offsetHeight;
		const rowTop = position * ROW_HEIGHT;
		const visible = scroller.clientHeight - offset;
		if (rowTop < scroller.scrollTop) {
			scroller.scrollTop = rowTop;
		} else if (rowTop + ROW_HEIGHT > scroller.scrollTop + visible) {
			scroller.scrollTop = rowTop + ROW_HEIGHT - visible;
		}
	}

	scroller.addEventListener('scroll', () => {
		if (!state.frame) {
			state.frame = win.requestAnimationFrame(() => draw(false));
		}
	});
	body.addEventListener('click', (event) => {
		const row = event.target.closest('.mce-grid-row');
		if (row) {
			options.onSelect(Number(row.dataset.position));
		}
	});
	body.addEventListener('dblclick', (event) => {
		const row = event.target.closest('.mce-grid-row');
		if (row && options.onOpen) {
			options.onOpen(Number(row.dataset.position));
		}
	});
	scroller.addEventListener('keydown', (event) => {
		if (state.count === 0) {
			return;
		}
		if (event.key === 'Enter' && options.onOpen && state.selected >= 0) {
			options.onOpen(state.selected);
			return;
		}
		const page = Math.max(1, Math.floor((scroller.clientHeight - head.offsetHeight) / ROW_HEIGHT) - 1);
		const moves = { ArrowDown: 1, ArrowUp: -1, PageDown: page, PageUp: -page, Home: -Infinity, End: Infinity };
		if (!Object.prototype.hasOwnProperty.call(moves, event.key)) {
			return;
		}
		event.preventDefault();
		const to = Math.max(0, Math.min(state.count - 1, Math.max(state.selected, -1) + moves[event.key]));
		options.onSelect(to);
	});

	return {
		element: scroller,
		setColumns(columns) {
			head.textContent = '';
			state.columns = columns;
			layColumns();
			for (const column of columns) {
				const cell = doc.createElement('div');
				cell.className = 'aras-grid-head-cell mce-cell mce-head-cell' + (column.className ? ' ' + column.className : '');
				cell.setAttribute('role', 'columnheader');
				cell.textContent = column.label;
				cell.title = column.label;
				cell.append(resizer(column, cell));
				head.append(cell);
			}
		},
		setRows(count, renderRow) {
			state.count = count;
			state.renderRow = renderRow;
			state.selected = -1;
			state.drawn = false;
			body.style.height = Math.max(count, 2) * ROW_HEIGHT + 'px';
		},
		/** Selects a row. Rows in the page only change their mark, so that a double-click still reaches them. */
		select(position, show) {
			state.selected = position;
			if (show && position >= 0) {
				reveal(position);
			}
			if (!state.drawn) {
				draw();
				return;
			}
			for (const row of Array.from(body.children)) {
				if (row.dataset.position === undefined) {
					continue;
				}
				const on = Number(row.dataset.position) === position;
				row.classList.toggle('mce-selected', on);
				row.setAttribute('aria-selected', on ? 'true' : 'false');
			}
			// A row brought into view by reveal() is drawn on the scroll that follows.
		},
		draw
	};
}

// --- Uploading ---

/** Saves a Draft job that is being edited and finishes editing it, with the item window's Done command. */
async function finishEditing(view, say, notDone) {
	const aras = view.top.aras;
	const node = view.doc.item;
	const isNew = !!aras.isTempEx(node);
	if (!isNew && !aras.isLockedByUser(node)) {
		return false;
	}
	// Only a Draft job is saved by hand. After that the server Methods change the job, and a
	// save would be refused, so a job that is still locked is not saved here.
	if (!isNew && readJob(view).state !== 'Draft') {
		throw new Error('The job is locked for editing, and after Draft a job is not changed by hand. Close the edit with Discard, then try again; ' + notDone + '.');
	}
	say('Saving the job...');
	const saved = await view.itemWindow.onDoneCommand();
	if (!saved) {
		throw new Error('The job was not saved, so ' + notDone + '.');
	}
	return true;
}

/**
 * Uploading again deletes the job's earlier upload: its uploaded workbook and its validation
 * result. An Aras dialog asks first and names what is deleted. Resolves true when the
 * upload goes ahead. Asked once the file is chosen: a dialog before the file picker would
 * keep the browser from opening the picker.
 */
async function confirmReplacingUpload(view, memory, fileName) {
	const job = readJob(view);
	if (job.state !== 'Validated' || (!job.uploadFile && !job.resultsFile)) {
		return true;
	}
	if (memory.askingToUpload) {
		return false;
	}
	memory.askingToUpload = true;
	try {
		const listed = job.uploadFile && job.resultsFile ? 'the uploaded workbook and its validation result are' :
			job.uploadFile ? 'the uploaded workbook is' : 'the validation result is';
		const aras = view.top.aras;
		const modules = (view.itemWindow && view.itemWindow.ArasModules) || view.top.ArasModules;
		const answer = await modules.Dialog.confirm(
			'Uploading ' + fileName + ' replaces the earlier upload of job ' + job.number + '. When the new workbook is accepted, ' +
			listed + ' deleted from Aras, and the preview shows the new workbook. ' +
			'Copies you downloaded are not affected. Are you sure?',
			{ title: 'Replace the uploaded workbook?', okButtonText: 'Delete and upload', cancelButtonText: aras.getResource('', 'common.cancel') }
		);
		return answer === 'ok';
	} finally {
		memory.askingToUpload = false;
	}
}

/**
 * Checks a workbook with the picked files and stores the preview. book is the workbook just picked,
 * as pickedWorkbook() reads it: it is uploaded in place of the earlier one. Without it, the job's
 * uploaded workbook is checked again, after the picked files changed; it stays the uploaded workbook.
 */
async function upload(jobId, book) {
	const memory = memoryOf(jobId);
	const imp = importMemoryOf(jobId);
	if (importBusy(memory) || (memory.flow && memory.flow.kind === 'busy')) {
		return;
	}
	const view = memory.view;
	const top = view.top;
	const say = (message) => {
		imp.flow = { kind: 'busy', step: 'upload', message };
		repaintImport(jobId);
	};
	imp.confirming = false;
	let changed = false;
	try {
		const config = await configOf(view);
		let workbook;
		let workbookSize;
		let uploadFileId = null;
		if (book) {
			workbookSize = book.entry.file.size;
			const megabytes = workbookSize / (1024 * 1024);
			if (config.maxFileSizeMb > 0 && megabytes > config.maxFileSizeMb) {
				throw new Error('The workbook is ' + megabytes.toFixed(1) + ' MB, and an upload takes at most ' + config.maxFileSizeMb + ' MB.');
			}
			if (!(await confirmReplacingUpload(view, memory, book.entry.file.name))) {
				repaint(jobId);
				return;
			}
			changed = await finishEditing(view, say, 'the workbook was not uploaded');
			workbook = book.workbook;
		} else {
			const preview = shownPreview(view);
			if (!preview) {
				throw new Error('The job has no preview to check again. Upload the workbook.');
			}
			changed = await finishEditing(view, say, 'the workbook was not checked again');
			const job = readJob(view);
			workbook = await uploadedWorkbook(top, job, preview, say);
			workbookSize = preview.workbookSize;
			uploadFileId = job.uploadFile;
		}

		// The picked files: their MD5, for the check and for the import, which stores them.
		const picked = imp.picked.slice();
		if (config.maxAttachments > 0 && picked.length > config.maxAttachments) {
			throw new Error(plural(picked.length, 'file is', 'files are') + ' picked, and an upload takes at most ' + config.maxAttachments + '. Remove some under Files.');
		}
		const pickedMb = picked.reduce((sum, entry) => sum + entry.file.size, 0) / (1024 * 1024);
		if (config.maxAttachmentsTotalMb > 0 && pickedMb > config.maxAttachmentsTotalMb) {
			throw new Error('The picked files are ' + pickedMb.toFixed(1) + ' MB together, and an upload takes at most ' + config.maxAttachmentsTotalMb + ' MB. Remove some under Files.');
		}
		const unread = picked.filter((entry) => !entry.md5);
		if (unread.length > 0) {
			const md5 = await loadMd5();
			for (let index = 0; index < unread.length; index++) {
				const entry = unread[index];
				say('Reading the files: ' + (index + 1) + ' of ' + unread.length + ' (' + entry.file.name + ')...');
				entry.md5 = await md5.md5OfBlob(entry.file);
			}
		}
		// The same file picked twice, under two paths, is listed once: a cell that names it is not ambiguous.
		const manifest = [];
		const listed = new Set();
		for (const entry of picked) {
			const fileToken = imports.fileToken(entry.path, entry.md5).toLowerCase();
			if (!listed.has(fileToken)) {
				listed.add(fileToken);
				manifest.push({ path: entry.path, size: entry.file.size, md5: entry.md5 });
			}
		}
		const rowCount = workbook.sheets.reduce((sum, sheet) => sum + sheet.rows.length, 0);

		say('Checking ' + plural(rowCount, 'row', 'rows') + (book ? '' : ' again with the picked files') + '...');
		const result = await callMethod(top, VALIDATE_METHOD, {
			job_id: jobId, workbook: JSON.stringify(workbook), file_size: workbookSize ? String(workbookSize) : '',
			files: manifest.length > 0 ? JSON.stringify(manifest) : ''
		});
		if (result.status !== 'validated') {
			imp.flow = {
				kind: 'rejected', step: 'upload', at: Date.now(), issues: result.issues || [],
				message: result.status === 'failed' ? 'The workbook could not be checked. Nothing was changed.' : 'The workbook was not accepted, and nothing was changed:'
			};
		} else {
			imports.addKeyTexts(result, workbook);
			imports.addIndentedResult(result, workbook);
			// The list of the picked files goes with the preview: each batch is checked with it again.
			result.files = manifest;
			if (book) {
				say('Storing the workbook and the preview...');
				uploadFileId = await storeFile(top, book.entry.file.name, book.bytes);
			} else {
				say('Storing the preview...');
			}
			const resultsFileId = await storeFile(top, imports.resultsFileName(result.job && result.job.number),
				new TextEncoder().encode(JSON.stringify(result)), JSON_TYPE);
			await importStep.upload(top, jobId, workbook.meta && workbook.meta.job ? workbook.meta.job.token : '', uploadFileId, resultsFileId);
			changed = true;
			imp.preview = newPreview(resultsFileId, uploadFileId, result, workbook, workbookSize);
			imp.pickedOpen = null;
			imp.flow = null;
		}
	} catch (error) {
		imp.flow = { kind: 'error', step: 'upload', at: Date.now(), message: error && error.message ? error.message : String(error) };
	}
	await refreshJob(top, jobId, changed);
}

/** Shows the job as it is now on the server: the item window is refreshed when the job changed. */
async function refreshJob(top, jobId, changed) {
	const itemWindow = changed ? top.aras.uiFindWindowEx(jobId) : null;
	if (itemWindow && typeof itemWindow.onRefresh === 'function') {
		try {
			await itemWindow.onRefresh();
			return;
		} catch (error) {
			// The section is drawn again below, as it is.
		}
	}
	repaint(jobId);
}

// --- Importing ---

/** The uploaded workbook as parseWorkbook gives it: from memory when it is the job's, else read again from the vault. */
async function uploadedWorkbook(top, job, preview, say) {
	if (preview.workbook && preview.uploadFileId === job.uploadFile) {
		return preview.workbook;
	}
	if (!job.uploadFile) {
		throw new Error('The job has no uploaded workbook. Upload it again.');
	}
	say('Reading the uploaded workbook again...');
	const buffer = await readStoredFile(top, job.uploadFile);
	const workbook = await (await loadWorkbook()).parseWorkbook(buffer);
	preview.workbook = workbook;
	preview.workbookSize = buffer.byteLength;
	preview.uploadFileId = job.uploadFile;
	return workbook;
}

async function runImport(jobId) {
	const memory = memoryOf(jobId);
	const imp = importMemoryOf(jobId);
	if (importBusy(memory) || (memory.flow && memory.flow.kind === 'busy') || !imp.preview) {
		return;
	}
	const view = memory.view;
	const top = view.top;
	const preview = imp.preview;
	const result = preview.result;
	const progress = { batch: 0, batches: 0, written: 0, failed: 0, toWrite: 0 };
	const say = (message) => {
		imp.flow = { kind: 'busy', step: 'import', message, progress: { ...progress } };
		repaintImport(jobId);
	};
	imp.confirming = false;
	const failures = [];
	let begun = false;
	let changed = false;
	try {
		const job = readJob(view);
		changed = await finishEditing(view, say, 'nothing was imported');
		const workbook = await uploadedWorkbook(top, job, preview, say);
		const mismatch = imports.workbookMatchesResult(workbook, result);
		if (mismatch) {
			throw new Error(mismatch + ' Upload the workbook again.');
		}
		const token = workbook.meta && workbook.meta.job ? workbook.meta.job.token : '';

		say('Starting the import...');
		const begin = await importStep.begin(top, jobId, token);
		begun = true;
		// The Files stored for the import so far, by file token: a run that was broken off stored some.
		const stored = new Set();
		const keepStored = (list) => {
			for (const entry of list || []) {
				stored.add('file:' + String(entry.md5).toUpperCase() + ':' + entry.name);
			}
		};
		keepStored(begin.files);
		const manifest = Array.isArray(result.files) ? result.files : [];
		const pickedFiles = pickedByToken(imp);
		const plan = imports.planImport(result);
		const config = await configOf(view);
		const batches = imports.makeBatches(plan.units, Number(begin.batchSize) || config.batchSize);
		progress.toWrite = plan.units.reduce((sum, unit) => sum + imports.unitRowCount(unit), 0);
		progress.batches = batches.length;
		// The main-sheet rows not imported: a later batch leaves out the rows that point at the item one was to create.
		const failedMainRows = new Set();
		// Every row not imported: the rows below a line of an indented sheet that failed are left out too.
		const failedRows = new Set();
		const fail = (entries) => {
			for (const entry of entries) {
				failures.push(imports.failedRowEntry(result, entry));
				const failedRow = imports.rowKey(entry.sheetIndex, result.sheets[entry.sheetIndex].rows[entry.rowIndex].row);
				failedRows.add(failedRow);
				if (imports.failedRowLeavesNoItem(result, entry)) {
					failedMainRows.add(failedRow);
				}
			}
			progress.failed += entries.length;
		};

		// Batches go at once, up to the configured number; a batch waits for the earlier batches that
		// create the items its rows point at, or touch the same existing item or file (batchDependencies).
		const waitsFor = imports.batchDependencies(batches, result);
		const atOnce = Math.max(1, Number(begin.parallelBatches) || config.parallelBatches || 1);
		progress.running = [];
		// The Files being stored, by file token: a batch that links one another batch is storing waits for it.
		const storing = new Map();
		const storeFilesOf = async (batch) => {
			const toStore = imports.batchFileTokens(batch, result).filter((fileToken) => !stored.has(fileToken));
			const waits = toStore.filter((fileToken) => storing.has(fileToken)).map((fileToken) => storing.get(fileToken));
			const own = toStore.filter((fileToken) => !storing.has(fileToken));
			if (own.length > 0) {
				// Stored as the user, then added to the job.
				const work = (async () => {
					const storedIds = [];
					try {
						for (let position = 0; position < own.length; position++) {
							const fileToken = own[position];
							const file = pickedFiles.get(fileToken);
							if (!file) {
								throw new Error('The file ' + imports.tokenName(fileToken) + ' is not picked. Add it under Files, then import.');
							}
							say(imports.progressText(progress) + ' Storing ' + imports.tokenName(fileToken) + ' (' + (position + 1) + ' of ' + own.length + ')...');
							storedIds.push(await storeFile(top, imports.tokenName(fileToken), file, file.type || 'application/octet-stream'));
						}
					} finally {
						if (storedIds.length > 0) {
							keepStored((await withRetry(() => importStep.files(top, jobId, token, storedIds))).files);
						}
					}
				})();
				for (const fileToken of own) {
					storing.set(fileToken, work);
				}
				waits.push(work);
			}
			await Promise.all(waits);
		};
		const sendBatch = async (index) => {
			const batch = batches[index];
			fail(imports.dropRowsPointingAt(batch, result, failedMainRows));
			fail(imports.dropRowsBelowFailed(batch, result, workbook, failedRows));
			for (let position = batch.length - 1; position >= 0; position--) {
				if (!batch[position].writesMain && batch[position].rels.length === 0) {
					batch.splice(position, 1);
				}
			}
			progress.batch = index + 1;
			progress.running = progress.running.concat(index + 1);
			try {
				await storeFilesOf(batch);
				while (batch.length > 0) {
					say(imports.progressText(progress));
					let answer;
					try {
						answer = await withRetry(() => importStep.batch(top, jobId, imports.sliceWorkbook(workbook, result, batch), manifest));
					} catch (error) {
						// A write failed and the whole request was rolled back. The server names the rows
						// of that write; they are marked failed and the batch goes again without them.
						const failure = imports.writeFailureOf(error.fault, result);
						const dropped = failure ? imports.dropFailedWrite(batch, result, failure) : null;
						if (!dropped) {
							throw error;
						}
						fail(dropped);
						continue;
					}
					const read = imports.readBatchAnswer(answer, result);
					fail(read.failed);
					for (const other of read.unknown) {
						failures.push({ sheetName: String(other.sheet), row: other.row, key: '', message: other.message });
						progress.failed++;
					}
					progress.written += read.created + read.updated + read.deleted;
					break;
				}
			} finally {
				progress.running = progress.running.filter((number) => number !== index + 1);
			}
		};
		await imports.runScheduled(batches.length, waitsFor, atOnce, sendBatch);

		say('Finishing the import...');
		const done = await importStep.complete(top, jobId, token, failures, plan.counts);
		changed = true;
		const counts = (done && done.counts) || {};
		imp.flow = {
			kind: 'done', step: 'import', at: Date.now(),
			message: 'The import is finished: ' + (Number(counts.created) || 0) + ' created, ' + (Number(counts.updated) || 0) + ' updated, ' +
				(Number(counts.deleted) > 0 ? Number(counts.deleted) + ' deleted, ' : '') +
				(counts.failed !== undefined ? Number(counts.failed) : failures.length) + ' failed.'
		};
		imp.preview = null;
		imp.picked = [];
		imp.pickedOpen = null;
	} catch (error) {
		const message = error && error.message ? error.message : String(error);
		imp.flow = {
			kind: 'error', step: 'import', at: Date.now(),
			message: !begun ? message : 'The import stopped: ' + message + ' ' + (progress.written > 0
				? 'The ' + plural(progress.written, 'row', 'rows') + ' written so far stay written, and the job stays Validated. Upload the workbook again to see what is left to import.'
				: 'The job stays Validated.')
		};
	}
	await refreshJob(top, jobId, changed);
}

// --- An imported job ---

function drawImported(view, job) {
	const make = (tag, attributes, ...children) => element(view.doc, tag, attributes, ...children);
	const imp = importMemoryOf(view.jobId);
	const top = view.parts.head;
	// The message of a finished import is shown with the others, under the header (drawImport).
	const count = (name) => Number(job.counts[name]) || 0;
	const parts = [count('created') + ' created', count('updated') + ' updated', count('failed') + ' failed'];
	if (count('skipped') > 0) {
		parts.push(count('skipped') + ' skipped because of errors');
	}
	if (count('unchanged') > 0) {
		parts.push(count('unchanged') + ' unchanged');
	}
	top.append(make('p', { class: 'mce-summary' }, 'Of ' + plural(count('total'), 'row', 'rows') + ': ' + parts.join(', ') + '.'));
	const records = view.importModel.records || [];
	// With several targets, what the import did for each: its main sheet and its relationship sheets.
	const targets = view.importModel.targets || [];
	const several = targets.length > 1;
	if (several && records.length > 0) {
		top.append(make('ul', { class: 'mce-target-counts' },
			...imports.recordCountsByTarget(records, targets).map((counts, at) =>
				make('li', null, targetMark(view.doc, at, several), imports.targetCountsLine(counts) + '.'))));
	}
	// The target of each sheet, by name: its rows are marked with the target's colour.
	const targetOfSheet = new Map();
	targets.forEach((target, at) => target.sheets.forEach((name) => targetOfSheet.set(name, at)));
	// The job keeps its validation result: the Result tab can also show the preview as it was
	// before the import, read-only.
	const preview = importedPreview(view, job);
	if (preview) {
		const viewTab = (value, label, title) => {
			const on = imp.resultView === value;
			const button = make('button', {
				type: 'button', role: 'tab', id: 'mce-result-view-' + value, class: 'mce-tab' + (on ? ' mce-tab-on' : ''),
				'aria-selected': on ? 'true' : 'false', tabindex: on ? '0' : '-1', title
			}, label);
			button.addEventListener('click', () => {
				if (imp.resultView !== value) {
					imp.resultView = value;
					drawImport(view);
				}
			});
			return button;
		};
		top.append(make('div', { class: 'mce-tabs mce-result-views', role: 'tablist', 'aria-label': 'Result' },
			viewTab('records', 'Import result', 'What the import did with each row.'),
			viewTab('preview', 'Preview before the import', 'The preview of the upload, as it was when you confirmed the import.')));
		if (imp.resultView === 'preview') {
			top.append(make('p', { class: 'mce-muted' },
				'The preview as it was before the import, read-only. It shows what each row was to do; the import result says what it did, ' +
				'such as a row that failed.'));
			drawPreviewPart(view, preview);
			return;
		}
	} else {
		imp.resultView = 'records';
	}
	if (view.previewUi && view.previewUi.records === records) {
		return;
	}
	clearPreviewPart(view);
	if (records.length === 0) {
		top.append(make('p', { class: 'mce-muted' }, 'The import has no rows to list.'));
		return;
	}
	const part = view.parts.preview;
	const filter = RECORD_FILTERS.find((choice) => choice.value === imp.resultShow) || RECORD_FILTERS[0];
	const showSelect = make('select', { id: 'mce-records-show' },
		...RECORD_FILTERS.map((choice) => {
			const count = choice.action ? records.filter((record) => record.action === choice.action).length : records.length;
			return make('option', { value: choice.value, selected: choice === filter }, choice.label + ' (' + count + ')');
		}));
	// The sheets, in the order of the records (the order of the workbook), when there is more than one.
	const sheetNames = Array.from(new Set(records.map((record) => record.sheet)));
	if (!sheetNames.includes(imp.resultSheet)) {
		imp.resultSheet = '';
	}
	const sheetSelect = sheetNames.length > 1
		? make('select', { id: 'mce-records-sheet' },
			make('option', { value: '', selected: imp.resultSheet === '' }, 'All sheets'),
			...sheetNames.map((name) => make('option', { value: name, selected: imp.resultSheet === name },
				name + ' (' + records.filter((record) => record.sheet === name).length + ')')))
		: null;
	const shownCount = make('span', { class: 'mce-muted mce-shown', role: 'status' });
	part.append(make('div', { class: 'mce-toolbar' }, make('div', { class: 'mce-filters' },
		sheetSelect ? make('label', { class: 'mce-filter', for: 'mce-records-sheet' }, 'Sheet') : null, sheetSelect ? arasSelect(view.doc, sheetSelect) : null,
		make('label', { class: 'mce-filter', for: 'mce-records-show' }, 'Show'), arasSelect(view.doc, showSelect), shownCount)));
	let shown = [];
	const grid = createGrid(view.doc, {
		onSelect: (position) => grid.select(position, true),
		onOpen: (position) => {
			const record = shown[position];
			if (record && record.targetId && record.targetType && (record.action === RECORD_ACTIONS.Created || record.action === RECORD_ACTIONS.Updated || record.action === RECORD_ACTIONS.Deleted)) {
				view.top.aras.uiShowItem(record.targetType, record.targetId);
			}
		}
	});
	view.previewUi = { records, grid, pane: null };
	grid.setColumns([
		{ label: 'Sheet', width: 150 }, { label: 'Row', width: 56, className: 'mce-cell-number' }, { label: 'Result', width: 92 },
		{ label: 'Item', width: 220 }, { label: 'Message', width: 520 }
	]);
	const fill = () => {
		const action = (RECORD_FILTERS.find((choice) => choice.value === imp.resultShow) || RECORD_FILTERS[0]).action;
		shown = records.filter((record) => (!action || record.action === action) && (imp.resultSheet === '' || record.sheet === imp.resultSheet));
		shownCount.textContent = shown.length === records.length ? plural(records.length, 'row', 'rows') : shown.length + ' of ' + records.length + ' rows';
		grid.setRows(shown.length, (rowElement, position) => {
			const record = shown[position];
			const cell = (text, className) => {
				const node = view.doc.createElement('div');
				node.className = 'mce-cell' + (className ? ' ' + className : '');
				node.textContent = text;
				node.title = text;
				return node;
			};
			const kind = record.action === RECORD_ACTIONS.Created ? 'new' : record.action === RECORD_ACTIONS.Updated ? 'update'
				: record.action === RECORD_ACTIONS.Deleted ? 'delete' : record.action === RECORD_ACTIONS.Skipped ? 'skipped' : 'error';
			const action = cell('', 'mce-cell-action');
			action.append(element(view.doc, 'span', { class: 'mce-action mce-action-' + kind }, record.action));
			const sheetCell = cell(record.sheet);
			const mark = targetMark(view.doc, targetOfSheet.get(record.sheet), several);
			if (mark) {
				sheetCell.prepend(mark);
			}
			rowElement.append(sheetCell, cell(String(record.row), 'mce-cell-number'), action, cell(record.name),
				cell(record.message, record.action === RECORD_ACTIONS.Failed || record.action === RECORD_ACTIONS.Skipped ? 'mce-cell-error' : ''));
		});
		grid.select(-1, false);
	};
	showSelect.addEventListener('change', () => {
		imp.resultShow = showSelect.value;
		fill();
	});
	if (sheetSelect) {
		sheetSelect.addEventListener('change', () => {
			imp.resultSheet = sheetSelect.value;
			fill();
		});
	}
	part.append(make('div', { class: 'mce-records' }, grid.element),
		make('p', { class: 'mce-muted' }, 'Double-click a created or updated row to open its item.'));
	fill();
}

// ---------------------------------------------------------------------------
// Closing a new job
// ---------------------------------------------------------------------------

/**
 * In the job's window, closing a new job that was never saved asks whether to
 * save it, as closing an edited job with unsaved changes does, so that its
 * choices (the target ItemType, the items selected for it) are not lost.
 */
function askBeforeClosingNewJob(itemWindow) {
	if (itemWindow.mceAsksBeforeClosing || typeof itemWindow.confirmExitEditMode !== 'function') {
		return;
	}
	const original = itemWindow.confirmExitEditMode;
	itemWindow.mceAsksBeforeClosing = true;
	itemWindow.confirmExitEditMode = async function confirmExitEditMode(...args) {
		const aras = itemWindow.aras;
		const item = itemWindow.item;
		if (!item || !aras || !aras.isTempEx(item) || itemWindow.ignorePageCloseHooks || !itemWindow.ArasModules) {
			return original.apply(this, args);
		}
		const answer = await itemWindow.ArasModules.Dialog.confirm(aras.getResource('', 'item_methods_ex.changes_not_saved'), {
			title: aras.getResource('', 'item_methods_ex.unsaved_changes'),
			okButtonText: aras.getResource('', 'common.save'),
			additionalButtons: [{ text: aras.getResource('', 'common.discard'), actionName: 'discard' }],
			buttonsOrdering: ['ok', 'discard', 'cancel']
		});
		if (!answer || answer === 'cancel') {
			return false;
		}
		if (answer === 'ok') {
			return !!(await itemWindow.onDoneCommand());
		}
		return true;
	};
}
