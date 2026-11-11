/**
 * Mass Create & Edit - the export step, without the page.
 *
 * Everything the job Form needs to work out before it can build a workbook,
 * kept free of the browser and of the Aras client so that it can be tested on
 * its own:
 *
 *   - the export options of a job (which relationships become sheets, the match
 *     key of each, and the columns left out of each sheet), as they are stored
 *     in the job's export_options. A job may have several targets, each with its
 *     own main sheet: `omit` is that of the first, and `targets` lists the others;
 *   - which columns can be left out, and the schema without them;
 *   - what the sheet picker offers for an ItemType, from the template schema;
 *   - the AML requests that read the items to pre-fill, and the conversion of
 *     their answers into the rows buildWorkbook takes.
 *
 * Items come in as plain objects, { id, props }, where props maps a property
 * name to { value, keyedName, item }: the text of the property, the keyed
 * name the server sent with an item property, and the nested item when the
 * request asked for one. A property that is empty in Aras has value ''.
 */

const OPTIONS_VERSION = 1;

/** How many ids go into one request. A request with 250 ids is about 9 KB. */
export const IDS_PER_REQUEST = 250;

/** The most rows a worksheet has below its header. */
export const SHEET_ROW_LIMIT = 1048575;

const escapeXml = (text) =>
	String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');

const isId = (text) => /^[0-9A-Fa-f]{32}$/.test(text);

/** The values of an AML "in" condition: each one quoted, with any quote inside doubled. */
const inList = (values) => values.map((value) => "'" + String(value).replace(/'/g, "''") + "'").join(',');

// ---------------------------------------------------------------------------
// Export options
// ---------------------------------------------------------------------------

/** A list of column ids as stored: text entries only, each once. */
const idList = (value) => (Array.isArray(value) ? Array.from(new Set(value.filter((id) => typeof id === 'string' && id !== ''))) : []);

/** Puts the columns left out on an object of the options, only when there are any. */
const withOmit = (target, omit) => {
	if (omit && omit.length > 0) {
		target.omit = omit.slice();
	}
	return target;
};

/** The layout of a relationship sheet written as an indented structure, with a Level column in place of the parent. */
export const INDENTED_LAYOUT = 'indented';

/** Puts the layout on a sheet of the options, only when it is the indented one. */
const withLayout = (target, layout) => {
	if (layout === INDENTED_LAYOUT) {
		target.layout = INDENTED_LAYOUT;
	}
	return target;
};

/**
 * Reads the export options of a job.
 *
 * @param {string|null} text The job's export_options.
 * @returns {{sheets: Array<{relationshipTypeId: string, name: string, matchKey: string[], omit?: string[]}>, omit?: string[],
 *   targets?: Array<{itemTypeId: string, omit?: string[]}>}}
 *   Text that is empty or can't be read gives options without sheets. omit lists the
 *   columns left out of the workbook, by column id: on the options themselves for the
 *   main sheet of the first target, on an entry of targets for the main sheet of another
 *   target, on a sheet for that relationship sheet. It is there only when a column
 *   is left out; without it the sheet has every column. targets lists the targets after
 *   the first, in their order, and is there only when the job has more than one.
 */
export function parseOptions(text) {
	const empty = { sheets: [] };
	if (typeof text !== 'string' || text.trim() === '') {
		return empty;
	}
	let stored;
	try {
		stored = JSON.parse(text);
	} catch (error) {
		return empty;
	}
	if (!stored || !Array.isArray(stored.sheets)) {
		return empty;
	}
	const sheets = [];
	for (const sheet of stored.sheets) {
		if (!sheet || !isId(String(sheet.relationshipTypeId || '')) || sheets.some((known) => known.relationshipTypeId === sheet.relationshipTypeId)) {
			continue;
		}
		sheets.push(withLayout(withOmit({
			relationshipTypeId: sheet.relationshipTypeId,
			name: typeof sheet.name === 'string' ? sheet.name : '',
			matchKey: Array.isArray(sheet.matchKey) ? sheet.matchKey.filter((id) => typeof id === 'string' && id !== '') : []
		}, idList(sheet.omit)), sheet.layout));
	}
	const options = withOmit({ sheets }, idList(stored.omit));
	const targets = [];
	for (const target of Array.isArray(stored.targets) ? stored.targets : []) {
		if (target && isId(String(target.itemTypeId || '')) && !targets.some((known) => known.itemTypeId === target.itemTypeId)) {
			targets.push(withOmit({ itemTypeId: target.itemTypeId }, idList(target.omit)));
		}
	}
	if (targets.length > 0) {
		options.targets = targets;
	}
	return options;
}

/**
 * The export options as the text stored on the job. Options without sheets are
 * still written out: the server tells "no relationship sheets" (an empty list)
 * from "never chosen on the Form" (no text at all).
 */
export function serializeOptions(options) {
	const sheets = (options && options.sheets) || [];
	const targets = (options && options.targets) || [];
	const stored = withOmit({
		version: targets.length > 0 ? 2 : OPTIONS_VERSION,
		sheets: sheets.map((sheet) => withLayout(withOmit({ relationshipTypeId: sheet.relationshipTypeId, name: sheet.name, matchKey: sheet.matchKey }, sheet.omit), sheet.layout))
	}, options && options.omit);
	if (targets.length > 0) {
		stored.targets = targets.map((target) => withOmit({ itemTypeId: target.itemTypeId }, target.omit));
	}
	return JSON.stringify(stored);
}

// ---------------------------------------------------------------------------
// The targets of a job
// ---------------------------------------------------------------------------

/**
 * The targets of a job: its target ItemTypes in the order of their main sheets, each with
 * the items to pre-fill. A job with several holds them in its targets property, as JSON;
 * a job with one has its target_itemtype, with the items of prefill_idlist. A list that
 * can't be read, or whose first entry is not target_itemtype, counts as not there.
 *
 * @param {string} text The job's targets.
 * @param {string} targetId The job's target_itemtype.
 * @param {string} prefillText The job's prefill_idlist.
 * @returns {Array<{itemTypeId: string, prefillIds: string[]}>} Empty for a job without a target.
 */
export function parseTargets(text, targetId, prefillText) {
	const first = String(targetId || '').toUpperCase();
	const own = isId(first) ? [{ itemTypeId: first, prefillIds: parseIdList(prefillText) }] : [];
	if (typeof text !== 'string' || text.trim() === '') {
		return own;
	}
	let stored;
	try {
		stored = JSON.parse(text);
	} catch (error) {
		return own;
	}
	if (!Array.isArray(stored)) {
		return own;
	}
	const targets = [];
	for (const entry of stored) {
		const id = entry && typeof entry.itemTypeId === 'string' ? entry.itemTypeId.toUpperCase() : '';
		if (!isId(id) || targets.some((known) => known.itemTypeId === id)) {
			continue;
		}
		const ids = Array.isArray(entry.prefillIds) ? entry.prefillIds.filter((value) => typeof value === 'string' && isId(value)) : [];
		targets.push({ itemTypeId: id, prefillIds: unique(ids) });
	}
	return targets.length > 0 && targets[0].itemTypeId === first ? targets : own;
}

/**
 * The job's targets property for a list of targets: empty for a job with one target, which
 * target_itemtype and prefill_idlist hold on their own.
 */
export function serializeTargets(targets) {
	if (!Array.isArray(targets) || targets.length < 2) {
		return '';
	}
	return JSON.stringify(targets.map((target) => ({ itemTypeId: target.itemTypeId, prefillIds: target.prefillIds.slice() })));
}

// ---------------------------------------------------------------------------
// The allow-list
// ---------------------------------------------------------------------------

/**
 * Which allow-listed ItemTypes the user may use. A row of the allow-list names an
 * ItemType and, optionally, an identity: with one, only members of that identity
 * (direct or through other identities) may use the ItemType; without one, everyone
 * who can use the tool may. An ItemType on several rows is allowed when any of them
 * allows it.
 *
 * @param {Array<{ itemType: { id, name, label }, identityId: string, identityName: string }>} rows
 * @param {Iterable<string>} identityIds every identity of the user, as the client lists them
 * @returns {{ allowed: Array<{ id, name, label }>, restricted: Array<{ id, name, label, identityName }> }}
 *   allowed: the ItemTypes the user may use, sorted by label. restricted: the others, each
 *   with the name of an identity that may use it.
 */
export function allowedForUser(rows, identityIds) {
	const mine = new Set(Array.from(identityIds || [], (id) => String(id).trim().toUpperCase()));
	const allowed = new Map();
	const restricted = new Map();
	for (const row of rows) {
		if (!row || !row.itemType || !row.itemType.id) {
			continue;
		}
		const identityId = String(row.identityId || '').toUpperCase();
		if (identityId === '' || mine.has(identityId)) {
			allowed.set(row.itemType.id, row.itemType);
			restricted.delete(row.itemType.id);
		} else if (!allowed.has(row.itemType.id) && !restricted.has(row.itemType.id)) {
			restricted.set(row.itemType.id, { ...row.itemType, identityName: row.identityName || identityId });
		}
	}
	const byLabel = (a, b) => a.label.localeCompare(b.label);
	return { allowed: Array.from(allowed.values()).sort(byLabel), restricted: Array.from(restricted.values()).sort(byLabel) };
}

// ---------------------------------------------------------------------------
// What the Form shows for an ItemType
// ---------------------------------------------------------------------------

const labelOf = (thing) => (thing && (thing.label || thing.name)) || '';

/**
 * The key of the main sheet, for display.
 *
 * @param {object} schema A template schema.
 * @returns {{itemType: string, createOnly: boolean, columns: Array<{id: string, label: string}>}}
 */
export function describeKey(schema) {
	const main = schema.sheets[0];
	return {
		itemType: labelOf(main.itemType),
		createOnly: !!main.createOnly,
		columns: main.columns.filter((column) => column.isKey).map((column) => ({ id: column.id, label: column.label || column.property }))
	};
}

/**
 * What the sheet picker offers for an ItemType: one choice per RelationshipType.
 *
 * A choice that can be picked has `groups`, the parts a match key is made of.
 * A group stands for one or more columns that are in the match key together:
 *   - `parent`   the parent's Unique key. Fixed: a row has to name its parent.
 *   - `related`  the related item's Unique key, as a whole.
 *   - one group per property of the relationship itself.
 * `byDefault` marks the groups of the default match key.
 *
 * @param {object} schema The template schema asked for with all_relationships.
 * @returns {Array<object>} In the order of the relationship tabs.
 */
export function sheetChoices(schema) {
	const mainLabel = labelOf(schema.sheets[0].itemType);
	return (schema.relationships || []).map((relationship) => {
		const sheet = schema.sheets.find((candidate) => candidate.relationshipType && candidate.relationshipType.id === relationship.id) || null;
		const relatedLabel = labelOf(relationship.relatedItemType);
		const choice = {
			relationshipTypeId: relationship.id,
			name: relationship.name,
			label: relationship.label || relationship.name,
			relatedLabel,
			relatedItemTypeId: relationship.relatedItemType ? String(relationship.relatedItemType.id || '').toUpperCase() : '',
			selectable: !!relationship.canBeSheet && !!sheet,
			canBeIndented: !!relationship.canBeIndented,
			reason: relationship.canBeSheet ? null : relationship.reason,
			createOnly: !!(sheet && sheet.createOnly),
			relatedRequired: !!relationship.relatedRequired,
			hasRelated: !!relationship.relatedItemType,
			groups: []
		};
		if (!choice.selectable) {
			return choice;
		}
		if (choice.createOnly && choice.relatedRequired) {
			// A row would need a related item and has no way to name one.
			choice.selectable = false;
			choice.reason = 'A row of this relationship needs a related ' + relatedLabel + ', and ' + relatedLabel +
				' has no Unique properties a row could name it by.';
			return choice;
		}
		const columnsOf = (role) => sheet.columns.filter((column) => column.role === role);
		const names = (columns) => columns.map((column) => column.label || column.property).join(' + ');
		const parent = columnsOf('parentKey');
		const related = columnsOf('relatedKey');
		choice.groups.push({ key: 'parent', ids: parent.map((column) => column.id), label: mainLabel + ' (parent): ' + names(parent), fixed: true, byDefault: true });
		if (choice.createOnly) {
			return choice;
		}
		if (related.length > 0) {
			choice.groups.push({
				key: 'related', ids: related.map((column) => column.id), label: relatedLabel + ' (related): ' + names(related), fixed: false, byDefault: true
			});
		}
		for (const column of columnsOf('property')) {
			choice.groups.push({ key: column.id, ids: [column.id], label: column.label || column.property, fixed: false, byDefault: !!column.isMatchKey });
		}
		return choice;
	});
}

/** The match key a choice gets when the user changes nothing, as column ids. */
export function defaultMatchKey(choice) {
	return choice.groups.filter((group) => group.fixed || group.byDefault).flatMap((group) => group.ids);
}

/**
 * The match key an indented sheet starts with: the parent and the related item. Sequence stays an
 * ordinary column, so a line can be matched and renumbered without one being given.
 */
export function indentedMatchKey(choice) {
	return matchKeyFor(choice, ['related']);
}

/** The groups of a choice that a stored match key covers. The fixed groups always count. */
export function selectedGroups(choice, matchKey) {
	const ids = new Set(matchKey || []);
	return choice.groups.filter((group) => group.fixed || group.ids.every((id) => ids.has(id))).map((group) => group.key);
}

/** The match key for a set of ticked groups, as column ids in the order of the sheet's columns. */
export function matchKeyFor(choice, groupKeys) {
	const ticked = new Set(groupKeys || []);
	return choice.groups.filter((group) => group.fixed || ticked.has(group.key)).flatMap((group) => group.ids);
}

/**
 * The export options with everything removed that the ItemType does not offer:
 * sheets of relationships that are gone or can't be sheets, and match-key
 * columns that no longer exist. The sheets come back in the order of the choices.
 * A sheet stored without a match key gets the default one, as it does on the server.
 * Columns left out that a sheet no longer has are dropped. For the main sheet that takes
 * its sheet of the schema; without one, the main sheet's list is kept as it is.
 *
 * @param {object} options The options as parseOptions gives them.
 * @param {Array<object>} choices The choices of the first target (sheetChoices).
 * @param {object|null} mainSheet The main sheet of the first target.
 * @param {Array<{itemTypeId: string, choices: Array<object>, mainSheet: object|null}>} [more] The other targets,
 *   in their order. Their relationship sheets follow those of the first target, and options.targets
 *   comes back with one entry for each of them, when there are any.
 */
export function tidyOptions(options, choices, mainSheet = null, more = []) {
	const sheets = [];
	for (const choice of [choices, ...more.map((target) => target.choices)].flat()) {
		const stored = options.sheets.find((sheet) => sheet.relationshipTypeId === choice.relationshipTypeId);
		if (!stored || !choice.selectable) {
			continue;
		}
		const known = new Set(choice.groups.flatMap((group) => group.ids));
		const layout = choice.canBeIndented ? stored.layout : undefined;
		sheets.push(withLayout(withOmit({
			relationshipTypeId: choice.relationshipTypeId,
			name: choice.name,
			matchKey: stored.matchKey.length === 0 ? (layout === INDENTED_LAYOUT ? indentedMatchKey(choice) : defaultMatchKey(choice))
				: matchKeyFor(choice, selectedGroups(choice, stored.matchKey))
		}, (stored.omit || []).filter((id) => known.has(id))), layout));
	}
	const keptOmit = (omit, sheet) => {
		const ids = sheet ? new Set(sheet.columns.map((column) => column.id)) : null;
		return (omit || []).filter((id) => !ids || ids.has(id));
	};
	const tidy = withOmit({ sheets }, keptOmit(options.omit, mainSheet));
	if (more.length > 0) {
		tidy.targets = more.map((target) => {
			const stored = (options.targets || []).find((candidate) => candidate.itemTypeId === target.itemTypeId);
			return withOmit({ itemTypeId: target.itemTypeId }, keptOmit(stored && stored.omit, target.mainSheet));
		});
	}
	return tidy;
}

// ---------------------------------------------------------------------------
// The columns of a sheet
// ---------------------------------------------------------------------------

/**
 * Why a column can't be left out of the workbook, as a short text, or null when it can.
 * The key and the match key find the item or row a row stands for, the parent and
 * related key columns name the items a relationship row joins, and a required property
 * without a default needs a value in every new row.
 *
 * @param {object} column A column of the template schema.
 * @param {string[]} [matchKey] The match key chosen for a relationship sheet, by column
 *   id. Without it the column's own isMatchKey counts.
 */
export function lockedReason(column, matchKey) {
	if (column.role === 'parentKey') {
		return 'parent';
	}
	if (column.role === 'relatedKey') {
		return 'related item';
	}
	if (column.isKey) {
		return 'key';
	}
	if (matchKey ? matchKey.includes(column.id) : column.isMatchKey) {
		return 'match key';
	}
	if (column.required && (column.defaultValue === null || column.defaultValue === undefined || column.defaultValue === '')) {
		return 'required';
	}
	return null;
}

/**
 * The columns of a sheet as the column picker shows them, in the order of the sheet.
 *
 * @param {object} sheet A sheet of the template schema.
 * @param {string[]} [omit] The columns left out, by id.
 * @param {string[]} [matchKey] The match key chosen for a relationship sheet.
 * @returns {Array<{id: string, label: string, locked: string|null, included: boolean}>}
 *   A locked column is always included, also when omit names it.
 */
export function columnChoices(sheet, omit, matchKey) {
	const left = new Set(omit || []);
	return sheet.columns.map((column) => {
		const locked = lockedReason(column, matchKey);
		return { id: column.id, label: column.label || column.property || column.id, locked, included: locked !== null || !left.has(column.id) };
	});
}

/** A sheet without the columns only an indented sheet has. */
export function withoutIndentedColumns(sheet) {
	return sheet.columns.some((column) => column.indentedOnly) ? { ...sheet, columns: sheet.columns.filter((column) => !column.indentedOnly) } : sheet;
}

/** The first column of an indented sheet. */
const LEVEL_COLUMN = {
	id: 'level', role: 'level', property: null, label: 'Level', dataType: 'integer', required: true, isKey: false, isMatchKey: false, isClassification: false,
	source: 'grid', storedLength: null, precision: null, scale: null, pattern: null, minRange: 0, maxRange: null, rangeInclusive: true,
	defaultValue: null, classPath: null, listId: null, itemTarget: null
};

/**
 * The schema the workbook is built from: the schema with the columns that the options
 * leave out taken off their sheets. A column that can't be left out stays (see
 * lockedReason), and everything else of the schema is kept. Its definitionHash describes
 * the definitions, which leaving columns out does not change. The upload is checked
 * against the whole schema, where a column the workbook does not have leaves that value
 * of the item as it is.
 *
 * @param {object} schema The schema begin returned. Its match keys are the job's.
 * @param {object} options The export options the workbook is generated with.
 *   A sheet with the indented layout gets a Level column in place of its parent columns.
 */
export function narrowSchema(schema, options) {
	const storedOf = (sheet) => ((options && options.sheets) || [])
		.find((candidate) => !!sheet.relationshipType && candidate.relationshipTypeId === sheet.relationshipType.id);
	const laidOut = (sheet) => {
		const stored = storedOf(sheet);
		return !stored || stored.layout !== INDENTED_LAYOUT ? withoutIndentedColumns(sheet)
			: { ...sheet, layout: INDENTED_LAYOUT, columns: [LEVEL_COLUMN, ...sheet.columns.filter((column) => column.role !== 'parentKey')] };
	};
	const omitOf = (sheet, index) => {
		if (index === 0) {
			return (options && options.omit) || [];
		}
		if (sheet.kind === 'main') {
			const target = ((options && options.targets) || []).find((candidate) => candidate.itemTypeId === sheet.itemType.id);
			return (target && target.omit) || [];
		}
		const stored = storedOf(sheet);
		return (stored && stored.omit) || [];
	};
	return {
		...schema,
		sheets: schema.sheets.map((sheet, index) => {
			const left = new Set(omitOf(sheet, index));
			return laidOut(left.size === 0 ? sheet : { ...sheet, columns: sheet.columns.filter((column) => !left.has(column.id) || lockedReason(column) !== null) });
		})
	};
}

/**
 * What a match key means for the user, as sentences to show beside the choice.
 *
 * @param {object} choice A choice of sheetChoices that can be picked.
 * @param {string[]} groupKeys The ticked groups.
 */
export function matchKeyConsequences(choice, groupKeys, indented = false) {
	if (choice.createOnly) {
		return [
			'Create only: ' + choice.relatedLabel + ' has no Unique properties, so a row can\'t name a related item and can\'t be matched with an existing row. ' +
			'Every row of this sheet is added. Existing rows are not written to the workbook.'
		];
	}
	const ticked = new Set(groupKeys || []);
	const free = choice.groups.filter((group) => !group.fixed);
	const inKey = free.filter((group) => ticked.has(group.key));
	const lines = [];
	lines.push(
		'An import finds an existing row by the ticked values, so it can\'t change them: a row whose ticked values match no existing row is added as a new row.'
	);
	const sequence = inKey.find((group) => group.key === 'sort_order');
	if (sequence) {
		lines.push('Changing a row\'s ' + sequence.label + ' in the workbook therefore adds a new row; it doesn\'t renumber the old one.');
		if (indented) {
			lines.push('Every line of the indented sheet then needs a ' + sequence.label + ', and the same part can be listed more than once under a parent with different values.');
		}
	}
	const related = free.find((group) => group.key === 'related');
	if (related && !ticked.has('related')) {
		lines.push('The related item is not ticked: a row is found by the other values alone, and an import can change which item the row points to.');
	}
	if (inKey.length === 0) {
		lines.push('Only the parent is ticked: a row can be matched only where the parent has a single row of this relationship.');
	}
	if (choice.hasRelated && !choice.relatedRequired && ticked.has('related')) {
		lines.push('A row without a related item can\'t be matched: it is always added, and existing rows without one are not written to the workbook.');
	}
	return lines;
}

// ---------------------------------------------------------------------------
// Reading the items to pre-fill
// ---------------------------------------------------------------------------

/** A list cut into pieces of at most `size` entries. */
export function chunk(list, size = IDS_PER_REQUEST) {
	const pieces = [];
	for (let start = 0; start < list.length; start += size) {
		pieces.push(list.slice(start, start + size));
	}
	return pieces;
}

const unique = (list) => Array.from(new Set(list));

/**
 * The request that tells which item each selected id belongs to. An id names one
 * generation of a versionable item, and the item may have been versioned since
 * it was selected; its config_id stays the same through every generation.
 */
export function configIdsAml(itemTypeName, ids) {
	return '<Item type="' + escapeXml(itemTypeName) + '" action="get" select="config_id" idlist="' + ids.filter(isId).join(',') + '"/>';
}

/** The selected ids of a job, as its prefill_idlist holds them: in order, each once. */
export function parseIdList(text) {
	return unique(String(text || '').split(',').map((id) => id.trim()).filter(isId));
}

/**
 * The selection after adding and removing items: the ids kept stay in their order, the
 * added ones go at the end, and an id that is there already is not added again.
 */
export function editSelection(ids, { add = [], remove = [] } = {}) {
	const removed = new Set(remove);
	return unique([...ids, ...add].filter((id) => isId(id) && !removed.has(id)));
}

/** The request for the names of the selected items, for the list on the job's Form. */
export function selectionNamesAml(itemTypeName, ids) {
	return '<Item type="' + escapeXml(itemTypeName) + '" action="get" select="keyed_name" idlist="' + ids.filter(isId).join(',') + '"/>';
}

/** The properties to ask for on the main sheet's items. */
export function mainSelect(mainSheet) {
	return unique(['config_id', 'modified_on', ...mainSheet.columns.map((column) => column.property)]).join(',');
}

/**
 * The request for the items of the main sheet in their current generation.
 *
 * @param {object} mainSheet The first sheet of the schema.
 * @param {string[]} keys config_ids when the ItemType is versionable, ids otherwise.
 */
export function mainItemsAml(mainSheet, keys) {
	const type = escapeXml(mainSheet.itemType.name);
	const select = mainSelect(mainSheet);
	if (mainSheet.itemType.isVersionable) {
		return '<Item type="' + type + '" action="get" select="' + select + '">' +
			'<config_id condition="in">' + inList(keys.filter(isId)) + '</config_id><is_current>1</is_current></Item>';
	}
	return '<Item type="' + type + '" action="get" select="' + select + '" idlist="' + keys.filter(isId).join(',') + '"/>';
}

/** The request for the rows of a relationship sheet under the given parents. */
export function relationshipRowsAml(sheet, parentIds, relatedProperties = []) {
	const relatedKeys = sheet.columns.filter((column) => column.role === 'relatedKey').map((column) => column.property).concat(relatedProperties);
	const own = sheet.columns.filter((column) => column.role === 'property').map((column) => column.property);
	const select = unique(['source_id', relatedKeys.length ? 'related_id(' + relatedKeys.join(',') + ')' : 'related_id', 'sort_order', 'modified_on', ...own]);
	return '<Item type="' + escapeXml(sheet.itemType.name) + '" action="get" select="' + select.join(',') + '">' +
		'<source_id condition="in">' + inList(parentIds.filter(isId)) + '</source_id></Item>';
}

/** The current generations of versionable items, by their config_id. */
export function currentItemsAml(itemTypeName, configIds) {
	return '<Item type="' + escapeXml(itemTypeName) + '" action="get" select="config_id">' +
		'<config_id condition="in">' + inList(configIds.filter(isId)) + '</config_id><is_current>1</is_current></Item>';
}

/**
 * Reads the lines of a structure, level by level, from the given items down. A line's related item is followed
 * to its current generation, because that is the generation an import writes to.
 *
 * @param {function(string): Promise<Array<object>>} query Runs AML and gives the items found.
 * @param {object} sheet The indented sheet.
 * @param {Array<object>} parents The items at the top.
 * @param {string} relatedTypeName The ItemType of the related items.
 * @returns {Promise<{linesOf: Map<string, Array<object>>, currentOf: Map<string, string>}>} The lines by the id
 *   of their parent, and the id of the current generation by config_id.
 */
export async function readIndentedStructure(query, sheet, parents, relatedTypeName) {
	const linesOf = new Map();
	const currentOf = new Map();
	let frontier = parents.map((parent) => parent.id);
	while (frontier.length > 0) {
		frontier.forEach((id) => linesOf.set(id, []));
		const found = [];
		for (const part of chunk(frontier)) {
			found.push(...(await query(relationshipRowsAml(sheet, part, ['config_id']))));
		}
		const configs = new Set();
		for (const line of found) {
			(linesOf.get(textOf(line, 'source_id')) || []).push(line);
			const config = configOfLine(line);
			if (config && !currentOf.has(config)) {
				configs.add(config);
			}
		}
		for (const part of chunk(Array.from(configs))) {
			for (const item of await query(currentItemsAml(relatedTypeName, part))) {
				currentOf.set(textOf(item, 'config_id'), item.id);
			}
		}
		frontier = Array.from(new Set(found.map((line) => currentOf.get(configOfLine(line))).filter((id) => id && !linesOf.has(id))));
	}
	return { linesOf, currentOf };
}

const relatedItemOf = (line) => (line.props.related_id && line.props.related_id.item) || null;
const configOfLine = (line) => (relatedItemOf(line) ? textOf(relatedItemOf(line), 'config_id') : '');

/**
 * The rows of an indented sheet for existing items: each item on level 0, with its lines below it, in Sequence
 * order, at every level. A part used in several places is listed in full in each of them. A line that leads back
 * to a part above it is listed, but not followed.
 *
 * @param {object} sheet The indented sheet.
 * @param {Array<object>} parents The items at the top.
 * @param {{linesOf: Map, currentOf: Map}} structure What readIndentedStructure read.
 * @param {number} [limit] The most rows to make.
 * @returns {{rows: Array<object>, unreadable: number, withoutRelated: number, loops: number, sameLine: number, tooMany: boolean}}
 *   The counts are of distinct lines, except sameLine, the parents that have two lines the match columns can't tell apart.
 */
export function indentedRows(sheet, parents, structure, limit = SHEET_ROW_LIMIT) {
	const relatedKeys = sheet.columns.filter((column) => column.role === 'relatedKey');
	const own = sheet.columns.filter((column) => column.role === 'property');
	const relatedInMatchKey = relatedKeys.some((column) => column.isMatchKey);
	const matchColumns = [...relatedKeys, ...own].filter((column) => column.isMatchKey);
	const result = { rows: [], unreadable: 0, withoutRelated: 0, loops: 0, sameLine: 0, tooMany: false };
	const sequence = (line) => {
		const number = Number(textOf(line, 'sort_order'));
		return textOf(line, 'sort_order') !== '' && Number.isFinite(number) ? number : Number.MAX_SAFE_INTEGER;
	};
	const valuesOf = (line) => {
		const values = {};
		for (const column of relatedKeys) {
			values[column.id] = relatedItemOf(line) ? cellValue(column, relatedItemOf(line).props[column.property]) : '';
		}
		for (const column of own) {
			values[column.id] = cellValue(column, line.props[column.property]);
		}
		return values;
	};
	const known = new Map();
	const linesBelow = (id) => {
		if (!known.has(id)) {
			const lines = (structure.linesOf.get(id) || []).filter((line) => {
				if (relatedKeys.length === 0 || relatedItemOf(line)) {
					return true;
				}
				const related = line.props.related_id;
				if (related && (related.value || related.restricted)) {
					result.unreadable++;
					return false;
				}
				if (relatedInMatchKey) {
					result.withoutRelated++;
					return false;
				}
				return true;
			}).sort((a, b) => sequence(a) - sequence(b) || (a.id < b.id ? -1 : 1));
			const seen = new Set();
			let same = false;
			for (const line of lines) {
				const values = valuesOf(line);
				const key = JSON.stringify(matchColumns.map((column) => String(values[column.id]).trim().toLowerCase()));
				same = same || seen.has(key);
				seen.add(key);
			}
			result.sameLine += same ? 1 : 0;
			known.set(id, lines);
		}
		return known.get(id);
	};
	const closing = new Set();
	const addRow = (row) => {
		if (result.rows.length >= limit) {
			result.tooMany = true;
			return false;
		}
		result.rows.push(row);
		return true;
	};
	const walk = (id, level, path) => {
		for (const line of linesBelow(id)) {
			if (!addRow({ values: { level, ...valuesOf(line) }, modifiedOn: textOf(line, 'modified_on') || null })) {
				return false;
			}
			const child = structure.currentOf.get(configOfLine(line));
			if (!child) {
				continue;
			}
			if (path.has(child)) {
				closing.add(line.id);
				continue;
			}
			path.add(child);
			const more = walk(child, level + 1, path);
			path.delete(child);
			if (!more) {
				return false;
			}
		}
		return true;
	};
	for (const parent of parents) {
		const values = { level: 0 };
		for (const column of relatedKeys) {
			values[column.id] = cellValue(column, parent.props[column.property]);
		}
		if (!addRow({ values, modifiedOn: null }) || !walk(parent.id, 1, new Set([parent.id]))) {
			break;
		}
	}
	result.loops = closing.size;
	return result;
}

/** The value a workbook cell gets for a property: the keyed name for an item property (a File's is its file name), the stored text otherwise. */
function cellValue(column, prop) {
	if (!prop) {
		return '';
	}
	if (column.dataType === 'item' || column.dataType === 'file') {
		return prop.keyedName || '';
	}
	return prop.value === null || prop.value === undefined ? '' : String(prop.value);
}

const textOf = (item, name) => (item.props[name] ? String(item.props[name].value || '') : '');

/**
 * Puts the items read for the main sheet in the order they were selected in.
 *
 * @param {string[]} selectedIds The ids as selected.
 * @param {Map<string, string>|null} configIdOf For a versionable ItemType: selected id -> config_id. Null otherwise.
 * @param {Array<object>} items The items read, current generations.
 * @returns {{items: Array<object>, missing: number}} `missing` counts the selected items that were not read:
 *   deleted since, or not visible to the user.
 */
export function orderMainItems(selectedIds, configIdOf, items) {
	const byKey = new Map();
	for (const item of items) {
		byKey.set(configIdOf ? textOf(item, 'config_id') : item.id, item);
	}
	const ordered = [];
	const taken = new Set();
	let missing = 0;
	for (const id of selectedIds) {
		const key = configIdOf ? configIdOf.get(id) : id;
		const item = key ? byKey.get(key) : null;
		if (!item) {
			missing++;
		} else if (!taken.has(item)) {
			taken.add(item);
			ordered.push(item);
		}
	}
	return { items: ordered, missing };
}

/** The pre-fill rows of the main sheet. */
export function mainRows(mainSheet, items) {
	return items.map((item) => {
		const values = {};
		for (const column of mainSheet.columns) {
			values[column.id] = cellValue(column, item.props[column.property]);
		}
		return { values, modifiedOn: textOf(item, 'modified_on') || null };
	});
}

/**
 * The pre-fill rows of a relationship sheet.
 *
 * Rows are grouped by parent, in the order of the main sheet, and sorted by
 * their sequence within a parent. Two kinds of existing rows are left out,
 * and counted:
 *   - `unreadable`: the related item exists but the user can't read it, so
 *     its key is not known;
 *   - `withoutRelated`: the row has no related item while the related key is
 *     part of the match key. Such a row can't be matched on import and would
 *     come back as a new row.
 *
 * @param {object} sheet A relationship sheet of the schema.
 * @param {Array<object>} parents The items of the main sheet, in sheet order.
 * @param {Array<object>} relationships The relationship items read.
 * @returns {{rows: Array<object>, unreadable: number, withoutRelated: number}}
 */
export function relationshipRows(sheet, parents, relationships) {
	const parentKeys = sheet.columns.filter((column) => column.role === 'parentKey');
	const relatedKeys = sheet.columns.filter((column) => column.role === 'relatedKey');
	const own = sheet.columns.filter((column) => column.role === 'property');
	const relatedInMatchKey = relatedKeys.some((column) => column.isMatchKey);
	const parentIndex = new Map(parents.map((parent, index) => [parent.id, index]));
	const sequence = (item) => {
		const number = Number(textOf(item, 'sort_order'));
		return textOf(item, 'sort_order') !== '' && Number.isFinite(number) ? number : Number.MAX_SAFE_INTEGER;
	};

	const result = { rows: [], unreadable: 0, withoutRelated: 0 };
	const kept = [];
	for (const relationship of relationships) {
		const index = parentIndex.get(textOf(relationship, 'source_id'));
		if (index === undefined) {
			continue;
		}
		const related = relationship.props.related_id || null;
		const relatedItem = related && related.item ? related.item : null;
		if (relatedKeys.length > 0 && !relatedItem) {
			if (related && (related.value || related.restricted)) {
				result.unreadable++;
				continue;
			}
			if (relatedInMatchKey) {
				result.withoutRelated++;
				continue;
			}
		}
		kept.push({ relationship, relatedItem, index });
	}
	kept.sort((a, b) =>
		a.index - b.index || sequence(a.relationship) - sequence(b.relationship) || (a.relationship.id < b.relationship.id ? -1 : 1));

	for (const { relationship, relatedItem, index } of kept) {
		const parent = parents[index];
		const values = {};
		for (const column of parentKeys) {
			values[column.id] = cellValue(column, parent.props[column.property]);
		}
		for (const column of relatedKeys) {
			values[column.id] = relatedItem ? cellValue(column, relatedItem.props[column.property]) : '';
		}
		for (const column of own) {
			values[column.id] = cellValue(column, relationship.props[column.property]);
		}
		result.rows.push({ values, modifiedOn: textOf(relationship, 'modified_on') || null });
	}
	return result;
}

/**
 * A file name for the workbook of a job: the job number ("MCE-000123.xlsx"), whatever its targets are.
 * Letters that a file system refuses are replaced.
 */
export function workbookFileName(jobNumber) {
	const name = String(jobNumber || '').replace(/[\\/:*?"<>|\u0000-\u001F]/g, '_').trim();
	return (name || 'workbook') + '.xlsx';
}
