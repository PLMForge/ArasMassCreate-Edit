/**
 * A schema for a made-up ItemType that has every data type and every rule of
 * the template schema contract: required and optional variants, ranges with
 * and without their ends, date patterns, lists with awkward labels, a class
 * structure and a class-specific property.
 */

function column(id, dataType, overrides = {}) {
	return {
		id,
		role: 'property',
		property: id,
		label: id.replace(/_/g, ' ').replace(/^./, (first) => first.toUpperCase()),
		dataType,
		required: false,
		isKey: false,
		isMatchKey: false,
		isClassification: false,
		source: 'form',
		storedLength: null,
		precision: null,
		scale: null,
		pattern: null,
		minRange: null,
		maxRange: null,
		rangeInclusive: false,
		defaultValue: null,
		classPath: null,
		listId: null,
		itemTarget: null,
		...overrides
	};
}

const required = { required: true };

export const UNIT_LIST = 'C0FFEE00C0FFEE00C0FFEE00C0FFEE01';
export const REGION_LIST = 'C0FFEE00C0FFEE00C0FFEE00C0FFEE02';
export const EMPTY_LIST = 'C0FFEE00C0FFEE00C0FFEE00C0FFEE03';

export const CLASS_PATHS = ['Hardware/Fastener', 'Hardware/Bracket, Angle', 'Software', 'Monthly, Numbered Day'];

export function allTypesSchema() {
	return {
		formatVersion: 1,
		database: 'SampleDatabase',
		generatedOn: '2026-11-05T09:30:00',
		language: 'en',
		sheets: [
			{
				kind: 'main',
				itemType: { id: 'AB12AB12AB12AB12AB12AB12AB12AB12', name: 'Sample Item', label: 'Sample Item', isVersionable: false },
				relationshipType: null,
				createOnly: false,
				classPaths: CLASS_PATHS,
				columns: [
					column('code', 'string', { ...required, isKey: true, storedLength: 32 }),
					column('title', 'ml_string', { storedLength: 64 }),
					column('short_code', 'string', { storedLength: 4 }),
					column('notes', 'text'),
					column('body', 'formatted text'),
					column('classification', 'string', { storedLength: 512, isClassification: true }),
					column('count_required', 'integer', { ...required, minRange: 1, maxRange: 999, rangeInclusive: true }),
					column('count_optional', 'integer', { minRange: 0, maxRange: 10, rangeInclusive: false }),
					column('count_free', 'integer'),
					column('price_required', 'decimal', { ...required, precision: 10, scale: 2, minRange: 0, maxRange: 99999.99, rangeInclusive: true }),
					column('price_exclusive', 'decimal', { ...required, precision: 10, scale: 2, minRange: 0, maxRange: 1, rangeInclusive: false }),
					column('price_optional', 'decimal', { precision: 10, scale: 2, minRange: 0, maxRange: 99999.99, rangeInclusive: true }),
					column('ratio_min', 'float', { ...required, minRange: 0, rangeInclusive: false }),
					column('ratio_max', 'float', { ...required, maxRange: 100, rangeInclusive: true }),
					column('ratio_any', 'float', required),
					column('ratio_optional', 'float'),
					column('start_date', 'date', { ...required, pattern: 'short_date' }),
					column('end_date', 'date', { pattern: 'long_date' }),
					column('released_on', 'date', { ...required, pattern: 'short_date_time' }),
					column('checked_on', 'date'),
					column('is_active', 'boolean', required),
					column('is_template', 'boolean'),
					column('unit', 'list', { ...required, listId: UNIT_LIST }),
					column('unit_optional', 'list', { listId: UNIT_LIST }),
					column('region', 'filter list', { listId: REGION_LIST, pattern: 'country' }),
					column('tags', 'mv_list', { listId: UNIT_LIST }),
					column('owner', 'item', {
						itemTarget: { id: 'CD34CD34CD34CD34CD34CD34CD34CD34', name: 'Identity', label: 'Identity', isVersionable: false, isPolymorphic: false }
					}),
					column('colour', 'color', required),
					column('accent', 'color list'),
					column('thread_size', 'string', { storedLength: 16, classPath: 'Hardware' }),
					column('retired_choice', 'list', { ...required, listId: EMPTY_LIST })
				]
			}
		],
		lists: {
			[UNIT_LIST]: {
				name: 'Units',
				values: [
					{ value: 'EA', label: 'Each' },
					{ value: 'KG', label: 'Kilogram, net' },
					{ value: 'M, long', label: '' },
					{ value: 'L', label: 'Volume' },
					{ value: 'GAL', label: 'Volume' },
					{ value: 'PK', label: 'Pack' },
					{ value: 'Pack', label: '' }
				]
			},
			[REGION_LIST]: {
				name: 'Regions',
				values: [
					{ value: 'north', label: 'North' },
					{ value: 'south', label: 'South' }
				]
			},
			[EMPTY_LIST]: { name: 'Retired', values: [] }
		},
		definitionHash: 'sha256:0a1b2c3d4e5f60718293a4b5c6d7e8f90a1b2c3d4e5f60718293a4b5c6d7e8f9'
	};
}

/** The 1-based position and the letter of a column of the fixture's only sheet. */
export function columnAt(schema, id) {
	const position = schema.sheets[0].columns.findIndex((entry) => entry.id === id) + 1;
	let letter = '';
	for (let n = position; n > 0; n = Math.floor((n - 1) / 26)) {
		letter = String.fromCharCode(65 + ((n - 1) % 26)) + letter;
	}
	return { position, letter };
}
