// Method:  FRG_MCE_BulkEditInit
// Type:    client Method, JavaScript (init handler of a command bar item)
// Purpose: decides whether a "Bulk Edit" entry of a search grid is shown and enabled.
//          It is the init handler of both entries: the one in the grid's right-click
//          menu and the button on the search toolbar.
//
// Result:
//   { hidden: true }                   the grid's ItemType is not on the allow-list for
//                                      the user's identities, or the settings could not
//                                      be read
//   { hidden: false, disabled: ... }   disabled while no saved row is selected
//
// The allow-list and the row limit are kept on the Mass Create & Edit configuration
// item. They are read once per client session, with the user's own permissions, and
// kept on the main window as JSON text (FRG_MCE_bulkEditSettings), where
// FRG_MCE_BulkEditClick finds them too. A change to the configuration item shows
// after the client is loaded again.
//
// This handler answers at once, so the settings are requested in the background and
// the entry stays hidden until they arrive. The toolbar button is then drawn again from
// here; the right-click menu needs no redraw.

const SETTINGS_KEY = 'FRG_MCE_bulkEditSettings';
const REQUEST_KEY = 'FRG_MCE_bulkEditSettingsRequest';
const mainWindow = aras.getMainWindow();

// A row of an item that is not saved yet has nothing to fill a workbook with.
const savedRowCount = () => {
	let selectedRows;
	try {
		selectedRows = options.getState().grid.selectedRows;
	} catch (error) {
		// The search grid was closed while the settings were on their way.
		return 0;
	}
	const itemState = ArasModules.item;
	if (!itemState || typeof itemState.isNew !== 'function') {
		return selectedRows.length;
	}
	return selectedRows.filter((rowId) => !itemState.isNew(options.itemTypeName, rowId)).length;
};

const decide = (settings) => {
	if (!options.itemType || !settings.allowedTypeIds.includes(options.itemType.id)) {
		return { hidden: true };
	}
	return { hidden: false, disabled: savedRowCount() === 0 };
};

const stored = mainWindow[SETTINGS_KEY];
if (stored) {
	return decide(JSON.parse(stored));
}

// One request per window. After a failure the entries stay hidden in this window, and
// the next search grid that opens asks again.
if (!window[REQUEST_KEY]) {
	const request =
		'<Item type="FRG_MassCreateEdit_Config" action="get" select="max_rows">' +
		'<Relationships>' +
		'<Item type="FRG_MassCreateEdit_AllowedType" action="get" select="related_id,identity_id" related_expand="0"/>' +
		'</Relationships>' +
		'</Item>';
	window[REQUEST_KEY] = ArasModules.soap(request, { method: 'ApplyItem', async: true })
		.then((result) => {
			// There is one configuration item. Should there be more, the first one counts.
			const config = ArasModules.xml.selectSingleNode(result, '//Result/Item');
			// A row with an identity allows its ItemType only to members of that identity, direct
			// or through other identities.
			const mine = new Set(String(aras.getIdentityList() || '').split(',').map((id) => id.trim().toUpperCase()));
			const allowedTypeIds = Array.from(ArasModules.xml.selectNodes(config, 'Relationships/Item'))
				.filter((row) => {
					const identityId = (aras.getItemProperty(row, 'identity_id') || '').toUpperCase();
					return identityId === '' || mine.has(identityId);
				})
				.map((row) => aras.getItemProperty(row, 'related_id'));
			mainWindow[SETTINGS_KEY] = JSON.stringify({
				maxRows: parseInt(aras.getItemProperty(config, 'max_rows'), 10) || 0,
				allowedTypeIds
			});
		})
		.catch(() => {
			// No configuration item the user can read: Bulk Edit is not offered.
		});
}

if (target.include_events) {
	window[REQUEST_KEY].then(() => {
		const loaded = mainWindow[SETTINGS_KEY];
		const current = control.isConnected && control.data && control.data.get(target.id);
		if (!loaded || !current) {
			return;
		}
		control.data.set(target.id, { ...current, ...decide(JSON.parse(loaded)) });
		control.render();
	});
}

return { hidden: true };
