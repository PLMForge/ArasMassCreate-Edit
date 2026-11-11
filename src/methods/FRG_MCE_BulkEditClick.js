// Method:  FRG_MCE_BulkEditClick
// Type:    client Method, JavaScript (click handler of a command bar item)
// Purpose: starts a Mass Create & Edit job for the rows selected in a search grid.
//          It is the click handler of both "Bulk Edit" entries: the one in the grid's
//          right-click menu and the button on the search toolbar.
//
// The job is a new FRG_MassCreateEdit item, opened in an item window the way
// "Create New" opens one. Nothing is saved here: the user chooses the options on the
// job's Form and saves it. Two properties are filled in:
//   target_itemtype   the ItemType of the search grid
//   prefill_idlist    the ids of the selected rows, separated by commas
//
// When more rows are selected than the configuration item allows (max_rows), no job
// is started and the user gets a notification.

const stored = aras.getMainWindow().FRG_MCE_bulkEditSettings;
if (!stored) {
	// FRG_MCE_BulkEditInit shows the entries only after it has stored the settings.
	return;
}
const settings = JSON.parse(stored);
const itemType = options.itemType;
if (!itemType || !settings.allowedTypeIds.includes(itemType.id)) {
	return;
}

// A row of an item that is not saved yet has nothing to fill a workbook with.
const itemState = ArasModules.item;
const canTellNew = Boolean(itemState) && typeof itemState.isNew === 'function';
const ids = options.getState().grid.selectedRows.filter(
	(rowId) => !(canTellNew && itemState.isNew(options.itemTypeName, rowId))
);
if (ids.length === 0) {
	return;
}

if (settings.maxRows > 0 && ids.length > settings.maxRows) {
	const notify = aras.getNotifyByContext(window);
	notify(
		ids.length + ' items are selected, and Bulk Edit takes no more than ' +
			settings.maxRows + ' at a time. Select fewer items and try again.',
		{ type: 'warning' }
	);
	return;
}

return Promise.resolve(aras.newItem('FRG_MassCreateEdit')).then((job) => {
	if (!job) {
		return;
	}
	aras.setItemProperty(job, 'target_itemtype', itemType.id);
	// The ItemType's name goes on the property too, for the Form to show.
	aras.setItemPropertyAttribute(job, 'target_itemtype', 'type', 'ItemType');
	aras.setItemPropertyAttribute(job, 'target_itemtype', 'keyed_name', itemType.name);
	aras.setItemProperty(job, 'prefill_idlist', ids.join(','));
	aras.itemsCache.updateItem(job);
	return aras.uiShowItemEx(job);
});
