// Method:  FRG_MCE_JobFormPopulated
// Type:    client Method, JavaScript (Form event onformpopulated of Form FRG_MassCreateEdit)
// Purpose: hands the job Form to the client module that draws the tool on it.
//
// The Form has one HTML field with an empty element, "mce-app". This Method loads
// mce-job-form.js and calls its attach() with the Form's window; attach() builds the tool
// in that element, and the tool fills the job's item window.
//
// The module is loaded into the main window, not into the Form's frame, so that
// generating a workbook or importing one goes on when the job is saved or refreshed.
// The main window is asked to do the import.
//
// CLIENT_VERSION goes into the address of the module, so a browser loads the new files
// after an update: after a change to the tool's client files, change it.

const CLIENT_VERSION = '1.0.17';

const formWindow = window;
const mainWindow = aras.getMainWindow();
const address = aras.getBaseURL() + '/customer/masscreateedit/mce-job-form.js?v=' + CLIENT_VERSION;

mainWindow.Function('address', 'return import(address);')(address)
	.then((module) => module.attach(formWindow))
	.catch((error) => {
		const host = formWindow.document.getElementById('mce-app');
		if (host) {
			host.textContent = 'Mass Create & Edit could not be loaded: ' + (error && error.message ? error.message : error);
		}
	});
