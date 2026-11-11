/**
 * Runs the workbook module the way a browser page would: as an ES module
 * served over HTTP, in a realm of its own that has an AMD loader and `fetch`
 * but no `require`, no `process` and no Buffer. A module it imports comes
 * from the client folder, as the page's server would serve it. It builds a
 * workbook, reads it back, and prints what happened as one line of JSON.
 *
 * It is a script, not a test: mce-workbook-loader.test.mjs starts it in a
 * child process, because loading an ES module into a sandbox needs the node
 * option --experimental-vm-modules.
 */

import vm from 'node:vm';
import { readFileSync } from 'node:fs';
import { webcrypto } from 'node:crypto';

const clientFolder = new URL('../../../1.0.0/Innovator/Client/customer/masscreateedit/', import.meta.url);
const pageFolder = 'http://server.example/Innovator/Client/customer/masscreateedit/';
const schema = JSON.parse(readFileSync(new URL('../fixtures/part-with-bom.schema.json', import.meta.url), 'utf8'));

const fetched = [];
const registered = [];
const page = {
	console, setTimeout, clearTimeout, setInterval, clearInterval, queueMicrotask, TextEncoder, TextDecoder, URL, performance,
	CompressionStream,
	DecompressionStream,
	crypto: webcrypto,
	navigator: { userAgent: 'sandbox' },
	// The page's server: it serves the files of the client folder.
	fetch: async (url) => {
		fetched.push(String(url));
		const source = readFileSync(new URL(String(url).slice(pageFolder.length), clientFolder), 'utf8');
		return { ok: true, status: 200, text: async () => source };
	},
	define: Object.assign((...parameters) => registered.push(parameters), { amd: {} })
};
page.window = page;
page.self = page;
const context = vm.createContext(page);

// The page's server also serves the other modules of the client folder, at their address, which may
// carry the cache tag of the page (?v=...). A browser keeps one module per address, tag included.
// Anything else (a bare name, a node: module) is something a browser can not provide.
const modules = new Map();
function loadModule(url) {
	const path = url.split('?')[0];
	if (!path.startsWith(pageFolder)) {
		throw new Error('The module imported ' + url + ', which a browser can not provide.');
	}
	if (!modules.has(url)) {
		modules.set(url, (async () => {
			const loaded = new vm.SourceTextModule(readFileSync(new URL(path.slice(pageFolder.length), clientFolder), 'utf8'), {
				context,
				identifier: url,
				initializeImportMeta: (meta) => {
					meta.url = url;
				},
				importModuleDynamically: (specifier) => loadModule(new URL(specifier, url).href)
			});
			await loaded.link((specifier) => loadModule(new URL(specifier, url).href));
			await loaded.evaluate();
			return loaded;
		})());
	}
	return modules.get(url);
}

const module = await loadModule(pageFolder + 'mce-workbook.js');

const { buildWorkbook, parseWorkbook } = module.namespace;
const file = await buildWorkbook(schema, {
	job: { id: '0123456789ABCDEF0123456789ABCDEF', token: 'c2FtcGxlLXRva2Vu' },
	rows: [[{ values: { item_number: 'P-0001', unit: 'KG', make_buy: 'Buy', effective_date: '2026-03-01T00:00:00' }, modifiedOn: '2026-10-31T08:15:42' }]]
});
const parsed = await parseWorkbook(file);

console.log(JSON.stringify({
	fetched,
	registeredWithAmd: registered.length,
	isArrayBuffer: Object.prototype.toString.call(file) === '[object ArrayBuffer]',
	bytes: file.byteLength,
	jobId: parsed.meta.job.id,
	rows: parsed.sheets[0].rows,
	leakedGlobals: ['module', 'exports'].filter((name) => name in page)
}));
