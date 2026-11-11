/**
 * Tests for the MD5 of the browser (mce-md5.js), against the MD5 of Node.
 */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';

const { Md5, md5OfBlob } = await import(new URL('../../1.0.0/Innovator/Client/customer/masscreateedit/mce-md5.js', import.meta.url).href);

const nodeMd5 = (bytes) => createHash('md5').update(bytes).digest('hex').toUpperCase();

function bytesOf(length, seed) {
	const bytes = new Uint8Array(length);
	let value = seed;
	for (let index = 0; index < length; index++) {
		value = (value * 1103515245 + 12345) >>> 0;
		bytes[index] = value >>> 24;
	}
	return bytes;
}

test('md5: the test values of RFC 1321', () => {
	const text = (value) => new Md5().update(new TextEncoder().encode(value)).digest();
	assert.equal(text(''), 'D41D8CD98F00B204E9800998ECF8427E');
	assert.equal(text('a'), '0CC175B9C0F1B6A831C399E269772661');
	assert.equal(text('abc'), '900150983CD24FB0D6963F7D28E17F72');
	assert.equal(text('message digest'), 'F96B697D7CB7938D525A2F31AAF161D0');
	assert.equal(text('12345678901234567890123456789012345678901234567890123456789012345678901234567890'), '57EDF4A22BE3C955AC49DA2E2107B67A');
});

test('md5: every length around a block, in one part and in many', () => {
	for (let length = 0; length <= 200; length++) {
		const bytes = bytesOf(length, length + 1);
		assert.equal(new Md5().update(bytes).digest(), nodeMd5(bytes), 'length ' + length);
		const parts = new Md5();
		for (let start = 0; start < length; start += 7) {
			parts.update(bytes.subarray(start, Math.min(length, start + 7)));
		}
		assert.equal(parts.digest(), nodeMd5(bytes), 'length ' + length + ' in parts');
	}
});

test('md5: a file read in parts, larger than one part', async () => {
	const bytes = bytesOf(9 * 1024 * 1024 + 13, 42);
	const seen = [];
	const md5 = await md5OfBlob(new Blob([bytes]), (read) => seen.push(read));
	assert.equal(md5, nodeMd5(bytes));
	assert.deepEqual(seen, [4 * 1024 * 1024, 8 * 1024 * 1024, bytes.length]);
	assert.equal(await md5OfBlob(new Blob([])), 'D41D8CD98F00B204E9800998ECF8427E');
});
