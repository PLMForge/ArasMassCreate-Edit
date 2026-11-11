/**
 * Mass Create & Edit - the MD5 of a file, computed in the browser.
 *
 * The import names each file the user picked by its MD5, which the server checks
 * against the checksum of the stored File. The browser's own crypto API has no
 * MD5, so it is computed here, a chunk at a time, so that a large file is never
 * held in memory as a whole.
 *
 * MD5 is used to tell files apart, not to protect anything (RFC 1321).
 */

// The per-round shift amounts and the sine-derived constants of RFC 1321.
const SHIFTS = [
	7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22,
	5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14, 20,
	4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23,
	6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21
];
const CONSTANTS = new Int32Array(64);
for (let index = 0; index < 64; index++) {
	CONSTANTS[index] = Math.floor(Math.abs(Math.sin(index + 1)) * 0x100000000) | 0;
}

/** The size of the parts a file is read in. */
const CHUNK_BYTES = 4 * 1024 * 1024;

/** An MD5 computed over bytes given in parts. */
export class Md5 {
	constructor() {
		this.state = new Int32Array([0x67452301, 0xefcdab89 | 0, 0x98badcfe | 0, 0x10325476]);
		this.block = new Uint8Array(64);
		this.blockLength = 0;
		this.length = 0;
		this.words = new Int32Array(16);
	}

	/** Adds bytes. */
	update(bytes) {
		const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
		this.length += data.length;
		let offset = 0;
		if (this.blockLength > 0) {
			const take = Math.min(64 - this.blockLength, data.length);
			this.block.set(data.subarray(0, take), this.blockLength);
			this.blockLength += take;
			offset = take;
			if (this.blockLength < 64) {
				return this;
			}
			this.transform(this.block, 0);
			this.blockLength = 0;
		}
		for (; offset + 64 <= data.length; offset += 64) {
			this.transform(data, offset);
		}
		if (offset < data.length) {
			this.block.set(data.subarray(offset), 0);
			this.blockLength = data.length - offset;
		}
		return this;
	}

	/** The MD5 of the bytes added, as 32 uppercase hexadecimal digits. Ends the computation. */
	digest() {
		const bits = this.length * 8;
		const padding = new Uint8Array((this.blockLength < 56 ? 56 : 120) - this.blockLength + 8);
		padding[0] = 0x80;
		const view = new DataView(padding.buffer);
		view.setUint32(padding.length - 8, bits >>> 0, true);
		view.setUint32(padding.length - 4, Math.floor(bits / 0x100000000), true);
		this.update(padding);
		let hex = '';
		for (const word of this.state) {
			for (let shift = 0; shift < 32; shift += 8) {
				hex += ((word >>> shift) & 0xff).toString(16).padStart(2, '0');
			}
		}
		return hex.toUpperCase();
	}

	/** One 64-byte block. */
	transform(bytes, offset) {
		const words = this.words;
		for (let index = 0; index < 16; index++) {
			const at = offset + index * 4;
			words[index] = bytes[at] | (bytes[at + 1] << 8) | (bytes[at + 2] << 16) | (bytes[at + 3] << 24);
		}
		let [a, b, c, d] = this.state;
		for (let index = 0; index < 64; index++) {
			let f;
			let g;
			if (index < 16) {
				f = (b & c) | (~b & d);
				g = index;
			} else if (index < 32) {
				f = (d & b) | (~d & c);
				g = (5 * index + 1) % 16;
			} else if (index < 48) {
				f = b ^ c ^ d;
				g = (3 * index + 5) % 16;
			} else {
				f = c ^ (b | ~d);
				g = (7 * index) % 16;
			}
			const sum = (a + f + CONSTANTS[index] + words[g]) | 0;
			a = d;
			d = c;
			c = b;
			b = (b + ((sum << SHIFTS[index]) | (sum >>> (32 - SHIFTS[index])))) | 0;
		}
		this.state[0] = (this.state[0] + a) | 0;
		this.state[1] = (this.state[1] + b) | 0;
		this.state[2] = (this.state[2] + c) | 0;
		this.state[3] = (this.state[3] + d) | 0;
	}
}

/**
 * The MD5 of a file (a Blob), read a part at a time.
 *
 * @param {Blob} blob
 * @param {(read: number) => void} [onProgress] Called with the bytes read so far.
 * @returns {Promise<string>} 32 uppercase hexadecimal digits.
 */
export async function md5OfBlob(blob, onProgress) {
	const md5 = new Md5();
	for (let start = 0; start < blob.size; start += CHUNK_BYTES) {
		const part = await blob.slice(start, Math.min(blob.size, start + CHUNK_BYTES)).arrayBuffer();
		md5.update(new Uint8Array(part));
		if (onProgress) {
			onProgress(Math.min(blob.size, start + CHUNK_BYTES));
		}
	}
	return md5.digest();
}
