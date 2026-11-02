// lib/pptx/ole.js — 极简 CFBF（OLE 复合文档，MS-CFB）读取器。
// 只需按名字提取根目录下的流（加密 pptx 里的 EncryptionInfo / EncryptedPackage）。
// 支持：FAT/DIFAT 链、目录链、mini-FAT/mini 流（小流）。

const OLE_MAGIC = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;
const MAX_CHAIN = 64 * 1024 * 1024; // 扇区链上限，防死循环

export function isOLE(buf) {
	return buf.length >= 8 && buf.subarray(0, 8).equals(OLE_MAGIC);
}

export class OleError extends Error {
	constructor(message) {
		super(message);
		this.name = "OleError";
	}
}

export function readOLE(buf) {
	if (!isOLE(buf)) throw new OleError("不是 OLE 复合文档");
	if (buf.length < 512) throw new OleError("OLE 头不完整");
	const sectorShift = buf.readUInt16LE(0x1e);
	const miniShift = buf.readUInt16LE(0x20);
	if (sectorShift < 7 || sectorShift > 16) throw new OleError("OLE sectorShift 异常");
	const ssz = 1 << sectorShift;
	const msz = 1 << miniShift;
	const numFat = buf.readUInt32LE(0x2c);
	const firstDir = buf.readUInt32LE(0x30);
	const miniCutoff = buf.readUInt32LE(0x38);
	const firstMiniFat = buf.readUInt32LE(0x3c);
	const firstDifat = buf.readUInt32LE(0x44);

	const sectorSlice = (sid) => {
		const off = (sid + 1) * ssz;
		if (off + ssz > buf.length) throw new OleError("OLE 扇区越界");
		return buf.subarray(off, off + ssz);
	};

	// ── DIFAT：FAT 扇区 id 列表（头 109 个 + 链式扩展）──
	const fatSids = [];
	for (let i = 0; i < 109; i++) {
		const v = buf.readUInt32LE(0x4c + i * 4);
		if (v === FREESECT || v === ENDOFCHAIN) break;
		fatSids.push(v);
	}
	{
		let sid = firstDifat;
		let guard = 0;
		while (sid !== FREESECT && sid !== ENDOFCHAIN && sid < FREESECT) {
			if (++guard > MAX_CHAIN) throw new OleError("DIFAT 链过长");
			const sec = sectorSlice(sid);
			const n = ssz >> 2;
			for (let i = 0; i < n - 1; i++) {
				const v = sec.readUInt32LE(i * 4);
				if (v === FREESECT) continue;
				fatSids.push(v);
			}
			sid = sec.readUInt32LE((n - 1) * 4);
		}
	}
	if (numFat > 0 && fatSids.length < numFat) throw new OleError("FAT 扇区不足");

	// ── FAT：扇区分配表 ──
	const fat = new Uint32Array(fatSids.length * (ssz >> 2));
	{
		let p = 0;
		for (const sid of fatSids) {
			const sec = sectorSlice(sid);
			for (let i = 0; i < ssz >> 2; i++) fat[p++] = sec.readUInt32LE(i * 4);
		}
	}
	const chain = (start) => {
		const out = [];
		let s = start;
		while (s < ENDOFCHAIN) {
			if (s >= fat.length) throw new OleError("扇区链越界");
			out.push(s);
			if (out.length > MAX_CHAIN) throw new OleError("扇区链过长");
			s = fat[s];
		}
		return out;
	};

	// ── 目录项 ──
	const entries = [];
	{
		let idx = 0;
		for (const sid of chain(firstDir)) {
			const sec = sectorSlice(sid);
			for (let o = 0; o + 128 <= ssz; o += 128) {
				const e = sec.subarray(o, o + 128);
				const type = e[0x42];
				const nameLen = e.readUInt16LE(0x40);
				let name = "";
				if (type === 1 || type === 2 || type === 5) {
					name = e.subarray(0, Math.max(0, nameLen - 2)).toString("utf16le");
				}
				const start = e.readUInt32LE(0x74);
				const size = e.readUInt32LE(0x78) + e.readUInt32LE(0x7c) * 0x100000000;
				entries.push({ id: idx++, name, type, start, size });
			}
		}
	}
	const root = entries.find((e) => e.type === 5);
	if (root === undefined) throw new OleError("OLE 缺少根存储");

	// ── mini 流与 mini-FAT ──
	let miniStream = Buffer.alloc(0);
	if (root.size > 0) {
		const parts = chain(root.start).map(sectorSlice);
		miniStream = Buffer.concat(parts, Math.min(root.size, parts.length * ssz)).subarray(0, root.size);
	}
	let miniFat = new Uint32Array(0);
	if (firstMiniFat < ENDOFCHAIN) {
		const parts = chain(firstMiniFat).map(sectorSlice);
		const raw = Buffer.concat(parts);
		miniFat = new Uint32Array(raw.buffer, raw.byteOffset, raw.length >> 2);
	}

	const readStream = (e) => {
		const out = [];
		if (e.size < miniCutoff) {
			let s = e.start;
			let guard = 0;
			while (s < ENDOFCHAIN) {
				if (s * msz + msz > miniStream.length && s * msz >= miniStream.length) throw new OleError("mini 扇区越界");
				out.push(miniStream.subarray(s * msz, Math.min((s + 1) * msz, miniStream.length)));
				if (++guard > MAX_CHAIN) throw new OleError("mini 链过长");
				if (s >= miniFat.length) break;
				s = miniFat[s];
			}
		} else {
			for (const sid of chain(e.start)) out.push(sectorSlice(sid));
		}
		const all = Buffer.concat(out);
		return all.subarray(0, Math.min(e.size, all.length));
	};

	return {
		names: entries.filter((e) => e.type === 2).map((e) => e.name),
		getStream(name) {
			const e = entries.find((x) => x.type === 2 && x.name === name);
			if (e === undefined) return null;
			return readStream(e);
		},
	};
}
