// lib/pptx/decrypt.js — 解密密码保护的 OOXML（pptx/docx/xlsx）。
// 加密文件是 OLE 复合文档（MS-CFB），内含 EncryptionInfo + EncryptedPackage 流。
// 支持 MS-OFFCRYPTO 标准加密（ECMA-376 Standard，4.2）与 Agile 加密（4.4，XML 描述符）。
import crypto from "node:crypto";
import { XMLParser } from "fast-xml-parser";
import { readOLE, isOLE } from "./ole.js";

export class WrongPasswordError extends Error {
	constructor(message = "密码错误") {
		super(message);
		this.name = "WrongPasswordError";
		this.code = "wrong-password";
	}
}

export class DecryptError extends Error {
	constructor(message) {
		super(message);
		this.name = "DecryptError";
		this.code = "decrypt-error";
	}
}

export function isEncryptedOOXML(buf) {
	return isOLE(buf);
}

// ── 工具 ──────────────────────────────────────────────────────────────────────

const u32 = (n) => {
	const b = Buffer.alloc(4);
	b.writeUInt32LE(n >>> 0);
	return b;
};

function hash(alg, ...parts) {
	const h = crypto.createHash(alg);
	for (const p of parts) h.update(p);
	return h.digest();
}

// 短于 blockSize 时用 0x36 填充（MS-OFFCRYPTO _normalize_key）
function normalize(buf, size) {
	if (buf.length >= size) return buf.subarray(0, size);
	return Buffer.concat([buf, Buffer.alloc(size - buf.length, 0x36)]);
}

function aesCbc(key, iv, data) {
	const d = crypto.createDecipheriv(`aes-${key.length * 8}-cbc`, key, iv);
	d.setAutoPadding(false);
	return Buffer.concat([d.update(data), d.final()]);
}

function aesEcb(key, data) {
	const d = crypto.createDecipheriv(`aes-${key.length * 8}-ecb`, key, null);
	d.setAutoPadding(false);
	return Buffer.concat([d.update(data), d.final()]);
}

const AGILE_ALGS = { sha1: "sha1", sha256: "sha256", sha384: "sha384", sha512: "sha512" };
function agileHash(name) {
	const a = AGILE_ALGS[String(name ?? "SHA1").toLowerCase()];
	if (a === undefined) throw new DecryptError(`不支持的哈希算法: ${name}`);
	return a;
}

// Agile 块密钥常量（MS-OFFCRYPTO 2.3.4.11）
const BLK_VERIFIER_INPUT = Buffer.from([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]);
const BLK_VERIFIER_VALUE = Buffer.from([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]);
const BLK_ENCRYPTED_KEY = Buffer.from([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]);

// h = H(salt + utf16le(pw)); for i in 0..spin-1: h = H(LE32(i) + h)
function iteratedHash(alg, salt, password, spin) {
	let h = hash(alg, salt, Buffer.from(password, "utf16le"));
	for (let i = 0; i < spin; i++) h = hash(alg, u32(i), h);
	return h;
}

// ── Agile 加密（EncryptionInfo 4.4，XML 描述符）──────────────────────────────

function decryptAgile(info, pkg, password) {
	const xml = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_", parseTagValue: false }).parse(
		info.subarray(8).toString("utf8"),
	);
	const enc = xml?.encryption;
	if (enc === undefined) throw new DecryptError("EncryptionInfo XML 无法解析");
	const keyData = enc.keyData;
	if (keyData === undefined) throw new DecryptError("EncryptionInfo 缺少 keyData");
	const b64 = (s) => Buffer.from(String(s ?? ""), "base64");

	// 可能有多个 keyEncryptor（密码/证书），取 password 类型
	let ke = enc.keyEncryptors?.keyEncryptor;
	if (Array.isArray(ke)) ke = ke.find((k) => String(k["@_uri"] ?? "").includes("keyEncryptor/password")) ?? ke[0];
	const ek = ke?.["p:encryptedKey"] ?? ke?.encryptedKey;
	if (ek === undefined) throw new DecryptError("未找到密码加密器（可能仅支持证书加密）");

	const alg = agileHash(ek["@_hashAlgorithm"] ?? keyData["@_hashAlgorithm"]);
	const keyBits = Number(ek["@_keyBits"] ?? keyData["@_keyBits"]);
	const blockSize = Number(keyData["@_blockSize"] ?? 16);
	if (!(keyBits === 128 || keyBits === 192 || keyBits === 256)) throw new DecryptError(`不支持的密钥长度: ${keyBits}`);
	const spin = Number(ek["@_spinCount"]);
	const salt = b64(ek["@_saltValue"]);
	const eVHI = b64(ek["@_encryptedVerifierHashInput"]);
	const eVHV = b64(ek["@_encryptedVerifierHashValue"]);
	const eKV = b64(ek["@_encryptedKeyValue"]);
	if (salt.length === 0 || eKV.length === 0) throw new DecryptError("EncryptionInfo 字段缺失");

	const h = iteratedHash(alg, salt, password, spin);
	const keyLen = keyBits >> 3;
	const key1 = hash(alg, h, BLK_VERIFIER_INPUT).subarray(0, keyLen);
	const key2 = hash(alg, h, BLK_VERIFIER_VALUE).subarray(0, keyLen);
	const key3 = hash(alg, h, BLK_ENCRYPTED_KEY).subarray(0, keyLen);
	const iv = normalize(salt, blockSize);

	// 校验密码
	const verifierInput = aesCbc(key1, iv, eVHI);
	const expected = aesCbc(key2, iv, eVHV);
	const actual = hash(alg, verifierInput);
	if (!crypto.timingSafeEqual(actual.subarray(0, expected.length), expected)) throw new WrongPasswordError();

	// 会话密钥（secretKey 直接作为包密钥）
	const secret = aesCbc(key3, iv, eKV);

	// 包：u64 totalSize + 4096B 分段 AES-CBC，iv_i = H(keyDataSalt + LE32(i))[:16]
	const dataAlg = agileHash(keyData["@_hashAlgorithm"] ?? ek["@_hashAlgorithm"]);
	const kdSalt = b64(keyData["@_saltValue"]);
	const total = Number(pkg.readBigUInt64LE(0));
	const out = [];
	let remaining = total;
	let i = 0;
	for (let off = 8; off < pkg.length && remaining > 0; off += 4096, i++) {
		const seg = pkg.subarray(off, Math.min(off + 4096, pkg.length));
		const ivN = hash(dataAlg, kdSalt, u32(i)).subarray(0, 16);
		let dec = aesCbc(secret, ivN, seg);
		if (dec.length > remaining) dec = dec.subarray(0, remaining);
		out.push(dec);
		remaining -= dec.length;
	}
	if (remaining > 0) throw new DecryptError("EncryptedPackage 数据不完整");
	return Buffer.concat(out);
}

// ── 标准加密（EncryptionInfo 4.2，二进制头，SHA1 + AES-ECB）──────────────────

// SHA1 派生（迭代 50000 次 + 0x36/0x5c 异或展开），keySizeBits: 128/192/256
export function standardKeyFromPassword(password, keySizeBits, salt) {
	let h = hash("sha1", salt, Buffer.from(password, "utf16le"));
	for (let i = 0; i < 50000; i++) h = hash("sha1", u32(i), h);
	const hFinal = hash("sha1", h, u32(0));
	const xor = (fill) => {
		const b = Buffer.alloc(64, fill);
		for (let i = 0; i < 20; i++) b[i] = hFinal[i] ^ fill;
		return b;
	};
	return Buffer.concat([hash("sha1", xor(0x36)), hash("sha1", xor(0x5c))]).subarray(0, keySizeBits >> 3);
}

// AES-ECB 校验密钥是否正确
export function standardVerifyKey(key, encVerifier, encVerifierHash) {
	const verifier = aesEcb(key, encVerifier);
	const expected = aesEcb(key, encVerifierHash).subarray(0, 20);
	return crypto.timingSafeEqual(hash("sha1", verifier), expected);
}

function decryptStandard(info, pkg, password) {
	// 8B 版本头 + 32B 头（flags,sizeExtra,algId,algIdHash,keySize,providerType,r1,r2）
	if (info.length < 40) throw new DecryptError("EncryptionInfo 头不完整");
	const keySize = info.readUInt32LE(16); // bits
	const keyLen = keySize >> 3;
	if (!(keyLen === 16 || keyLen === 24 || keyLen === 32)) throw new DecryptError(`不支持的密钥长度: ${keySize}`);
	// verifier 块从偏移 40 起：u32 saltSize, salt(16), encVerifier(16), u32 hashSize, encVerifierHash(32)
	let off = 40;
	const saltSize = info.readUInt32LE(off);
	off += 4;
	const salt = info.subarray(off, off + Math.max(16, saltSize)).subarray(0, 16);
	off += 16;
	const encVerifier = info.subarray(off, off + 16);
	off += 16;
	off += 4; // verifierHashSize
	const encVerifierHash = info.subarray(off, off + 32);

	const key = standardKeyFromPassword(password, keySize, salt);
	if (!standardVerifyKey(key, encVerifier, encVerifierHash)) throw new WrongPasswordError();

	// 包：u32 totalSize（u64 低 32 位），数据从偏移 8 起，AES-ECB
	const total = pkg.readUInt32LE(0);
	const dec = aesEcb(key, pkg.subarray(8, 8 + Math.ceil(total / 16) * 16));
	return dec.subarray(0, total);
}

// ── 入口 ──────────────────────────────────────────────────────────────────────

// buf: 加密文件全文；password: 打开密码。返回解密后的 zip Buffer。
export function decryptOOXML(buf, password) {
	const ole = readOLE(buf);
	const info = ole.getStream("EncryptionInfo");
	if (info === null) throw new DecryptError("不是加密的 pptx（未找到 EncryptionInfo；旧版 .ppt 二进制格式不支持）");
	const pkg = ole.getStream("EncryptedPackage");
	if (pkg === null || pkg.length < 8) throw new DecryptError("未找到 EncryptedPackage 流");
	const major = info.readUInt16LE(0);
	const minor = info.readUInt16LE(2);
	let out;
	if (major === 4 && minor === 4) {
		out = decryptAgile(info, pkg, String(password ?? ""));
	} else if (minor === 2 && (major === 2 || major === 3 || major === 4)) {
		out = decryptStandard(info, pkg, String(password ?? ""));
	} else {
		throw new DecryptError(`不支持的加密版本 ${major}.${minor}（可扩展加密不支持）`);
	}
	// 解密成功必须是 zip
	if (!(out[0] === 0x50 && out[1] === 0x4b && out[2] === 0x03 && out[3] === 0x04)) throw new WrongPasswordError();
	return out;
}
