import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { decryptOOXML, isEncryptedOOXML, WrongPasswordError, standardKeyFromPassword, standardVerifyKey } from "../lib/pptx/decrypt.js";
import { readOLE, isOLE } from "../lib/pptx/ole.js";

const fixtures = new URL("./fixtures/", import.meta.url).pathname;

let failures = 0;
async function check(name, fn) {
	try {
		await fn();
		console.log(`PASS ${name}`);
	} catch (error) {
		failures += 1;
		console.log(`FAIL ${name}: ${error.stack?.split("\n")[0]}`);
	}
}

// ── OLE 识别 ─────────────────────────────────────────────────────────────────
const encBuf = await readFile(path.join(fixtures, "enc-ascii.pptx"));
const plainBuf = await readFile(path.join(fixtures, "composed.pptx"));
await check("isEncryptedOOXML 识别加密/明文", () => {
	assert.equal(isEncryptedOOXML(encBuf), true);
	assert.equal(isEncryptedOOXML(plainBuf), false);
});

await check("OLE 读取流目录", () => {
	const ole = readOLE(encBuf);
	assert.ok(ole.names.includes("EncryptionInfo"));
	assert.ok(ole.names.includes("EncryptedPackage"));
	const info = ole.getStream("EncryptionInfo");
	assert.ok(info.length > 8);
	assert.equal(info.readUInt16LE(0), 4); // agile 4.4
	assert.equal(info.readUInt16LE(2), 4);
});

// ── Agile 端到端 ─────────────────────────────────────────────────────────────
await check("agile 解密（ASCII 密码）", () => {
	const out = decryptOOXML(encBuf, "secret123");
	assert.equal(out.equals(plainBuf), true);
});

await check("agile 解密（中文密码）", async () => {
	const buf = await readFile(path.join(fixtures, "enc-cjk.pptx"));
	const rich = await readFile(path.join(fixtures, "rich.pptx"));
	const out = decryptOOXML(buf, "测试密码123");
	assert.equal(out.equals(rich), true);
});

await check("错误密码抛 WrongPasswordError", () => {
	assert.throws(() => decryptOOXML(encBuf, "wrong"), (e) => e instanceof WrongPasswordError);
	assert.throws(() => decryptOOXML(encBuf, ""), (e) => e instanceof WrongPasswordError);
});

// ── 标准加密：msoffcrypto doctest 向量 ───────────────────────────────────────
await check("standard 密钥派生向量", () => {
	const salt = Buffer.from("e88266490c5bd1eebd2b4394e3f830ef", "hex");
	const key = standardKeyFromPassword("Password1234_", 128, salt);
	assert.equal(key.toString("hex"), "40b13a71f90b966e375408f2d181a1aa");
});

await check("standard 校验向量", () => {
	const key = Buffer.from("40b13a71f90b966e375408f2d181a1aa", "hex");
	// encryptedVerifier = b'Qos.\x96o\xac\x17\xb1\xc5\xd7\xd8\xcc6\xc9('
	const ev = Buffer.from([0x51, 0x6f, 0x73, 0x2e, 0x96, 0x6f, 0xac, 0x17, 0xb1, 0xc5, 0xd7, 0xd8, 0xcc, 0x36, 0xc9, 0x28]);
	const evh = Buffer.from([0x2b, 0x61, 0x68, 0xda, 0xbe, 0x29, 0x11, 0xad, 0x2b, 0xd3, 0x7c, 0x17, 0x46, 0x74, 0x5c, 0x14, 0xd3, 0xcf, 0x1b, 0xb1, 0x40, 0xa4, 0x8f, 0x4e, 0x6f, 0x3d, 0x23, 0x88, 0x08, 0x72, 0xb1, 0x6a]);
	assert.equal(standardVerifyKey(key, ev, evh), true);
	assert.equal(standardVerifyKey(Buffer.alloc(16), ev, evh), false);
});

console.log(failures === 0 ? "ALL PASS" : `${failures} FAILURES`);
process.exit(failures === 0 ? 0 : 1);
