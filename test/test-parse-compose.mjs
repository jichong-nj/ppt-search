import assert from "node:assert/strict";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { parsePptx } from "../lib/pptx/parse.js";
import { composePptx, applyTitleOverride, applyReplace } from "../lib/pptx/compose.js";

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

// ── parse ────────────────────────────────────────────────────────────────────
const a = await parsePptx(path.join(fixtures, "deck-a", "source.pptx"));
await check("parse deckA count/size", () => {
	assert.equal(a.slideCount, 3);
	assert.equal(a.slides.length, 3);
	assert.equal(a.slideSize.cx, 9144000);
});
await check("parse deckA titles", () => {
	assert.equal(a.slides[0].title, "年度总结");
	assert.equal(a.slides[2].title, "数据概览");
});
await check("parse deckA body + table", () => {
	assert.ok(a.slides[0].body.includes("营收增长 25%"));
	assert.ok(a.slides[2].body.includes("营收"));
	assert.ok(a.slides[2].body.includes("1.2亿"));
});
await check("parse deckA notes", () => {
	assert.ok(a.slides[0].notes.includes("先讲营收"));
	assert.equal(a.slides[1].notes, "");
});
const b = await parsePptx(path.join(fixtures, "deck-b", "source.pptx"));
await check("parse deckB", () => {
	assert.equal(b.slideCount, 2);
	assert.ok(b.slides[0].body.includes("revenue grew 25 percent"));
});
await check("parse invalid file rejects", async () => {
	const bad = path.join(fixtures, "bad.pptx");
	await writeFile(bad, "not a zip");
	await assert.rejects(() => parsePptx(bad));
});

// ── compose ──────────────────────────────────────────────────────────────────
const outPath = path.join(fixtures, "composed.pptx");
const res = await composePptx({
	items: [
		{ deckId: "deck-a", n: 1 },
		{ deckId: "deck-b", n: 1, titleOverride: "改过的标题X" },
		{ deckId: "deck-a", n: 3 },
		{ deckId: "deck-b", n: 2, replace: [{ find: "old-brand", replace: "new-brand" }] },
	],
	outPath,
	decksRoot: fixtures,
	baseDeckId: "deck-a",
});
await check("compose returns count", () => {
	assert.equal(res.slideCount, 4);
	assert.ok(res.bytes > 1000);
});

const c = await parsePptx(outPath);
await check("composed slide order & count", () => {
	assert.equal(c.slideCount, 4);
	assert.equal(c.slides[0].title, "年度总结");
	assert.equal(c.slides[1].title, "改过的标题X");
	assert.equal(c.slides[2].title, "数据概览");
	assert.equal(c.slides[3].title, "技术架构 Architecture");
});
await check("composed title override removed old title text", () => {
	assert.ok(!c.slides[1].body.includes("Quarterly Revenue"));
	assert.ok(c.slides[1].body.includes("revenue grew 25 percent"));
});
await check("composed replace applied", () => {
	// deckA slide1 无 old-brand，deckB slide2 有 —— 检查 deckB 那页
	assert.ok(c.slides[3].body.includes("new-brand 出现") || c.slides[3].body.includes("new-brand"));
	assert.ok(!c.slides[3].body.includes("old-brand"));
});
await check("composed keeps table content", () => {
	assert.ok(c.slides[2].body.includes("1.2亿"));
});
await check("composed base slideSize kept", () => {
	assert.equal(c.slideSize.cx, 9144000);
});

// ── compose 边界 ─────────────────────────────────────────────────────────────
await check("compose empty items rejects", async () => {
	await assert.rejects(() => composePptx({ items: [], outPath, decksRoot: fixtures }));
});
await check("compose out-of-range rejects", async () => {
	await assert.rejects(() => composePptx({ items: [{ deckId: "deck-a", n: 99 }], outPath, decksRoot: fixtures }));
});
await check("compose same-deck only", async () => {
	const out2 = path.join(fixtures, "composed2.pptx");
	const r = await composePptx({ items: [{ deckId: "deck-b", n: 2 }, { deckId: "deck-b", n: 1 }], outPath: out2, decksRoot: fixtures });
	const c2 = await parsePptx(out2);
	assert.equal(c2.slideCount, 2);
	assert.equal(c2.slides[0].title, "技术架构 Architecture");
});

// ── 单元：文本修改函数 ───────────────────────────────────────────────────────
await check("applyReplace only touches a:t", () => {
	const xml = '<p:c name="old-brand"><a:t>old-brand 文本</a:t></p:c>';
	const out = applyReplace(xml, [{ find: "old-brand", replace: "NB" }]);
	assert.ok(out.includes("<a:t>NB 文本</a:t>"));
	assert.ok(out.includes('p:c name="old-brand"'));
});
await check("applyReplace handles entities", () => {
	const xml = "<a:t>A &amp; B old</a:t>";
	const out = applyReplace(xml, [{ find: "old", replace: "新&amp;旧" }]);
	assert.ok(out.includes("A &amp; B 新&amp;amp;旧") || out.includes("A &amp; B 新&amp;旧"));
});
await check("applyTitleOverride no title keeps xml", () => {
	const xml = "<p:sp><p:txBody><a:t>x</a:t></p:txBody></p:sp>";
	assert.equal(applyTitleOverride(xml, "T"), xml);
});

// ── 纯 JS 校验合成包合法性（不依赖 LibreOffice）─────────────────────────────
await check("composed pptx is a valid package (4 slides)", async () => {
	const JSZip = (await import("jszip")).default;
	const zip = await JSZip.loadAsync(await readFile(outPath));
	assert.ok(zip.file("[Content_Types].xml") !== null, "has [Content_Types].xml");
	const slideNames = Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide[^/]+\.xml$/.test(n));
	assert.ok(slideNames.length >= 4, `expected >=4 slide parts, got ${slideNames.length}`);
	const parsed = await parsePptx(outPath);
	assert.equal(parsed.slideCount, 4);
});

process.exit(failures === 0 ? 0 : 1);
