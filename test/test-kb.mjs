// test/test-kb.mjs — kb 集成测试（纯 JS 渲染模型）。
import assert from "node:assert/strict";
import { cp, mkdtemp, mkdir, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { createKb } from "../lib/kb.js";

const here = path.dirname(fileURLToPath(import.meta.url));
const fixtures = path.join(here, "fixtures");

let passed = 0;
function ok(cond, msg) {
	assert.ok(cond, msg);
	passed += 1;
	console.log(`  ✓ ${msg}`);
}

async function waitForStatus(kb, deckId, wanted, timeoutMs = 120_000) {
	const deadline = Date.now() + timeoutMs;
	while (Date.now() < deadline) {
		const { decks } = kb.list();
		const deck = decks.find((d) => d.id === deckId);
		if (deck !== undefined && (deck.status === wanted || deck.status === "error")) return deck;
		await new Promise((r) => setTimeout(r, 250));
	}
	throw new Error(`timeout waiting for deck ${deckId} to become ${wanted}`);
}

async function main() {
	const dataDir = await mkdtemp(path.join(os.tmpdir(), "ps-kb-"));
	const kb = createKb({ dataDir });
	await kb.init();
	ok(true, "init on empty dir");

	// 上传两个 fixture
	const tmp = await kb.store.tmpDir();
	const a = path.join(tmp, "a.pptx");
	const b = path.join(tmp, "b.pptx");
	await cp(path.join(fixtures, "deck-a", "source.pptx"), a);
	await cp(path.join(fixtures, "deck-b", "source.pptx"), b);
	const deckA = await kb.addUpload({ name: "年度总结.pptx", tmpPath: a });
	const deckB = await kb.addUpload({ name: "english deck.pptx", tmpPath: b });

	const readyA = await waitForStatus(kb, deckA.id, "ready");
	ok(readyA.status === "ready", `deck A ready (error=${readyA.error ?? "none"})`);
	ok(readyA.slideCount === 3, "deck A has 3 slides");
	ok(readyA.renderCount === 3, "deck A built 3 render models");
	const readyB = await waitForStatus(kb, deckB.id, "ready");
	ok(readyB.status === "ready" && readyB.slideCount === 2, "deck B ready with 2 slides");

	// 检索
	const r1 = kb.search("营收");
	ok(r1.hits.length >= 1 && r1.hits.some((h) => h.deckId === deckA.id && h.n === 3), "cjk search 营收 hits deck-a#3");
	ok(r1.hits[0].renderUrl.includes("/render/"), "hit has renderUrl");
	const r2 = kb.search("architecture");
	ok(r2.hits.some((h) => h.deckId === deckB.id), "english search architecture hits deck-b");
	const r3 = kb.search("营收", { deckId: deckB.id });
	ok(!r3.hits.some((h) => h.deckId === deckA.id), "deckId filter excludes deck-a");

	// 详情
	const detail = await kb.detail(deckA.id);
	ok(detail.slides.length === 3 && detail.slides[0].title === "年度总结", "detail returns slides");

	// 渲染模型文件（含懒生成：删除后应自动重建）
	const rj = await kb.renderFile(deckA.id, 1);
	ok(rj.endsWith(".json"), "renderFile path");
	await rm(path.join(kb.store.deckDir(deckA.id), "render"), { recursive: true, force: true });
	const rj2 = await kb.renderFile(deckA.id, 2);
	ok(rj2.endsWith(".json"), "renderFile lazy regeneration");

	// 合成
	const comp = await kb.compose({
		items: [
			{ deckId: deckA.id, n: 1, titleOverride: "新封面" },
			{ deckId: deckB.id, n: 2, replace: [{ find: "old-brand", replace: "new-brand" }] },
		],
		name: "我的合成",
	});
	ok(comp.slideCount === 2 && comp.name === "我的合成.pptx", "compose returns 2 slides with name");
	const comps = await kb.compositions();
	ok(comps.items.some((c) => c.id === comp.id), "compositions lists new item");
	const dl = await kb.downloadFile(comp.id);
	ok(dl.name === "我的合成.pptx", "downloadFile resolves composition");

	// deck 源文件下载
	const dlDeck = await kb.downloadFile(deckA.id);
	ok(dlDeck.name === "年度总结.pptx", "downloadFile falls back to deck source");

	// 删除 → 检索消失
	await kb.remove(deckB.id);
	ok(kb.search("architecture").hits.length === 0, "search empty after delete");

	// 重启：新 kb 实例重建索引
	const kb2 = createKb({ dataDir });
	await kb2.init();
	const { decks } = kb2.list();
	ok(decks.length === 1 && decks[0].id === deckA.id, "library persisted across restart");
	ok(kb2.search("营收").hits.some((h) => h.n === 3), "index rebuilt from slides.json on restart");

	console.log(`\nALL ${passed} KB CHECKS PASSED`);
}

main().catch((e) => {
	console.error("FAIL:", e);
	process.exit(1);
});
