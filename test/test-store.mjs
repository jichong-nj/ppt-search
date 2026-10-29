import assert from "node:assert/strict";
import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { createStore, assertId } from "../lib/store.js";

let failures = 0;
async function check(name, fn) {
	try {
		await fn();
		console.log(`PASS ${name}`);
	} catch (error) {
		failures += 1;
		console.log(`FAIL ${name}: ${error.message}`);
	}
}

const dir = await mkdtemp(path.join(tmpdir(), "ps-store-"));
const store = createStore(dir);

await check("ensureDirs", async () => {
	await store.ensureDirs();
	await readFile(path.join(dir, "decks"), () => undefined).catch(() => undefined);
});

await check("empty library", async () => {
	const lib = await store.loadLibrary();
	assert.deepEqual(lib, { decks: {} });
});

await check("library round-trip", async () => {
	const lib = await store.loadLibrary();
	lib.decks["deck-a1"] = { id: "deck-a1", name: "x.pptx", status: "ready" };
	await store.saveLibrary(lib);
	const again = await store.loadLibrary();
	assert.equal(again.decks["deck-a1"].name, "x.pptx");
});

await check("deck meta atomic write + read", async () => {
	await store.writeDeckMeta({ id: "deck-a1", status: "ready", slideCount: 3 });
	const meta = await store.readDeckMeta("deck-a1");
	assert.equal(meta.slideCount, 3);
});

await check("deck slides round-trip", async () => {
	await store.writeDeckSlides("deck-a1", [{ n: 1, title: "t", body: "b", notes: "" }]);
	const slides = await store.readDeckSlides("deck-a1");
	assert.equal(slides[0].title, "t");
});

await check("genId format", () => {
	const id = store.genId("deck");
	assert.match(id, /^deck-[a-z0-9]+-[0-9a-f]{4}$/);
});

await check("assertId rejects traversal", () => {
	assert.throws(() => assertId("../etc"));
	assert.throws(() => assertId("a/b"));
	assert.throws(() => assertId("A-b"));
	assert.throws(() => assertId(""));
	assert.equal(assertId("deck-abc123"), "deck-abc123");
});

await check("renderPath validates page", () => {
	assert.throws(() => store.renderPath("deck-a1", 0));
	assert.throws(() => store.renderPath("deck-a1", "1/../2"));
});

await check("mediaPath rejects traversal", () => {
	assert.throws(() => store.mediaPath("deck-a1", "../secret"));
	assert.throws(() => store.mediaPath("deck-a1", "a/b.png"));
	assert.throws(() => store.mediaPath("deck-a1", ""));
	assert.ok(store.mediaPath("deck-a1", "image1.png").endsWith("/media/image1.png"));
});

await check("composition list", async () => {
	await writeFile(store.compositionPath("comp-1"), "fake-pptx-bytes");
	await store.writeCompositionMeta("comp-1", { name: "合成一.pptx", slideCount: 2 });
	const items = await store.listCompositions();
	assert.equal(items.length, 1);
	assert.equal(items[0].id, "comp-1");
	assert.equal(items[0].name, "合成一.pptx");
	assert.equal(items[0].slideCount, 2);
	await store.removeComposition("comp-1");
	assert.equal((await store.listCompositions()).length, 0);
});

await check("removeDeckFiles", async () => {
	await store.removeDeckFiles("deck-a1");
	assert.equal(await store.readDeckMeta("deck-a1"), null);
});

await check("tmpDir unique", async () => {
	const a = await store.tmpDir();
	const b = await store.tmpDir();
	assert.notEqual(a, b);
	assert.ok(a.startsWith(store.tmpRoot));
});

process.exit(failures === 0 ? 0 : 1);
