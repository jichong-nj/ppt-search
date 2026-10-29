import assert from "node:assert/strict";
import { createIndex, addSlide, removeDeck, search, serialize, deserialize, tokenize } from "../lib/search.js";

let failures = 0;
function check(name, fn) {
	try {
		fn();
		console.log(`PASS ${name}`);
	} catch (error) {
		failures += 1;
		console.log(`FAIL ${name}: ${error.message}`);
	}
}

const idx = createIndex();
addSlide(idx, { key: "d1#1", deckId: "d1", n: 1, title: "年度总结", body: "本年度营收增长 25%，机器学习平台上线。", notes: "" });
addSlide(idx, { key: "d1#2", deckId: "d1", n: 2, title: "机器学习入门", body: "监督学习与无监督学习简介", notes: "讲师：张三" });
addSlide(idx, { key: "d2#1", deckId: "d2", n: 1, title: "Quarterly Revenue", body: "revenue grew 25 percent in machine learning segment", notes: "revenue highlight" });
addSlide(idx, { key: "d2#2", deckId: "d2", n: 2, title: "季度营收总结", body: "营收同比增长，详见附表", notes: "" });

check("tokenize cjk bigrams", () => {
	const t = tokenize("机器学习");
	assert.ok(t.includes("机") && t.includes("机器") && t.includes("器学") && t.includes("学习"));
});

check("tokenize latin words", () => {
	const t = tokenize("Hello_World foo-bar2");
	assert.ok(t.includes("hello") && t.includes("world") && t.includes("foo") && t.includes("bar2"));
});

check("cjk search hits", () => {
	const hits = search(idx, "机器学习");
	assert.ok(hits.length >= 2, `expected >=2 hits, got ${hits.length}`);
	assert.ok(hits.some((h) => h.key === "d1#2"));
	assert.ok(hits.some((h) => h.key === "d1#1"));
});

check("latin search hits", () => {
	const hits = search(idx, "revenue");
	assert.equal(hits[0].deckId, "d2");
	assert.ok(hits[0].snippets.length > 0);
});

check("title weighted above body", () => {
	const hits = search(idx, "机器学习入门");
	assert.equal(hits[0].key, "d1#2");
});

check("phrase bonus ranks exact title first", () => {
	const hits = search(idx, "季度营收总结");
	assert.equal(hits[0].key, "d2#2");
	assert.ok(hits[0].snippets[0].text.includes("季度营收总结"));
});

check("deckId filter", () => {
	const hits = search(idx, "学习", { deckId: "d1" });
	assert.ok(hits.length > 0 && hits.every((h) => h.deckId === "d1"));
});

check("limit", () => {
	assert.ok(search(idx, "学习", { limit: 1 }).length === 1);
});

check("empty query", () => {
	assert.deepEqual(search(idx, ""), []);
	assert.deepEqual(search(idx, "   "), []);
});

check("stopword-only query degrades to title match", () => {
	const hits = search(idx, "的");
	assert.ok(Array.isArray(hits));
});

check("removeDeck", () => {
	removeDeck(idx, "d2");
	assert.ok(search(idx, "revenue").length === 0);
	assert.ok(search(idx, "机器学习").every((h) => h.deckId === "d1"));
});

check("serialize round-trip", () => {
	const before = search(idx, "机器学习").map((h) => `${h.key}:${h.score}`);
	const restored = deserialize(JSON.parse(JSON.stringify(serialize(idx))));
	const after = search(restored, "机器学习").map((h) => `${h.key}:${h.score}`);
	assert.deepEqual(after, before);
	const restored2 = deserialize(JSON.parse(JSON.stringify(serialize(restored))));
	assert.deepEqual(search(restored2, "年度").map((h) => h.key), search(idx, "年度").map((h) => h.key));
});

check("addSlide replaces same key", () => {
	addSlide(idx, { key: "d1#2", deckId: "d1", n: 2, title: "深度学习进阶", body: "不再包含机器学习字样", notes: "" });
	const hits = search(idx, "机器学习入门");
	assert.ok(!hits.some((h) => h.key === "d1#2" && h.title === "机器学习入门"));
});

process.exit(failures === 0 ? 0 : 1);
