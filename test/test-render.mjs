// test/test-render.mjs — 渲染模型冒烟测试：rich.pptx / deck-a → render JSON
import assert from "node:assert/strict";
import { mkdtemp, readFile, readdir } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { buildRenderModels } from "../lib/pptx/renderModel.js";

const here = path.dirname(fileURLToPath(import.meta.url));
let passed = 0;
const ok = (name, fn) => {
	try {
		fn();
		passed += 1;
		console.log("  ok -", name);
	} catch (e) {
		console.error("  FAIL -", name, "\n   ", e.message);
		process.exitCode = 1;
	}
};

const tmp = await mkdtemp(path.join(os.tmpdir(), "ps-render-"));
const rich = await buildRenderModels(path.join(here, "fixtures", "rich.pptx"), tmp);
const load = async (n) => JSON.parse(await readFile(path.join(tmp, "render", `${n}.json`), "utf8"));

console.log("rich.pptx:");
ok("4 页", () => assert.equal(rich.count, 4));
const s1 = await load(1);
const s2 = await load(2);
const s3 = await load(3);
const s4 = await load(4);

ok("slide 尺寸 16:9", () => {
	assert.ok(s1.w > 12100000 && s1.w < 12300000, "w=" + s1.w);
	assert.equal(s1.h, 6858000);
});
ok("背景为实心色", () => assert.equal(s1.bg.t, "solid"));
ok("所有元素坐标有限", () => {
	for (const s of [s1, s2, s3, s4]) {
		for (const el of s.els) {
			assert.ok(Number.isFinite(el.x) && Number.isFinite(el.y) && Number.isFinite(el.cx) && Number.isFinite(el.cy), `${el.k} 坐标异常`);
		}
	}
});

const title = s1.els.find((e) => e.k === "sp" && e.text?.paras.some((p) => p.runs.some((r) => r.t.includes("渲染模型测试"))));
ok("标题文本存在", () => assert.ok(title));
const body = s1.els.find((e) => e.k === "sp" && e.text?.paras.some((p) => p.runs.some((r) => r.t.includes("第二条"))));
ok("多 run 样式（粗体红色）", () => {
	const p = body.text.paras.find((pp) => pp.runs.some((r) => r.t.includes("第二条")));
	const r = p.runs.find((r) => r.t.includes("第二条"));
	assert.equal(r.b, 1);
	assert.equal(r.c, "#c00000");
	const it = p.runs.find((r) => r.t.includes("斜体"));
	assert.equal(it.i, 1);
});
ok("右对齐段落", () => {
	assert.ok(body.text.paras.some((p) => p.algn === "r" && p.runs.some((r) => r.t.includes("右对齐"))));
});

const charts2 = s2.els.filter((e) => e.k === "chart");
ok("slide2 两个图表", () => assert.equal(charts2.length, 2));
ok("柱状图数据", () => {
	const bar = charts2.find((c) => c.chart.type === "bar");
	assert.equal(bar.chart.dir, "col");
	assert.equal(bar.chart.grouping, "clustered");
	assert.deepEqual(bar.chart.cats, ["一月", "二月", "三月"]);
	assert.equal(bar.chart.series.length, 2);
	assert.deepEqual(bar.chart.series[0].vals, [1.2, 3.4, 2.8]);
	assert.ok(bar.chart.series[0].color.startsWith("#"));
});
ok("折线图数据", () => {
	const line = charts2.find((c) => c.chart.type === "line");
	assert.deepEqual(line.chart.cats, ["Q1", "Q2", "Q3", "Q4"]);
	assert.deepEqual(line.chart.series[0].vals, [10, 22, 15, 30]);
});

const charts3 = s3.els.filter((e) => e.k === "chart");
ok("slide3 饼图+环形图", () => {
	assert.ok(charts3.some((c) => c.chart.type === "pie" && c.chart.series[0].vals.join() === "45,30,25"));
	assert.ok(charts3.some((c) => c.chart.type === "doughnut"));
});
const tbl = s3.els.find((e) => e.k === "tbl");
ok("表格 3x3", () => {
	assert.equal(tbl.cols.length, 3);
	assert.equal(tbl.rows.length, 3);
	assert.equal(tbl.cells.length, 9);
	const c00 = tbl.cells.find((c) => c.r === 0 && c.c === 0);
	assert.ok(c00.text.paras[0].runs[0].t === "地区");
});

const pic = s4.els.find((e) => e.k === "pic");
ok("图片元素 + 媒体提取", () => {
	assert.ok(pic && /\.(png|jpeg|jpg|gif)$/i.test(pic.src));
	assert.ok(rich.media.includes(pic.src));
});
const geoms = s4.els.filter((e) => e.k === "sp").map((e) => e.geom);
ok("形状几何（roundRect/ellipse/rightArrow/triangle）", () => {
	assert.ok(geoms.includes("roundRect"), "roundRect: " + geoms.join(","));
	assert.ok(geoms.includes("ellipse"));
	assert.ok(geoms.includes("rightArrow"));
	assert.ok(geoms.includes("triangle"), "组合内三角形被展平");
});
const grad = s4.els.find((e) => e.k === "sp" && e.fill?.t === "grad");
ok("渐变填充", () => assert.ok(grad && grad.fill.stops.length >= 2));
const tbx = s4.els.find((e) => e.k === "sp" && e.text?.paras.some((p) => p.runs.some((r) => r.t.includes("混排换行"))));
ok("文本框文本", () => assert.ok(tbx));

const media = await readdir(path.join(tmp, "media"));
ok("media 目录含图片文件", () => assert.ok(media.length >= 1));

console.log("deck-a（旧夹具）:");
const tmp2 = await mkdtemp(path.join(os.tmpdir(), "ps-render2-"));
const ra = await buildRenderModels(path.join(here, "fixtures", "deck-a", "source.pptx"), tmp2);
ok("deck-a 3 页", () => assert.equal(ra.count, 3));
const a1 = await load2(1);
async function load2(n) {
	return JSON.parse(await readFile(path.join(tmp2, "render", `${n}.json`), "utf8"));
}
ok("deck-a 第 1 页有文本", () => {
	assert.ok(a1.els.some((e) => e.text));
});

console.log(`\n${passed} 项通过${process.exitCode ? "（有失败）" : "，全部通过"}`);
