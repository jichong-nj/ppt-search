// test/test-tools.mjs — 用真实 dsh-tools defineTool 校验 lib/tools.js 的 schema DSL，
// 并对拍 dslToRaw/paramsToRaw 编译结果与 defineTool 内部编译的 raw JSON Schema。
import assert from "node:assert/strict";
import { pathToFileURL } from "node:url";
import path from "node:path";

const DSH_ROOT = process.env.DSH_ROOT ?? "/home/jichong/.nvm/versions/node/v24.15.0/lib/node_modules/@deepseek-ai/dsh";
const toolsMod = await import(pathToFileURL(path.join(DSH_ROOT, "node_modules/@deepseek-ai/dsh-tools/lib/index.js")).href);
const { defineTool } = toolsMod;
assert.ok(typeof defineTool === "function", "defineTool 应可从 dsh-tools 导入");

const { buildToolDefs } = await import("../lib/tools.js");
const defs = buildToolDefs();

/** 忽略键序的深比较。 */
function deepEqualJson(a, b, path = "$") {
	if (a === b) return;
	if (Array.isArray(a) || Array.isArray(b)) {
		assert.ok(Array.isArray(a) && Array.isArray(b), `${path}: 类型不一致`);
		assert.equal(a.length, b.length, `${path}: 数组长度不一致`);
		a.forEach((_, i) => deepEqualJson(a[i], b[i], `${path}[${i}]`));
		return;
	}
	if (a !== null && b !== null && typeof a === "object" && typeof b === "object") {
		const ka = Object.keys(a).sort();
		const kb = Object.keys(b).sort();
		assert.deepEqual(ka, kb, `${path}: 键集合不一致`);
		for (const k of ka) deepEqualJson(a[k], b[k], `${path}.${k}`);
		return;
	}
	assert.equal(a, b, `${path}: 值不一致`);
}

let passed = 0;
for (const def of defs) {
	// 1) DSL 必须被真实 defineTool 接受（内部会 assertSupportedJsonSchema）
	const tool = defineTool({
		name: def.name,
		description: def.description,
		parameters: def.parameters,
		output: {
			schema: def.outputSchema,
			render: (args, value) => [{ type: "text", text: def.render(args ?? {}, value) }],
		},
		execute: async () => ({}),
	});
	passed += 1;
	console.log(`OK defineTool 接受 ${def.name}`);

	// 2) 我们的 raw 编译结果必须与 defineTool 编译结果语义一致
	deepEqualJson(tool.parameters, def.parametersRaw, `${def.name}.parameters`);
	deepEqualJson(tool.output.schema, def.outputSchemaRaw, `${def.name}.output.schema`);
	passed += 1;
	console.log(`OK raw 对拍一致 ${def.name}`);

	// 3) render 契约：数组 of {type:"text",text}
	const rendered = tool.output.render({}, sampleValue(def.name));
	assert.ok(Array.isArray(rendered) && rendered.length > 0, `${def.name}: render 必须返回非空数组`);
	for (const block of rendered) {
		assert.equal(block.type, "text");
		assert.equal(typeof block.text, "string");
	}
	passed += 1;
	console.log(`OK render 契约 ${def.name}`);
}

// 4) 参数校验行为：必填缺失应抛 ToolArgsError
{
	const def = defs.find((d) => d.name === "ppt_slide_get");
	const tool = defineTool({
		name: def.name,
		description: def.description,
		parameters: def.parameters,
		output: { schema: def.outputSchema, render: () => [{ type: "text", text: "x" }] },
		execute: async (args) => args,
	});
	await assert.rejects(() => tool.execute({ deckId: "a" }), /invalid arguments/);
	const ok = await tool.execute({ deckId: "a", n: 3 });
	assert.deepEqual(ok, { deckId: "a", n: 3 });
	passed += 1;
	console.log("OK 必填参数校验（缺 n 拒绝 / 齐全通过）");
}

// 5) ppt_compose 嵌套 items 校验
{
	const def = defs.find((d) => d.name === "ppt_compose");
	const tool = defineTool({
		name: def.name,
		description: def.description,
		parameters: def.parameters,
		output: { schema: def.outputSchema, render: () => [{ type: "text", text: "x" }] },
		execute: async (args) => args,
	});
	await assert.rejects(() => tool.execute({}), /invalid arguments/);
	await assert.rejects(() => tool.execute({ items: [{ deckId: "a" }] }), /invalid arguments/);
	await assert.rejects(() => tool.execute({ items: [{ deckId: "a", n: 1, bogus: 1 }] }), /invalid arguments/);
	const ok = await tool.execute({
		items: [{ deckId: "a", n: 1, titleOverride: "T", replace: [{ find: "x", replace: "y" }] }],
		name: "out",
	});
	assert.equal(ok.items[0].n, 1);
	passed += 1;
	console.log("OK ppt_compose 嵌套参数校验");
}

function sampleValue(name) {
	switch (name) {
		case "ppt_search": return { total: 1, hits: [{ deckId: "d", n: 1, title: "t", score: 1.5, snippets: [{ field: "body", text: "x" }] }] };
		case "ppt_slide_get": return { deckId: "d", deckName: "n", n: 1, title: "t", body: "b", notes: "", renderUrl: "/r" };
		case "ppt_deck_list": return { total: 1, decks: [{ id: "d", name: "n", status: "ready", slideCount: 3, error: "" }] };
		case "ppt_compose": return { id: "c", name: "n.pptx", slideCount: 2, size: 100, downloadUrl: "/dl" };
		default: throw new Error(name);
	}
}

console.log(`\n全部通过：${passed} 项`);
