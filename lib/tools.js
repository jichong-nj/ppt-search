// lib/tools.js — 智能体工具定义（DSH ctx.tools 与 MCP tools/list 共用实现）。
//
// schema 以 dsh-tools 的 DSL 形式书写（parameters = property map；对象节点显式
// additionalProperties；属性用 required:true 声明必填），再由本文件的 dslToRaw /
// paramsToRaw 编译为 raw JSON Schema 供 ctx.tools.register 与 MCP inputSchema 使用。
// 编译语义与 dsh-tools 的 parameterSchemaSpecToJsonSchema / valueSchemaSpecToJsonSchema
// 一致（test/test-tools.mjs 用真实 defineTool 对拍验证）。

const RENDER_BASE = "/ppt-search/api/render";

const ANNOTATIONS = ["description", "title", "default", "examples"];

/** DSL value-schema 节点 → raw JSON Schema。 */
function dslToRaw(node) {
	const out = {};
	for (const k of ANNOTATIONS) if (node[k] !== undefined) out[k] = node[k];
	if (Object.hasOwn(node, "oneOf")) {
		out.oneOf = node.oneOf.map(dslToRaw);
		return out;
	}
	switch (node.type) {
		case "object": {
			out.type = "object";
			out.additionalProperties = node.additionalProperties;
			if (node.properties !== undefined) {
				out.properties = {};
				const required = [];
				for (const [key, child] of Object.entries(node.properties)) {
					if (child.required === true) required.push(key);
					out.properties[key] = dslToRaw(child);
				}
				if (required.length > 0) out.required = required;
			}
			return out;
		}
		case "array": {
			out.type = "array";
			if (node.items !== undefined) out.items = dslToRaw(node.items);
			return out;
		}
		default: {
			out.type = node.type;
			if (node.enum !== undefined) out.enum = node.enum;
			if (node.const !== undefined) out.const = node.const;
			return out;
		}
	}
}

/** DSL parameters（property map）→ raw JSON Schema（object 根）。 */
function paramsToRaw(map) {
	const out = { type: "object", properties: {} };
	const required = [];
	for (const [key, child] of Object.entries(map)) {
		if (child.required === true) required.push(key);
		out.properties[key] = dslToRaw(child);
	}
	if (required.length > 0) out.required = required;
	return out;
}

/** DSL 定义表（parameters 为 property map，outputSchema 为 value schema）。 */
const DEFS = [
	{
		name: "ppt_search",
		description:
			"在已上传的 PPT 知识库中做全文检索（支持中文，BM25 排序）。返回命中页面列表（deckId、页码 n、标题、得分、摘要片段、渲染模型 URL）。query 留空且指定 deckId 时返回该 deck 的页目录。",
		parameters: {
			query: { type: "string", description: "检索关键词或短语（中文/英文）。留空配合 deckId 可列出该 deck 全部页目录" },
			deckId: { type: "string", description: "可选：限定在某个 deck 内检索（来自 ppt_deck_list）" },
			limit: { type: "integer", description: "最多返回条数，默认 10，上限 50" },
		},
		outputSchema: {
			type: "object",
			additionalProperties: false,
			properties: {
				total: { type: "integer", required: true, description: "命中总数" },
				hits: {
					type: "array",
					required: true,
					items: {
						type: "object",
						additionalProperties: true,
						properties: {
							deckId: { type: "string", required: true },
							deckName: { type: "string" },
							n: { type: "integer", required: true },
							title: { type: "string", required: true },
							score: { type: "number", required: true },
							snippets: {
								type: "array",
								items: {
									type: "object",
									additionalProperties: false,
									properties: { field: { type: "string", required: true }, text: { type: "string", required: true } },
								},
							},
							renderUrl: { type: "string" },
						},
					},
				},
			},
		},
		async run(kb, args) {
			const query = typeof args.query === "string" ? args.query.trim() : "";
			const opts = { limit: Math.min(Number(args.limit) || 10, 50) };
			if (typeof args.deckId === "string" && args.deckId !== "") opts.deckId = args.deckId;
			if (query === "") {
				if (opts.deckId === undefined) throw new Error("query 与 deckId 至少提供一个");
				const { deck, slides } = await kb.detail(opts.deckId);
				const hits = slides.map((s) => ({
					deckId: deck.id,
					deckName: deck.name,
					n: s.n,
					title: s.title,
					score: 0,
					snippets: [{ field: "body", text: s.body.slice(0, 120) }],
					renderUrl: `${RENDER_BASE}/${deck.id}/${s.n}.json`,
				}));
				return { total: hits.length, hits: hits.slice(0, opts.limit) };
			}
			return kb.search(query, opts);
		},
		render(args, value) {
			if (value.hits.length === 0) return "没有命中。";
			const lines = value.hits.map((h) => `[${h.deckId} #${h.n}] ${h.title} (score ${h.score.toFixed(2)})`);
			for (const h of value.hits) {
				for (const s of h.snippets ?? []) lines.push(`    ${s.field}: ${s.text}`);
			}
			return `共 ${value.total} 条命中：\n${lines.join("\n")}`;
		},
	},
	{
		name: "ppt_slide_get",
		description: "读取知识库中某一页 PPT 的完整内容：标题、正文全文、演讲备注、渲染模型 URL。deckId/n 来自 ppt_search 或 ppt_deck_list。",
		parameters: {
			deckId: { type: "string", description: "deck id", required: true },
			n: { type: "integer", description: "页码（1 起）", required: true },
		},
		outputSchema: {
			type: "object",
			additionalProperties: false,
			properties: {
				deckId: { type: "string", required: true },
				deckName: { type: "string" },
				n: { type: "integer", required: true },
				title: { type: "string", required: true },
				body: { type: "string", required: true },
				notes: { type: "string" },
				renderUrl: { type: "string" },
			},
		},
		async run(kb, args) {
			const { deck, slides } = await kb.detail(args.deckId);
			const slide = slides.find((s) => s.n === Number(args.n));
			if (slide === undefined) throw new Error(`deck ${args.deckId} 没有第 ${args.n} 页（共 ${slides.length} 页）`);
			return {
				deckId: deck.id,
				deckName: deck.name,
				n: slide.n,
				title: slide.title,
				body: slide.body,
				notes: slide.notes ?? "",
				renderUrl: `${RENDER_BASE}/${deck.id}/${slide.n}.json`,
			};
		},
		render(args, value) {
			const parts = [`# ${value.deckName} 第 ${value.n} 页`, `标题: ${value.title}`, `正文:\n${value.body}`];
			if (value.notes !== "") parts.push(`备注:\n${value.notes}`);
			return parts.join("\n");
		},
	},
	{
		name: "ppt_deck_list",
		description: "列出 PPT 知识库中全部 deck（文件）：id、文件名、状态（pending/processing/ready/error）、页数、错误信息。检索前可先用它了解库结构。",
		parameters: {},
		outputSchema: {
			type: "object",
			additionalProperties: false,
			properties: {
				total: { type: "integer", required: true },
				decks: {
					type: "array",
					required: true,
					items: {
						type: "object",
						additionalProperties: false,
						properties: {
							id: { type: "string", required: true },
							name: { type: "string", required: true },
							status: { type: "string", required: true },
							slideCount: { type: "integer" },
							error: { type: "string" },
						},
					},
				},
			},
		},
		async run(kb) {
			const { decks } = kb.list();
			return {
				total: decks.length,
				decks: decks.map((d) => ({ id: d.id, name: d.name, status: d.status, slideCount: d.slideCount ?? 0, error: d.error ?? "" })),
			};
		},
		render(args, value) {
			if (value.decks.length === 0) return "知识库为空。";
			return value.decks.map((d) => `- ${d.id} | ${d.name} | ${d.status} | ${d.slideCount} 页${d.error ? ` | 错误: ${d.error}` : ""}`).join("\n");
		},
	},
	{
		name: "ppt_compose",
		description:
			"按内容提纲从知识库挑选 PPT 页面合成新的 pptx 文件。items 按顺序给出每页 {deckId, n}，可选 titleOverride 改标题、replace 数组做该页文字查找替换（[{find,replace}]）。返回新文件 id 与下载地址。",
		parameters: {
			items: {
				type: "array",
				required: true,
				description: "提纲：按顺序的页面列表",
				items: {
					type: "object",
					additionalProperties: false,
					properties: {
						deckId: { type: "string", required: true },
						n: { type: "integer", required: true },
						titleOverride: { type: "string", description: "可选：覆盖该页标题" },
						replace: {
							type: "array",
							description: "可选：该页文字替换",
							items: {
								type: "object",
								additionalProperties: false,
								properties: { find: { type: "string", required: true }, replace: { type: "string", required: true } },
							},
						},
					},
				},
			},
			name: { type: "string", description: "可选：输出文件名（自动补 .pptx）" },
		},
		outputSchema: {
			type: "object",
			additionalProperties: false,
			properties: {
				id: { type: "string", required: true },
				name: { type: "string", required: true },
				slideCount: { type: "integer", required: true },
				size: { type: "integer" },
				downloadUrl: { type: "string" },
			},
		},
		async run(kb, args) {
			return await kb.compose({ items: args.items, name: args.name });
		},
		render(args, value) {
			return `已合成 ${value.name}（${value.slideCount} 页，${value.size} 字节），下载: ${value.downloadUrl}`;
		},
	},
];

/**
 * 构建工具定义表。
 * @returns {Array<{name,description,parameters,outputSchema,parametersRaw,outputSchemaRaw,run,render}>}
 *   parameters/outputSchema 为 DSL 形式（可用 dsh-tools defineTool 校验），
 *   parametersRaw/outputSchemaRaw 为 raw JSON Schema（供 ctx.tools.register 与 MCP）。
 */
export function buildToolDefs() {
	return DEFS.map((def) => ({
		...def,
		parametersRaw: paramsToRaw(def.parameters),
		outputSchemaRaw: dslToRaw(def.outputSchema),
	}));
}

/** 注册到 DSH 工具系统。@returns 取消注册函数 */
export function registerTools(ctx, kb) {
	const disposers = [];
	for (const def of buildToolDefs()) {
		const dispose = ctx.tools.register({
			name: def.name,
			description: def.description,
			parameters: def.parametersRaw,
			output: {
				schema: def.outputSchemaRaw,
				render: (args, value) => [{ type: "text", text: def.render(args ?? {}, value) }],
			},
			execute: async (args) => await def.run(kb, args ?? {}),
		});
		disposers.push(dispose);
	}
	return () => {
		for (const d of disposers) d?.();
	};
}
