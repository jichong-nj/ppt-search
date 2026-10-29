// lib/index.js — PPT-Search 插件入口（cordis host plugin）。
// 提供：/ppt-search/api HTTP 路由、/ppt-search/mcp MCP 端点、ppt_* 智能体工具。
import os from "node:os";
import path from "node:path";
import { createKb } from "./kb.js";
import { registerRoutes } from "./api.js";
import { registerMcp } from "./mcp.js";
import { registerTools } from "./tools.js";

const name = "ppt-search";
const inject = ["webServer", "tools"];

function apply(ctx) {
	const dataDir = process.env.PPT_SEARCH_DATA ?? path.join(os.homedir(), ".dsh", "data", "ppt-search");
	const kb = createKb({ dataDir });
	ctx.effect(() => {
		kb.init().then(
			() => ctx.logger?.info?.(`[ppt-search] ready: ${kb.list().total} decks, ${kb.indexDocCount} slides indexed (${dataDir})`),
			(error) => ctx.logger?.warn?.(`[ppt-search] init failed: ${error?.stack ?? error}`),
		);
		const offApi = registerRoutes(ctx, kb);
		const offMcp = registerMcp(ctx, kb);
		const offTools = registerTools(ctx, kb);
		return () => {
			offApi?.();
			offMcp?.();
			offTools?.();
		};
	}, "ppt-search: api + mcp + tools");
}

export { apply, inject, name };
