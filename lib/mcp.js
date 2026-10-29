// lib/mcp.js — 无状态 MCP（JSON-RPC 2.0 over HTTP POST）：/ppt-search/mcp
import { buildToolDefs } from "./tools.js";

const SERVER_INFO = { name: "ppt-search", version: "0.1.0" };
const DEFAULT_PROTOCOL = "2025-03-26";

function jsonRpc(res, code, payload) {
	const body = JSON.stringify(payload);
	res.writeHead(code, { "content-type": "application/json; charset=utf-8", "content-length": Buffer.byteLength(body) });
	res.end(body);
}

function rpcError(id, code, message) {
	return { jsonrpc: "2.0", id: id ?? null, error: { code, message } };
}

function rpcResult(id, result) {
	return { jsonrpc: "2.0", id: id ?? null, result };
}

async function readBody(req, limit = 4 * 1024 * 1024) {
	const chunks = [];
	let size = 0;
	for await (const chunk of req) {
		size += chunk.length;
		if (size > limit) throw new RangeError("body too large");
		chunks.push(chunk);
	}
	return Buffer.concat(chunks).toString("utf8");
}

async function handleRequest(kb, defs, msg) {
	const id = msg?.id;
	const method = msg?.method;
	if (typeof method !== "string") return { reply: rpcError(id, -32600, "invalid request"), status: 400 };

	if (method === "initialize") {
		const params = msg.params ?? {};
		return {
			reply: rpcResult(id, {
				protocolVersion: typeof params.protocolVersion === "string" ? params.protocolVersion : DEFAULT_PROTOCOL,
				capabilities: { tools: {} },
				serverInfo: SERVER_INFO,
			}),
			status: 200,
		};
	}
	if (method.startsWith("notifications/")) return { reply: null, status: 202 };
	if (method === "ping") return { reply: rpcResult(id, {}), status: 200 };

	if (method === "tools/list") {
		const tools = defs.map((d) => ({ name: d.name, description: d.description, inputSchema: d.parametersRaw }));
		return { reply: rpcResult(id, { tools }), status: 200 };
	}

	if (method === "tools/call") {
		const name = msg.params?.name;
		const args = msg.params?.arguments ?? {};
		const def = defs.find((d) => d.name === name);
		if (def === undefined) return { reply: rpcError(id, -32602, `unknown tool: ${name}`), status: 200 };
		try {
			const value = await def.run(kb, args);
			return {
				reply: rpcResult(id, {
					content: [{ type: "text", text: def.render(args, value) }],
					structuredContent: value,
					isError: false,
				}),
				status: 200,
			};
		} catch (error) {
			return {
				reply: rpcResult(id, {
					content: [{ type: "text", text: `工具执行失败: ${error?.message ?? error}` }],
					isError: true,
				}),
				status: 200,
			};
		}
	}

	return { reply: rpcError(id, -32601, `method not found: ${method}`), status: 200 };
}

/**
 * 注册 /ppt-search/mcp 端点。
 * @returns 取消注册函数
 */
export function registerMcp(ctx, kb) {
	const defs = buildToolDefs();
	const handler = async (req, res) => {
		if (req.method !== "POST") {
			res.writeHead(405, { allow: "POST" });
			res.end();
			return;
		}
		try {
			const text = await readBody(req);
			let msg;
			try {
				msg = JSON.parse(text);
			} catch {
				return jsonRpc(res, 400, rpcError(null, -32700, "parse error"));
			}
			if (Array.isArray(msg)) {
				// 批量：逐条处理，仅返回有回复的
				const replies = [];
				for (const m of msg) {
					const { reply } = await handleRequest(kb, defs, m);
					if (reply !== null) replies.push(reply);
				}
				return jsonRpc(res, 200, replies);
			}
			const { reply, status } = await handleRequest(kb, defs, msg);
			if (reply === null) {
				res.writeHead(status);
				res.end();
				return;
			}
			return jsonRpc(res, status, reply);
		} catch (error) {
			ctx.logger?.warn?.(`[ppt-search] mcp error: ${error?.message ?? error}`);
			if (!res.headersSent) return jsonRpc(res, 500, rpcError(null, -32603, String(error?.message ?? error)));
			res.destroy();
		}
	};

	return ctx.webServer.register({ kind: "prefix", path: "/ppt-search/mcp", handler });
}
