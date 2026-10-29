// lib/api.js — HTTP 路由：POST /ppt-search/api/<endpoint>，GET render/media/download。
import { createReadStream } from "node:fs";
import { createWriteStream } from "node:fs";
import { rm } from "node:fs/promises";

const JSON_LIMIT = 2 * 1024 * 1024;
const UPLOAD_LIMIT = 500 * 1024 * 1024;

const MEDIA_TYPES = {
	png: "image/png",
	jpg: "image/jpeg",
	jpeg: "image/jpeg",
	gif: "image/gif",
	webp: "image/webp",
	svg: "image/svg+xml",
	bmp: "image/bmp",
	tiff: "image/tiff",
};

function writeJson(res, code, value) {
	const body = JSON.stringify(value);
	res.writeHead(code, { "content-type": "application/json; charset=utf-8", "content-length": Buffer.byteLength(body) });
	res.end(body);
}

function writeError(res, code, kind, message) {
	writeJson(res, code, { ok: false, error: { code: kind, message: String(message) } });
}

async function readJsonBody(req) {
	const chunks = [];
	let size = 0;
	for await (const chunk of req) {
		size += chunk.length;
		if (size > JSON_LIMIT) throw new RangeError("request body too large");
		chunks.push(chunk);
	}
	if (size === 0) return {};
	return JSON.parse(Buffer.concat(chunks).toString("utf8"));
}

/** 流式接收上传到目标文件；超限则销毁请求并抛错。 */
async function readUpload(req, destPath) {
	return await new Promise((resolve, reject) => {
		const out = createWriteStream(destPath);
		let size = 0;
		let failed = null;
		out.on("error", (e) => {
			failed = e;
			req.destroy();
			reject(e);
		});
		req.on("data", (chunk) => {
			if (failed !== null) return;
			size += chunk.length;
			if (size > UPLOAD_LIMIT) {
				failed = new RangeError("upload too large (limit 500MB)");
				req.destroy();
				out.destroy();
				reject(failed);
				return;
			}
			out.write(chunk);
		});
		req.on("error", (e) => {
			if (failed === null) reject(e);
		});
		req.on("end", () => {
			if (failed !== null) return;
			out.end(() => {
				if (failed === null) resolve(size);
			});
		});
	});
}

function disposition(filename) {
	const ascii = filename.replace(/[^\x20-\x7e]/g, "_").replace(/"/g, "'");
	return `attachment; filename="${ascii}"; filename*=UTF-8''${encodeURIComponent(filename)}`;
}

const ENDPOINTS = {
	async list(kb) {
		return kb.list();
	},
	async detail(kb, payload) {
		return await kb.detail(payload?.deckId);
	},
	async search(kb, payload) {
		if (typeof payload?.query !== "string" || payload.query.trim() === "") throw new TypeError("query 必须是非空字符串");
		const opts = {};
		if (typeof payload.deckId === "string" && payload.deckId !== "") opts.deckId = payload.deckId;
		if (Number.isInteger(payload.limit)) opts.limit = Math.min(Math.max(payload.limit, 1), 100);
		return kb.search(payload.query, opts);
	},
	async delete(kb, payload) {
		return await kb.remove(payload?.deckId);
	},
	async retry(kb, payload) {
		return { deck: await kb.retry(payload?.deckId) };
	},
	async compose(kb, payload) {
		return await kb.compose({ items: payload?.items, name: payload?.name });
	},
	async compositions(kb) {
		return await kb.compositions();
	},
};

function streamFile(res, file, contentType, size, extraHeaders = {}) {
	const headers = { "content-type": contentType, ...size !== undefined ? { "content-length": String(size) } : {}, ...extraHeaders };
	res.writeHead(200, headers);
	createReadStream(file).pipe(res);
}

/**
 * 注册 /ppt-search/api 前缀路由。
 * @returns 取消注册函数（若宿主提供）。
 */
export function registerRoutes(ctx, kb) {
	const handler = async (req, res) => {
		const url = new URL(req.url ?? "/", "http://localhost");
		const route = url.pathname.replace(/^\/ppt-search\/api\/?/, "");
		try {
			// GET 渲染模型 JSON
			const renderMatch = /^render\/([^/]+)\/(\d+)\.json$/.exec(route);
			if (renderMatch !== null) {
				if (req.method !== "GET" && req.method !== "HEAD") return writeError(res, 405, "method-not-allowed", "render requires GET");
				let file;
				try {
					file = await kb.renderFile(decodeURIComponent(renderMatch[1]), Number(renderMatch[2]));
				} catch {
					return writeError(res, 404, "not-found", "render model not found");
				}
				return streamFile(res, file, "application/json; charset=utf-8", undefined, { "cache-control": "no-cache" });
			}
			// GET 媒体文件（图片）
			const mediaMatch = /^media\/([^/]+)\/([^/]+)$/.exec(route);
			if (mediaMatch !== null) {
				if (req.method !== "GET" && req.method !== "HEAD") return writeError(res, 405, "method-not-allowed", "media requires GET");
				let file;
				try {
					file = await kb.mediaFile(decodeURIComponent(mediaMatch[1]), decodeURIComponent(mediaMatch[2]));
				} catch {
					return writeError(res, 404, "not-found", "media not found");
				}
				return streamFile(res, file, MEDIA_TYPES[mediaMatch[2].split(".").pop()?.toLowerCase()] ?? "application/octet-stream", undefined, {
					"cache-control": "public, max-age=31536000, immutable",
				});
			}
			// GET 下载（合成产物或 deck 源文件）
			const downloadMatch = /^download\/([^/]+)$/.exec(route);
			if (downloadMatch !== null) {
				if (req.method !== "GET" && req.method !== "HEAD") return writeError(res, 405, "method-not-allowed", "download requires GET");
				let info;
				try {
					info = await kb.downloadFile(decodeURIComponent(downloadMatch[1]));
				} catch {
					return writeError(res, 404, "not-found", "composition or deck not found");
				}
				return streamFile(res, info.file, "application/vnd.openxmlformats-officedocument.presentationml.presentation", info.size, {
					"content-disposition": disposition(info.name),
				});
			}
			// POST 上传
			if (route === "upload") {
				if (req.method !== "POST") return writeError(res, 405, "method-not-allowed", "upload requires POST");
				const name = url.searchParams.get("name") ?? "untitled.pptx";
				const dir = await kb.store.tmpDir();
				const tmpPath = `${dir}/upload.pptx`;
				let size;
				try {
					size = await readUpload(req, tmpPath);
				} catch (e) {
					await rm(dir, { recursive: true, force: true }).catch(() => undefined);
					if (e instanceof RangeError) return writeError(res, 413, "payload-too-large", e.message);
					throw e;
				}
				if (size < 4) {
					await rm(dir, { recursive: true, force: true }).catch(() => undefined);
					return writeError(res, 400, "bad-request", "empty upload");
				}
				// zip magic: PK\x03\x04
				const { open } = await import("node:fs/promises");
				const fh = await open(tmpPath, "r");
				const head = Buffer.alloc(4);
				await fh.read(head, 0, 4, 0);
				await fh.close();
				if (!(head[0] === 0x50 && head[1] === 0x4b && head[2] === 0x03 && head[3] === 0x04)) {
					await rm(dir, { recursive: true, force: true }).catch(() => undefined);
					return writeError(res, 400, "bad-request", "文件不是有效的 pptx（zip）格式");
				}
				const deck = await kb.addUpload({ name, tmpPath });
				return writeJson(res, 200, { ok: true, value: { deck } });
			}
			// POST JSON 端点
			if (req.method !== "POST") return writeError(res, 405, "method-not-allowed", "requires POST");
			const endpoint = ENDPOINTS[route];
			if (endpoint === undefined) return writeError(res, 404, "not-found", `unknown endpoint: ${route}`);
			let payload;
			try {
				payload = await readJsonBody(req);
			} catch (e) {
				if (e instanceof RangeError) return writeError(res, 413, "payload-too-large", e.message);
				return writeError(res, 400, "bad-request", `invalid JSON body: ${e.message}`);
			}
			const value = await endpoint(kb, payload);
			return writeJson(res, 200, { ok: true, value });
		} catch (error) {
			ctx.logger?.warn?.(`[ppt-search] api error: ${error?.stack ?? error}`);
			if (!res.headersSent) return writeError(res, 500, "internal-error", error?.message ?? String(error));
			res.destroy();
		}
	};

	return ctx.webServer.register({ kind: "prefix", path: "/ppt-search/api", handler });
}
