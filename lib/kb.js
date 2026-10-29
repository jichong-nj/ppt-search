// lib/kb.js — 知识库服务：deck 生命周期、串行摄取队列、检索索引、合成。
import { mkdir, rename, stat } from "node:fs/promises";
import { createStore, assertId } from "./store.js";
import { createIndex, addSlide, removeDeck, search as searchIndex } from "./search.js";
import { parsePptx } from "./pptx/parse.js";
import { composePptx } from "./pptx/compose.js";
import { buildRenderModels } from "./pptx/renderModel.js";

function ensurePptxName(name) {
	return /\.pptx$/i.test(name) ? name : `${name}.pptx`;
}

/**
 * @param {{dataDir: string}} opts
 */
export function createKb({ dataDir }) {
	const store = createStore(dataDir);
	/** @type {{decks: Record<string, any>}} */
	let library = { decks: {} };
	let index = createIndex();
	let tail = Promise.resolve();

	function persist() {
		return store.saveLibrary(library);
	}

	function enqueue(task) {
		const run = tail.then(task, task);
		tail = run.then(
			() => undefined,
			() => undefined,
		);
		return run;
	}

	async function indexDeck(deckId) {
		const slides = (await store.readDeckSlides(deckId)) ?? [];
		for (const s of slides) {
			addSlide(index, { key: `${deckId}#${s.n}`, deckId, n: s.n, title: s.title ?? "", body: s.body ?? "", notes: s.notes ?? "" });
		}
	}

	async function rebuildIndex() {
		index = createIndex();
		for (const deck of Object.values(library.decks)) {
			if (deck.status !== "ready") continue;
			await indexDeck(deck.id);
		}
	}

	/** 启动：建目录、载 library、恢复中断状态、重建索引。 */
	async function init() {
		await store.ensureDirs();
		library = await store.loadLibrary();
		let dirty = false;
		for (const deck of Object.values(library.decks)) {
			if (deck.status === "processing" || deck.status === "pending") {
				deck.status = "error";
				deck.error = "处理被重启中断，请重试";
				dirty = true;
			}
		}
		if (dirty) await persist();
		await rebuildIndex();
	}

	async function ingest(deckId) {
		const deck = library.decks[deckId];
		if (deck === undefined) return;
		deck.status = "processing";
		deck.error = undefined;
		await persist();
		try {
			const src = store.sourcePath(deckId);
			const parsed = await parsePptx(src);
			await store.writeDeckSlides(deckId, parsed.slides);
			deck.slideSize = parsed.slideSize;
			deck.slideCount = parsed.slideCount;
			deck.slides = parsed.slides.map((s) => ({ n: s.n, title: s.title, charCount: `${s.title}${s.body}${s.notes}`.length }));
			await persist();
			const rendered = await buildRenderModels(src, store.deckDir(deckId));
			deck.renderCount = rendered.count;
			removeDeck(index, deckId);
			await indexDeck(deckId);
			deck.status = "ready";
			deck.readyAt = Date.now();
		} catch (error) {
			deck.status = "error";
			deck.error = String(error?.message ?? error);
		}
		await persist();
	}

	/**
	 * 接收已落盘的上传文件（tmpPath 在 store.tmpRoot 内），登记 deck 并排队摄取。
	 * @param {{name: string, tmpPath: string}} opts
	 */
	async function addUpload({ name, tmpPath }) {
		const deckId = store.genId("deck");
		await mkdir(store.deckDir(deckId), { recursive: true });
		await rename(tmpPath, store.sourcePath(deckId));
		const deck = { id: deckId, name: name || "untitled.pptx", status: "pending", uploadedAt: Date.now(), slideCount: 0 };
		library.decks[deckId] = deck;
		await persist();
		enqueue(() => ingest(deckId));
		return deck;
	}

	async function retry(deckId) {
		const deck = library.decks[assertId(deckId, "deckId")];
		if (deck === undefined) throw new Error(`no such deck: ${deckId}`);
		if (deck.status === "ready" || deck.status === "processing") return deck;
		deck.status = "pending";
		deck.error = undefined;
		await persist();
		enqueue(() => ingest(deckId));
		return deck;
	}

	async function remove(deckId) {
		const id = assertId(deckId, "deckId");
		if (library.decks[id] === undefined) throw new Error(`no such deck: ${id}`);
		delete library.decks[id];
		removeDeck(index, id);
		await persist();
		await store.removeDeckFiles(id);
		return { removed: id };
	}

	function list() {
		const decks = Object.values(library.decks).sort((a, b) => (b.uploadedAt ?? 0) - (a.uploadedAt ?? 0));
		return { decks, total: decks.length };
	}

	async function detail(deckId) {
		const id = assertId(deckId, "deckId");
		const deck = library.decks[id];
		if (deck === undefined) throw new Error(`no such deck: ${id}`);
		const slides = (await store.readDeckSlides(id)) ?? [];
		return { deck, slides };
	}

	function search(query, opts = {}) {
		const hits = searchIndex(index, query, opts).map((h) => ({
			...h,
			deckName: library.decks[h.deckId]?.name ?? h.deckId,
			renderUrl: `/ppt-search/api/render/${h.deckId}/${h.n}.json`,
		}));
		return { hits, total: hits.length };
	}

	/**
	 * @param {{items: Array<{deckId:string,n:number,titleOverride?:string,replace?:Array}>, name?: string}} opts
	 */
	async function compose({ items, name }) {
		if (!Array.isArray(items) || items.length === 0) throw new Error("items 不能为空");
		for (const it of items) {
			const deck = library.decks[assertId(it?.deckId, "deckId")];
			if (deck === undefined || deck.status !== "ready") throw new Error(`deck 不可用: ${it?.deckId}`);
			if (!Number.isInteger(it.n) || it.n < 1 || it.n > deck.slideCount) throw new Error(`页码越界: ${it?.deckId}#${it?.n}`);
		}
		const id = store.genId("comp");
		const outPath = store.compositionPath(id);
		const res = await composePptx({ items, outPath, decksRoot: store.decksRoot });
		const finalName = typeof name === "string" && name.trim() !== "" ? ensurePptxName(name.trim()) : `composition-${id}.pptx`;
		await store.writeCompositionMeta(id, { name: finalName, slideCount: res.slideCount, createdAt: Date.now() });
		return { id, name: finalName, slideCount: res.slideCount, size: res.bytes, downloadUrl: `/ppt-search/api/download/${id}` };
	}

	async function compositions() {
		const items = await store.listCompositions();
		return {
			items: items.map((it) => ({
				id: it.id,
				name: it.name,
				slideCount: it.slideCount,
				size: it.size,
				mtime: it.mtime,
				downloadUrl: `/ppt-search/api/download/${it.id}`,
			})),
		};
	}

	/** 合成产物或 deck 源文件的下载路径。 */
	async function downloadFile(id) {
		const comp = store.compositionPath(id);
		try {
			const st = await stat(comp);
			const meta = await store.readCompositionMeta?.(id);
			return { file: comp, name: meta?.name ?? `${id}.pptx`, size: st.size };
		} catch {
			/* fall through */
		}
		const deck = library.decks[assertId(id, "id")];
		if (deck === undefined) throw new Error(`no such composition or deck: ${id}`);
		const src = store.sourcePath(deck.id);
		const st = await stat(src);
		return { file: src, name: deck.name, size: st.size };
	}

	/** 旧 deck（无渲染模型）首次访问时懒生成；文件缺失即重建；同 deck 并发请求共享一次生成。 */
	const renderEnsuring = new Map();

	async function ensureRender(deckId, file) {
		try {
			await stat(file);
			return;
		} catch {
			/* 缺失，重建 */
		}
		const deck = library.decks[deckId];
		if (deck === undefined) throw new Error(`no such deck: ${deckId}`);
		let p = renderEnsuring.get(deckId);
		if (p === undefined) {
			p = (async () => {
				const rendered = await buildRenderModels(store.sourcePath(deckId), store.deckDir(deckId));
				deck.renderCount = rendered.count;
				await persist();
			})().finally(() => renderEnsuring.delete(deckId));
			renderEnsuring.set(deckId, p);
		}
		await p;
	}

	async function renderFile(deckId, n) {
		assertId(deckId, "deckId");
		const file = store.renderPath(deckId, n);
		await ensureRender(deckId, file);
		await stat(file);
		return file;
	}

	async function mediaFile(deckId, name) {
		assertId(deckId, "deckId");
		const file = store.mediaPath(deckId, name);
		await ensureRender(deckId, file);
		await stat(file);
		return file;
	}

	return {
		init,
		list,
		detail,
		search,
		addUpload,
		retry,
		remove,
		compose,
		compositions,
		downloadFile,
		renderFile,
		mediaFile,
		store,
		get indexDocCount() {
			return index.docs.size;
		},
	};
}
