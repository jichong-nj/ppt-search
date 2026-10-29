// Persistence layer: directory layout, atomic writes, id hygiene.

import { randomBytes } from "node:crypto";
import { mkdir, mkdtemp, readdir, readFile, rename, rm, stat, writeFile } from "node:fs/promises";
import path from "node:path";

const ID_PATTERN = /^[a-z0-9][a-z0-9-]{0,64}$/;

export function assertId(value, what = "id") {
	if (typeof value !== "string" || !ID_PATTERN.test(value)) {
		throw new Error(`invalid ${what}: ${JSON.stringify(value)}`);
	}
	return value;
}

async function readJson(filePath, fallback) {
	try {
		return JSON.parse(await readFile(filePath, "utf8"));
	} catch {
		return fallback;
	}
}

async function writeJsonAtomic(filePath, value) {
	await mkdir(path.dirname(filePath), { recursive: true });
	const tmp = `${filePath}.${randomBytes(4).toString("hex")}.tmp`;
	await writeFile(tmp, JSON.stringify(value, null, 1), "utf8");
	await rename(tmp, filePath);
}

/**
 * @param {string} dataDir absolute data root
 */
export function createStore(dataDir) {
	const decksRoot = path.join(dataDir, "decks");
	const compositionsRoot = path.join(dataDir, "compositions");
	const tmpRoot = path.join(dataDir, "tmp");

	function deckDir(deckId) {
		return path.join(decksRoot, assertId(deckId, "deckId"));
	}

	return {
		dataDir,
		decksRoot,
		compositionsRoot,
		tmpRoot,

		async ensureDirs() {
			await mkdir(decksRoot, { recursive: true });
			await mkdir(compositionsRoot, { recursive: true });
			await mkdir(tmpRoot, { recursive: true });
		},

		async loadLibrary() {
			const lib = await readJson(path.join(dataDir, "library.json"), null);
			if (lib === null || typeof lib !== "object" || lib.decks === null || typeof lib.decks !== "object") {
				return { decks: {} };
			}
			return lib;
		},

		saveLibrary(lib) {
			return writeJsonAtomic(path.join(dataDir, "library.json"), lib);
		},

		async readDeckMeta(deckId) {
			return await readJson(path.join(deckDir(deckId), "meta.json"), null);
		},

		writeDeckMeta(deck) {
			return writeJsonAtomic(path.join(deckDir(deck.id), "meta.json"), deck);
		},

		async readDeckSlides(deckId) {
			const slides = await readJson(path.join(deckDir(deckId), "slides.json"), null);
			return Array.isArray(slides) ? slides : null;
		},

		writeDeckSlides(deckId, slides) {
			return writeJsonAtomic(path.join(deckDir(deckId), "slides.json"), slides);
		},

		deckDir,

		sourcePath(deckId) {
			return path.join(deckDir(deckId), "source.pptx");
		},

		renderPath(deckId, n) {
			const page = Number(n);
			if (!Number.isInteger(page) || page < 1 || page > 10_000) throw new Error(`invalid slide number: ${n}`);
			return path.join(deckDir(deckId), "render", `${page}.json`);
		},

		mediaPath(deckId, name) {
			if (typeof name !== "string" || !/^[\w][\w.\-+]*$/.test(name) || name.includes("..")) {
				throw new Error(`invalid media name: ${JSON.stringify(name)}`);
			}
			return path.join(deckDir(deckId), "media", name);
		},

		compositionPath(id) {
			return path.join(compositionsRoot, `${assertId(id, "compositionId")}.pptx`);
		},

		compositionMetaPath(id) {
			return path.join(compositionsRoot, `${assertId(id, "compositionId")}.json`);
		},

		async writeCompositionMeta(id, meta) {
			return await writeJsonAtomic(this.compositionMetaPath(id), meta);
		},

		async readCompositionMeta(id) {
			return await readJson(this.compositionMetaPath(id), {});
		},

		async tmpDir() {
			return await mkdtemp(path.join(tmpRoot, "job-"));
		},

		genId(prefix) {
			return `${prefix}-${Date.now().toString(36)}-${randomBytes(2).toString("hex")}`;
		},

		removeDeckFiles(deckId) {
			return rm(deckDir(deckId), { recursive: true, force: true });
		},

		async removeComposition(id) {
			await rm(this.compositionPath(id), { force: true });
			await rm(this.compositionMetaPath(id), { force: true });
		},

		async listCompositions() {
			let names = [];
			try {
				names = await readdir(compositionsRoot);
			} catch {
				return [];
			}
			const items = [];
			for (const name of names) {
				if (!name.endsWith(".pptx")) continue;
				const id = name.slice(0, -".pptx".length);
				if (!ID_PATTERN.test(id)) continue;
				const file = path.join(compositionsRoot, name);
				try {
					const info = await stat(file);
					const meta = await readJson(this.compositionMetaPath(id), {});
					items.push({ id, file, size: info.size, mtime: info.mtimeMs, name: meta.name ?? `${id}.pptx`, slideCount: meta.slideCount });
				} catch {
					// ignore unreadable entries
				}
			}
			items.sort((a, b) => b.mtime - a.mtime);
			return items;
		},
	};
}
