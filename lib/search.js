// Full-text search: BM25 over title/body/notes with CJK-aware tokenization.
// Zero dependencies. Index is a plain in-memory structure that serializes to JSON.

const K1 = 1.2;
const B = 0.75;
const FIELD_WEIGHT = { title: 3, body: 1, notes: 0.5 };
const FIELDS = ["title", "body", "notes"];
const STOPWORDS = new Set(["的", "了", "是", "在", "和", "与", "或", "吗", "呢", "a", "an", "the", "of", "to", "in", "on", "and", "or", "is", "are", "for", "with"]);
const PHRASE_BONUS = 10;
const SNIPPET_WINDOW = 60;

function isCjk(cp) {
	return (
		(cp >= 0x4e00 && cp <= 0x9fff) || // CJK unified
		(cp >= 0x3400 && cp <= 0x4dbf) || // ext A
		(cp >= 0xf900 && cp <= 0xfaff) || // compatibility
		(cp >= 0x3040 && cp <= 0x30ff) // kana
	);
}

function isWordChar(cp) {
	return (cp >= 0x61 && cp <= 0x7a) || (cp >= 0x30 && cp <= 0x39) || cp >= 0x80 && !isCjk(cp);
}

/** Tokenize: latin/digit words (lowercased) + CJK unigrams & adjacent bigrams. */
export function tokenize(text) {
	const tokens = [];
	const s = String(text ?? "").toLowerCase();
	let word = "";
	let prevCjk = null;
	const flushWord = () => {
		if (word !== "") {
			tokens.push(word);
			word = "";
		}
	};
	for (const ch of s) {
		const cp = ch.codePointAt(0);
		if (isCjk(cp)) {
			flushWord();
			tokens.push(ch);
			if (prevCjk !== null) tokens.push(prevCjk + ch);
			prevCjk = ch;
		} else if (isWordChar(cp)) {
			if (prevCjk !== null) prevCjk = null;
			word += ch;
		} else {
			flushWord();
			prevCjk = null;
		}
	}
	flushWord();
	return tokens;
}

export function createIndex() {
	return {
		docs: new Map(), // key -> {key, deckId, n, title, body, notes}
		postings: { title: new Map(), body: new Map(), notes: new Map() }, // field -> term -> Map(key -> tf)
		docLen: { title: new Map(), body: new Map(), notes: new Map() }, // field -> key -> token count
		totalLen: { title: 0, body: 0, notes: 0 },
	};
}

function indexDoc(index, doc) {
	for (const field of FIELDS) {
		const tokens = tokenize(doc[field] ?? "");
		index.docLen[field].set(doc.key, tokens.length);
		index.totalLen[field] += tokens.length;
		const postings = index.postings[field];
		for (const term of tokens) {
			let list = postings.get(term);
			if (list === undefined) {
				list = new Map();
				postings.set(term, list);
			}
			list.set(doc.key, (list.get(doc.key) ?? 0) + 1);
		}
	}
}

function unindexDoc(index, key) {
	for (const field of FIELDS) {
		const len = index.docLen[field].get(key);
		if (len === undefined) continue;
		index.totalLen[field] -= len;
		index.docLen[field].delete(key);
		for (const [term, list] of index.postings[field]) {
			if (list.delete(key) && list.size === 0) index.postings[field].delete(term);
		}
	}
	index.docs.delete(key);
}

/** Insert or replace a slide document ({key, deckId, n, title, body, notes}). */
export function addSlide(index, doc) {
	if (index.docs.has(doc.key)) unindexDoc(index, doc.key);
	index.docs.set(doc.key, { key: doc.key, deckId: doc.deckId, n: doc.n, title: doc.title ?? "", body: doc.body ?? "", notes: doc.notes ?? "" });
	indexDoc(index, index.docs.get(doc.key));
}

export function removeDeck(index, deckId) {
	for (const key of [...index.docs.keys()]) {
		if (index.docs.get(key).deckId === deckId) unindexDoc(index, key);
	}
}

function avgdl(index, field) {
	const count = index.docLen[field].size;
	return count === 0 ? 1 : Math.max(1, index.totalLen[field] / count);
}

function fieldScore(index, field, term, key, tf) {
	const df = index.postings[field].get(term)?.size ?? 0;
	if (df === 0) return 0;
	const n = index.docLen[field].size;
	const idf = Math.log(1 + (n - df + 0.5) / (df + 0.5));
	const dl = index.docLen[field].get(key) ?? 0;
	const denom = tf + K1 * (1 - B + B * (dl / avgdl(index, field)));
	return idf * ((tf * (K1 + 1)) / denom);
}

function windowSnippet(text, at) {
	const chars = [...text];
	const start = Math.max(0, at - SNIPPET_WINDOW);
	const end = Math.min(chars.length, at + SNIPPET_WINDOW);
	const prefix = start > 0 ? "…" : "";
	const suffix = end < chars.length ? "…" : "";
	return prefix + chars.slice(start, end).join("").trim() + suffix;
}

function firstHit(text, terms) {
	const lower = String(text).toLowerCase();
	let best = -1;
	for (const term of terms) {
		const at = lower.indexOf(term);
		if (at >= 0 && (best < 0 || at < best)) best = at;
	}
	return best;
}

function buildSnippets(doc, terms, phrase) {
	const snippets = [];
	if (phrase !== "" && phrase.length >= 2) {
		for (const field of FIELDS) {
			const at = String(doc[field] ?? "").toLowerCase().indexOf(phrase.toLowerCase());
			if (at >= 0) {
				snippets.push({ field, text: windowSnippet(doc[field], at) });
				break;
			}
		}
	}
	for (const field of FIELDS) {
		if (snippets.length >= 3) break;
		const text = doc[field] ?? "";
		if (text === "") continue;
		const at = firstHit(text, terms);
		if (at < 0) continue;
		const snippet = windowSnippet(text, at);
		if (!snippets.some((s) => s.text === snippet)) snippets.push({ field, text: snippet });
	}
	return snippets.slice(0, 3);
}

/**
 * Search the index.
 * @param {ReturnType<typeof createIndex>} index
 * @param {string} query
 * @param {{deckId?: string, limit?: number}} [opts]
 * @returns {Array<{key, deckId, n, title, score, snippets}>}
 */
export function search(index, query, opts = {}) {
	const raw = String(query ?? "").trim();
	if (raw === "" || index.docs.size === 0) return [];
	const limit = Math.max(1, Math.min(100, opts.limit ?? 20));
	const deckFilter = typeof opts.deckId === "string" && opts.deckId !== "" ? opts.deckId : null;
	const tokens = tokenize(raw);
	const meaningful = tokens.filter((t) => !STOPWORDS.has(t));
	const scores = new Map();
	const bump = (key, delta) => scores.set(key, (scores.get(key) ?? 0) + delta);

	if (meaningful.length === 0) {
		// all-stopword query: degrade to title substring match
		const needle = raw.toLowerCase();
		for (const doc of index.docs.values()) {
			if (deckFilter !== null && doc.deckId !== deckFilter) continue;
			if (doc.title.toLowerCase().includes(needle)) bump(doc.key, 1);
		}
	} else {
		for (const term of meaningful) {
			for (const field of FIELDS) {
				const list = index.postings[field].get(term);
				if (list === undefined) continue;
				for (const [key, tf] of list) {
					const doc = index.docs.get(key);
					if (doc === undefined) continue;
					if (deckFilter !== null && doc.deckId !== deckFilter) continue;
					bump(key, FIELD_WEIGHT[field] * fieldScore(index, field, term, key, tf));
				}
			}
		}
		if (raw.length >= 2) {
			for (const [key, score] of [...scores]) {
				const doc = index.docs.get(key);
				if (doc === undefined) continue;
				if (FIELDS.some((f) => String(doc[f]).toLowerCase().includes(raw.toLowerCase()))) bump(key, PHRASE_BONUS);
			}
		}
	}

	const hits = [];
	for (const [key, score] of scores) {
		const doc = index.docs.get(key);
		if (doc === undefined) continue;
		hits.push({
			key,
			deckId: doc.deckId,
			n: doc.n,
			title: doc.title,
			score: Math.round(score * 1000) / 1000,
			snippets: buildSnippets(doc, meaningful.length > 0 ? meaningful : [raw.toLowerCase()], raw),
		});
	}
	hits.sort((a, b) => b.score - a.score || a.deckId.localeCompare(b.deckId) || a.n - b.n);
	return hits.slice(0, limit);
}

export function serialize(index) {
	const postings = {};
	for (const field of FIELDS) {
		postings[field] = {};
		for (const [term, list] of index.postings[field]) {
			postings[field][term] = [...list.entries()];
		}
	}
	return {
		version: 1,
		docs: [...index.docs.values()],
		postings,
		docLen: {
			title: [...index.docLen.title.entries()],
			body: [...index.docLen.body.entries()],
			notes: [...index.docLen.notes.entries()],
		},
		totalLen: { ...index.totalLen },
	};
}

export function deserialize(obj) {
	const index = createIndex();
	if (obj === null || typeof obj !== "object" || obj.version !== 1) return index;
	for (const doc of obj.docs ?? []) {
		if (typeof doc?.key !== "string") continue;
		index.docs.set(doc.key, { key: doc.key, deckId: doc.deckId, n: doc.n, title: doc.title ?? "", body: doc.body ?? "", notes: doc.notes ?? "" });
	}
	for (const field of FIELDS) {
		const src = obj.postings?.[field];
		if (src !== null && typeof src === "object") {
			for (const [term, entries] of Object.entries(src)) {
				index.postings[field].set(term, new Map(entries));
			}
		}
		const lens = obj.docLen?.[field];
		if (Array.isArray(lens)) {
			for (const [key, len] of lens) index.docLen[field].set(key, len);
		}
		if (typeof obj.totalLen?.[field] === "number") index.totalLen[field] = obj.totalLen[field];
	}
	return index;
}
