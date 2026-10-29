// lib/pptx/compose.js — 契约 B：按提纲从多个 deck 克隆幻灯片合成新 pptx。
// 策略：以 baseDeck 的 zip 为底，递归复制所选 slide 及其全部依赖 part（layout→master→theme、
// media、chart 等），重命名避免冲突，重写 rels 与 [Content_Types].xml，重建 sldIdLst。

import fs from "node:fs/promises";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { toArray, resolveRel } from "./parse.js";

const FX = new XMLParser({ ignoreAttributes: false, parseTagValue: false, preserveOrder: false });
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const SLIDE_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

function dirname(p) {
	const i = p.lastIndexOf("/");
	return i < 0 ? "" : p.slice(0, i);
}
function basename(p) {
	return p.slice(p.lastIndexOf("/") + 1);
}
function extname(p) {
	const b = basename(p);
	const i = b.lastIndexOf(".");
	return i <= 0 ? "" : b.slice(i).toLowerCase();
}
function relsPathFor(partPath) {
	const d = dirname(partPath);
	const base = partPath.slice(d.length + 1);
	return (d ? d + "/" : "") + "_rels/" + base + ".rels";
}
function isTextPart(p) {
	const e = extname(p);
	return e === ".xml" || e === ".rels" || e === ".txt";
}
function escapeXml(s) {
	return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function unescapeXml(s) {
	return String(s)
		.replace(/&lt;/g, "<")
		.replace(/&gt;/g, ">")
		.replace(/&quot;/g, '"')
		.replace(/&apos;/g, "'")
		.replace(/&amp;/g, "&");
}
function xmlEscapeAttr(s) {
	return escapeXml(s);
}

/** 从 fromDir 到 toPath 的相对路径（package 内部，斜杠分隔） */
function relativePath(fromDir, toPath) {
	const from = fromDir === "" ? [] : fromDir.split("/");
	const to = toPath.split("/");
	let i = 0;
	while (i < from.length && i < to.length && from[i] === to[i]) i += 1;
	return "../".repeat(from.length - i) + to.slice(i).join("/");
}

function parseRels(xmlText) {
	if (xmlText === null) return [];
	const doc = FX.parse(xmlText);
	const root = doc.Relationships ?? {};
	const out = [];
	for (const r of toArray(root.Relationship)) {
		if (r === null || typeof r !== "object") continue;
		out.push({
			Id: String(r["@_Id"] ?? ""),
			Type: String(r["@_Type"] ?? ""),
			Target: String(r["@_Target"] ?? ""),
			TargetMode: String(r["@_TargetMode"] ?? ""),
		});
	}
	return out;
}

function buildRels(rels) {
	const body = rels
		.map((r) => {
			const mode = r.TargetMode === "External" ? ' TargetMode="External"' : "";
			return `<Relationship Id="${xmlEscapeAttr(r.Id)}" Type="${xmlEscapeAttr(r.Type)}" Target="${xmlEscapeAttr(r.Target)}"${mode}/>`;
		})
		.join("");
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${REL_NS}">${body}</Relationships>`;
}

function parseContentTypes(xmlText) {
	const doc = FX.parse(xmlText);
	const root = doc.Types ?? {};
	const defaults = new Map();
	const overrides = new Map();
	for (const d of toArray(root.Default)) {
		if (d === null || typeof d !== "object") continue;
		defaults.set(String(d["@_Extension"] ?? "").toLowerCase(), String(d["@_ContentType"] ?? ""));
	}
	for (const o of toArray(root.Override)) {
		if (o === null || typeof o !== "object") continue;
		overrides.set(String(o["@_PartName"] ?? ""), String(o["@_ContentType"] ?? ""));
	}
	return { defaults, overrides };
}

function buildContentTypes(ct) {
	const defs = [...ct.defaults.entries()].map(([ext, type]) => `<Default Extension="${xmlEscapeAttr(ext)}" ContentType="${xmlEscapeAttr(type)}"/>`).join("");
	const overs = [...ct.overrides.entries()].map(([part, type]) => `<Override PartName="${xmlEscapeAttr(part)}" ContentType="${xmlEscapeAttr(type)}"/>`).join("");
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="${CT_NS}">${defs}${overs}</Types>`;
}

// ── 文本修改 ────────────────────────────────────────────────────────────────

const TITLE_PH_RE = /<p:ph\b[^>]*\btype="(title|ctrTitle)"/;

function replaceATtexts(txBodyXml, text) {
	let first = true;
	return txBodyXml.replace(/<a:t>([\s\S]*?)<\/a:t>/g, (_m, inner) => {
		if (first && String(inner).trim() !== "") {
			first = false;
			return `<a:t>${escapeXml(text)}</a:t>`;
		}
		if (first && String(inner).trim() === "") {
			// 首个 run 为空也用它承载标题
			first = false;
			return `<a:t>${escapeXml(text)}</a:t>`;
		}
		return "<a:t></a:t>";
	});
}

/** 替换标题占位符文本；无标题占位符时原样返回。 */
export function applyTitleOverride(slideXml, text) {
	const spRe = /<p:sp>[\s\S]*?<\/p:sp>/g;
	let result = null;
	let match;
	while ((match = spRe.exec(slideXml)) !== null) {
		const block = match[0];
		if (!TITLE_PH_RE.test(block)) continue;
		const tbMatch = block.match(/<p:txBody>[\s\S]*?<\/p:txBody>/);
		if (tbMatch === null) continue;
		const newTb = replaceATtexts(tbMatch[0], text);
		const newBlock = block.replace(tbMatch[0], newTb);
		result = slideXml.slice(0, match.index) + newBlock + slideXml.slice(match.index + block.length);
		break;
	}
	return result === null ? slideXml : result;
}

/** 全局查找替换：仅作用于 <a:t> 文本节点内容。 */
export function applyReplace(slideXml, pairs) {
	const list = (pairs ?? []).filter((p) => typeof p?.find === "string" && p.find !== "" && typeof p.replace === "string");
	if (list.length === 0) return slideXml;
	return slideXml.replace(/<a:t>([\s\S]*?)<\/a:t>/g, (_m, inner) => {
		let t = unescapeXml(inner);
		for (const { find, replace } of list) t = t.split(find).join(replace);
		return `<a:t>${escapeXml(t)}</a:t>`;
	});
}

// ── 合成 ────────────────────────────────────────────────────────────────────

/**
 * @param {{items: Array<{deckId:string, n:number, titleOverride?:string, replace?:Array<{find:string,replace:string}>}>, outPath:string, decksRoot:string, baseDeckId?:string}} opts
 */
export async function composePptx({ items, outPath, decksRoot, baseDeckId }) {
	if (!Array.isArray(items) || items.length === 0) throw new Error("compose: items 不能为空");
	for (const it of items) {
		if (typeof it?.deckId !== "string" || !Number.isInteger(it?.n) || it.n < 1) throw new Error(`compose: 非法条目 ${JSON.stringify(it)}`);
	}
	const baseId = baseDeckId ?? items[0].deckId;
	const loadZip = async (deckId) => {
		const buf = await fs.readFile(`${decksRoot}/${deckId}/source.pptx`);
		return await JSZip.loadAsync(buf);
	};

	const baseZip = await loadZip(baseId);
	const deckZips = new Map([[baseId, baseZip]]);
	const getDeckZip = async (deckId) => {
		let z = deckZips.get(deckId);
		if (z === undefined) {
			z = await loadZip(deckId);
			deckZips.set(deckId, z);
		}
		return z;
	};

	// 已占用 part 名
	const used = new Set();
	baseZip.forEach((relPath, entry) => {
		if (!entry.dir) used.add(relPath);
	});

	function allocPart(srcPart) {
		const dir = dirname(srcPart);
		const b = basename(srcPart);
		const ext = extname(srcPart);
		const stem = ext === "" ? b : b.slice(0, -ext.length);
		const candidate = (name) => (dir === "" ? name : `${dir}/${name}`);
		let name = b;
		if (!used.has(candidate(name))) {
			used.add(candidate(name));
			return candidate(name);
		}
		for (let k = 2; k < 10_000; k += 1) {
			name = `${stem}-${k}${ext}`;
			if (!used.has(candidate(name))) {
				used.add(candidate(name));
				return candidate(name);
			}
		}
		throw new Error(`compose: 无法为 ${srcPart} 分配名称`);
	}

	const ct = parseContentTypes(await baseZip.file("[Content_Types].xml").async("string"));
	const remaps = new Map(); // deckId -> Map<srcPart, dstPart>

	/** 递归复制 srcZip 的 part（含其 rels 依赖）到 baseZip，返回目标路径 */
	async function copyPart(deckId, srcZip, srcPart) {
		let deckMap = remaps.get(deckId);
		if (deckMap === undefined) {
			deckMap = new Map();
			remaps.set(deckId, deckMap);
		}
		const seen = deckMap.get(srcPart);
		if (seen !== undefined) return seen;

		const srcFile = srcZip.file(srcPart);
		if (srcFile === null) throw new Error(`compose: 源 part 缺失 ${srcPart} (deck ${deckId})`);
		const dst = allocPart(srcPart);
		deckMap.set(srcPart, dst); // 先登记，防环

		// 处理 rels
		const relsFile = srcZip.file(relsPathFor(srcPart));
		if (relsFile !== null) {
			const rels = parseRels(await relsFile.async("string"));
			const newRels = [];
			for (const rel of rels) {
				if (rel.TargetMode === "External") {
					newRels.push(rel);
					continue;
				}
				if (rel.Type.endsWith("/notesSlide")) continue; // 备注页不复制
				const abs = resolveRel(dirname(srcPart), rel.Target);
				const dstDep = await copyPart(deckId, srcZip, abs);
				newRels.push({ ...rel, Target: relativePath(dirname(dst), dstDep) });
			}
			baseZip.file(relsPathFor(dst), buildRels(newRels));
		}

		const data = isTextPart(dst) ? await srcFile.async("string") : await srcFile.async("uint8array");
		baseZip.file(dst, data);

		// content types：优先源 Override，其次 Default ext
		const partName = `/${srcPart}`;
		const srcCt = await getDeckCt(deckId, srcZip);
		const override = srcCt.overrides.get(partName);
		const dstOverrideKey = `/${dst}`;
		if (!ct.overrides.has(dstOverrideKey)) {
			if (override !== undefined) ct.overrides.set(dstOverrideKey, override);
			else {
				const ext = extname(dst).replace(/^\./, "");
				const defType = srcCt.defaults.get(ext) ?? (ext === "xml" ? "application/xml" : "application/octet-stream");
				if (!ct.defaults.has(ext)) ct.defaults.set(ext, defType);
			}
		}
		return dst;
	}

	const deckCts = new Map();
	async function getDeckCt(deckId, zip) {
		let c = deckCts.get(deckId);
		if (c === undefined) {
			c = parseContentTypes(await zip.file("[Content_Types].xml").async("string"));
			deckCts.set(deckId, c);
		}
		return c;
	}

	// 预取各源 deck 的 slide part 列表
	const deckSlideParts = new Map();
	async function slidePartOf(deckId, n) {
		let list = deckSlideParts.get(deckId);
		if (list === undefined) {
			const { listSlideParts } = await import("./parse.js");
			list = await listSlideParts(await getDeckZip(deckId));
			deckSlideParts.set(deckId, list);
		}
		const part = list[n - 1];
		if (part === undefined) throw new Error(`compose: deck ${deckId} 没有第 ${n} 页（共 ${list.length} 页）`);
		return part;
	}

	// 逐条复制并应用修改
	const newSlideDsts = [];
	for (const it of items) {
		const srcZip = await getDeckZip(it.deckId);
		const srcPart = await slidePartOf(it.deckId, it.n);
		const dst = await copyPart(it.deckId, srcZip, srcPart);
		let xml = await baseZip.file(dst).async("string");
		if (typeof it.titleOverride === "string" && it.titleOverride !== "") xml = applyTitleOverride(xml, it.titleOverride);
		if (Array.isArray(it.replace) && it.replace.length > 0) xml = applyReplace(xml, it.replace);
		baseZip.file(dst, xml);
		newSlideDsts.push(dst);
	}

	// 重写 presentation.xml.rels：保留非 slide 关系，追加新 slide 关系
	const presRelsPath = "ppt/_rels/presentation.xml.rels";
	const presRels = parseRels(await baseZip.file(presRelsPath).async("string"));
	const kept = presRels.filter((r) => !r.Type.endsWith("/slide"));
	let maxRid = 0;
	for (const r of kept) {
		const m = /^rId(\d+)$/.exec(r.Id);
		if (m !== null) maxRid = Math.max(maxRid, Number(m[1]));
	}
	const addedRels = [];
	const sldIds = [];
	let maxSldId = 255;
	const presPath = "ppt/presentation.xml";
	const presXml = await baseZip.file(presPath).async("string");
	for (const m of presXml.matchAll(/<p:sldId\b[^>]*\bid="(\d+)"/g)) maxSldId = Math.max(maxSldId, Number(m[1]));
	newSlideDsts.forEach((dst, i) => {
		maxRid += 1;
		maxSldId += 1;
		const rid = `rId${maxRid}`;
		addedRels.push({ Id: rid, Type: SLIDE_TYPE, Target: relativePath("ppt", dst), TargetMode: "" });
		sldIds.push(`<p:sldId id="${maxSldId}" r:id="${rid}"/>`);
	});
	baseZip.file(presRelsPath, buildRels([...kept, ...addedRels]));

	// 重写 sldIdLst
	const newLst = `<p:sldIdLst>${sldIds.join("")}</p:sldIdLst>`;
	const lstRe = /<p:sldIdLst\b(?:[^>]*)\/>|<p:sldIdLst\b(?:[^>]*)>[\s\S]*?<\/p:sldIdLst>/;
	let newPres;
	if (lstRe.test(presXml)) newPres = presXml.replace(lstRe, newLst);
	else if (presXml.includes("<p:sldSz")) newPres = presXml.replace("<p:sldSz", `${newLst}<p:sldSz`);
	else throw new Error("compose: presentation.xml 中找不到 sldIdLst/sldSz 锚点");
	baseZip.file(presPath, newPres);

	// content types 重建
	baseZip.file("[Content_Types].xml", buildContentTypes(ct));

	const bytes = await baseZip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
	await fs.mkdir(dirname(outPath), { recursive: true }).catch(() => undefined);
	await fs.writeFile(outPath, bytes);
	return { slideCount: newSlideDsts.length, bytes: bytes.length };
}
