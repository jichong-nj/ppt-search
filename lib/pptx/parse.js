// lib/pptx/parse.js — PPTX 文本解析（契约 A）
import fs from "node:fs/promises";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";

export const FX = new XMLParser({
  ignoreAttributes: false,
  parseTagValue: false,
  preserveOrder: false,
});

/** 单元素/数组归一化 */
export function toArray(v) {
  if (v === undefined || v === null) return [];
  return Array.isArray(v) ? v : [v];
}

export function isObj(v) {
  return typeof v === "object" && v !== null;
}

export function getAttr(o, name) {
  if (!isObj(o)) return undefined;
  const v = o["@_" + name];
  return v === undefined || v === null ? undefined : String(v);
}

/** 相对路径解析：baseDir 为 "ppt/slides" 这类目录，target 为 rels 里的 Target */
export function resolveRel(baseDir, target) {
  const t = String(target || "");
  const parts = t.startsWith("/") ? [] : baseDir ? baseDir.split("/") : [];
  for (const seg of t.split("/")) {
    if (seg === "" || seg === ".") continue;
    if (seg === "..") parts.pop();
    else parts.push(seg);
  }
  return parts.join("/");
}

export function dirname(p) {
  const i = p.lastIndexOf("/");
  return i < 0 ? "" : p.slice(0, i);
}

export function relsPathFor(partPath) {
  const d = dirname(partPath);
  const base = partPath.slice(d.length + 1);
  return (d ? d + "/" : "") + "_rels/" + base + ".rels";
}

export async function readText(zip, path) {
  const f = zip.file(path);
  return f ? f.async("string") : null;
}

export async function parseXmlFile(zip, path) {
  const txt = await readText(zip, path);
  if (txt === null) return null;
  try {
    return FX.parse(txt);
  } catch {
    return {};
  }
}

/** 读取 .rels → Map<Id, {Target, Type, TargetMode}> */
export async function readRels(zip, relsPath) {
  const map = new Map();
  const doc = await parseXmlFile(zip, relsPath);
  if (!doc) return map;
  const root = doc.Relationships || doc["xml:Relationships"] || {};
  for (const r of toArray(root.Relationship)) {
    if (!isObj(r)) continue;
    const id = getAttr(r, "Id");
    if (!id) continue;
    map.set(id, {
      Target: getAttr(r, "Target") || "",
      Type: getAttr(r, "Type") || "",
      TargetMode: getAttr(r, "TargetMode") || "",
    });
  }
  return map;
}

/** 段落文本：拼接 a:r/a:t、a:fld、a:br（a:t 可能直接挂段落或嵌在 a:r 内） */
function paraText(p) {
	if (!isObj(p)) return "";
	let out = "";
	const collect = (node, key) => {
		for (const v of toArray(node[key])) {
			if (v === null || v === undefined) continue;
			if (typeof v === "string" || typeof v === "number") {
				out += String(v);
				continue;
			}
			if (isObj(v)) {
				if (v["#text"] !== undefined && v["#text"] !== null) out += String(v["#text"]);
				else collect(v, "a:t"); // a:r → a:t
			}
		}
	};
	for (const key of Object.keys(p)) {
		if (key.startsWith("@_")) continue;
		if (key === "a:t" || key === "a:r" || key === "a:fld") collect(p, key);
		else if (key === "a:br") out += "\n";
	}
	return out;
}

function txBodyParas(txBody) {
  const paras = [];
  for (const b of toArray(txBody)) {
    if (!isObj(b)) continue;
    for (const p of toArray(b["a:p"])) paras.push(paraText(p));
  }
  return paras;
}

function txBodyText(txBody) {
  return txBodyParas(txBody).join("\n");
}

function phType(sp) {
  if (!isObj(sp)) return undefined;
  const nvSpPr = sp["p:nvSpPr"];
  if (!isObj(nvSpPr)) return undefined;
  const nvPr = nvSpPr["p:nvPr"];
  if (!isObj(nvPr)) return undefined;
  for (const ph of toArray(nvPr["p:ph"])) {
    const t = getAttr(ph, "type");
    if (t) return t;
  }
  return undefined;
}

function tableText(tbl) {
  const lines = [];
  for (const t of toArray(tbl)) {
    if (!isObj(t)) continue;
    for (const tr of toArray(t["a:tr"])) {
      if (!isObj(tr)) continue;
      const cells = [];
      for (const tc of toArray(tr["a:tc"])) {
        const paras = [];
        for (const sp of toArray(tc["a:txBody"])) paras.push(...txBodyParas([sp]));
        cells.push(paras.filter((s) => s !== "").join(" ").replace(/[\t\n]/g, " "));
      }
      lines.push(cells.join("\t"));
    }
  }
  return lines;
}

/** 遍历 sp 树（grpSp 递归、graphicFrame 表格），收集段落行 */
function walkTree(node, lines) {
  if (!isObj(node)) return;
  for (const sp of toArray(node["p:sp"])) {
    const paras = txBodyParas(sp["p:txBody"]);
    for (const s of paras) if (s !== "") lines.push(s);
  }
  for (const sp of toArray(node["p:cxnSp"])) {
    const paras = txBodyParas(sp["p:txBody"]);
    for (const s of paras) if (s !== "") lines.push(s);
  }
  for (const gf of toArray(node["p:graphicFrame"])) {
    if (!isObj(gf)) continue;
    const g = gf["a:graphic"];
    for (const gd of toArray(isObj(g) ? g["a:graphicData"] : null)) {
      if (isObj(gd) && gd["a:tbl"]) lines.push(...tableText(gd["a:tbl"]));
      else if (isObj(gd)) {
        for (const tb of toArray(gd["p:txBody"])) for (const s of txBodyParas([tb])) if (s !== "") lines.push(s);
      }
    }
  }
  for (const g of toArray(node["p:grpSp"])) walkTree(g, lines);
}

/** 提取一页 slide XML 的 { title, body } */
export function extractSlideTexts(slideXml) {
  const doc = typeof slideXml === "string" ? FX.parse(slideXml) : slideXml;
  const pres = isObj(doc) ? doc["p:sld"] : null;
  const cSld = isObj(pres) ? pres["p:cSld"] : null;
  const tree = isObj(cSld) ? cSld["p:spTree"] : null;
  const lines = [];
  walkTree(tree, lines);
  const body = lines.join("\n");

  let title = "";
  const candidates = [
    ...toArray(tree && tree["p:sp"]),
    ...collectGrpSps(tree),
  ];
  for (const sp of candidates) {
    const t = phType(sp);
    if (t === "title" || t === "ctrTitle") {
      title = txBodyText(sp["p:txBody"]).trim();
      break;
    }
  }
  if (!title) {
    const first = lines.find((s) => s.trim() !== "");
    title = first ? first.trim().slice(0, 80) : "";
  }
  return { title, body };
}

function collectGrpSps(tree) {
  const out = [];
  const walk = (node) => {
    if (!isObj(node)) return;
    for (const g of toArray(node["p:grpSp"])) {
      out.push(...toArray(g["p:sp"]));
      walk(g);
    }
  };
  walk(tree);
  return out;
}

function notesText(notesDoc) {
  const ns = isObj(notesDoc) ? notesDoc["p:notes"] : null;
  const cSld = isObj(ns) ? ns["p:cSld"] : null;
  const tree = isObj(cSld) ? cSld["p:spTree"] : null;
  const parts = [];
  const scan = (node) => {
    if (!isObj(node)) return;
    const check = (sp) => {
      if (!isObj(sp)) return;
      const t = phType(sp);
      if (t === "body") {
        const s = txBodyText(sp["p:txBody"]).trim();
        if (s) parts.push(s);
      }
    };
    for (const sp of toArray(node["p:sp"])) check(sp);
    for (const g of toArray(node["p:grpSp"])) {
      for (const sp of toArray(g["p:sp"])) check(sp);
      scan(g);
    }
  };
  scan(tree);
  return parts.join("\n");
}

/** 按 presentation.xml sldIdLst 顺序返回 slide part 路径列表（供 compose 复用）。 */
export async function listSlideParts(zip) {
	const presDoc = await parseXmlFile(zip, "ppt/presentation.xml");
	if (!presDoc || !isObj(presDoc["p:presentation"])) throw new Error("invalid pptx: missing presentation.xml");
	const presRels = await readRels(zip, "ppt/_rels/presentation.xml.rels");
	const pres = presDoc["p:presentation"];
	const sldIdLst = isObj(pres["p:sldIdLst"]) ? pres["p:sldIdLst"] : {};
	const out = [];
	for (const sldId of toArray(sldIdLst["p:sldId"])) {
		const rid = getAttr(sldId, "r:id");
		const rel = rid ? presRels.get(rid) : undefined;
		if (!rel) continue;
		out.push(resolveRel("ppt", rel.Target));
	}
	return out;
}

/**
 * 解析 pptx 文件。
 * @returns {Promise<{slideSize:{cx:number,cy:number}, slideCount:number, slides:Array<{n:number,title:string,body:string,notes:string}>}>}
 */
export async function parsePptx(filePath) {
  let buf;
  try {
    buf = await fs.readFile(filePath);
  } catch (e) {
    throw new Error(`invalid pptx: cannot read file: ${e.message}`);
  }
  let zip;
  try {
    zip = await JSZip.loadAsync(buf);
  } catch (e) {
    throw new Error(`invalid pptx: ${e.message}`);
  }
  const presDoc = await parseXmlFile(zip, "ppt/presentation.xml");
  if (!presDoc || !isObj(presDoc["p:presentation"])) {
    throw new Error("invalid pptx: missing presentation.xml");
  }
  const pres = presDoc["p:presentation"];
  const sldSz = pres["p:sldSz"];
  const slideSize = {
    cx: sldSz ? Number(getAttr(sldSz, "cx")) || 0 : 0,
    cy: sldSz ? Number(getAttr(sldSz, "cy")) || 0 : 0,
  };

  const presRels = await readRels(zip, "ppt/_rels/presentation.xml.rels");
  const sldIdLst = isObj(pres["p:sldIdLst"]) ? pres["p:sldIdLst"] : {};
  const entries = [];
  for (const sldId of toArray(sldIdLst["p:sldId"])) {
    const rid = getAttr(sldId, "r:id");
    const rel = rid ? presRels.get(rid) : undefined;
    if (!rel) continue;
    entries.push(resolveRel("ppt", rel.Target));
  }

  const slides = [];
  let n = 0;
  for (const slidePath of entries) {
    n += 1;
    let title = "";
    let body = "";
    let notes = "";
    const xml = await readText(zip, slidePath);
    if (xml !== null) {
      try {
        const t = extractSlideTexts(xml);
        title = t.title;
        body = t.body;
      } catch {
        /* 空页/异常结构不抛错 */
      }
      const rels = await readRels(zip, relsPathFor(slidePath));
      let notesPath = null;
      for (const rel of rels.values()) {
        if (rel.Type.includes("notesSlide")) {
          notesPath = resolveRel(dirname(slidePath), rel.Target);
          break;
        }
      }
      if (notesPath) {
        const ndoc = await parseXmlFile(zip, notesPath);
        if (ndoc) {
          try {
            notes = notesText(ndoc);
          } catch {
            notes = "";
          }
        }
      }
    }
    slides.push({ n: slides.length + 1, title, body, notes });
  }

  return { slideSize, slideCount: slides.length, slides };
}
