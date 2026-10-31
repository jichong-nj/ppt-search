// lib/pptx/renderModel.js — 纯 JS 渲染模型构建（Qoder 式）：pptx → 每页渲染模型 JSON + 媒体提取。
// 不依赖 LibreOffice。坐标一律 EMU；颜色在服务端完成主题解析，客户端只管画。
import path from "node:path";
import { mkdir, writeFile } from "node:fs/promises";
import JSZip from "jszip";
import { FX, toArray, getAttr, isObj, resolveRel, readRels, parseXmlFile, dirname, relsPathFor } from "./parse.js";

const EMU_PER_PT = 12700;

/* ---------- 小工具 ---------- */

function num(v, d = 0) {
	const n = Number(v);
	return Number.isFinite(n) ? n : d;
}

function textOf(node) {
	if (node === null || node === undefined) return "";
	if (typeof node === "string" || typeof node === "number") return String(node);
	if (isObj(node)) {
		if (node["#text"] !== undefined) return String(node["#text"]);
		return "";
	}
	return "";
}

/** 递归查找 key 以 suffix 结尾的第一个节点 */
function findDeep(node, suffix) {
	if (!isObj(node)) return null;
	for (const k of Object.keys(node)) {
		if (k.startsWith("@_")) continue;
		if (k.endsWith(suffix)) return node[k];
		const r = findDeep(node[k], suffix);
		if (r !== null) return r;
	}
	return null;
}

/* ---------- 颜色 ---------- */

function hex6(v) {
	const s = String(v || "").replace("#", "").toUpperCase();
	if (/^[0-9A-F]{6}$/.test(s)) return "#" + s.toLowerCase();
	return null;
}

function applyShade(hex, f) {
	const r = parseInt(hex.slice(1, 3), 16) * f;
	const g = parseInt(hex.slice(3, 5), 16) * f;
	const b = parseInt(hex.slice(5, 7), 16) * f;
	const c = (x) => Math.max(0, Math.min(255, Math.round(x)));
	return `#${[c(r), c(g), c(b)].map((x) => x.toString(16).padStart(2, "0")).join("")}`;
}

function applyTint(hex, f) {
	const mix = (x) => Math.round(x + (255 - x) * f);
	const r = mix(parseInt(hex.slice(1, 3), 16));
	const g = mix(parseInt(hex.slice(3, 5), 16));
	const b = mix(parseInt(hex.slice(5, 7), 16));
	return `#${[r, g, b].map((x) => x.toString(16).padStart(2, "0")).join("")}`;
}

/** 解析主题 clrScheme → {dk1,lt1,dk2,lt2,accent1..6,hlink,folHlink} */
function parseTheme(doc) {
	const out = { dk1: "#000000", lt1: "#ffffff", dk2: "#44546a", lt2: "#e7e6e5", hlink: "#0563c1", folHlink: "#954f72" };
	for (let i = 1; i <= 6; i += 1) out["accent" + i] = "#4472c4";
	const scheme =
		doc?.["a:theme"]?.["a:themeElements"]?.["a:clrScheme"] ??
		doc?.["a:theme"]?.themeElements?.clrScheme ??
		doc?.theme?.themeElements?.clrScheme;
	if (!isObj(scheme)) return out;
	for (const key of Object.keys(out)) {
		const entry = scheme["a:" + key] ?? scheme[key];
		if (!isObj(entry)) continue;
		const srgb = entry["a:srgbClr"];
		const sys = entry["a:sysClr"];
		const hex = (isObj(srgb) ? hex6(getAttr(srgb, "val")) : null) ?? (isObj(sys) ? hex6(getAttr(sys, "lastClr")) : null);
		if (hex) out[key] = hex;
	}
	return out;
}

/** clrMap：schemeClr val → clrScheme 键 */
function parseClrMap(masterDoc) {
	const map = { bg1: "lt1", tx1: "dk1", bg2: "lt2", tx2: "dk2", hlink: "hlink", folHlink: "folHlink" };
	const cm = masterDoc?.sldMaster?.p?.clrMap ?? masterDoc?.["p:sldMaster"]?.["p:clrMap"];
	if (isObj(cm)) {
		for (const k of ["bg1", "tx1", "bg2", "tx2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]) {
			const v = getAttr(cm, k);
			if (v) map[k] = v;
		}
	}
	return map;
}

/** 解析颜色节点（srgbClr/sysClr/schemeClr，含 alpha/shade/tint/lumMod/lumOff）→ {c,a} | null */
function colorFromNode(node, ctx) {
	if (!isObj(node)) return null;
	let hex = null;
	let alpha = 1;
	const srgb = node["a:srgbClr"];
	const sys = node["a:sysClr"];
	const scheme = node["a:schemeClr"];
	const src = isObj(srgb) ? srgb : isObj(sys) ? sys : isObj(scheme) ? scheme : null;
	if (src === null) return null;
	if (isObj(srgb)) hex = hex6(getAttr(srgb, "val"));
	else if (isObj(sys)) hex = hex6(getAttr(sys, "lastClr"));
	else {
		const val = getAttr(scheme, "val") || "tx1";
		const key = ctx.clrMap[val] ?? val;
		hex = ctx.theme[key] ?? "#000000";
		for (const m of toArray(scheme["a:lumMod"])) alpha *= num(getAttr(m, "val"), 100000) / 100000;
		for (const m of toArray(scheme["a:lumOff"])) {
			const f = num(getAttr(m, "val"), 0) / 100000;
			hex = applyTint(hex, f);
		}
	}
	if (!hex) return null;
	const holder = isObj(scheme) ? scheme : src;
	for (const m of toArray(holder["a:alpha"])) alpha *= num(getAttr(m, "val"), 100000) / 100000;
	for (const m of toArray(holder["a:shade"])) hex = applyShade(hex, num(getAttr(m, "val"), 100000) / 100000);
	for (const m of toArray(holder["a:tint"])) hex = applyTint(hex, num(getAttr(m, "val"), 100000) / 100000);
	return { c: hex, a: Math.max(0, Math.min(1, alpha)) };
}

/* ---------- 填充 / 描边 ---------- */

/** 从含 fill 子元素的节点（spPr/tcPr/bgPr…）解析填充 */
function fillFromPr(pr, ctx) {
	if (!isObj(pr)) return null;
	if (pr["a:noFill"] !== undefined) return { t: "none" };
	const solid = toArray(pr["a:solidFill"])[0];
	if (isObj(solid)) {
		const c = colorFromNode(solid, ctx);
		if (c) return { t: "solid", c: c.c, a: c.a };
	}
	const grad = toArray(pr["a:gradFill"])[0];
	if (isObj(grad)) {
		const stops = [];
		for (const lst of toArray(grad["a:gsLst"])) {
			for (const gs of toArray(lst["a:gs"])) {
				const c = colorFromNode(gs, ctx);
				if (c) stops.push({ p: num(getAttr(gs, "pos"), 0) / 100000, c: c.c, a: c.a });
			}
		}
		if (stops.length > 0) {
			const lin = toArray(grad["a:lin"])[0];
			const pathEl = toArray(grad["a:path"])[0];
			if (isObj(lin)) return { t: "grad", kind: "lin", ang: num(getAttr(lin, "ang"), 0) / 60000, stops };
			if (isObj(pathEl)) return { t: "grad", kind: "radial", stops };
			return { t: "grad", kind: "lin", ang: 90, stops };
		}
	}
	const blipFill = toArray(pr["a:blipFill"])[0];
	if (isObj(blipFill)) {
		const blip = toArray(blipFill["a:blip"])[0];
		const rid = isObj(blip) ? getAttr(blip, "r:embed") : undefined;
		const src = rid ? ctx.registerMedia(rid) : null;
		if (src) {
			let crop = null;
			const sr = toArray(blipFill["a:srcRect"])[0];
			if (isObj(sr)) crop = ["l", "t", "r", "b"].map((k) => num(getAttr(sr, k), 0) / 100000);
			return { t: "img", src, crop };
		}
	}
	const patt = toArray(pr["a:pattFill"])[0];
	if (isObj(patt)) {
		const fg = toArray(patt["a:fgClr"])[0];
		const c = colorFromNode(fg, ctx);
		if (c) return { t: "solid", c: c.c, a: 0.6 };
	}
	return null;
}

/** style 里的 fillRef/lnRef idx → 主题色 */
function refColor(idx, ctx, shade) {
	const i = num(idx, 0);
	if (i === 0) return null;
	let hex = null;
	if (i >= 1 && i <= 6) hex = ctx.theme["accent" + i];
	else if (i === 9) hex = ctx.theme[ctx.clrMap.tx1] ?? "#000000";
	else if (i === 10) hex = ctx.theme[ctx.clrMap.tx2] ?? "#000000";
	else if (i === 11) hex = ctx.theme[ctx.clrMap.bg1] ?? "#ffffff";
	else if (i === 12) hex = ctx.theme[ctx.clrMap.bg2] ?? "#ffffff";
	else hex = ctx.theme.accent1;
	if (!hex) return null;
	if (shade !== undefined) hex = applyShade(hex, num(shade, 100000) / 100000);
	return { t: "solid", c: hex, a: 1 };
}

function parseLn(pr, ctx) {
	const ln = toArray(pr?.["a:ln"])[0];
	if (!isObj(ln)) return null;
	if (ln["a:noFill"] !== undefined) return { w: 0, c: null };
	const c = colorFromNode(toArray(ln["a:solidFill"])[0], ctx);
	const dash = toArray(ln["a:prstDash"])[0];
	return {
		w: num(getAttr(ln, "w"), 12700),
		c: c ? c.c : "#000000",
		a: c ? c.a : 1,
		dash: isObj(dash) ? getAttr(dash, "val") ?? null : null,
	};
}

function parseXfrm(node) {
	if (!isObj(node)) return null;
	const off = node["a:off"];
	const ext = node["a:ext"];
	if (!isObj(off) || !isObj(ext)) return null;
	return {
		x: num(getAttr(off, "x"), 0),
		y: num(getAttr(off, "y"), 0),
		cx: num(getAttr(ext, "cx"), 0),
		cy: num(getAttr(ext, "cy"), 0),
		rot: num(getAttr(node, "rot"), 0) / 60000,
		fh: getAttr(node, "flipH") === "1",
		fv: getAttr(node, "flipV") === "1",
	};
}

/* ---------- 文本 ---------- */

function fldText(type, slideNo) {
	if (type === "sldNum") return String(slideNo);
	if (type === "datetime") return new Date().toISOString().slice(0, 10);
	return "";
}

function parseRunsOfPara(p, ctx, opts) {
	const runs = [];
	for (const key of Object.keys(p)) {
		if (key.startsWith("@_")) continue;
		if (key === "a:r") {
			for (const r of toArray(p[key])) {
				if (!isObj(r)) continue;
				const t = textOf(r["a:t"]);
				if (t === "") continue;
				runs.push(runOf(r["a:rPr"], t, ctx, opts));
			}
		} else if (key === "a:fld") {
			for (const f of toArray(p[key])) {
				if (!isObj(f)) continue;
				const t = textOf(f) || fldText(getAttr(f, "type") || "", opts.slideNo);
				if (t !== "") runs.push(runOf(f["a:rPr"], t, ctx, opts));
			}
		} else if (key === "a:br") {
			runs.push({ t: "\n", sz: 18, br: 1 });
		} else if (key === "a:t") {
			for (const t of toArray(p[key])) {
				const s = textOf(t);
				if (s !== "") runs.push(runOf(null, s, ctx, opts));
			}
		}
	}
	return runs;
}

function runOf(rPr, text, ctx, opts) {
	const run = { t: text, sz: 18, c: null };
	if (isObj(rPr)) {
		if (rPr["@_sz"] !== undefined) run.sz = num(getAttr(rPr, "sz"), 1800) / 100;
		if (rPr["@_spc"] !== undefined) run.spc = num(getAttr(rPr, "spc"), 0) / 100;
		else if (typeof opts.spcDef === "number") run.spc = opts.spcDef;
		if (getAttr(rPr, "b") === "1") run.b = 1;
		if (getAttr(rPr, "i") === "1") run.i = 1;
		const u = getAttr(rPr, "u");
		if (u && u !== "none") run.u = u;
		const st = getAttr(rPr, "strike");
		if (st && st !== "noStrike") run.s = st;
		const c = colorFromNode(toArray(rPr["a:solidFill"])[0], ctx);
		if (c) run.c = c.c, (run.a = c.a);
		const latin = toArray(rPr["a:latin"])[0];
		if (isObj(latin)) run.f = getAttr(latin, "typeface") || undefined;
		const ea = toArray(rPr["a:ea"])[0];
		if (isObj(ea)) run.ea = getAttr(ea, "typeface") || undefined;
		const base = num(getAttr(rPr, "baseline"), 0);
		if (base !== 0) run.base = base / 1000;
	}
	if (run.c === null) {
		const c = colorFromNode({ "a:schemeClr": { "@_val": "tx1" } }, ctx);
		run.c = c ? c.c : "#000000";
	}
	if (opts.szScale !== 1) run.sz = run.sz * opts.szScale;
	if (opts.fontScale !== 1) run.sz = run.sz * opts.fontScale;
	return run;
}

function parseTextBody(txBody, ctx, opts) {
	const b = toArray(txBody)[0];
	if (!isObj(b)) return null;
	const bodyPr = isObj(b["a:bodyPr"]) ? b["a:bodyPr"] : {};
	const tb = {
		anchor: getAttr(bodyPr, "anchor") || "t",
		ctr: getAttr(bodyPr, "anchorCtr") === "1",
		ins: [num(getAttr(bodyPr, "lIns"), 91440), num(getAttr(bodyPr, "tIns"), 45720), num(getAttr(bodyPr, "rIns"), 91440), num(getAttr(bodyPr, "bIns"), 45720)],
		wrap: getAttr(bodyPr, "wrap") !== "none",
		paras: [],
	};
	let fontScale = 1;
	for (const na of toArray(bodyPr["a:normAutofit"])) {
		const fs = num(getAttr(na, "fontScale"), 100000);
		if (fs > 0 && fs < 100000) fontScale = fs / 100000;
	}
	const paras = toArray(b["a:p"]);
	for (const p of paras) {
		if (!isObj(p)) {
			tb.paras.push({ runs: [] });
			continue;
		}
		const pPr = isObj(p["a:pPr"]) ? p["a:pPr"] : {};
		const para = { algn: getAttr(pPr, "algn") || "l", marL: 0, indent: 0, bullet: null, runs: [] };
		if (pPr["@_marL"] !== undefined) para.marL = num(getAttr(pPr, "marL"), 0);
		if (pPr["@_indent"] !== undefined) para.indent = num(getAttr(pPr, "indent"), 0);
		const lnSpc = toArray(pPr["a:lnSpc"])[0];
		if (isObj(lnSpc)) {
			const pct = toArray(lnSpc["a:spcPct"])[0];
			const pts = toArray(lnSpc["a:spcPts"])[0];
			if (isObj(pct)) para.lnSpc = num(getAttr(pct, "val"), 100000) / 100000;
			else if (isObj(pts)) para.lnPts = num(getAttr(pts, "val"), 0) / 100;
		}
		const spcBef = toArray(pPr["a:spcBef"])[0];
		if (isObj(spcBef)) {
			const pts = toArray(spcBef["a:spcPts"])[0];
			const pct = toArray(spcBef["a:spcPct"])[0];
			if (isObj(pts)) para.spcBef = num(getAttr(pts, "val"), 0) / 100;
			else if (isObj(pct)) para.spcBefPct = num(getAttr(pct, "val"), 0) / 100000;
		}
		if (pPr["a:buNone"] !== undefined) para.bullet = null;
		else {
			const buChar = toArray(pPr["a:buChar"])[0];
			const buAuto = toArray(pPr["a:buAutoNum"])[0];
			if (isObj(buChar)) para.bullet = { char: getAttr(buChar, "char") || "•" };
			else if (isObj(buAuto)) para.bullet = { auto: getAttr(buAuto, "type") || "arabicPeriod" };
			else if (para.marL !== 0 || para.indent !== 0) para.bullet = { char: "•" };
			if (para.bullet) {
				const buClr = toArray(pPr["a:buClr"])[0];
				const c = isObj(buClr) ? colorFromNode(buClr, ctx) : null;
				if (c) para.bullet.c = c.c;
				const buFont = toArray(pPr["a:buFont"])[0];
				if (isObj(buFont)) para.bullet.f = getAttr(buFont, "typeface") || undefined;
			}
		}
		const defRPr = toArray(pPr["a:defRPr"])[0];
		let spcDef;
		if (isObj(defRPr) && defRPr["@_spc"] !== undefined) spcDef = num(getAttr(defRPr, "spc"), 0) / 100;
		para.runs = parseRunsOfPara(p, ctx, { ...opts, fontScale: (opts.fontScale ?? 1) * fontScale, spcDef });
		tb.paras.push(para);
	}
	if (tb.paras.length === 0) return null;
	return tb;
}

/* ---------- 几何 ---------- */

function parseGeom(spPr) {
	const pg = toArray(spPr?.["a:prstGeom"])[0];
	if (isObj(pg)) {
		const prst = getAttr(pg, "prst") || "rect";
		const adj = {};
		for (const lst of toArray(pg["a:avLst"])) {
			for (const gd of toArray(lst["a:gd"])) {
				const fmla = getAttr(gd, "fmla") || "";
				const m = /val\s+(-?\d+)/.exec(fmla);
				if (m) adj[getAttr(gd, "name") || "adj"] = num(m[1], 0);
			}
		}
		return { geom: prst, adj };
	}
	if (spPr?.["a:custGeom"] !== undefined) return { geom: "rect" };
	return { geom: "rect" };
}

/* ---------- 媒体注册 ---------- */

function makeMediaRegistry(zip, rels, partDir, mediaMap) {
	return function registerMedia(rid) {
		const rel = rels.get(rid);
		if (!rel) return null;
		const part = resolveRel(partDir, rel.Target);
		if (mediaMap.has(part)) return mediaMap.get(part);
		const base = part.split("/").pop().replace(/[^\w.\-+]/g, "_");
		let name = base;
		let i = 1;
		const used = new Set([...mediaMap.values()]);
		while (used.has(name)) {
			const ext = base.includes(".") ? "." + base.split(".").pop() : "";
			name = `${base.slice(0, base.length - ext.length)}_${i++}${ext}`;
		}
		mediaMap.set(part, name);
		return name;
	};
}

/* ---------- 元素构建 ---------- */

/** 构建 sp/cxnSp 元素；tf 为组变换 {sx,sy,dx,dy,rot,szScale} */
function buildSp(sp, ctx, tf) {
	const spPr = isObj(sp["p:spPr"]) ? sp["p:spPr"] : {};
	const x = parseXfrm(spPr["a:xfrm"]) ?? ctx.inheritXfrm?.(sp) ?? null;
	if (x === null) return null;
	const g = parseGeom(spPr);
	let fill = fillFromPr(spPr, ctx);
	let ln = parseLn(spPr, ctx);
	const style = isObj(sp["p:style"]) ? sp["p:style"] : null;
	if (style) {
		const fillRef = toArray(style["a:fillRef"])[0];
		if (fill === null && isObj(fillRef)) {
			const c = colorFromNode(fillRef, ctx) ?? refColor(getAttr(fillRef, "idx"), ctx, getAttr(fillRef, "shade"));
			if (c !== null) fill = { t: "solid", c: c.c, a: c.a };
		}
		const lnRef = toArray(style["a:lnRef"])[0];
		if (ln === null && isObj(lnRef)) {
			const c = colorFromNode(lnRef, ctx) ?? refColor(getAttr(lnRef, "idx"), ctx, getAttr(lnRef, "shade"));
			if (c !== null) ln = { w: 9525, c: c.c, a: c.a };
		}
	}
	const el = {
		k: "sp",
		x: tf.dx + (x.x - tf.ox) * tf.sx,
		y: tf.dy + (x.y - tf.oy) * tf.sy,
		cx: x.cx * tf.sx,
		cy: x.cy * tf.sy,
		rot: x.rot + tf.rot,
		geom: g.geom,
		fill,
		ln,
	};
	if (x.fh) el.fh = 1;
	if (x.fv) el.fv = 1;
	const text = parseTextBody(sp["p:txBody"], ctx, { slideNo: ctx.slideNo, szScale: 1, fontScale: 1 });
	if (text !== null) el.text = text;
	return el;
}

function buildPic(pic, ctx, tf) {
	const spPr = isObj(pic["p:spPr"]) ? pic["p:spPr"] : {};
	const x = parseXfrm(spPr["a:xfrm"]);
	if (x === null) return null;
	// 视频 → 占位
	const nvPr = pic["p:nvPicPr"]?.["p:nvPr"];
	if (isObj(nvPr) && (findDeep(nvPr, "videoFile") !== null || findDeep(nvPr, "mediaFile") !== null)) {
		return { k: "ph", x: tf.dx + (x.x - tf.ox) * tf.sx, y: tf.dy + (x.y - tf.oy) * tf.sy, cx: x.cx * tf.sx, cy: x.cy * tf.sy, label: "视频" };
	}
	const blipFill = toArray(pic["p:blipFill"])[0];
	const blip = isObj(blipFill) ? toArray(blipFill["a:blip"])[0] : null;
	const rid = isObj(blip) ? getAttr(blip, "r:embed") : undefined;
	let src = rid ? ctx.registerMedia(rid) : null;
	let svg = null;
	const svgBlip = findDeep(pic, "svgBlip");
	if (isObj(svgBlip)) {
		const srid = getAttr(svgBlip, "r:embed");
		if (srid) svg = ctx.registerMedia(srid);
	}
	if (src === null && svg === null) return null;
	if (src !== null && /\.(emf|wmf)$/i.test(src)) {
		return { k: "ph", x: tf.dx + (x.x - tf.ox) * tf.sx, y: tf.dy + (x.y - tf.oy) * tf.sy, cx: x.cx * tf.sx, cy: x.cy * tf.sy, label: "图元文件" };
	}
	const el = {
		k: "pic",
		x: tf.dx + (x.x - tf.ox) * tf.sx,
		y: tf.dy + (x.y - tf.oy) * tf.sy,
		cx: x.cx * tf.sx,
		cy: x.cy * tf.sy,
		rot: x.rot + tf.rot,
		src: src ?? svg,
		svg: src === null ? svg : null,
	};
	const sr = isObj(blipFill) ? toArray(blipFill["a:srcRect"])[0] : null;
	if (isObj(sr)) el.crop = ["l", "t", "r", "b"].map((k) => num(getAttr(sr, k), 0) / 100000);
	return el;
}

function buildTable(gf, tbl, ctx, tf) {
	const grid = toArray(tbl["a:tblGrid"])[0];
	const cols = grid ? toArray(grid["a:gridCol"]).map((c) => num(getAttr(c, "w"), 0)) : [];
	const trs = toArray(tbl["a:tr"]);
	const rows = trs.map((tr) => num(getAttr(tr, "h"), 0));
	const cells = [];
	trs.forEach((tr, r) => {
		const row = [];
		let c = 0;
		for (const tc of toArray(tr["a:tc"])) {
			if (!isObj(tc)) continue;
			while (cells.some((pc) => pc.r + pc.rs > r && pc.c <= c && pc.c + pc.cs > c)) c += 1;
			const cell = { r, c, cs: num(getAttr(tc, "gridSpan"), 1), rs: num(getAttr(tc, "rowSpan"), 1) };
			if (getAttr(tc, "hMerge") === "1" || getAttr(tc, "vMerge") === "1") {
				c += cell.cs;
				continue;
			}
			const tcPr = isObj(tc["a:tcPr"]) ? tc["a:tcPr"] : {};
			const fill = fillFromPr(tcPr, ctx);
			if (fill !== null) cell.fill = fill;
			const anchor = getAttr(tcPr, "anchor") ?? (toArray(tcPr["a:anchor"])[0] !== undefined ? "ctr" : "t");
			const ins = [num(getAttr(tcPr, "marL"), 91440), num(getAttr(tcPr, "marT"), 45720), num(getAttr(tcPr, "marR"), 91440), num(getAttr(tcPr, "marB"), 45720)];
			const borders = {};
			for (const [key, side] of [["a:lnL", "l"], ["a:lnR", "r"], ["a:lnT", "t"], ["a:lnB", "b"]]) {
				const ln = toArray(tcPr[key])[0];
				if (!isObj(ln)) continue;
				if (ln["a:noFill"] !== undefined) {
					borders[side] = null;
					continue;
				}
				const col = colorFromNode(toArray(ln["a:solidFill"])[0], ctx);
				borders[side] = { w: num(getAttr(ln, "w"), 12700), c: col ? col.c : "#000000" };
			}
			if (Object.keys(borders).length > 0) cell.borders = borders;
			const text = parseTextBody(tc["a:txBody"], ctx, { slideNo: ctx.slideNo, szScale: 1, fontScale: 1 });
			if (text !== null) {
				text.anchor = anchor === "ctr" ? "ctr" : anchor === "b" ? "b" : "t";
				text.ins = ins;
				cell.text = text;
			}
			c += cell.cs;
			row.push(cell);
		}
		cells.push(...row);
	});
	const xf = parseXfrm(gf["p:xfrm"]);
	if (xf === null) return null;
	return {
		k: "tbl",
		x: tf.dx + (xf.x - tf.ox) * tf.sx,
		y: tf.dy + (xf.y - tf.oy) * tf.sy,
		cx: xf.cx * tf.sx,
		cy: xf.cy * tf.sy,
		cols: cols.map((w) => w * tf.sx),
		rows: rows.map((h) => h * tf.sy),
		cells,
	};
}

/* ---------- 图表 ---------- */

function chartCachePts(node, ctx) {
	if (!isObj(node)) return null;
	const ref = toArray(node["c:strRef"])[0] ?? toArray(node["c:numRef"])[0] ?? node;
	const strCache = toArray(ref["c:strCache"])[0] ?? toArray(ref["c:numCache"])[0];
	if (!isObj(strCache)) return null;
	const pts = toArray(strCache["c:pt"]);
	const out = [];
	for (const pt of pts) {
		if (!isObj(pt)) continue;
		const idx = num(getAttr(pt, "idx"), out.length);
		out[idx] = textOf(pt["c:v"]);
	}
	for (let i = 0; i < out.length; i += 1) if (out[i] === undefined) out[i] = "";
	return out;
}

function chartSeriesColor(ser, idx, ctx) {
	const spPr = isObj(ser["c:spPr"]) ? ser["c:spPr"] : null;
	if (spPr) {
		const f = fillFromPr(spPr, ctx);
		if (f && f.t === "solid") return f.c;
	}
	return ctx.theme["accent" + ((idx % 6) + 1)] ?? "#4472c4";
}

function chartTitle(chartNode, ctx) {
	const title = toArray(chartNode["c:title"])[0];
	if (!isObj(title)) return "";
	const rich = toArray(title["c:tx"])[0];
	const paras = toArray(isObj(rich) ? rich["c:rich"]?.["a:p"] : null);
	const lines = [];
	for (const p of paras) lines.push(...parseRunsOfPara(p, ctx, { slideNo: 1, szScale: 1, fontScale: 1 }).map((r) => r.t));
	return lines.join("").trim();
}

function parseChartPart(doc, ctx) {
	const chart = doc?.["c:chartSpace"]?.["c:chart"];
	if (!isObj(chart)) return null;
	const plotArea = isObj(chart["c:plotArea"]) ? chart["c:plotArea"] : {};
	const kinds = [
		["c:barChart", "bar"],
		["c:lineChart", "line"],
		["c:pieChart", "pie"],
		["c:doughnutChart", "doughnut"],
		["c:areaChart", "area"],
		["c:scatterChart", "scatter"],
		["c:radarChart", "radar"],
		["c:bubbleChart", "scatter"],
	];
	for (const [key, type] of kinds) {
		const plot = toArray(plotArea[key])[0];
		if (!isObj(plot)) continue;
		const model = { type, dir: "col", grouping: "clustered", title: chartTitle(chart, ctx), cats: [], series: [], legend: true };
		if (type === "bar") {
			model.dir = getAttr(toArray(plot["c:barDir"])[0], "val") || "col";
			model.grouping = getAttr(toArray(plot["c:grouping"])[0], "val") || "clustered";
		}
		if (type === "doughnut") model.hole = num(getAttr(toArray(plot["c:holeSize"])[0], "val"), 50) / 100;
		let si = 0;
		for (const ser of toArray(plot["c:ser"])) {
			if (!isObj(ser)) continue;
			const s = { name: "", vals: [], color: chartSeriesColor(ser, si, ctx) };
			const tx = toArray(ser["c:tx"])[0];
			if (isObj(tx)) {
				const pts = chartCachePts(tx["c:strRef"], ctx);
				if (pts && pts.length > 0) s.name = pts[0];
			}
			if (type === "scatter") {
				const xv = chartCachePts(toArray(ser["c:xVal"])[0], ctx);
				const yv = chartCachePts(toArray(ser["c:yVal"])[0], ctx);
				if (yv) s.vals = yv.map((v) => num(v, 0));
				if (xv && model.cats.length === 0) model.cats = xv.map((v) => String(v));
			} else {
				const val = chartCachePts(toArray(ser["c:val"])[0], ctx);
				if (val) s.vals = val.map((v) => num(v, 0));
				const cat = chartCachePts(toArray(ser["c:cat"])[0], ctx);
				if (cat && model.cats.length === 0) model.cats = cat.map((v) => String(v));
			}
			model.series.push(s);
			si += 1;
		}
		const legend = toArray(chart["c:legend"])[0];
		if (isObj(legend)) {
			if (getAttr(toArray(legend["c:legendPos"])[0] ?? {}, "val") === "none") model.legend = false;
			const del = toArray(legend["c:delete"])[0];
			if (isObj(del) && getAttr(del, "val") === "1") model.legend = false;
		}
		if (model.series.length === 0) return null;
		return model;
	}
	return null;
}

function buildGraphicFrame(gf, ctx, tf) {
	const gd = toArray(gf["a:graphic"]?.["a:graphicData"])[0];
	if (!isObj(gd)) return null;
	if (isObj(gd["a:tbl"])) return buildTable(gf, gd["a:tbl"], ctx, tf);
	const uri = getAttr(gd, "uri") || "";
	const xf = parseXfrm(gf["p:xfrm"]);
	const box = xf !== null ? { x: tf.dx + (xf.x - tf.ox) * tf.sx, y: tf.dy + (xf.y - tf.oy) * tf.sy, cx: xf.cx * tf.sx, cy: xf.cy * tf.sy } : null;
	if (uri.includes("/chart") && box !== null) {
		const rid = getAttr(gd["c:chart"] ?? {}, "r:id");
		const rel = rid ? ctx.rels.get(rid) : undefined;
		if (rel) {
			const part = resolveRel(ctx.partDir, rel.Target);
			const entry = ctx.charts?.get(part);
			if (entry !== undefined && entry.doc !== null) {
				const cctx = { ...ctx, rels: entry.rels, partDir: dirname(part), registerMedia: makeMediaRegistry(ctx.zip, entry.rels, dirname(part), ctx.mediaMap) };
				const model = parseChartPart(entry.doc, cctx);
				if (model !== null) return { k: "chart", ...box, chart: model };
			}
		}
		return { k: "ph", ...box, label: "图表" };
	}
	if (uri.includes("/diagram") && box !== null) return { k: "ph", ...box, label: "SmartArt" };
	if (uri.includes("/inkDrawing") && box !== null) return { k: "ph", ...box, label: "墨迹" };
	// 其他 graphicFrame：若有 txBody 则当文本框
	const text = parseTextBody(gd["p:txBody"], ctx, { slideNo: ctx.slideNo, szScale: 1, fontScale: 1 });
	if (text !== null && box !== null) return { k: "sp", ...box, geom: "rect", fill: null, ln: null, text };
	return null;
}

const IDENTITY_TF = { sx: 1, sy: 1, ox: 0, oy: 0, dx: 0, dy: 0, rot: 0, fh: false, fv: false };

const RANK_FALLBACK = Number.MAX_SAFE_INTEGER;

/** 取元素节点自身 cNvPr id 对应的文档顺序序号 */
function nodeRank(node, ctx) {
	if (!isObj(node)) return RANK_FALLBACK;
	for (const k of ["p:nvSpPr", "p:nvGrpSpPr", "p:nvPicPr", "p:nvGraphicFramePr", "p:nvCxnSpPr"]) {
		const nv = node[k];
		if (!isObj(nv)) continue;
		const c = nv["p:cNvPr"];
		if (isObj(c)) {
			const id = getAttr(c, "id");
			if (id !== undefined) {
				const r = ctx?.ranks?.get(String(id));
				return r === undefined ? RANK_FALLBACK : r;
			}
		}
	}
	return RANK_FALLBACK;
}

function buildTree(tree, ctx, tf) {
	const els = [];
	if (!isObj(tree)) return els;
	const push = (el, node) => {
		if (el !== null && el !== undefined) {
			el.__r = nodeRank(node, ctx);
			els.push(el);
		}
	};
	for (const sp of toArray(tree["p:sp"])) {
		if (ctx.isPlaceholder?.(sp)) continue;
		push(buildSp(sp, ctx, tf), sp);
	}
	for (const sp of toArray(tree["p:cxnSp"])) push(buildSp(sp, ctx, tf), sp);
	for (const pic of toArray(tree["p:pic"])) push(buildPic(pic, ctx, tf), pic);
	for (const gf of toArray(tree["p:graphicFrame"])) push(buildGraphicFrame(gf, ctx, tf), gf);
	for (const grp of toArray(tree["p:grpSp"])) {
		const grpPr = isObj(grp["p:grpSpPr"]) ? grp["p:grpSpPr"] : {};
		const xf = parseXfrm(grpPr["a:xfrm"]);
		const chOff = isObj(grpPr["a:xfrm"]) ? grpPr["a:xfrm"]["a:chOff"] : null;
		const chExt = isObj(grpPr["a:xfrm"]) ? grpPr["a:xfrm"]["a:chExt"] : null;
		const childTf = { ...tf };
		if (xf !== null && isObj(chOff) && isObj(chExt) && num(getAttr(chExt, "cx"), 0) > 0 && num(getAttr(chExt, "cy"), 0) > 0) {
			const sx = (xf.cx / num(getAttr(chExt, "cx"), 1)) * tf.sx;
			const sy = (xf.cy / num(getAttr(chExt, "cy"), 1)) * tf.sy;
			const gx = tf.dx + (xf.x - tf.ox) * tf.sx;
			const gy = tf.dy + (xf.y - tf.oy) * tf.sy;
			childTf.sx = sx;
			childTf.sy = sy;
			childTf.ox = num(getAttr(chOff, "x"), 0);
			childTf.oy = num(getAttr(chOff, "y"), 0);
			childTf.dx = gx;
			childTf.dy = gy;
			childTf.rot = tf.rot + xf.rot;
		}
		els.push(...buildTree(grp, ctx, childTf));
	}
	// 按 spTree 文档顺序排序（PPT 的 z-order = 文档顺序，越靠后越在上层）
	els.sort((a, b) => (a.__r ?? 0) - (b.__r ?? 0));
	return els;
}

/* ---------- 占位符继承 / 背景 ---------- */

function phKey(sp) {
	const nvPr = sp["p:nvSpPr"]?.["p:nvPr"];
	if (!isObj(nvPr)) return null;
	for (const ph of toArray(nvPr["p:ph"])) {
		const type = getAttr(ph, "type") || "obj";
		const idx = getAttr(ph, "idx");
		return idx !== undefined ? `idx:${idx}` : `type:${type}`;
	}
	return null;
}

function collectPlaceholders(tree) {
	const map = new Map();
	if (!isObj(tree)) return map;
	const scan = (node) => {
		if (!isObj(node)) return;
		for (const sp of [...toArray(node["p:sp"]), ...toArray(node["p:cxnSp"])]) {
			const key = phKey(sp);
			if (key === null) continue;
			const x = parseXfrm(sp["p:spPr"]?.["a:xfrm"]);
			if (x !== null && !map.has(key)) map.set(key, x);
		}
		for (const g of toArray(node["p:grpSp"])) scan(g);
	};
	scan(tree);
	return map;
}

function isPlaceholderSp(sp) {
	return phKey(sp) !== null;
}

function bgFromCSld(cSld, ctx) {
	const bg = isObj(cSld) ? cSld["p:bg"] : null;
	if (!isObj(bg)) return null;
	const bgPr = isObj(bg["p:bgPr"]) ? bg["p:bgPr"] : null;
	if (bgPr !== null) {
		const f = fillFromPr(bgPr, ctx);
		if (f !== null && f.t !== "none") return f;
	}
	const bgRef = isObj(bg["p:bgRef"]) ? bg["p:bgRef"] : null;
	if (bgRef !== null) {
		const inner = colorFromNode(bgRef, ctx);
		if (inner) return { t: "solid", c: inner.c, a: inner.a };
		const f = refColor(getAttr(bgRef, "idx"), ctx);
		if (f !== null) return f;
	}
	return null;
}

/* ---------- 主流程 ---------- */

/**
 * 构建整个 deck 的渲染模型。
 * @param {string} pptxPath 源文件
 * @param {string} outDir decks/<deckId> 目录（写 render/<n>.json 与 media/*）
 * @returns {Promise<{count:number, media:string[]}>}
 */
export async function buildRenderModels(pptxPath, outDir) {
	const { readFile } = await import("node:fs/promises");
	const zip = await JSZip.loadAsync(await readFile(pptxPath));
	const presDoc = await parseXmlFile(zip, "ppt/presentation.xml");
	const pres = presDoc?.["p:presentation"];
	if (!isObj(pres)) throw new Error("invalid pptx: missing presentation.xml");
	const sldSz = isObj(pres["p:sldSz"]) ? pres["p:sldSz"] : {};
	const slideSize = { cx: num(getAttr(sldSz, "cx"), 12192000), cy: num(getAttr(sldSz, "cy"), 6858000) };
	const presRels = await readRels(zip, "ppt/_rels/presentation.xml.rels");
	const sldIdLst = isObj(pres["p:sldIdLst"]) ? pres["p:sldIdLst"] : {};
	const slidePaths = [];
	for (const sldId of toArray(sldIdLst["p:sldId"])) {
		const rid = getAttr(sldId, "r:id");
		const rel = rid ? presRels.get(rid) : undefined;
		if (rel) slidePaths.push(resolveRel("ppt", rel.Target));
	}

	const renderDir = path.join(outDir, "render");
	const mediaDir = path.join(outDir, "media");
	await mkdir(renderDir, { recursive: true });
	await mkdir(mediaDir, { recursive: true });

	const mediaMap = new Map(); // part → filename
	const pendingXmls = new Map(); // part → promise
	// 返回 { doc, ranks }：ranks 为 cNvPr id → spTree 文档顺序（z-order）序号。
	// XMLParser 按标签名归组会丢失不同标签间的文档顺序，故从原始 XML 扫描顺序。
	const ensureXml = (part) => {
		if (!pendingXmls.has(part)) {
			const f = zip.file(part);
			pendingXmls.set(
				part,
				f ? f.async("string").then((s) => {
					let doc = null;
					try {
						doc = FX.parse(s);
					} catch {
						doc = null;
					}
					const ranks = new Map();
					let r = 0;
					const re = /<p:cNvPr\b[^>]*?\bid="(\d+)"/g;
					for (;;) {
						const m = re.exec(s);
						if (m === null) break;
						if (!ranks.has(m[1])) ranks.set(m[1], (r += 1));
					}
					return { doc, ranks };
				}, () => ({ doc: null, ranks: new Map() })) : Promise.resolve({ doc: null, ranks: new Map() }),
			);
		}
		return pendingXmls.get(part);
	};
	const relsCache = new Map();
	const ensureRels = (part) => {
		const rp = relsPathFor(part);
		if (!relsCache.has(rp)) relsCache.set(rp, readRels(zip, rp));
		return relsCache.get(rp);
	};

	// 预解析所有 slide/layout/master/theme 的 rels 与 xml（两阶段异步）
	const slideInfos = [];
	for (const slidePath of slidePaths) {
		const rels = await ensureRels(slidePath);
		let layoutPath = null;
		for (const rel of rels.values()) if (rel.Type.endsWith("/slideLayout")) layoutPath = resolveRel(dirname(slidePath), rel.Target);
		slideInfos.push({ slidePath, rels, layoutPath });
	}
	const layoutInfos = new Map();
	for (const info of slideInfos) {
		if (info.layoutPath === null || layoutInfos.has(info.layoutPath)) continue;
		const lrels = await ensureRels(info.layoutPath);
		let masterPath = null;
		for (const rel of lrels.values()) if (rel.Type.endsWith("/slideMaster")) masterPath = resolveRel(dirname(info.layoutPath), rel.Target);
		layoutInfos.set(info.layoutPath, { rels: lrels, masterPath });
	}
	const masterInfos = new Map();
	for (const li of layoutInfos.values()) {
		if (li.masterPath === null || masterInfos.has(li.masterPath)) continue;
		const mrels = await ensureRels(li.masterPath);
		let themePath = null;
		for (const rel of mrels.values()) if (rel.Type.endsWith("/theme")) themePath = resolveRel(dirname(li.masterPath), rel.Target);
		masterInfos.set(li.masterPath, { rels: mrels, themePath });
	}
	const themeInfos = new Map();
	for (const mi of masterInfos.values()) {
		if (mi.themePath === null || themeInfos.has(mi.themePath)) continue;
		themeInfos.set(mi.themePath, (await ensureXml(mi.themePath)).doc);
	}
	const slideDocs = new Map();
	for (const info of slideInfos) slideDocs.set(info.slidePath, await ensureXml(info.slidePath));
	const layoutDocs = new Map();
	for (const lp of layoutInfos.keys()) layoutDocs.set(lp, await ensureXml(lp));
	const masterDocs = new Map();
	for (const mp of masterInfos.keys()) masterDocs.set(mp, await ensureXml(mp));

	// 预加载全部 chart part（含各自 rels，供图表内图片注册）
	const chartInfos = new Map();
	const chartParts = [];
	zip.forEach((p) => {
		if (/^ppt\/charts\/chart\d+\.xml$/.test(p)) chartParts.push(p);
	});
	for (const cp of chartParts) chartInfos.set(cp, { doc: (await ensureXml(cp)).doc, rels: await ensureRels(cp) });

	const count = slideInfos.length;
	for (let i = 0; i < count; i += 1) {
		const info = slideInfos[i];
		const n = i + 1;
		const slideEntry = slideDocs.get(info.slidePath);
		const slideDoc = slideEntry?.doc ?? null;
		const layoutInfo = info.layoutPath !== null ? layoutInfos.get(info.layoutPath) : undefined;
		const masterInfo = layoutInfo?.masterPath != null ? masterInfos.get(layoutInfo.masterPath) : undefined;
		const themeDoc = masterInfo?.themePath != null ? themeInfos.get(masterInfo.themePath) : null;
		const masterEntry = layoutInfo?.masterPath != null ? masterDocs.get(layoutInfo.masterPath) : null;
		const masterDoc = masterEntry?.doc ?? null;
		const layoutEntry = info.layoutPath !== null ? layoutDocs.get(info.layoutPath) : null;
		const layoutDoc = layoutEntry?.doc ?? null;

		const theme = parseTheme(themeDoc);
		const clrMap = parseClrMap(masterDoc);
		const slideTree = slideDoc?.["p:sld"]?.["p:cSld"]?.["p:spTree"];
		const layoutTree = layoutDoc?.["sldLayout"]?.["p:cSld"]?.["p:spTree"] ?? layoutDoc?.["p:sldLayout"]?.["p:cSld"]?.["p:spTree"];
		const masterTree = masterDoc?.["sldMaster"]?.["p:cSld"]?.["p:spTree"] ?? masterDoc?.["p:sldMaster"]?.["p:cSld"]?.["p:spTree"];

		const ctx = {
			theme,
			clrMap,
			zip,
			mediaMap,
			rels: info.rels,
			partDir: dirname(info.slidePath),
			slideNo: n,
			charts: chartInfos,
			registerMedia: makeMediaRegistry(zip, info.rels, dirname(info.slidePath), mediaMap),
			inheritXfrm: (() => {
				const lmap = collectPlaceholders(layoutTree);
				const mmap = collectPlaceholders(masterTree);
				return (sp) => {
					const key = phKey(sp);
					if (key === null) return null;
					return lmap.get(key) ?? mmap.get(key) ?? null;
				};
			})(),
			isPlaceholder: isPlaceholderSp,
			ranks: slideEntry?.ranks ?? new Map(),
		};

		let bg = bgFromCSld(slideDoc?.["p:sld"]?.["p:cSld"], ctx);
		if (bg === null && layoutDoc != null) {
			const lctx = { ...ctx, rels: layoutInfo.rels, registerMedia: makeMediaRegistry(zip, layoutInfo.rels, dirname(info.layoutPath), mediaMap) };
			bg = bgFromCSld(layoutDoc?.["sldLayout"]?.["p:cSld"] ?? layoutDoc?.["p:sldLayout"]?.["p:cSld"], lctx);
		}
		if (bg === null && masterDoc != null) {
			const mctx = { ...ctx, rels: masterInfo.rels, registerMedia: makeMediaRegistry(zip, masterInfo.rels, dirname(layoutInfo.masterPath), mediaMap) };
			bg = bgFromCSld(masterDoc?.["sldMaster"]?.["p:cSld"] ?? masterDoc?.["p:sldMaster"]?.["p:cSld"], mctx);
		}
		if (bg === null) bg = { t: "solid", c: theme[clrMap.bg1] ?? "#ffffff", a: 1 };

		const els = [];
		// master 装饰（非占位符形状）
		if (masterTree != null && getAttr(layoutDoc?.["sldLayout"] ?? layoutDoc?.["p:sldLayout"] ?? {}, "showMasterSp") !== "0") {
			const mctx = { ...ctx, ranks: masterEntry?.ranks ?? new Map(), rels: masterInfo.rels, registerMedia: makeMediaRegistry(zip, masterInfo.rels, dirname(layoutInfo.masterPath), mediaMap) };
			els.push(...buildTree(masterTree, mctx, IDENTITY_TF));
		}
		// layout 装饰
		if (layoutTree != null) {
			const lctx = { ...ctx, ranks: layoutEntry?.ranks ?? new Map(), rels: layoutInfo.rels, registerMedia: makeMediaRegistry(zip, layoutInfo.rels, dirname(info.layoutPath), mediaMap) };
			els.push(...buildTree(layoutTree, lctx, IDENTITY_TF));
		}
		// slide 内容（占位符是正文，不过滤）
		els.push(...buildTree(slideTree, { ...ctx, isPlaceholder: null }, IDENTITY_TF));

		for (const el of els) delete el.__r;
		const model = { w: slideSize.cx, h: slideSize.cy, bg, els };
		await writeFile(path.join(renderDir, `${n}.json`), JSON.stringify(model));
	}

	// 提取媒体
	const mediaFiles = [];
	for (const [part, name] of mediaMap.entries()) {
		const f = zip.file(part);
		if (!f) continue;
		const bytes = await f.async("nodebuffer");
		await writeFile(path.join(mediaDir, name), bytes);
		mediaFiles.push(name);
	}

	return { count, media: mediaFiles };
}
