// ppt-search — 浏览器端 UI（PPT 知识库：上传 / 检索 / 合成）。
// 本文件是 build.js 生成的 module-loader factory 的函数体：React 来自 require("react")，
// 样式注入为独立 <style> 标签。构建：node build.js

const React = require("react");
const { useState, useEffect, useCallback, useRef } = React;
const h = React.createElement;

const API = "/ppt-search/api";
const CSS_TAG_ID = "@jichong/ppt-search/ps.css";
const BASKET_KEY = "ppt-search:basket";

const inject = ["slots"];

// ── styles ───────────────────────────────────────────────────────────────────

const CSS = `
.ps { display:flex; flex-direction:column; height:100%; min-height:0;
  --ps-bg:#ffffff; --ps-fg:#1f2328; --ps-bg2:#f6f8fa; --ps-sf:#ffffff;
  --ps-bd:#d0d7de; --ps-mut:#59636e; --ps-dim:#818b96;
  --ps-blue:#0969da; --ps-green:#1a7f37; --ps-red:#cf222e; --ps-yellow:#9a6700;
  background:var(--ps-bg); color:var(--ps-fg);
  font-family:var(--font-ui,system-ui,sans-serif); font-size:13px; }
/* 深色态跟随 dsh 外观设置（浅色/深色/跟随系统），由 shell 在 body 上落 data-ds-dark-theme；
   不用 prefers-color-scheme，否则“浅色”设置下 OS 深色会误判 */
body[data-ds-dark-theme] .ps { --ps-bg:#0d1117; --ps-fg:#e6edf3; --ps-bg2:#161b22; --ps-sf:#161b22;
  --ps-bd:#30363d; --ps-mut:#9198a1; --ps-dim:#6e7681;
  --ps-blue:#4493f8; --ps-green:#3fb950; --ps-red:#f85149; --ps-yellow:#d29922; }
.ps * { box-sizing:border-box; }
.ps-bar { display:flex; align-items:center; gap:8px; padding:8px 12px;
  border-bottom:1px solid var(--ps-bd); background:var(--ps-bg2); flex:0 0 auto; flex-wrap:wrap; }
.ps-tabs { display:flex; gap:4px; }
.ps-title { font-size:14px; font-weight:700; margin-right:10px; white-space:nowrap; }
.ps-tab { border:1px solid var(--ps-bd); background:var(--ps-sf); color:var(--ps-fg);
  border-radius:6px; padding:4px 12px; font-size:12.5px; cursor:pointer; }
.ps-tab.on { border-color:var(--ps-blue); color:var(--ps-blue); font-weight:600; }
.ps-badge { font-size:11px; padding:1px 8px; border-radius:10px; background:var(--ps-bg2);
  color:var(--ps-mut); border:1px solid var(--ps-bd); }
.ps-badge.ready { color:var(--ps-green); border-color:var(--ps-green); }
.ps-badge.error { color:var(--ps-red); border-color:var(--ps-red); }
.ps-badge.processing, .ps-badge.pending { color:var(--ps-yellow); border-color:var(--ps-yellow); }
.ps-main { flex:1 1 auto; min-height:0; overflow-y:auto; padding:12px; }
.ps-zoom { position:fixed; inset:0; z-index:2147483000; display:flex; flex-direction:column;
  align-items:center; justify-content:center; gap:10px; background:rgba(0,0,0,.72); cursor:zoom-out; }
.ps-zoom canvas { width:min(92vw, calc(84vh * var(--ps-ratio, 1.7778))); border-radius:6px; background:#fff;
  box-shadow:0 12px 48px rgba(0,0,0,.55); }
.ps-zoom-cap { color:#e6edf3; font-size:13px; max-width:90vw; text-align:center; }
.ps-zoom-hint { color:#9aa4af; font-size:11.5px; }
.ps-btn { background:var(--ps-sf); color:var(--ps-fg); border:1px solid var(--ps-bd);
  border-radius:6px; padding:3px 10px; font-size:12px; cursor:pointer; }
.ps-btn:hover { border-color:var(--ps-blue); color:var(--ps-blue); }
.ps-btn.primary { background:var(--ps-blue); border-color:var(--ps-blue); color:#fff; }
.ps-btn.primary:hover { opacity:.9; color:#fff; }
.ps-btn.danger:hover { border-color:var(--ps-red); color:var(--ps-red); }
.ps-btn:disabled { opacity:.5; cursor:default; }
a.ps-btn { text-decoration:none; display:inline-block; }
.ps-input, .ps-select { background:var(--ps-sf); color:var(--ps-fg); border:1px solid var(--ps-bd);
  border-radius:6px; padding:4px 8px; font-size:12.5px; }
.ps-input:focus, .ps-select:focus { outline:none; border-color:var(--ps-blue); }
.ps-err { margin:0 12px 8px; padding:6px 10px; border:1px solid var(--ps-red);
  color:var(--ps-red); border-radius:6px; background:rgba(207,34,46,.08); font-size:12px; }
.ps-ok { margin:0 12px 8px; padding:6px 10px; border:1px solid var(--ps-green);
  color:var(--ps-green); border-radius:6px; background:rgba(26,127,55,.08); font-size:12px; }
.ps-empty { padding:40px 16px; text-align:center; color:var(--ps-dim); }
.ps-card { border:1px solid var(--ps-bd); border-radius:8px; margin-bottom:10px;
  background:var(--ps-sf); overflow:hidden; }
.ps-card-head { display:flex; align-items:center; gap:8px; padding:8px 10px; cursor:pointer; }
.ps-card-head:hover { background:var(--ps-bg2); }
.ps-name { font-weight:600; }
.ps-mut { color:var(--ps-mut); font-size:11.5px; }
.ps-grid { display:grid; grid-template-columns:repeat(auto-fill,minmax(200px,1fr)); gap:10px;
  padding:10px; border-top:1px solid var(--ps-bd); }
.ps-slide { border:1px solid var(--ps-bd); border-radius:6px; overflow:hidden; background:var(--ps-bg2); }
.ps-slide canvas { width:100%; display:block; aspect-ratio:var(--ps-ratio, 1.7778); background:#fff; cursor:zoom-in; }
.ps-slide-foot { display:flex; align-items:center; gap:6px; padding:5px 7px; }
.ps-slide-title { flex:1 1 auto; font-size:11.5px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.ps-hit { display:flex; gap:10px; border:1px solid var(--ps-bd); border-radius:8px;
  background:var(--ps-sf); padding:8px; margin-bottom:8px; }
.ps-hit canvas { width:180px; min-width:180px; aspect-ratio:var(--ps-ratio, 1.7778); cursor:zoom-in;
  border:1px solid var(--ps-bd); border-radius:4px; background:#fff; }
.ps-hit-body { flex:1 1 auto; min-width:0; }
.ps-hit-title { font-weight:600; margin-bottom:2px; }
.ps-snip { color:var(--ps-mut); font-size:12px; margin-top:3px; }
.ps-snip b { color:var(--ps-fg); }
.ps-full { white-space:pre-wrap; font-size:12px; color:var(--ps-mut); margin-top:6px;
  border-top:1px dashed var(--ps-bd); padding-top:6px; max-height:260px; overflow-y:auto; }
.ps-row { display:flex; align-items:center; gap:6px; flex-wrap:wrap; }
.ps-basket-item { border:1px solid var(--ps-bd); border-radius:8px; background:var(--ps-sf);
  padding:8px 10px; margin-bottom:8px; }
.ps-basket-item canvas { width:120px; aspect-ratio:var(--ps-ratio, 1.7778); border:1px solid var(--ps-bd);
  border-radius:4px; background:#fff; float:left; margin-right:10px; cursor:zoom-in; }
.ps-basket-item::after { content:""; display:block; clear:both; }
.ps-field { display:flex; align-items:center; gap:6px; margin-top:4px; font-size:12px; }
.ps-field label { color:var(--ps-mut); min-width:64px; }
.ps-rep { display:flex; gap:6px; margin-top:3px; align-items:center; }
.ps-comp { display:flex; align-items:center; gap:10px; border:1px solid var(--ps-bd);
  border-radius:8px; background:var(--ps-sf); padding:8px 10px; margin-bottom:8px; }
.ps-drop { border:2px dashed var(--ps-bd); border-radius:8px; padding:18px; text-align:center;
  color:var(--ps-dim); margin-bottom:12px; }
.ps-drop.over { border-color:var(--ps-blue); color:var(--ps-blue); }
.ps-a { color:var(--ps-blue); text-decoration:none; }
.ps-a:hover { text-decoration:underline; }
/* ── 对话内工具调用行（tool.call.toolview）── */
.ps-tv { min-width:0; max-width:100%; }
.ps-tv-row { display:flex; align-items:center; gap:6px; min-height:26px; padding:1px 4px; border-radius:6px;
  color:var(--dsw-alias-label-secondary,#59636e); font-size:var(--dsh-content-font-size-secondary,13px); min-width:0; }
.ps-tv-click { cursor:pointer; }
.ps-tv-click:hover { background:rgba(128,148,170,.10); }
.ps-tv-icon { flex:0 0 auto; display:inline-flex; align-items:center; }
.ps-tv-title { flex:0 0 auto; }
.ps-tv-dot { flex:0 0 auto; width:2px; height:2px; border-radius:1px; background:currentColor; opacity:.45; }
.ps-tv-sum { flex:1 1 auto; min-width:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.ps-tv-row[data-state="error"] .ps-tv-sum { color:var(--dsw-alias-state-error-primary,#cf222e); }
.ps-tv-row[data-state="stopped"] .ps-tv-sum { opacity:.7; }
.ps-tv-chev { flex:0 0 auto; display:inline-flex; opacity:.55; }
.ps-tv-body { display:flex; flex-direction:column; gap:8px; margin:4px 0 6px 24px; padding:10px 12px;
  border:1px solid var(--dsw-alias-border-l1,#d0d7de); border-radius:8px;
  background:var(--dsw-alias-bg-container,#f6f8fa); max-width:760px; }
.ps-tv-hits { display:flex; flex-direction:column; gap:8px; }
.ps-tv-hit { display:flex; gap:10px; min-width:0; }
.ps-tv-thumb { flex:0 0 auto; cursor:zoom-in; }
.ps-tv-thumb canvas { width:150px; display:block; aspect-ratio:var(--ps-ratio,1.7778);
  border:1px solid var(--dsw-alias-border-l1,#d0d7de); border-radius:4px; background:#fff; }
.ps-tv-hit-body { flex:1 1 auto; min-width:0; }
.ps-tv-hit-title { font-weight:600; font-size:13px; color:var(--dsw-alias-label-primary,#1f2328);
  overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.ps-tv-hit-meta { font-size:11.5px; color:var(--dsw-alias-label-secondary,#59636e); margin:2px 0 1px; }
.ps-tv-snip { font-size:12px; color:var(--dsw-alias-label-secondary,#59636e); margin-top:2px;
  overflow:hidden; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; }
.ps-tv-big { display:block; cursor:zoom-in; }
.ps-tv-big canvas { width:100%; display:block; aspect-ratio:var(--ps-ratio,1.7778);
  border:1px solid var(--dsw-alias-border-l1,#d0d7de); border-radius:6px; background:#fff; }
.ps-tv-text { white-space:pre-wrap; font-size:12px; line-height:1.55; color:var(--dsw-alias-label-secondary,#59636e);
  max-height:260px; overflow-y:auto; }
.ps-tv-loading { font-size:12px; color:var(--dsw-alias-label-tertiary,#818b96); padding:2px 0; }
.ps-tv-grid { display:grid; grid-template-columns:repeat(auto-fill,minmax(140px,1fr)); gap:8px; }
.ps-tv-grid canvas { width:100%; display:block; aspect-ratio:var(--ps-ratio,1.7778);
  border:1px solid var(--dsw-alias-border-l1,#d0d7de); border-radius:4px; background:#fff; }
.ps-tv-grid-cap { font-size:11px; color:var(--dsw-alias-label-secondary,#59636e); margin-top:2px;
  overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.ps-tv-a { color:var(--dsw-alias-link-primary,#0969da); text-decoration:none; align-self:flex-start; }
.ps-tv-a:hover { text-decoration:underline; }
`;

function ensureCss() {
	if (typeof document === "undefined") return;
	if (document.querySelector(`style[data-plugin-css=${JSON.stringify(CSS_TAG_ID)}]`) !== null) return;
	const tag = document.createElement("style");
	tag.dataset.plugin = "@dsh-plugins/ppt-search";
	tag.dataset.pluginCss = CSS_TAG_ID;
	tag.textContent = CSS;
	document.head.appendChild(tag);
}

// ── transport ────────────────────────────────────────────────────────────────

async function call(endpoint, payload) {
	const res = await fetch(`${API}/${endpoint}`, {
		method: "POST",
		headers: { "content-type": "application/json" },
		body: JSON.stringify(payload ?? {}),
	});
	const body = await res.json().catch(() => null);
	if (!res.ok || body === null || body.ok !== true) throw new Error(body?.error?.message ?? `HTTP ${res.status}`);
	return body.value;
}

async function uploadFile(file) {
	let password;
	for (;;) {
		let q = `name=${encodeURIComponent(file.name)}`;
		if (password !== undefined) q += `&password=${encodeURIComponent(password)}`;
		const res = await fetch(`${API}/upload?${q}`, { method: "POST", body: file });
		const body = await res.json().catch(() => null);
		if (!res.ok || body === null || body.ok !== true) throw new Error(body?.error?.message ?? `HTTP ${res.status}`);
		const v = body.value ?? {};
		if (v.needPassword !== true) return v;
		const msg = v.wrongPassword ? "密码错误，请重新输入该 PPT 的打开密码：" : "该 PPT 已加密，请输入打开密码：";
		const p = window.prompt(msg, password ?? "");
		if (p === null) throw new Error(`已取消上传「${file.name}」（未输入密码）`);
		password = p;
	}
}

function fmtSize(n) {
	if (typeof n !== "number") return "";
	if (n < 1024) return `${n} B`;
	if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
	return `${(n / 1024 / 1024).toFixed(1)} MB`;
}

function fmtTime(ms) {
	if (typeof ms !== "number") return "";
	return new Date(ms).toLocaleString("zh-CN", { hour12: false });
}

const STATUS_LABEL = { pending: "排队中", processing: "解析中", ready: "就绪", error: "失败" };

// ── 渲染引擎：服务端把 pptx 解析成渲染模型 JSON，浏览器用 canvas 2D 绘制 ──────
// 渲染模型：{w,h,bg,els}，坐标单位 EMU（1pt = 12700 EMU），颜色已在服务端解析为 hex+alpha。

const PT = 12700;
const FONT_FALLBACK = '"PingFang SC","Microsoft YaHei","Noto Sans CJK SC",sans-serif';
const CHART_PAL = ["#4472c4", "#ed7d31", "#a5a5a5", "#ffc000", "#5b9bd5", "#70ad47", "#264478", "#9e480e", "#76608a", "#46788a"];
const modelCache = new Map();
const imgCache = new Map();

// canvas 字体尺寸有上限（Chrome 钳制到 10000px），EMU 坐标系下字号会超限导致文本不渲染；
// 文本统一换算到设备像素空间绘制，同时保留旋转/翻转矩阵。
function emuFontPx(fontEMU, s) {
	return fontEMU.replace(/(\d+(?:\.\d+)?)px/, (m, n) => `${Math.max(0.5, parseFloat(n) * s)}px`);
}
function fillTextEMU(g, text, fontEMU, x, y, maxW, spcEMU) {
	if (text === "" || text === null || text === undefined) return;
	const t = g.getTransform();
	const s = Math.hypot(t.a, t.b) || 1;
	g.save();
	g.setTransform(t.a / s, t.b / s, t.c / s, t.d / s, t.a * x + t.c * y + t.e, t.b * x + t.d * y + t.f);
	g.font = emuFontPx(fontEMU, s);
	if (spcEMU && "letterSpacing" in g) g.letterSpacing = `${spcEMU * s}px`;
	if (maxW !== undefined && maxW !== null) g.fillText(text, 0, 0, Math.max(1, maxW * s));
	else g.fillText(text, 0, 0);
	g.restore();
}
function measureEMU(g, text, fontEMU, spcEMU) {
	const t = g.getTransform();
	const s = Math.hypot(t.a, t.b) || 1;
	g.save();
	g.setTransform(1, 0, 0, 1, 0, 0);
	g.font = emuFontPx(fontEMU, s);
	if (spcEMU && "letterSpacing" in g) g.letterSpacing = `${spcEMU * s}px`;
	const w = g.measureText(text).width / s;
	g.restore();
	return w;
}

function loadModel(url) {
	let p = modelCache.get(url);
	if (p === undefined) {
		p = fetch(url)
			.then((r) => (r.ok ? r.json() : null))
			.catch(() => null)
			.then((m) => {
				if (m === null || typeof m.w !== "number" || !Array.isArray(m.els)) { modelCache.delete(url); return null; }
				return m;
			});
		modelCache.set(url, p);
	}
	return p;
}

/** 同步取缓存图片；未加载完则登记回调，加载完成后触发整页重绘。 */
function getImg(src, onReady) {
	let e = imgCache.get(src);
	if (e === undefined) {
		e = { img: null, cb: null };
		const im = new Image();
		im.onload = () => { e.img = im; const cb = e.cb; e.cb = null; if (cb !== null) cb(); };
		im.onerror = () => { imgCache.delete(src); };
		im.src = src;
		imgCache.set(src, e);
	}
	if (e.img === null) e.cb = onReady;
	return e.img;
}

function rgba(c, a) {
	if (typeof c !== "string" || c.charAt(0) !== "#") return "transparent";
	const n = parseInt(c.slice(1), 16);
	if (!Number.isFinite(n)) return "transparent";
	const al = typeof a === "number" ? Math.max(0, Math.min(1, a)) : 1;
	return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${al})`;
}

function safeFont(name) {
	return String(name ?? "").replace(/["'\\;]/g, "");
}

function runFont(r) {
	const sz = ((r.sz ?? 18) * PT).toFixed(2);
	const fam = safeFont(r.f ?? r.ea ?? "");
	return `${r.i ? "italic " : ""}${r.b ? "bold " : ""}${sz}px ${fam !== "" ? `"${fam}",` : ""}${FONT_FALLBACK}`;
}

function roman(n) {
	const map = [[10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];
	let out = "";
	for (const [v, s] of map) while (n >= v) { out += s; n -= v; }
	return out === "" ? "i" : out;
}

function fmtNum(v) {
	if (!Number.isFinite(v)) return "";
	if (Math.abs(v) >= 10000) return String(Math.round(v));
	return String(Math.round(v * 100) / 100);
}

function niceCeil(v) {
	if (!(v > 0)) return 1;
	const p = Math.pow(10, Math.floor(Math.log10(v)));
	for (const m of [1, 2, 2.5, 5, 10]) if (v <= m * p) return m * p;
	return 10 * p;
}

function rrect(g, x, y, w, h, r) {
	r = Math.min(r, w / 2, h / 2);
	g.moveTo(x + r, y);
	g.arcTo(x + w, y, x + w, y + h, r);
	g.arcTo(x + w, y + h, x, y + h, r);
	g.arcTo(x, y + h, x, y, r);
	g.arcTo(x, y, x + w, y, r);
	g.closePath();
}

/** 预置几何路径（adj 未随模型下发，使用 PowerPoint 默认调整值）。 */
function geomPath(g, geom, x, y, w, h) {
	g.beginPath();
	switch (geom) {
		case "ellipse":
			g.ellipse(x + w / 2, y + h / 2, Math.max(0, w / 2), Math.max(0, h / 2), 0, 0, Math.PI * 2);
			break;
		case "roundRect":
			rrect(g, x, y, w, h, Math.min(w, h) * 0.16667);
			break;
		case "triangle":
			g.moveTo(x + w / 2, y); g.lineTo(x + w, y + h); g.lineTo(x, y + h); g.closePath();
			break;
		case "rtTriangle":
			g.moveTo(x, y); g.lineTo(x + w, y + h); g.lineTo(x, y + h); g.closePath();
			break;
		case "diamond":
			g.moveTo(x + w / 2, y); g.lineTo(x + w, y + h / 2); g.lineTo(x + w / 2, y + h); g.lineTo(x, y + h / 2); g.closePath();
			break;
		case "hexagon": {
			const c = w * 0.25;
			g.moveTo(x + c, y); g.lineTo(x + w - c, y); g.lineTo(x + w, y + h / 2); g.lineTo(x + w - c, y + h); g.lineTo(x + c, y + h); g.lineTo(x, y + h / 2); g.closePath();
			break;
		}
		case "parallelogram": {
			const c = Math.min(w * 0.25, h);
			g.moveTo(x + c, y); g.lineTo(x + w, y); g.lineTo(x + w - c, y + h); g.lineTo(x, y + h); g.closePath();
			break;
		}
		case "rightArrow": case "leftArrow": case "upArrow": case "downArrow": {
			if (geom === "rightArrow" || geom === "leftArrow") {
				const bh = h * 0.5;
				const hd = w * 0.5;
				if (geom === "rightArrow") {
					g.moveTo(x, y + (h - bh) / 2); g.lineTo(x + w - hd, y + (h - bh) / 2); g.lineTo(x + w - hd, y); g.lineTo(x + w, y + h / 2); g.lineTo(x + w - hd, y + h); g.lineTo(x + w - hd, y + (h + bh) / 2); g.lineTo(x, y + (h + bh) / 2); g.closePath();
				} else {
					g.moveTo(x + w, y + (h - bh) / 2); g.lineTo(x + hd, y + (h - bh) / 2); g.lineTo(x + hd, y); g.lineTo(x, y + h / 2); g.lineTo(x + hd, y + h); g.lineTo(x + hd, y + (h + bh) / 2); g.lineTo(x + w, y + (h + bh) / 2); g.closePath();
				}
			} else {
				const bw = w * 0.5;
				const hd = h * 0.5;
				if (geom === "upArrow") {
					g.moveTo(x + w / 2, y); g.lineTo(x + w, y + hd); g.lineTo(x + (w + bw) / 2, y + hd); g.lineTo(x + (w + bw) / 2, y + h); g.lineTo(x + (w - bw) / 2, y + h); g.lineTo(x + (w - bw) / 2, y + hd); g.lineTo(x, y + hd); g.closePath();
				} else {
					g.moveTo(x + w / 2, y + h); g.lineTo(x + w, y + h - hd); g.lineTo(x + (w + bw) / 2, y + h - hd); g.lineTo(x + (w + bw) / 2, y); g.lineTo(x + (w - bw) / 2, y); g.lineTo(x + (w - bw) / 2, y + h - hd); g.lineTo(x, y + h - hd); g.closePath();
				}
			}
			break;
		}
		case "pie":
			g.moveTo(x + w / 2, y + h / 2);
			g.arc(x + w / 2, y + h / 2, Math.min(w, h) / 2, 0, Math.PI * 1.5);
			g.closePath();
			break;
		case "line": case "straightConnector1":
			g.moveTo(x, y); g.lineTo(x + w, y + h);
			break;
		default:
			g.rect(x, y, w, h);
	}
}

/** 填充当前路径；返回是否实际填充。 */
function fillEl(g, fill, x, y, w, h, mediaBase, onReady) {
	if (fill === null || fill === undefined || fill.t === "none") return false;
	if (fill.t === "solid") {
		g.fillStyle = rgba(fill.c, fill.a);
		g.fill();
		return true;
	}
	if (fill.t === "grad") {
		let gr;
		if (fill.kind === "radial") {
			gr = g.createRadialGradient(x + w / 2, y + h / 2, 0, x + w / 2, y + h / 2, Math.max(w, h) / 2);
		} else {
			const a = ((fill.ang ?? 90) * Math.PI) / 180; // OOXML：0=左→右，90°=上→下
			const len = Math.abs(w * Math.cos(a)) + Math.abs(h * Math.sin(a));
			const cx = x + w / 2;
			const cy = y + h / 2;
			gr = g.createLinearGradient(cx - (Math.cos(a) * len) / 2, cy - (Math.sin(a) * len) / 2, cx + (Math.cos(a) * len) / 2, cy + (Math.sin(a) * len) / 2);
		}
		for (const s of fill.stops ?? []) gr.addColorStop(Math.max(0, Math.min(1, s.p)), rgba(s.c, s.a));
		g.fillStyle = gr;
		g.fill();
		return true;
	}
	if (fill.t === "img") {
		const im = getImg(mediaBase + (fill.src ?? ""), onReady);
		if (im === null || im.naturalWidth === 0) return false;
		let sx = 0;
		let sy = 0;
		let sw = im.naturalWidth;
		let sh = im.naturalHeight;
		if (fill.crop) { sx = fill.crop[0] * sw; sy = fill.crop[1] * sh; sw *= 1 - fill.crop[0] - fill.crop[2]; sh *= 1 - fill.crop[1] - fill.crop[3]; }
		if (sw <= 0 || sh <= 0) return false;
		const sc = Math.max(w / sw, h / sh);
		const dw = sw * sc;
		const dh = sh * sc;
		g.save();
		g.clip();
		g.drawImage(im, sx, sy, sw, sh, x + (w - dw) / 2, y + (h - dh) / 2, dw, dh);
		g.restore();
		return true;
	}
	return false;
}

function strokeEl(g, ln) {
	if (ln === null || ln === undefined || !ln.c || !(ln.w > 0)) return;
	g.strokeStyle = rgba(ln.c, ln.a);
	g.lineWidth = ln.w;
	if (ln.dash !== null && ln.dash !== undefined && ln.dash !== "solid") {
		const d = Math.max(ln.w, 12700);
		const map = { dash: [4 * d, 3 * d], sysDash: [4 * d, 3 * d], lgDash: [8 * d, 4 * d], dot: [d, 2 * d], sysDot: [d, 2 * d], roundDot: [d, d], dashDot: [4 * d, d, d, d], sysDashDot: [4 * d, d, d, d], lgDashDot: [8 * d, d, d, d], lgDashDotDot: [8 * d, d, d, d, d, d] };
		g.setLineDash(map[ln.dash] ?? [4 * d, 3 * d]);
	}
	g.stroke();
	g.setLineDash([]);
}

/** 换行单元：CJK 逐字、拉丁按词（含尾随空格）。 */
function tokenize(s) {
	const toks = [];
	let cur = "";
	for (const ch of s) {
		if (ch === "\n") { if (cur !== "") { toks.push(cur); cur = ""; } toks.push("\n"); continue; }
		const code = ch.codePointAt(0);
		const cjk = (code >= 0x2e80 && code <= 0x303f) || (code >= 0x3040 && code <= 0x33ff) || (code >= 0x3400 && code <= 0x9fff)
			|| (code >= 0xf900 && code <= 0xfaff) || (code >= 0xff00 && code <= 0xffef) || code >= 0x20000;
		if (cjk) { if (cur !== "") { toks.push(cur); cur = ""; } toks.push(ch); }
		else { cur += ch; if (ch === " ") { toks.push(cur); cur = ""; } }
	}
	if (cur !== "") toks.push(cur);
	return toks;
}

/** 段落 → 行列表（含行高/段前距/项目符号文本），用当前 ctx 量宽。 */
function layoutParas(g, tb, boxW) {
	const lines = [];
	let bulletNo = 0;
	for (const p of tb.paras ?? []) {
		const runs = (p.runs ?? []).filter((r) => r !== null && r !== undefined);
		let maxSz = 10;
		for (const r of runs) maxSz = Math.max(maxSz, r.sz ?? 18);
		const lh = p.lnPts ? p.lnPts * PT : (p.lnSpc ?? 1.2) * maxSz * PT;
		let spc = 0;
		if (typeof p.spcBef === "number") spc = p.spcBef * PT;
		else if (typeof p.spcBefPct === "number") spc = p.spcBefPct * lh;
		let bulletTxt = null;
		if (p.bullet) {
			if (p.bullet.char) bulletTxt = p.bullet.char;
			else {
				bulletNo += 1;
				const t = p.bullet.auto ?? "arabicPeriod";
				const suf = t.includes("Paren") ? ")" : ".";
				if (t.startsWith("alpha")) bulletTxt = String.fromCharCode(96 + bulletNo) + suf;
				else if (t.startsWith("roman")) bulletTxt = roman(bulletNo) + suf;
				else bulletTxt = `${bulletNo}${suf}`;
			}
		} else bulletNo = 0;
		const base = { lh, spc, para: p, bulletTxt };
		if (runs.length === 0) { lines.push(Object.assign(base, { items: [], w: 0, first: true })); continue; }
		const stream = [];
		for (const r of runs) {
			if (r.br === 1 || r.t === "\n") { stream.push({ br: true }); continue; }
			const fnt = runFont(r);
			const spcEMU = (r.spc ?? 0) * PT;
			for (const tk of tokenize(r.t)) {
				if (tk === "\n") stream.push({ br: true });
				else stream.push({ tk, run: r, w: measureEMU(g, tk, fnt, spcEMU) });
			}
		}
		let line = Object.assign({}, base, { items: [], w: 0, first: true });
		const flush = () => { lines.push(line); line = Object.assign({}, base, { items: [], w: 0, first: false }); };
		for (const s of stream) {
			if (s.br === true) { flush(); continue; }
			const indent = line.items.length === 0 ? Math.max(0, (p.marL ?? 0) + Math.min(0, p.indent ?? 0)) : Math.max(0, p.marL ?? 0);
			const avail = tb.wrap === false ? Infinity : boxW - indent;
			if (line.w + s.w > avail && line.items.length > 0) flush();
			const last = line.items[line.items.length - 1];
			if (last !== undefined && last.run === s.run) { last.text += s.tk; last.w += s.w; }
			else line.items.push({ run: s.run, text: s.tk, w: s.w });
			line.w += s.w;
		}
		flush();
		if (lines.length > 0) lines[lines.length - 1].last = 1;
	}
	return lines;
}

function drawText(g, tb, x, y, w, h) {
	const ins = tb.ins ?? [0, 0, 0, 0];
	const bx = x + ins[0];
	const by = y + ins[1];
	const bw = Math.max(0, w - ins[0] - ins[2]);
	const bh = Math.max(0, h - ins[1] - ins[3]);
	const lines = layoutParas(g, tb, bw);
	let total = 0;
	for (const l of lines) total += l.lh + l.spc;
	let cy = tb.anchor === "ctr" ? by + (bh - total) / 2 : tb.anchor === "b" ? by + bh - total : by;
	const prevAlign = g.textAlign;
	const prevBaseline = g.textBaseline;
	g.textBaseline = "middle";
	for (const l of lines) {
		cy += l.spc;
		const p = l.para;
		const marL = Math.max(0, p.marL ?? 0);
		const indent = Math.min(0, p.indent ?? 0);
		let lx = bx + marL;
		if (p.algn === "ctr") lx = bx + (bw - l.w) / 2;
		else if (p.algn === "r") lx = bx + bw - l.w;
		const ty = cy + l.lh / 2;
		if (l.bulletTxt !== null && l.first) {
			const fr = l.items.length > 0 ? l.items[0].run : { sz: 18 };
			g.fillStyle = rgba(p.bullet.c ?? "#000000", 1);
			g.textAlign = "left";
			fillTextEMU(g, l.bulletTxt, runFont({ sz: fr.sz ?? 18, f: p.bullet.f, i: fr.i, b: fr.b }), bx + marL + indent, ty);
		}
		let distGap = 0;
		if (tb.wrap !== false && l.items.length > 0 && (p.algn === "dist" || (p.algn === "just" && !l.last))) {
			let chars = 0;
			for (const it of l.items) chars += Array.from(it.text).length;
			if (chars > 1) distGap = Math.max(0, bw - marL - l.w) / chars;
		}
		for (const it of l.items) {
			const r = it.run;
			g.fillStyle = rgba(r.c ?? "#000000", r.a ?? 1);
			g.textAlign = "left";
			const sz = (r.sz ?? 18) * PT;
			const dy = r.base ? -((r.base / 100) * sz) : 0;
			const fnt = runFont(r);
			const spcEMU = (r.spc ?? 0) * PT;
			const sx = lx;
			if (distGap > 0) {
				for (const ch of Array.from(it.text)) {
					fillTextEMU(g, ch, fnt, lx, ty + dy, null, spcEMU);
					lx += measureEMU(g, ch, fnt, spcEMU) + distGap;
				}
			} else {
				fillTextEMU(g, it.text, fnt, lx, ty + dy, null, spcEMU);
				lx += it.w;
			}
			if ((r.u !== undefined && r.u !== null) || (r.s !== undefined && r.s !== null)) {
				g.strokeStyle = rgba(r.c ?? "#000000", r.a ?? 1);
				g.lineWidth = Math.max(9525, sz * 0.05);
				if (r.u !== undefined && r.u !== null) { g.beginPath(); g.moveTo(sx, ty + dy + sz * 0.42); g.lineTo(sx + it.w, ty + dy + sz * 0.42); g.stroke(); }
				if (r.s !== undefined && r.s !== null) { g.beginPath(); g.moveTo(sx, ty + dy - sz * 0.28); g.lineTo(sx + it.w, ty + dy - sz * 0.28); g.stroke(); }
			}
		}
		cy += l.lh;
	}
	g.textAlign = prevAlign;
	g.textBaseline = prevBaseline;
}

function drawLegend(g, items, x, y, w, h, fs) {
	if (items.length === 0) return;
	const lfont = `${fs * 0.85}px ${FONT_FALLBACK}`;
	let total = 0;
	const widths = items.map((it) => { const tw = measureEMU(g, it.name, lfont); total += fs * 2.3 + tw; return tw; });
	let lx = x + (w - total) / 2;
	const ly = y + h - fs * 0.9;
	items.forEach((it, i) => {
		g.fillStyle = it.color;
		g.fillRect(lx, ly - fs * 0.4, fs * 0.8, fs * 0.8);
		g.fillStyle = "#444444";
		g.textAlign = "left";
		fillTextEMU(g, it.name, lfont, lx + fs * 1.1, ly, w * 0.4);
		lx += fs * 2.3 + widths[i];
	});
}

function drawChart(g, c, x, y, w, h) {
	if (c === null || c === undefined || !Array.isArray(c.series) || c.series.length === 0) return;
	const fs = Math.max(9, Math.min(16, h / PT / 22)) * PT;
	g.save();
	g.fillStyle = "#ffffff";
	g.fillRect(x, y, w, h);
	g.textBaseline = "middle";
	let top = y + fs * 0.8;
	let bottom = y + h - fs * 1.0;
	if (c.title) {
		g.fillStyle = "#333333";
		g.textAlign = "center";
		fillTextEMU(g, c.title, `bold ${fs}px ${FONT_FALLBACK}`, x + w / 2, top + fs * 0.5, w);
		top += fs * 1.3;
	}
	const isPie = c.type === "pie" || c.type === "doughnut";
	const showLegend = c.legend !== false && (isPie || c.series.length > 1);
	if (showLegend) bottom -= fs * 1.5;
	const left = x + fs * 3.4;
	const right = x + w - fs * 0.5;
	const pw = right - left;
	const ph = bottom - top;
	if (pw <= fs || ph <= fs) { g.restore(); return; }
	const allVals = [];
	for (const s of c.series) for (const v of s.vals) if (typeof v === "number") allVals.push(v);
	if (allVals.length === 0) { g.restore(); return; }

	if (isPie) {
		const vals = c.series[0].vals.map((v) => (typeof v === "number" && v > 0 ? v : 0));
		const total = vals.reduce((a, b) => a + b, 0);
		const cx = x + w / 2;
		const cy = top + ph / 2;
		const R = (Math.min(pw, ph) / 2) * 0.92;
		const hole = c.type === "doughnut" ? R * Math.max(0.1, Math.min(0.9, c.hole ?? 0.5)) : 0;
		let a0 = -Math.PI / 2;
		vals.forEach((v, i) => {
			const frac = total > 0 ? v / total : 0;
			if (frac <= 0) return;
			const a1 = a0 + frac * Math.PI * 2;
			g.beginPath();
			g.moveTo(cx + Math.cos(a0) * hole, cy + Math.sin(a0) * hole);
			g.arc(cx, cy, R, a0, a1);
			if (hole > 0) g.arc(cx, cy, hole, a1, a0, true);
			else g.lineTo(cx, cy);
			g.closePath();
			g.fillStyle = CHART_PAL[i % CHART_PAL.length];
			g.fill();
			g.strokeStyle = "#ffffff";
			g.lineWidth = fs * 0.08;
			g.stroke();
			if (frac > 0.05) {
				const am = (a0 + a1) / 2;
				const rr = hole > 0 ? (R + hole) / 2 : R * 0.66;
				g.fillStyle = hole > 0 ? "#333333" : "#ffffff";
				g.textAlign = "center";
				fillTextEMU(g, `${Math.round(frac * 100)}%`, `${fs * 0.8}px ${FONT_FALLBACK}`, cx + Math.cos(am) * rr, cy + Math.sin(am) * rr);
			}
			a0 = a1;
		});
		if (showLegend) {
			const names = c.cats.length > 0 ? c.cats : vals.map((_, i) => `项${i + 1}`);
			drawLegend(g, names.map((t, i) => ({ name: String(t), color: CHART_PAL[i % CHART_PAL.length] })), x, y, w, h, fs);
		}
		g.restore();
		return;
	}

	const n = Math.max(1, c.cats.length);
	const horiz = c.type === "bar" && c.dir === "bar";
	const stacked = c.grouping === "stacked" || c.grouping === "percentStacked";
	let vmax = 0;
	if (stacked) {
		for (let i = 0; i < n; i += 1) {
			let sum = 0;
			for (const s of c.series) sum += Math.max(0, s.vals[i] ?? 0);
			vmax = Math.max(vmax, sum);
		}
	} else for (const v of allVals) vmax = Math.max(vmax, v);
	if (c.grouping === "percentStacked") vmax = 100;
	vmax = niceCeil(vmax || 1);
	const ticks = 4;
	g.lineWidth = Math.max(9525, fs * 0.06);
	const afont = `${fs * 0.78}px ${FONT_FALLBACK}`;
	for (let t = 0; t <= ticks; t += 1) {
		const f = t / ticks;
		if (horiz) {
			const gx = left + pw * f;
			g.strokeStyle = t === 0 ? "#999999" : "#e3e3e3";
			g.beginPath(); g.moveTo(gx, top); g.lineTo(gx, bottom); g.stroke();
			g.fillStyle = "#666666";
			g.textAlign = "center";
			fillTextEMU(g, fmtNum(vmax * f), afont, gx, bottom + fs * 0.8);
		} else {
			const gy = bottom - ph * f;
			g.strokeStyle = t === 0 ? "#999999" : "#e3e3e3";
			g.beginPath(); g.moveTo(left, gy); g.lineTo(right, gy); g.stroke();
			g.fillStyle = "#666666";
			g.textAlign = "right";
			fillTextEMU(g, fmtNum(vmax * f), afont, left - fs * 0.4, gy);
		}
	}
	g.fillStyle = "#666666";
	const catLabel = (i) => String(c.cats[i] ?? i + 1);
	if (horiz) {
		g.textAlign = "right";
		const step = Math.ceil(n / Math.max(1, Math.floor(ph / (fs * 1.4))));
		for (let i = 0; i < n; i += 1) if (i % step === 0) fillTextEMU(g, catLabel(i), afont, left - fs * 0.4, top + ((ph / n) * (i + 0.5)), fs * 3);
	} else if (c.type !== "scatter") {
		g.textAlign = "center";
		const step = Math.ceil(n / Math.max(1, Math.floor(pw / (fs * 3.2))));
		for (let i = 0; i < n; i += 1) if (i % step === 0) fillTextEMU(g, catLabel(i), afont, left + (pw / n) * (i + 0.5), bottom + fs * 0.8, (pw / n) * step);
	}

	if (c.type === "bar") {
		for (let i = 0; i < n; i += 1) {
			if (stacked) {
				let acc = 0;
				const bw = (((horiz ? ph : pw) / n) * 0.62);
				for (const s of c.series) {
					const v = Math.max(0, s.vals[i] ?? 0);
					const f0 = acc / vmax;
					acc += v;
					const f1 = acc / vmax;
					if (v <= 0) continue;
					g.fillStyle = s.color;
					if (horiz) g.fillRect(left + (pw * f0), top + ((ph / n) * (i + 0.5)) - (bw / 2), pw * (f1 - f0), bw);
					else g.fillRect(left + ((pw / n) * (i + 0.5)) - (bw / 2), bottom - (ph * f1), bw, ph * (f1 - f0));
				}
			} else {
				const k = c.series.length;
				const bw = (((horiz ? ph : pw) / n) * 0.72) / k;
				c.series.forEach((s, j) => {
					const v = Math.max(0, s.vals[i] ?? 0);
					if (v <= 0) return;
					const f = v / vmax;
					g.fillStyle = s.color;
					if (horiz) g.fillRect(left, top + ((ph / n) * (i + 0.5)) - ((bw * k) / 2) + (bw * j), pw * f, bw);
					else g.fillRect(left + ((pw / n) * (i + 0.5)) - ((bw * k) / 2) + (bw * j), bottom - (ph * f), bw, ph * f);
				});
			}
		}
	}
	if (c.type === "line" || c.type === "area") {
		c.series.forEach((s) => {
			const pts = [];
			for (let i = 0; i < n; i += 1) {
				const v = typeof s.vals[i] === "number" ? s.vals[i] : 0;
				pts.push([left + ((pw / n) * (i + 0.5)), bottom - ((v / vmax) * ph)]);
			}
			if (pts.length === 0) return;
			if (c.type === "area") {
				g.beginPath();
				g.moveTo(pts[0][0], bottom);
				for (const p of pts) g.lineTo(p[0], p[1]);
				g.lineTo(pts[pts.length - 1][0], bottom);
				g.closePath();
				g.fillStyle = rgba(s.color, 0.55);
				g.fill();
			}
			g.beginPath();
			pts.forEach((p, i) => { if (i === 0) g.moveTo(p[0], p[1]); else g.lineTo(p[0], p[1]); });
			g.strokeStyle = s.color;
			g.lineWidth = fs * 0.12;
			g.stroke();
			g.fillStyle = s.color;
			for (const p of pts) { g.beginPath(); g.arc(p[0], p[1], fs * 0.16, 0, Math.PI * 2); g.fill(); }
		});
	}
	if (c.type === "scatter") {
		const xs = c.cats.map((v) => Number(v)).filter((v) => Number.isFinite(v));
		const xmin = xs.length > 0 ? Math.min(...xs) : 0;
		const xmax = niceCeil(xs.length > 0 ? Math.max(...xs) : 1);
		c.series.forEach((s) => {
			g.fillStyle = s.color;
			s.vals.forEach((v, i) => {
				const xv = Number(c.cats[i]);
				if (!Number.isFinite(xv) || typeof v !== "number") return;
				const px = left + ((xv - xmin) / ((xmax - xmin) || 1)) * pw;
				const py = bottom - ((Math.max(0, v) / vmax) * ph);
				g.beginPath(); g.arc(px, py, fs * 0.18, 0, Math.PI * 2); g.fill();
			});
		});
	}
	if (c.type === "radar" && n > 2) {
		const cxr = left + pw / 2;
		const cyr = top + ph / 2;
		const R = (Math.min(pw, ph) / 2) * 0.78;
		for (let t = 1; t <= ticks; t += 1) {
			g.beginPath();
			for (let i = 0; i <= n; i += 1) {
				const a = -Math.PI / 2 + (((i % n) / n) * Math.PI * 2);
				const rr = (R * t) / ticks;
				if (i === 0) g.moveTo(cxr + Math.cos(a) * rr, cyr + Math.sin(a) * rr);
				else g.lineTo(cxr + Math.cos(a) * rr, cyr + Math.sin(a) * rr);
			}
			g.strokeStyle = "#e3e3e3";
			g.lineWidth = Math.max(9525, fs * 0.05);
			g.stroke();
		}
		c.series.forEach((s) => {
			g.beginPath();
			for (let i = 0; i <= n; i += 1) {
				const a = -Math.PI / 2 + (((i % n) / n) * Math.PI * 2);
				const f = (Math.max(0, s.vals[i % n] ?? 0)) / vmax;
				if (i === 0) g.moveTo(cxr + Math.cos(a) * R * f, cyr + Math.sin(a) * R * f);
				else g.lineTo(cxr + Math.cos(a) * R * f, cyr + Math.sin(a) * R * f);
			}
			g.strokeStyle = s.color;
			g.lineWidth = fs * 0.1;
			g.stroke();
		});
	}
	if (showLegend) {
		drawLegend(g, c.series.map((s, i) => ({ name: s.name !== "" && s.name !== undefined ? s.name : `系列${i + 1}`, color: s.color })), x, y, w, h, fs);
	}
	g.restore();
}

function drawTable(g, el, mediaBase, onReady) {
	const colX = [el.x];
	for (const cw of el.cols ?? []) colX.push(colX[colX.length - 1] + cw);
	const rowY = [el.y];
	for (const rh of el.rows ?? []) rowY.push(rowY[rowY.length - 1] + rh);
	for (const cell of el.cells ?? []) {
		const x0 = colX[Math.min(cell.c, colX.length - 1)] ?? el.x;
		const y0 = rowY[Math.min(cell.r, rowY.length - 1)] ?? el.y;
		const x1 = colX[Math.min(cell.c + cell.cs, colX.length - 1)] ?? x0 + el.cx;
		const y1 = rowY[Math.min(cell.r + cell.rs, rowY.length - 1)] ?? y0 + el.cy;
		const cw = x1 - x0;
		const ch = y1 - y0;
		if (cw <= 0 || ch <= 0) continue;
		if (cell.fill !== undefined && cell.fill !== null) {
			g.beginPath();
			g.rect(x0, y0, cw, ch);
			if (!fillEl(g, cell.fill, x0, y0, cw, ch, mediaBase, onReady)) { g.fillStyle = "#ffffff"; g.fill(); }
		} else { g.fillStyle = "#ffffff"; g.fillRect(x0, y0, cw, ch); }
		const sides = [["l", x0, y0, x0, y1], ["t", x0, y0, x1, y0], ["r", x1, y0, x1, y1], ["b", x0, y1, x1, y1]];
		for (const [side, ax, ay, bx2, by2] of sides) {
			const b = cell.borders !== undefined && cell.borders !== null ? cell.borders[side] : undefined;
			if (b === null) continue;
			g.strokeStyle = b !== undefined ? rgba(b.c, 1) : "#d0d0d0";
			g.lineWidth = b !== undefined ? Math.max(9525, b.w) : Math.max(9525, el.cx / 800);
			g.beginPath(); g.moveTo(ax, ay); g.lineTo(bx2, by2); g.stroke();
		}
		if (cell.text !== undefined && cell.text !== null) {
			g.save();
			g.beginPath(); g.rect(x0, y0, cw, ch); g.clip();
			drawText(g, cell.text, x0, y0, cw, ch);
			g.restore();
		}
	}
}

function drawSlide(g, m, mediaBase, onReady) {
	g.fillStyle = m.bg !== null && m.bg !== undefined && m.bg.c ? rgba(m.bg.c, m.bg.a) : "#ffffff";
	g.fillRect(0, 0, m.w, m.h);
	for (const el of m.els ?? []) {
		if (el === null || el === undefined) continue;
		if (el.k === "tbl") { drawTable(g, el, mediaBase, onReady); continue; }
		if (el.k === "chart") { drawChart(g, el.chart, el.x, el.y, el.cx, el.cy); continue; }
		if (el.k === "ph") {
			g.save();
			g.strokeStyle = "#b8bfc6";
			g.lineWidth = Math.max(12700, el.cx / 250);
			g.setLineDash([el.cx / 30, el.cx / 45]);
			g.strokeRect(el.x, el.y, el.cx, el.cy);
			g.setLineDash([]);
			g.fillStyle = "#8a94a0";
			g.textAlign = "center";
			g.textBaseline = "middle";
			fillTextEMU(g, el.label ?? "", `${Math.max(1, el.cy * 0.18)}px ${FONT_FALLBACK}`, el.x + el.cx / 2, el.y + el.cy / 2);
			g.restore();
			continue;
		}
		g.save();
		const cx = el.x + el.cx / 2;
		const cy = el.y + el.cy / 2;
		if (el.rot) { g.translate(cx, cy); g.rotate((el.rot * Math.PI) / 180); g.translate(-cx, -cy); }
		if (el.fh === 1 || el.fv === 1) { g.translate(cx, cy); g.scale(el.fh === 1 ? -1 : 1, el.fv === 1 ? -1 : 1); g.translate(-cx, -cy); }
		if (el.k === "pic") {
			const im = getImg(mediaBase + (el.src ?? ""), onReady);
			if (im !== null && im !== undefined && im.naturalWidth > 0) {
				let sx = 0;
				let sy = 0;
				let sw = im.naturalWidth;
				let sh = im.naturalHeight;
				if (el.crop) { sx = el.crop[0] * sw; sy = el.crop[1] * sh; sw *= 1 - el.crop[0] - el.crop[2]; sh *= 1 - el.crop[1] - el.crop[3]; }
				if (sw > 0 && sh > 0) g.drawImage(im, sx, sy, sw, sh, el.x, el.y, el.cx, el.cy);
			}
			g.restore();
			continue;
		}
		geomPath(g, el.geom ?? "rect", el.x, el.y, el.cx, el.cy);
		fillEl(g, el.fill, el.x, el.y, el.cx, el.cy, mediaBase, onReady);
		strokeEl(g, el.ln);
		if (el.text !== undefined && el.text !== null) {
			g.save();
			geomPath(g, el.geom ?? "rect", el.x, el.y, el.cx, el.cy);
			g.clip();
			drawText(g, el.text, el.x, el.y, el.cx, el.cy);
			g.restore();
		}
		g.restore();
	}
}

/** 单页幻灯片 canvas：拉取渲染模型 JSON → canvas 2D 绘制；尺寸自适应 + dpr 缩放。 */
function SlideCanvas({ deckId, n, className, style, onClick }) {
	const ref = useRef(null);
	const [ratio, setRatio] = useState(16 / 9);
	const stRef = useRef({ model: null, error: false });
	const drawRef = useRef(null);
	const url = `${API}/render/${deckId}/${n}.json`;
	const mediaBase = `${API}/media/${deckId}/`;

	drawRef.current = () => {
		const cv = ref.current;
		if (cv === null) return;
		const cw = cv.clientWidth;
		const ch = cv.clientHeight;
		if (cw === 0 || ch === 0) return;
		const dpr = window.devicePixelRatio || 1;
		const bw = Math.max(1, Math.round(cw * dpr));
		const bh = Math.max(1, Math.round(ch * dpr));
		if (cv.width !== bw) cv.width = bw;
		if (cv.height !== bh) cv.height = bh;
		const g = cv.getContext("2d");
		g.setTransform(dpr, 0, 0, dpr, 0, 0);
		g.clearRect(0, 0, cw, ch);
		const m = stRef.current.model;
		if (m === null) {
			if (stRef.current.error) {
				g.fillStyle = "#9aa4af";
				g.font = `${Math.max(12, cw / 18)}px ${FONT_FALLBACK}`;
				g.textAlign = "center";
				g.textBaseline = "middle";
				g.fillText("渲染失败", cw / 2, ch / 2);
			}
			return;
		}
		const s = Math.min(cw / m.w, ch / m.h);
		g.translate((cw - m.w * s) / 2, (ch - m.h * s) / 2);
		g.scale(s, s);
		drawSlide(g, m, mediaBase, () => { const f = drawRef.current; if (f !== null) f(); });
	};

	useEffect(() => {
		let alive = true;
		stRef.current = { model: null, error: false };
		const f0 = drawRef.current;
		if (f0 !== null) f0();
		loadModel(url).then((m) => {
			if (!alive) return;
			if (m === null) stRef.current.error = true;
			else { stRef.current.model = m; setRatio(m.w / m.h); }
			const f = drawRef.current;
			if (f !== null) f();
		});
		return () => { alive = false; };
	}, [url]);

	useEffect(() => {
		const cv = ref.current;
		if (cv === null || typeof ResizeObserver === "undefined") return undefined;
		const ro = new ResizeObserver(() => { const f = drawRef.current; if (f !== null) f(); });
		ro.observe(cv);
		return () => ro.disconnect();
	}, []);

	const st = Object.assign({ "--ps-ratio": String(ratio) }, style);
	return h("canvas", { ref, className, style: st, onClick });
}

// ── 知识库面板 ───────────────────────────────────────────────────────────────

function KbPanel({ decks, error, setError, refresh, basket, addToBasket }) {
	const [openId, setOpenId] = useState(null);
	const [details, setDetails] = useState({});
	const [uploading, setUploading] = useState(false);
	const [dragOver, setDragOver] = useState(false);
	const fileRef = useRef(null);

	const openDeck = useCallback(
		async (id) => {
			if (openId === id) return setOpenId(null);
			setOpenId(id);
			if (details[id] === undefined) {
				try {
					const d = await call("detail", { deckId: id });
					setDetails((cur) => ({ ...cur, [id]: d.slides }));
				} catch (e) {
					setError(e.message);
				}
			}
		},
		[openId, details, setError],
	);

	const doUpload = useCallback(
		async (files) => {
			setUploading(true);
			setError(null);
			try {
				for (const f of files) await uploadFile(f);
				await refresh();
			} catch (e) {
				setError(`上传失败: ${e.message}`);
			} finally {
				setUploading(false);
			}
		},
		[refresh, setError],
	);

	const act = useCallback(
		async (fn, confirmMsg) => {
			if (confirmMsg !== undefined && !window.confirm(confirmMsg)) return;
			setError(null);
			try {
				await fn();
				await refresh();
			} catch (e) {
				setError(e.message);
			}
		},
		[refresh, setError],
	);

	return h(
		"div",
		{
			onDragOver: (e) => {
				e.preventDefault();
				setDragOver(true);
			},
			onDragLeave: () => setDragOver(false),
			onDrop: (e) => {
				e.preventDefault();
				setDragOver(false);
				const files = [...(e.dataTransfer?.files ?? [])].filter((f) => /\.pptx$/i.test(f.name));
				if (files.length > 0) doUpload(files);
			},
		},
		h(
			"div",
			{ className: `ps-drop${dragOver ? " over" : ""}` },
			uploading ? "上传中…" : "拖拽 .pptx 文件到此处，或 ",
			h(
				"button",
				{ className: "ps-btn primary", disabled: uploading, onClick: () => fileRef.current?.click() },
				"选择文件",
			),
			h("input", {
				ref: fileRef,
				type: "file",
				accept: ".pptx",
				multiple: true,
				style: { display: "none" },
				onChange: (e) => {
					const files = [...(e.target.files ?? [])];
					e.target.value = "";
					if (files.length > 0) doUpload(files);
				},
			}),
		),
		decks.length === 0
			? h("div", { className: "ps-empty" }, "知识库为空，请上传 pptx 文件")
			: decks.map((d) =>
					h(
						"div",
						{ className: "ps-card", key: d.id },
						h(
							"div",
							{ className: "ps-card-head", onClick: () => openDeck(d.id) },
							h("span", { className: `ps-badge ${d.status}` }, STATUS_LABEL[d.status] ?? d.status),
							h("span", { className: "ps-name" }, d.name),
							h("span", { className: "ps-mut" }, `${d.slideCount ?? 0} 页`),
							h("span", { className: "ps-mut" }, fmtTime(d.uploadedAt)),
							d.error !== undefined ? h("span", { className: "ps-badge error" }, d.error) : null,
							h("span", { style: { flex: "1 1 auto" } }),
							h(
								"a",
								{ className: "ps-btn", href: `${API}/download/${d.id}`, onClick: (e) => e.stopPropagation() },
								"下载原文件",
							),
							d.status === "error"
								? h(
										"button",
										{ className: "ps-btn", onClick: (e) => { e.stopPropagation(); act(() => call("retry", { deckId: d.id })); } },
										"重试",
									)
								: null,
							h(
								"button",
								{
									className: "ps-btn danger",
									onClick: (e) => { e.stopPropagation(); act(() => call("delete", { deckId: d.id }), `确定删除「${d.name}」？`); },
								},
								"删除",
							),
							h("span", { className: "ps-mut" }, openId === d.id ? "收起 ▲" : "展开 ▼"),
						),
						openId === d.id && details[d.id] !== undefined
							? h(
									"div",
									{ className: "ps-grid" },
									details[d.id].map((s) =>
										h(
											"div",
											{ className: "ps-slide", key: s.n },
											h(SlideCanvas, {
												deckId: d.id,
												n: s.n,
												onClick: () => openZoom(d.id, s.n, `${d.name} #${s.n} · ${s.title || "(无标题)"}`),
											}),
											h(
												"div",
												{ className: "ps-slide-foot" },
												h("span", { className: "ps-mut" }, `#${s.n}`),
												h("span", { className: "ps-slide-title", title: s.title }, s.title || "(无标题)"),
												h(
													"button",
													{
														className: "ps-btn",
														onClick: () => addToBasket({ deckId: d.id, deckName: d.name, n: s.n, title: s.title }),
													},
													"＋",
												),
											),
										),
									),
								)
							: null,
					),
				),
	);
}

// ── 检索面板 ─────────────────────────────────────────────────────────────────

function SearchPanel({ decks, setError, basket, addToBasket }) {
	const [query, setQuery] = useState("");
	const [deckId, setDeckId] = useState("");
	const [hits, setHits] = useState(null);
	const [searching, setSearching] = useState(false);
	const [expanded, setExpanded] = useState({});
	const [full, setFull] = useState({});

	const doSearch = useCallback(
		async (q) => {
			const text = (q ?? query).trim();
			if (text === "") return;
			setSearching(true);
			setError(null);
			try {
				const payload = { query: text, limit: 30 };
				if (deckId !== "") payload.deckId = deckId;
				const v = await call("search", payload);
				setHits(v.hits);
			} catch (e) {
				setError(e.message);
			} finally {
				setSearching(false);
			}
		},
		[query, deckId, setError],
	);

	const toggleFull = useCallback(
		async (hit) => {
			const key = hit.key;
			if (expanded[key] === true) return setExpanded((c) => ({ ...c, [key]: false }));
			if (full[key] === undefined) {
				try {
					const d = await call("detail", { deckId: hit.deckId });
					const s = d.slides.find((x) => x.n === hit.n);
					setFull((c) => ({ ...c, [key]: s }));
				} catch (e) {
					setError(e.message);
					return;
				}
			}
			setExpanded((c) => ({ ...c, [key]: true }));
		},
		[expanded, full, setError],
	);

	return h(
		"div",
		null,
		h(
			"div",
			{ className: "ps-row", style: { marginBottom: "10px" } },
			h("input", {
				className: "ps-input",
				style: { flex: "1 1 200px" },
				placeholder: "全文检索（支持中文），回车搜索…",
				value: query,
				onChange: (e) => setQuery(e.target.value),
				onKeyDown: (e) => e.key === "Enter" && doSearch(),
			}),
			h(
				"select",
				{ className: "ps-select", value: deckId, onChange: (e) => setDeckId(e.target.value) },
				h("option", { value: "" }, "全部 deck"),
				decks
					.filter((d) => d.status === "ready")
					.map((d) => h("option", { key: d.id, value: d.id }, d.name)),
			),
			h("button", { className: "ps-btn primary", disabled: searching, onClick: () => doSearch() }, searching ? "检索中…" : "搜索"),
		),
		hits === null
			? h("div", { className: "ps-empty" }, "输入关键词开始检索")
			: hits.length === 0
				? h("div", { className: "ps-empty" }, "没有命中")
				: hits.map((hit) => {
						const s = full[hit.key];
						return h(
							"div",
							{ className: "ps-hit", key: hit.key },
							h(SlideCanvas, {
								deckId: hit.deckId,
								n: hit.n,
								onClick: () => openZoom(hit.deckId, hit.n, `${hit.deckName} #${hit.n} · ${hit.title || "(无标题)"}`),
							}),
							h(
								"div",
								{ className: "ps-hit-body" },
								h(
									"div",
									{ className: "ps-hit-title" },
									`${hit.title || "(无标题)"} `,
									h("span", { className: "ps-mut" }, `— ${hit.deckName} #${hit.n} · 相关度 ${hit.score.toFixed(1)}`),
								),
								(hit.snippets ?? []).map((sn, i) =>
									h("div", { className: "ps-snip", key: i }, h("b", null, sn.field === "title" ? "标题" : sn.field === "notes" ? "备注" : "正文"), ` ${sn.text}`),
								),
								h(
									"div",
									{ className: "ps-row", style: { marginTop: "6px" } },
									h("button", { className: "ps-btn", onClick: () => addToBasket({ deckId: hit.deckId, deckName: hit.deckName, n: hit.n, title: hit.title }) }, "＋ 加入合成篮"),
									h("button", { className: "ps-btn", onClick: () => toggleFull(hit) }, expanded[hit.key] ? "收起全文" : "查看全文"),
								),
								expanded[hit.key] === true && s !== undefined
									? h("div", { className: "ps-full" }, `${s.body}${s.notes ? `\n\n【备注】\n${s.notes}` : ""}`)
									: null,
							),
						);
					}),
	);
}

// ── 合成篮面板 ───────────────────────────────────────────────────────────────

function BasketPanel({ basket, setBasket, setError, setErrorMsg, decks }) {
	const [name, setName] = useState("");
	const [result, setResult] = useState(null);
	const [busy, setBusy] = useState(false);

	const patch = (i, p) => setBasket((cur) => cur.map((it, j) => (j === i ? { ...it, ...p } : it)));
	const move = (i, dir) =>
		setBasket((cur) => {
			const next = [...cur];
			const j = i + dir;
			if (j < 0 || j >= next.length) return cur;
			[next[i], next[j]] = [next[j], next[i]];
			return next;
		});
	const removeAt = (i) => setBasket((cur) => cur.filter((_, j) => j !== i));

	const patchReplace = (i, k, p) =>
		patch(i, { replace: (basket[i].replace ?? []).map((r, j) => (j === k ? { ...r, ...p } : r)) });

	const doCompose = useCallback(async () => {
		setBusy(true);
		setResult(null);
		setError(null);
		try {
			const items = basket.map((it) => {
				const out = { deckId: it.deckId, n: it.n };
				if (it.titleOverride !== undefined && it.titleOverride.trim() !== "") out.titleOverride = it.titleOverride.trim();
				const rep = (it.replace ?? []).filter((r) => r.find !== "" );
				if (rep.length > 0) out.replace = rep;
				return out;
			});
			const v = await call("compose", { items, name: name.trim() === "" ? undefined : name.trim() });
			setResult(v);
			setErrorMsg(`合成成功：${v.name}`);
		} catch (e) {
			setError(e.message);
		} finally {
			setBusy(false);
		}
	}, [basket, name, setError, setErrorMsg]);

	if (basket.length === 0) return h("div", { className: "ps-empty" }, "合成篮为空 — 在知识库或检索结果中点「＋ 加入合成篮」按提纲顺序挑选页面");

	return h(
		"div",
		null,
		basket.map((it, i) =>
			h(
				"div",
				{ className: "ps-basket-item", key: `${it.deckId}-${it.n}-${i}` },
				h(SlideCanvas, {
					deckId: it.deckId,
					n: it.n,
					onClick: () => openZoom(it.deckId, it.n, `${it.deckName} #${it.n} · ${it.title || "(无标题)"}`),
				}),
				h("div", null, h("b", null, `${i + 1}. ${it.title || "(无标题)"}`), h("div", { className: "ps-mut" }, `${it.deckName} #${it.n}`)),
				h(
					"div",
					{ className: "ps-row", style: { marginTop: "4px" } },
					h("button", { className: "ps-btn", disabled: i === 0, onClick: () => move(i, -1) }, "↑"),
					h("button", { className: "ps-btn", disabled: i === basket.length - 1, onClick: () => move(i, 1) }, "↓"),
					h("button", { className: "ps-btn danger", onClick: () => removeAt(i) }, "移除"),
				),
				h(
					"div",
					{ className: "ps-field" },
					h("label", null, "改标题"),
					h("input", {
						className: "ps-input",
						style: { flex: "1 1 auto" },
						placeholder: it.title || "留空则不改",
						value: it.titleOverride ?? "",
						onChange: (e) => patch(i, { titleOverride: e.target.value }),
					}),
				),
				h(
					"div",
					{ className: "ps-field", style: { alignItems: "flex-start" } },
					h("label", null, "文字替换"),
					h(
						"div",
						{ style: { flex: "1 1 auto" } },
						(it.replace ?? []).map((r, k) =>
							h(
								"div",
								{ className: "ps-rep", key: k },
								h("input", { className: "ps-input", placeholder: "查找", value: r.find, onChange: (e) => patchReplace(i, k, { find: e.target.value }) }),
								h("span", { className: "ps-mut" }, "→"),
								h("input", { className: "ps-input", placeholder: "替换为", value: r.replace, onChange: (e) => patchReplace(i, k, { replace: e.target.value }) }),
								h("button", { className: "ps-btn danger", onClick: () => patch(i, { replace: (it.replace ?? []).filter((_, j) => j !== k) }) }, "×"),
							),
						),
						h(
							"button",
							{
								className: "ps-btn",
								onClick: () => patch(i, { replace: [...(it.replace ?? []), { find: "", replace: "" }] }),
							},
							"＋ 加一条替换",
						),
					),
				),
			),
		),
		h(
			"div",
			{ className: "ps-row", style: { marginTop: "10px" } },
			h("input", { className: "ps-input", style: { flex: "1 1 200px" }, placeholder: "输出文件名（自动补 .pptx）", value: name, onChange: (e) => setName(e.target.value) }),
			h("button", { className: "ps-btn primary", disabled: busy, onClick: doCompose }, busy ? "合成中…" : `合成 PPT（${basket.length} 页）`),
			h("button", { className: "ps-btn danger", onClick: () => setBasket([]) }, "清空"),
		),
		result !== null
			? h(
					"div",
					{ className: "ps-ok", style: { marginTop: "8px" } },
					`已生成 ${result.name}（${result.slideCount} 页，${fmtSize(result.size)}）— `,
					h("a", { className: "ps-a", href: result.downloadUrl }, "点击下载"),
				)
			: null,
	);
}

// ── 合成记录面板 ─────────────────────────────────────────────────────────────

function CompsPanel({ setError }) {
	const [items, setItems] = useState(null);

	const load = useCallback(async () => {
		setError(null);
		try {
			const v = await call("compositions", {});
			setItems(v.items);
		} catch (e) {
			setError(e.message);
		}
	}, [setError]);

	useEffect(() => {
		load();
	}, [load]);

	if (items === null) return h("div", { className: "ps-empty" }, "加载中…");
	if (items.length === 0) return h("div", { className: "ps-empty" }, "还没有合成记录");
	return h(
		"div",
		null,
		items.map((c) =>
			h(
				"div",
				{ className: "ps-comp", key: c.id },
				h("b", null, c.name),
				h("span", { className: "ps-mut" }, `${c.slideCount} 页 · ${fmtSize(c.size)} · ${fmtTime(c.mtime)}`),
				h("span", { style: { flex: "1 1 auto" } }),
				h("a", { className: "ps-a", href: c.downloadUrl }, "下载"),
			),
		),
	);
}

// ── 主视图 ───────────────────────────────────────────────────────────────────

function loadBasket() {
	try {
		const raw = localStorage.getItem(BASKET_KEY);
		const arr = raw === null ? [] : JSON.parse(raw);
		return Array.isArray(arr) ? arr : [];
	} catch {
		return [];
	}
}

// ── 预览图灯箱：点击缩略图放大，点击遮罩/Esc 关闭 ────────────────────────────

/** 缩略图 click：广播放大事件。 */
function openZoom(deckId, n, label) {
	document.dispatchEvent(new CustomEvent("ps-zoom", { detail: { deckId, n, label } }));
}

function ZoomOverlay() {
	const [zoom, setZoom] = useState(null);
	useEffect(() => {
		const onZoom = (e) => setZoom(e.detail);
		const onKey = (e) => {
			if (e.key === "Escape") setZoom(null);
		};
		document.addEventListener("ps-zoom", onZoom);
		window.addEventListener("keydown", onKey);
		return () => {
			document.removeEventListener("ps-zoom", onZoom);
			window.removeEventListener("keydown", onKey);
		};
	}, []);
	if (zoom === null) return null;
	return h(
		"div",
		{ className: "ps-zoom", onClick: () => setZoom(null) },
		h(SlideCanvas, { deckId: zoom.deckId, n: zoom.n }),
		zoom.label ? h("div", { className: "ps-zoom-cap" }, zoom.label) : null,
		h("div", { className: "ps-zoom-hint" }, "点击关闭（或按 Esc）"),
	);
}

function View() {
	const [tab, setTab] = useState("kb");
	const [decks, setDecks] = useState([]);
	const [error, setError] = useState(null);
	const [notice, setNotice] = useState(null);
	const [basket, setBasket] = useState(loadBasket);
	const timer = useRef(null);

	useEffect(() => {
		localStorage.setItem(BASKET_KEY, JSON.stringify(basket));
	}, [basket]);

	const refresh = useCallback(async () => {
		try {
			const v = await call("list", {});
			setDecks(v.decks);
		} catch (e) {
			setError(e.message);
		}
	}, []);

	useEffect(() => {
		refresh();
	}, [refresh]);

	// 有 deck 在处理时轮询刷新
	useEffect(() => {
		const busy = decks.some((d) => d.status === "pending" || d.status === "processing");
		if (!busy) return undefined;
		timer.current = setTimeout(refresh, 2500);
		return () => clearTimeout(timer.current);
	}, [decks, refresh]);

	useEffect(() => {
		if (notice === null) return undefined;
		const t = setTimeout(() => setNotice(null), 6000);
		return () => clearTimeout(t);
	}, [notice]);

	const addToBasket = useCallback((item) => {
		setBasket((cur) => [...cur, { ...item, titleOverride: "", replace: [] }]);
		setNotice(`已加入合成篮：${item.title || `第 ${item.n} 页`}（共 ${loadBasket().length + 1} 页）`);
	}, []);

	const basketCount = basket.length;

	return h(
		"div",
		{ className: "ps" },
		h(
			"div",
			{ className: "ps-bar" },
			h("span", { className: "ps-title" }, "PPT 知识库"),
			h(
				"div",
				{ className: "ps-tabs" },
				h("button", { className: `ps-tab${tab === "kb" ? " on" : ""}`, onClick: () => setTab("kb") }, `知识库 (${decks.length})`),
				h("button", { className: `ps-tab${tab === "search" ? " on" : ""}`, onClick: () => setTab("search") }, "全文检索"),
				h("button", { className: `ps-tab${tab === "basket" ? " on" : ""}`, onClick: () => setTab("basket") }, `合成篮 (${basketCount})`),
				h("button", { className: `ps-tab${tab === "comps" ? " on" : ""}`, onClick: () => setTab("comps") }, "合成记录"),
			),
			h("span", { style: { flex: "1 1 auto" } }),
			h("button", { className: "ps-btn", onClick: refresh }, "刷新"),
		),
		error !== null
			? h(
					"div",
					{ className: "ps-err" },
					error,
					h("button", { className: "ps-btn", style: { marginLeft: "8px" }, onClick: () => setError(null) }, "关闭"),
				)
			: null,
		notice !== null ? h("div", { className: "ps-ok" }, notice) : null,
		h(
			"div",
			{ className: "ps-main" },
			tab === "kb" ? h(KbPanel, { decks, error, setError, refresh, basket, addToBasket }) : null,
			tab === "search" ? h(SearchPanel, { decks, setError, basket, addToBasket }) : null,
			tab === "basket" ? h(BasketPanel, { basket, setBasket, setError, setErrorMsg: setNotice, decks }) : null,
			tab === "comps" ? h(CompsPanel, { setError }) : null,
		),
		h(ZoomOverlay),
	);
}

// ── 对话内工具调用预览（tool.call.toolview keyed slot）────────────────────────
//
// DSH 会话流中每个工具调用行通过 renderSlot("tool.call.toolview", owner, {entryKey: toolName,
// fallback: 通用行}) 渲染；插件用 key=工具名 注册组件即可整行替换通用样式。
// 组件 props（来自 ToolCall owner）：{ callId, toolName, phase, block, openFile, cwd, home,
// loadImage, useDisclosure, inspect, sessionId, useSessions, t }。
// block：结果态含 kind/call.argsRaw/content/isError/error；运行态 phase 为 preparing|start。

/** 从工具 block 解析调用参数；流式未完成或解析失败返回 null。 */
function parseToolArgs(block) {
	const raw = "kind" in block ? block.call?.argsRaw ?? "" : block.phase === "start" ? block.argsRaw : null;
	if (typeof raw !== "string" || raw === "") return null;
	try {
		const v = JSON.parse(raw);
		return v !== null && typeof v === "object" ? v : null;
	} catch {
		return null;
	}
}

/** 调用状态：preparing | running | stopped | error | ok */
function toolState(block) {
	if (!("kind" in block)) return block.phase === "preparing" ? "preparing" : "running";
	if (block.error?.code === "interrupted") return "stopped";
	return block.isError ? "error" : "ok";
}

/** 工具结果的全部 text 内容拼接。 */
function toolResultText(block) {
	if (!("kind" in block)) return "";
	const c = block.content;
	if (!Array.isArray(c)) return "";
	return c.map((p) => (p?.type === "text" ? p.text ?? "" : "")).filter((s) => s !== "").join("\n");
}

function firstLine(s) {
	const i = s.indexOf("\n");
	return (i < 0 ? s : s.slice(0, i)).trim();
}

/**
 * 展开状态：本地 state，默认展开（结果预览需保留在最终会话中可见）。
 * 不再复用宿主 useDisclosure——宿主会在回合结束/新消息时自动收起，
 * 导致预览"一闪而过"。用户仍可点击工具行手动收起。
 */
function useToolExpand() {
	const [own, setOwn] = useState(true);
	return { expanded: own, toggle: () => setOwn((v) => !v) };
}

function tvSvg(paths) {
	return h(
		"svg",
		{
			width: 15,
			height: 15,
			viewBox: "0 0 24 24",
			fill: "none",
			stroke: "currentColor",
			strokeWidth: 1.8,
			strokeLinecap: "round",
			strokeLinejoin: "round",
			"aria-hidden": "true",
		},
		...paths,
	);
}
const ICON_SEARCH = tvSvg([h("circle", { cx: 11, cy: 11, r: 7 }), h("path", { d: "M20.5 20.5L16.7 16.7" })]);
const ICON_PAGE = tvSvg([
	h("path", { d: "M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" }),
	h("path", { d: "M14 2v6h6" }),
	h("path", { d: "M8.5 13h7M8.5 17h7" }),
]);
const ICON_COMPOSE = tvSvg([
	h("rect", { x: 3, y: 3.5, width: 18, height: 12, rx: 2 }),
	h("path", { d: "M12 15.5V21" }),
	h("path", { d: "M8.5 21h7" }),
	h("path", { d: "M7 11.5l3-3 2.5 2.5L17 7" }),
]);

function Chevron({ up }) {
	return h(
		"svg",
		{
			width: 12,
			height: 12,
			viewBox: "0 0 24 24",
			fill: "none",
			stroke: "currentColor",
			strokeWidth: 2,
			strokeLinecap: "round",
			strokeLinejoin: "round",
			"aria-hidden": "true",
		},
		h("path", { d: up ? "M18 15l-6-6-6 6" : "M6 9l6 6 6-6" }),
	);
}

/** 工具行外壳：图标 + 标题 + 摘要 + 状态 + 可选展开体。 */
function ToolRow({ icon, title, summary, state, expandable, expanded, onToggle, children }) {
	const interactive = expandable === true;
	return h(
		"div",
		{ className: "ps-tv" },
		h(
			"div",
			{
				className: "ps-tv-row" + (interactive ? " ps-tv-click" : ""),
				"data-state": state,
				role: interactive ? "button" : undefined,
				tabIndex: interactive ? 0 : undefined,
				"aria-expanded": interactive ? expanded : undefined,
				onClick: interactive ? onToggle : undefined,
				onKeyDown: interactive
					? (e) => {
							if (e.key !== "Enter" && e.key !== " ") return;
							e.preventDefault();
							onToggle();
						}
					: undefined,
			},
			h("span", { className: "ps-tv-icon" }, icon),
			h("span", { className: "ps-tv-title" }, title),
			h("span", { className: "ps-tv-dot" }),
			h("span", { className: "ps-tv-sum" }, summary),
			interactive ? h("span", { className: "ps-tv-chev" }, h(Chevron, { up: expanded })) : null,
		),
		interactive && expanded ? h("div", { className: "ps-tv-body" }, children) : null,
	);
}

/** 全屏放大浮层（复用 .ps-zoom 样式）。 */
function TvZoom({ zoom, onClose }) {
	if (zoom === null) return null;
	return h(
		"div",
		{ className: "ps-zoom", onClick: onClose },
		h(SlideCanvas, { deckId: zoom.deckId, n: zoom.n }),
		h("div", { className: "ps-zoom-cap" }, zoom.label ?? ""),
		h("div", { className: "ps-zoom-hint" }, "点击任意处关闭"),
	);
}

/** 按调用参数重新查询结构化命中（与服务端工具逻辑一致，结果确定性相同）。 */
async function fetchHits(args) {
	const limit = Math.min(Number(args.limit) || 10, 50);
	const query = typeof args.query === "string" ? args.query.trim() : "";
	if (query === "" && typeof args.deckId === "string" && args.deckId !== "") {
		const d = await call("detail", { deckId: args.deckId });
		const hits = (d.slides ?? []).map((s) => ({
			deckId: args.deckId,
			deckName: d.deck?.name ?? args.deckId,
			n: s.n,
			title: s.title,
			score: 0,
			snippets: [{ field: "body", text: String(s.body ?? "").slice(0, 120) }],
		}));
		return { total: hits.length, hits: hits.slice(0, limit) };
	}
	const payload = { query, limit };
	if (typeof args.deckId === "string" && args.deckId !== "") payload.deckId = args.deckId;
	const v = await call("search", payload);
	return { total: v.total ?? 0, hits: v.hits ?? [] };
}

/** ppt_search 展开体：命中卡片（缩略图 + 标题 + 元信息 + 摘要），点缩略图放大。 */
function SearchHits({ args }) {
	const [st, setSt] = useState({ loading: true, error: "", total: 0, hits: [] });
	const [zoom, setZoom] = useState(null);
	const q = typeof args.query === "string" ? args.query : "";
	const dk = typeof args.deckId === "string" ? args.deckId : "";
	const lim = args.limit;
	useEffect(() => {
		let alive = true;
		setSt({ loading: true, error: "", total: 0, hits: [] });
		fetchHits({ query: q, deckId: dk, limit: lim }).then(
			(v) => {
				if (alive) setSt({ loading: false, error: "", total: v.total, hits: v.hits });
			},
			(e) => {
				if (alive) setSt({ loading: false, error: String(e?.message ?? e), total: 0, hits: [] });
			},
		);
		return () => {
			alive = false;
		};
	}, [q, dk, lim]);
	if (st.loading) return h("div", { className: "ps-tv-loading" }, "加载预览…");
	if (st.error !== "") return h("div", { className: "ps-tv-loading", style: { color: "var(--ps-red,#cf222e)" } }, st.error);
	if (st.hits.length === 0) return h("div", { className: "ps-tv-loading" }, "没有命中");
	return h(
		"div",
		{ className: "ps-tv-hits" },
		st.hits.map((hit) =>
			h(
				"div",
				{ className: "ps-tv-hit", key: `${hit.deckId}#${hit.n}` },
				h(
					"span",
					{
						className: "ps-tv-thumb",
						title: "点击放大",
						onClick: () =>
							setZoom({
								deckId: hit.deckId,
								n: hit.n,
								label: `${hit.deckName ?? hit.deckId} · 第 ${hit.n} 页 · ${hit.title ?? ""}`,
							}),
					},
					h(SlideCanvas, { deckId: hit.deckId, n: hit.n }),
				),
				h(
					"div",
					{ className: "ps-tv-hit-body" },
					h("div", { className: "ps-tv-hit-title" }, hit.title ?? `第 ${hit.n} 页`),
					h(
						"div",
						{ className: "ps-tv-hit-meta" },
						`${hit.deckName ?? hit.deckId} · 第 ${hit.n} 页${typeof hit.score === "number" && hit.score > 0 ? ` · 相关度 ${hit.score.toFixed(2)}` : ""}`,
					),
					(hit.snippets ?? []).slice(0, 2).map((s, i) =>
						h("div", { className: "ps-tv-snip", key: i }, s.field === "body" ? s.text : `[${s.field}] ${s.text}`),
					),
				),
			),
		),
		h(TvZoom, { zoom, onClose: () => setZoom(null) }),
	);
}

/** ppt_search 工具行。 */
function PptSearchRow({ block, useDisclosure }) {
	const args = parseToolArgs(block) ?? {};
	const state = toolState(block);
	const text = toolResultText(block);
	const { expanded, toggle } = useToolExpand(useDisclosure);
	const query = typeof args.query === "string" ? args.query.trim() : "";
	let summary;
	if (state === "error" || state === "stopped") summary = firstLine(text) || (state === "stopped" ? "已中断" : "检索失败");
	else if (state === "ok") {
		const m = /共 (\d+) 条命中/.exec(text);
		summary = (query !== "" ? `检索「${query}」` : `${args.deckId ?? ""} 页目录`) + (m !== null ? ` · 命中 ${m[1]} 页` : "");
	} else summary = "检索中…";
	const expandable = state === "ok";
	return h(
		ToolRow,
		{ icon: ICON_SEARCH, title: "PPT检索", summary, state, expandable, expanded, onToggle: toggle },
		expandable && expanded ? h(SearchHits, { args }) : null,
	);
}

/** ppt_slide_get 工具行：大图预览 + 结果全文。 */
function PptSlideGetRow({ block, useDisclosure }) {
	const args = parseToolArgs(block) ?? {};
	const state = toolState(block);
	const text = toolResultText(block);
	const { expanded, toggle } = useToolExpand(useDisclosure);
	const [zoom, setZoom] = useState(null);
	const deckId = typeof args.deckId === "string" ? args.deckId : "";
	const n = Number(args.n);
	let summary;
	if (state === "error" || state === "stopped") summary = firstLine(text) || (state === "stopped" ? "已中断" : "读取失败");
	else if (state === "ok") {
		const nameM = /^# (.+?) 第 (\d+) 页/.exec(text);
		const titleM = /^标题: (.*)$/m.exec(text);
		summary = `${nameM?.[1] ?? deckId} 第 ${n} 页${titleM !== null ? ` · ${titleM[1]}` : ""}`;
	} else summary = "读取中…";
	const expandable = state === "ok" && deckId !== "" && Number.isFinite(n);
	return h(
		ToolRow,
		{ icon: ICON_PAGE, title: "PPT页", summary, state, expandable, expanded, onToggle: toggle },
		expandable && expanded
			? h(
					"div",
					{ style: { display: "flex", flexDirection: "column", gap: "8px" } },
					h(
						"span",
						{
							className: "ps-tv-big",
							title: "点击放大",
							onClick: () => setZoom({ deckId, n, label: summary }),
						},
						h(SlideCanvas, { deckId, n }),
					),
					h("div", { className: "ps-tv-text" }, text),
					h(TvZoom, { zoom, onClose: () => setZoom(null) }),
				)
			: null,
	);
}

/** ppt_compose 工具行：合成页缩略图网格 + 下载链接。 */
function PptComposeRow({ block, useDisclosure }) {
	const args = parseToolArgs(block) ?? {};
	const state = toolState(block);
	const text = toolResultText(block);
	const { expanded, toggle } = useToolExpand(useDisclosure);
	const items = Array.isArray(args.items) ? args.items : [];
	let summary;
	if (state === "error" || state === "stopped") summary = firstLine(text) || (state === "stopped" ? "已中断" : "合成失败");
	else if (state === "ok") {
		const nameM = /^已合成 (.+?)（/.exec(text);
		summary = `${nameM?.[1] ?? args.name ?? "新文件"} · ${items.length} 页`;
	} else summary = "合成中…";
	const urlM = /下载: (\S+)/.exec(text);
	const expandable = state === "ok" && items.length > 0;
	return h(
		ToolRow,
		{ icon: ICON_COMPOSE, title: "PPT合成", summary, state, expandable, expanded, onToggle: toggle },
		expandable && expanded
			? h(
					"div",
					{ style: { display: "flex", flexDirection: "column", gap: "8px" } },
					h(
						"div",
						{ className: "ps-tv-grid" },
						items.map((it, i) =>
							h(
								"div",
								{ className: "ps-tv-grid-item", key: i },
								h(SlideCanvas, { deckId: it.deckId, n: Number(it.n) }),
								h("div", { className: "ps-tv-grid-cap" }, `${i + 1}. ${it.titleOverride ?? `${it.deckId}#${it.n}`}`),
							),
						),
					),
					urlM !== null ? h("a", { className: "ps-tv-a", href: urlM[1], target: "_blank", rel: "noreferrer" }, "下载 pptx") : null,
				)
			: null,
	);
}

// ── entry ────────────────────────────────────────────────────────────────────

/** 左侧面板列表图标：演示屏 + 折线。 */
function PanelIcon({ size = 18 }) {
	return h(
		"svg",
		{
			width: size,
			height: size,
			viewBox: "0 0 24 24",
			fill: "none",
			stroke: "currentColor",
			strokeWidth: 1.8,
			strokeLinecap: "round",
			strokeLinejoin: "round",
			"aria-hidden": "true",
		},
		h("rect", { x: 3, y: 3.5, width: 18, height: 12, rx: 2 }),
		h("path", { d: "M12 15.5V21" }),
		h("path", { d: "M8.5 21h7" }),
		h("path", { d: "M7 11.5l3-3 2.5 2.5L17 7" }),
	);
}

function apply(ctx) {
	ensureCss();
	// 左侧菜单面板：图标行（位于"新会话"下方、工作区上方）
	ctx.slots.inject("sidebar.panellist", () =>
		ctx.slots.register(
			{
				name: "sidebar.panellist",
				id: "ppt-search",
				order: 20,
				label: "PPT知识库",
			},
			PanelIcon,
		),
	);
	// 面板内容：选中该面板时渲染到 main 区域
	ctx.slots.inject("main", () =>
		ctx.slots.register(
			{
				name: "main",
				key: "ppt-search",
			},
			View,
		),
	);
	// 会话流工具调用行：key=工具名，整行替换通用样式，展开后显示幻灯片预览
	ctx.slots.inject("tool.call.toolview", () =>
		ctx.slots.register(
			{
				name: "tool.call.toolview",
				key: "ppt_search",
			},
			PptSearchRow,
		),
	);
	ctx.slots.inject("tool.call.toolview", () =>
		ctx.slots.register(
			{
				name: "tool.call.toolview",
				key: "ppt_slide_get",
			},
			PptSlideGetRow,
		),
	);
	ctx.slots.inject("tool.call.toolview", () =>
		ctx.slots.register(
			{
				name: "tool.call.toolview",
				key: "ppt_compose",
			},
			PptComposeRow,
		),
	);
}

module.exports = { name: "ppt-search", inject, apply };
