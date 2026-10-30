# @jichong/ppt-search

A [DSH](https://www.npmjs.com/search?q=dsh) plugin that turns PowerPoint files into a searchable knowledge base — with **client-side canvas slide previews** (no server-side image rendering).

中文说明见下方。

## Features

- **Upload `.pptx` decks** — each slide is parsed into text (indexed for search) plus a **render model** JSON.
- **Faithful slide previews in the browser** — the server parses the OOXML (shapes, pictures, gradients, tables, charts, theme colors, group transforms, z-order) into a compact render model in EMU coordinates with pre-resolved colors; the browser draws it with Canvas 2D. No PNGs are ever generated server-side.
- **Full-text search** — BM25 ranking over all slides (`ppt_search` agent tool, HTTP API, and an MCP endpoint).
- **Agent tools** — `ppt_deck_list`, `ppt_search`, `ppt_slide_get`, `ppt_compose` for LLM agents.
- **Deck composition** — build a new `.pptx` by picking slides (by deck id + slide number) from the knowledge base, with optional per-slide title override and find/replace.

## How it works

```
.pptx ──(server: jszip + fast-xml-parser)──▶ render model {w,h,bg,els[]} (EMU)
                                            └▶ slide text ──▶ BM25 index
browser ──(GET /ppt-search/api/render/<deckId>/<n>.json)──▶ Canvas 2D painter
```

The render model carries everything the canvas engine needs: solid/gradient fills, lines, pictures (served from `/ppt-search/api/media/…`), tables, charts, bullets, CJK text layout, rotation and flip.

## Install

Install into a DSH web profile (the plugin is loaded as a cordis host plugin):

```sh
npm install @jichong/ppt-search --prefix ~/.dsh/profiles/web
```

Then restart the DSH web service. The plugin injects into `webServer` and `tools`.

## HTTP API

Base path: `/ppt-search/api`

| Endpoint | Method | Purpose |
| --- | --- | --- |
| `POST /upload` | POST | Upload a `.pptx` (multipart or raw body) |
| `POST /list` | POST | List decks (status: pending/processing/ready/error) |
| `POST /detail` | POST | Deck detail with per-slide text |
| `POST /search` | POST | Full-text search across slides |
| `POST /delete` | POST | Remove a deck |
| `POST /retry` | POST | Re-run parsing for a failed deck |
| `POST /compose` | POST | Compose a new deck from an outline of `{deckId, n}` pages |
| `GET /render/<deckId>/<n>.json` | GET | Slide render model (regenerated on demand if missing) |
| `GET /media/<deckId>/<name>` | GET | Slide media (images extracted from the deck) |

An MCP endpoint is also exposed at `/ppt-search/mcp`.

## Configuration

| Env var | Default | Purpose |
| --- | --- | --- |
| `PPT_SEARCH_DATA` | `~/.dsh/data/ppt-search` | Data directory (decks, render models, media, index) |

## Development

```sh
npm run build        # bundle src/client.js → lib/client.js (canvas engine + React UI)
node test/test-render.mjs   # run test suites (test-*.mjs)
```

`src/client.js` is the client source; `build.js` produces the loadable `lib/client.js`. The `prepare` script runs the build automatically on install/publish.

## License

MIT

---

## 中文说明

**ppt-search** 是一个 DSH 插件：把 PPT 变成可检索的知识库。

- 上传 `.pptx`，逐页抽取文本并建立 BM25 全文索引；
- 幻灯片预览**完全在浏览器 Canvas 上绘制**：服务端只把 pptx 解析成渲染模型 JSON（EMU 坐标、预解析的主题色/渐变/线条/组合变换/层级顺序），浏览器 Canvas 2D 引擎负责绘制，服务端不生成任何图片；
- 提供 HTTP API、MCP 端点和智能体工具（`ppt_search`、`ppt_compose` 等）；
- 支持按"提纲"（deck id + 页码列表）从知识库拼装新的 pptx，可对每页做标题覆盖与文字查找替换。

数据目录默认为 `~/.dsh/data/ppt-search`，可用环境变量 `PPT_SEARCH_DATA` 覆盖。

安装：`npm install @jichong/ppt-search --prefix ~/.dsh/profiles/web`，然后重启 DSH web 服务。
