# PPT-Search 插件架构契约（团队共享）

插件 id: `ppt-search`，包名 `@dsh-plugins/ppt-search`，仓库根目录即插件目录。

目标：PPT 知识库管理 —— 上传 .pptx → 逐页提取文本 + 生成渲染模型（render model JSON）→ 全文检索（Web 页面 + Agent 工具 + MCP 端点）→ 按内容提纲组合拼接新 PPT 并做文本修改。

## 运行时约定

- Node ≥ 24，纯 ESM（`"type": "module"`）。宿主是 cordis 插件模型：入口导出 `name` / `inject` / `apply(ctx)`。
- 依赖只允许纯 JS：`jszip`（^3.10）、`fast-xml-parser`（^5.x）。不要引入原生编译依赖。
- 无外部命令依赖：预览渲染为纯 JS 管线（服务端解析 pptx → 渲染模型 JSON，浏览器 canvas 2D 绘制），不使用 soffice/pdftoppm。
- 测试夹具：本机有 `python3 -m pptx`（python-pptx 1.0.2），可用 `test/make-fixture.py` 生成样例 pptx；测试脚本 `test/test-*.mjs` 用 `node` 直接跑，输出 `PASS`/`FAIL` 并以退出码表达结果。
- 所有模块**不得**访问 `process.env` 之外的全局宿主状态；纯函数优先，便于离线测试。

## 数据目录（dataDir）

默认 `path.join(os.homedir(), ".dsh", "data", "ppt-search")`，可用环境变量 `PPT_SEARCH_DATA` 覆盖。

```
<dataDir>/
├── library.json                  # { decks: { [deckId]: Deck } }
├── index.json                    # 序列化检索索引
├── decks/<deckId>/
│   ├── source.pptx               # 原始上传文件
│   ├── meta.json                 # Deck（同 library 中的条目）
│   ├── render/<n>.json           # 第 n 页渲染模型（EMU 坐标，颜色已解析为 hex+alpha）
│   └── media/<file>              # 从包内提取的图片（扁平文件名，已去重/净化）
└── compositions/<id>.pptx        # 合成输出
└── tmp/<uuid>/                   # 处理中的临时目录
```

## 核心数据结构

```js
// Deck
{
  id: "deck-<base36ts>-<rand4>",
  name: "原始文件名.pptx",
  addedAt: 1697456400000,          // ms epoch
  status: "pending" | "processing" | "ready" | "error",
  error: null | "错误信息",
  slideCount: 12,
  slideSize: { cx: 12192000, cy: 6858000 },   // EMU
  slides: [ { n: 1, title: "标题文本", charCount: 350 } ],
  updatedAt: 1697456400000
}

// SlideDoc（检索索引输入）
{
  key: "<deckId>#<n>",
  deckId: "deck-...", n: 3,
  title: "标题", body: "正文纯文本（段落间用 \n 连接）", notes: "演讲备注"
}

// 检索命中
{
  key, deckId, n, title, score,
  deckName: "文件名.pptx",
  snippets: [ { field: "title"|"body"|"notes", text: "…含关键词的上下文片段…" } ],  // ≤3 条
  renderUrl: "/ppt-search/api/render/<deckId>/<n>.json"
}
```

## 模块接口（写作用域互不重叠）

### A. `lib/pptx/parse.js` —— 负责人：pptx-engine

```js
export async function parsePptx(filePath) -> {
  slideSize: { cx, cy },
  slideCount,
  slides: [ { n, title, body, notes } ]   // n 从 1 开始，按 presentation.xml 的 sldIdLst 顺序
}
```

要求：
- JSZip 解压；`fast-xml-parser`（`preserveOrder: false` 亦可，但需处理单元素数组归一化，建议 `parseTagValue: false`）。
- 顺序以 `ppt/presentation.xml` 的 `<p:sldIdLst>` 中 rId 经 `ppt/_rels/presentation.xml.rels` 映射到 slide part 为准（不要用文件名数字排序，那可能错序）。
- 每页文本：遍历 `<p:sp>` 的 `<p:txBody>`，段落 `<a:p>` 内 `<a:r>/<a:t>` 拼接为段落，段落间 `\n`；`<a:tbl>` 表格按行 `\t` 分隔、行间 `\n`；图形分组 `<p:grpSp>` 递归。
- title = 第一个 `<p:ph type="title"|"ctrTitle">` 的文本（去首尾空白）；无则取第一段非空文本前 80 字符。
- notes = `ppt/notesSlides/notesSlideN.xml`（经 slide 的 rels 找 notesSlide）中 body placeholder 文本；无则 ""。
- 忽略母版/版式中的文本（只提取 slide 自身 part）。
- 导出辅助：`export function extractSlideTexts(slideXml, parsed)` 可选，内部使用即可。
- 健壮性：缺 notes、空 slide、无 title 都不能抛错；非 zip / 缺 presentation.xml 抛 `Error("invalid pptx: …")`。

### B. `lib/pptx/compose.js` —— 负责人：pptx-engine

```js
export async function composePptx({
  items,        // [{ deckId, n, titleOverride?: string, replace?: [{find, replace}] }]
  outPath,      // 输出 .pptx 绝对路径
  decksRoot,    // <dataDir>/decks 绝对路径；每个 deck 目录含 source.pptx
  baseDeckId,   // 可选：以哪个 deck 的包为基底（版式/母版/主题）；缺省取 items[0].deckId
}) -> { slideCount, outPath }
```

实现路线（JSZip 克隆法）：
1. 以 base deck 的 source.pptx 为基底打开 JSZip 实例。
2. 删除基底中所有原幻灯片（presentation.xml 的 sldIdLst 清空、删 slide part 及其 rels 中不再被引用的项——**保留** layouts/masters/theme/media，简单起见可不物理删除 slide part，只断开引用；但新克隆的 slide 必须用**未占用**的 part 名）。
3. 对每个 item：打开其来源 deck 的 source.pptx，取第 n 页对应 slide part：
   - 递归收集该 slide 的 `.rels` 引用闭包（layout → master → theme、media 图片、tags 等；notesSlide 可丢弃并删 rels 中该关系）。
   - 跨 deck 冲突处理：目标包中若已有同名 part（如 `slides/slide5.xml`），克隆到新的空闲名字（`slide900+.xml` 递增）。layout/master/theme/media 按「源包内相对路径」映射到目标包的新名字，同一源 deck 同一 part 只复制一次（Map 去重）；不同 deck 的同名 part 视为不同内容，各自复制。
   - 重写克隆 slide 及其依赖 part 的 `.rels` 中的 Target 指向新名字（`../slideLayouts/slideLayout3.xml` 这类相对路径要保持相对语义）。
   - 在目标 `[Content_Types].xml` 中为新 part 补 Override（slide 的 contentType 用源包里该 slide 原 Override 的值；layout/master/theme/media 同理，Default 扩展名已有的跳过）。
4. 文本修改（在克隆出的 slide XML 字符串上做，作用于 `<a:t>` 文本节点）：
   - `titleOverride`：替换 title placeholder（ph type title/ctrTitle）所在 txBody 的**第一个非空 `<a:t>`** 为覆盖文本，清空同 txBody 其余 `<a:t>`。
   - `replace`: 对整页所有 `<a:t>` 做字面 find→replace（区分大小写，全局）。
   - 修改 XML 文本时必须做 XML 转义（& < >），且不得破坏其余结构；用字符串级替换 `<a:t>…</a:t>` 内容。
5. 追加 `presentation.xml` 的 `<p:sldIdLst>`（新 rId 从现有最大 rId+1，sldId 用不冲突的 256+ 递增），`presentation.xml.rels` 增加对应 Relationship，`[Content_Types].xml` 补 Override。
6. 输出 buffer 写 outPath（mkdir -p）。
7. 尺寸不一致的 deck：保持 base 的 sldSz，不做缩放（README 说明限制）。

测试：用 python-pptx 生成 2 个 deck（各 3 页、含图片可选），compose 后：
- 用 parsePptx 解析输出，断言页数、顺序、标题、replace 生效；
- 断言输出可被 JSZip 打开且 `parsePptx` 解析出预期页数（证明包合法，纯 JS 校验，不依赖 soffice）。

### C. `lib/search.js` —— 负责人：search-store

```js
export function createIndex() -> Index            // 内存对象
export function addSlide(index, slideDoc) -> void // 覆盖同 key 旧文档
export function removeDeck(index, deckId) -> void
export function search(index, query, opts) -> [hit]
   // opts: { deckId?, limit = 20 }
export function serialize(index) -> plainObject
export function deserialize(obj) -> Index
export function tokenize(text) -> [string]
```

要求：
- 分词：拉丁字母/数字词（小写、≥1 字符）；CJK 字符逐字 unigram + 相邻双字 bigram；下划线/连字符视为分隔。
- 打分：BM25（k1=1.2, b=0.75），字段加权 title×3、body×1、notes×0.5（字段独立算长度归一）。多查询词命中求和；短语精确命中（query 原文连续出现在 title/body/notes）额外 +10 分并把该完整句作为首条 snippet。
- snippets：对每个命中字段取第一个命中词位置 ±60 字符窗口（词边界对齐），`…` 包裹；最多 3 条，按 title>body>notes 排序。
- 空 query 返回 []；query 全为停用词（的/了/是/a/the/of 等小表）时退化为按 title 前缀匹配。
- serialize/deserialize 必须 round-trip 后搜索结果一致；索引可 JSON.stringify。

### D. `lib/store.js` —— 负责人：search-store

```js
export function createStore(dataDir) -> {
  dataDir,
  ensureDirs(),                       // mkdir -p 全部目录
  loadLibrary() -> { decks: {} },     // 缺文件返回空
  saveLibrary(lib),                   // 原子写（tmp+rename）
  readDeckMeta(deckId) -> Deck|null,
  writeDeckMeta(deck),                // 原子写 decks/<id>/meta.json
  deckDir(deckId), sourcePath(deckId), renderPath(deckId, n), mediaPath(deckId, name),
  compositionPath(id), tmpDir(),      // tmpDir 返回新建的唯一临时目录绝对路径
  genId(prefix) -> `${prefix}-${base36now}-${rand4}`,
  removeDeckFiles(deckId), removeComposition(id),
  listCompositions() -> [ { id, file, size, mtime } ],
}
```

- deckId 校验：只允许 `[a-z0-9-]`，防路径穿越；所有对外 API 传入的 deckId 都要过 `assertId`。
- 原子写：同目录 `.tmp` 文件 + `fs.rename`。

### E. `src/client.js` —— 负责人：client-ui

浏览器侧，module-loader 工厂体（`require("react")` 可用；无 JSX，全部 `React.createElement`）。
入口：`module.exports = { name: "ppt-search", inject: ["slots"], apply }`，
`apply(ctx)` 里 `ctx.slots.inject("conversation.view", () => ctx.slots.register({ name:"conversation.view", id:"ppt-search", order: 98, label: "PPT知识库" }, View))`。
（参考同级仓库 conv-4-inspector 的 `src/client.js` 的整体写法：CSS 注入、fetch 封装、深浅色 prefers-color-scheme。）

后端 API（前缀 `/ppt-search/api`，除注明外均 POST JSON，响应 `{ ok: true, value }` / `{ ok:false, error:{ code, message } }`）：

| 端点 | 请求 | 响应 value |
|---|---|---|
| `POST /list` | `{}` | `{ decks: Deck[] }`（按 addedAt 降序） |
| `POST /detail` | `{ deckId }` | `{ deck, slides: [{ n, title, body, notes }] }` |
| `POST /search` | `{ query, deckId?, limit? }` | `{ hits: Hit[] }` |
| `POST /delete` | `{ deckId }` | `{ deleted: true }` |
| `POST /compose` | `{ items: [{deckId,n,titleOverride?,replace?}], name? }` | `{ id, downloadUrl, slideCount }` |
| `POST /compositions` | `{}` | `{ items: [{id,name,size,mtime,downloadUrl}] }` |
| `POST /retry` | `{ deckId }` | `{ ok: true }`（重新入队处理） |
| `POST /upload` | body = 文件二进制，`?name=<encodeURIComponent 文件名>`，content-type: application/octet-stream | `{ deck }`（status=processing） |
| `GET /render/<deckId>/<n>.json` | — | application/json（渲染模型；no-cache；旧 deck 惰性生成） |
| `GET /media/<deckId>/<file>` | — | image/*（immutable 缓存） |
| `GET /download/<compositionId>` | — | application/vnd.openxmlformats-officedocument.presentationml.presentation，Content-Disposition attachment |

UI 结构（单页三栏/两态）：
1. 顶栏：上传按钮（`<input type=file accept=".pptx">`，读 ArrayBuffer POST）、搜索框（回车检索）、deck 过滤下拉、结果计数。
2. 主体左：结果网格 —— 每卡片 = 预览（`SlideCanvas` 组件拉取 `/render/...` JSON 用 canvas 绘制，失败显示占位）+ deck 名 + 页码 + 标题 + snippet（命中词 `<mark>` 高亮，snippet 文本自行按 query 分词高亮）。
3. 卡片点击 → 右侧抽屉/弹层：大图预览 + title/body/notes 全文 + 「加入提纲」按钮。
4. 「提纲」面板（可折叠侧栏）：已选 slide 列表（deck#页码+标题），支持上移/下移/移除、每页可编辑 titleOverride、全局替换输入框（find/replace 两行）、deck 名输入、「合成 PPT」按钮 → 成功后显示下载链接（`<a href=downloadUrl>`）。
5. deck 管理：下拉旁「管理」弹层列 deck（状态徽标 processing/ready/error + error 详情 + 重试/删除按钮）；processing 的 deck 每 2s 轮询 list 直到 ready。
6. 空态提示：无 deck 时引导上传。
7. 所有文案中文；CSS 前缀 `ps-`，注入 `<style>`，深浅色两套变量（参考 conv-4-inspector 的 `--ci-*` 模式，用 `--ps-*`）。

### F. 宿主集成（负责人：lead，勿动）

`lib/index.js`（inject `["webServer","tools"]`）、`lib/api.js`、`lib/tools.js`、`lib/mcp.js`、`lib/kb.js`、`lib/pptx/renderModel.js`。

处理管线：upload 落盘 → status=processing → parsePptx（文本索引）+ buildRenderModels（逐页 render JSON + 抽取 media）→ 写 meta + 索引 addSlide → status=ready。失败 status=error + error 信息。启动时从 library.json 重建索引（ready 的 deck）；旧 deck 缺 render JSON 时首次请求惰性重建。

Agent 工具（`ctx.tools.register(defineTool 形状的裸对象)`，`output.render(args,value)` 返回字符串）：
- `ppt_search`、`ppt_slide_get`、`ppt_deck_list`、`ppt_compose`（参数见 lib/tools.js 实现）。

MCP：`/ppt-search/mcp` 无状态 streamable-HTTP JSON-RPC（initialize / notifications/initialized(202) / tools/list / tools/call / ping），protocolVersion 回 `2025-03-26`。

## 测试基线

- `node test/test-parse.mjs`、`node test/test-compose.mjs`（pptx-engine）
- `node test/test-search.mjs`、`node test/test-store.mjs`（search-store）
- 全部离线可跑、不依赖网络；fixture 由 `test/make-fixture.py` 生成到 `test/fixtures/`（脚本需幂等）。
