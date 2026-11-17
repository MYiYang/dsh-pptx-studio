# dsh-pptx-studio 需求分析

> 目标一句话：在 DSH 右侧开一块「PPTX 实时画布」，用户给模板或长提示词，就能看着 deck 一页页画出来；发现哪页/哪个元素不对，点一下写句意见就发给 DSH 去改。
>
> 本文是**先分析需求**的产物（用户明确要求「先分析需求，再实现」）。实现状态与偏差记录见 `README.md`。

## 1. 用户与场景

目标用户：用 DSH 做汇报、方案、答辩材料的人。

现有链路（没有本插件时）：

1. 用户给模板或长提示词 → Agent 写 python-pptx 脚本 → 生成 `.pptx`；
2. 用户必须下载/在 PowerPoint 打开才知道长什么样；
3. 发现问题只能用文字描述（「第 3 页那个方块压住字了」），Agent 要靠猜；
4. 每次修改都要重开文件 → 一轮反馈 3~5 分钟，且改动过程不可见。

三个核心痛点：

| 痛点 | 本插件怎么解 |
|---|---|
| 生成过程是黑箱 | 右侧画布随 Agent 的写入增量绘制（页/元素级动画 + 状态条） |
| 问题指不准 | 画布上点元素或框选区域，自动带上 slide/元素/坐标/文本，一键发给 DSH |
| 模板复用难 | 解析 `.pptx` 模板的母版/版式/占位符/主题色，生成时继承，画布先显示模板骨架 |

## 2. 功能需求（分级）

### P0（本次必须可用）

1. **右侧实时画布**
   - 自动或一键在右侧栏打开；左侧缩略图带 + 右侧当前页大画布 + 底部状态条。
   - 缩略图与当前页同源：按 deck 的页尺寸（16:9 / 4:3 / 16:10 / A4）出图，用页面自己的底色，
     元素位置、字号、表格列宽与实际页成同一比例；先整页布局再整体缩放，不是把元素按 1.65px
     小字硬画进 88×60 的白底里。
   - 元素级渲染：文本、矩形、图片、线条、表格占位、图表占位；坐标系与 PowerPoint 一致（英寸/pt，16:9 13.333×7.5in）。
   - 支持缩放（适应宽度/100%）、翻页、全屏。
2. **绘制过程可见（不是「跑完突然出现」）**
   - **打开时机（所有步骤，不只是 open）**：12 个 `pptx_studio_*` 工具各注册一张 `tool.call.toolview` 卡片，
     在 `preparing`/`start` 阶段就 `openTab` 把右侧画布打开并切到前台，不等工具返回。带 `deck` 参数的步骤
     （模板/调研/预览/大纲/写页/改元素/校验/导出）直接按参数里的 deck 导航；`pptx_studio_open` 还不知道新 deck，
     就以 `pending` 方式打开，画布渲染该次调用的等待骨架（工具名 + 耗时），**不加载最近一次落盘的旧 deck**；
     结果带回 `deckId` 后再把已打开的 tab 导航过去。
   - **会话级进度流**：宿主把每次调用的 `{tool, deckId, state, startedAt, seq}` 发布到
     `GET /api/session?sessionId=`（快照）与 `GET /api/session/events?sessionId=`（SSE）。画布订阅它：
     deck 一变就切（工具卡片已经被卸载也一样）；`open` 在 running 且无 deck 时先把上一份 deck 移出屏幕；
     顶部显示「正在〈某步〉… deck: xxx Ns」，页签标题出现呼吸圆点。
   - **不会退回旧 deck**：`/state` 响应回来时比对当前 `deckId`，过期的响应直接丢弃；`pending` 参数带时间戳，
     页面刷新后的过期 pending 自动失效（否则会把画布卡死在空态）。
   - **没有 deck 时也有编辑区**：画布渲染完整的编辑区/功能区 shell（纸面骨架 + 3 个灰色缩略图占位 +
     分隔条），而不是一句空提示；Agent 一开画，内容就在已经存在的编辑区里生长，不挤压、不跳位。
   - **逐页**：`pptx_studio_plan` 先发 `planning` 状态（「正在摆放 N 页骨架」），再每 160ms 落一页骨架；
     写页前先发 `drawing(begin)`（「正在准备第 N 页」），220ms 后内容才到。
   - **逐元素**：整页内容到达时，客户端先整页画成铅笔幽灵（虚线 + 低透明度），再按数组顺序每 42–75ms
     落定一个元素；文本自左向右「写出来」，图形/图片/表格淡入 + 轻微缩放。
   - **进度**：编辑区标题出现 ✎ 绘制徽标，画布左下角有绘制 HUD（第 N 页 · 元素 x/y + 进度条），
     状态条同步；画完 0.7s 后 HUD 自动收起。
   - **稳定性**：缩略图栏常驻（无页时放骨架格），编辑区有 `min-height` 兜底；生成期间编辑区/功能区
     位置和高度不跳。
   - 「跟随最新」开关：生成中自动跳到新页；用户手动回看某一页时暂停跟随（不会抢镜头）。
3. **两种输入模式**
   - **模板模式**：用户提供 `.pptx` 模板（路径或上传）→ Host 解析母版、版式、占位符、主题色与字体 → 画布先画模板骨架（灰）→ Agent 按版式填内容。
   - **提示词模式**：长提示词 → Agent 先产大纲（画布上逐张出现页卡骨架），再逐页填充。
4. **问题指认 → 修正闭环**
   - 点元素选中（高亮框 + 元素信息：类型/文本/坐标/字号），或拖拽框选区域（bbox）。
   - 意见浮层：写一句自然语言意见 + 快捷标签（「太挤」「字号太小」「标题改短」「换配色」「图错了」）。
   - 「发给 DSH」：把结构化消息送进当前会话（slide 序号、元素 id、bbox、当前文本、用户意见、期望动作）。
   - Agent 修改后画布自动刷新，被改动的元素高亮/脉冲；支持「改动前/后」临时对比。
5. **落盘与导出**
   - 画布模型与真实 `.pptx` 同源；随时导出 `deck.pptx`（原生文本框/形状/图片，PowerPoint 打开可编辑）。
   - 工作区落 `.pptx-studio/<deckId>/deck.json` + `deck.pptx` + `assets/`，会话重开可恢复画布。

### P1（本轮之后）

- 占位符映射（模板里的 `{{title}}` / 空占位框自动填）。
- 版本历史（每次 patch 一个快照，可回退/对比）。
- 缩略图拖拽排序、增删页。
- 导出 PNG / PDF 预览。
- 一键「按模板重排全部页」。

### P2（不做，只记录）

- 多人协同标注、评论线程、审阅模式。
- 组合图形 / SmartArt / 动画时间轴的 1:1 渲染（只做近似与占位）。
- 直接编辑 OOXML 级特效（阴影、渐变、3D）。

## 3. 非功能需求

| 维度 | 要求 |
|---|---|
| 实时性 | Agent 写入到画布可见 ≤ 800ms（Host 500ms 防抖批量 + SSE 推送） |
| 安全 | 只读写当前工作区内路径；路径规范化后必须落在 session cwd 之下；不执行上传文件里的宏；不把 deck 内容发到任何外部服务 |
| 兼容 | Windows 优先；导出的 pptx 在 PowerPoint 打开无「修复」提示；中文用微软雅黑（latin+ea+cs 都设） |
| 健壮 | 复杂元素（组合/Freeform/图表）降级为占位并标注「近似渲染」；单个坏元素不能拖垮整页 |
| 性能 | 50 页 / 2000 元素以内滚动与翻页不卡；增量 diff 只重绘变化元素 |
| 可恢复 | 插件重启 / 页面刷新后从 deck.json 恢复；SSE 断线自动重连并拉全量 |

## 4. 技术方案与取舍

### 4.1 三种候选

| 方案 | 做法 | 优点 | 缺点 |
|---|---|---|---|
| A. 文档预览器扩展 | `documentPreviews.register` 接管 `pptx`，替换 DSH 原生 Office→PDF 预览 | 最贴原生：`present` 打开文件即在右侧看 | 只能「看文件」，绘制过程要靠文件多次重写 + 轮询 diff；标注能力受预览器 props 限制 |
| B. 自定义右侧栏 Tab | `sidebarRightTabs.register` 新 kind + `sidebarRight.openTab` | 完全可控：画布、动画、标注、属性面板都自由 | 依赖未公开稳定的 client API；需要自己管理 tab 生命周期 |
| C. 会话内工具卡片 + 右侧栏 | 工具卡片给入口/状态，右侧栏承载画布（B 的补充） | 有兜底入口；卡片可显示进度、可直接开画布 | 卡片本身不承载画布 |

**选定：C + B（B 为主，C 兜底）**。若 B 的 `openTab` 在插件上下文不可用，则降级为「点工具卡片按钮 → 打开右侧栏 tab」，再不行就退到 A（document preview 挂 pptx）。

### 4.2 架构

```
┌─ Agent(模型) ──────────────┐        ┌─ Host (index.js) ─────────────────┐
│ pptx_studio_* 工具调用      │ ─────► │ 1. DeckStore: deck.json (内存+盘) │
│ 读回 inspect 结果自检        │ ◄───── │ 2. 防抖写盘 + python 导出 pptx    │
└────────────────────────────┘        │ 3. SSE /events 广播增量事件        │
                                       │ 4. 路由 /state /feedback /export  │
                                       │ 5. 模板解析(pptx → template.json) │
                                       └──────────────┬────────────────────┘
                                                      │ SSE + fetch
                                       ┌──────────────▼────────────────────┐
                                       │ Client (client.js)                │
                                       │ 右侧栏 tab: 缩略图/画布/属性/意见  │
                                       │ 工具卡片: 入口 + 进度              │
                                       │ 发送意见 → /feedback → 会话       │
                                       └───────────────────────────────────┘
```

### 4.3 数据契约（v1）

`deck.json`：

```json
{
  "version": 1,
  "id": "deck_20261113_ab12",
  "title": "2026 年度复盘",
  "size": { "w": 13.333, "h": 7.5 },
  "theme": { "colors": { "ink": "#0F172A", "accent": "#2563EB" }, "fonts": { "latin": "微软雅黑", "ea": "微软雅黑" } },
  "template": { "path": "templates/brand.pptx", "layouts": [{ "name": "Title and Content", "placeholders": [] }] },
  "slides": [
    {
      "id": "s1",
      "layout": "title-content",
      "elements": [
        {
          "id": "e1",
          "type": "text",
          "box": [0.8, 0.5, 11.7, 0.9],
          "text": "2026 年度复盘",
          "style": { "size": 40, "bold": true, "color": "#0F172A", "align": "left", "font": "微软雅黑" }
        }
      ]
    }
  ]
}
```

SSE 事件：

```json
{ "type": "deck",   "seq": 12, "deck": { "...": "全量（首次/重连）" } }
{ "type": "slide",  "seq": 13, "slide": { "...": "新增/替换一页" }, "index": 2, "total": 8 }
{ "type": "element","seq": 14, "slideId": "s3", "element": { "...": "新增/替换一个元素" } }
{ "type": "status", "seq": 15, "phase": "drawing", "slideIndex": 3, "total": 8, "message": "正在生成第 3 页" }
{ "type": "saved",  "seq": 16, "path": ".pptx-studio/<deckId>/deck.pptx", "bytes": 102141 }
```

工具面（模型可见）：

| 工具 | 作用 | 关键参数 |
|---|---|---|
| `pptx_studio_open` | 新建/打开画布；模板模式传 template | `title`, `path`(可选), `template`(可选), `size` |
| `pptx_studio_plan` | 提交/更新大纲（页数、每页主题） | `deck`, `slides: [{title, intent}]` |
| `pptx_studio_slide` | 写入/替换整页 | `deck`, `index`, `slide` |
| `pptx_studio_patch` | 局部改元素（模型最常用） | `deck`, `slideId`, `elementId`, `patch` |
| `pptx_studio_inspect` | 读回当前 deck（供模型自检/续写） | `deck`, `slideId?` |
| `pptx_studio_export` | 立即导出 pptx 并返回路径/字节数 | `deck`, `path?` |

反馈消息（发给会话的用户消息模板）：

```
[PPTX 画布反馈] deck=<deckId> slide=3/8 element=e12(type=text)
当前文本：「2026 年度复盘」
位置：x=0.80in y=0.50in w=11.70in h=0.90in
用户意见：标题太长，压到副标题了，缩短到 12 字以内
期望动作：只改这个元素，其它页不要动；改完重新导出。
```

## 5. 关键流程

**模板模式**：给 `.pptx` → `pptx_studio_open(template=...)` 解析模板 → 画布显示模板骨架 → 用户说需求 → Agent `pptx_studio_plan` → `pptx_studio_slide` 逐页 → 画布动画 → 点元素写意见 → 发送 → Agent `pptx_studio_patch` → 画布高亮 → `pptx_studio_export`。

**提示词模式**：长提示词 → `pptx_studio_open` → `pptx_studio_plan` 大纲 → 画布页卡逐个出现 → 逐页 `pptx_studio_slide` 填充 → 同上闭环。

**异常路径**：SSE 断线 → 客户端指数退避重连并拉 `/state` 全量；Python 导出失败 → 工具返回错误码 + stderr 摘要，画布保留上次成功版本；模板解析失败 → 降级为空白 deck 并提示。

## 6. 验收标准

1. 右侧栏在 `pptx_studio_open` 一开始（`start`）就打开画布；`pptx_studio_slide` 连续写入 5 页时，
   每页先出铅笔幽灵、再逐个元素落定，HUD/状态条显示当前页与元素进度；
1.1 还没建 deck 时，画布也只有「编辑区 + 骨架纸面 + 缩略图占位 + 功能区」的完整 shell，不出现空白提示；
1.2 缩略图的底色、页尺寸比例、元素位置与当前页一致（深色稿的缩略图也是深色；4:3 稿不会被压成 16:9）；
2. 点击任意文本元素能选中并显示 slide/element/bbox 信息；
3. 点「发给 DSH」后，当前会话出现含 slide 序号、element id、bbox、用户意见的结构化消息；
4. Agent 执行 `pptx_studio_patch` 后画布 ≤1s 更新，被改元素有高亮；
5. `pptx_studio_export` 产出的 `.pptx` 能被 PowerPoint 打开（无修复提示），文本可编辑；
6. 模板模式下新页继承模板母版与主题色（导出后检查 `p:sldLayout` 引用与主题色）。

## 7. 风险与对策

| 风险 | 影响 | 对策 |
|---|---|---|
| 右侧栏/输入框 API 非公开稳定 | 打不开画布、发不出反馈 | 先做 API 侦察；准备 A/C 降级路径；全部调用做 `typeof === 'function'` 守卫 |
| 复杂元素无法 1:1 渲染 | 画布与 PowerPoint 不一致 | 组合/图表/SmartArt 用占位 + 「近似渲染」角标；导出以 PowerPoint 为准 |
| 字体度量差异 | 画布换行与导出不一致 | 画布用与导出相同的字体族与字号换算；导出后允许一次「渲染回看」 |
| python-pptx 能力有限 | 复杂版式做不出来 | 模板模式走占位符；自由模式只承诺基础几何 |
| 大量元素导致卡顿 | 体验差 | 增量 diff、虚拟化缩略图、单页渲染 |
| 反馈绕开权限体系 | 安全风险 | 只走会话正常用户消息通道，不直接调用工具 |

## 8. 里程碑

- M1（本轮）：完成 API 侦察 + 需求文档 + 插件骨架（package/cordis.patch/host 工具/skeleton client），装进 web profile 并能被加载。
- M2：Host DeckStore + SSE + python 导出跑通；`pptx_studio_*` 工具可用。
- M3：右侧画布渲染 + 增量动画 + 缩略图。
- M4：选中/框选 → 意见 → 发回会话闭环。
- M5：模板解析与版式继承；验收 1-6 全过。

## 9. v2 增量需求：调研 → HTML → PPTX，以及幻灯片特效

> v1 解决的是「画布实时可见 + 点元素反馈」；v2 要把生成过程变成有依据的设计流程，并让画布编辑覆盖 PPTX 的常用能力。

### 9.1 目标

1. **先生成依据，再生成文件**：模型不得直接吐 PPTX；先交需求分析/调研（受众、目标、大纲、事实与来源、视觉方向），
   用户能先审这份分析。
2. **先看 HTML 效果，再落 PPTX**：第二步交一份自包含 HTML 视觉稿，画布用沙箱 iframe 直接渲染；
   用户在生成 PPTX 前先确认版式、配色和字体。
3. **画布编辑具备 PPTX 常用能力**：至少覆盖幻灯片转场、元素入场动画、演讲者备注、页面背景、隐藏页、
   主题色/字体、画布尺寸、原生图表、表格、形状/阴影、富文本段和模板版式选择；这些能力必须真正写进 `.pptx`，
   不能只停留在画布。

### 9.2 功能需求（P0）

1. **三步流水线**
   - 新 deck 的 `pipeline.stage` 依次为 `research → preview → build`；阶段由「是否有 research / preview / builtAt」派生。
   - 新 deck 直接调 `plan/slide/add_element/export` 返回 `stage-research` 或 `stage-preview`，并给出下一步提示与 `force:true` 逃生口。
   - 已有页面的旧 deck 或用户手工加页的 deck 视为已构建，不被流程卡住（向后兼容）。
   - 画布顶栏显示三步条；已完成打勾，点击切换面板；默认跟随当前阶段。
2. **调研分析**
   - `pptx_studio_research` 接收 `brief`：摘要、受众、目标、语气、约束、逐页大纲、关键结论、来源、
     配色与字体方向；写入 `deck.pipeline.research` 并持久化。
   - 画布「① 调研」面板显示上述内容；来源 URL 只接受 http(s)，可点击；空 brief 被拒绝。
   - 调研面板的意见通过 `/feedback` 的 `stage=research` 送回会话，提示模型先更新分析。
3. **HTML 效果预览**
   - `pptx_studio_preview` 接收自包含 HTML（≤400KB）与可选的设计画布尺寸 `width/height`（px），写入
     `deck.pipeline.preview` 与 `preview.html`。画布按该尺寸等比缩放到侧栏宽度，不拉伸、不裁掉整页；
     未声明尺寸时从 HTML 的 `.slide` 宽度或 deck 尺寸推断，并预留常见的外边距。
   - 画布「② HTML」用 `<iframe sandbox="">` 指向 `GET /preview?deck=`；服务端同时下发 CSP `sandbox`，
     禁止脚本与外链样式，避免模型产出的 HTML 在 DSH 源里执行。
   - 支持刷新、新窗口打开、下载 HTML；预览面板的意见以 `stage=preview` 送回会话，要求先改预览再同步幻灯片。
4. **幻灯片特效与 PPTX 能力**
   - **转场**：`slide.transition` 支持 none/fade/push/wipe/split/cover/pull/cut/zoom/dissolve/circle/random，
     speed slow/medium/fast，direction left/right/up/down，advanceOnClick/advanceAfterMs；画布可预览；
     导出为 `<p:transition>`。
   - **入场动画**：`element.anim` 支持 appear/fade/wipe/zoom + direction/duration/delay/order；画布可单元素预览，
     转场预览时按 order/delay 依次播放；导出为 `<p:timing>` 的 clickEffect 序列。
   - **演讲者备注**：`slide.notes`，画布可编辑，导出到 notes slide。
   - **页面背景**：`slide.background.color`，画布与导出同步。
   - **隐藏页**：`slide.hidden`，缩略图标记，导出 `p:sld show="0"`。
   - **主题与尺寸**：`deck.theme.colors/fonts`、`deck.size`；画布有主题面板和 16:9/4:3/16:10/A4 预设；
     导出应用默认字体/文字色与画布尺寸。
   - **原生图表**：`element.chart` 支持 column/bar/line/pie/doughnut/area（categories + series），
     画布用 SVG/HTML 近似，导出为真实 chart part。
   - **表格 / 形状 / 阴影 / 富文本**：表格原生导出；rect 支持圆角/填充/描边/阴影；文本段落支持 runs。
   - **模板版式**：模板模式下页面设置可选 `layoutIndex`，导出引用对应 `p:sldLayout`。
5. **画布操作补充**
   - 页面设置面板：转场、速度、方向、换页方式、背景色、隐藏页、备注、版式、预览切换。
   - 主题面板：5 个主题色、字体、画布尺寸。
   - 插入按钮：文本、图片、矩形、表格、图表。
   - 属性面板：下划线、阴影开关、入场动画效果/方向/时长/预览。
   - 撤销/重做 diff 覆盖 deck 级（title/theme/size）与 slide 级（transition/notes/background/hidden/layout）字段。

### 9.3 数据契约增量（deck.json v2）

```jsonc
{
  "version": 2,
  "theme": { "colors": { "ink": "#0F172A", "muted": "#64748B", "accent": "#2563EB", "surface": "#F1F5F9", "background": "#FFFFFF" },
             "fonts": { "latin": "微软雅黑", "ea": "微软雅黑", "heading": "微软雅黑" } },
  "pipeline": {
    "research": { "summary": "…", "audience": "…", "goal": "…", "tone": "…", "constraints": "…",
                  "outline": [{ "title": "背景", "intent": "…", "source": "…" }],
                  "facts": [{ "claim": "…", "source": "…" }],
                  "sources": [{ "title": "…", "url": "https://…", "note": "…" }],
                  "palette": ["#22D3EE"], "fonts": { "latin": "微软雅黑" }, "updatedAt": "…" },
    "preview": { "html": "<!doctype html>…", "title": "视觉稿", "notes": "…", "width": 1008, "height": 588, "updatedAt": "…" },
    "builtAt": "2026-11-13T09:00:00.000Z",
    "stage": "build"
  },
  "slides": [{
    "id": "s1", "notes": "演讲者备注", "hidden": true,
    "background": { "color": "#060B18" },
    "transition": { "type": "push", "speed": "slow", "direction": "up", "advanceOnClick": false, "advanceAfterMs": 2000 },
    "elements": [{
      "id": "e1", "type": "text", "box": [1, 1, 4, 1], "text": "标题",
      "anim": { "effect": "fade", "direction": "left", "duration": 600, "delay": 0, "order": 0 },
      "shadow": { "blur": 4, "distance": 3, "direction": 90, "color": "#000000", "alpha": 0.35 },
      "paragraphs": [{ "text": "段", "runs": [{ "text": "run", "bold": true, "color": "#FFFFFF" }] }]
    }]
  }]
}
```

### 9.4 验收标准

1. 新 deck 未调 research 时 `plan/slide/export` 返回 `stage-research`；有 research 无 preview 时返回 `stage-preview`；
   完成 preview 后放行；`force:true` 可显式跳过。
2. 调研/HTML 两个面板能在画布上打开并显示宿主数据；HTML iframe 的 `src` 打到 `/preview`，带 sandbox。
3. 画布改转场/备注/背景/隐藏/主题/尺寸后，`deck.json` 与 `.pptx` 同步；切换预览能在画布上看到 `data-fx` 动画。
4. 导出的 `.pptx` 回读后：转场页数、timing 页数、动画数、备注页数、隐藏页数、图表数与模型一致；
   OOXML 子元素顺序为 `cSld → clrMapOvr → transition → timing`。
5. 图表为原生 chart part（`ppt/charts/chartN.xml`），可在 PowerPoint 中编辑数据；文本/形状/表格保持可编辑。
6. `npm test` 全绿（model/store/python/e2e/client）。

### 9.5 非目标（v2 仍不做）

- 组合形状、SmartArt、group 的 1:1 渲染；仍以占位/近似处理。
- PowerPoint 的复杂动画时间轴、motion path、触发器、退出动画、转场 morph。
- 多人协同评论、审阅模式、版本历史（v1 P1 仍未做）。
- HTML 预览执行脚本或访问外网；它只用于静态视觉确认。


## 10. v3 增量需求：编辑区/功能区视觉分界与 PPTX 快捷键

### 10.1 目标

1. **一眼能分清在哪儿编辑、在哪儿操作**：编辑画布与控件面板不能混成一条长列表。
2. **常用 PPTX 快捷键可用**：保存、剪切/复制/粘贴、撤销/重做、复制、新建页、删除、微调、缩放。

### 10.2 功能需求（P0）

1. **视觉分界**
   - 编辑区为独立卡片：`--ps-border-strong` 边框、标题行「编辑区 · 幻灯片编辑/调研/HTML 预览」、浅灰底；
     幻灯片纸面保留自有背景。
   - 功能区为独立卡片：标题行「功能区」，包含状态、缩放、页操作、插入、属性、动效、主题和意见；
     `max-height` + 内部滚动，内容多时只滚动功能区，不挤压编辑区。
   - 中间有可见分隔条，带「编辑区 / 功能区」说明；窄侧栏也保持上下分界。
2. **快捷键**
   - `Ctrl+S` 保存/导出 PPTX（输入框里也生效）；`Ctrl+Shift+S` 导出并下载。
   - `Ctrl+Z` / `Ctrl+Shift+Z` / `Ctrl+Y` 撤销/重做；输入框内保留浏览器原生文字撤销。
   - `Ctrl+X/C/V` 剪切/复制/粘贴元素；内部剪贴板 + 系统剪贴板 JSON；输入框内保留原生复制粘贴。
   - `Ctrl+A` 选中本页全部元素（多选显示虚线框 + 批量面板）；`Ctrl+D` 复制一份；`Ctrl+M` 新增一页。
   - `Delete/Backspace` 删除选中元素；方向键微调；`Ctrl+0` / `Ctrl±` 缩放；`Esc` 取消；`?` 切换帮助。
   - 多选批量删除/粘贴应作为**一步**进入撤销栈。
3. **帮助与可发现性**
   - 功能区右上角有「快捷键」按钮，打开快捷键表；画布 `?` 也可切换。

### 10.3 验收标准

1. DOM 中存在 `data-ps-editor`、`data-ps-controls`、`data-ps-split`，且样式包含强边框/分隔条。
2. 画布根节点的按键处理覆盖上述快捷键；`Ctrl+S` 触发 `/export`，`Ctrl+M` 触发 `slide.add`，
   `Ctrl+C/V` 产生 `element.add`，`Ctrl+X` 产生 `element.patch {remove:true}`，`Ctrl+Z` 产生撤销 diff。
3. 在 `textarea` 内按 `Ctrl+Z` 不触发画布撤销；`Ctrl+S` 仍保存。
4. `Ctrl+A` 后 `Delete` 一次提交 N 条删除 op；撤销可一步恢复。
5. 快捷键帮助面板列出主要键位。
6. `npm test` 全绿（model 23 / store 7 / python 4 / e2e 19 / client 19）。

---

## 11. v3 增量：插件自带技能，约束「写 pptx 必须走插件」

### 11.1 背景

真机会话里官人说「生成一个充满科技感风格的 ppt」，模型没调任何 `pptx_studio_*`，而是照
`~/.dsh/skills/pptx-research-deck/SKILL.md` 写了 `build_deck.py` 走 python-pptx。
根因：工具表只在模型「想找工具」时说话；技能表里那份 user 技能 description 逐字就是「用 python-pptx 从零生成」，
先声夺人。插件当时没有任何约束通道（`index.js` 里只有 `tools.register` 与 `webServer.register`）。

### 11.2 需求

1. 插件的约束要随包走：装插件即生效，卸载即撤回；不往 `~/.dsh/skills` 抄副本，也不写进长期记忆。
2. 约束以「技能」形式出现在宿主技能表里，description 与正文都要写明：`.pptx` 目的产物只用 `pptx_studio_*`，
   禁止 python-pptx 从零拼；与 `pptx-research-deck` / `office-pptx` 的手搓路线冲突时以本技能为准；
   例外只有「用户明说不要用插件」与「工具表里没有」两条。
3. 技能里给全七步工具链与坑位（英寸/pt/px 三种单位、阶段闸门、越界、图表字段）。
4. 可单独关闭（`config.skill === false`），且宿主缺 `skills` 服务时降级为一条日志，不影响插件加载。

### 11.3 实现

- `skills/pptx-studio/SKILL.md`：规则四条 + When to use / Steps / Pitfalls / Verification。
- `lib/skills.js`：`parseFrontmatter` / `readSkill` / `loadBundledSkills` / `skillRegistration` / `registerBundledSkills`；
  注册走 `ctx.inject(['skills'])` → `skills.register()`，返回的 disposer 再包进 `scope.effect`，随 fiber 卸载撤回。
- `skillRegistration()` **必须带 `source`**：注册表按 `SkillDefinition` 校验，缺 `source` 时 `skill` 工具报
  `loaded skill "pptx-studio" source must be a string`（只有 `invocation` / `provider` 有默认值）。
- `index.js`：`apply()` 末尾接线；`GET /api/health` 回带 `skills: [...]`。

### 11.4 验收标准

1. `npm test` 全绿（model 26 / template 27 / store 7 / python 7 / skills 14 / e2e 29 / client 30 = 140）。
2. 禁用再启用 `include:dsh-pptx-studio` 后，本会话技能表里出现 `pptx-studio`；再禁用则消失（证明注册挂在插件 fiber 上）。
3. `skill` 工具能加载 `pptx-studio`，返回正文与 `resourceBase` 目录。
4. 新会话只给「生成一个充满科技感风格的 ppt」：首个制作动作是 `pptx_studio_open`，全程无 `.py` 落盘。
5. 工作区 `.tmp/skill-probe` 探针留下的只有 `.pptx-studio/deck_*/{deck.json,preview.html}`，均为插件自产。

### 11.5 已知缺口

- 本层是「可被发现、可被说服」，不是硬拦：模型若执意手写脚本，仍需 `tools.guard` 兜底（未做，待批）。
- `~/.dsh/skills/pptx-research-deck` 原样保留，只被本技能第 3 条压住；要收窄它得单独动家目录。


## 12. v4 增量：模板分解 → 需求分解 → HTML 模版 → 幻灯片（一致性）

> v1 的模板模式只解析母版/版式，内容仍由模型自由摆放；v2/v3 解决了流程与技能约束。v4 把「用户给的模板」变成可复用的版式 + 文字槽，需求先落到模板页上，HTML 模版与幻灯片从同一份映射生成，导出前校验一致性。

### 12.1 目标

1. **模板可复用，而不是只换母版**：把 `.pptx` 每页拆成「角色 + 干净底图 + 文字槽」，新内容填进槽位，版式/配色/字距沿用模板。
2. **需求先落到模板页上**：需求分解（outline）每项必须写用哪一页模板、每个字段写什么；放不下就拒收，而不是生成阶段才发现。
3. **HTML 与幻灯片同源**：同一份 `pages` 映射、同一套模板几何；改文案两边一起变，不存在「HTML 好看、PPTX 跑版」。
4. **导出前可判定**：模板模式逐页校验底图、字段位置、字号与文字，不一致拒绝导出并指出页号/字段。

### 12.2 功能需求（P0）

1. **⓪ 模板分解（`pptx_studio_template`）**
   - `pptx_studio_open(template='...')` 进模板模式；未分解时 `plan/slide/add_element/export` 返回 `stage-template`
     （`force:true` 只用于用户明确要求跳过流程）。
   - 画布阶段条在原来三步前加一格 `⓪ 模板`：模板未分解时默认停在模板面板（页数/角色/字段/配色/映射），
     分解完成后才允许进入 ① 调研。
   - PowerPoint COM 每页渲两遍：原样 `reference`（对照）与隐藏全部文字后的干净底图 `plate`；解析每页角色
     （`cover/toc/section/content/ending`）、文字字段（`id/sample/box/style`、`fit.maxCjk` 可放字数）、
     配色/字体、同款底版分组 `families`。
   - 逐字拆分的单字文本框合并成**一个**字段（`perChar:true`，取各框并集 box，重填按整行居中）；
     `footer/date/slideNumber` 不作为可填字段。
   - 产物落 `<deck>/template/{template.json,plates,thumbs,reference}`；`pipeline.template` 记
     `path/pages/engine/textHidden/design`；`deck.size` 与主题色/字体沿用模板。
   - `list:true` 只读已分解结果；`render:false` 只解析结构（无底图）；没装 PowerPoint 时先试 Pillow 近似。
2. **① 需求分解（`pptx_studio_research`）**
   - 模板模式下 `brief.outline` 每项必须 `{title, templateSlide, fields:{字段id: 文案}}`；无映射返回
     `template-mapping-required`。
   - `checkMapping` 校验并拒收：缺 `templateSlide`、页号越界、字段不存在（错误里给出该页可用字段）、
     字数 > `fit.maxCjk × 1.35`；略超字数、章节页塞超过 3 个字段、连续两页用同一模板页只记 warning。
   - 拒收码 `template-mapping-invalid`，一次列出全部问题；通过后 `research.outline` 即需求分解结果。
3. **② HTML 模版（`pptx_studio_preview`）**
   - 模板模式**不接收手写 HTML**（传了也被忽略）：`pages` 省略时直接用 `research.outline` 的
     `templateSlide/fields`；宿主用 `renderTemplateHtml` 生成每页 `.slide`（底图 data URL +
     按模板 box/style 绝对定位的文字）。
   - `preview.mode='template'` + `preview.pages` 存映射；`preview.html` 落盘，但 `GET /preview`
     按需重算（映射改了不会看到旧稿）；缺底图返回 `preview-missing-plate`。
4. **③ 幻灯片（`pptx_studio_plan` / `pptx_studio_slide`）**
   - `plan` 每页可带 `templateSlide`，模板模式骨架页不发明标题框；`slide` 传 `templateSlide` + `fields`，
     可加 `extra` 图元。
   - 位置/字号/颜色来自模板字段，调用方不能自带 `box`；传了 `templateSlide`/`fields` 时手写的
     `slide.elements` 被忽略（加图元用 `extra`）；不走模板字段、自己写元素的页会被 verify 判为不一致。
   - `slide.templateSlide` 与 `element.field` 记录来源；底图元素固定 `field='__plate'`。
5. **④ 校验与导出（`pptx_studio_verify` / `pptx_studio_export`）**
   - `verify` 逐页检查：是否从模板页建、底图是否同页、字段是否模板所有、字段 box 漂移 > 0.06in、
     字号差 > 1pt、模板里有内容的字段是否漏填；并逐字比对幻灯片文字与 HTML 映射
     （`previewConsistency`）。返回 `ok/consistent/issues/warnings`。
   - `export` 在模板模式先跑 `previewConsistency`：页数、`templateSlide`、字段文字任何一处不一致返回
     `template-inconsistent` 且不写出文件。
   - 没装 PowerPoint 且渲染不可用（`render.ok=false`，如 `render:false`）时缺底图是降级而非错误：
     `verify` 记 warning，导出照常完成；Pillow 近似底图存在但不是原版，属另一种降级（收据提示）。

### 12.3 数据契约增量

`deck.json` 在 v2 模型上增量（`DECK_VERSION` 仍为 2）：

```jsonc
{
  "pipeline": {
    "template": {
      "path": "templates/brand.pptx", "required": true,
      "decomposedAt": "2026-11-16T09:00:00.000Z",
      "pages": 6, "engine": "powerpoint", "textHidden": true,
      "design": { "palette": ["#0F172A"], "background": "#0F172A", "fonts": { "title": "…", "body": "…" } }
    },
    "preview": {
      "mode": "template",
      "pages": [{ "templateSlide": 0, "title": "封面", "fields": { "title": "…", "subtitle": "…" } }],
      "title": "HTML 模版", "width": 1280, "height": 720, "updatedAt": "…"
    },
    "stage": "preview"
  },
  "slides": [{
    "id": "s1", "templateSlide": 0,
    "elements": [
      { "id": "s1-e1", "type": "image", "box": [0, 0, 13.333, 7.5], "src": "template/plates/s0.jpg", "field": "__plate" },
      { "id": "s1-e2", "type": "text", "box": [0.9, 4.6, 6.2, 1.1], "text": "2026 年度复盘",
        "style": { "size": 54, "bold": true, "align": "left" }, "field": "title" }
    ]
  }]
}
```

`<deck>/template/template.json`（`assets/pptx_template.py` 产出、`lib/template.js` 归一化，`version: 2`）：

```jsonc
{
  "version": 2, "path": "templates/brand.pptx",
  "size": { "w": 13.333, "h": 7.5 }, "engine": "powerpoint",
  "render": { "ok": true, "engine": "powerpoint", "textHidden": true, "width": 1600, "height": 900 },
  "design": { "palette": [{ "hex": "#0F172A", "share": 0.42 }], "background": "#0F172A",
              "fonts": { "title": "…", "body": "…" } },
  "families": [{ "id": "f1", "role": "content", "slides": [1, 2, 4] }],
  "slides": [{
    "index": 0, "role": "cover",
    "plate": "plates/s0.jpg", "reference": "reference/s0.jpg",
    "fields": [{
      "id": "title", "sample": "敦煌夜游", "perChar": true,
      "box": [0.9, 4.6, 6.2, 1.1],
      "style": { "size": 54, "bold": true, "align": "ctr", "anchor": "middle", "color": "#F8FAFC", "font": "…" },
      "fit": { "lines": 1, "cjkPerLine": 6, "maxCjk": 6, "maxLatin": 11 }
    }],
    "capacity": { "fields": 2, "maxCjk": 24, "bodyMaxCjk": 18 }
  }]
}
```

字段口径：`slides[].fields[].id` 是**唯一可填的字面**（`title/subtitle/body1…`，按字号从大到小命名）；
`box` 单位英寸、来自模板；`perChar` 表示由多个单字框合并；`fit.maxCjk` 是该格字数预算
（`checkMapping` 与 `buildSlideFromTemplate` 都据此判断溢出）。

### 12.4 验收标准

1. 给模板后未调 `pptx_studio_template`：`plan/slide/export` 返回 `stage-template`；分解后 `pipeline.stage`
   依次为 `research → preview → build`。
2. `outline` 写漏 `templateSlide`、页号越界、字段名写错或超字数时，`pptx_studio_research` 返回
   `template-mapping-required` / `template-mapping-invalid`，错误里列出该页可用字段；修正后通过。
3. 模板模式 `pptx_studio_preview` 不传 HTML 也能生成 HTML 模版：`/preview` 每页 `.slide` 的底图来自对应
   模板页 `plate`，文字位置与模板 `fields[].box` 同比例（缩放后），`preview.mode='template'` 且
   `pages` 与 `research.outline` 一致。
4. **导出后的 `.pptx` 与模板同款底图与字段位置**：回读每页第一层是与 `template/plates/` 同源的整页图片；
   文字框位置偏差 ≤ 0.06in、字号差 ≤ 1pt、文字内容与 `fields` 一致。
5. **HTML 与 PPTX 文字必须一致**：只改映射（`research`/`preview`）再重写该页时两边同步变化；若只改单边
   （例如绕过映射 patch 幻灯片），`pptx_studio_verify` 报不一致，`pptx_studio_export` 返回
   `template-inconsistent` 且不写出文件。
6. 没装 PowerPoint 的降级：Pillow 近似（底图存在但非原版，收据提示）或 `render:false` 纯色底
   （无底图，`verify` 记 warning）时导出仍可完成。
7. `npm test` 全绿（model / store / python / skills / template / e2e / client）。

### 12.5 非目标与已知边界

- 不重排模板版式：只填文字槽 + `extra` 图元；页内装饰不提供单独编辑，要改装饰回 PowerPoint 改模板再分解。
- 不做 OOXML 占位符自动映射（`{{title}}` 自动填）：仍由模型显式给 `fields`。
- 不继承模板的转场/动画：导出特效按 deck 模型字段，模板只提供版式、底图、配色与字体。
- 底图是整页图片：动态图表/表格等要原生可编辑的内容放 `extra`，不要指望从底图里抠出来。


- 继承色文字按底图取色：无显式颜色的字段按字段框内的底图亮度在深/浅墨色间选择（`apply_contrast_colors`），
  显式 `style.color` 永远优先。
- 模板可在工作区外：`pptx_studio_open(template=绝对路径)` 会按内容哈希复制到 `.pptx-studio/_templates/`，
  后续步骤只读工作区内副本；`ensureInside` 同时修正了「跨盘符路径被当成工作区内」的旧漏洞。