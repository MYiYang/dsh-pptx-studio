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
   - 元素级渲染：文本、矩形、图片、线条、表格占位、图表占位；坐标系与 PowerPoint 一致（英寸/pt，16:9 13.333×7.5in）。
   - 支持缩放（适应宽度/100%）、翻页、全屏。
2. **绘制过程可见**
   - Agent 每新增/修改一页或一个元素，画布增量出现（淡入 + 描边动画），状态条显示「正在生成第 N 页 / 共 M 页」。
   - 「跟随最新」开关：生成中自动跳到新页；用户手动回看某一页时暂停跟随。
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
{ "type": "saved",  "seq": 16, "path": "F:/.../deck.pptx", "bytes": 102141 }
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
| `pptx_studio_finish` | 收尾：导出 + 状态置完成 | `deck`, `path?` |

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

1. 右侧栏能打开画布，`pptx_studio_slide` 连续写入 5 页时能看到逐页绘制动画；
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
