# dsh-pptx-studio

在 DSH 右侧开一块 **PPTX 实时画布**。它不是「一句话直接吐一个 .pptx」：默认走
**调研分析 → HTML 效果预览 → 可编辑 PPTX** 三步，让用户在花时间生成幻灯片之前，先看见
分析结论和视觉方向；确认后再逐页画出带转场、入场动画、备注、主题和图表原生可编辑的 `.pptx`。
用户给一份 `.pptx` 模板时改走**模板优先**流程：先分解模板，再把需求逐页映射到模板字段，
HTML 与幻灯片从同一份映射生成（见下文「模板优先」一节）。

```
Agent：pptx_studio_template ─► 画布「⓪ 模板」面板（每页角色/可填字段/配色；仅模板模式）
       pptx_studio_research ──► 画布「① 调研」面板（受众/目标/大纲/来源/配色）
       pptx_studio_preview  ──► 画布「② HTML 效果」沙箱 iframe（先确认视觉方向）
       pptx_studio_plan/slide ─► 画布「③ 幻灯片」逐页/逐元素绘制
       pptx_studio_export   ──► 带转场/动画/备注/图表的真实 .pptx
        ▲                                          │
        └── 用户点元素或阶段面板写意见 ◄────────────┘
```

## 装了什么

| 文件 | 作用 |
|---|---|
| `index.js` | Host 半侧：12 个模型工具 + SSE + 调研/HTML 预览/模板分解/反馈/导出/资源路由 + 消息注入 |
| `client.js` | 客户端半侧：右侧栏画布（阶段条 / 模板面板 / 调研面板 / HTML 预览 / 缩略图 / 画布 / 属性 / 动效 / 主题 / 意见）+ 工具卡片 |
| `lib/model.js` | deck 模型 v2：调研/预览流水线、转场、入场动画、备注、背景、阴影、主题、原生图表 |
| `lib/store.js` | deck 存储：内存 + 400ms 防抖落盘 + 事件广播 |
| `lib/python.js` | python-pptx 桥：解释器探测 + 子进程 + JSON 解析 |
| `lib/template.js` | 模板优先核心：spec 归一化、需求映射校验（`checkMapping`）、HTML 模版渲染、按模板字段建页、一致性比对 |
| `lib/skills.js` | 把包内 `skills/` 注册进宿主技能表（`ctx.inject(['skills'])`，随插件卸载自动撤回） |
| `skills/pptx-studio/SKILL.md` | **插件自带的技能**：写 pptx 走插件的规则 + 七步工具链（含模板模式分支）+ 坑位清单 |
| `assets/pptx_build.py` | `deck.json → 原生可编辑 .pptx`（含 OOXML 转场/动画/备注/背景/隐藏页/阴影/图表） |
| `assets/pptx_inspect.py` | 解析模板：尺寸、版式、占位符、主题色与字体 |
| `assets/pptx_template.py` | ⓪ 拆模板：逐页解析角色/文字槽/配色，渲「原样 + 去字」两套底图（无 PowerPoint 时 Pillow 近似） |
| `assets/pptx_render.ps1` | PowerPoint COM 渲图脚本：每页渲两遍（原样 `reference` + 隐藏全部文字后的 `plate`） |
| `REQUIREMENTS.md` | 需求分析（场景、分级、技术取舍、验收标准、风险 + v2 增量） |

## 为什么写 pptx 必须走插件（自带的技能就是约束）

工具只在你「想找工具」时说话；模型想「自己写脚本」时，工具表不在它眼里。所以这个包还自带一个技能
`skills/pptx-studio/SKILL.md`，`apply()` 里通过 `ctx.inject(['skills'])` 把它注册进宿主技能表
（`runtime` 层，随插件卸载自动撤回，不往 `~/.dsh/skills` 抄任何副本）。技能开头的「规则」四条是这么写的：

1. 任何 `.pptx` 目的产物只用 `pptx_studio_*`，第一条命令是 `pptx_studio_open`；
2. 禁止用 python-pptx（`Presentation()`、`slide_layouts[...]`、`add_slide(...)`）从零拼幻灯片；
3. `pptx-research-deck`、`office-pptx` 等技能讲 python-pptx 手搓的地方与本技能冲突时，**以本技能为准**，
   它们只借「核查已有 pptx」的部分（溢出检测、COM 导图）；
4. 例外只有两条：用户明说「不要用插件，手写脚本」，或 `pptx_studio_*` 不在当前工具表里。

| 关掉它 | 怎么关 |
|---|---|
| 只想关技能、工具照用 | profile 里给这个 bundle 传 `config: { skill: false }` |
| 整插件关 | `plugin_manager` 禁用 `include:dsh-pptx-studio` —— 技能随之消失（注册挂在插件 fiber 上） |
| 想用工作区自己的规矩压过它 | 在 `.dsh/skills/` 放同名 `pptx-studio` 技能：同一层里 project 条目排名高于 runtime |

验证：改完 host 代码后用 `plugin_manager` 先禁用再启用 `include:dsh-pptx-studio`（模块会从磁盘重新导入），
技能表里出现 `pptx-studio` 就说明注册成功；`GET /dsh-pptx-studio/api/health` 也会返回 `skills: ["pptx-studio"]`。

## 三步工作流（默认强制执行）

| 阶段 | 工具 | 画布 | 解锁条件 |
|---|---|---|---|
| ① 调研分析 | `pptx_studio_research` | 受众、目标、摘要、叙事大纲、关键结论、来源链接、配色/字体方向 | 新 deck 的起点 |
| ② HTML 效果 | `pptx_studio_preview` | 沙箱 iframe 直接渲染自包含 HTML 视觉稿；可刷新/新窗口/下载 HTML | 已有调研 |
| ③ 幻灯片/导出 | `pptx_studio_plan` → `pptx_studio_slide` → `pptx_studio_export` | 逐页骨架 → 逐元素绘制 → 动效/备注/主题编辑 | 已有调研 + HTML 预览 |

- 新 deck 直接调 `plan` / `slide` / `export` 会被拒绝，返回 `stage-research` 或 `stage-preview`
  与下一步提示；确实要跳过流程时传 `force:true`（用户明确要求快速草稿才用）。
- 已经有页面的旧 deck（或用户在画布上手工加过页）视为已完成流程，不会被卡住。
- 每个阶段都有意见输入框：调研阶段发给 `pptx_studio_research`，HTML 阶段先改
  `pptx_studio_preview` 再同步到幻灯片，幻灯片阶段仍走元素/区域定位反馈。
- 带 `.pptx` 模板时，这三步之前还有「⓪ 模板分解」（阶段闸门返回 `stage-template`），见下一节。

## 模板优先：PPTX 模版 → 需求分解 → HTML 模版 → 幻灯片

用户给一份 `.pptx` 模板时，流程多出前置的「⓪ 模板分解」，之后每一步都从模板派生：

```
       给了 .pptx 模板（pptx_studio_open 传 template）
                    │
                    ▼
⓪ pptx_studio_template ──► <deck>/template/{template.json,plates,thumbs,reference}
   每页渲两遍：原样 reference + 隐藏全部文字后的干净底图 plate
   解析每页角色（封面/目录/章节/内容/结尾）与文字字段（id/box/style/可放字数）
                    │
                    ▼
① pptx_studio_research ──► brief.outline 每项 {title, templateSlide, fields}
   需求直接落到模板页上；checkMapping 校验页号/字段名/字数，不合格拒收并列出可用字段
                    │
                    ▼
② pptx_studio_preview ──► preview.mode='template' + pages（同一份映射）
   不手写 HTML：宿主用 renderTemplateHtml 把映射 + 模板底图/字段位置渲成 HTML 模版
                    │
                    ▼
③ pptx_studio_slide ───► 每页 templateSlide + fields → 底图 + 模板 box/style 的文字
   pptx_studio_plan 用同一顺序列页（每页可带 templateSlide）；位置不由调用方给，不会跑版
                    │
                    ▼
④ pptx_studio_verify ──► pptx_studio_export
   逐页比对底图/字段位置/字号/文字；不一致时 export 返回 template-inconsistent，拒绝导出
```

| 阶段 | 工具 | 输入 | 输出 |
|---|---|---|---|
| ⓪ 模板分解 | `pptx_studio_template` | deck（`open` 时传了 `template`）；可选 `pages`（只分解前 N 页）、`render:false`（只解析结构）、`list:true`（只读上次结果） | 每页角色、干净底图、文字字段（收据里是 `id=示例（≤N字）`）；`pipeline.template` 记 path/pages/engine/textHidden；spec 落 `template/template.json` |
| ① 需求分解 | `pptx_studio_research` | `brief`：`summary/audience/goal/tone/constraints` + `outline:[{title,intent,source,templateSlide,fields}]` | 写 `pipeline.research.outline`；缺映射返回 `template-mapping-required`，页号越界/字段不存在/超字数返回 `template-mapping-invalid` 并列出该页可用字段 |
| ② HTML 模版 | `pptx_studio_preview` | `pages`（可省略，默认取 `research.outline` 的 `templateSlide/fields`） | `preview.mode='template'` + `preview.pages`；HTML 由 `/preview` 路由按需生成 |
| ③ 幻灯片 | `pptx_studio_plan` → `pptx_studio_slide` | `plan`：按同一顺序列页（每页可带 `templateSlide`）；`slide`：`templateSlide` + `fields`（可加 `extra` 图元） | 每页 = 模板底图 + 按模板 `box/style` 生成的文字；`slide.templateSlide` / `element.field` 记录来源 |
| ④ 校验/导出 | `pptx_studio_verify` → `pptx_studio_export` | deck | `verify` 返回 `consistent/issues/warnings`；模板模式不一致时 `export` 返回 `template-inconsistent`，文件不写出 |

**为什么不会跑版、不会漂移**：模板模式只维护一份数据——`preview.pages` 那一份逐页映射
（第 N 页用模板第几页、每个字段写什么）。HTML 模版由 `renderTemplateHtml(spec, pages)` 生成，
幻灯片由 `buildSlideFromTemplate(spec, page, fields)` 生成：几何都来自模板字段的 `box/style`
（`template.json` 的 `slides[].fields[]`），调用方没有机会自带坐标，所以「手滑写错位置」跑不了版；
文字都来自同一份 `fields`，所以改一句文案不会出现「HTML 改了、PPTX 没改」的漂移。
导出前再对一遍：页数、`templateSlide`、每个字段的文字必须逐字一致，否则 `pptx_studio_export`
直接返回 `template-inconsistent`，不写出文件。

模板模式的最小调用序列：

```
pptx_studio_open(template: 'templates/brand.pptx')
pptx_studio_template
pptx_studio_research(brief: { summary: '…', outline: [
  { title: '封面', templateSlide: 0, fields: { title: '2026 年度复盘', subtitle: '管理层评审' } },
  { title: '结论', templateSlide: 3, fields: { title: '三件事', body1: '……' } },
] })
pptx_studio_preview            # 不传 html/pages，宿主按 outline 生成 HTML 模版
pptx_studio_plan(slides: [{ title: '封面', templateSlide: 0 }, …])
pptx_studio_slide(index: 1, templateSlide: 0, fields: { title: '2026 年度复盘' })
pptx_studio_verify → pptx_studio_export
```

产物落在 `<deck>/template/`：`template.json`（页/角色/字段/配色）、`plates/`（去字干净底图，
导出时铺满整页）、`thumbs/`（HTML 模版内联用的小图，避免预览文件几十 MB）、
`reference/`（原样整页渲染，对照用）。

## 怎么用（对话里）

```
用 pptx_studio_open 打开画布，标题「2026 年度复盘」，模板用 templates/brand.pptx
pptx_studio_template：先把模板拆成每页版式与可填字段（模板模式下这一步必须先做）
先用 pptx_studio_research：受众是管理层，目标是批准预算，先调研行业数据并给 8 页大纲
  —— 模板模式：outline 每项写 templateSlide 与 fields，把每页落到模板页上
pptx_studio_preview：模板模式不用给 HTML，宿主按同一份映射生成 HTML 模版
pptx_studio_plan：8 页，封面/背景/数据/结论/建议/…
pptx_studio_slide 第 1 页：……（带 transition、notes、元素的 anim）
pptx_studio_deck：主题 accent 改成 #22D3EE，画布改成 16:9
pptx_studio_patch：第 3 页 s3-e2 的字号改成 24、文字改成「……」
pptx_studio_export：导出到 reports/2026-review.pptx
```

## 画布上能做什么

### 生成过程是「看得见的」

不是「跑完突然出现一份 deck」，而是从调用开始就在你眼前长出来：

| 阶段 | 画布上发生什么 | 依据 |
|---|---|---|
| 工具调用开始 | **任何** `pptx_studio_*` 调用一开始就把右侧画布打开并切到前台（不等工具返回）；新画布还没有 deck 时显示「正在准备新的画布…」的进度骨架，**不会**退回上一份 deck | 12 个工具各有一张 `tool.call.toolview` 卡片，在 `preparing`/`start` 阶段就 `openTab`；`pptx_studio_open` 带 `pending` 参数，画布据此渲染等待态而不是就近加载旧 deck |
| 调用进行中 | 画布顶部有一条进度条：`正在〈绘制幻灯片〉… deck: xxx 12s`；页签标题出现一个呼吸圆点。会话里对应步骤的卡片也写明「画布已切到这一步」 | 宿主把每次调用的 `{tool, deckId, state, startedAt}` 发布到 `GET …/session/events?sessionId=`（SSE）与会话快照 `GET …/session?sessionId=` |
| deck 一落地 | 画布立刻切到这份 deck（工具返回前就切；工具卡片已经不在了也一样） | 客户端订阅会话进度流：`deckId` 变化 → 切 deck；`open` 的 `running` 阶段没有 deck 时先把旧 deck 移出屏幕 |
| 大纲 | 状态条先报「正在摆放 N 页骨架」，随后逐页（每页 160ms）落到缩略图栏 | `pptx_studio_plan` 的 `planning` + `slide(skeleton)` 事件 |
| 写一页 | 先报「正在准备第 N 页」，220ms 后整页内容才写入；画布上这页先以铅笔幽灵（虚线 + 低透明度 + 扫光）出现，再逐个元素落定（42–75ms/个）；幻灯片下方有 ✎ 绘制 HUD 与元素进度条，编辑区标题同时显示「逐页绘制中」徽标 | `slide(begin)` → `slide(content)`；客户端的逐元素 reveal 队列 |
| 加一个元素 | 同样先报「正在往第 N 页加元素」，元素再以幽灵→闪一下的方式出现 | `add_element` 的 `drawing(begin)` + `element` 事件 |
| 画完 | HUD 自动收起；期间编辑区、缩略图栏、功能区的位置和高度不跳 | 缩略图栏常驻（无页时放 3 个骨架格），编辑区卡片 `min-height` 兜底 |

- **「切到画布」的范围**：不只是 `pptx_studio_open`。模板分解、调研、HTML 效果稿、大纲、写页、改元素、
  校验、导出——每一步都有自己的卡片，都会把画布拉到前台并指向该步写的 deck；用户在聊天里翻旧消息时
  卡片不会反向抢焦点（只有正在跑的那次调用会切）。
- **不会显示上一份 pptx**：新一次 `pptx_studio_open` 在 preparing 阶段就用 `pending` 打开画布，
  画布显示该次调用的等待骨架（带工具名/耗时）而不是最近一次落盘的 deck；旧 deck 的 `/state` 响应若在
  切换后才回来会被丢掉（`deckIdRef` 比对），不会把旧画布盖回来。`pending` 参数带时间戳，页面刷新后
  过期的 pending 会自动失效、退回正常选 deck，不会把画布卡死在空态。
- `GET …/session` 与 `…/session/events` 不可达时（老宿主/老客户端）一切照旧：卡片走 `openTab` 的
  navigation params 切换 deck，只是少了顶部那条进度条。
- 每个元素落定时都有入场动画：文本是自左向右「写出来」（clip reveal），图形/图片/表格是淡入 + 轻微缩放。
- 偏好减少动效（`prefers-reduced-motion`）时动画关闭，但幽灵→落定的节奏与进度提示仍在。
- 缩略图不是另一套简化渲染：它和画布用同一份元素模型、页面自己的底色与页尺寸（16:9 / 4:3 / 16:10 / A4），
  先按 0.5× 完整布局整页、再整体缩到 88px 宽，所以深色稿的缩略图也是深色，标题字号、表格列宽、图表与图片位置
  都与实际页成同一比例（旧实现把元素直接按 1.65px 字号硬画在 88×60 的白底里，缩略图与实际页对不上）。
- 切页时「跟随最新」默认开；用户点过缩略图就会暂停跟随，避免抢镜头。
- 客户端改动要刷新页面（或让 HMR 重新载入）才会生效；宿主改动用重载插件行（`set_plugin`）即可。

### 编辑区 / 功能区的分界

画布按上下两块组织，并各自带标签、边框和底色：

- **编辑区**：套在带 `--ps-border-strong` 边框的卡片里，标题行写「编辑区 · 幻灯片编辑 / 调研 / HTML 预览」，
  里面是缩略图 + 幻灯片画布或阶段面板；画布外是浅灰底，幻灯片纸面保持白/自有背景。
- **分隔条**：中间一条 `编辑区 / 功能区` 的可见分隔线，把两块区域明确断开。
- **功能区**：独立卡片，标题行写「功能区」，里面是状态、缩放、页操作、插入、属性、动效、主题和意见；
  内容超高时只滚动功能区，不会挤压编辑区。

### 快捷键

焦点在画布上时生效；`Ctrl+S` 在输入框里也会保存，其余编辑快捷键在 `input / textarea / select`
里保留系统原生行为，不会抢走文字撤销/复制。

| 快捷键 | 动作 | 说明 |
|---|---|---|
| `Ctrl+S` | 保存 | 写进 `.pptx-studio/<deckId>/deck.pptx`（按钮「保存」），面板回报文件名与体积 |
| `Ctrl+Shift+S` | 下载 / 另存 | 弹宿主目录框选位置，导出到该目录（同名自动改名，绝不覆盖） |
| `Ctrl+Z` / `Ctrl+Shift+Z` / `Ctrl+Y` | 撤销 / 重做 | 与工具栏 ↶ ↷ 同源 |
| `Ctrl+X` / `Ctrl+C` / `Ctrl+V` | 剪切 / 复制 / 粘贴元素 | 支持多选；内部剪贴板 + 系统剪贴板 JSON 双通道 |
| `Ctrl+A` | 选中本页全部元素 | 多选后显示虚线框和批量剪切/删除面板 |
| `Ctrl+D` | 复制一份 | 多选时批量复制 |
| `Ctrl+M` | 新增一页 | 同 `+ 页` |
| `Delete` / `Backspace` | 删除选中元素 | 多选时一次批量删除，一步撤销 |
| `← ↑ → ↓` / `Shift` + 方向键 | 微调位置 | 0.05in / 0.2in |
| `Ctrl+0` / `Ctrl+=` / `Ctrl+-` | 适应窗口 / 放大 / 缩小 | 与缩放按钮一致 |
| `Esc` | 取消选择 / 关闭浮层 | |
| `?` | 打开/关闭快捷键帮助 | 功能区右上角也有「快捷键」按钮 |

### 阶段条与三个前置面板

- 顶栏下方阶段条 `⓪ 模板 / ① 调研 / ② HTML / ③ 幻灯片`：已完成打勾，点击随时切换查看；默认跟随当前阶段。
  模板模式下 `⓪` 未完成时阶段闸门停在 `stage-template`，画布默认停在模板面板；空白 deck 的模板面板显示「没有使用模板」。
- **模板面板**：模板页数、渲染引擎、文字是否已隐去，配色/字体/文字颜色；每页角色（封面/目录/章节/内容/结尾）、
  可填字段（含 `≤N 字` 字数预算）与配色；以及本 deck 的页映射（第 N 页来自模板第几页、已填哪些字段）。
  模板还没分解时显示「正在分解模板…」，分解失败可点「刷新模板」重新拉取。
- **调研面板**：分析摘要、受众/目标/语气/约束、编号大纲、关键结论、可点击来源、配色色板与字体方向。
- **HTML 预览面板**：服务端 CSP `sandbox` + iframe `sandbox=""` 双保险，脚本和外链样式不执行；
  提供刷新、新窗口打开、下载 `preview.html`。预览是静态视觉稿，图片用内联 data URL 最稳。
  面板按视觉稿的设计尺寸（`pptx_studio_preview` 的 `width/height`，或从 HTML 的 `.slide` 宽度推断）
  **等比缩放到侧栏宽度**，不拉伸、不把整页压扁；多页稿在 iframe 内滚动。

### 幻灯片编辑

| 动作 | 怎么操作 | 落到哪里 |
|---|---|---|
| 选中元素 | 单击 | 蓝框 + 8 个缩放手柄 + 属性面板 |
| 移动 / 缩放 | 拖元素本体 / 拖手柄（`Shift` 锁一个方向） | `POST /edit` → `element.patch {box}` |
| 改文字 | 双击元素，或在属性面板改 | `element.patch {text}` |
| 改样式 | 字号、加粗/斜体/下划线、对齐、文字色/填充色、X/Y/W/H | `element.patch {style}` / `{box}` |
| 阴影 | 属性面板「阴影」开关（导出为 `a:outerShdw`） | `element.patch {shadow}` |
| 入场动画 | 属性面板选 出现/淡入/擦除/缩放 + 方向/时长 + 预览 | `element.patch {anim}` |
| 层级 / 复制 / 删除 | 面板右上 ⤒ ⤓ ⧉ ✕，或 `Delete`、`Ctrl+D` | `element.reorder` / `duplicate` / `patch{remove}` |
| 微调位置 | 方向键 0.05in，`Shift` 0.2in | `element.patch {box}` |
| 撤销 / 重做 | ↶ ↷，或 `Ctrl+Z` / `Ctrl+Shift+Z` | 模型 diff 后一次性补发 ops |
| 加页 / 删页 / 调序 | 页操作行 `+ 页` `←` `→` `🗑` | `slide.add` / `slide.remove` / `slide.move` |
| 页面设置 | 「页面/动效」：切换效果、速度、方向、点击/自动换页、背景色、隐藏页、演讲者备注、模板版式 | `slide.patch` |
| 预览切换 | 页面设置里「预览切换」：画布按当前转场播放，同时带动画元素按 order/delay 入场 | 仅画布预览，不改模型 |
| 主题与尺寸 | 「主题」：5 个主题色、字体、16:9/4:3/16:10/A4 预设 | `deck.theme` / `deck.size` |
| 插入元素 | 「插入文本 / 图片 / 矩形 / 表格 / 图表」；图片也可拖入、`Ctrl+V` | `POST /upload` 或 `element.add` |
| 改标题 | 双击顶栏标题 | `deck.title` |
| 意见反馈 | 点元素 / 框选区域 / 阶段面板写意见 → **发给 DSH** | 原生输入框 `setDraft+submit`；没有输入框时由宿主注入 |

反馈消息带 `deck / slide / element id / box / 当前文本 / 用户意见 / 期望动作 / 阶段`。
阶段面板的反馈不带元素，宿主会提示模型先改对应阶段。

### 尺寸口径（与 PowerPoint 对齐）

画布与 `.pptx` 用同一套换算：**1in = 96 CSS px**，字号 = `pt × 96 / 72 × zoom`。
所以 100% 缩放时 13.333×7.5in 的页面正好 1280×720px，30pt 的标题是 40px——
不是「1pt = 1px」的近似值。默认「适应窗口」，可切 `100%` 或 ±10% 步进。

### 主题

画布的控件全部读 `--dsw-alias-*` 令牌，随 DSH 明暗主题变：主按钮 =
`button-primary-fill` 底 + `label-primary-foreground` 字；文字 `label-primary/secondary`、
描边 `border-l1`、悬停 `interactive-bg-hover`、状态点 `state-*`。幻灯片纸面颜色来自
`slide.background`（默认白），元素颜色来自 deck 模型。

## 幻灯片能力与导出实现

| 能力 | deck 模型 | 画布 | `.pptx` |
|---|---|---|---|
| 幻灯片转场 | `slide.transition {type, speed, direction, advanceOnClick, advanceAfterMs}` | `data-fx` CSS 动画预览 | `<p:transition spd advClick advTm>` + `p:fade/push/p:wipe/p:split/p:cover/p:pull/p:cut/p:zoom/p:dissolve/p:circle/p:random` |
| 元素入场动画 | `element.anim {effect, direction, duration, delay, order}` | 逐元素 CSS 动画 + 预览按钮 | `<p:timing>` 点击序列；支持 appear/fade/wipe/zoom |
| 演讲者备注 | `slide.notes` | 页面设置里的备注框 | `notesSlide` 文本 |
| 页面背景 | `slide.background {color}` | 画布纸面底色 | 幻灯片背景 solid fill |
| 隐藏页 | `slide.hidden` | 缩略图 ⊘ + 半透明 | `p:sld show="0"` |
| 主题 | `deck.theme {colors, fonts}` | 主题面板 + 默认字体/色板 | 导出默认字体与文字色 |
| 尺寸 | `deck.size` | 预设切换 | `prs.slide_width/height` |
| 原生图表 | `element.chart {kind, categories, series}` | SVG/HTML 近似预览 | `ppt/charts/chartN.xml`（column/bar/line/pie/doughnut/area） |
| 表格 | `element.table {rows, header}` | HTML 表格 | 原生表格；表头加粗/浅色底 |
| 形状 / 圆角 / 阴影 | `rect`, `radius`, `shadow` | CSS | 自选图形 + `a:outerShdw` |
| 富文本段 | `paragraphs[].runs` | span 级字号/颜色/粗斜体/下划线 | run 级 `a:rPr` |
| 模板版式 | `slide.layoutIndex` | 页面设置里选版式 | `p:sldLayout` 引用（模板模式） |

导出时 Python 侧会回读文件并返回 `features`：转场页数、timing 页数、动画数、备注页数、
隐藏页数、图表数；工具收据里显示成「切换 3 页/动画 3 处/备注 3 页/图表 1 个」。
需要兼容性回退时可直接给 `pptx_build.py` 传 `--no-effects`（不带转场/timing）。

## 工具面

| 工具 | 说明 |
|---|---|
| `pptx_studio_open` | 建/续画布；`template` 进模板模式（沿用母版与版式） |
| `pptx_studio_template` | 模板模式第 0 步：分解模板（每页角色、去字干净底图、可填文字字段）；`list:true` 只读、`render:false` 只解析结构 |
| `pptx_studio_research` | 第一步：交需求分析/调研 brief（摘要、受众、目标、大纲、结论、来源、配色/字体） |
| `pptx_studio_preview` | 第二步：交自包含 HTML 视觉稿（可选 `width/height` 设计画布尺寸），写入 `preview.html`，画布用沙箱 iframe 等比缩放展示 |
| `pptx_studio_deck` | 整份 deck 的标题、画布尺寸、主题色/字体 |
| `pptx_studio_plan` | 交页大纲，画布先出骨架卡片 |
| `pptx_studio_slide` | 整页写入/替换（1-based index，可带 transition/background/notes/hidden） |
| `pptx_studio_add_element` | 追加单个元素（可带 chart/anim/shadow） |
| `pptx_studio_patch` | 改/删一个元素（修正闭环的主力） |
| `pptx_studio_inspect` | 读回模型（页 id、元素 id/box/文本、阶段、转场/备注） |
| `pptx_studio_verify` | 一致性校验：模板底图/字段位置/字号 + 幻灯片文字 vs HTML 模版；不一致时 export 拒绝导出 |
| `pptx_studio_export` | 导出真实 `.pptx`，返回路径/页数/字节/特效统计 |

模型可见的收据都很短（`deck=…；阶段=…；3 页；…`），deck 本体走画布与 SSE，不烧 token。
写页/加元素时模型层是**宽容**的（磁盘旧 deck、画布补丁都要能读），所以会在收据里回一行
`参数已修正 N 处：…`；参数干净时没有这一行。

## 数据与落盘

```
<cwd>/.pptx-studio/<deckId>/deck.json               # 模型（唯一真相，含 pipeline/theme/特效）
<cwd>/.pptx-studio/<deckId>/deck.pptx               # 最近一次导出
<cwd>/.pptx-studio/<deckId>/preview.html            # 第二步的 HTML 视觉稿（模板模式是生成结果的副本）
<cwd>/.pptx-studio/<deckId>/assets/                 # 元素引用的图片
<cwd>/.pptx-studio/<deckId>/template.json           # 打开时 pptx_inspect 的版式摘要（母版/占位符）
<cwd>/.pptx-studio/<deckId>/template/template.json  # ⓪ 分解出的模板 spec（页/角色/字段/配色，v2）
<cwd>/.pptx-studio/<deckId>/template/plates/        # 去字干净底图（导出时铺满整页）
<cwd>/.pptx-studio/<deckId>/template/thumbs/        # HTML 模版内联用的缩略底图
<cwd>/.pptx-studio/<deckId>/template/reference/     # 原样整页渲染（对照用）
```

SSE 事件：`deck`（全量）`pipeline`（阶段变化）`slide`（新增/替换）`element`（新增/替换/删除）
`status`（绘制进度：`ready` / `planning` / `drawing`，其中写页与加元素前会先发 `begin:true` + `op`）
`saved`（导出，含 `features`）`feedback`（用户反馈已送达）`deck-removed`。

## 画布 ↔ 宿主路由

| 路由 | 用途 |
|---|---|
| `GET /dsh-pptx-studio/api/decks?sessionId=` | 画布列表：内存里的 + **磁盘上有 `deck.json` 但本进程还没加载的**（重启/重载后仍能找回，`fromDisk:true`） |
| `GET …/state?deck=&sessionId=` | 全量模型（含 pipeline/theme）+ 模板摘要 + `preview` 路径 + 工作区 `cwd` |
| `GET …/events?deck=` | SSE 增量流 |
| `GET …/session?sessionId=` | 会话当前/最近一次 pptx 调用的进度快照：`{task: {tool,label,deckId,state,message,startedAt,seq}}`；从未调用过为 `null` |
| `GET …/session/events?sessionId=` | 会话级 SSE：先推一份快照，之后每次调用开始/结束推一条 `{type:'task', task}`。画布据此把「显示上一份 deck 的 tab」切到正在写的 deck，并在顶部显示「正在〈某步〉」进度条 |
| `GET …/preview?deck=` | 第二步 HTML 视觉稿；`?download=1` 下载；响应带 CSP `sandbox`（无脚本/无外链样式） |
| `POST …/edit` | 画布编辑 `{deckId, ops:[…]}`：`element.patch/replace/add/duplicate/reorder`、`slide.add/remove/move`、`slide.patch`、`deck.title/theme/size` |
| `POST …/upload` | 图片：`{deckId, slideId, dataUrl, box}` → 内容哈希去重后写入 `assets/` |
| `GET …/asset?deck=&path=` | 元素图片（`assets/xxx.png` 按 deck 目录解析；越界 403） |
| `POST …/feedback` | 「阶段 + 页 + 元素 + bbox + 意见」→ 会话用户消息 |
| `POST …/deck/delete` | 删画布：先清内存副本与防抖写入，再删目录；id 必须是 `deck_<日期>_<hash>` |
| `POST …/export` | 导出：`{deckId, path?}` 写进 deck 目录或工作区内路径；`{deckId, dir}` 写进用户用宿主目录框选中的**任意绝对目录**（不合规 400 `bad-directory`，同名自动 `名称 (1).pptx`） |
| `GET …/download?deck=` | 附件形式的 `.pptx`（未导出时按需先导出）；现在是画布「下载」在**没有**目录选择器时的兜底通道（页面 fetch 成 blob 再交给浏览器，不再用 `window.open`） |

## 依赖与配置

- **必需**：能 `import pptx` 的 Python。探测顺序：插件配置 `python` → `$PPTX_STUDIO_PYTHON`
  → `$DSH_HOME/dsh-runtimes/*/dependencies/python` → PATH。
- 没有 python-pptx 时：画布/调研/HTML 预览/大纲/编辑仍可用，只有导出不可用（工具返回
  `python-unavailable` 与下一步提示）。

在 profile 的 `cordis.patch.yml` 里给本插件行加 `config`：

```yaml
- insert:
    - id: dsh-pptx-studio
      name: 'dsh-pptx-studio'
      config:
        python: 'C:/path/to/python.exe'   # 可选：固定解释器
```

## 安装 / 卸载 / 更新

```bash
dsh plugin --profile <p> add link:<本仓库绝对路径>
# Electron 管理的 profile（如 desktop）用：
#   plugin_manager(action:'install_bundle', target:'link:<绝对路径>')
# 之后改 host 代码，免重启重载：
#   plugin_manager(action:'set_plugin', target:'include:dsh-pptx-studio', enabled:true)
```

卸载：`plugin_manager(action:'remove_bundle', target:'dsh-pptx-studio')`，并删掉 profile 里的
`dsh-pptx-studio` 依赖。**改 profile 前先备份** `~/.dsh/profiles/<p>/{package.json,cordis.patch.yml,cordis.yml}`。

## 测试

```bash
npm test     # model / template / store / python / skills / e2e / client 七组
```

- `test/model.test.js`（26 项）：模型 v2 的主题/转场/动画/阴影/图表/流水线归一化与阶段推导。
- `test/template.test.js`（27 项）：模板 spec 归一化、`checkMapping` 需求映射校验、`renderTemplateHtml`、
  `buildSlideFromTemplate`、`previewConsistency`；不依赖 Python / PowerPoint。
- `test/store.test.js`（7 项）：deck 目录存取、防抖落盘、路径越界。
- `test/python.test.js`（7 项）：真跑 python-pptx；回读 zip 校验 OOXML 子元素顺序、`p:transition`、
  `p:timing`、`a:outerShdw`、notes、隐藏页、chart 部件；`--no-effects` 回退。
- `test/skills.test.js`（14 项）：frontmatter 解析、`readSkill` / `loadBundledSkills`、
  `skillRegistration`（含必填 `source`）与 `registerBundledSkills` 注入/撤回。
- `test/e2e.test.js`（31 项）：假 `ctx` 真跑 `apply()` → 12 个工具 + 全部 HTTP 路由；重点覆盖
  「未调研/未预览被拦住 → research → preview → plan/slide → export features」、画布编辑、
  `slide.patch`/`deck.theme`/`deck.size`、预览 CSP、上传去重、资产越界、导出字节、
  「导出到用户选的绝对目录且同名不覆盖 / 非法目录 400」；模板模式覆盖
  「未分解被拦 → 分解 → 需求映射拒收/通过 → 同一映射生成 HTML → 按模板建页 → verify/export 拒绝漂移」；
  另有「每页写入前先发 begin 状态、planning 先于骨架页、写页真的停顿 ≥180ms」的可见性回归。
- `test/client.test.js`（41 项）：按 Web 载体的方式加载 `client.js`，渲染画布，验证编辑区/功能区
  分界、阶段条与模板/调研/HTML 面板、iframe 沙箱与等比缩放、转场设置与预览、主题/尺寸、图表插入、
  元素动画/阴影、保存/剪贴/撤销/新增页/多选快捷键、输入框原生按键、主题令牌与 1in=96px；
  以及「下载按钮常显、点它走宿主目录框带 `dir` 导出、取消不导出、无目录服务时退化到 blob 下载」、
  「无 deck 时编辑区已在」「SSE 新页先是幽灵再逐个落定 + HUD」「工具卡片 start 阶段就开画布」
  「tab 参数换 deck」「等待时轮询 deck」；缩略图按页尺寸/底色出图、整页缩放而不是 1.6px 小字硬画。
- 真 DOM 冒烟（真 React 18 + jsdom，65 项）：点选、拖动提交 box、拖角缩放、双击改字、缩略图切页、
  方向键微调、框选、插图上传、反馈 sendDraft+submit、导出/下载/关闭；外加无 deck 骨架 shell、
  工具卡片 start 阶段自动打开、SSE 新页幽灵→逐个落定→HUD 收起、生成中编辑区不消失。
  脚本在 `%TEMP%/dsh-pptx-smoke/smoke.mjs`（依赖本机其它 checkout 的 node_modules，故不入库）。

## 已知边界

- **HTML 预览是静态沙箱**：脚本、外链 CSS/字体、相对路径图片不会加载；请用内联 `<style>`、
  data URL 或 `assets/` 绝对 API 地址。它用于确认视觉方向，不是可交互网页。
- **入场动画只写基础效果**（appear/fade/wipe/zoom + 时长/延迟/顺序），不写 PowerPoint 的
  motion path、触发器、退出动画和复杂时间轴；复杂动效请在 PowerPoint 里补。
- **转场覆盖常用集合**（fade/push/wipe/split/cover/pull/cut/zoom/dissolve/circle/random）；
  版本特有转场（morph 等）不承诺。
- 组合形状 / SmartArt / group 仍是占位（标注「近似」）；图表、表格、阴影、富文本段是原生导出。
- 画布是**屏幕近似**：字体度量、行距与 PowerPoint 不完全一致；导出后请以 PowerPoint 为准。
- 主题/字号等默认值只影响未显式设置颜色的元素；已有元素颜色不会被整体重刷。
- `openTab` 需要「有会话在屏幕上」；`openTabIn` 只对已挂载（adopted）的会话有效，两者都不行时
  卡片会先打开会话再重试一次，仍失败才提示。
- 编辑以「本地先改、随发随存」为准：宿主不可达时改动留在画布上并提示，不会静默丢失。
- 只在当前页编辑；跨页批量改（如「全部页标题字号统一」）仍走对话让模型用工具改。
- 反馈默认走原生输入框管线（`setDraft` + `submit`，等同用户按 Enter）；没有 composer 时才回落到 Host 注入。
- **模板底图是整页图片**：⓪ 之后每页的装饰/背景都以 `template/plates/*.jpg` 整页图参与导出，
  页内装饰**不能单独选中改**；要改装饰回 PowerPoint 改模板，再重新跑 `pptx_studio_template`。
  能改的是字段文字和 `extra` 图元。
- **没装 PowerPoint 时的降级**：分解先试 PowerPoint COM；失败回退 Pillow 近似（只拼图片与纯色形状，
  渐变/文字不参与，不是原版渲染；收据会提示底图非 PowerPoint 渲染，成品需人工核对），或传
  `render:false` 只解析结构（没有底图，导出是纯色底；此时 `verify` 把缺底图记为 warning 而不是 issue）。
- **逐字拆分的标题会合并成一个字段**：模板把标题拆成单字文本框（方便字距）时，字段表只暴露一个
  `perChar` 字段，重填按整行居中，不会散成单字、也不会沿用原来的单字间距。
- **继承色的文字按底图取色**：模板里没写显式颜色的文字，会按底图在字段框内的亮度自动选深/浅色
  （`apply_contrast_colors`），避免「深色字压在深色牌匾上」看不见；想要别的颜色就显式传 `style.color`。
- **模板可以在工作区外**：`pptx_studio_open(template=...)` 接受绝对路径（如 `~/Downloads/xxx.pptx`），
  宿主会按内容哈希把它复制到 `.pptx-studio/_templates/` 再使用，后续步骤不会再去读工作区外的文件。
