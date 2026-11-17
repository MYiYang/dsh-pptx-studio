---
name: pptx-studio
description: 生成或修改 .pptx 的唯一通道（dsh-pptx-studio 插件）：调研 → HTML 效果稿 → 可编辑幻灯片，右侧画布逐页可见；用户给了 .pptx 模板时先做模板分解（pptx_studio_template），需求逐页落到模板字段上，HTML 与幻灯片同源。当用户要求「做一份 PPT / 生成 pptx / 把调研做成演示文稿 / 改这份幻灯片」时使用；不要用 python-pptx 之类脚本从零拼 .pptx，其它讲 python-pptx 的 pptx 技能与本技能冲突时以本技能为准。
---

# PPTX 走插件，不走脚本

## 规则（优先于本机其它 pptx 技能）

1. 任何 `.pptx` 的**目的产物**——新建、改版、把调研或大纲落成幻灯片——只用本插件的 `pptx_studio_*` 工具。
   第一条命令就是 `pptx_studio_open`，让画布先立起来，用户全程看得见。之后每一步（模板分解 / 调研 / HTML 预览 /
   大纲 / 写页 / 改元素 / 校验 / 导出）都会自动把右侧画布切到该步正在写的 deck 并显示进度条，
   不需要让用户自己去点页签，也不会停在上一份 deck 上。
2. **禁止**用 python-pptx（`Presentation()`、`slide_layouts[...]`、`add_slide(...)`）或 `build_deck.py`
   一类脚本从零拼装幻灯片本体。那不是本插件的做法，也改不动。
3. `pptx-research-deck`、`office-pptx` 等技能里若出现 python-pptx 手搓路线，与本技能冲突时**以本技能为准**；
   它们的价值只在「核查已有 pptx」——溢出检测、PowerPoint COM 导图，可以借来用。
4. 例外仅两条：用户明说「不要用插件，手写脚本」；或 `pptx_studio_*` 不在当前工具表里。除此之外没有第三种。

## When to use

- 用户要一份新的演示文稿（调研汇报、方案、复盘、路演、课程……）。
- 用户要给已有的 `.pptx` 改几页、换主题、加动画或备注。
- 用户已有一份 HTML 视觉稿或一份调研结论，要落成可编辑的幻灯片。
- 用户**给了一份 `.pptx` 模板**、要求新内容沿用它的版式/配色 → **模板模式**：`pptx_studio_template`
  （模板分解）是第一步，不是 research。

## Steps

1. **`pptx_studio_open`**：定标题；有品牌模板就传 `template`（工作区内的 `.pptx`），没有就从空白 16:9 起。
   返回 `deckId` —— 之后每个工具都要带它。
   传了 `template` 就进**模板模式**：下一步不是 research，而是 `pptx_studio_template`（见下面的模板模式分支）。
2. **`pptx_studio_research`**：先 `web_search` 取数，再把 `brief` 交给画布：
   `{summary, audience, goal, tone, constraints, outline:[{title,intent,source}], facts:[{claim,source}], sources:[{title,url,note}], palette, fonts}`。
   用户在画布上先审分析，再谈视觉。
   模板模式下 `outline` 每项还要写 `templateSlide` 与 `fields`（`{字段id: 文案}`），否则会被拒收。
3. **`pptx_studio_preview`**：交一份自包含 HTML 视觉稿（每页一个 `.slide`，尺寸与 deck 一致，内联 CSS，
   不引外部脚本；图片用内联 data URL）。用户确认方向后再进幻灯片。
   模板模式下**不要自己写 HTML**：宿主会按同一份 `pages` 映射 + 模板底图/字段位置生成 HTML 模版。
4. **`pptx_studio_plan`**：给页大纲，`slides:[{title, intent}]`，画布逐页落骨架。模板模式每页可带 `templateSlide`。
5. **`pptx_studio_slide`**：按 `index`（1 起）写整页。元素坐标单位是**英寸**，
   `box=[x, y, w, h]`，页宽 13.333、页高 7.5；`type` 支持 `text/rect/image/line/table/chart`。
   页上可带 `transition`、`background`、`notes`、`hidden`；元素可带 `anim`、`shadow`。
   追加单个元素用 `pptx_studio_add_element`，局部改动用 `pptx_studio_patch`。
   模板模式下**不要自己算 box**：传 `templateSlide` + `fields`，位置/字号由模板字段决定；加图元用 `extra`。
6. **`pptx_studio_deck` / `pptx_studio_inspect` / `pptx_studio_verify`**：改画布尺寸、主题配色字体；
   回读当前模型核对；校验幻灯片与模板/HTML 模版是否一致（导出前跑）。
7. **`pptx_studio_export`**：导出真 `.pptx`（默认 `.pptx-studio/<deckId>/deck.pptx`，或传 `path`）。
   返回绝对路径、页数、字节数与特效统计；把路径交给用户，或用 `present` 交付。

## 模板模式分支（给了 .pptx 模板就走这条）

模板模式不是「换个母版」，而是把模板拆成可复用版式，需求、HTML、幻灯片三者同源：

1. **⓪ `pptx_studio_template`**（模板模式第一步）：分解模板。收据里有每页 `page/role` 与字段
   `id=示例（≤N字）`；`list:true` 只读、`render:false` 只解析结构（没有底图）。
2. **① `pptx_studio_research`**：`outline` 每项写成 `{title, templateSlide, fields}`——
   `templateSlide` 用模板第几页（0 起），`fields` 是 `{字段id: 文案}`，字段 id 必须来自上一步的 pages。
   页号越界、字段名不存在、超字数会被 `checkMapping` 拒收（`template-mapping-invalid` 会列出该页可用字段）。
3. **② `pptx_studio_preview`**：**不传 `html`**；`pages` 也可以省略（默认用 research 的映射）。
   宿主用同一份映射生成 HTML 模版（`preview.mode='template'`），用户在画布上确认。
4. **③ `pptx_studio_plan` → `pptx_studio_slide`**：`plan` 按同一顺序列页（每页可带 `templateSlide`）；
   每页调 `pptx_studio_slide` 传 `templateSlide` + `fields`，可加 `extra` 图元。
5. **④ `pptx_studio_verify` → `pptx_studio_export`**：先 verify 拿到 `consistent:true`；
   模板模式下若幻灯片文字与 HTML 映射对不上，`export` 返回 `template-inconsistent` 拒绝导出，
   回去改 preview 映射或重写对应页，不要绕过。

模板模式的三条硬约束：

- 不许自己写 HTML 预览（preview 会自动生成）；
- 不许自己算元素 box（位置由模板字段决定；传了 `templateSlide`/`fields` 时手写的 `slide.elements`
  会被忽略，加图元用 `extra`；不走模板字段自己写元素的页会被 verify 判为不一致）；
- 不许把模板页当普通空白页改：底图是整页图片，装饰不能单独改；要加原生图表/图片用 `extra`。

## Pitfalls

- 忘了带 `deck`：除 `open` 外每个工具都要求 `deckId`，报 `deck-not-found` 就是这一步。
- 新 deck 直接 `plan`／`slide`／`export` 会被阶段闸门挡住（`stage-research` / `stage-preview`）。
  按提示先补 research 与 preview；`force:true` 只在用户明确要「快速草稿」时才用，别拿它绕过流程。
- **给了模板却直接 research/plan**：模板模式闸门先返回 `stage-template`。第一步永远是 `pptx_studio_template`。
- **模板字段名写错会被拒**：`pptx_studio_research` 用 `checkMapping` 校验，字段不存在 / 页号越界 / 超字数
  返回 `template-mapping-invalid`，错误里列出该页可用字段 id；别猜字段名，先看 `pptx_studio_template` 的 pages 收据。
- **逐字标题会合并**：模板把标题拆成单字文本框时，字段表只暴露一个 `perChar` 字段，重填是整行居中，
  不会按原来的单字间距散开；不要试图逐字填。
- **没 PowerPoint 时的降级**：分解先试 PowerPoint COM，失败回退 Pillow 近似（只拼图片与纯色形状，
  渐变/文字不参与合成，底图是近似）或 `render:false` 只解析结构（没有底图，导出是纯色底）。
  导出仍可完成；Pillow 底图存在但不是原版（收据会提示），`render:false`/渲染失败时缺底图在 verify 里记
  warning。两种降级的成品都必须人工核对并告知用户。
- 单位混淆是最高频的错：`box` 用**英寸**，`style.size` 用 **pt**，预览 HTML 用 **px**。三者不要串。
- 元素越界：`x + w ≤ 13.333`、`y + h ≤ 7.5`（换过画布尺寸以 `pptx_studio_deck` 后的为准）。
- 一页塞太多：卡片正文 3–5 条、每条 ≤ 2 行；标题 28–40pt，正文 12–18pt。
- 图表要 `chart.categories` + `chart.series:[{name, values}]`，别只给一张图。
- 图片 `src` 用 deck 目录下的 `assets/...`（画布上传会落到这里），预览稿里才用 data URL。
- 导出后不要再用 python 重开改写；要改就回画布改，再导出一次——画布才是唯一事实来源。

## Verification

- 画布页数与 `plan` 一致，每页 `elements` 非空，标题不是占位文字。
- `pptx_studio_export` 返回 `path` 存在，页数等于画布页数；`present` 或直接给路径。
- 抽查 `pptx_studio_inspect`：备注、转场、图表系列都在；元素无越界。
- 模板模式：`pptx_studio_verify` 返回 `ok:true`（issues 为空）再 export；导出后每页底图与模板同款、
  文字位置/字号未漂移；`template-inconsistent` 一律视为「HTML 与幻灯片对不上」，先补齐再导。
- 工作区里**没有** `build_deck.py` 一类脚本落盘——落了就说明走错了路。
