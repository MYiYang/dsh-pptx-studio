---
name: pptx-studio
description: 生成或修改 .pptx 的唯一通道（dsh-pptx-studio 插件）：调研 → HTML 效果稿 → 可编辑幻灯片，右侧画布逐页可见。当用户要求「做一份 PPT / 生成 pptx / 把调研做成演示文稿 / 改这份幻灯片」时使用；不要用 python-pptx 之类脚本从零拼 .pptx，其它讲 python-pptx 的 pptx 技能与本技能冲突时以本技能为准。
---

# PPTX 走插件，不走脚本

## 规则（优先于本机其它 pptx 技能）

1. 任何 `.pptx` 的**目的产物**——新建、改版、把调研或大纲落成幻灯片——只用本插件的 `pptx_studio_*` 工具。
   第一条命令就是 `pptx_studio_open`，让画布先立起来，用户全程看得见。
2. **禁止**用 python-pptx（`Presentation()`、`slide_layouts[...]`、`add_slide(...)`）或 `build_deck.py`
   一类脚本从零拼装幻灯片本体。那不是本插件的做法，也改不动。
3. `pptx-research-deck`、`office-pptx` 等技能里若出现 python-pptx 手搓路线，与本技能冲突时**以本技能为准**；
   它们的价值只在「核查已有 pptx」——溢出检测、PowerPoint COM 导图，可以借来用。
4. 例外仅两条：用户明说「不要用插件，手写脚本」；或 `pptx_studio_*` 不在当前工具表里。除此之外没有第三种。

## When to use

- 用户要一份新的演示文稿（调研汇报、方案、复盘、路演、课程……）。
- 用户要给已有的 `.pptx` 改几页、换主题、加动画或备注。
- 用户已有一份 HTML 视觉稿或一份调研结论，要落成可编辑的幻灯片。

## Steps

1. **`pptx_studio_open`**：定标题；有品牌模板就传 `template`（工作区内的 `.pptx`），没有就从空白 16:9 起。
   返回 `deckId` —— 之后每个工具都要带它。
2. **`pptx_studio_research`**：先 `web_search` 取数，再把 `brief` 交给画布：
   `{summary, audience, goal, tone, constraints, outline:[{title,intent,source}], facts:[{claim,source}], sources:[{title,url,note}], palette, fonts}`。
   用户在画布上先审分析，再谈视觉。
3. **`pptx_studio_preview`**：交一份自包含 HTML 视觉稿（每页一个 `.slide`，尺寸与 deck 一致，内联 CSS，
   不引外部脚本；图片用内联 data URL）。用户确认方向后再进幻灯片。
4. **`pptx_studio_plan`**：给页大纲，`slides:[{title, intent}]`，画布逐页落骨架。
5. **`pptx_studio_slide`**：按 `index`（1 起）写整页。元素坐标单位是**英寸**，
   `box=[x, y, w, h]`，页宽 13.333、页高 7.5；`type` 支持 `text/rect/image/line/table/chart`。
   页上可带 `transition`、`background`、`notes`、`hidden`；元素可带 `anim`、`shadow`。
   追加单个元素用 `pptx_studio_add_element`，局部改动用 `pptx_studio_patch`。
6. **`pptx_studio_deck` / `pptx_studio_inspect`**：改画布尺寸、主题配色字体；或回读当前模型核对。
7. **`pptx_studio_export`**：导出真 `.pptx`（默认 `.pptx-studio/<deckId>/deck.pptx`，或传 `path`）。
   返回绝对路径、页数、字节数与特效统计；把路径交给用户，或用 `present` 交付。

## Pitfalls

- 忘了带 `deck`：除 `open` 外每个工具都要求 `deckId`，报 `deck-not-found` 就是这一步。
- 新 deck 直接 `plan`／`slide`／`export` 会被阶段闸门挡住（`stage-research` / `stage-preview`）。
  按提示先补 research 与 preview；`force:true` 只在用户明确要「快速草稿」时才用，别拿它绕过流程。
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
- 工作区里**没有** `build_deck.py` 一类脚本落盘——落了就说明走错了路。
