# dsh-pptx-studio 使用说明书

一句话说清它做什么：**你说要一份 PPT，它在右侧开一块画布，先给你看调研结论与视觉稿，你点头之后才一页一页画成真正可编辑的 `.pptx`。** 不必等导出才知道长什么样。

这份说明书面向使用者，讲怎么装、怎么开口、画布上怎么下手、出问题怎么查。想了解实现细节与设计取舍，看 [README](../README.md)；想了解为什么必须走插件而不能让模型自己写脚本，看 [REQUIREMENTS.md](../REQUIREMENTS.md)。

## 目录

1. [这是什么](#这是什么)
2. [适合谁用](#适合谁用)
3. [安装](#安装)
4. [五分钟上手](#五分钟上手)
5. [在对话里怎么说](#在对话里怎么说)
6. [画布怎么用](#画布怎么用)
7. [元素级编辑](#元素级编辑)
8. [模板模式](#模板模式)
9. [快捷键](#快捷键)
10. [工具速查](#工具速查)
11. [产物放在哪里](#产物放在哪里)
12. [Python 依赖与配置](#python-依赖与配置)
13. [常见问题](#常见问题)
14. [已知边界](#已知边界)
15. [自查与升级](#自查与升级)

## 这是什么

一个 DSH 插件。装上之后：

- 多出一块**右侧 PPTX 画布**，页签标题就是这份 deck 的名字；
- 多出 **12 个模型工具**（`pptx_studio_*`），模型按「调研 → 视觉稿 → 幻灯片」三步走；
- 自带一个**技能**，强制模型别绕过插件去手写 python-pptx 脚本（绕过就得不到画布，也得不到转场、动画、备注、原生图表）。

它跟「一句话直接吐个 pptx」的区别在于：中间有两次可拦截、可修改的机会——一次在调研分析，一次在 HTML 视觉稿。方向不对时，你改这两处，比改二十页幻灯片便宜得多。

## 适合谁用

- 要把调研、汇报、方案讲清楚，而不只是「凑几页字」；
- 手上已有公司模板（`.pptx`），希望沿用母版、配色、版式，只换内容；
- 需要对成品做**元素级微调**（字号、位置、颜色、入场动画），又不想开 PowerPoint；
- 需要转场、演讲者备注、原生图表这些「真 PPT 才有的东西」，而不是一堆图片拼起来的幻灯片。

如果你的需求只是把一段文字塞进几页白底幻灯片，这插件也能干，只是大材小用。

## 安装

给某个 profile 装上（以本仓库绝对路径为例）：

```bash
dsh plugin --profile <profile> add link:<本仓库的绝对路径>
```

Electron 管理的 profile（例如桌面版 `desktop`）走插件管理器：

```
plugin_manager(action: 'install_bundle', target: 'link:<本仓库绝对路径>')
```

改过宿主侧代码之后免重启重载：

```
plugin_manager(action: 'set_plugin', target: 'include:dsh-pptx-studio', enabled: true)
```

改动只涉及 `client.js`（画布）时，重载插件不足以让浏览器换掉旧脚本，需刷新页面（`Ctrl+R`）。

卸载：

```
plugin_manager(action: 'remove_bundle', target: 'dsh-pptx-studio')
```

再删掉 profile 的 `package.json` 里 `dsh-pptx-studio` 依赖键即可。**动 profile 前先备份** `~/.dsh/profiles/<profile>/{package.json, cordis.patch.yml, cordis.yml}`。

## 五分钟上手

第一份 deck，可以这样说：

```
用 pptx_studio_open 开一块新画布，标题「2026 年度复盘」。
```

说过这句，右侧就会亮起画布（不等工具返回，调用一开始就开），顶部显示一份等待骨架。接着交给模型继续：

```
先用 pptx_studio_research：受众是管理层，目标是批准明年的预算，给 8 页大纲。
```

画布切到「① 调研」面板，你能看到摘要、受众、目标、大纲、来源与配色。觉得方向对了，说「继续」；不对就直说要改哪一条。

```
pptx_studio_preview：出一版 16:9 的视觉稿，深色底，青色做强调色。
pptx_studio_plan：按大纲排 8 页。
然后逐页 pptx_studio_slide，写完导出到 reports/2026-review.pptx。
```

画布会在「② HTML」停下让你确认视觉方向，再到「③ 幻灯片」逐页逐元素画出来。导出完成时工具收据会写明路径、页数、字节与特效统计（`切换 3 页/动画 3 处/备注 3 页/图表 1 个`）。

## 在对话里怎么说

不必背工具名，说人话即可；但知道工具名能让模型少猜。常用说法与落点：

| 你想做的事 | 可以这么说 | 落到哪个工具 |
|---|---|---|
| 开一份新 deck | 「开一块画布，标题是……，模板用 xxx.pptx」 | `pptx_studio_open` |
| 先做需求分析 | 「先调研，受众是……，目标是……」 | `pptx_studio_research` |
| 看视觉稿 | 「出个 HTML 效果稿，看方向」 | `pptx_studio_preview` |
| 排大纲 | 「排 8 页，封面/背景/数据/结论……」 | `pptx_studio_plan` |
| 写一页 | 「第 3 页写数据对比，带柱状图」 | `pptx_studio_slide` |
| 加一个元素 | 「封面右下角加个徽标图片」 | `pptx_studio_add_element` |
| 改一处 | 「第 2 页标题字号改 32，文字改成……」 | `pptx_studio_patch` |
| 换整份风格 | 「主色改成 #22D3EE，尺寸换 4:3」 | `pptx_studio_deck` |
| 核对模型现状 | 「现在几页、每页有哪些元素」 | `pptx_studio_inspect` |
| 导出 | 「导出到 reports/xxx.pptx」 | `pptx_studio_export` |

三处**闸门**会拦住抢跑：没调研就排页、没视觉稿就写页、模板模式没分解就写页——都会被拒并回一句提示。这是有意的，宁可拦住，也不让你的幻灯片在方向未定时白画。

## 画布怎么用

画布分上下两块：上面是**编辑区**（缩略图 + 幻灯片纸面，或模板/调研/HTML 阶段面板），下面是**功能区**（状态、缩放、页操作、插入、属性、动效、主题、意见）。两块各有标题与边框，中间一条可见分隔线；内容超高时只滚功能区，不挤压编辑区。

顶栏给你的东西：

- **标题**：双击可改，改的即 deck 标题；
- **页数与页选择器**：跳页；
- **保存**：写进 `.pptx-studio/<deckId>/deck.pptx`，回报文件名与体积；
- **下载**：弹宿主目录框让你选落点（同名会自动改名，绝不覆盖）；没有目录服务时退回浏览器下载；
- **阶段条** `⓪ 模板 → ① 调研 → ② HTML → ③ 幻灯片`：做完打勾，随时点击回看，默认跟随当前阶段。

缩略图条常驻。新 deck 还没页时放三个骨架格，不会让编辑区塌下去。缩略图不是另一套简化画法，它跟大画布共用同一份元素模型，页面底色、页尺寸、字号比例都一致，所以深色稿的缩略图也是深色。

「跟随最新」默认开着，模型写到哪页就翻到哪页；你点过任意缩略图，它就暂停跟随，免得抢镜头，再点回开关继续。

缩放默认「适应窗口」，可切 100% 或按 ±10% 步进。换算口径与 PowerPoint 一致：**1 英寸 = 96 CSS px**，字号按 `pt × 96 / 72` 折算，所以 100% 时 13.333×7.5 英寸的页面正好 1280×720 像素。

## 元素级编辑

在画布上点一下元素，出现蓝框、八个缩放手柄与属性面板。之后：

| 想做什么 | 怎么下手 |
|---|---|
| 移动 | 拖元素本体 |
| 缩放 | 拖手柄；按住 `Shift` 锁一个方向 |
| 改文字 | 双击元素直接改，或在属性面板的文本框里改 |
| 改样式 | 字号、粗体、斜体、下划线、对齐、文字色、填充色、X/Y/W/H |
| 加阴影 | 属性面板「阴影」开关，导出为 PowerPoint 的 `outerShdw` |
| 加入场动画 | 选 出现/淡入/擦除/缩放，配方向、时长、延迟，可先预览 |
| 调层级 | 面板右上 ⤒ ⤓ 上移下移，⧉ 复制，✕ 删除 |
| 微调位置 | 方向键 0.05 英寸，`Shift` + 方向键 0.2 英寸 |
| 撤销 | ↶ 或 `Ctrl+Z`，重做 ↷ |
| 加页 / 删页 / 调序 | 页操作行 `+ 页`、`←`、`→`、🗑 |
| 页面设置 | 「页面/动效」里改转场、速度、方向、换页方式、背景色、隐藏页、备注、模板版式 |
| 换主题与尺寸 | 「主题」里挑 5 个主题色与字体，尺寸选 16:9 / 4:3 / 16:10 / A4 |
| 插入 | 「插入文本 / 图片 / 矩形 / 表格 / 图表」；图片也可拖入或 `Ctrl+V` 粘贴 |

改得不满意，不必打字描述：点元素，或框选一块区域，在底部输入框写意见，点**发给 DSH**，消息里会带你选中的页、元素 id、位置框、当前文字与阶段，模型据此改。

画布只改当前页；要「所有页标题统一字号」这类跨页批量改，仍在对话里说，让模型用工具改。

## 模板模式

手上已有公司模板时，这条路能保住母版、配色与版式。最小调用序列：

```
pptx_studio_open(template: 'templates/brand.pptx')
pptx_studio_template                                  # 先分解模板，这步不做会被拦
pptx_studio_research(brief: { summary: '…', outline: [
  { title: '封面', templateSlide: 0, fields: { title: '2026 年度复盘', subtitle: '管理层评审' } },
  { title: '结论', templateSlide: 3, fields: { title: '三件事', body1: '……' } },
] })
pptx_studio_preview                                    # 不传 HTML，宿主按同一份映射生成
pptx_studio_plan(slides: [{ title: '封面', templateSlide: 0 }, …])
pptx_studio_slide(index: 1, templateSlide: 0, fields: { title: '2026 年度复盘' })
pptx_studio_verify → pptx_studio_export
```

要点：

- `pptx_studio_template` 会把模板逐页拆成「角色 + 可填字段 + 配色」，并渲染两套底图——**原样整页**做对照，**去掉文字后的干净底图**在导出时铺满整页。产物落在 `<deck>/template/`。
- 需求大纲的每一项都要写 `templateSlide`（第几页模板）与 `fields`（填哪些字段）。字段填错、填多、超字数，会在这一步被拒，附一句为什么。
- **能改的**是字段文字与 `extra` 图元；**不能单独改**的是模板页里的装饰——底图是一整张图片。要动装饰，回 PowerPoint 改模板，再重跑一次 `pptx_studio_template`。
- 导出前 `pptx_studio_verify` 会逐页比对底图、字段位置与文字是否与 HTML 模版一致；不一致时导出直接拒绝（`template-inconsistent`），不写文件。
- 模板放在工作区外也行，脚本会把 `~/Downloads/xxx.pptx` 这类路径按内容哈希复制到 `.pptx-studio/_templates/` 再用。

## 快捷键

焦点在画布上时生效。`Ctrl+S` 在输入框里也会保存；其余编辑快捷键在 `input / textarea / select` 内保留系统原生行为，不会抢走文字撤销与复制。

| 快捷键 | 动作 |
|---|---|
| `Ctrl+S` | 保存到 `.pptx-studio/<deckId>/deck.pptx` |
| `Ctrl+Shift+S` | 下载 / 另存（宿主目录框） |
| `Ctrl+Z` / `Ctrl+Shift+Z` / `Ctrl+Y` | 撤销 / 重做 |
| `Ctrl+X` / `Ctrl+C` / `Ctrl+V` | 剪切 / 复制 / 粘贴元素（支持多选） |
| `Ctrl+A` | 选中本页全部元素 |
| `Ctrl+D` | 复制一份 |
| `Ctrl+M` | 新增一页 |
| `Delete` / `Backspace` | 删除选中元素 |
| `← ↑ → ↓` | 微调 0.05 英寸 |
| `Shift` + 方向键 | 微调 0.2 英寸 |
| `Ctrl+0` / `Ctrl+=` / `Ctrl+-` | 适应窗口 / 放大 / 缩小 |
| `Esc` | 取消选择、关闭浮层 |
| `?` | 快捷键帮助 |

## 工具速查

| 工具 | 一句话说明 |
|---|---|
| `pptx_studio_open` | 新建或续接画布；给 `template` 就进模板模式 |
| `pptx_studio_template` | 模板第 0 步：分解每页角色、可填字段、去字底图 |
| `pptx_studio_research` | 第一步：交调研与需求分析（摘要、受众、目标、大纲、来源、配色） |
| `pptx_studio_preview` | 第二步：交自包含 HTML 视觉稿，画布沙箱展示 |
| `pptx_studio_deck` | 整份 deck 的标题、尺寸、主题色与字体 |
| `pptx_studio_plan` | 交页大纲，画布先出骨架卡片 |
| `pptx_studio_slide` | 整页写入或替换（1-based 页号，可带转场、背景、备注、隐藏） |
| `pptx_studio_add_element` | 追加单个元素（可带图表、动画、阴影） |
| `pptx_studio_patch` | 改或删一个元素，修正闭环的主力 |
| `pptx_studio_inspect` | 读回模型：页 id、元素 id 与位置、文本、阶段、转场、备注 |
| `pptx_studio_verify` | 一致性校验：底图、字段位置、字号、文字是否与 HTML 模版一致 |
| `pptx_studio_export` | 导出真实 `.pptx`，回报路径、页数、字节与特效统计 |

模型看到的收据都很短（`deck=…；阶段=…；3 页；…`），deck 本体走画布与事件流，不吃你的上下文额度。

## 产物放在哪里

一切落在当前工作目录的 `.pptx-studio/` 下：

```
<cwd>/.pptx-studio/<deckId>/deck.json                模型本体（唯一真相）
<cwd>/.pptx-studio/<deckId>/deck.pptx                最近一次导出
<cwd>/.pptx-studio/<deckId>/preview.html             第二步的 HTML 视觉稿
<cwd>/.pptx-studio/<deckId>/assets/                  元素引用的图片
<cwd>/.pptx-studio/<deckId>/template.json            打开时的模板版式摘要
<cwd>/.pptx-studio/<deckId>/template/template.json   分解出的模板 spec
<cwd>/.pptx-studio/<deckId>/template/plates/         去字干净底图（导出时铺满整页）
<cwd>/.pptx-studio/<deckId>/template/thumbs/         HTML 模版内联用的缩略底图
<cwd>/.pptx-studio/<deckId>/template/reference/      原样整页渲染，对照用
<cwd>/.pptx-studio/_templates/                       从工作区外导入的模板副本
```

`deckId` 形如 `deck_20261117_e7f1c1`。删画布走画布上的删除键（要按两次），会先清内存与防抖写入，再删目录；id 不合规的删除请求一律拒收。

重启 DSH 之后画布不会丢：画布列表除了内存里的 deck，还会把磁盘上有 `deck.json` 的也列出来，标 `fromDisk`。

## Python 依赖与配置

导出 `.pptx` 需要能 `import pptx` 的 Python（python-pptx）。探测顺序：

1. 插件配置里的 `python`；
2. 环境变量 `PPTX_STUDIO_PYTHON`；
3. `$DSH_HOME/dsh-runtimes/*/dependencies/python`；
4. PATH 上的 `python` / `python3` / `py`。

要固定解释器，在 profile 的 `cordis.patch.yml` 里给本插件行加 `config`：

```yaml
- insert:
    - id: dsh-pptx-studio
      name: 'dsh-pptx-studio'
      config:
        python: '<python 可执行文件的绝对路径>'
```

没有 python-pptx 时，画布、调研、HTML 预览、大纲、元素编辑照常能用，**只有导出不可用**——工具会返回 `python-unavailable` 并告诉你下一步。

## 常见问题

**导出报 `python-unavailable`。** 机器上没有能 `import pptx` 的 Python。装一个 `python-pptx`，或用上面的 `config.python` 指定解释器。

**模板底图是糊的、或配色不对。** 分解模板优先调 PowerPoint COM 渲染；没有 PowerPoint 时退回 Pillow 近似（只拼图片与纯色形状，渐变与文字不参与），收据会写明「非 PowerPoint 渲染」。要精确，装 PowerPoint，或在有 PowerPoint 的机器上分解后把 `template/` 目录带过来。

**`verify` 不通过，导不出来。** 多半是 HTML 模版与幻灯片文字漂移了：改过某页文字却只改了 HTML，或反过来。按 `verify` 报出的页与字段，重新用 `pptx_studio_slide` 写那一页，或重跑 `pptx_studio_preview` 让两边同源。

**「下载」没弹目录框。** 宿主没提供目录选择服务时会退到浏览器下载通道，文件到系统下载目录，不是保存失败。

**画布空白或停在上一份 deck。** 画布为了不让你误以为在看新 deck，新一次 `open` 会显示等待骨架而不是旧内容。若确实空着：确认那次调用是否报错，点页签标题旁的下拉重新选 deck，或刷新页面。

**改了代码没反应。** 宿主侧（`index.js`、`lib/`）用 `plugin_manager(action:'set_plugin', …)` 重载即可；客户端（`client.js`）要刷新页面，`pnpm run dev:web` 在跑时才会自动热更。

**HTML 预览里脚本、外链字体、相对路径图片都不动。** 预览是双保险沙箱（CSP `sandbox` + iframe `sandbox=""`），本就只用于确认视觉方向。图片请用内联 data URL 或 `assets/` 的 API 地址。

**画布上字的位置与 PowerPoint 里差一点点。** 画布是屏幕近似，字体度量与行距不可能与 PowerPoint 完全一致。以导出后的 PowerPoint 为准；差得多时改字号或位置，别指望像素级一致。

## 已知边界

- 入场动画只写基础效果（出现/淡入/擦除/缩放 + 时长/延迟/顺序），不含 motion path、触发器、退出动画与复杂时间轴；
- 转场覆盖常用十一类（淡入/推入/擦除/分割/覆盖/拉出/切出/缩放/溶解/圆形/随机），版本特有的 morph 不承诺；
- 组合形状与 SmartArt 目前是占位近似；图表、表格、阴影、富文本段是原生导出；
- 模板页的装饰不可单独选中修改（底图整页图片）；
- 逐字拆开的标题会合并成一个 `perChar` 字段，重填时整行居中，不还原单字间距；
- 只在当前页编辑，跨页批量改走对话；
- 主题与字号默认值只影响未显式设色的元素，已有颜色不会被整体重刷。

## 自查与升级

在本仓库里跑测试：

```bash
npm test    # model / template / store / python / skills / e2e / client 七组，共 153 项
```

其中 `python` 组会真跑 python-pptx 并回读 zip 校验 OOXML，故本机没有 python-pptx 时会失败——那是环境问题，不是代码问题。

升级插件：仓库里 `git pull` 之后重载插件行；若是 npm 装的，`npm i dsh-pptx-studio@latest` 再重载。版本改动见 [README 的已知边界与测试章节](../README.md#测试) 与仓库提交历史。
