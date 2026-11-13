# dsh-pptx-studio

在 DSH 右侧开一块 **PPTX 实时画布**：给模型一个 `.pptx` 模板或一句长提示词，deck 会一页页、一个元素一个元素地画出来；哪一页哪个元素不对，点一下、写句意见就发给 DSH 去改。

```
Agent 调 pptx_studio_open ──► Host 写 deck.json + 推 SSE ──► 右侧画布逐页绘制
        ▲                                                        │
        └──── pptx_studio_patch / add_element ◄── 用户点元素 + 写意见 ◄─┘
```

## 装了什么

| 文件 | 作用 |
|---|---|
| `index.js` | Host 半侧：7 个模型工具 + SSE + 反馈/导出/资源路由 + 消息注入 |
| `client.js` | 客户端半侧：右侧栏画布（缩略图/画布/属性/意见）+ `pptx_studio_open` 工具卡片 |
| `lib/model.js` | deck 模型：归一化、clamp、patch 合并、摘要 |
| `lib/store.js` | deck 存储：内存 + 400ms 防抖落盘 + 事件广播 |
| `lib/python.js` | python-pptx 桥：解释器探测 + 子进程 + JSON 解析 |
| `assets/pptx_build.py` | `deck.json → 原生可编辑 .pptx`（模板模式继承母版/版式/主题） |
| `assets/pptx_inspect.py` | 解析模板：尺寸、版式、占位符、主题色与字体 |
| `REQUIREMENTS.md` | 需求分析（场景、分级、技术取舍、验收标准、风险） |

## 怎么用（对话里）

```
用 pptx_studio_open 打开画布，标题「2026 年度复盘」，模板用 agent-ppt-research/…pptx
pptx_studio_plan：8 页，封面/背景/数据/结论/建议/…
pptx_studio_slide 第 1 页：……（逐页写）
pptx_studio_patch：第 3 页 s3-e2 的字号改成 24、文字改成「……」
pptx_studio_export：导出到 reports/2026-review.pptx
```

画布上：点元素 → 写意见（或点「太挤 / 字号太小 / 换配色」）→ **发给 DSH**。消息会带
`deck / slide / element id / box / 当前文本 / 用户意见 / 期望动作`，模型据此只改这一处。

## 工具面

| 工具 | 说明 |
|---|---|
| `pptx_studio_open` | 建/续画布；`template` 进模板模式（沿用母版与版式） |
| `pptx_studio_plan` | 交页大纲，画布先出骨架卡片 |
| `pptx_studio_slide` | 整页写入/替换（1-based index） |
| `pptx_studio_add_element` | 追加单个元素（最细粒度，逐个画出来） |
| `pptx_studio_patch` | 改/删一个元素（修正闭环的主力） |
| `pptx_studio_inspect` | 读回模型（页 id、元素 id/box/文本） |
| `pptx_studio_export` | 导出真实 `.pptx`，返回路径/页数/字节 |

模型可见的收据都很短（`deck=…；3 页；…`），deck 本体走画布与 SSE，不烧 token。

## 数据与落盘

```
<cwd>/.pptx-studio/<deckId>/deck.json      # 模型（唯一真相）
<cwd>/.pptx-studio/<deckId>/deck.pptx      # 最近一次导出
<cwd>/.pptx-studio/<deckId>/assets/        # 元素引用的图片
<cwd>/.pptx-studio/<deckId>/template.json  # 模板解析结果
```

SSE 事件：`deck`（全量）`slide`（新增/替换）`element`（新增/替换/删除）`status`（绘制进度）`saved`（导出）`feedback`（用户反馈已送达）。

## 依赖与配置

- **必需**：能 `import pptx` 的 Python。探测顺序：插件配置 `python` → `$PPTX_STUDIO_PYTHON` → `$DSH_HOME/dsh-runtimes/*/dependencies/python` → PATH。
- 没有 python-pptx 时：画布/大纲/编辑仍可用，只有导出不可用（工具返回 `python-unavailable` 与下一步提示）。

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
npm test     # model / store / python / e2e 四组，共 25 项
```

`test/e2e.test.js` 会用假 `ctx` 真跑一遍 `apply()` → 7 个工具全链路，并按 DSH 的
`snapshotJsonValue` 规则校验每个返回值（plain object/array、无 `undefined`、无 NaN、无 -0）。

## 已知边界

- 组合形状 / SmartArt / 图表用**占位**渲染与导出（标注「近似」），不承诺 1:1。
- 画布是**屏幕近似**：字体度量、行距与 PowerPoint 不完全一致；导出后请以 PowerPoint 为准。
- `openTab` 需要「有会话在屏幕上」；工具卡片在会话不在屏时会提示先点开会话。
- 反馈默认走原生输入框管线（`setDraft` + `submit`，等同用户按 Enter）；没有 composer 时才回落到 Host 注入。
