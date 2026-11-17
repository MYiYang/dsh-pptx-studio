/**
 * dsh-pptx-studio — client half (classic script consumed by the Web carrier).
 *
 * Two surfaces:
 *   1. a right-side canvas tab (`sidebarRightTabs` + `sidebar.right.pane.tab`)
 *      that draws the deck slide by slide over SSE, lets the user select /
 *      move / resize / retype elements, insert images, and send one structured
 *      note back to the session;
 *   2. a tool-call card for **every** `pptx_studio_*` tool — each one switches
 *      the canvas to the deck that step writes (and focuses it) as soon as the
 *      call appears, and reports the step's progress next to the result.
 *
 * Three invariants this file is built around:
 *   - **PowerPoint geometry.** One inch is 96 CSS px, so at 100% zoom a 30pt run
 *     is `30 * 96 / 72 = 40px`. Anything else makes the canvas disagree with
 *     the exported .pptx (text too small, boxes drifting).
 *   - **Theme by token.** Every piece of chrome reads `--dsw-alias-*`; in light
 *     mode `--dsw-alias-brand-primary` is near-black and in dark mode near-white,
 *     so a primary button must pair `button-primary-fill` with
 *     `label-primary-foreground` instead of a hard-coded white label.
 *   - **Never the previous deck.** A call that has not produced its deck yet
 *     opens the tab "pending" and shows that call's progress; the host's
 *     session feed (`/api/session`) names the deck a running call writes, so a
 *     tab left on an older .pptx switches the moment the deck exists.
 * @module dsh-pptx-studio/client
 */
window.__ModuleLoader__.load({
  id: 'dsh-pptx-studio',
  factory(require) {
    const React = require('react')
    const h = React.createElement
    const { useCallback, useEffect, useMemo, useRef, useState } = React

    /** Cordis ctx of this plugin instance (one per page); lets components call back. */
    let hostCtx = null

    /** Locale namespace. */
    const NS = 'pptxStudio'
    /** Tab definition id; also the key its body and title register under. */
    const TAB_ID = 'dsh-pptx-studio/canvas'
    /** Tab kind passed to `sidebarRight.openTab`. */
    const TAB_KIND = 'pptx-studio-canvas'
    /** Host route prefix (absolute path served by the Web carrier). */
    const API_PATH = 'dsh-pptx-studio/api'
    /** Thumbnails lay the page out at this zoom, then scale the whole page down. */
    const THUMB_DESIGN_ZOOM = 0.5
    /** CSS px per inch (the CSS reference pixel); 1pt = 96/72 px. */
    const PX_PER_IN = 96
    /** Points per inch. */
    const PT_PER_IN = 72
    /** Zoom limits (1 = 100%, i.e. 96 px per inch). */
    const ZOOM_MIN = 0.2
    const ZOOM_MAX = 4
    /** Smallest element edge the canvas will drag to, in inches. */
    const MIN_EDGE = 0.12
    /** Palette offered for text / fill edits. */
    const SWATCHES = ['#0F172A', '#334155', '#64748B', '#2563EB', '#0EA5E9', '#10B981', '#F59E0B', '#EF4444', '#FFFFFF']
    /** Cap on undo snapshots (a deck is small; 40 is plenty). */
    const UNDO_LIMIT = 40
    /** How long a "pending" tab navigation stays meaningful (ms). */
    const PENDING_GRACE_MS = 20_000
    /** Workflow stages, in order. Template first: it defines what can be built. */
    const STAGE_ORDER = ['template', 'research', 'preview', 'build']
    /** Slide transition kinds the canvas can preview (mirrors lib/model.js). */
    const TRANSITION_TYPES = ['none', 'fade', 'push', 'wipe', 'split', 'cover', 'pull', 'cut', 'zoom', 'dissolve', 'circle', 'random']
    /** Transitions whose direction matters. */
    const DIRECTED_TRANSITIONS = new Set(['push', 'wipe', 'split', 'cover', 'pull', 'zoom'])
    /** Entrance effects the canvas can preview. */
    const ANIM_EFFECTS = ['none', 'appear', 'fade', 'wipe', 'zoom']
    /** Canvas size presets (inches), matching lib/model.js. */
    const SIZE_PRESETS = [
      { id: '16:9', w: 13.333, h: 7.5, label: 'size.169' },
      { id: '4:3', w: 10, h: 7.5, label: 'size.43' },
      { id: '16:10', w: 12.8, h: 8, label: 'size.1610' },
      { id: 'a4-landscape', w: 11.693, h: 8.268, label: 'size.a4' },
    ]
    /** Fallback theme (same defaults as the host model). */
    const DEFAULT_THEME = {
      colors: { ink: '#0F172A', muted: '#64748B', accent: '#2563EB', surface: '#F1F5F9', background: '#FFFFFF' },
      fonts: { latin: '微软雅黑', ea: '微软雅黑', heading: '微软雅黑' },
    }
    /** Series palette for chart previews. */
    const CHART_PALETTE = ['#2563EB', '#0EA5E9', '#10B981', '#F59E0B', '#EF4444', '#8B5CF6']

    const zh = {
      'tab.title': 'PPTX 画布',
      'guide.title': 'PPTX 画布',
      'guide.description': '实时看 deck 被画出来，点元素改文字/位置/图片',
      'empty.refresh': '刷新画布列表',
      'empty.shellTitle': '编辑区已就绪',
      'empty.shellBody': '画布已经开好，Agent 一开始绘制就会在这里逐页出现。',
      'state.loading': '正在加载画布…',
      'state.error': '加载失败',
      'state.offline': '宿主未响应',
      'status.idle': '等待写入',
      'status.drawing': '正在绘制第 {index} 页 / 共 {total} 页',
      'status.planned': '大纲已就位，共 {total} 页',
      'status.ready': '画布已就绪，等待第一页',
      'status.preparing': '正在准备第 {index} 页…',
      'status.planning': '正在摆放 {total} 页骨架…',
      'status.elements': '正在绘制第 {index} 页 · 元素 {done}/{count}',
      'status.skeleton': '正在准备画布…',
      'status.saved': '已导出 {name}',
      'status.follow': '跟随最新',
      'actions.export': '导出 PPTX',
      'actions.save': '保存',
      'actions.saveHint': '保存到画布目录（Ctrl+S）',
      'actions.exporting': '导出中…',
      'actions.download': '下载',
      'actions.downloadHint': '导出到任意位置（Ctrl+Shift+S）',
      'actions.downloading': '导出中…',
      'actions.savedAt': '已保存 {name}（{size}）',
      'actions.exportedAt': '已导出到 {path}',
      'actions.exportCancelled': '已取消，未导出',
      'actions.handedOver': '已交给浏览器下载：{name}',
      'actions.noChannel': '这个环境里没有可用的保存通道，导出失败',
      'actions.staleHost': '宿主还没重载：文件其实写到了 {path}。重载插件后再试。',
      'actions.refresh': '重新加载',
      'actions.close': '关闭',
      'actions.send': '发给 DSH',
      'actions.sending': '发送中…',
      'actions.copy': '复制反馈文本',
      'actions.copied': '已复制',
      'actions.undo': '撤销',
      'actions.redo': '重做',
      'actions.fit': '适应窗口',
      'actions.actual': '100%',
      'actions.zoomIn': '放大',
      'actions.zoomOut': '缩小',
      'actions.addSlide': '新增一页',
      'actions.removeSlide': '删除这一页',
      'actions.slideLeft': '前移这一页',
      'actions.slideRight': '后移这一页',
      'actions.addText': '插入文本',
      'actions.addImage': '插入图片',
      'actions.duplicate': '复制元素',
      'actions.delete': '删除元素',
      'actions.front': '置顶',
      'actions.back': '置底',
      'actions.upload': '上传图片',
      'select.hint': '点元素选中；拖动可移动，拖角可缩放；双击改文字；空白处拖拽可框选区域',
      'select.element': '已选元素',
      'select.region': '已选区域',
      'select.clear': '取消选择',
      'select.multi': '已选 {n} 个元素',
      'note.placeholder': '哪里不对？例如：标题太长压到副标题了，缩到 12 字以内',
      'note.chips': ['太挤', '字号太小', '文字压到了', '换配色', '删除这个元素'],
      'note.needText': '先写一句意见，或点一个快捷标签',
      'note.needTarget': '先在画布上点一个元素，或框选一块区域',
      'note.sent': '已发给 DSH：第 {slide} 页 {target}',
      'note.fallback': '已填入输入框，按 Enter 发送',
      'note.copy': '无法自动发送，请复制这段话发给 DSH：',
      'edit.title': '标题',
      'edit.text': '文字',
      'edit.size': '字号',
      'edit.color': '颜色',
      'edit.fill': '填充',
      'edit.align': '对齐',
      'edit.box': '位置与大小（英寸）',
      'edit.bold': '加粗',
      'edit.italic': '斜体',
      'edit.applied': '已保存',
      'edit.unsaved': '正在保存…',
      'edit.failed': '保存失败：{message}',
      'edit.noHost': '宿主不支持画布编辑（缺少 /edit 路由），改动只在本地',
      'edit.lastSlide': '这是最后一页，不能删除',
      'edit.deleted': '已删除',
      'image.uploading': '正在插入图片…',
      'image.failed': '图片插入失败：{message}',
      'image.tooBig': '图片太大（上限 20MB）',
      'image.hint': '把图片拖到画布上，或 Ctrl+V 粘贴',
      'card.title': 'PPTX 画布已创建',
      'card.titleRunning': 'PPTX 画布正在准备',
      'card.badgeRunning': '绘制中',
      'card.badgeDone': '已就绪',
      'card.open': '在右侧查看绘制',
      'card.opened': '画布已打开',
      'card.noTab': '右侧栏暂不可用：先在某个会话里点一下再试',
      'card.slides': '{count} 页',
      'card.opening': '正在打开右侧画布，马上就能看到逐页绘制…',
      'card.live': '画布已跟进本次生成',
      'card.following': '画布已切到这一步，进度会同步显示',
      'card.switching': '正在把右侧画布切到这一步的 deck…',
      'card.failed': '这一步失败了：{message}',
      'card.noDeckYet': '这次调用还没有 deck（新画布要等 pptx_studio_open 返回）',
      'task.running': '正在{label}…',
      'task.done': '已完成{label}',
      'task.failed': '{label}失败',
      'task.deck': 'deck：{id}',
      'task.elapsed': '{n}s',
      'task.waitingTitle': '正在准备新的画布…',
      'task.waitingHint': 'Agent 正在{label}；新 deck 一落地，画布就会自动切过去，不会停在上一份 pptx 上。',
      'task.waitingStep': '等这一步把新 deck 建出来…',
      'tool.open': '打开画布',
      'tool.template': '拆解模板',
      'tool.research': '分析与调研',
      'tool.preview': '生成 HTML 效果稿',
      'tool.deck': '更新演示文稿设置',
      'tool.plan': '规划页面骨架',
      'tool.slide': '绘制幻灯片',
      'tool.addElement': '添加元素',
      'tool.patch': '修改画布',
      'tool.inspect': '检查画布',
      'tool.verify': '一致性校验',
      'tool.export': '导出 PPTX',
      'tool.busy': '处理中',
      'deck.picker': '画布',
      'deck.delete': '删除这份画布',
      'deck.confirm': '再点一次删除',
      'deck.deleted': '已删除，已切到下一份画布',
      'deck.empty': '已经没有画布了',
      'deck.failed': '删除失败：{message}',
      'live.on': '实时',
      'live.off': '离线',
      'saved.hint': '文件：{path}',
      'approx': '近似渲染',
      'zone.editor': '编辑区',
      'zone.controls': '功能区',
      'zone.split': '编辑区 / 功能区',
      'zone.slide': '幻灯片编辑',
      'zone.waiting': '等待 Agent 开始绘制',
      'zone.drawing': '逐页绘制中',
      'actions.shortcuts': '快捷键',
      'shortcuts.title': '画布快捷键',
      'shortcuts.hint': '焦点在画布上时生效；输入框里保留系统编辑快捷键。',
      'shortcut.save': '保存（导出 PPTX）',
      'shortcut.saveAs': '另存 / 下载 PPTX',
      'shortcut.undo': '撤销',
      'shortcut.redo': '重做',
      'shortcut.cut': '剪切元素',
      'shortcut.copy': '复制元素',
      'shortcut.paste': '粘贴元素',
      'shortcut.duplicate': '复制一份',
      'shortcut.newSlide': '新增一页',
      'shortcut.delete': '删除元素',
      'shortcut.selectAll': '选中本页全部元素',
      'shortcut.deselect': '取消选择 / 关闭浮层',
      'shortcut.nudge': '微调位置（Shift 步长更大）',
      'shortcut.fit': '适应窗口',
      'shortcut.zoom': '缩放画布',
      'shortcut.help': '快捷键帮助',
      'shortcut.saving': '正在保存…',
      'shortcut.saved': '已保存到 deck.pptx',
      'clipboard.copied': '已复制 {n} 个元素',
      'clipboard.cut': '已剪切 {n} 个元素',
      'clipboard.pasted': '已粘贴 {n} 个元素',
      'zoom.label': '{pct}%',
      'stage.template': '模板',
      'stage.research': '调研分析',
      'stage.preview': 'HTML 效果',
      'stage.build': '幻灯片',
      'stage.stepTemplate': '⓪ 模板',
      'stage.stepResearch': '① 调研',
      'stage.stepPreview': '② HTML',
      'stage.stepSlides': '③ 幻灯片',
      'stage.done': '已完成',
      'stage.pending': '待进行',
      'stage.current': '进行中',
      'stage.nextResearch': '先让 Agent 完成需求分析与调研。',
      'stage.nextPreview': '让 Agent 先出一份 HTML 效果预览。',
      'stage.nextBuild': '把 HTML 设计落成可编辑幻灯片。',
      'research.emptyTitle': '还没有调研分析',
      'research.emptyBody': '让 Agent 先调研并调用 pptx_studio_research 提交分析、大纲与来源；这里会显示出来供确认。',
      'research.audience': '受众',
      'research.goal': '目标',
      'research.summary': '分析摘要',
      'research.tone': '语气',
      'research.constraints': '约束',
      'research.outline': '叙事大纲',
      'research.facts': '关键结论',
      'research.sources': '来源',
      'research.palette': '配色方向',
      'research.fonts': '字体方向',
      'research.ask': '对调研提意见',
      'preview.emptyTitle': '还没有 HTML 效果预览',
      'preview.emptyBody': '让 Agent 调用 pptx_studio_preview 提交一份自包含 HTML 视觉稿；确认效果后再生成 PPTX。',
      'preview.open': '新窗口打开',
      'preview.download': '下载 HTML',
      'preview.reload': '刷新预览',
      'preview.notes': '设计说明',
      'preview.sandboxHint': '沙箱预览：脚本与外链样式不会加载；图片用内联 data URL。',
      'preview.size': 'HTML {kb} KB',
      'preview.ask': '对 HTML 预览提意见',
      'template.emptyTitle': '这个 deck 没有使用模板',
      'template.emptyBody': '空白模式没有模板可分。用 pptx_studio_open 传 template 参数打开模板模式的 deck，Agent 分解后这里会列出每页版式、可填字段与配色。',
      'template.decomposingTitle': '正在分解模板…',
      'template.decomposingBody': 'Agent 正在调用 pptx_studio_template 把 .pptx 模板拆成可复用版式；完成后这里会列出每页角色、字段与配色。',
      'template.unavailableTitle': '模板数据读取失败',
      'template.unavailableBody': '画布拿不到分解结果（{error}）。让 Agent 重新调用 pptx_studio_template 分解一次。',
      'template.reload': '刷新模板',
      'template.engine': '引擎',
      'template.pagesShort': '{n} 页',
      'template.fieldsShort': '{n} 个字段',
      'template.textHidden': '文字已隐去',
      'template.noPages': '模板里没有可用的页',
      'template.design': '模板设计',
      'template.palette': '配色',
      'template.fonts': '字体',
      'template.textColors': '文字颜色',
      'template.mapped': '本 deck 第 {n} 页',
      'template.mappingEmpty': '本 deck 还没有页映射：让 Agent 在 pptx_studio_research 的 outline 里写 templateSlide + fields。',
      'template.mappedFields': '已填字段',
      'template.noFields': '这一页没有可填字段',
      'template.fieldCap': '≤{n} 字',
      'template.fromSlide': '来自模板第 {n} 页',
      'template.role.cover': '封面',
      'template.role.toc': '目录',
      'template.role.section': '章节',
      'template.role.content': '内容',
      'template.role.ending': '结尾',
      'stage.send': '发给 DSH',
      'stage.sending': '发送中…',
      'stage.needText': '先写一句意见',
      'stage.sent': '已发给 DSH',
      'slide.settings': '页面/动效',
      'transition.fade': '淡入',
      'transition.push': '推入',
      'transition.wipe': '擦除',
      'transition.split': '分割',
      'transition.cover': '覆盖',
      'transition.pull': '拉入',
      'transition.cut': '切出',
      'transition.zoom': '缩放',
      'transition.dissolve': '溶解',
      'transition.circle': '圆形',
      'transition.random': '随机',
      'slide.transition': '切换',
      'slide.speed': '速度',
      'slide.direction': '方向',
      'slide.advance': '换页',
      'slide.advanceClick': '点击后',
      'slide.advanceAfter': '自动（毫秒）',
      'slide.notes': '演讲者备注',
      'slide.background': '背景色',
      'slide.hidden': '隐藏此页',
      'slide.layout': '版式',
      'slide.play': '预览切换',
      'slide.noneTransition': '无',
      'speed.slow': '慢',
      'speed.medium': '中',
      'speed.fast': '快',
      'dir.left': '向左',
      'dir.right': '向右',
      'dir.up': '向上',
      'dir.down': '向下',
      'edit.anim': '入场动画',
      'edit.animNone': '无',
      'edit.animAppear': '出现',
      'edit.animFade': '淡入',
      'edit.animWipe': '擦除',
      'edit.animZoom': '缩放',
      'edit.animPlay': '预览',
      'edit.shadow': '阴影',
      'edit.underline': '下划线',
      'theme.title': '主题与尺寸',
      'theme.colors': '主题色',
      'theme.font': '字体',
      'theme.size': '画布尺寸',
      'size.169': '16:9 宽屏',
      'size.43': '4:3 标准',
      'size.1610': '16:10',
      'size.a4': 'A4 横向',
      'actions.addRect': '矩形',
      'actions.addTable': '表格',
      'actions.addChart': '图表',
      'actions.settings': '页面/动效',
      'actions.theme': '主题',
      'chart.approx': '图表预览',
    }
    const en = {
      'tab.title': 'PPTX canvas',
      'guide.title': 'PPTX canvas',
      'guide.description': 'Watch a deck being drawn; retype, move and re-image elements',
      'empty.refresh': 'Refresh decks',
      'empty.shellTitle': 'Editor ready',
      'empty.shellBody': 'The canvas is open; pages will appear here as the agent draws them.',
      'state.loading': 'Loading canvas…',
      'state.error': 'Failed to load',
      'state.offline': 'Host not reachable',
      'status.idle': 'Waiting',
      'status.drawing': 'Drawing slide {index} / {total}',
      'status.planned': 'Outline ready, {total} slides',
      'status.ready': 'Canvas ready — waiting for the first slide',
      'status.preparing': 'Preparing slide {index}…',
      'status.planning': 'Laying out {total} slides…',
      'status.elements': 'Drawing slide {index} · element {done}/{count}',
      'status.skeleton': 'Preparing the canvas…',
      'status.saved': 'Exported {name}',
      'status.follow': 'Follow latest',
      'actions.export': 'Export PPTX',
      'actions.save': 'Save',
      'actions.saveHint': 'Save into the deck folder (Ctrl+S)',
      'actions.exporting': 'Exporting…',
      'actions.download': 'Download',
      'actions.downloadHint': 'Export to any folder (Ctrl+Shift+S)',
      'actions.downloading': 'Exporting…',
      'actions.savedAt': 'Saved {name} ({size})',
      'actions.exportedAt': 'Exported to {path}',
      'actions.exportCancelled': 'Cancelled — nothing exported',
      'actions.handedOver': 'Handed to the browser ({name})',
      'actions.noChannel': 'No save channel is available in this environment',
      'actions.staleHost': 'The Host has not reloaded yet: the file actually went to {path}. Reload the plugin and retry.',
      'actions.refresh': 'Reload',
      'actions.close': 'Close',
      'actions.send': 'Send to DSH',
      'actions.sending': 'Sending…',
      'actions.copy': 'Copy feedback',
      'actions.copied': 'Copied',
      'actions.undo': 'Undo',
      'actions.redo': 'Redo',
      'actions.fit': 'Fit',
      'actions.actual': '100%',
      'actions.zoomIn': 'Zoom in',
      'actions.zoomOut': 'Zoom out',
      'actions.addSlide': 'Add slide',
      'actions.removeSlide': 'Delete slide',
      'actions.slideLeft': 'Move slide left',
      'actions.slideRight': 'Move slide right',
      'actions.addText': 'Insert text',
      'actions.addImage': 'Insert image',
      'actions.duplicate': 'Duplicate element',
      'actions.delete': 'Delete element',
      'actions.front': 'Bring to front',
      'actions.back': 'Send to back',
      'actions.upload': 'Upload image',
      'select.hint': 'Click to select; drag to move; drag a corner to resize; double-click to retype; drag on empty space to marquee',
      'select.element': 'Element',
      'select.region': 'Region',
      'select.clear': 'Clear',
      'select.multi': '{n} elements',
      'note.placeholder': 'What is wrong? e.g. the title is too long and overlaps the subtitle',
      'note.chips': ['Too tight', 'Text too small', 'Overlapping', 'Recolor', 'Delete this element'],
      'note.needText': 'Write a note or pick a quick tag first',
      'note.needTarget': 'Pick an element (or drag a box) on the canvas first',
      'note.sent': 'Sent to DSH: slide {slide}, {target}',
      'note.fallback': 'Filled into the composer — press Enter to send',
      'note.copy': 'Automatic send failed; copy this to DSH:',
      'edit.title': 'Title',
      'edit.text': 'Text',
      'edit.size': 'Size',
      'edit.color': 'Color',
      'edit.fill': 'Fill',
      'edit.align': 'Align',
      'edit.box': 'Box (inches)',
      'edit.bold': 'Bold',
      'edit.italic': 'Italic',
      'edit.applied': 'Saved',
      'edit.unsaved': 'Saving…',
      'edit.failed': 'Save failed: {message}',
      'edit.noHost': 'Host does not support canvas editing (no /edit route); changes stay local',
      'edit.lastSlide': 'The last slide cannot be deleted',
      'edit.deleted': 'Deleted',
      'image.uploading': 'Inserting image…',
      'image.failed': 'Image insert failed: {message}',
      'image.tooBig': 'Image too large (20MB max)',
      'image.hint': 'Drop an image on the canvas, or press Ctrl+V',
      'card.title': 'PPTX canvas created',
      'card.titleRunning': 'Preparing the PPTX canvas',
      'card.badgeRunning': 'drawing',
      'card.badgeDone': 'ready',
      'card.open': 'Open the live canvas',
      'card.opened': 'Canvas opened',
      'card.noTab': 'Right sidebar unavailable: click inside a session first',
      'card.slides': '{count} slides',
      'card.opening': 'Opening the right-hand canvas — slides will draw in live…',
      'card.live': 'Canvas is following this build',
      'card.following': 'The canvas switched to this step; progress shows live',
      'card.switching': 'Switching the right-hand canvas to this step’s deck…',
      'card.failed': 'This step failed: {message}',
      'card.noDeckYet': 'No deck yet (a new canvas appears once pptx_studio_open returns)',
      'task.running': '{label}…',
      'task.done': '{label} done',
      'task.failed': '{label} failed',
      'task.deck': 'deck: {id}',
      'task.elapsed': '{n}s',
      'task.waitingTitle': 'Preparing the new canvas…',
      'task.waitingHint': 'The agent is {label}; the canvas switches to the new deck as soon as it exists, instead of staying on the previous .pptx.',
      'task.waitingStep': 'Waiting for this step to create the new deck…',
      'tool.open': 'opening the canvas',
      'tool.template': 'decomposing the template',
      'tool.research': 'researching the brief',
      'tool.preview': 'writing the HTML mockup',
      'tool.deck': 'updating the deck settings',
      'tool.plan': 'planning the slides',
      'tool.slide': 'drawing slides',
      'tool.addElement': 'adding elements',
      'tool.patch': 'editing the canvas',
      'tool.inspect': 'inspecting the canvas',
      'tool.verify': 'verifying consistency',
      'tool.export': 'exporting the .pptx',
      'tool.busy': 'working',
      'deck.picker': 'Deck',
      'deck.delete': 'Delete this deck',
      'deck.confirm': 'Click again to delete',
      'deck.deleted': 'Deleted; switched to the next deck',
      'deck.empty': 'No deck left',
      'deck.failed': 'Delete failed: {message}',
      'live.on': 'live',
      'live.off': 'offline',
      'saved.hint': 'File: {path}',
      'approx': 'approx',
      'zone.editor': 'Editor',
      'zone.controls': 'Controls',
      'zone.split': 'Editor / Controls',
      'zone.slide': 'Slide editor',
      'zone.waiting': 'Waiting for the agent',
      'zone.drawing': 'Drawing',
      'actions.shortcuts': 'Shortcuts',
      'shortcuts.title': 'Canvas shortcuts',
      'shortcuts.hint': 'Active while the canvas has focus; inputs keep native editing keys.',
      'shortcut.save': 'Save (export PPTX)',
      'shortcut.saveAs': 'Save as / download PPTX',
      'shortcut.undo': 'Undo',
      'shortcut.redo': 'Redo',
      'shortcut.cut': 'Cut elements',
      'shortcut.copy': 'Copy elements',
      'shortcut.paste': 'Paste elements',
      'shortcut.duplicate': 'Duplicate',
      'shortcut.newSlide': 'New slide',
      'shortcut.delete': 'Delete elements',
      'shortcut.selectAll': 'Select all on this slide',
      'shortcut.deselect': 'Deselect / close popovers',
      'shortcut.nudge': 'Nudge (Shift = larger step)',
      'shortcut.fit': 'Fit to window',
      'shortcut.zoom': 'Zoom canvas',
      'shortcut.help': 'Shortcut help',
      'shortcut.saving': 'Saving…',
      'shortcut.saved': 'Saved to deck.pptx',
      'clipboard.copied': 'Copied {n} element(s)',
      'clipboard.cut': 'Cut {n} element(s)',
      'clipboard.pasted': 'Pasted {n} element(s)',
      'zoom.label': '{pct}%',
      'stage.template': 'Template',
      'stage.research': 'Research',
      'stage.preview': 'HTML preview',
      'stage.build': 'Slides',
      'stage.stepTemplate': '⓪ Template',
      'stage.stepResearch': '① Research',
      'stage.stepPreview': '② HTML',
      'stage.stepSlides': '③ Slides',
      'stage.done': 'done',
      'stage.pending': 'pending',
      'stage.current': 'current',
      'stage.nextResearch': 'Ask the agent to finish analysis and research first.',
      'stage.nextPreview': 'Ask the agent for an HTML visual preview first.',
      'stage.nextBuild': 'Turn the HTML design into editable slides.',
      'research.emptyTitle': 'No research brief yet',
      'research.emptyBody': 'Ask the agent to research and call pptx_studio_research with the analysis, outline and sources; it will show up here for review.',
      'research.audience': 'Audience',
      'research.goal': 'Goal',
      'research.summary': 'Summary',
      'research.tone': 'Tone',
      'research.constraints': 'Constraints',
      'research.outline': 'Narrative outline',
      'research.facts': 'Key findings',
      'research.sources': 'Sources',
      'research.palette': 'Palette',
      'research.fonts': 'Type',
      'research.ask': 'Comment on the research',
      'preview.emptyTitle': 'No HTML preview yet',
      'preview.emptyBody': 'Ask the agent to call pptx_studio_preview with a self-contained HTML mockup; review it before the PPTX is built.',
      'preview.open': 'Open in new window',
      'preview.download': 'Download HTML',
      'preview.reload': 'Reload preview',
      'preview.notes': 'Design notes',
      'preview.sandboxHint': 'Sandboxed preview: scripts and external styles do not load; inline data URL images only.',
      'preview.size': 'HTML {kb} KB',
      'preview.ask': 'Comment on the HTML preview',
      'template.emptyTitle': 'This deck has no template',
      'template.emptyBody': 'Blank-mode decks have no template to decompose. Open a template-mode deck with pptx_studio_open(template); after the agent decomposes it, every layout, field and the palette show up here.',
      'template.decomposingTitle': 'Decomposing the template…',
      'template.decomposingBody': 'The agent is calling pptx_studio_template to turn the .pptx into reusable layouts; page roles, fields and the palette will appear here when it is done.',
      'template.unavailableTitle': 'Template data unavailable',
      'template.unavailableBody': 'The canvas could not load the decomposition ({error}). Ask the agent to run pptx_studio_template again.',
      'template.reload': 'Reload template',
      'template.engine': 'Engine',
      'template.pagesShort': '{n} pages',
      'template.fieldsShort': '{n} fields',
      'template.textHidden': 'text hidden',
      'template.noPages': 'No usable template pages',
      'template.design': 'Template design',
      'template.palette': 'Palette',
      'template.fonts': 'Fonts',
      'template.textColors': 'Text colours',
      'template.mapped': 'Deck page {n}',
      'template.mappingEmpty': 'No page mapping yet: ask the agent to add templateSlide + fields to each pptx_studio_research outline item.',
      'template.mappedFields': 'Filled fields',
      'template.noFields': 'No fillable fields on this page',
      'template.fieldCap': '≤{n} chars',
      'template.fromSlide': 'From template page {n}',
      'template.role.cover': 'Cover',
      'template.role.toc': 'Contents',
      'template.role.section': 'Section',
      'template.role.content': 'Content',
      'template.role.ending': 'Ending',
      'stage.send': 'Send to DSH',
      'stage.sending': 'Sending…',
      'stage.needText': 'Write a note first',
      'stage.sent': 'Sent to DSH',
      'slide.settings': 'Slide / effects',
      'transition.fade': 'Fade',
      'transition.push': 'Push',
      'transition.wipe': 'Wipe',
      'transition.split': 'Split',
      'transition.cover': 'Cover',
      'transition.pull': 'Pull',
      'transition.cut': 'Cut',
      'transition.zoom': 'Zoom',
      'transition.dissolve': 'Dissolve',
      'transition.circle': 'Circle',
      'transition.random': 'Random',
      'slide.transition': 'Transition',
      'slide.speed': 'Speed',
      'slide.direction': 'Direction',
      'slide.advance': 'Advance',
      'slide.advanceClick': 'On click',
      'slide.advanceAfter': 'After (ms)',
      'slide.notes': 'Speaker notes',
      'slide.background': 'Background',
      'slide.hidden': 'Hide slide',
      'slide.layout': 'Layout',
      'slide.play': 'Preview transition',
      'slide.noneTransition': 'None',
      'speed.slow': 'slow',
      'speed.medium': 'medium',
      'speed.fast': 'fast',
      'dir.left': 'left',
      'dir.right': 'right',
      'dir.up': 'up',
      'dir.down': 'down',
      'edit.anim': 'Entrance animation',
      'edit.animNone': 'None',
      'edit.animAppear': 'Appear',
      'edit.animFade': 'Fade',
      'edit.animWipe': 'Wipe',
      'edit.animZoom': 'Zoom',
      'edit.animPlay': 'Preview',
      'edit.shadow': 'Shadow',
      'edit.underline': 'Underline',
      'theme.title': 'Theme & size',
      'theme.colors': 'Theme colours',
      'theme.font': 'Font',
      'theme.size': 'Canvas size',
      'size.169': '16:9 widescreen',
      'size.43': '4:3 standard',
      'size.1610': '16:10',
      'size.a4': 'A4 landscape',
      'actions.addRect': 'Rectangle',
      'actions.addTable': 'Table',
      'actions.addChart': 'Chart',
      'actions.settings': 'Slide / effects',
      'actions.theme': 'Theme',
      'chart.approx': 'Chart preview',
    }

    /**
     * Stylesheet.
     *
     * Rule of thumb: chrome = `--dsw-alias-*` tokens only (so the canvas follows
     * the DSH theme, light or dark); the *slide paper* stays white with its own
     * document colours, because a .pptx is a white page in both themes.
     */
    const CSS = `
/* Token aliases live on both surfaces: the canvas tab and the tool cards that
   the conversation renders (a card is not inside .ps-root, so without this its
   colours and borders would fall back to nothing). */
.ps-root,.ps-scope{
  --ps-bg:var(--dsw-alias-bg-layer-1,#fff);
  --ps-bg-soft:var(--dsw-alias-bg-layer-2,#f5f6f8);
  --ps-border:var(--dsw-alias-border-l1,rgba(0,0,0,.08));
  --ps-border-strong:var(--dsw-alias-border-l2,rgba(0,0,0,.16));
  --ps-fg:var(--dsw-alias-label-primary,#1b1b1b);
  --ps-fg-muted:var(--dsw-alias-label-secondary,#6b7280);
  --ps-accent:var(--dsw-alias-brand-primary,#1b1b1b);
  --ps-on-accent:var(--dsw-alias-label-primary-foreground,#fff);
  --ps-primary:var(--dsw-alias-button-primary-fill,#1b1b1b);
  --ps-primary-hover:var(--dsw-alias-button-primary-hover,#333);
  --ps-elevated:var(--dsw-alias-button-elevated-fill,#fff);
  --ps-hover:var(--dsw-alias-interactive-bg-hover,rgba(0,0,0,.06));
  --ps-danger:var(--dsw-alias-state-error-primary,#b91c1c);
  --ps-success:var(--dsw-alias-state-success-primary,#047857);
  --ps-idle:var(--dsw-alias-state-idle-primary,#94a3b8);
  --ps-font:var(--dsh-font-family-text-default,-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif);
  color:var(--ps-fg);font:12px/1.45 var(--ps-font)}
.ps-root{
  display:flex;flex-direction:column;height:100%;min-height:0;gap:8px;padding:8px 10px 10px;box-sizing:border-box;
  outline:none}

.ps-editor{display:flex;flex-direction:column;flex:1 1 auto;min-height:180px;border:1px solid var(--ps-border-strong);
  border-radius:10px;background:var(--ps-bg-soft);overflow:hidden;box-shadow:0 1px 2px rgba(15,23,42,.05)}
.ps-editor-head{display:flex;align-items:center;gap:6px;min-height:26px;padding:0 8px;flex:0 0 auto;
  border-bottom:1px solid var(--ps-border);background:var(--ps-bg)}
.ps-zone-label{font-size:10px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:var(--ps-fg-muted)}
.ps-zone-sub{font-size:10px;color:var(--ps-fg-muted);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ps-editor .ps-thumbs{padding:6px 8px 4px;border-bottom:1px solid var(--ps-border);background:var(--ps-bg)}
.ps-editor .ps-stage{border-radius:0;background:transparent}
.ps-split{display:flex;align-items:center;justify-content:center;height:10px;flex:0 0 auto;position:relative}
.ps-split::before{content:'';position:absolute;left:4px;right:4px;top:50%;height:1px;background:var(--ps-border-strong)}
.ps-split-grip{position:relative;z-index:1;padding:0 8px;line-height:14px;border:1px solid var(--ps-border);border-radius:999px;
  background:var(--ps-bg);color:var(--ps-fg-muted);font-size:9px;letter-spacing:.04em}
.ps-controls{display:flex;flex-direction:column;flex:0 1 auto;min-height:84px;max-height:48%;border:1px solid var(--ps-border-strong);
  border-radius:10px;background:var(--ps-bg);overflow:hidden}
.ps-controls-head{display:flex;align-items:center;gap:6px;min-height:26px;padding:0 8px;flex:0 0 auto;
  border-bottom:1px solid var(--ps-border);background:var(--ps-bg-soft)}
.ps-controls-body{display:flex;flex-direction:column;gap:6px;flex:1 1 auto;min-height:0;overflow:auto;padding:8px;
  background:var(--ps-bg)}
.ps-shortcuts{display:flex;flex-direction:column;gap:5px;border:1px solid var(--ps-border);border-radius:8px;
  padding:8px;background:var(--ps-bg-soft);margin-bottom:2px}
.ps-shortcuts .ps-shortcut{display:flex;align-items:center;gap:8px;min-width:0}
.ps-shortcuts .ps-shortcut span{font-size:11px;color:var(--ps-fg-muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.ps-shortcuts kbd{flex:0 0 auto;min-width:88px;box-sizing:border-box;padding:1px 6px;border:1px solid var(--ps-border-strong);
  border-bottom-width:2px;border-radius:4px;background:var(--ps-bg);color:var(--ps-fg);font:10px/1.5 var(--ps-font);text-align:center}
.ps-frame-multi{border-style:dashed;box-shadow:none;opacity:.9}
.ps-foot.ps-controls-body{flex:1 1 auto;min-height:0;overflow:auto}

/* ── buttons: three flavours, all token-driven ─────────────────────────── */
.ps-btn{appearance:none;display:inline-flex;align-items:center;justify-content:center;gap:4px;
  border:1px solid var(--ps-border);background:var(--ps-elevated);color:var(--ps-fg);
  border-radius:6px;padding:0 8px;height:24px;font:inherit;font-size:11px;cursor:pointer;white-space:nowrap;
  transition:background .12s,border-color .12s,color .12s}
.ps-btn:hover:not(:disabled){background:var(--ps-hover)}
.ps-btn:active:not(:disabled){transform:translateY(.5px)}
.ps-btn:focus-visible{outline:2px solid color-mix(in srgb,var(--ps-accent) 55%,transparent);outline-offset:1px}
.ps-btn:disabled{opacity:.45;cursor:default}
.ps-btn.primary{background:var(--ps-primary);color:var(--ps-on-accent);border-color:transparent;font-weight:500}
.ps-btn.primary:hover:not(:disabled){background:var(--ps-primary-hover);color:var(--ps-on-accent)}
.ps-btn.ghost{background:transparent;border-color:transparent;color:var(--ps-fg-muted)}
.ps-btn.ghost:hover:not(:disabled){background:var(--ps-hover);color:var(--ps-fg)}
.ps-btn.icon{width:24px;padding:0;font-size:12px;line-height:1}
.ps-select{padding-right:4px}
.ps-select option{background:var(--ps-elevated);color:var(--ps-fg)}
.ps-btn[data-active="true"]{background:var(--ps-hover);color:var(--ps-fg);border-color:var(--ps-border-strong)}
.ps-btn.danger{color:var(--ps-danger)}
.ps-seg{display:inline-flex;align-items:center;border:1px solid var(--ps-border);border-radius:6px;overflow:hidden;background:var(--ps-elevated)}
.ps-seg .ps-btn{border:0;border-radius:0;height:22px}
.ps-seg .ps-btn+.ps-btn{border-left:1px solid var(--ps-border)}

/* ── top bar / thumbnails ──────────────────────────────────────────────── */
.ps-bar{display:flex;align-items:center;gap:6px;min-height:24px}
.ps-bar .ps-title{font-weight:600;font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;flex:1;min-width:0}
.ps-title-input{flex:1;min-width:0;height:22px;border:1px solid transparent;background:transparent;color:var(--ps-fg);
  font:inherit;font-weight:600;font-size:12px;border-radius:5px;padding:0 4px}
.ps-title-input:hover{border-color:var(--ps-border)}
.ps-title-input:focus{outline:none;border-color:var(--ps-accent);background:var(--ps-bg)}
.ps-meta{color:var(--ps-fg-muted);font-size:11px;white-space:nowrap}
.ps-dot{width:7px;height:7px;border-radius:50%;background:var(--ps-idle);flex:0 0 auto}
.ps-dot.live{background:var(--ps-success)}
.ps-thumbs{display:flex;align-items:center;gap:6px;overflow-x:auto;overflow-y:hidden;padding:2px 0 4px;flex:0 0 auto;min-height:60px}
.ps-thumb{position:relative;flex:0 0 auto;width:88px;aspect-ratio:16/9;border-radius:5px;
  border:1px solid var(--ps-border);background:#fff;cursor:pointer;overflow:hidden;padding:0;
  transition:border-color .15s,box-shadow .15s}
.ps-thumb[data-active="true"]{border-color:var(--ps-accent);box-shadow:0 0 0 2px color-mix(in srgb,var(--ps-accent) 24%,transparent)}
.ps-thumb:focus-visible{outline:2px solid color-mix(in srgb,var(--ps-accent) 55%,transparent);outline-offset:1px}
.ps-thumb .ps-thumb-inner{position:absolute;left:0;top:0;transform-origin:top left;pointer-events:none}
.ps-thumb .n{position:absolute;left:3px;bottom:2px;font-size:9px;color:#475569;background:rgba(255,255,255,.88);
  border-radius:3px;padding:0 3px;line-height:1.4}
.ps-thumb .up{position:absolute;top:2px;right:2px;width:6px;height:6px;border-radius:50%;background:var(--ps-success)}

/* ── stage: the slide surface ──────────────────────────────────────────── */
.ps-stage{position:relative;flex:1 1 auto;min-height:0;display:flex;align-items:flex-start;justify-content:center;
  overflow:auto;background:var(--ps-bg-soft);border-radius:8px;padding:10px}
.ps-stage-inner{display:flex;flex-direction:column;align-items:center;gap:8px;min-height:0;max-width:100%}
.ps-canvas{position:relative;background:#fff;box-shadow:0 1px 3px rgba(15,23,42,.18);flex:0 0 auto;
  cursor:default;user-select:none;touch-action:none}
.ps-canvas[data-marquee="true"]{cursor:crosshair}
.ps-slide{position:absolute;inset:0;animation:ps-slide-in .3s ease both}
.ps-el{position:absolute;box-sizing:border-box;overflow:hidden}
.ps-el[data-fresh="true"]{animation:ps-el-in .36s cubic-bezier(.2,.7,.3,1) both}
.ps-el[data-flash="true"]{animation:ps-flash .9s ease both}
/* ── live drawing: pencil ghost → ink-in, plus the drawing HUD ──────────── */
.ps-el[data-sketch="true"]{opacity:.45;filter:saturate(.6);outline:1px dashed color-mix(in srgb,var(--ps-accent) 60%,transparent);
  outline-offset:0;transition:opacity .16s ease,filter .16s ease;position:absolute}
.ps-el[data-sketch="true"]::after{content:'';position:absolute;inset:0;pointer-events:none;
  background:linear-gradient(100deg,transparent 25%,rgba(255,255,255,.55) 50%,transparent 75%);
  background-size:220% 100%;animation:ps-shimmer 1.2s linear infinite}
.ps-el[data-type="text"][data-fresh="true"]{animation:ps-el-ink .4s cubic-bezier(.2,.7,.3,1) both}
.ps-el[data-type="image"][data-fresh="true"],.ps-el[data-type="chart"][data-fresh="true"],
.ps-el[data-type="table"][data-fresh="true"]{animation:ps-el-ink .42s ease both}
.ps-draw-badge{display:inline-flex;align-items:center;gap:5px;height:18px;padding:0 7px;border-radius:999px;
  border:1px solid var(--ps-border);background:var(--ps-bg-soft);color:var(--ps-fg-muted);
  font-size:10px;letter-spacing:.04em;white-space:nowrap}
.ps-draw-dot{width:6px;height:6px;border-radius:50%;background:var(--ps-accent);animation:ps-pulse 1s ease-in-out infinite}
.ps-draw{position:relative;z-index:30;min-width:180px;max-width:100%;
  display:flex;flex-direction:column;gap:5px;padding:7px 9px;border:1px solid var(--ps-border-strong);
  border-radius:8px;background:color-mix(in srgb,var(--ps-bg) 94%,transparent);
  box-shadow:0 4px 14px rgba(15,23,42,.14);pointer-events:none;animation:ps-hud-in .2s ease both}
.ps-draw-top{display:flex;align-items:center;gap:6px;min-width:0;color:var(--ps-fg);font-size:11px}
.ps-draw-pen{color:var(--ps-accent);font-size:11px}
.ps-draw-text{flex:1 1 auto;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.ps-draw-count{color:var(--ps-fg-muted);font-variant-numeric:tabular-nums}
.ps-draw-bar{height:3px;border-radius:999px;background:var(--ps-border);overflow:hidden}
.ps-draw-bar span{display:block;height:100%;border-radius:inherit;background:var(--ps-accent);
  transition:width .22s cubic-bezier(.2,.7,.3,1)}
.ps-stage[data-drawing="true"]{outline:1px solid color-mix(in srgb,var(--ps-accent) 22%,transparent);outline-offset:-1px}
/* ── skeleton: the editing surface exists before any slide does ─────────── */
.ps-skeleton{display:flex;flex-direction:column;align-items:center;gap:8px;width:100%;padding:14px;box-sizing:border-box}
.ps-skeleton-paper{position:relative;width:min(560px,88%);aspect-ratio:16/9;background:#fff;border-radius:6px;
  border:1px dashed color-mix(in srgb,var(--ps-accent) 38%,var(--ps-border-strong));
  box-shadow:0 1px 3px rgba(15,23,42,.14);display:flex;flex-direction:column;padding:7%;box-sizing:border-box;overflow:hidden}
.ps-skeleton-paper::after{content:'';position:absolute;inset:0;
  background:linear-gradient(100deg,transparent 20%,rgba(148,163,184,.16) 50%,transparent 80%);
  background-size:220% 100%;animation:ps-shimmer 1.6s linear infinite;pointer-events:none}
.sk-line,.sk-block{display:block;border-radius:4px;background:rgba(148,163,184,.2)}
.ps-skeleton-hint{color:var(--ps-fg-muted);font-size:11px;text-align:center;max-width:300px;line-height:1.5}
.ps-thumb.ghost{display:block;border-style:dashed;border-color:color-mix(in srgb,var(--ps-border-strong) 70%,transparent);
  background:linear-gradient(100deg,rgba(148,163,184,.10),rgba(148,163,184,.24),rgba(148,163,184,.10));
  background-size:220% 100%;animation:ps-shimmer 1.6s linear infinite;cursor:default}
.ps-thumbs[data-empty="true"]{opacity:.8}
.ps-el[data-type="text"]{white-space:pre-wrap;word-break:break-word}
.ps-el[data-type="rect"]{border-style:none}
.ps-el[data-type="image"] img{width:100%;height:100%;object-fit:fill;display:block;pointer-events:none}
.ps-el[data-type="image"] .ps-img-fallback{width:100%;height:100%;display:flex;align-items:center;justify-content:center;
  background:#f1f5f9;color:#94a3b8;font-size:11px;border:1px dashed #cbd5e1}
.ps-el[data-type="line"]{background:#e2e8f0}
.ps-el[data-type="table"] table{width:100%;height:100%;border-collapse:collapse;table-layout:fixed}
.ps-el[data-type="table"] td{border:1px solid #e2e8f0;padding:2px 4px;overflow:hidden;color:#0f172a}
.ps-el[data-type="group"]{background:#f8fafc;border:1px dashed #cbd5e1;color:#64748b;
  display:flex;align-items:center;justify-content:center;font-size:11px;text-align:center}
/* A chart keeps the slide's own background (theming happens in renderChart);
   only a dashed frame hints at the chart area, like the export's noFill. */
.ps-el[data-type="chart"]{border:1px dashed #cbd5e1;color:#64748b;
  display:flex;align-items:center;justify-content:center;font-size:11px;text-align:center}
.ps-hit{position:absolute;inset:0;cursor:move}
.ps-el[data-selected="true"]{outline:1.5px solid var(--ps-accent);outline-offset:0}

/* selection frame + handles */
.ps-frame{position:absolute;pointer-events:none;border:1.5px solid var(--ps-accent);border-radius:2px;
  box-shadow:0 0 0 1px rgba(255,255,255,.6) inset;z-index:5}
.ps-handle{position:absolute;width:9px;height:9px;margin:-5px 0 0 -5px;border-radius:2px;background:var(--ps-bg);
  border:1.5px solid var(--ps-accent);pointer-events:auto;z-index:6}
.ps-handle[data-h="nw"]{left:0;top:0;cursor:nwse-resize}
.ps-handle[data-h="n"]{left:50%;top:0;cursor:ns-resize}
.ps-handle[data-h="ne"]{left:100%;top:0;cursor:nesw-resize}
.ps-handle[data-h="e"]{left:100%;top:50%;cursor:ew-resize}
.ps-handle[data-h="se"]{left:100%;top:100%;cursor:nwse-resize}
.ps-handle[data-h="s"]{left:50%;top:100%;cursor:ns-resize}
.ps-handle[data-h="sw"]{left:0;top:100%;cursor:nesw-resize}
.ps-handle[data-h="w"]{left:0;top:50%;cursor:ew-resize}
.ps-region{position:absolute;border:2px dashed var(--ps-accent);background:color-mix(in srgb,var(--ps-accent) 12%,transparent);
  pointer-events:none;z-index:5}
.ps-inline{position:absolute;z-index:20;box-sizing:border-box;margin:0;padding:0;border:1.5px solid var(--ps-accent);
  background:#fff;color:#0f172a;resize:none;overflow:hidden;outline:none;box-shadow:0 4px 14px rgba(15,23,42,.18)}

/* ── insp8ector / footer ───────────────────────────────────────────────── */
.ps-foot{display:flex;flex-direction:column;gap:6px;flex:0 0 auto}
.ps-row{display:flex;align-items:center;gap:6px;flex-wrap:wrap;min-width:0}
.ps-status{font-size:11px;color:var(--ps-fg-muted);flex:1 1 auto;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.ps-insp{display:flex;flex-direction:column;gap:6px;border:1px solid var(--ps-border);border-radius:8px;
  padding:8px;background:var(--ps-bg)}
.ps-insp .ps-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:6px}
.ps-field{display:flex;flex-direction:column;gap:2px;min-width:0}
.ps-field label{font-size:10px;color:var(--ps-fg-muted);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ps-input{height:22px;min-width:0;width:100%;box-sizing:border-box;border:1px solid var(--ps-border);border-radius:5px;
  background:var(--ps-bg);color:var(--ps-fg);font:inherit;font-size:11px;padding:0 5px}
.ps-input:focus{outline:none;border-color:var(--ps-accent)}
.ps-textarea{width:100%;box-sizing:border-box;resize:vertical;min-height:44px;max-height:120px;border:1px solid var(--ps-border);
  border-radius:6px;padding:5px 7px;font:inherit;font-size:12px;background:var(--ps-bg);color:var(--ps-fg)}
.ps-textarea:focus{outline:none;border-color:var(--ps-accent)}
.ps-swatches{display:flex;gap:4px;flex-wrap:wrap}
.ps-swatch{width:16px;height:16px;border-radius:4px;border:1px solid var(--ps-border-strong);cursor:pointer;padding:0}
.ps-swatch[data-active="true"]{box-shadow:0 0 0 2px var(--ps-accent)}
.ps-note{width:100%;box-sizing:border-box;resize:vertical;min-height:48px;max-height:140px;border:1px solid var(--ps-border);
  border-radius:6px;padding:6px 8px;font:inherit;font-size:12px;background:var(--ps-bg);color:var(--ps-fg)}
.ps-note:focus{outline:none;border-color:var(--ps-accent)}
.ps-chips{display:flex;gap:5px;flex-wrap:wrap}
.ps-chip{border:1px solid var(--ps-border);background:var(--ps-elevated);border-radius:999px;
  padding:2px 9px;font-size:11px;cursor:pointer;color:var(--ps-fg-muted);font:inherit;font-size:11px}
.ps-chip:hover{border-color:var(--ps-accent);color:var(--ps-fg)}
.ps-selecting{display:flex;align-items:center;gap:8px;font-size:11px;color:var(--ps-fg-muted);
  background:var(--ps-bg-soft);border-radius:6px;padding:4px 8px;min-height:22px;box-sizing:border-box}
.ps-selecting b{color:var(--ps-fg);font-weight:600}
.ps-selecting span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;min-width:0}
.ps-empty{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:10px;height:100%;padding:24px;
  text-align:center;color:var(--ps-fg-muted)}
.ps-empty h4{margin:0;font-size:13px;color:var(--ps-fg)}
.ps-empty p{margin:0;font-size:12px;max-width:280px}
.ps-card{display:flex;flex-direction:column;gap:8px;border:1px solid var(--ps-border);border-radius:8px;
  padding:10px 12px;margin:6px 0;background:var(--ps-bg)}
.ps-card .hd{display:flex;align-items:center;gap:8px;font-size:12px;font-weight:600}
.ps-card .sub{font-size:11px;color:var(--ps-fg-muted)}
.ps-card .badge{font-size:10px;border-radius:999px;padding:1px 7px;background:var(--ps-bg-soft);color:var(--ps-fg-muted)}
/* Session task strip: which step is running and which deck it writes. */
.ps-task{display:flex;align-items:center;gap:6px;min-height:24px;padding:2px 8px;border-radius:6px;
  border:1px solid color-mix(in srgb,var(--ps-accent) 22%,transparent);
  background:color-mix(in srgb,var(--ps-accent) 7%,transparent);color:var(--ps-fg);font-size:11px;
  box-sizing:border-box;flex:0 0 auto}
.ps-task[data-ps-task="failed"]{border-color:color-mix(in srgb,var(--ps-danger) 32%,transparent);
  background:color-mix(in srgb,var(--ps-danger) 8%,transparent)}
.ps-task-dot{width:6px;height:6px;border-radius:50%;background:var(--ps-accent);flex:0 0 auto;
  animation:ps-pulse 1.1s ease-in-out infinite}
.ps-task[data-ps-task="failed"] .ps-task-dot{background:var(--ps-danger);animation:none}
.ps-task-text{font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;min-width:0}
.ps-task-msg{color:var(--ps-fg-muted);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;min-width:0}
.ps-task-elapsed{color:var(--ps-fg-muted);font-variant-numeric:tabular-nums;flex:0 0 auto}
.ps-tab-dot{display:inline-block;width:5px;height:5px;margin-left:5px;border-radius:50%;background:var(--ps-accent);
  vertical-align:middle;animation:ps-pulse 1.1s ease-in-out infinite}
.ps-err{color:var(--ps-danger);font-size:11px}
.ps-ok{color:var(--ps-success);font-size:11px}.ps-drop{position:absolute;inset:6px;border:2px dashed var(--ps-accent);border-radius:8px;pointer-events:none;
  background:color-mix(in srgb,var(--ps-accent) 8%,transparent);z-index:30;display:flex;align-items:center;justify-content:center;
  color:var(--ps-fg);font-size:12px;font-weight:500}
@keyframes ps-slide-in{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}
@keyframes ps-el-in{from{opacity:.2;transform:scale(.99)}to{opacity:1;transform:none}}
@keyframes ps-flash{0%{box-shadow:0 0 0 0 rgba(37,99,235,.5)}70%{box-shadow:0 0 0 7px rgba(37,99,235,0)}100%{box-shadow:0 0 0 0 rgba(37,99,235,0)}}

/* ── research → HTML → slides pipeline ─────────────────────────────────── */
.ps-steps{display:flex;align-items:center;gap:5px;flex:0 0 auto;overflow-x:auto}
.ps-step{display:inline-flex;align-items:center;gap:5px;height:24px;padding:0 8px;border:1px solid var(--ps-border);
  border-radius:6px;background:var(--ps-elevated);color:var(--ps-fg-muted);font:inherit;font-size:11px;cursor:pointer;white-space:nowrap}
.ps-step:hover{background:var(--ps-hover);color:var(--ps-fg)}
.ps-step[data-active="true"]{border-color:var(--ps-accent);color:var(--ps-fg);box-shadow:0 0 0 1px color-mix(in srgb,var(--ps-accent) 30%,transparent)}
.ps-step .n{display:inline-flex;align-items:center;justify-content:center;min-width:15px;height:15px;border-radius:50%;
  background:var(--ps-bg-soft);color:var(--ps-fg-muted);font-size:9px}
.ps-step[data-done="true"] .n{background:var(--ps-success);color:#fff}
.ps-step[data-active="true"] .n{background:var(--ps-accent);color:var(--ps-on-accent)}
.ps-step .sub{font-size:10px;color:var(--ps-fg-muted)}
.ps-steps[data-pending="true"] .ps-step{cursor:default;opacity:.72}
.ps-panel{flex:1 1 auto;min-height:0;overflow:auto;border:1px solid var(--ps-border);border-radius:8px;background:var(--ps-bg);padding:10px;box-sizing:border-box}
.ps-panel-head{display:flex;align-items:center;gap:8px;margin-bottom:8px;flex-wrap:wrap}
.ps-panel-head h4{margin:0;font-size:13px;font-weight:600}
.ps-badge{font-size:10px;border-radius:999px;padding:1px 7px;background:var(--ps-bg-soft);color:var(--ps-fg-muted);border:1px solid var(--ps-border)}
.ps-badge.ok{color:var(--ps-success);border-color:color-mix(in srgb,var(--ps-success) 35%,transparent)}
.ps-brief{display:flex;flex-direction:column;gap:8px}
.ps-brief .card{border:1px solid var(--ps-border);border-radius:8px;padding:8px 10px;background:var(--ps-bg-soft)}
.ps-brief h5{margin:0 0 5px;font-size:11px;font-weight:600;color:var(--ps-fg-muted);letter-spacing:.02em}
.ps-brief p{margin:0;font-size:12px;line-height:1.55;white-space:pre-wrap;word-break:break-word}
.ps-kv{display:grid;grid-template-columns:repeat(auto-fit,minmax(110px,1fr));gap:6px}
.ps-kv .k{font-size:10px;color:var(--ps-fg-muted)}
.ps-kv .v{font-size:12px;color:var(--ps-fg);word-break:break-word}
.ps-list{margin:0;padding-left:18px;display:flex;flex-direction:column;gap:3px;font-size:12px}
.ps-list li{line-height:1.45}
.ps-list .intent{color:var(--ps-fg-muted);font-size:11px}
.ps-source{display:flex;flex-direction:column;gap:2px;font-size:12px}
.ps-source a{color:var(--ps-accent);text-decoration:none;word-break:break-all}
.ps-source a:hover{text-decoration:underline}
.ps-palette{display:flex;gap:6px;flex-wrap:wrap;align-items:center}
.ps-palette span{width:22px;height:22px;border-radius:5px;border:1px solid var(--ps-border-strong);display:inline-block}
.ps-hint{font-size:10px;color:var(--ps-fg-muted);line-height:1.45}
.ps-preview-wrap{position:relative;flex:0 0 auto;width:100%;min-height:220px;border:1px solid var(--ps-border);border-radius:8px;overflow:hidden;background:#fff;display:flex;justify-content:center}
.ps-preview-frame{border:0;display:block;background:#fff}
.ps-stage-note{margin-top:8px;display:flex;flex-direction:column;gap:6px}
.ps-stage-note .ps-row{align-items:center}

/* ── template walkthrough (⓪ before research) ──────────────────────────── */
.ps-tpl{display:flex;flex-direction:column;gap:8px;min-width:0}
.ps-tpl-summary{display:flex;align-items:center;gap:6px;flex-wrap:wrap}
.ps-tpl-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:8px;min-width:0}
.ps-tpl-page{display:flex;flex-direction:column;gap:5px;min-width:0;border:1px solid var(--ps-border);border-radius:8px;
  padding:6px;background:var(--ps-bg-soft)}
.ps-tpl-page.mapped{border-color:var(--ps-accent);background:var(--ps-bg);
  box-shadow:0 0 0 1px color-mix(in srgb,var(--ps-accent) 32%,transparent)}
.ps-tpl-thumb{position:relative;aspect-ratio:16/9;border-radius:5px;overflow:hidden;background:var(--ps-bg);
  border:1px solid var(--ps-border)}
.ps-tpl-thumb img{width:100%;height:100%;object-fit:cover;display:block}
.ps-tpl-no{position:absolute;left:3px;top:3px;font-size:9px;line-height:1.4;border-radius:3px;padding:0 3px;
  background:var(--ps-elevated);color:var(--ps-fg);opacity:.92}
.ps-tpl-role{position:absolute;right:3px;bottom:3px;font-size:9px;line-height:1.4;border-radius:999px;padding:0 5px;
  background:var(--ps-primary);color:var(--ps-on-accent)}
.ps-tpl-fields{display:flex;flex-direction:column;gap:2px;min-width:0}
.ps-tpl-field{display:flex;align-items:baseline;gap:5px;font-size:10px;min-width:0;overflow:hidden}
.ps-tpl-field code{flex:0 0 auto;color:var(--ps-accent);font-size:10px}
.ps-tpl-field .sample{flex:1 1 auto;min-width:0;color:var(--ps-fg-muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.ps-tpl-field .cap{flex:0 0 auto;color:var(--ps-fg-muted);opacity:.85}
.ps-tpl-title{font-size:11px;color:var(--ps-fg);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.ps-tpl-meta{display:flex;align-items:center;gap:6px;flex-wrap:wrap;min-height:14px}
.ps-tpl-mapped{font-size:10px;font-weight:600;color:var(--ps-accent);white-space:nowrap}
.ps-tpl-chip{display:inline-flex;align-items:center;height:16px;padding:0 5px;border-radius:999px;font-size:9px;font-weight:600;
  border:1px solid color-mix(in srgb,var(--ps-accent) 42%,transparent);color:var(--ps-accent);white-space:nowrap}
.ps-thumb .ps-tpl-badge{position:absolute;left:50%;top:2px;transform:translateX(-50%);font-size:8px;line-height:1.4;
  border-radius:3px;padding:0 3px;background:var(--ps-primary);color:var(--ps-on-accent);opacity:.9}

/* ── slide settings / theme ─────────────────────────────────────────────── */
.ps-settings{display:flex;flex-direction:column;gap:8px;border:1px solid var(--ps-border);border-radius:8px;padding:8px;background:var(--ps-bg)}
.ps-settings .ps-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:6px}
.ps-settings .ps-grid.two{grid-template-columns:repeat(2,minmax(0,1fr))}
.ps-settings .ps-grid.three{grid-template-columns:repeat(3,minmax(0,1fr))}
.ps-settings .ps-field label{font-size:10px}
.ps-check{display:inline-flex;align-items:center;gap:4px;font-size:11px;color:var(--ps-fg-muted);white-space:nowrap}
.ps-check input{margin:0}
.ps-color{width:30px;height:22px;border:1px solid var(--ps-border);border-radius:5px;background:transparent;padding:0;cursor:pointer}
.ps-notes{width:100%;box-sizing:border-box;resize:vertical;min-height:38px;max-height:110px;border:1px solid var(--ps-border);
  border-radius:6px;padding:5px 7px;font:inherit;font-size:11px;background:var(--ps-bg);color:var(--ps-fg)}
.ps-notes:focus{outline:none;border-color:var(--ps-accent)}
.ps-thumb[data-hidden="true"]{opacity:.5}
.ps-thumb .h{position:absolute;top:2px;left:3px;font-size:9px;color:#475569;background:rgba(255,255,255,.9);
  border-radius:3px;padding:0 3px;line-height:1.4}
.ps-thumb .fx{position:absolute;right:3px;bottom:2px;font-size:8px;color:#64748b;background:rgba(255,255,255,.9);
  border-radius:3px;padding:0 2px;line-height:1.4}

/* ── chart preview ──────────────────────────────────────────────────────── */
.ps-chart{width:100%;height:100%;display:flex;flex-direction:column;gap:4px;padding:6px;box-sizing:border-box;overflow:hidden}
.ps-chart .t{font-size:10px;color:#334155;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ps-chart .cats{display:flex;gap:6px;font-size:8px;color:#64748b;flex:0 0 auto}
.ps-chart .cats span{flex:1;text-align:center;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.ps-chart svg{width:100%;height:100%;display:block}

/* ── transition / animation previews ───────────────────────────────────── */
.ps-slide[data-fx="fade"]{animation:ps-fx-fade .6s ease both}
.ps-slide[data-fx="cut"]{animation:ps-fx-fade .15s steps(1,end) both}
.ps-slide[data-fx="dissolve"]{animation:ps-fx-dissolve .7s ease both}
.ps-slide[data-fx="circle"]{animation:ps-fx-circle .7s ease both}
.ps-slide[data-fx="zoom"]{animation:ps-fx-zoom .6s ease both}
.ps-slide[data-fx="push"][data-fx-dir="left"]{animation:ps-fx-push-left .55s ease both}
.ps-slide[data-fx="push"][data-fx-dir="right"]{animation:ps-fx-push-right .55s ease both}
.ps-slide[data-fx="push"][data-fx-dir="up"]{animation:ps-fx-push-up .55s ease both}
.ps-slide[data-fx="push"][data-fx-dir="down"]{animation:ps-fx-push-down .55s ease both}
.ps-slide[data-fx="cover"][data-fx-dir="left"]{animation:ps-fx-push-left .55s ease both}
.ps-slide[data-fx="cover"][data-fx-dir="right"]{animation:ps-fx-push-right .55s ease both}
.ps-slide[data-fx="cover"][data-fx-dir="up"]{animation:ps-fx-push-up .55s ease both}
.ps-slide[data-fx="cover"][data-fx-dir="down"]{animation:ps-fx-push-down .55s ease both}
.ps-slide[data-fx="pull"][data-fx-dir="left"]{animation:ps-fx-push-left .55s ease both}
.ps-slide[data-fx="pull"][data-fx-dir="right"]{animation:ps-fx-push-right .55s ease both}
.ps-slide[data-fx="pull"][data-fx-dir="up"]{animation:ps-fx-push-up .55s ease both}
.ps-slide[data-fx="pull"][data-fx-dir="down"]{animation:ps-fx-push-down .55s ease both}
.ps-slide[data-fx="wipe"][data-fx-dir="left"]{animation:ps-fx-wipe-left .6s ease both}
.ps-slide[data-fx="wipe"][data-fx-dir="right"]{animation:ps-fx-wipe-right .6s ease both}
.ps-slide[data-fx="wipe"][data-fx-dir="up"]{animation:ps-fx-wipe-up .6s ease both}
.ps-slide[data-fx="wipe"][data-fx-dir="down"]{animation:ps-fx-wipe-down .6s ease both}
.ps-slide[data-fx="split"][data-fx-dir="left"]{animation:ps-fx-split .6s ease both}
.ps-slide[data-fx="split"][data-fx-dir="right"]{animation:ps-fx-split .6s ease both}
.ps-slide[data-fx="split"][data-fx-dir="up"]{animation:ps-fx-split-vert .6s ease both}
.ps-slide[data-fx="split"][data-fx-dir="down"]{animation:ps-fx-split-vert .6s ease both}
.ps-slide[data-fx="random"]{animation:ps-fx-dissolve .7s ease both}
.ps-el[data-anim="appear"]{animation:ps-anim-appear .18s steps(1,end) both}
.ps-el[data-anim="fade"]{animation:ps-anim-fade .6s ease both}
.ps-el[data-anim="zoom"]{animation:ps-anim-zoom .6s ease both}
.ps-el[data-anim="wipe"][data-anim-dir="left"]{animation:ps-anim-wipe-left .6s ease both}
.ps-el[data-anim="wipe"][data-anim-dir="right"]{animation:ps-anim-wipe-right .6s ease both}
.ps-el[data-anim="wipe"][data-anim-dir="up"]{animation:ps-anim-wipe-up .6s ease both}
.ps-el[data-anim="wipe"][data-anim-dir="down"]{animation:ps-anim-wipe-down .6s ease both}
@keyframes ps-fx-fade{from{opacity:0}to{opacity:1}}
@keyframes ps-fx-dissolve{from{opacity:0;filter:blur(6px)}to{opacity:1;filter:none}}
@keyframes ps-fx-circle{from{clip-path:circle(0% at 50% 50%)}to{clip-path:circle(75% at 50% 50%)}}
@keyframes ps-fx-zoom{from{opacity:0;transform:scale(1.12)}to{opacity:1;transform:none}}
@keyframes ps-fx-push-left{from{transform:translateX(100%)}to{transform:none}}
@keyframes ps-fx-push-right{from{transform:translateX(-100%)}to{transform:none}}
@keyframes ps-fx-push-up{from{transform:translateY(100%)}to{transform:none}}
@keyframes ps-fx-push-down{from{transform:translateY(-100%)}to{transform:none}}
@keyframes ps-fx-wipe-left{from{clip-path:inset(0 100% 0 0)}to{clip-path:inset(0 0 0 0)}}
@keyframes ps-fx-wipe-right{from{clip-path:inset(0 0 0 100%)}to{clip-path:inset(0 0 0 0)}}
@keyframes ps-fx-wipe-up{from{clip-path:inset(100% 0 0 0)}to{clip-path:inset(0 0 0 0)}}
@keyframes ps-fx-wipe-down{from{clip-path:inset(0 0 100% 0)}to{clip-path:inset(0 0 0 0)}}
@keyframes ps-fx-split{from{opacity:.2;transform:scaleX(.2)}to{opacity:1;transform:none}}
@keyframes ps-fx-split-vert{from{opacity:.2;transform:scaleY(.2)}to{opacity:1;transform:none}}
@keyframes ps-anim-appear{from{opacity:0}to{opacity:1}}
@keyframes ps-anim-fade{from{opacity:0}to{opacity:1}}
@keyframes ps-anim-zoom{from{opacity:0;transform:scale(.82)}to{opacity:1;transform:none}}
@keyframes ps-anim-wipe-left{from{clip-path:inset(0 100% 0 0)}to{clip-path:inset(0 0 0 0)}}
@keyframes ps-anim-wipe-right{from{clip-path:inset(0 0 0 100%)}to{clip-path:inset(0 0 0 0)}}
@keyframes ps-anim-wipe-up{from{clip-path:inset(100% 0 0 0)}to{clip-path:inset(0 0 0 0)}}
@keyframes ps-anim-wipe-down{from{clip-path:inset(0 0 100% 0)}to{clip-path:inset(0 0 0 0)}}
/* live-drawing keyframes */
@keyframes ps-el-ink{from{clip-path:inset(0 100% 0 0);opacity:.3}to{clip-path:inset(0 0 0 0);opacity:1}}
@keyframes ps-pulse{0%,100%{opacity:.35;transform:scale(.85)}50%{opacity:1;transform:scale(1.2)}}
@keyframes ps-hud-in{from{opacity:0;transform:translateY(4px)}to{opacity:1;transform:none}}
@keyframes ps-shimmer{from{background-position:180% 0}to{background-position:-80% 0}}
@media (prefers-reduced-motion: reduce){
  .ps-el[data-fresh="true"],.ps-el[data-flash="true"],.ps-slide,.ps-draw,.ps-draw-dot,
  .ps-skeleton-paper::after,.ps-thumb.ghost{animation:none!important}
  .ps-draw-bar span,.ps-el[data-sketch="true"]{transition:none!important}
}
`

    // ── helpers ──────────────────────────────────────────────────────────────

    /**
     * Resolve one host route against the page base.
     * @param {string} path - path after the API prefix.
     * @returns {string} absolute URL.
     */
    function api(path) {
      const relative = `${API_PATH}/${String(path).replace(/^\/+/, '')}`
      try {
        const base = new URL(document.baseURI)
        base.search = ''
        base.hash = ''
        if (!base.pathname.endsWith('/')) base.pathname += '/'
        return new URL(relative, base.href).toString()
      } catch {
        return `/${relative}`
      }
    }

    /**
     * JSON fetch with a timeout and a readable failure.
     * @param {string} path - API path.
     * @param {{method?: string, body?: unknown, timeoutMs?: number}} [options] - request options.
     * @returns {Promise<any>} parsed JSON.
     */
    async function request(path, options = {}) {
      const controller = typeof AbortController === 'function' ? new AbortController() : null
      const timer = controller ? setTimeout(() => controller.abort(), options.timeoutMs ?? 15000) : null
      try {
        const response = await fetch(api(path), {
          method: options.method ?? 'GET',
          headers: options.body === undefined ? undefined : { 'content-type': 'application/json' },
          body: options.body === undefined ? undefined : JSON.stringify(options.body),
          signal: controller ? controller.signal : undefined,
          credentials: 'same-origin',
        })
        const text = await response.text()
        let data = {}
        try {
          data = text ? JSON.parse(text) : {}
        } catch {
          data = { ok: false, code: 'bad-json', message: text.slice(0, 200) }
        }
        if (!response.ok) {
          const error = new Error(data?.message || data?.code || `HTTP ${response.status}`)
          // Keep the machine code so callers can special-case "host too old".
          ;/** @type {any} */ (error).code = data?.code
          ;/** @type {any} */ (error).status = response.status
          throw error
        }
        return data
      } finally {
        if (timer) clearTimeout(timer)
      }
    }

    /**
     * Format a translation string with simple placeholders.
     * @param {string} template - template text.
     * @param {Record<string, unknown>} [values] - placeholder values.
     * @returns {string} formatted text.
     */
    function fmt(template, values) {
      return String(template).replace(/\{(\w+)\}/g, (_, key) => (values && values[key] !== undefined ? String(values[key]) : `{${key}}`))
    }

    /**
     * Last path segment of a file path (handles both separators).
     * @param {unknown} path - file path.
     * @returns {string} file name, or an empty string.
     */
    function fileNameOf(path) {
      const parts = String(path ?? '').split(/[\\/]+/).filter(Boolean)
      return parts.length > 0 ? parts[parts.length - 1] : ''
    }

    /**
     * Human-sized byte count for receipts (`12345` → `12.1 KB`).
     * @param {unknown} bytes - byte count.
     * @returns {string} readable size.
     */
    function byteSize(bytes) {
      const value = Number(bytes)
      if (!Number.isFinite(value) || value < 0) return '—'
      if (value < 1024) return `${Math.round(value)} B`
      if (value < 1024 * 1024) return `${Math.round(value / 102.4) / 10} KB`
      return `${Math.round(value / 104857.6) / 10} MB`
    }

    // ── session activity feed ────────────────────────────────────────────────

    /** Short dictionary key of each `pptx_studio_*` tool. */
    const TOOL_SHORT = {
      pptx_studio_open: 'open',
      pptx_studio_template: 'template',
      pptx_studio_research: 'research',
      pptx_studio_preview: 'preview',
      pptx_studio_deck: 'deck',
      pptx_studio_plan: 'plan',
      pptx_studio_slide: 'slide',
      pptx_studio_add_element: 'addElement',
      pptx_studio_patch: 'patch',
      pptx_studio_inspect: 'inspect',
      pptx_studio_verify: 'verify',
      pptx_studio_export: 'export',
    }

    /**
     * Localized label of one tool (the fallback is the wire name).
     * @param {(key: string, values?: Record<string, unknown>) => string} t - translator.
     * @param {string} name - tool name.
     * @returns {string} label.
     */
    function toolLabel(t, name) {
      const short = TOOL_SHORT[String(name || '')]
      return short ? t(`tool.${short}`) : (String(name || '') || t('tool.busy'))
    }

    /**
     * Per-session activity feeds. The host publishes the `pptx_studio_*` call in
     * flight (which deck it writes, which step it is); the canvas tab and the
     * tool cards share one stream per session so a tab showing the previous deck
     * can switch to the deck being drawn, and the cards can report progress.
     *
     * @type {Map<string, {task: Record<string, any>|null, listeners: Set<(task: Record<string, any>) => void>, refs: number, started: boolean, source: any, poll: any}>}
     */
    const taskFeeds = new Map()

    /**
     * The feed record of one session (created on first use).
     * @param {string} sessionId - session id.
     * @returns {any} feed record.
     */
    function taskFeedOf(sessionId) {
      let feed = taskFeeds.get(sessionId)
      if (!feed) {
        feed = { task: null, listeners: new Set(), refs: 0, started: false, source: null, poll: null }
        taskFeeds.set(sessionId, feed)
      }
      return feed
    }

    /**
     * Accept one published task (the host's `seq` only grows).
     * @param {string} sessionId - session id.
     * @param {any} task - task payload.
     * @returns {void}
     */
    function acceptTask(sessionId, task) {
      if (!task || typeof task !== 'object') return
      const feed = taskFeedOf(sessionId)
      const previous = feed.task
      if (previous && Number(task.seq || 0) <= Number(previous.seq || 0)) return
      feed.task = task
      for (const listener of [...feed.listeners]) {
        try {
          listener(task)
        } catch {
          /* one broken subscriber must not stop the others */
        }
      }
    }

    /**
     * Start (or restart) the stream of one session: a snapshot fetch plus SSE,
     * with a slow poll when the page has no EventSource.
     * @param {string} sessionId - session id.
     * @returns {void}
     */
    function startTaskFeed(sessionId) {
      const feed = taskFeedOf(sessionId)
      if (feed.started) return
      feed.started = true
      const refresh = () => {
        request(`session?sessionId=${encodeURIComponent(sessionId)}`, { timeoutMs: 8000 })
          .then((data) => {
            if (data && data.task) acceptTask(sessionId, data.task)
          })
          .catch(() => { /* host may be reloading; SSE retries on its own */ })
      }
      refresh()
      if (typeof EventSource === 'function') {
        try {
          const source = new EventSource(api(`session/events?sessionId=${encodeURIComponent(sessionId)}`))
          source.onmessage = (message) => {
            let event
            try {
              event = JSON.parse(message.data)
            } catch {
              return
            }
            if (event && event.type === 'task') acceptTask(sessionId, event.task)
          }
          feed.source = source
        } catch {
          feed.source = null
        }
      }
      if (!feed.source) {
        feed.poll = setInterval(refresh, 2500)
        if (feed.poll && typeof feed.poll.unref === 'function') feed.poll.unref()
      }
    }

    /**
     * Stop one session's stream once nothing listens any more.
     * @param {string} sessionId - session id.
     * @returns {void}
     */
    function stopTaskFeed(sessionId) {
      const feed = taskFeedOf(sessionId)
      if (feed.source) {
        try {
          feed.source.close()
        } catch {
          /* already closed */
        }
        feed.source = null
      }
      if (feed.poll) {
        clearInterval(feed.poll)
        feed.poll = null
      }
      feed.started = false
    }

    /**
     * Subscribe to one session's task feed.
     * @param {string} sessionId - session id.
     * @param {(task: any) => void} listener - sink.
     * @returns {() => void} unsubscribe.
     */
    function subscribeTask(sessionId, listener) {
      const feed = taskFeedOf(sessionId)
      feed.listeners.add(listener)
      feed.refs += 1
      startTaskFeed(sessionId)
      return () => {
        feed.listeners.delete(listener)
        feed.refs = Math.max(0, feed.refs - 1)
        if (feed.refs === 0) stopTaskFeed(sessionId)
      }
    }

    /**
     * Read the running (or last) pptx step of one session.
     * @param {string|undefined} sessionId - session id.
     * @returns {any} task or null.
     */
    function useSessionTask(sessionId) {
      const [task, setTask] = useState(null)
      useEffect(() => {
        const id = typeof sessionId === 'string' ? sessionId : ''
        if (id === '') {
          setTask(null)
          return undefined
        }
        setTask(taskFeedOf(id).task)
        return subscribeTask(id, setTask)
      }, [sessionId])
      return task
    }

    /**
     * Elapsed seconds of a running task, refreshed every second.
     * @param {any} task - task payload.
     * @returns {number} seconds (0 when the task is not running).
     */
    function useElapsed(task) {
      const running = Boolean(task && task.state === 'running')
      const [now, setNow] = useState(() => Date.now())
      useEffect(() => {
        if (!running) return undefined
        setNow(Date.now())
        const timer = setInterval(() => setNow(Date.now()), 1000)
        return () => clearInterval(timer)
      }, [running, task && task.seq])
      if (!running) return 0
      const started = typeof task.startedAt === 'string' ? Date.parse(task.startedAt) : NaN
      if (!Number.isFinite(started)) return 0
      return Math.max(0, Math.round((now - started) / 1000))
    }


    /**
     * Translation key of one workflow-step chip (template is step ⓪, before
     * research). Kept in one place so the pending shell and the live canvas
     * cannot drift apart.
     * @param {string} key - stage key from `STAGE_ORDER`.
     * @returns {string} dictionary key.
     */
    function stageStepKey(key) {
      if (key === 'template') return 'stage.stepTemplate'
      if (key === 'research') return 'stage.stepResearch'
      if (key === 'preview') return 'stage.stepPreview'
      return 'stage.stepSlides'
    }

    /**
     * Translation key of one workflow-stage title.
     * @param {string} key - stage key from `STAGE_ORDER`.
     * @returns {string} dictionary key.
     */
    function stageTitleKey(key) {
      if (key === 'template') return 'stage.template'
      if (key === 'research') return 'stage.research'
      if (key === 'preview') return 'stage.preview'
      return 'stage.build'
    }

    /**
     * Read an array-valued translation entry safely (a bound `t` returns the
     * raw key when the dictionary has no such entry).
     * @param {(key: string, values?: any) => any} t - translator.
     * @param {string} key - dictionary key.
     * @param {string[]} [fallback] - fallback list.
     * @returns {string[]} list.
     */
    function tList(t, key, fallback = []) {
      let value
      try {
        value = t(key)
      } catch {
        value = undefined
      }
      return Array.isArray(value) ? value : fallback
    }

    /**
     * Deep clone a plain JSON model.
     * @param {any} value - JSON value.
     * @returns {any} clone.
     */
    function clone(value) {
      try {
        return JSON.parse(JSON.stringify(value))
      } catch {
        return value
      }
    }

    /**
     * Clamp a number.
     * @param {number} n - value.
     * @param {number} min - lower bound.
     * @param {number} max - upper bound.
     * @returns {number} clamped value.
     */
    function clamp(n, min, max) {
      return Math.min(max, Math.max(min, n))
    }

    /**
     * Round to 3 decimals (0.001in), matching the host model.
     * @param {number} n - value.
     * @returns {number} rounded value.
     */
    function round3(n) {
      return Math.round(n * 1000) / 1000
    }

    /**
     * CSS px per point at a given zoom. 1 = 100% = 96 CSS px per inch.
     * @param {number} zoom - zoom factor.
     * @returns {number} px per point.
     */
    function pxPerPt(zoom) {
      return (PX_PER_IN / PT_PER_IN) * zoom
    }

    /**
     * Canonicalise a deck colour for inline CSS.
     *
     * Deck models may spell a colour `060B18` (what python-pptx accepts) or
     * `#060B18`; only the second is a valid CSS value, and an invalid one is
     * dropped silently — the element then renders transparent. Never hand a
     * bare hex string to the DOM.
     * @param {unknown} value - candidate colour.
     * @param {string} fallback - used when the value is empty or unusable.
     * @returns {string} CSS colour.
     */
    function cssColor(value, fallback) {
      if (typeof value !== 'string') return fallback
      const raw = value.trim()
      if (raw === '') return fallback
      const hex = /^#?([0-9a-fA-F]{3}|[0-9a-fA-F]{4}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/.exec(raw)
      if (!hex) return raw
      const digits = hex[1]
      return `#${digits.length <= 4 ? [...digits].map((d) => d + d).join('') : digits}`
    }

    /**
     * Rough brightness test for a CSS colour, mirroring pptx_build.py's luma:
     * decides whether a chart drawn on this background needs light or dark
     * ink. Anything unparseable counts as light.
     * @param {unknown} value - candidate colour.
     * @returns {boolean} true when the colour reads as dark.
     */
    function isDarkBg(value) {
      const raw = String(value || '').trim().replace(/^#/, '')
      if (!/^([0-9a-fA-F]{3}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/.test(raw)) return false
      const hex = raw.length <= 4 ? [...raw].map((d) => d + d).join('') : raw
      const r = parseInt(hex.slice(0, 2), 16)
      const g = parseInt(hex.slice(2, 4), 16)
      const b = parseInt(hex.slice(4, 6), 16)
      return 0.2126 * r + 0.7152 * g + 0.0722 * b < 110
    }

    /**
     * Workflow stage of a deck, tolerant of decks written before the pipeline.
     * @param {any} deck - deck model.
     * @returns {'template'|'research'|'preview'|'build'} stage.
     */
    function stageOf(deck) {
      const pipeline = deck && deck.pipeline
      if (pipeline && pipeline.preview) return 'build'
      if (pipeline && pipeline.research) return 'preview'
      if (pipeline && pipeline.builtAt) return 'build'
      // Template mode is its own first step: nothing can be decomposed before
      // the template says which layouts and fields exist. The host exposes the
      // stage directly (`pipeline.stage`); the `template.required` fallback keeps
      // decks written before that field from skipping it.
      if (pipeline && (pipeline.stage === 'template' || (pipeline.template && pipeline.template.required === true && !pipeline.template.decomposedAt))) return 'template'
      return ((deck && deck.slides) || []).length > 0 ? 'build' : 'research'
    }

    /**
     * Theme of a deck with the model defaults merged in.
     * @param {any} deck - deck model.
     * @returns {any} theme.
     */
    function themeOf(deck) {
      const theme = (deck && deck.theme) || {}
      const colors = theme.colors || {}
      const fonts = theme.fonts || {}
      return {
        colors: {
          ink: cssColor(colors.ink, DEFAULT_THEME.colors.ink),
          muted: cssColor(colors.muted, DEFAULT_THEME.colors.muted),
          accent: cssColor(colors.accent, DEFAULT_THEME.colors.accent),
          surface: cssColor(colors.surface, DEFAULT_THEME.colors.surface),
          background: cssColor(colors.background, DEFAULT_THEME.colors.background),
        },
        fonts: {
          latin: fonts.latin || DEFAULT_THEME.fonts.latin,
          ea: fonts.ea || DEFAULT_THEME.fonts.ea,
          heading: fonts.heading || DEFAULT_THEME.fonts.heading,
        },
      }
    }

    /**
     * Work out the design-canvas size of an HTML preview.
     *
     * The mockup is authored at a fixed pixel canvas (often 960×540 or
     * 1280×720). The panel then scales that canvas uniformly; without a size
     * the iframe would either clip the slide or stretch it to a wrong aspect.
     * @param {any} preview - `pipeline.preview`.
     * @param {number} slideW - deck width in inches.
     * @param {number} slideH - deck height in inches.
     * @returns {{width: number, height: number}} design canvas in CSS px.
     */
    function inferPreviewSize(preview, slideW, slideH) {
      const deckWidth = Math.max(320, Math.round(slideW * PX_PER_IN))
      const declaredWidth = Number(preview && preview.width) > 0
      const declaredHeight = Number(preview && preview.height) > 0
      let width = declaredWidth ? Number(preview.width) : 0
      let height = declaredHeight ? Number(preview.height) : 0
      if (width <= 0 && preview && typeof preview.html === 'string') {
        // Respect an explicit `.slide { width: …px }` declaration. Use the
        // most frequent value so a tiny side label cannot win.
        /** @type {Map<number, number>} */
        const counts = new Map()
        const pattern = /(?:^|[^-\w])width\s*:\s*(\d{3,4})px/gi
        let match
        while ((match = pattern.exec(preview.html)) !== null) {
          const value = Number(match[1])
          if (value >= 480 && value <= 3840) counts.set(value, (counts.get(value) || 0) + 1)
        }
        if (counts.size > 0) {
          width = [...counts.entries()].sort((a, b) => b[1] - a[1])[0][0]
        }
      }
      if (width <= 0) width = deckWidth
      if (height <= 0) height = (width * slideH) / slideW
      // A mockup usually wraps its slide in a padded stage; a fixed `.slide`
      // width alone would crop those gutters. Leave room for common 24px
      // gutters unless the caller declared the full document viewport.
      if (!declaredWidth) width += 48
      if (!declaredHeight) height += 48
      return {
        width: Math.max(200, Math.min(3840, Math.round(width))),
        height: Math.max(150, Math.min(2160, Math.round(height))),
      }
    }

    /**
     * Match a deck size to one of the presets.
     * @param {any} size - `{w,h}`.
     * @returns {string} preset id or 'custom'.
     */
    function sizePresetOf(size) {
      const w = Number(size && size.w) || 0
      const h = Number(size && size.h) || 0
      const hit = SIZE_PRESETS.find((item) => Math.abs(item.w - w) < 0.02 && Math.abs(item.h - h) < 0.02)
      return hit ? hit.id : 'custom'
    }

    /**
     * CSS box-shadow for one element shadow model.
     * @param {any} shadow - shadow model.
     * @returns {string} CSS value ('' when absent).
     */
    function shadowCss(shadow) {
      if (!shadow || typeof shadow !== 'object') return ''
      const blur = Number(shadow.blur) > 0 ? Number(shadow.blur) : 4
      const distance = Number(shadow.distance) >= 0 ? Number(shadow.distance) : 3
      const angle = ((Number(shadow.direction) || 90) * Math.PI) / 180
      const dx = Math.round(Math.cos(angle) * distance * 100) / 100
      const dy = Math.round(Math.sin(angle) * distance * 100) / 100
      const color = cssColor(shadow.color, '#000000')
      const alpha = Number.isFinite(Number(shadow.alpha)) ? Math.max(0, Math.min(1, Number(shadow.alpha))) : 0.35
      const rgba = /^#([0-9a-f]{6})$/i.exec(color)
      const fill = rgba
        ? `rgba(${parseInt(rgba[1].slice(0, 2), 16)},${parseInt(rgba[1].slice(2, 4), 16)},${parseInt(rgba[1].slice(4, 6), 16)},${alpha})`
        : color
      return `${dx}px ${dy}px ${blur}px ${fill}`
    }

    /**
     * Fit a slide of `slideW x slideH` inches into a viewport.
     * @param {number} slideW - slide width (in).
     * @param {number} slideH - slide height (in).
     * @param {number} boxW - viewport width (px).
     * @param {number} boxH - viewport height (px).
     * @param {number} pad - padding inside the viewport (px), both sides.
     * @returns {number} zoom factor that fits (clamped to the zoom range).
     */
    function fitZoom(slideW, slideH, boxW, boxH, pad = 16) {
      const usableW = Math.max(40, boxW - pad * 2)
      const usableH = Math.max(40, boxH - pad * 2)
      const byWidth = usableW / (slideW * PX_PER_IN)
      const byHeight = usableH / (slideH * PX_PER_IN)
      return clamp(Math.min(byWidth, byHeight), ZOOM_MIN, ZOOM_MAX)
    }

    /**
     * Absolute URL for an element image.
     *
     * `src` may be a deck-relative asset (`assets/img_ab12cd34.png`), a plain
     * file name, a workspace path, an absolute path, a data URL or an http URL.
     * @param {string} deckId - deck id.
     * @param {string} src - element src.
     * @param {number|string} [rev] - cache buster.
     * @returns {string} URL.
     */
    function assetUrl(deckId, src, rev) {
      const value = String(src || '')
      if (value === '') return ''
      if (/^(data:|https?:|blob:)/i.test(value)) return value
      const params = new URLSearchParams({ deck: deckId, path: value })
      if (rev !== undefined && rev !== null && rev !== '') params.set('v', String(rev))
      return api(`asset?${params.toString()}`)
    }

    /**
     * Build a unique element id inside one slide.
     * @param {any[]} elements - existing elements.
     * @param {string} base - preferred id.
     * @returns {string} unique id.
     */
    function uniqueId(elements, base) {
      const taken = new Set((elements || []).map((el) => el.id))
      if (!taken.has(base)) return base
      for (let i = 2; i < 1000; i += 1) {
        const candidate = `${base}-${i}`
        if (!taken.has(candidate)) return candidate
      }
      return `${base}-${Date.now().toString(36)}`
    }

    /**
     * Locate a slide by id.
     * @param {any} deck - deck model.
     * @param {string} slideId - slide id.
     * @returns {number} index, or -1.
     */
    function slideIndexById(deck, slideId) {
      return ((deck && deck.slides) || []).findIndex((slide) => slide.id === slideId)
    }

    /**
     * Locate an element by id.
     * @param {any} slide - slide model.
     * @param {string} elementId - element id.
     * @returns {number} index, or -1.
     */
    function elementIndexById(slide, elementId) {
      return ((slide && slide.elements) || []).findIndex((el) => el.id === elementId)
    }

    /**
     * Apply one edit op to a deck model locally (optimistic UI).
     *
     * The host applies the same op authoritatively and pushes the result back
     * over SSE; this local pass exists so a drag feels instant.
     * @param {any} deck - deck model (not mutated).
     * @param {Record<string, any>} op - one op.
     * @returns {any} new deck model.
     */
    function applyLocal(deck, op) {
      if (!deck) return deck
      const next = clone(deck)
      const slides = next.slides || (next.slides = [])
      const slide = op.slideId ? slides.find((s) => s.id === op.slideId) : null
      if (op.kind === 'deck.title') {
        next.title = String(op.title ?? '').slice(0, 200)
        return next
      }
      if (op.kind === 'deck.theme') {
        const patch = op.theme && typeof op.theme === 'object' ? op.theme : {}
        const current = next.theme || { colors: {}, fonts: {} }
        next.theme = {
          ...current,
          colors: { ...(current.colors || {}), ...(patch.colors || {}) },
          fonts: { ...(current.fonts || {}), ...(patch.fonts || {}) },
        }
        return next
      }
      if (op.kind === 'deck.size') {
        next.size = { w: round3(Number(op.width) || next.size?.w || 13.333), h: round3(Number(op.height) || next.size?.h || 7.5) }
        return next
      }
      if (op.kind === 'slide.add') {
        const id = op.slide?.id || uniqueId(slides, `s${slides.length + 1}`)
        const created = {
          ...(op.slide || {}),
          id,
          title: op.slide?.title || `第 ${slides.length + 1} 页`,
          elements: op.slide?.elements || [],
        }
        const at = op.index === undefined ? slides.length : clamp(op.index, 0, slides.length)
        slides.splice(at, 0, created)
        return next
      }
      if (op.kind === 'slide.patch' && slide) {
        const patch = op.patch && typeof op.patch === 'object' ? op.patch : {}
        for (const [key, value] of Object.entries(patch)) {
          if (value === null) delete slide[key]
          else if (value && typeof value === 'object' && !Array.isArray(value) && slide[key] && typeof slide[key] === 'object' && !Array.isArray(slide[key])) {
            slide[key] = { ...slide[key], ...value }
          } else slide[key] = value
        }
        return next
      }
      if (op.kind === 'slide.remove' && slide) {
        const at = slides.indexOf(slide)
        if (slides.length > 1) slides.splice(at, 1)
        return next
      }
      if (op.kind === 'slide.move' && slide) {
        const from = slides.indexOf(slide)
        const to = clamp(Number(op.index) || 0, 0, slides.length - 1)
        slides.splice(from, 1)
        slides.splice(to, 0, slide)
        return next
      }
      if (!slide) return next
      const elements = slide.elements || (slide.elements = [])
      if (op.kind === 'element.add' && op.element) {
        const el = clone(op.element)
        el.id = el.id || uniqueId(elements, `${slide.id}-e${elements.length + 1}`)
        // The host echoes every add back over SSE; keep the two paths merged.
        const existing = elements.findIndex((item) => item.id === el.id)
        if (existing >= 0) elements[existing] = el
        else {
          const index = op.index === undefined ? elements.length : clamp(op.index, 0, elements.length)
          elements.splice(index, 0, el)
        }
        return next
      }
      const at = op.elementId ? elements.findIndex((el) => el.id === op.elementId) : -1
      if (op.kind === 'element.patch') {
        if (at < 0) return next
        const patch = op.patch || {}
        if (op.remove === true) {
          elements.splice(at, 1)
          return next
        }
        const el = elements[at]
        for (const [key, value] of Object.entries(patch)) {
          if (value === null) delete el[key]
          else if (value && typeof value === 'object' && !Array.isArray(value) && el[key] && typeof el[key] === 'object') {
            el[key] = { ...el[key], ...value }
          } else el[key] = value
        }
        if (Array.isArray(el.box)) {
          el.box = [round3(el.box[0]), round3(el.box[1]), round3(Math.max(MIN_EDGE, el.box[2])), round3(Math.max(MIN_EDGE, el.box[3]))]
        }
        return next
      }
      if (op.kind === 'element.replace' && op.element) {
        const el = clone(op.element)
        if (at >= 0) elements[at] = el
        else elements.push(el)
        return next
      }
      if (op.kind === 'element.duplicate' && at >= 0) {
        const src = elements[at]
        const copy = clone(src)
        copy.id = uniqueId(elements, `${src.id}-copy`)
        const [dx, dy] = Array.isArray(op.offset) ? op.offset : [0.2, 0.2]
        copy.box = [round3(src.box[0] + dx), round3(src.box[1] + dy), src.box[2], src.box[3]]
        elements.splice(at + 1, 0, copy)
        return next
      }
      if (op.kind === 'element.reorder' && at >= 0) {
        const [el] = elements.splice(at, 1)
        let to = op.to
        if (to === 'front') to = elements.length
        else if (to === 'back') to = 0
        elements.splice(clamp(Number(to) || 0, 0, elements.length), 0, el)
        return next
      }
      if (op.kind === 'slide.replace' && op.slide) {
        const index = slides.indexOf(slide)
        slides[index] = { ...slide, ...op.slide }
        return next
      }
      return next
    }

    /**
     * Box of an element, guarding against a malformed model.
     * @param {any} el - element.
     * @param {number} slideW - slide width (in).
     * @param {number} slideH - slide height (in).
     * @returns {number[]} `[x, y, w, h]`.
     */
    function boxOf(el, slideW, slideH) {
      const box = Array.isArray(el?.box) ? el.box : []
      const x = Number(box[0]) || 0
      const y = Number(box[1]) || 0
      const w = Number(box[2]) > 0 ? Number(box[2]) : slideW / 2
      const hh = Number(box[3]) > 0 ? Number(box[3]) : slideH / 4
      return [x, y, w, hh]
    }

    /**
     * Shift a box and keep it reachable on the slide.
     * @param {number[]} box - source `[x, y, w, h]` in inches.
     * @param {number} dx - x offset.
     * @param {number} dy - y offset.
     * @param {number} slideW - slide width (in).
     * @param {number} slideH - slide height (in).
     * @returns {number[]} shifted box.
     */
    function shiftBox(box, dx, dy, slideW, slideH) {
      const [x, y, w, hh] = Array.isArray(box) ? box : [0, 0, slideW / 2, slideH / 4]
      const width = Number(w) > 0 ? Number(w) : slideW / 2
      const height = Number(hh) > 0 ? Number(hh) : slideH / 4
      return [
        round3(clamp((Number(x) || 0) + dx, -width + MIN_EDGE, slideW - MIN_EDGE)),
        round3(clamp((Number(y) || 0) + dy, -height + MIN_EDGE, slideH - MIN_EDGE)),
        width,
        height,
      ]
    }

    /** MIME-ish marker used when putting canvas elements on the text clipboard. */
    const ELEMENT_CLIPBOARD_KIND = 'dsh-pptx-studio/elements'

    /**
     * Serialise elements for the system clipboard.
     * @param {any[]} elements - element models.
     * @returns {string} JSON payload.
     */
    function serializeElements(elements) {
      return JSON.stringify({ kind: ELEMENT_CLIPBOARD_KIND, version: 1, elements })
    }

    /**
     * Read a canvas-element payload from clipboard text.
     * @param {unknown} text - clipboard text.
     * @returns {any[]|null} elements or null.
     */
    function parseElementClipboard(text) {
      if (typeof text !== 'string' || text.trim() === '') return null
      try {
        const data = JSON.parse(text)
        if (data && data.kind === ELEMENT_CLIPBOARD_KIND && Array.isArray(data.elements)) {
          return data.elements.filter((el) => el !== null && typeof el === 'object')
        }
      } catch {
        /* not our payload */
      }
      return null
    }

    /**
     * Absolute CSS placement for one element.
     * @param {number[]} box - `[x, y, w, h]` in inches.
     * @param {number} zoom - zoom factor.
     * @returns {Record<string, string>} CSS.
     */
    function placement(box, zoom) {
      const unit = PX_PER_IN * zoom
      // Round to 1/1000 px: `0.8 * 96 * 0.5` otherwise renders as
      // `38.400000000000006px`, which makes every miniature differ from the
      // page by a hair and makes thumb-vs-page comparisons untestable.
      const at = (value) => `${round3(value * unit)}px`
      return {
        left: at(box[0]),
        top: at(box[1]),
        width: at(box[2]),
        height: at(box[3]),
      }
    }

    // ── element rendering ────────────────────────────────────────────────────

    /**
     * Nice round value-axis ticks covering [0, max(data)] (and the negatives,
     * if any), the way PowerPoint picks automatic axis bounds.
     * @param {number[]} numbers - every data point of the chart.
     * @returns {{bottom: number, top: number, ticks: number[]}} axis scale.
     */
    function axisScale(numbers) {
      const flat = numbers.length > 0 ? numbers : [0]
      const lo = Math.min(0, ...flat)
      const hi = Math.max(1, ...flat)
      const span = hi - lo || 1
      const raw = span / 4
      const mag = Math.pow(10, Math.floor(Math.log10(raw)))
      const norm = raw / mag
      const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10) * mag
      const bottom = Math.floor(lo / step) * step
      const top = Math.ceil(hi / step) * step
      const ticks = []
      for (let value = bottom; value <= top + step / 2; value += step) {
        ticks.push(Math.round(value * 1000) / 1000)
      }
      return { bottom, top: bottom + (ticks.length - 1) * step, ticks }
    }

    /**
     * Format an axis tick the short way (integers stay integers).
     * @param {number} value - tick value.
     * @returns {string} label text.
     */
    function formatTick(value) {
      return String(Math.round(value * 100) / 100)
    }

    /**
     * Draw a native chart element as a lightweight SVG/HTML approximation.
     * The exported .pptx carries a real chart part; this canvas view mirrors
     * that part's structure (pptx_build.py#add_chart / #theme_chart): the same
     * palette, an automatic value axis with gridlines, a bottom legend for
     * multi-series charts, the single-series auto title, and value labels on
     * pie/doughnut — so what the agent draws is what the file will contain.
     * @param {Record<string, any>} el - chart element.
     * @param {Record<string, any>} ctx - render context.
     * @returns {import('react').ReactNode} node.
     */
    function renderChart(el, ctx) {
      const chart = el.chart || {}
      const kind = String(chart.kind || 'column')
      const categories = Array.isArray(chart.categories) ? chart.categories : []
      const series = Array.isArray(chart.series) ? chart.series : []
      // The canvas approximation follows the deck theme exactly like the
      // exported chart part (pptx_build.py#theme_chart): accent-led palette,
      // transparent canvas, light ink on a dark slide.
      const theme = (ctx.theme && ctx.theme.colors) || DEFAULT_THEME.colors
      const dark = isDarkBg(cssColor(ctx.background, theme.background || '#FFFFFF'))
      const palette = [cssColor(theme.accent, CHART_PALETTE[0]), ...CHART_PALETTE.slice(1)]
      const colors = series.map((item, index) => cssColor(item.color, palette[index % palette.length]))
      const textColor = cssColor(dark ? theme.surface : theme.ink, '#0f172a')
      const axisColor = cssColor(theme.muted, '#64748b')
      const trackColor = dark ? cssColor(theme.surface, '#f1f5f9') : '#eef2f7'
      const trackOpacity = dark ? 0.14 : 1
      // pptx_build sets chart.font.size = style.size || 11pt.
      const fontPx = Math.max(6, (Number(el.style && el.style.size) || 11) * pxPerPt(ctx.zoom))
      // Only an explicit chart title shows: the design sheet owns page titles,
      // and the export hides the auto title (autoTitleDeleted=1).
      const titleText = String(chart.title || '')
      const title = titleText
        ? h('div', { className: 't', style: { fontSize: `${Math.max(7, fontPx * 1.1)}px`, color: textColor } }, titleText)
        : null
      // Chart styling the design sheet can ask for (see model.js#normalizeChart);
      // pptx_build.py#add_chart honours the very same fields.
      const bareAxis = String(chart.axis || 'auto') === 'none'
      const gradient = Array.isArray(chart.gradient) && chart.gradient.length >= 2
        ? `linear-gradient(180deg, ${cssColor(chart.gradient[0], colors[0])}, ${cssColor(chart.gradient[1], colors[0])})`
        : null
      const valueLabels = chart.dataLabels === true
      const cornerCss = typeof chart.corner === 'number' && chart.corner > 0
        ? `${Math.round(chart.corner * 100)}% ${Math.round(chart.corner * 100)}% 0 0`
        : '0'
      const barBackground = gradient || colors[0]
      // Multi-series charts carry a bottom legend in the export.
      const legend = series.length > 1
        ? h('div', {
          style: {
            display: 'flex', flexWrap: 'wrap', gap: 8, justifyContent: 'center', alignItems: 'center',
            flex: '0 0 auto', fontSize: `${fontPx}px`, color: textColor,
          },
        },
          series.map((item, index) => h('span', { key: `l${index}`, style: { display: 'inline-flex', alignItems: 'center', gap: 3, minWidth: 0 } },
            h('span', { style: { width: 7, height: 7, borderRadius: 2, background: colors[index], flex: '0 0 auto' } }),
            h('span', { style: { overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' } }, String(item.name || `系列 ${index + 1}`)))))
        : null
      if (series.length === 0 || categories.length === 0) {
        return h('div', { className: 'ps-chart' }, title, h('div', { className: 'ps-hint' }, ctx.approx))
      }
      const values = series.map((item) => (Array.isArray(item.values) ? item.values : []))
      const scale = axisScale(values.flat().map((value) => Number(value) || 0))
      const pct = (value) => ((Number(value) || 0) - scale.bottom) / (scale.top - scale.bottom) * 100
      const yAxisW = Math.round(fontPx * 2.6) + 6
      const yAxisLabels = h('div', {
        style: {
          width: yAxisW, flex: '0 0 auto', display: 'flex', flexDirection: 'column',
          justifyContent: 'space-between', alignItems: 'flex-end', paddingRight: 4,
          fontSize: `${fontPx}px`, color: textColor, textAlign: 'right',
        },
      }, [...scale.ticks].reverse().map((tick, index) => h('span', { key: `y${index}`, style: { lineHeight: 1 } }, formatTick(tick))))
      const gridlines = scale.ticks.map((tick, index) => h('div', {
        key: `g${index}`,
        style: { position: 'absolute', left: 0, right: 0, bottom: `${pct(tick)}%`, height: 1, background: axisColor, opacity: 0.35 },
      }))
      const zeroLine = h('div', { style: { position: 'absolute', left: 0, right: 0, bottom: `${pct(0)}%`, height: 1, background: axisColor } })

      if (kind === 'pie' || kind === 'doughnut') {
        // Export turns data labels on for pie/doughnut and shows values only,
        // so the canvas lists the values (category names ride in the tooltip).
        const total = values[0].reduce((sum, value) => sum + Math.max(0, Number(value) || 0), 0) || 1
        let cursor = 0
        const stops = values[0].map((value, index) => {
          const start = cursor
          cursor += (Math.max(0, Number(value) || 0) / total) * 100
          return `${colors[index % colors.length]} ${start}% ${cursor}%`
        })
        const hole = kind === 'doughnut' ? 'radial-gradient(circle at 50% 50%, ' + cssColor(el.fill, 'transparent') + ' 0 42%, transparent 43%)' : null
        return h('div', { className: 'ps-chart' },
          title,
          h('div', { style: { flex: '1 1 auto', minHeight: 0, display: 'flex', alignItems: 'center', gap: 8 } },
            h('div', { style: { flex: '0 0 auto', width: '58%', height: '100%', position: 'relative' } },
              h('div', {
                style: {
                  position: 'absolute', inset: '8%', borderRadius: '50%',
                  background: hole ? `${hole}, conic-gradient(${stops.join(',')})` : `conic-gradient(${stops.join(',')})`,
                },
              })),
            h('div', { style: { display: 'flex', flexDirection: 'column', gap: 2, fontSize: `${fontPx}px`, color: textColor, minWidth: 0 } },
              categories.map((category, index) => h('div', { key: `l${index}`, title: String(category), style: { display: 'flex', alignItems: 'center', gap: 4 } },
                h('span', { style: { width: 7, height: 7, borderRadius: 2, background: colors[index % colors.length], display: 'inline-block', flex: '0 0 auto' } }),
                h('span', { style: { overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' } }, formatTick(Number(values[0][index]) || 0)))))),
          legend)
      }

      if (kind === 'line' || kind === 'area') {
        const step = categories.length > 1 ? 100 / (categories.length - 1) : 0
        const paths = values.map((line) => line.map((value, index) => {
          const x = categories.length > 1 ? index * step : 50
          const y = 100 - pct(value)
          return `${Math.round(x * 100) / 100},${Math.round(y * 100) / 100}`
        }))
        const markerPx = Math.max(3, 2.4 * ctx.zoom)
        return h('div', { className: 'ps-chart' },
          title,
          h('div', { style: { flex: '1 1 auto', minHeight: 0, display: 'flex' } },
            yAxisLabels,
            h('div', { style: { flex: '1 1 auto', minWidth: 0, position: 'relative' } },
              h('svg', { viewBox: '0 0 100 100', preserveAspectRatio: 'none', style: { width: '100%', height: '100%' } },
                scale.ticks.map((tick, index) => h('line', { key: `g${index}`, x1: 0, x2: 100, y1: 100 - pct(tick), y2: 100 - pct(tick), stroke: axisColor, strokeWidth: 0.4, opacity: 0.35 })),
                h('line', { x1: 0, x2: 100, y1: 100 - pct(0), y2: 100 - pct(0), stroke: axisColor, strokeWidth: 0.6 }),
                paths.map((points, index) => h(kind === 'area' ? 'polygon' : 'polyline', {
                  key: `s${index}`,
                  points: kind === 'area' ? `0,100 ${points.join(' ')} 100,100` : points.join(' '),
                  fill: kind === 'area' ? colors[index % colors.length] : 'none',
                  fillOpacity: kind === 'area' ? 0.18 : 0,
                  stroke: colors[index % colors.length],
                  strokeWidth: 1.6,
                }))),
              // markers sit outside the stretched svg so they stay round
              ...(kind === 'line' ? values.flatMap((line, seriesIndex) => line.map((value, index) => h('div', {
                key: `m${seriesIndex}-${index}`,
                title: `${String(series[seriesIndex].name || '')} · ${String(categories[index] || '')}`,
                style: {
                  position: 'absolute', left: `${categories.length > 1 ? (index / (categories.length - 1)) * 100 : 50}%`,
                  top: `${100 - pct(value)}%`, width: markerPx, height: markerPx, borderRadius: '50%',
                  background: colors[seriesIndex % colors.length], transform: 'translate(-50%, -50%)',
                },
              }))) : []))),
          h('div', { className: 'cats', style: { marginLeft: yAxisW, fontSize: `${fontPx}px`, color: textColor } },
            categories.map((category, index) => h('span', { key: `c${index}` }, String(category)))),
          legend)
      }

      if (kind === 'bar') {
        const labelW = Math.round(fontPx * 4) + 8
        const rowH = series.length > 1 ? 3 : 8
        return h('div', { className: 'ps-chart' },
          title,
          h('div', { style: { flex: '1 1 auto', minHeight: 0, display: 'flex', flexDirection: 'column' } },
            h('div', { style: { flex: '1 1 auto', minHeight: 0, position: 'relative', display: 'flex', flexDirection: 'column', justifyContent: 'space-around' } },
              scale.ticks.map((tick, index) => h('div', {
                key: `g${index}`,
                style: { position: 'absolute', top: 0, bottom: 0, left: `${pct(tick)}%`, width: 1, background: axisColor, opacity: 0.35 },
              })),
              categories.map((category, index) => h('div', { key: `b${index}`, style: { display: 'flex', alignItems: 'center', gap: 4 } },
                h('span', { style: { width: labelW, flex: '0 0 auto', fontSize: `${fontPx}px`, color: textColor, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' } }, String(category)),
                h('span', { style: { flex: '1 1 auto', minWidth: 0, display: 'flex', flexDirection: 'column', gap: 1 } },
                  series.map((item, seriesIndex) => h('span', { key: `s${seriesIndex}`, title: String(item.name || ''), style: { display: 'block', height: rowH, background: trackColor, opacity: trackOpacity, borderRadius: 1, overflow: 'hidden' } },
                    h('span', { style: { display: 'block', height: '100%', width: `${Math.max(0, pct(Number(values[seriesIndex][index]) || 0))}%`, background: gradient || colors[seriesIndex % colors.length], borderRadius: 1 } }))))))),
            h('div', { style: { marginLeft: labelW + 4, display: 'flex', justifyContent: 'space-between', fontSize: `${fontPx}px`, color: textColor, flex: '0 0 auto' } },
              scale.ticks.map((tick, index) => h('span', { key: `t${index}` }, formatTick(tick))))),
          legend)
      }

      // column: value axis on the left, gridlines, zero axis, category axis —
      // or, with axis:"none", bare bars with value labels like the design sheet.
      const plotArea = h('div', { style: { flex: '1 1 auto', minWidth: 0, position: 'relative' } },
        bareAxis ? null : gridlines,
        h('div', { style: { position: 'absolute', inset: 0, display: 'flex', alignItems: 'flex-end', gap: 6 } },
          categories.map((category, index) => h('div', { key: `g${index}`, style: { flex: 1, minWidth: 0, display: 'flex', alignItems: 'flex-end', justifyContent: 'center', gap: 2, height: '100%' } },
            series.map((item, seriesIndex) => h('div', {
              key: `b${seriesIndex}`,
              style: {
                width: '38%', maxWidth: 22, minHeight: 1,
                height: `${Math.max(0, pct(Number(values[seriesIndex][index]) || 0))}%`,
                display: 'flex', flexDirection: 'column', justifyContent: 'flex-end',
              },
            },
              valueLabels
                ? h('span', { style: { fontSize: `${fontPx}px`, fontWeight: 600, color: textColor, textAlign: 'center', lineHeight: 1.2, marginBottom: 2 } }, formatTick(Number(values[seriesIndex][index]) || 0))
                : null,
              h('span', { style: { flex: '1 1 auto', minHeight: 1, background: gradient || colors[seriesIndex % colors.length], borderRadius: cornerCss } }))))),
        bareAxis ? null : zeroLine,
        bareAxis ? null : h('div', { style: { position: 'absolute', left: 0, top: 0, bottom: 0, width: 1, background: axisColor } })))
      return h('div', { className: 'ps-chart' },
        title,
        h('div', { style: { flex: '1 1 auto', minHeight: 0, display: 'flex' } },
          bareAxis ? null : yAxisLabels,
          plotArea),
        h('div', { className: 'cats', style: { marginLeft: bareAxis ? 0 : yAxisW, fontSize: `${fontPx}px`, color: textColor } },
          categories.map((category, index) => h('span', { key: `c${index}` }, String(category)))),
        legend)
    }

    /**
     * Render one element on the slide surface.
     * @param {Record<string, any>} el - element model.
     * @param {Record<string, any>} ctx - render context.
     * @returns {import('react').ReactNode} node.
     */
    function renderElement(el, ctx) {
      const style = el.style || {}
      const { zoom, deckId, rev } = ctx
      const box = ctx.overrideBox || boxOf(el, ctx.slideW, ctx.slideH)
      const fontPx = Math.max(3, (Number(style.size) || 18) * pxPerPt(zoom))
      const color = cssColor(style.color, '#0f172a')
      const css = { ...placement(box, zoom), fontSize: `${fontPx}px`, color }
      if (style.font) css.fontFamily = `${style.font}, "Microsoft YaHei", sans-serif`
      if (style.italic) css.fontStyle = 'italic'
      if (style.bold) css.fontWeight = 700
      if (style.align) css.textAlign = style.align
      if (el.fill) css.background = cssColor(el.fill, 'transparent')
      // A `line` without `lineWidth` still draws 1pt in the .pptx, so mirror
      // that instead of dropping the stroke. No `line` means no border at all —
      // the old blanket `border-style: solid` painted a 3px black frame on
      // every rectangle.
      if (el.line) {
        const stroke = Number(el.lineWidth) > 0 ? Number(el.lineWidth) : 1
        css.border = `${Math.max(1, stroke * pxPerPt(zoom))}px solid ${cssColor(el.line, '#94a3b8')}`
      } else {
        css.border = '0'
      }
      if (el.shadow) {
        const shadow = shadowCss(el.shadow)
        if (shadow) {
          css.boxShadow = shadow
          if (el.type === 'text') css.textShadow = shadow
        }
      }

      let body = null
      if (el.type === 'text') {
        const anchor = style.anchor || 'top'
        const paragraphs = Array.isArray(el.paragraphs) && el.paragraphs.length > 0 ? el.paragraphs : [{ text: el.text || '' }]
        body = h('div', {
          style: {
            width: '100%', height: '100%', display: 'flex', flexDirection: 'column',
            justifyContent: anchor === 'middle' ? 'center' : anchor === 'bottom' ? 'flex-end' : 'flex-start',
          },
        }, paragraphs.map((para, i) => {
          const runs = Array.isArray(para.runs) && para.runs.length > 0 ? para.runs : null
          return h('div', {
            key: `p${i}`,
            style: {
              fontSize: `${Math.max(3, (para.size ? Number(para.size) : Number(style.size) || 18) * pxPerPt(zoom))}px`,
              fontWeight: para.bold || style.bold ? 700 : 400,
              fontStyle: para.italic || style.italic ? 'italic' : 'normal',
              textDecoration: para.underline || style.underline ? 'underline' : 'none',
              color: cssColor(para.color, color),
              textAlign: para.align || style.align || 'left',
              lineHeight: String(Number(style.lineSpacing) > 0 ? style.lineSpacing : 1.2),
              whiteSpace: 'pre-wrap',
              wordBreak: 'break-word',
            },
          }, runs
            ? runs.map((run, index) => h('span', {
                key: `r${index}`,
                style: {
                  fontSize: run.size ? `${Math.max(3, Number(run.size) * pxPerPt(zoom))}px` : undefined,
                  fontWeight: run.bold === true ? 700 : undefined,
                  fontStyle: run.italic === true ? 'italic' : undefined,
                  textDecoration: run.underline === true ? 'underline' : undefined,
                  color: run.color ? cssColor(run.color, color) : undefined,
                  fontFamily: run.font ? `${run.font}, "Microsoft YaHei", sans-serif` : undefined,
                },
              }, String(run.text ?? '')))
            : String(para.text ?? ''))
        }))
      } else if (el.type === 'rect') {
        const radius = Number(el.radius)
        if (radius > 0) css.borderRadius = `${radius * PX_PER_IN * zoom * 0.25}px`
        body = el.text
          ? h('div', {
              style: {
                width: '100%', height: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center',
                textAlign: style.align || 'center', padding: `${2 * zoom}px ${6 * zoom}px`,
                whiteSpace: 'pre-wrap', wordBreak: 'break-word', fontSize: `${fontPx}px`, color,
                fontWeight: style.bold ? 700 : 400, fontStyle: style.italic ? 'italic' : 'normal',
                textDecoration: style.underline ? 'underline' : 'none', lineHeight: '1.2',
              },
            }, String(el.text))
          : null
      } else if (el.type === 'chart') {
        body = renderChart(el, ctx)
      } else if (el.type === 'image') {
        const src = assetUrl(deckId, el.src, rev)
        body = src
          ? h('img', { src, alt: el.name || 'image', draggable: false, loading: 'lazy' })
          : h('div', { className: 'ps-img-fallback' }, ctx.approx)
      } else if (el.type === 'line') {
        css.background = cssColor(el.line, cssColor(el.fill, '#e2e8f0'))
        css.border = '0'
      } else if (el.type === 'table') {
        const rows = (el.table && el.table.rows) || []
        const cellPx = Math.max(3, (Number(style.size) || 11) * pxPerPt(zoom))
        body = h('table', null, h('tbody', null, rows.map((row, r) => h('tr', { key: `r${r}` },
          (Array.isArray(row) ? row : []).map((cell, c) => h('td', {
            key: `c${c}`,
            style: { fontSize: `${cellPx}px`, fontWeight: r === 0 ? 600 : 400 },
          }, String(cell ?? '')))))))
      } else {
        body = h('div', null, `${ctx.approx} · ${el.type}${el.text ? ` · ${String(el.text).slice(0, 40)}` : ''}`)
      }

      const handlers = ctx.interactive
        ? {
            onPointerDown: (event) => ctx.onPointerDown(event, el, box),
            onPointerMove: ctx.onPointerMove,
            onPointerUp: ctx.onPointerUp,
            onPointerCancel: ctx.onPointerUp,
            onDoubleClick: (event) => {
              event.stopPropagation()
              ctx.onDoubleClick(el)
            },
          }
        : {}

      const anim = ctx.animPlay && el.anim && el.anim.effect && el.anim.effect !== 'none' ? el.anim : null
      if (anim) {
        css.animationDuration = `${Math.max(100, Number(anim.duration) || 500)}ms`
        css.animationDelay = `${Math.max(0, Number(anim.delay) || 0) + (Number(anim.order) || 0) * 80}ms`
      }
      return h('div', {
        key: el.id,
        className: 'ps-el',
        'data-type': el.type,
        'data-id': el.id,
        'data-selected': ctx.selected ? 'true' : 'false',
        'data-fresh': ctx.fresh ? 'true' : 'false',
        'data-flash': ctx.flash ? 'true' : 'false',
        // Still being drawn: shows as a pencil ghost until its turn comes.
        'data-sketch': ctx.sketch ? 'true' : 'false',
        ...(anim ? { 'data-anim': anim.effect, 'data-anim-dir': anim.direction || 'left' } : {}),
        style: css,
      }, body, h('div', { className: 'ps-hit', ...handlers }))
    }

    /**
     * Miniature of one slide for the thumbnail rail (real geometry, no text
     * editing: purely a visual cue for which page is which).
     * @param {Record<string, any>} slide - slide model.
     * @param {Record<string, any>} ctx - render context without interaction.
     * @returns {import('react').ReactNode} node.
     */
    function renderThumb(slide, ctx) {
      // Render the page at a design size, then scale the whole page down. Drawing
      // the elements directly at a 1.65px font size (the old `zoom = 88 / slidePx`)
      // made text smear, tables collapse and line breaks land differently, so the
      // miniature did not look like the page at all. Scaling a 0.5× page keeps the
      // geometry, colours and relative type sizes intact.
      const designZoom = Number(ctx.designZoom) > 0 ? Number(ctx.designZoom) : THUMB_DESIGN_ZOOM
      const designWidth = ctx.slideW * PX_PER_IN * designZoom
      const designHeight = ctx.slideH * PX_PER_IN * designZoom
      const scale = (Number(ctx.thumbW) || 88) / designWidth
      const background = ctx.background || cssColor(null, '#FFFFFF')
      return h('div', {
        className: 'ps-thumb-inner',
        'data-ps-thumb-inner': '1',
        'data-ps-thumb-scale': String(Math.round(scale * 10000) / 10000),
        style: {
          position: 'absolute',
          left: '0',
          top: '0',
          width: `${designWidth}px`,
          height: `${designHeight}px`,
          background: cssColor(background, '#FFFFFF'),
          transform: `scale(${scale})`,
          transformOrigin: 'top left',
          pointerEvents: 'none',
        },
      }, (slide.elements || []).map((el) => renderElement(el, {
        ...ctx,
        zoom: designZoom,
        interactive: false,
        selected: false,
        fresh: false,
        flash: false,
        sketch: false,
        animPlay: false,
        overrideBox: null,
      })))
    }

    // ── canvas body ──────────────────────────────────────────────────────────

    /**
     * Right-side canvas body.
     * @param {any} props - session-scope slot props.
     * @returns {import('react').ReactNode} node.
     */
    function CanvasBody(props) {
      const t = typeof props.t === 'function' ? props.t : (key, values) => fmt(key, values)
      const sessionId = props.sessionId
      const tabInfo = readTabInfo(props.useTabInfo)
      const tab = tabInfo?.tab || null
      const params = (tab && tab.navigation && tab.navigation.params) || {}

      const [decks, setDecks] = useState([])
      const [confirmDelete, setConfirmDelete] = useState(false)
      const paramDeckId = typeof params.deckId === 'string' ? params.deckId : ''
      // A tool card opens the canvas before a *new* deck exists and marks the
      // navigation "pending": the shell has to show that step's progress
      // instead of falling back to the deck of the previous build.
      const paramPending = typeof params.pending === 'string' && params.pending !== ''
      const paramNonce = params.nonce === undefined ? '' : String(params.nonce)
      // Every `openTab` records its own navigation; `at` makes a repeat open of
      // the same deck a new intent instead of an ignored duplicate.
      const paramSignature = `${paramDeckId}|${params.at === undefined ? '' : String(params.at)}`
      const [deckId, setDeckId] = useState(paramDeckId)
      const [deck, setDeck] = useState(null)
      const [phase, setPhase] = useState({ phase: 'idle', message: '', slideIndex: 0, total: 0 })
      const [saved, setSaved] = useState(null)
      const [error, setError] = useState('')
      const [notice, setNotice] = useState('')
      const [loading, setLoading] = useState(true)
      const [live, setLive] = useState(false)
      const [follow, setFollow] = useState(true)
      const [slideIndex, setSlideIndex] = useState(0)
      const [selection, setSelection] = useState(null) // {kind:'element', slideId, elementId} | {kind:'region', slideId, rect}
      const [multiSelection, setMultiSelection] = useState(null) // element ids (Ctrl+A / paste)
      const [note, setNote] = useState('')
      const [sending, setSending] = useState(false)
      const [feedbackHint, setFeedbackHint] = useState('')
      const [freshEls, setFreshEls] = useState({})
      const [flashEls, setFlashEls] = useState({})
      // Elements that exist in the model but are still being "drawn": they show
      // as pencil ghosts until their turn comes, so a page never lands in one
      // silent frame.
      const [sketchEls, setSketchEls] = useState({})
      // Live drawing HUD: which slide is being drawn and how many elements of it
      // have settled. `null` means idle.
      const [draw, setDraw] = useState(null)
      const [zoomMode, setZoomMode] = useState('fit')
      const [fit, setFit] = useState(0.5)
      const [preview, setPreview] = useState(null) // {elementId, box}
      const [marquee, setMarquee] = useState(null) // {x1,y1,x2,y2} in inches
      const [editing, setEditing] = useState(null) // {elementId, value}
      const [dirty, setDirty] = useState(0)
      const [busy, setBusy] = useState('') // '' | 'export' | 'upload'
      const [dragOver, setDragOver] = useState(false)
      const [history, setHistory] = useState({ undo: 0, redo: 0 })
      const [titleDraft, setTitleDraft] = useState('')
      const [titleEditing, setTitleEditing] = useState(false)
      const [view, setView] = useState('auto')
      const [template, setTemplate] = useState(null)
      // Decomposed template catalogue (pages/roles/fields/palette), served by
      // `GET /api/template`. `template` above is the PowerPoint layout list that
      // comes with `/api/state`; this one is the template-first workflow spec.
      const [templateInfo, setTemplateInfo] = useState(null)
      const [templateError, setTemplateError] = useState('')
      const [templateNonce, setTemplateNonce] = useState(0)
      const [settingsOpen, setSettingsOpen] = useState(false)
      const [themeOpen, setThemeOpen] = useState(false)
      const [fx, setFx] = useState(null)
      const [notesDraft, setNotesDraft] = useState('')
      const [stageNote, setStageNote] = useState('')
      const [stageSending, setStageSending] = useState(false)
      const [previewNonce, setPreviewNonce] = useState(0)
      const [previewBoxWidth, setPreviewBoxWidth] = useState(0)
      const [shortcutsOpen, setShortcutsOpen] = useState(false)
      // The running pptx step of this session: it names the deck being written
      // and the step, so the canvas switches to it and says what is happening.
      const task = useSessionTask(sessionId)
      const taskRunning = Boolean(task && task.state === 'running')
      const taskSeconds = useElapsed(task)

      const stageRef = useRef(null)
      const previewRef = useRef(null)
      const fxTimer = useRef(null)
      const canvasRef = useRef(null)
      const fileRef = useRef(null)
      const freshTimers = useRef(new Map())
      const drawTimers = useRef(new Set())
      const drawIdleTimer = useRef(null)
      const drawToken = useRef(0)
      /** @type {any} */
      const deleteTimer = useRef(null)
      const dragRef = useRef(null)
      const clipboardRef = useRef(null)
      /** True while the tab waits for the deck of a call that has not created it yet. */
      const awaitingDeckRef = useRef(paramPending)
      /** First task snapshot of this mount (a finished one must not yank the tab). */
      const firstTaskRef = useRef(true)
      /** Navigation signature already applied by the param effect. */
      const lastParamRef = useRef(paramSignature)
      const undoRef = useRef([])
      const redoRef = useRef([])
      const deckRef = useRef(null)
      /** Deck id of the current render; in-flight loads compare against it. */
      const deckIdRef = useRef(deckId)
      const followRef = useRef(true)
      const selectionRef = useRef(null)
      const pendingRef = useRef(0)
      const slideRef = useRef(null)
      const rootRef = useRef(null)
      // `t` may be a fresh closure every render; the SSE stream must not be
      // torn down and rebuilt on each of those (that would reconnect forever).
      const tRef = useRef(null)

      const slides = (deck && deck.slides) || []
      const size = (deck && deck.size) || { w: 13.333, h: 7.5 }
      const slideW = Number(size.w) || 13.333
      const slideH = Number(size.h) || 7.5
      const safeIndex = clamp(slideIndex, 0, Math.max(0, slides.length - 1))
      const current = slides[safeIndex] || null
      const zoom = zoomMode === 'fit' ? fit : Number(zoomMode) || 1
      const rev = (deck && deck.meta && deck.meta.updatedAt) || ''
      const stage = stageOf(deck)
      const autoView = stage === 'template' ? 'template' : stage === 'research' ? 'research' : stage === 'preview' ? 'preview' : (slides.length === 0 ? 'preview' : 'slides')
      const activeView = view === 'auto' ? autoView : view
      const theme = themeOf(deck)
      const research = (deck && deck.pipeline && deck.pipeline.research) || null
      const htmlPreview = (deck && deck.pipeline && deck.pipeline.preview) || null
      // Template-first workflow: `pipeline.template` exists only for decks opened
      // with a .pptx template. `decomposedAt` is what the host flips when
      // `pptx_studio_template` finished, so it drives the ⓪ step's done state.
      const pipelineTemplate = (deck && deck.pipeline && deck.pipeline.template) || null
      const templateDecomposed = Boolean(pipelineTemplate && pipelineTemplate.decomposedAt)
      // Re-decomposing an already-decomposed template changes `decomposedAt` but
      // not the boolean, so the timestamp is what the fetch effect keys on.
      const templateStamp = pipelineTemplate && pipelineTemplate.decomposedAt ? String(pipelineTemplate.decomposedAt) : ''
      const templatePages = templateInfo && Array.isArray(templateInfo.pages)
        ? templateInfo.pages.length
        : (pipelineTemplate && Number(pipelineTemplate.pages) > 0 ? Number(pipelineTemplate.pages) : 0)
      const previewUrl = deckId
        ? api(`preview?deck=${encodeURIComponent(deckId)}&v=${encodeURIComponent((htmlPreview && htmlPreview.updatedAt) || '')}&r=${previewNonce}`)
        : ''
      const previewSize = inferPreviewSize(htmlPreview, slideW, slideH)
      const viewportHeight = typeof window !== 'undefined' && Number(window.innerHeight) > 0 ? Number(window.innerHeight) : 720
      const fallbackPreviewWidth = typeof window !== 'undefined' && Number(window.innerWidth) > 0
        ? Math.min(860, Number(window.innerWidth) * 0.55)
        : 720
      const previewScale = htmlPreview
        ? Math.max(0.2, Math.min(1.4, (previewBoxWidth > 40 ? previewBoxWidth : fallbackPreviewWidth) / previewSize.width, (viewportHeight * 0.8) / previewSize.height))
        : 1
      const previewDisplayHeight = Math.max(220, Math.round(previewSize.height * previewScale))

      deckRef.current = deck
      deckIdRef.current = deckId
      followRef.current = follow
      selectionRef.current = selection
      slideRef.current = current
      tRef.current = t

      /**
       * Mark elements as freshly drawn (fade in) or patched (flash).
       * @param {string[]} ids - element ids.
       * @param {boolean} flashMode - true for the "changed" pulse.
       * @returns {void}
       */
      const markFresh = useCallback((ids, flashMode = false) => {
        const list = (ids || []).filter(Boolean)
        if (list.length === 0) return
        const stamp = Date.now()
        const next = {}
        for (const id of list) next[id] = stamp
        if (flashMode) setFlashEls((prev) => ({ ...prev, ...next }))
        else setFreshEls((prev) => ({ ...prev, ...next }))
        for (const id of list) {
          const key = `${flashMode ? 'f' : 'n'}:${id}`
          const existing = freshTimers.current.get(key)
          if (existing) clearTimeout(existing)
          const timer = setTimeout(() => {
            freshTimers.current.delete(key)
            if (flashMode) setFlashEls((prev) => { const copy = { ...prev }; delete copy[id]; return copy })
            else setFreshEls((prev) => { const copy = { ...prev }; delete copy[id]; return copy })
          }, flashMode ? 1000 : 700)
          freshTimers.current.set(key, timer)
        }
      }, [])

      /**
       * Run `fn` later, remembering the timer so a deck switch / unmount can
       * cancel the whole drawing queue in one go.
       * @param {() => void} fn - callback.
       * @param {number} delay - milliseconds.
       * @returns {void}
       */
      const later = useCallback((fn, delay) => {
        const timer = setTimeout(() => {
          drawTimers.current.delete(timer)
          fn()
        }, Math.max(0, delay))
        drawTimers.current.add(timer)
      }, [])

      /**
       * Stop every pending reveal and reset the drawing HUD.
       * @param {{keepHud?: boolean}} [options] - keep the HUD visible.
       * @returns {void}
       */
      const stopDrawing = useCallback((options = {}) => {
        for (const timer of drawTimers.current) clearTimeout(timer)
        drawTimers.current.clear()
        if (drawIdleTimer.current) {
          clearTimeout(drawIdleTimer.current)
          drawIdleTimer.current = null
        }
        drawToken.current += 1
        setSketchEls({})
        if (!options.keepHud) setDraw(null)
      }, [])

      /**
       * Draw one slide in front of the user: the page lands as pencil ghosts,
       * then each element settles in array order (≈40–75ms apart), and the HUD
       * counts the elements as they land.
       * @param {any} slide - slide model just written.
       * @param {number} index - zero-based slide index.
       * @param {number} total - slide count, for the HUD.
       * @returns {void}
       */
      const revealSlide = useCallback((slide, index, total) => {
        const elements = Array.isArray(slide && slide.elements) ? slide.elements : []
        const ids = elements.map((el) => el && el.id).filter(Boolean)
        if (ids.length === 0) return
        const token = drawToken.current + 1
        drawToken.current = token
        if (drawIdleTimer.current) {
          clearTimeout(drawIdleTimer.current)
          drawIdleTimer.current = null
        }
        // Only the current page can be on screen, so the ghost map is replaced
        // rather than merged: no stale pencil from a superseded page can linger.
        setSketchEls(() => {
          /** @type {Record<string, boolean>} */
          const next = {}
          for (const id of ids) next[id] = true
          return next
        })
        setDraw({ slideIndex: index + 1, total: Math.max(Number(total) || 0, index + 1), done: 0, count: ids.length })
        const step = ids.length >= 9 ? 42 : ids.length >= 5 ? 55 : 75
        ids.forEach((id, i) => {
          later(() => {
            // Always drop the ghost: a superseded reveal must not leave a
            // permanently pencilled element behind.
            setSketchEls((prev) => {
              if (!prev[id]) return prev
              const next = { ...prev }
              delete next[id]
              return next
            })
            if (drawToken.current !== token) return
            markFresh([id])
            setDraw((prev) => (prev ? { ...prev, done: Math.max(prev.done, i + 1) } : prev))
          }, 70 + i * step)
        })
        // Let the last element finish, then fade the HUD out. A newer slide (or
        // a stopDrawing) bumps the token and keeps this timer from clearing it.
        drawIdleTimer.current = setTimeout(() => {
          drawIdleTimer.current = null
          if (drawToken.current === token) setDraw(null)
        }, 70 + ids.length * step + 700)
      }, [later, markFresh])

      /**
       * Draw one element that was appended on its own (fine-grained drawing).
       * @param {string} id - element id.
       * @returns {void}
       */
      const revealElement = useCallback((id) => {
        if (!id) return
        setSketchEls((prev) => ({ ...prev, [id]: true }))
        later(() => {
          setSketchEls((prev) => {
            if (!prev[id]) return prev
            const next = { ...prev }
            delete next[id]
            return next
          })
          markFresh([id], true)
        }, 90)
      }, [later, markFresh])

      // Timers are cleaned up on unmount.
      useEffect(() => () => {
        for (const timer of freshTimers.current.values()) clearTimeout(timer)
        freshTimers.current.clear()
        for (const timer of drawTimers.current) clearTimeout(timer)
        drawTimers.current.clear()
        if (drawIdleTimer.current) clearTimeout(drawIdleTimer.current)
        if (fxTimer.current) clearTimeout(fxTimer.current)
      }, [])

      // Keep the notes editor in sync when the slide changes (or the model is
      // updated by the agent), without clobbering a draft on every keystroke.
      useEffect(() => {
        setNotesDraft(current ? String(current.notes || '') : '')
      }, [current && current.id, current && current.notes])

      /**
       * Load the deck list; keeps the current selection when possible.
       * @returns {Promise<any[]>} decks.
       */
      const loadDecks = useCallback(async () => {
        try {
          // The session id lets the host include decks that exist on disk but
          // were forgotten by a restart (they are not in its in-memory map).
          const query = sessionId ? `decks?sessionId=${encodeURIComponent(sessionId)}` : 'decks'
          const data = await request(query, {})
          const all = Array.isArray(data.decks) ? data.decks : []
          const mine = all.filter((item) => !item.sessionId || item.sessionId === sessionId)
          const list = mine.length > 0 ? mine : all
          setDecks(list)
          // A tab opened for a *new* call must not fall back to the newest deck
          // on disk: that is the previous build, and showing it is exactly the
          // "it opened the old pptx" complaint. Wait for the new deck instead.
          if (deckId === '' && list.length > 0 && !awaitingDeckRef.current) setDeckId(list[0].id)
          return list
        } catch (err) {
          setError(err instanceof Error ? err.message : String(err))
          return []
        }
      }, [deckId, sessionId])

      /**
       * Load one deck's model.
       * @param {string} id - deck id.
       * @returns {Promise<void>} resolved when loaded.
       */
      const loadState = useCallback(async (id) => {
        if (!id) return
        try {
          setLoading(true)
          const params = new URLSearchParams({ deck: id })
          if (sessionId) params.set('sessionId', sessionId)
          const data = await request(`state?${params.toString()}`, {})
          // The tab may have switched (or been cleared for a new call) while
          // this response was in flight: an older deck must not land on top of
          // the newer one — that is how "the old pptx reappeared" happens.
          if (deckIdRef.current !== id) return
          setDeck(data.deck || null)
          setTemplate(data.template || null)
          // A reload is a deliberate jump: drop any reveal still in flight so a
          // stale timer cannot ghost elements on the freshly loaded model.
          stopDrawing()
          setFreshEls({})
          setFlashEls({})
          setSelection(null)
          setMultiSelection(null)
          setEditing(null)
          setView('auto')
          setSettingsOpen(false)
          setThemeOpen(false)
          setFx(null)
          undoRef.current = []
          redoRef.current = []
          setHistory({ undo: 0, redo: 0 })
          setError('')
        } catch (err) {
          setError(err instanceof Error ? err.message : String(err))
        } finally {
          setLoading(false)
        }
      }, [sessionId, stopDrawing])

      /**
       * Delete the deck on screen, then fall back to the next one (or to the
       * empty state). Two clicks are required: the button arms itself first.
       * @returns {Promise<void>} resolved when the list and the model settled.
       */
      const removeDeck = useCallback(async () => {
        const id = deckId
        if (!id) return
        try {
          const params = new URLSearchParams()
          if (sessionId) params.set('sessionId', sessionId)
          const query = params.toString() ? `?${params.toString()}` : ''
          await request(`deck/delete${query}`, { method: 'POST', body: { deckId: id } })
          setDeck(null)
          setSelection(null)
          setMultiSelection(null)
          setEditing(null)
          setSaved(null)
          setSlideIndex(0)
          undoRef.current = []
          redoRef.current = []
          setHistory({ undo: 0, redo: 0 })
          const list = (await loadDecks()).filter((item) => item.id !== id)
          setDecks(list)
          setDeckId(list[0]?.id ?? '')
          setNotice(list.length > 0 ? t('deck.deleted') : t('deck.empty'))
          if (list.length === 0) setLoading(false)
        } catch (err) {
          setError(fmt('deck.failed', { message: err instanceof Error ? err.message : String(err) }))
        } finally {
          setConfirmDelete(false)
        }
      }, [deckId, sessionId, loadDecks, t])

      // An armed delete button disarms itself, so a stray click cannot linger.
      useEffect(() => {
        if (!confirmDelete) return undefined
        deleteTimer.current = setTimeout(() => setConfirmDelete(false), 4000)
        return () => { if (deleteTimer.current) clearTimeout(deleteTimer.current) }
      }, [confirmDelete])

      // First load: decks, then the selected deck. A pending call has no deck
      // yet on purpose — the shell must not load the previous build.
      useEffect(() => {
        let alive = true
        ;(async () => {
          const list = await loadDecks()
          if (!alive) return
          const target = deckId || (awaitingDeckRef.current ? '' : list[0]?.id) || ''
          if (target) await loadState(target)
          else setLoading(false)
        })()
        return () => { alive = false }
        // eslint-disable-next-line react-hooks/exhaustive-deps
      }, [sessionId])

      useEffect(() => { if (deckId) void loadState(deckId) }, [deckId, loadState])

      // Template catalogue for the ⓪ panel. Only fetched once the host reports
      // the template as decomposed (`pipeline.template.decomposedAt`); before
      // that the route answers 404, and the panel shows the "in progress" copy
      // instead of an error. `templateNonce` re-runs it on demand.
      useEffect(() => {
        if (!deckId || !templateDecomposed) {
          setTemplateInfo(null)
          setTemplateError('')
          return undefined
        }
        let alive = true
        const params = new URLSearchParams({ deck: deckId })
        if (sessionId) params.set('sessionId', sessionId)
        request(`template?${params.toString()}`, {})
          .then((data) => {
            if (!alive) return
            setTemplateInfo(data && data.ok !== false ? data : null)
            setTemplateError('')
          })
          .catch((err) => {
            if (!alive) return
            setTemplateInfo(null)
            setTemplateError(err instanceof Error ? err.message : String(err))
          })
        return () => { alive = false }
      }, [deckId, sessionId, templateDecomposed, templateStamp, templateNonce])

      // The tool card navigates the already-open tab to the deck it just opened
      // (`openTab` rewrites the tab's navigation params). Without this the tab
      // would keep showing the previous deck while the new one is drawn. Only a
      // *new* navigation may move the canvas: re-applying an unchanged param on
      // every render would drag it back after a running call took the screen.
      useEffect(() => {
        if (!paramDeckId) return
        awaitingDeckRef.current = false
        if (paramSignature === lastParamRef.current) return
        lastParamRef.current = paramSignature
        if (paramDeckId === deckId) return
        setDeckId(paramDeckId)
        setSlideIndex(0)
        setFollow(true)
        setSelection(null)
        setMultiSelection(null)
        setEditing(null)
        setView('auto')
      }, [paramDeckId, paramSignature, deckId])

      // A card that opened the canvas for a call whose deck does not exist yet
      // marks the navigation "pending". Do not fall back to the deck of the
      // previous build: show the preparing shell for *this* call instead. The
      // nonce changes on every open, so a second new call re-arms it. A stale
      // flag (a page reload, or a call that never produced a deck) must not
      // leave the canvas empty forever, so it expires with its nonce.
      useEffect(() => {
        if (!paramPending) return
        const armedAt = Number(paramNonce) || 0
        const stale = armedAt > 0 && Date.now() - armedAt > PENDING_GRACE_MS
        if (stale) {
          awaitingDeckRef.current = false
          void loadDecks()
          return
        }
        awaitingDeckRef.current = true
        setDeck(null)
        setDeckId('')
        setSlideIndex(0)
        setFollow(true)
        setSelection(null)
        setMultiSelection(null)
        setEditing(null)
        setView('auto')
        setLoading(false)
        setError('')
      }, [paramPending, paramNonce, loadDecks])

      // The host publishes the deck each running call writes. A tab that is
      // showing the previous pptx switches here — that is the whole point of the
      // canvas — even when the tool card that opened it is already gone.
      useEffect(() => {
        if (!task) return
        const running = Boolean(task.state === 'running')
        const target = typeof task.deckId === 'string' ? task.deckId : ''
        const first = firstTaskRef.current
        firstTaskRef.current = false
        // Opening the tab later must not drag the user off the deck they are
        // reading just because the last call (already finished) wrote another
        // one; only a call that is still running takes the screen.
        if (first && !running && target !== '') return
        // A brand-new canvas is being created and has no deck yet: whatever is
        // on screen is the previous build, so show this call's progress instead
        // of letting the old pptx pass for the new one.
        if (running && task.tool === 'pptx_studio_open' && target === '') {
          if (deckId === '' && deck === null) return
          awaitingDeckRef.current = true
          setDeck(null)
          setDeckId('')
          setSlideIndex(0)
          setFollow(true)
          setSelection(null)
          setMultiSelection(null)
          setEditing(null)
          setView('auto')
          return
        }
        if (target === '') return
        awaitingDeckRef.current = false
        if (target === deckId) return
        setDeckId(target)
        setSlideIndex(0)
        setFollow(true)
        setSelection(null)
        setMultiSelection(null)
        setEditing(null)
        setView('auto')
        // The new deck may be too fresh for the picker's snapshot.
        void loadDecks()
      }, [task && task.deckId, task && task.seq, task && task.state, task && task.tool, deckId, deck, loadDecks])

      // While no deck is on screen (the canvas was opened before the agent's
      // pptx_studio_open landed), look for one for a few seconds instead of
      // sitting on the skeleton forever. Stops as soon as a deck shows up.
      useEffect(() => {
        if (deckId) return undefined
        let alive = true
        let tries = 0
        const timer = setInterval(() => {
          if (!alive) return
          tries += 1
          void loadDecks().then((list) => {
            if (!alive) return
            if (list.length > 0 || tries >= 8) clearInterval(timer)
          })
        }, 1200)
        return () => {
          alive = false
          clearInterval(timer)
        }
      }, [deckId, loadDecks])

      // SSE: deck snapshots, slides, element changes, status, saved, feedback.
      useEffect(() => {
        if (!deckId) return undefined
        let source
        try {
          source = new EventSource(api(`events?deck=${encodeURIComponent(deckId)}`))
        } catch {
          setLive(false)
          return undefined
        }
        source.onopen = () => setLive(true)
        source.onerror = () => setLive(false)
        source.onmessage = (message) => {
          let event
          try { event = JSON.parse(message.data) } catch { return }
          setLive(true)
          if (event.type === 'deck') {
            // A snapshot from a different deck (or a fresh blank one) must not
            // inherit a half-finished pencil animation from the previous one.
            const previousId = (deckRef.current && deckRef.current.id) || ''
            if ((event.deck && event.deck.id) !== previousId) stopDrawing()
            setDeck(event.deck)
            return
          }
          if (event.type === 'pipeline') {
            setDeck((prev) => (prev ? { ...prev, pipeline: event.pipeline } : prev))
            return
          }
          if (event.type === 'deck-removed') {
            // Another tab (or this one, before it reloaded) deleted the deck.
            setDecks((prev) => prev.filter((item) => item.id !== event.deckId))
            if (event.deckId === deckId) {
              setDeck(null)
              setDeckId('')
            }
            return
          }
          if (event.type === 'slide') {
            setDeck((prev) => {
              if (!prev) return prev
              const next = { ...prev, slides: [...(prev.slides || [])] }
              const at = clamp(Number(event.index) || 0, 0, next.slides.length)
              // The same event covers "insert", "replace" and "move": if the
              // slide already exists, take it out first (otherwise a mid-list
              // insert would silently overwrite the page that was there).
              const existing = next.slides.findIndex((slide) => slide.id === event.slide.id)
              if (existing >= 0) next.slides.splice(existing, 1)
              next.slides.splice(clamp(at, 0, next.slides.length), 0, event.slide)
              return next
            })
            // Pencil-ghost the page, then settle its elements one by one; the
            // user sees the slide being drawn instead of a silent paste.
            revealSlide(event.slide, clamp(Number(event.index) || 0, 0, 999), Number(event.total) || 0)
            if (followRef.current) setSlideIndex(clamp(Number(event.index) || 0, 0, 999))
            return
          }
          if (event.type === 'element') {
            setDeck((prev) => {
              if (!prev) return prev
              const next = { ...prev, slides: (prev.slides || []).map((slide) => {
                if (slide.id !== event.slideId) return slide
                const elements = [...(slide.elements || [])]
                const targetId = event.elementId || (event.element && event.element.id)
                const at = elements.findIndex((el) => el.id === targetId)
                if (event.removed) return { ...slide, elements: elements.filter((el) => el.id !== targetId) }
                if (at >= 0) elements[at] = event.element
                else if (event.element) elements.push(event.element)
                return { ...slide, elements }
              }) }
              return next
            })
            // A removal is not drawn; only additions/patches get the pencil.
            if (event.removed !== true) revealElement(event.element?.id || event.elementId)
            return
          }
          if (event.type === 'status') {
            setPhase({ phase: event.phase, message: event.message || '', slideIndex: event.slideIndex || 0, total: event.total || 0, begin: event.begin === true })
            // `begin` announces a page that is not written yet — jumping to it
            // now would flash the previous page, so wait for the slide event.
            if (followRef.current && event.slideIndex && event.begin !== true) setSlideIndex(Math.max(0, Number(event.slideIndex) - 1))
            return
          }
          if (event.type === 'saved') {
            setSaved(event)
            setPhase((prev) => ({ ...prev, phase: 'saved', message: event.path || '' }))
            return
          }
          if (event.type === 'feedback') {
            const translate = tRef.current || t
            setFeedbackHint(fmt(translate('note.sent'), {
              slide: (Number(event.slideIndex) || 0) + 1,
              target: event.elementId || translate('select.region'),
            }))
          }
        }
        return () => { try { source.close() } catch { /* ignore */ } }
      }, [deckId, markFresh, revealElement, revealSlide, stopDrawing])

      // Measure the stage so "fit" tracks the sidebar width and height.
      useEffect(() => {
        if (activeView !== 'slides') return undefined
        const element = stageRef.current
        if (!element || typeof ResizeObserver !== 'function') return undefined
        const measure = () => {
          const rect = element.getBoundingClientRect()
          if (rect.width > 40 && rect.height > 40) setFit(fitZoom(slideW, slideH, rect.width, rect.height, 12))
        }
        measure()
        const observer = new ResizeObserver(measure)
        observer.observe(element)
        return () => observer.disconnect()
        // `deckId` is in the deps on purpose: the first render has no deck and
        // therefore no stage to measure, so without it "fit" would stay at its
        // seed zoom forever (the exact bug: a slide drawn at 50% and never fitted).
        // `activeView` matters too: leaving the research/preview panel remounts
        // the stage, and without the dep "fit" would keep the stale zoom.
      }, [slideW, slideH, deck ? 'deck' : 'empty', activeView])

      // Measure the HTML preview viewport so the design canvas can be scaled
      // to the sidebar instead of being clipped/stretched by a fixed iframe.
      useEffect(() => {
        if (activeView !== 'preview') return undefined
        const element = previewRef.current
        if (!element) return undefined
        const measure = () => {
          const rect = element.getBoundingClientRect()
          if (rect.width > 40) setPreviewBoxWidth(rect.width)
        }
        measure()
        if (typeof ResizeObserver !== 'function') return undefined
        const observer = new ResizeObserver(measure)
        observer.observe(element)
        return () => observer.disconnect()
      }, [activeView, deckId, htmlPreview && htmlPreview.updatedAt])

      // ── history ────────────────────────────────────────────────────────────

      /**
       * Push the current deck onto the undo stack.
       * @returns {void}
       */
      const pushUndo = useCallback(() => {
        const snapshot = clone(deckRef.current)
        if (!snapshot) return
        undoRef.current.push(snapshot)
        if (undoRef.current.length > UNDO_LIMIT) undoRef.current.shift()
        redoRef.current = []
        setHistory({ undo: undoRef.current.length, redo: 0 })
      }, [])

      /**
       * Undo the last local edit (also mirrored to the host model).
       * @returns {void}
       */
      const undo = useCallback(() => {
        const snapshot = undoRef.current.pop()
        if (!snapshot) return
        const before = clone(deckRef.current)
        if (before) redoRef.current.push(before)
        // deckRef must move with the state: another commit in the same tick
        // would otherwise diff against the pre-undo model.
        deckRef.current = snapshot
        setDeck(snapshot)
        setHistory({ undo: undoRef.current.length, redo: redoRef.current.length })
        setSelection(null)
        setEditing(null)
        void persistSnapshot(snapshot, before)
      }, [persistSnapshot])

      /**
       * Redo the last undone edit.
       * @returns {void}
       */
      const redo = useCallback(() => {
        const snapshot = redoRef.current.pop()
        if (!snapshot) return
        const before = clone(deckRef.current)
        if (before) undoRef.current.push(before)
        deckRef.current = snapshot
        setDeck(snapshot)
        setHistory({ undo: undoRef.current.length, redo: redoRef.current.length })
        void persistSnapshot(snapshot, before)
      }, [persistSnapshot])

      /**
       * Persist a whole-deck snapshot by diffing it against the previous model
       * and sending the minimal set of ops. Used by undo / redo.
       * @param {any} target - desired deck model.
       * @param {any} before - previous deck model.
       * @returns {Promise<void>} resolved when the host acknowledged.
       */
      async function persistSnapshot(target, before) {
        if (!deckId || !target || !before) return
        const ops = diffDecks(before, target)
        if (ops.length === 0) return
        await postOps(ops)
      }

      /**
       * Send ops to the host and keep the local dirty counter honest.
       * @param {Record<string, any>[]} ops - edit ops.
       * @returns {Promise<any>} host response (or null on failure).
       */
      const postOps = useCallback(async (ops) => {
        if (!deckId || ops.length === 0) return null
        pendingRef.current += 1
        setDirty(pendingRef.current)
        try {
          const result = await request('edit', {
            method: 'POST',
            body: { deckId, sessionId, ops },
          })
          setError('')
          return result
        } catch (err) {
          const message = err instanceof Error ? err.message : String(err)
          const code = /** @type {any} */ (err).code
          setError(code === 'not-found' ? t('edit.noHost') : fmt(t('edit.failed'), { message }))
          return null
        } finally {
          pendingRef.current = Math.max(0, pendingRef.current - 1)
          setDirty(pendingRef.current)
        }
      }, [deckId, sessionId, t])

      /**
       * Apply a batch of ops locally, remember one undo step, and persist them.
       * @param {Record<string, any>[]} ops - ops to apply in order.
       * @param {{undoable?: boolean}} [options] - options.
       * @returns {void}
       */
      const commitMany = useCallback((ops, options = {}) => {
        const currentDeck = deckRef.current
        if (!currentDeck || !Array.isArray(ops) || ops.length === 0) return
        if (options.undoable !== false) pushUndo()
        let next = currentDeck
        for (const op of ops) next = applyLocal(next, op)
        deckRef.current = next
        setDeck(next)
        void postOps(ops)
      }, [postOps, pushUndo])

      /**
       * Apply one op locally, remember it for undo, and persist it.
       * @param {Record<string, any>} op - one op.
       * @param {{undoable?: boolean}} [options] - options.
       * @returns {void}
       */
      const commit = useCallback((op, options = {}) => {
        commitMany([op], options)
      }, [commitMany])

      // ── selection / geometry ───────────────────────────────────────────────

      /**
       * Convert a pointer position into slide inches.
       * @param {number} clientX - pointer client X.
       * @param {number} clientY - pointer client Y.
       * @returns {number[]} `[x, y]` in inches.
       */
      const pointOf = useCallback((clientX, clientY) => {
        const canvas = canvasRef.current
        if (!canvas) return [0, 0]
        const rect = canvas.getBoundingClientRect()
        const x = rect.width > 0 ? ((clientX - rect.left) / rect.width) * slideW : 0
        const y = rect.height > 0 ? ((clientY - rect.top) / rect.height) * slideH : 0
        return [clamp(x, 0, slideW), clamp(y, 0, slideH)]
      }, [slideH, slideW])

      /**
       * Pointer-down on an element: select it and start a move drag.
       * @param {PointerEvent} event - pointer event.
       * @param {any} el - element.
       * @param {number[]} box - element box.
       * @returns {void}
       */
      const onElementPointerDown = useCallback((event, el, box) => {
        if (event.button !== undefined && event.button !== 0) return
        event.stopPropagation()
        if (!current) return
        // Keep keyboard shortcuts alive: arrow keys / Delete need focus here.
        try { rootRef.current?.focus?.({ preventScroll: true }) } catch { /* ignore */ }
        setMultiSelection(null)
        setSelection({ kind: 'element', slideId: current.id, elementId: el.id })
        const [px, py] = pointOf(event.clientX, event.clientY)
        dragRef.current = {
          mode: 'move',
          elementId: el.id,
          slideId: current.id,
          box0: box.slice(),
          start: [px, py],
          moved: false,
          liveBox: null,
          pointerId: event.pointerId,
        }
        try { event.currentTarget.setPointerCapture?.(event.pointerId) } catch { /* ignore */ }
      }, [current, pointOf])

      /**
       * Shared move handler for element drags (attached by the capture element).
       * @param {PointerEvent} event - pointer event.
       * @returns {void}
       */
      const onElementPointerMove = useCallback((event) => {
        const drag = dragRef.current
        if (!drag) return
        const [px, py] = pointOf(event.clientX, event.clientY)
        let dx = px - drag.start[0]
        let dy = py - drag.start[1]
        if (!drag.moved && Math.abs(dx) + Math.abs(dy) < 0.02) return
        drag.moved = true
        if (event.shiftKey) {
          if (Math.abs(dx) > Math.abs(dy)) dy = 0
          else dx = 0
        }
        const [x, y, w, hh] = drag.box0
        const box = [
          round3(clamp(x + dx, -w + MIN_EDGE, slideW - MIN_EDGE)),
          round3(clamp(y + dy, -hh + MIN_EDGE, slideH - MIN_EDGE)),
          w, hh,
        ]
        // Keep the live box on the drag record too: pointerup can fire before
        // React re-renders, so the state value alone would be one move behind.
        drag.liveBox = box
        setPreview({ elementId: drag.elementId, box })
      }, [pointOf, slideH, slideW])

      /**
       * Finish a drag: commit the previewed box as one undoable op.
       * @returns {void}
       */
      const onElementPointerUp = useCallback(() => {
        const drag = dragRef.current
        dragRef.current = null
        setPreview(null)
        const box = (drag && drag.liveBox) || (preview && preview.box) || null
        if (!drag || !drag.moved || !box) return
        commit({ kind: 'element.patch', slideId: drag.slideId, elementId: drag.elementId, patch: { box } })
      }, [commit, preview])

      /**
       * Pointer-down on a resize handle.
       * @param {PointerEvent} event - pointer event.
       * @param {string} handle - one of nw/n/ne/e/se/s/sw/w.
       * @returns {void}
       */
      const onHandlePointerDown = useCallback((event, handle) => {
        const sel = selectionRef.current
        const deckNow = deckRef.current
        if (!sel || sel.kind !== 'element' || !deckNow) return
        const slide = (deckNow.slides || []).find((s) => s.id === sel.slideId)
        const el = slide && (slide.elements || []).find((item) => item.id === sel.elementId)
        if (!el) return
        event.stopPropagation()
        event.preventDefault()
        const [px, py] = pointOf(event.clientX, event.clientY)
        dragRef.current = {
          mode: 'resize',
          handle,
          elementId: el.id,
          slideId: slide.id,
          box0: boxOf(el, slideW, slideH),
          start: [px, py],
          moved: false,
        }
        try { event.currentTarget.setPointerCapture?.(event.pointerId) } catch { /* ignore */ }
      }, [pointOf, slideH, slideW])

      /**
       * Move handler for a resize drag.
       * @param {PointerEvent} event - pointer event.
       * @returns {void}
       */
      const onHandlePointerMove = useCallback((event) => {
        const drag = dragRef.current
        if (!drag || drag.mode !== 'resize') return
        const [px, py] = pointOf(event.clientX, event.clientY)
        const dx = px - drag.start[0]
        const dy = py - drag.start[1]
        let [x, y, w, hh] = drag.box0
        if (drag.handle.includes('w')) { x += dx; w -= dx }
        if (drag.handle.includes('e')) { w += dx }
        if (drag.handle.includes('n')) { y += dy; hh -= dy }
        if (drag.handle.includes('s')) { hh += dy }
        if (w < MIN_EDGE) { if (drag.handle.includes('w')) x -= MIN_EDGE - w; w = MIN_EDGE }
        if (hh < MIN_EDGE) { if (drag.handle.includes('n')) y -= MIN_EDGE - hh; hh = MIN_EDGE }
        const box = [round3(x), round3(y), round3(w), round3(hh)]
        drag.moved = true
        drag.liveBox = box
        setPreview({ elementId: drag.elementId, box })
      }, [pointOf])

      /**
       * Pointer-down on empty canvas: start a marquee selection.
       * @param {PointerEvent} event - pointer event.
       * @returns {void}
       */
      const onCanvasPointerDown = useCallback((event) => {
        if (event.button !== undefined && event.button !== 0) return
        if (event.target !== event.currentTarget) return
        const [px, py] = pointOf(event.clientX, event.clientY)
        setSelection(null)
        setMultiSelection(null)
        setEditing(null)
        dragRef.current = { mode: 'marquee', start: [px, py] }
        setMarquee({ x1: px, y1: py, x2: px, y2: py })
        try { event.currentTarget.setPointerCapture?.(event.pointerId) } catch { /* ignore */ }
      }, [pointOf])

      /**
       * Marquee drag.
       * @param {PointerEvent} event - pointer event.
       * @returns {void}
       */
      const onCanvasPointerMove = useCallback((event) => {
        const drag = dragRef.current
        if (!drag || drag.mode !== 'marquee') return
        const [px, py] = pointOf(event.clientX, event.clientY)
        drag.rect = { x1: drag.start[0], y1: drag.start[1], x2: px, y2: py }
        setMarquee(drag.rect)
      }, [pointOf])

      /**
       * Finish a marquee: a big enough region becomes the feedback target.
       * @returns {void}
       */
      const onCanvasPointerUp = useCallback(() => {
        const drag = dragRef.current
        dragRef.current = null
        if (!drag || drag.mode !== 'marquee') return
        const area = drag.rect || marquee
        setMarquee(null)
        if (!area || !current) return
        const rect = [
          round3(Math.min(area.x1, area.x2)),
          round3(Math.min(area.y1, area.y2)),
          round3(Math.abs(area.x2 - area.x1)),
          round3(Math.abs(area.y2 - area.y1)),
        ]
        if (rect[2] > 0.12 && rect[3] > 0.12) setSelection({ kind: 'region', slideId: current.id, rect })
      }, [current, marquee])

      // ── editing helpers ────────────────────────────────────────────────────

      /**
       * Selected element of the current slide.
       * @returns {any|null} element.
       */
      const selectedElement = (() => {
        if (!selection || selection.kind !== 'element' || !current) return null
        const at = elementIndexById(current, selection.elementId)
        return at >= 0 ? current.elements[at] : null
      })()

      /** Element ids currently targeted by copy/cut/delete/arrow operations. */
      const selectedIds = (() => {
        if (!current) return []
        if (Array.isArray(multiSelection) && multiSelection.length > 0) {
          const live = new Set((current.elements || []).map((el) => el.id))
          return multiSelection.filter((id) => live.has(id))
        }
        return selectedElement ? [selectedElement.id] : []
      })()

      /**
       * Patch the selected element (or any element) on the current slide.
       * @param {Record<string, any>} patch - element fields.
       * @param {{undoable?: boolean}} [options] - options.
       * @returns {void}
       */
      const patchElement = useCallback((patch, options = {}) => {
        const sel = selectionRef.current
        if (!sel || sel.kind !== 'element' || !current) return
        commit({ kind: 'element.patch', slideId: current.id, elementId: sel.elementId, patch }, options)
      }, [commit, current])

      /**
       * Start the shared text draft for the selected element (inspector input).
       * @param {any} el - element.
       * @returns {void}
       */
      const beginInspectorEdit = useCallback((el) => {
        if (!el) return
        setEditing({ elementId: el.id, value: typeof el.text === 'string' ? el.text : '', mode: 'inspector' })
      }, [])

      /**
       * Inline text editing: open the overlay for one element.
       * @param {any} el - element.
       * @returns {void}
       */
      const beginTextEdit = useCallback((el) => {
        if (!el) return
        if (el.type !== 'text' && !(el.type === 'rect' && el.text)) return
        setEditing({ elementId: el.id, value: typeof el.text === 'string' ? el.text : '', mode: 'inline' })
      }, [])

      /**
       * Commit the inline text edit.
       * @returns {void}
       */
      const commitTextEdit = useCallback(() => {
        const edit = editing
        if (!edit) return
        setEditing(null)
        const slide = slideRef.current
        if (!slide) return
        const target = (slide.elements || []).find((el) => el.id === edit.elementId)
        if (target && String(target.text ?? '') === edit.value) return
        commit({ kind: 'element.patch', slideId: slide.id, elementId: edit.elementId, patch: { text: edit.value } })
      }, [commit, editing])

      /**
       * Insert a text box in the middle of the current slide.
       * @returns {void}
       */
      const addTextBox = useCallback(() => {
        if (!current) return
        const w = Math.min(6, slideW - 1)
        const element = {
          id: uniqueId(current.elements || [], `${current.id}-t${(current.elements || []).length + 1}`),
          type: 'text',
          box: [round3((slideW - w) / 2), round3(slideH / 2 - 0.4), round3(w), 0.8],
          text: '双击编辑文字',
          style: { size: 24, color: '#0F172A', align: 'left' },
        }
        commit({ kind: 'element.add', slideId: current.id, element })
        setMultiSelection(null)
        setSelection({ kind: 'element', slideId: current.id, elementId: element.id })
      }, [commit, current, slideH, slideW])

      /**
       * Delete the selected element.
       * @returns {void}
       */
      const deleteSelected = useCallback(() => {
        if (!current || selectedIds.length === 0) return
        const ops = selectedIds.map((id) => ({ kind: 'element.patch', slideId: current.id, elementId: id, remove: true }))
        commitMany(ops)
        setSelection(null)
        setMultiSelection(null)
        setEditing(null)
      }, [commitMany, current, selectedIds])

      /**
       * Copy the current selection into the internal and system clipboards.
       * @returns {number} number of elements copied.
       */
      const copySelection = useCallback(() => {
        if (!current || selectedIds.length === 0) return 0
        const elements = selectedIds
          .map((id) => (current.elements || []).find((el) => el.id === id))
          .filter(Boolean)
          .map((el) => clone(el))
        if (elements.length === 0) return 0
        clipboardRef.current = { elements, sourceSlideId: current.id }
        try {
          if (typeof navigator !== 'undefined' && navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
            const writing = navigator.clipboard.writeText(serializeElements(elements))
            if (writing && typeof writing.catch === 'function') writing.catch(() => {})
          }
        } catch { /* system clipboard is a bonus; internal clipboard still works */ }
        setNotice(fmt(t('clipboard.copied'), { n: elements.length }))
        return elements.length
      }, [current, selectedIds, t])

      /**
       * Paste element models onto the current slide.
       * @param {any[]} sourceElements - elements to paste.
       * @returns {number} number of elements pasted.
       */
      const pasteElements = useCallback((sourceElements) => {
        if (!current || !Array.isArray(sourceElements) || sourceElements.length === 0) return 0
        /** @type {any[]} */
        const created = []
        /** @type {Record<string, any>[]} */
        const ops = []
        for (const [index, source] of sourceElements.slice(0, 100).entries()) {
          if (source === null || typeof source !== 'object') continue
          const copy = clone(source)
          copy.id = uniqueId([...(current.elements || []), ...created], typeof source.id === 'string' && source.id !== '' ? `${source.id}-copy` : `${current.id}-paste`)
          copy.box = shiftBox(copy.box, 0.2 + index * 0.1, 0.2 + index * 0.1, slideW, slideH)
          created.push(copy)
          ops.push({ kind: 'element.add', slideId: current.id, element: copy })
        }
        if (ops.length === 0) return 0
        commitMany(ops)
        if (created.length === 1) {
          setMultiSelection(null)
          setSelection({ kind: 'element', slideId: current.id, elementId: created[0].id })
        } else {
          setSelection(null)
          setMultiSelection(created.map((el) => el.id))
        }
        setEditing(null)
        setNotice(fmt(t('clipboard.pasted'), { n: created.length }))
        return created.length
      }, [commitMany, current, slideH, slideW, t])

      /**
       * Paste the internal clipboard, if one exists.
       * @returns {number} number of elements pasted.
       */
      const pasteClipboard = useCallback(() => {
        const payload = clipboardRef.current
        return payload && Array.isArray(payload.elements) ? pasteElements(payload.elements) : 0
      }, [pasteElements])

      /** Cut the selection: copy it, then remove the source elements. */
      const cutSelection = useCallback(() => {
        const count = copySelection()
        if (count > 0) {
          deleteSelected()
          setNotice(fmt(t('clipboard.cut'), { n: count }))
        }
      }, [copySelection, deleteSelected, t])

      /**
       * Duplicate the selected element(s) and select the copies.
       * @returns {void}
       */
      const duplicateSelected = useCallback(() => {
        if (!current || selectedIds.length === 0) return
        /** @type {any[]} */
        const created = []
        /** @type {Record<string, any>[]} */
        const ops = []
        for (const id of selectedIds) {
          const source = (current.elements || []).find((el) => el.id === id)
          if (!source) continue
          const copy = clone(source)
          copy.id = uniqueId([...(current.elements || []), ...created], `${source.id}-copy`)
          copy.box = shiftBox(copy.box, 0.2, 0.2, slideW, slideH)
          created.push(copy)
          ops.push({ kind: 'element.add', slideId: current.id, element: copy })
        }
        if (ops.length === 0) return
        commitMany(ops)
        if (created.length === 1) {
          setMultiSelection(null)
          setSelection({ kind: 'element', slideId: current.id, elementId: created[0].id })
        } else {
          setSelection(null)
          setMultiSelection(created.map((el) => el.id))
        }
      }, [commitMany, current, selectedIds, slideH, slideW])

      /** Select every element on the current slide. */
      const selectAll = useCallback(() => {
        if (!current) return
        const ids = (current.elements || []).map((el) => el.id)
        setSelection(null)
        setMultiSelection(ids.length > 0 ? ids : null)
        setEditing(null)
      }, [current])

      /**
       * Reorder the selected element ("front" | "back").
       * @param {'front'|'back'} to - target z position.
       * @returns {void}
       */
      const reorderSelected = useCallback((to) => {
        if (!selectedElement || !current) return
        commit({ kind: 'element.reorder', slideId: current.id, elementId: selectedElement.id, to })
      }, [commit, current, selectedElement])

      /**
       * Add a slide after the current one.
       * @returns {void}
       */
      const addSlide = useCallback(() => {
        if (!deck) return
        const nextIndex = slides.length
        // The id travels with the op so the host and this model cannot drift.
        const id = uniqueId(slides, `s${nextIndex + 1}`)
        commit({ kind: 'slide.add', index: nextIndex, slide: { id, title: `第 ${nextIndex + 1} 页`, elements: [] } })
        setSlideIndex(nextIndex)
        setSelection(null)
        setMultiSelection(null)
      }, [commit, deck, slides, slides.length])

      /**
       * Delete the current slide.
       * @returns {void}
       */
      const removeSlide = useCallback(() => {
        if (!current) return
        if (slides.length <= 1) {
          setNotice(t('edit.lastSlide'))
          return
        }
        commit({ kind: 'slide.remove', slideId: current.id })
        setSlideIndex(clamp(safeIndex - 1, 0, slides.length - 2))
        setSelection(null)
        setMultiSelection(null)
      }, [commit, current, safeIndex, slides.length, t])

      /**
       * Move the current slide by one position.
       * @param {number} delta - -1 or 1.
       * @returns {void}
       */
      const moveSlide = useCallback((delta) => {
        if (!current) return
        const to = clamp(safeIndex + delta, 0, slides.length - 1)
        if (to === safeIndex) return
        commit({ kind: 'slide.move', slideId: current.id, index: to })
        setSlideIndex(to)
      }, [commit, current, safeIndex, slides.length])

      /**
       * Commit the deck title typed in the toolbar.
       * @returns {void}
       */
      const commitTitle = useCallback(() => {
        setTitleEditing(false)
        const nextTitle = String(titleDraft || '').trim()
        if (!deck || nextTitle === '' || nextTitle === deck.title) return
        commit({ kind: 'deck.title', title: nextTitle })
      }, [commit, deck, titleDraft])

      // ── pipeline / slide effects ───────────────────────────────────────────

      /**
       * Play one transition or element animation on the canvas.
       * @param {{kind: 'transition'|'element', type?: string, direction?: string, elementId?: string, animate?: boolean}} payload - effect descriptor.
       * @returns {void}
       */
      const playFx = useCallback((payload) => {
        if (fxTimer.current) clearTimeout(fxTimer.current)
        setFx({ ...payload, nonce: Date.now() })
        fxTimer.current = setTimeout(() => {
          fxTimer.current = null
          setFx(null)
        }, 1400)
      }, [])

      /**
       * Patch slide-level fields (transition / notes / background / hidden / layout).
       * @param {Record<string, unknown>} patch - fields to merge.
       * @returns {void}
       */
      const patchSlide = useCallback((patch) => {
        if (!current) return
        commit({ kind: 'slide.patch', slideId: current.id, patch })
      }, [commit, current])

      /**
       * Merge one field into the current slide transition.
       * @param {string} key - transition field.
       * @param {unknown} value - field value.
       * @returns {void}
       */
      const setTransitionField = useCallback((key, value) => {
        if (!current) return
        const base = current.transition || {
          type: 'none', speed: 'medium', direction: 'left', advanceOnClick: true, advanceAfterMs: 0,
        }
        const next = { ...base, [key]: value }
        patchSlide({ transition: next.type === 'none' ? null : next })
      }, [current, patchSlide])

      /**
       * Commit the speaker-notes draft.
       * @returns {void}
       */
      const commitNotes = useCallback(() => {
        if (!current) return
        const next = String(notesDraft || '').slice(0, 4000)
        if (String(current.notes || '') === next) return
        patchSlide({ notes: next })
      }, [current, notesDraft, patchSlide])

      /**
       * Set or clear this slide's solid background.
       * @param {string} color - CSS colour, or '' to clear.
       * @returns {void}
       */
      const setBackground = useCallback((color) => {
        patchSlide({ background: color ? { color } : null })
      }, [patchSlide])

      /**
       * Show/hide the current slide.
       * @param {boolean} hidden - new hidden state.
       * @returns {void}
       */
      const toggleHidden = useCallback((hidden) => {
        patchSlide({ hidden: hidden ? true : null })
      }, [patchSlide])

      /**
       * Pick the template layout used by this slide.
       * @param {string} value - layout index as a string.
       * @returns {void}
       */
      const commitLayout = useCallback((value) => {
        const index = Number(value)
        patchSlide({ layoutIndex: Number.isFinite(index) ? index : null })
      }, [patchSlide])

      /**
       * Patch the deck theme (colours / fonts).
       * @param {{colors?: Record<string, string>, fonts?: Record<string, string>}} patch - theme fields.
       * @returns {void}
       */
      const commitTheme = useCallback((patch) => {
        const currentTheme = themeOf(deckRef.current)
        const next = {
          colors: { ...currentTheme.colors, ...(patch.colors || {}) },
          fonts: { ...currentTheme.fonts, ...(patch.fonts || {}) },
        }
        commit({ kind: 'deck.theme', theme: next })
      }, [commit])

      /**
       * Switch the deck to one of the size presets.
       * @param {string} presetId - preset id.
       * @returns {void}
       */
      const commitSize = useCallback((presetId) => {
        const preset = SIZE_PRESETS.find((item) => item.id === presetId)
        if (!preset) return
        commit({ kind: 'deck.size', width: preset.w, height: preset.h })
      }, [commit])

      /**
       * Hand one Host route to the browser's own download channel.
       *
       * `window.open` cannot do this job: the Desktop shell answers every popup
       * with `deny` and only forwards http(s) to an external browser, which then
       * lacks the Host's signed cookie. So the page fetches the bytes itself and
       * lets a blob URL carry the file name.
       * @param {string} url - same-origin route to fetch.
       * @param {string} name - file name the browser should save as.
       * @returns {Promise<boolean>} whether the download was handed over.
       */
      const saveRouteAsFile = useCallback(async (url, name) => {
        if (!url) return false
        try {
          const response = await fetch(url, { credentials: 'same-origin' })
          if (!response.ok) {
            let detail = `HTTP ${response.status}`
            try {
              const data = JSON.parse(await response.text())
              detail = data?.message || data?.code || detail
            } catch { /* keep the status line */ }
            throw new Error(detail)
          }
          const blob = await response.blob()
          const doc = (rootRef.current && rootRef.current.ownerDocument) || (typeof document !== 'undefined' ? document : null)
          const view = (doc && doc.defaultView) || (typeof window !== 'undefined' ? window : null)
          if (!doc || !view || typeof doc.createElement !== 'function' || !view.URL || typeof view.URL.createObjectURL !== 'function') {
            throw new Error(t('actions.noChannel'))
          }
          const href = view.URL.createObjectURL(blob)
          const anchor = doc.createElement('a')
          anchor.href = href
          anchor.download = name
          anchor.rel = 'noopener'
          if (anchor.style) anchor.style.display = 'none'
          const body = doc.body
          if (body && typeof body.appendChild === 'function') body.appendChild(anchor)
          anchor.click()
          try {
            if (anchor.parentNode && typeof anchor.remove === 'function') anchor.remove()
            else if (body && typeof body.removeChild === 'function') body.removeChild(anchor)
          } catch { /* the click already happened */ }
          if (typeof view.setTimeout === 'function') view.setTimeout(() => view.URL.revokeObjectURL(href), 30000)
          setNotice(fmt(t('actions.handedOver'), { name }))
          return true
        } catch (err) {
          setError(err instanceof Error ? err.message : String(err))
          return false
        }
      }, [rootRef, t])

      /**
       * Open one URL in a new window/tab through the document that owns this
       * canvas (the harness may not expose `window.open`).
       * @param {string} url - absolute URL.
       * @returns {void}
       */
      const openUrl = useCallback((url) => {
        if (!url) return
        try {
          const view = rootRef.current && rootRef.current.ownerDocument && rootRef.current.ownerDocument.defaultView
          if (view && typeof view.open === 'function') {
            view.open(url, '_blank', 'noopener')
            return
          }
        } catch { /* fall through */ }
        try {
          if (typeof window !== 'undefined' && typeof window.open === 'function') window.open(url, '_blank', 'noopener')
        } catch { /* ignore */ }
      }, [])

      /** Open the HTML preview in a new window. */
      const openPreview = useCallback(() => openUrl(previewUrl), [openUrl, previewUrl])

      /**
       * Save the HTML mock-up through the browser's download channel.
       * @returns {void}
       */
      const downloadPreview = useCallback(() => {
        if (!previewUrl) return
        void saveRouteAsFile(`${previewUrl}&download=1`, `${String(deck?.title ?? 'preview')}.html`)
      }, [deck, previewUrl, saveRouteAsFile])

      /**
       * Compose a stage-level feedback message (research or HTML preview).
       * @returns {string} message.
       */
      const buildStageFeedbackText = useCallback(() => {
        const label = activeView === 'research' ? '调研分析' : 'HTML 效果预览'
        const lines = ['[PPTX 画布反馈]', `deck=${deckId}`, `阶段=${label}`]
        if (stageNote.trim() !== '') lines.push(`用户意见：${stageNote.trim()}`)
        lines.push(activeView === 'research'
          ? '请更新需求分析/调研（pptx_studio_research），再继续 HTML 效果与 PPTX。'
          : '请先更新 HTML 效果预览（pptx_studio_preview），再把同一套设计落到幻灯片（pptx_studio_slide）。')
        return lines.join('\n')
      }, [activeView, deckId, stageNote])

      /**
       * Send a stage-level note through the same composer/injection path as
       * element feedback.
       * @returns {Promise<void>} resolved when the attempt finished.
       */
      const sendStageFeedback = useCallback(async () => {
        if (stageNote.trim() === '') {
          setFeedbackHint(t('stage.needText'))
          return
        }
        if (stageSending) return
        const text = buildStageFeedbackText()
        setStageSending(true)
        setFeedbackHint('')
        let native = false
        try {
          const actions = props.inputActions
          const composer = document.querySelector('[data-composer-input]')
          if (actions && typeof actions.setDraft === 'function' && typeof actions.submit === 'function' && composer) {
            actions.setDraft(text)
            if (typeof actions.persistDraft === 'function') actions.persistDraft()
            actions.submit()
            native = true
          }
        } catch {
          native = false
        }
        try {
          const result = await request('feedback', {
            method: 'POST',
            body: {
              deckId,
              sessionId,
              note: stageNote.trim(),
              stage: activeView === 'research' ? 'research' : 'preview',
              inject: !native,
            },
          })
          setStageNote('')
          setFeedbackHint(native ? t('note.fallback') : result?.delivered === true ? t('stage.sent') : t('note.copy'))
        } catch (err) {
          setFeedbackHint(`${t('state.error')}: ${err instanceof Error ? err.message : String(err)}`)
        } finally {
          setStageSending(false)
        }
      }, [activeView, buildStageFeedbackText, deckId, props.inputActions, sessionId, stageNote, stageSending, t])

      // ── extra element kinds ───────────────────────────────────────────────

      /**
       * Insert a rectangle in the middle of the current slide.
       * @returns {void}
       */
      const addRect = useCallback(() => {
        if (!current) return
        const w = Math.min(4, slideW - 1)
        const h = Math.min(2, slideH - 1)
        const element = {
          id: uniqueId(current.elements || [], `${current.id}-r${(current.elements || []).length + 1}`),
          type: 'rect',
          box: [round3((slideW - w) / 2), round3((slideH - h) / 2), round3(w), round3(h)],
          fill: theme.colors.surface,
          radius: 0.08,
        }
        commit({ kind: 'element.add', slideId: current.id, element })
        setMultiSelection(null)
        setSelection({ kind: 'element', slideId: current.id, elementId: element.id })
      }, [commit, current, slideH, slideW, theme])

      /**
       * Insert a small table in the middle of the current slide.
       * @returns {void}
       */
      const addTable = useCallback(() => {
        if (!current) return
        const w = Math.min(7, slideW - 1)
        const h = Math.min(3, slideH - 1)
        const element = {
          id: uniqueId(current.elements || [], `${current.id}-tb${(current.elements || []).length + 1}`),
          type: 'table',
          box: [round3((slideW - w) / 2), round3((slideH - h) / 2), round3(w), round3(h)],
          table: { rows: [['列 1', '列 2', '列 3'], ['', '', ''], ['', '', '']], header: true },
          style: { size: 12, color: theme.colors.ink },
        }
        commit({ kind: 'element.add', slideId: current.id, element })
        setMultiSelection(null)
        setSelection({ kind: 'element', slideId: current.id, elementId: element.id })
      }, [commit, current, slideH, slideW, theme])

      /**
       * Insert a native chart in the middle of the current slide.
       * @returns {void}
       */
      const addChart = useCallback(() => {
        if (!current) return
        const w = Math.min(7, slideW - 1)
        const h = Math.min(4, slideH - 1)
        const element = {
          id: uniqueId(current.elements || [], `${current.id}-c${(current.elements || []).length + 1}`),
          type: 'chart',
          box: [round3((slideW - w) / 2), round3((slideH - h) / 2), round3(w), round3(h)],
          chart: {
            kind: 'column',
            title: '数据概览',
            categories: ['Q1', 'Q2', 'Q3', 'Q4'],
            series: [{ name: '系列 1', values: [12, 18, 9, 22], color: theme.colors.accent }],
          },
        }
        commit({ kind: 'element.add', slideId: current.id, element })
        setMultiSelection(null)
        setSelection({ kind: 'element', slideId: current.id, elementId: element.id })
      }, [commit, current, slideH, slideW, theme])

      // ── images ─────────────────────────────────────────────────────────────

      /**
       * Upload an image (data URL) and place it on the current slide.
       * @param {string} dataUrl - `data:image/...;base64,...`.
       * @param {{name?: string, mime?: string, at?: number[]|null}} [options] - placement options.
       * @returns {Promise<void>} resolved when placed.
       */
      const insertImage = useCallback(async (dataUrl, options = {}) => {
        if (!current || !deckId) return
        const match = /^data:(image\/[a-z+.-]+);base64,/i.exec(dataUrl)
        if (!match) {
          setError(fmt(t('image.failed'), { message: 'unsupported data URL' }))
          return
        }
        setBusy('upload')
        setError('')
        try {
          const natural = await imageSize(dataUrl)
          const width = Math.min(slideW * 0.45, 5)
          const height = natural[0] > 0 ? (natural[1] / natural[0]) * width : width * 0.6
          const at = Array.isArray(options.at) ? options.at : [slideW / 2, slideH / 2]
          const box = [
            round3(clamp(at[0] - width / 2, 0, Math.max(0, slideW - width))),
            round3(clamp(at[1] - height / 2, 0, Math.max(0, slideH - height))),
            round3(width),
            round3(height),
          ]
          const result = await request('upload', {
            method: 'POST',
            body: { deckId, slideId: current.id, name: options.name || 'image', mime: match[1], dataUrl, box },
            timeoutMs: 60000,
          })
          if (result && result.ok && result.element) {
            setDeck((prev) => (prev ? applyLocal(prev, { kind: 'element.add', slideId: current.id, element: result.element }) : prev))
            setMultiSelection(null)
            setSelection({ kind: 'element', slideId: current.id, elementId: result.element.id })
            markFresh([result.element.id], true)
          }
        } catch (err) {
          const message = err instanceof Error ? err.message : String(err)
          // `request()` attaches the host's machine code; the human message is
          // localised Chinese, so matching on text would never fire.
          const code = /** @type {any} */ (err).code
          setError(code === 'too-large' ? t('image.tooBig') : fmt(t('image.failed'), { message }))
        } finally {
          setBusy('')
        }
      }, [current, deckId, markFresh, slideH, slideW, t])

      /**
       * Read a dropped/selected File as a data URL and insert it.
       * @param {File} file - image file.
       * @param {number[]|null} [at] - slide point in inches.
       * @returns {Promise<void>} resolved when inserted.
       */
      const insertImageFile = useCallback(async (file, at = null) => {
        if (!file) return
        if (!/^image\//.test(file.type || '')) {
          setError(fmt(t('image.failed'), { message: file.type || 'not an image' }))
          return
        }
        if (file.size > 20 * 1024 * 1024) {
          setError(t('image.tooBig'))
          return
        }
        const dataUrl = await fileToDataUrl(file)
        await insertImage(dataUrl, { name: file.name, mime: file.type, at })
      }, [insertImage, t])

      // Clipboard paste + drag&drop wiring (best effort: both need a real DOM).
      useEffect(() => {
        if (typeof window === 'undefined' || typeof window.addEventListener !== 'function') return undefined
        /**
         * Paste handler: only intercept when the clipboard carries an image.
         * @param {ClipboardEvent} event - clipboard event.
         * @returns {void}
         */
        const onPaste = (event) => {
          if (activeView !== 'slides') return
          const target = event.target || {}
          const tag = String(target.tagName || '').toLowerCase()
          if (tag === 'input' || tag === 'textarea' || tag === 'select' || target.isContentEditable === true) return
          const clipboard = event.clipboardData
          const text = clipboard && typeof clipboard.getData === 'function' ? clipboard.getData('text/plain') : ''
          const elements = parseElementClipboard(text)
          if (elements && elements.length > 0) {
            event.preventDefault()
            pasteElements(elements)
            return
          }
          const items = (clipboard && clipboard.items) || []
          for (const item of items) {
            if (item && item.type && item.type.startsWith('image/')) {
              const file = item.getAsFile()
              if (file) {
                event.preventDefault()
                void insertImageFile(file, null)
                return
              }
            }
          }
        }
        window.addEventListener('paste', onPaste)
        return () => window.removeEventListener('paste', onPaste)
      }, [activeView, insertImageFile, pasteElements])

      // ── feedback ───────────────────────────────────────────────────────────

      /**
       * Compose the structured note that goes to the session.
       * @returns {string} message text.
       */
      const buildFeedbackText = useCallback(() => {
        const slideNo = safeIndex + 1
        const total = slides.length || 1
        const lines = ['[PPTX 画布反馈]', `deck=${deckId} slide=${slideNo}/${total}${current?.id ? ` (${current.id})` : ''}`]
        if (selection?.kind === 'element' && selectedElement) {
          const el = selectedElement
          lines.push(`element=${el.id} type=${el.type}${el.text ? ` text=「${String(el.text).split('\n')[0].slice(0, 80)}」` : ''}`)
          if (Array.isArray(el.box)) lines.push(`位置：x=${el.box[0]}in y=${el.box[1]}in w=${el.box[2]}in h=${el.box[3]}in`)
          if (el.style) lines.push(`样式：${JSON.stringify(el.style).slice(0, 200)}`)
        } else if (selection?.kind === 'region') {
          lines.push(`region=[${selection.rect.join(', ')}] (x,y,w,h in inches)`)
        }
        if (note.trim()) lines.push(`用户意见：${note.trim()}`)
        lines.push('请只改这一处（除非同一问题在别页重复），改完调用 pptx_studio_export 重新导出；不要重做整个 deck。')
        return lines.join('\n')
      }, [current, deckId, note, safeIndex, selectedElement, selection, slides.length])

      /**
       * Send the note to the session: composer first, host injection as the
       * fallback. Never claims success when nothing was delivered.
       * @returns {Promise<void>} resolved when the attempt finished.
       */
      const sendFeedback = useCallback(async () => {
        if (!selection) {
          setFeedbackHint(t('note.needTarget'))
          return
        }
        if (note.trim() === '') {
          setFeedbackHint(t('note.needText'))
          return
        }
        if (sending) return
        const text = buildFeedbackText()
        setSending(true)
        setFeedbackHint('')
        let native = false
        try {
          const actions = props.inputActions
          const composer = document.querySelector('[data-composer-input]')
          if (actions && typeof actions.setDraft === 'function' && typeof actions.submit === 'function' && composer) {
            actions.setDraft(text)
            if (typeof actions.persistDraft === 'function') actions.persistDraft()
            actions.submit()
            native = true
          }
        } catch {
          native = false
        }
        try {
          const result = await request('feedback', {
            method: 'POST',
            body: {
              deckId,
              sessionId,
              slideId: current?.id,
              elementId: selection.kind === 'element' ? selection.elementId : null,
              bbox: selection.kind === 'region' ? selection.rect : (selectedElement?.box ?? null),
              note: note.trim(),
              text,
              inject: !native,
            },
          })
          setNote('')
          if (native) {
            setFeedbackHint(t('note.fallback'))
          } else if (result?.delivered === true) {
            setFeedbackHint(fmt(t('note.sent'), {
              slide: safeIndex + 1,
              target: selection.kind === 'element' ? selection.elementId : t('select.region'),
            }))
          } else {
            setFeedbackHint(t('note.copy'))
          }
        } catch (err) {
          setFeedbackHint(`${t('state.error')}: ${err instanceof Error ? err.message : String(err)}`)
        } finally {
          setSending(false)
        }
      }, [buildFeedbackText, current, deckId, note, props.inputActions, safeIndex, selectedElement, selection, sending, sessionId, t])

      /**
       * Copy the composed note to the clipboard (works without a composer).
       * @returns {Promise<void>} resolved when copied.
       */
      const copyFeedback = useCallback(async () => {
        const text = buildFeedbackText()
        try {
          if (typeof navigator !== 'undefined' && navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
            await navigator.clipboard.writeText(text)
            setFeedbackHint(t('actions.copied'))
            return
          }
        } catch { /* fall through */ }
        setFeedbackHint(t('note.copy'))
      }, [buildFeedbackText, t])

      // ── export / download / close ──────────────────────────────────────────

      /**
       * Export the deck into its own store folder and remember the receipt.
       * @returns {Promise<Record<string, any>|null>} the receipt, or null when it failed.
       */
      const exportDeck = useCallback(async () => {
        if (!deckId || busy) return null
        setBusy('export')
        setError('')
        try {
          const data = await request('export', { method: 'POST', body: { deckId }, timeoutMs: 190000 })
          const value = data?.value ?? null
          if (value) setSaved(value)
          return value
        } catch (err) {
          setError(err instanceof Error ? err.message : String(err))
          return null
        } finally {
          setBusy('')
        }
      }, [busy, deckId])

      /**
       * Export to a folder the user picks with the Host's own directory picker.
       *
       * This is the "download" the toolbar promises: the file leaves the plugin
       * store and lands wherever the user pointed at. Builds without that Client
       * service fall back to the browser's download channel so the file still
       * reaches the disk.
       * @returns {Promise<void>} resolved when the attempt finished.
       */
      const downloadDeck = useCallback(async () => {
        if (!deckId || busy) return
        const uiWorkspace = hostCtx?.get?.('uiWorkspace') ?? null
        const pick = uiWorkspace && typeof uiWorkspace.pickDirectory === 'function'
          ? uiWorkspace.pickDirectory.bind(uiWorkspace)
          : null
        if (!pick) {
          const name = fileNameOf(saved?.path) || `${String(deck?.title ?? 'deck')}.pptx`
          await saveRouteAsFile(api(`download?deck=${encodeURIComponent(deckId)}`), name)
          return
        }
        setBusy('download')
        setError('')
        try {
          const directory = await pick()
          if (typeof directory !== 'string' || directory.trim() === '') {
            setNotice(t('actions.exportCancelled'))
            return
          }
          const data = await request('export', {
            method: 'POST',
            body: { deckId, dir: directory.trim() },
            timeoutMs: 190000,
          })
          const value = data?.value ?? null
          if (value) setSaved(value)
          const produced = String(value?.path ?? '')
          const picked = directory.trim().replace(/[\\/]+$/, '')
          const fold = (text) => text.replace(/\\/g, '/').toLowerCase()
          if (produced !== '' && !fold(produced).startsWith(`${fold(picked)}/`)) {
            // A Host that predates the `dir` parameter writes into the deck
            // folder instead; say so rather than claim the file went where it did not.
            setError(fmt(t('actions.staleHost'), { path: produced }))
            return
          }
          setNotice(fmt(t('actions.exportedAt'), { path: produced || picked }))
        } catch (err) {
          setError(err instanceof Error ? err.message : String(err))
        } finally {
          setBusy('')
        }
      }, [busy, deck, deckId, saveRouteAsFile, saved, t])

      /**
       * Close this canvas tab through whichever handle the sidebar exposes.
       * @returns {void}
       */
      const closeTab = useCallback(() => {
        const tabId = tab?.id
        try {
          if (typeof props.closeTab === 'function' && tabId) {
            props.closeTab(tabId)
            return
          }
        } catch { /* next */ }
        try {
          if (tab && tab.actions && typeof tab.actions.close === 'function') {
            tab.actions.close()
            return
          }
        } catch { /* next */ }
        try {
          hostCtx?.get?.('sidebarRight')?.closeIn?.(sessionId, tabId)
        } catch { /* ignore */ }
      }, [props, sessionId, tab])

      /**
       * Save the deck in place: the .pptx beside its own deck.json.
       * @returns {Promise<void>} resolved when the export attempt finished.
       */
      const saveDeck = useCallback(async () => {
        if (!deckId || busy === 'export' || slides.length === 0) return
        setNotice(t('shortcut.saving'))
        const value = await exportDeck()
        if (value) {
          setNotice(fmt(t('actions.savedAt'), {
            name: fileNameOf(value.path) || 'deck.pptx',
            size: byteSize(value.bytes),
          }))
        }
      }, [busy, deckId, exportDeck, slides.length, t])

      /**
       * Save a copy wherever the user points (Ctrl+Shift+S).
       * @returns {void}
       */
      const saveDeckAs = useCallback(() => {
        if (!deckId || slides.length === 0 || busy) return
        void downloadDeck()
      }, [busy, deckId, downloadDeck, slides.length])

      /**
       * One keyboard entry point for the canvas. Save works even inside an
       * input; native text editing keys (undo/copy/paste/select-all) are left
       * to the focused field instead of being hijacked.
       * @param {KeyboardEvent} event - keydown event.
       * @returns {void}
       */
      const handleShortcut = useCallback((event) => {
        if (!event || typeof event.key !== 'string') return
        const target = event.target || {}
        const tag = String(target.tagName || '').toLowerCase()
        const typing = tag === 'input' || tag === 'textarea' || tag === 'select' || target.isContentEditable === true
        const meta = event.ctrlKey === true || event.metaKey === true
        const key = event.key.toLowerCase()
        const onSlides = activeView === 'slides'
        if (meta && key === 's') {
          event.preventDefault()
          if (event.shiftKey) void saveDeckAs()
          else void saveDeck()
          return
        }
        if (typing) return
        if (meta && key === 'z') {
          event.preventDefault()
          if (event.shiftKey) redo()
          else undo()
          return
        }
        if (meta && key === 'y') {
          event.preventDefault()
          redo()
          return
        }
        if (onSlides && meta && key === 'x') {
          event.preventDefault()
          cutSelection()
          return
        }
        if (onSlides && meta && key === 'c') {
          event.preventDefault()
          copySelection()
          return
        }
        if (onSlides && meta && key === 'v') {
          if (clipboardRef.current && Array.isArray(clipboardRef.current.elements) && clipboardRef.current.elements.length > 0) {
            event.preventDefault()
            pasteClipboard()
          }
          return
        }
        if (onSlides && meta && key === 'd') {
          event.preventDefault()
          duplicateSelected()
          return
        }
        if (onSlides && meta && key === 'm') {
          event.preventDefault()
          addSlide()
          return
        }
        if (onSlides && meta && key === 'a') {
          event.preventDefault()
          selectAll()
          return
        }
        if (onSlides && meta && key === '0') {
          event.preventDefault()
          setZoomMode('fit')
          return
        }
        if (onSlides && meta && (key === '=' || key === '+')) {
          event.preventDefault()
          setZoomMode(clamp(zoom + 0.1, ZOOM_MIN, ZOOM_MAX))
          return
        }
        if (onSlides && meta && key === '-') {
          event.preventDefault()
          setZoomMode(clamp(zoom - 0.1, ZOOM_MIN, ZOOM_MAX))
          return
        }
        if (onSlides && (event.key === 'Delete' || event.key === 'Backspace')) {
          if (selectedIds.length === 0) return
          event.preventDefault()
          deleteSelected()
          return
        }
        if (event.key === 'Escape') {
          setSelection(null)
          setMultiSelection(null)
          setEditing(null)
          setShortcutsOpen(false)
          return
        }
        if (event.key === '?') {
          event.preventDefault()
          setShortcutsOpen((open) => !open)
          return
        }
        if (!onSlides || selectedIds.length === 0) return
        const step = event.shiftKey ? 0.2 : 0.05
        const moves = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] }
        const delta = moves[event.key]
        if (!delta) return
        event.preventDefault()
        const ops = selectedIds.map((id) => {
          const el = (current && current.elements || []).find((item) => item.id === id)
          if (!el) return null
          const box = boxOf(el, slideW, slideH)
          return {
            kind: 'element.patch',
            slideId: current.id,
            elementId: id,
            patch: { box: [round3(box[0] + delta[0]), round3(box[1] + delta[1]), box[2], box[3]] },
          }
        }).filter(Boolean)
        commitMany(ops)
      }, [
        activeView, addSlide, commitMany, copySelection, current, cutSelection, deleteSelected,
        duplicateSelected, pasteClipboard, redo, saveDeck, saveDeckAs, selectedIds, slideH, slideW,
        undo, zoom,
      ])

      // ── derived view data ──────────────────────────────────────────────────

      const selectionLabel = useMemo(() => {
        if (multiSelection && selectedIds.length > 0) return fmt(t('select.multi'), { n: selectedIds.length })
        if (!selection) return ''
        if (selection.kind === 'element') {
          if (!selectedElement) return ''
          const el = selectedElement
          const text = typeof el.text === 'string' ? el.text.split('\n')[0].slice(0, 24) : ''
          return `${el.type} · ${el.id}${text ? ` · ${text}` : ''}`
        }
        return `${t('select.region')} [${selection.rect.map((n) => Math.round(n * 100) / 100).join(', ')}]`
      }, [multiSelection, selectedElement, selectedIds.length, selection, t])

      const statusText = (() => {
        if (error) return error
        const stageLabel = t(stageTitleKey(stage))
        if (draw) return `${stageLabel} · ${fmt(t('status.elements'), { index: draw.slideIndex, done: Math.min(draw.done, draw.count), count: draw.count })}`
        // A running call is the most useful status there is: say which step.
        if (taskRunning) return `${stageLabel} · ${fmt(t('task.running'), { label: toolLabel(t, task.tool) })}`
        if (phase.phase === 'drawing') {
          return phase.begin
            ? `${stageLabel} · ${fmt(t('status.preparing'), { index: phase.slideIndex || safeIndex + 1 })}`
            : `${stageLabel} · ${fmt(t('status.drawing'), { index: phase.slideIndex || safeIndex + 1, total: phase.total || slides.length })}`
        }
        if (phase.phase === 'preparing') return `${stageLabel} · ${fmt(t('status.preparing'), { index: phase.slideIndex || safeIndex + 1 })}`
        if (phase.phase === 'planning') return `${stageLabel} · ${fmt(t('status.planning'), { total: phase.total || slides.length })}`
        if (phase.phase === 'planned') return `${stageLabel} · ${fmt(t('status.planned'), { total: phase.total || slides.length })}`
        if (phase.phase === 'ready') return `${stageLabel} · ${t('status.ready')}`
        if (phase.phase === 'saved') return `${stageLabel} · ${fmt(t('status.saved'), { name: String(phase.message || '').split(/[\\/]/).pop() || '' })}`
        if (saved?.path) return fmt(t('saved.hint'), { path: saved.path })
        if (notice) return notice
        if (activeView === 'template') {
          if (!pipelineTemplate) return `${stageLabel} · ${t('template.emptyTitle')}`
          if (!templateDecomposed) return `${stageLabel} · ${t('template.decomposingTitle')}`
          return `${stageLabel} · ${fmt(t('template.pagesShort'), { n: templatePages })}`
        }
        if (activeView === 'research' && !research) return `${stageLabel} · ${t('stage.nextResearch')}`
        if (activeView === 'preview' && !htmlPreview) return `${stageLabel} · ${t('stage.nextPreview')}`
        if (activeView === 'research' && research) return `${stageLabel} · ${t('stage.nextPreview')}`
        if (activeView === 'preview' && htmlPreview) return `${stageLabel} · ${t('stage.nextBuild')}`
        return `${stageLabel} · ${t('status.idle')}`
      })()

      const marqueeBox = marquee
        ? {
            left: Math.min(marquee.x1, marquee.x2),
            top: Math.min(marquee.y1, marquee.y2),
            width: Math.abs(marquee.x2 - marquee.x1),
            height: Math.abs(marquee.y2 - marquee.y1),
          }
        : null

      // ── empty / loading states ─────────────────────────────────────────────

      // A call opened this tab but its deck does not exist yet: say what the
      // agent is doing instead of showing the previous build's deck.
      const pendingCall = (Boolean(paramPending) || (taskRunning && task.tool === 'pptx_studio_open')) && !(task && task.state === 'failed')
      const taskBanner = task && (task.state === 'running' || task.state === 'failed')
        ? h('div', {
            className: 'ps-task',
            'data-ps-task': task.state,
            'data-ps-tool': task.tool || '',
            'data-ps-deck': task.deckId || '',
            title: task.message || '',
          },
            h('span', { className: 'ps-task-dot', 'aria-hidden': 'true' }),
            h('span', { className: 'ps-task-text' }, fmt(t(task.state === 'running' ? 'task.running' : 'task.failed'), { label: toolLabel(t, task.tool) })),
            task.state === 'failed' && task.message ? h('span', { className: 'ps-task-msg' }, task.message) : null,
            h('span', { style: { flex: 1 } }),
            task.deckId ? h('span', { className: 'ps-meta' }, fmt(t('task.deck'), { id: task.deckId })) : null,
            task.state === 'running' && taskSeconds > 0 ? h('span', { className: 'ps-task-elapsed' }, fmt(t('task.elapsed'), { n: taskSeconds })) : null,
          )
        : null

      // No deck yet (the agent has not called pptx_studio_open, or it is still
      // on its way): keep the very same editor/controls shell on screen with a
      // paper-shaped skeleton. The area the user will edit is reserved from the
      // first frame, so the deck does not "appear out of nowhere" and shove the
      // layout around when the first event lands.
      if (!deck) {
        const waiting = pendingCall ? t('task.waitingTitle') : loading ? t('state.loading') : t('empty.shellTitle')
        const waitingHint = pendingCall
          ? fmt(t('task.waitingHint'), { label: task ? toolLabel(t, task.tool) : t('tool.busy') })
          : loading ? t('status.skeleton') : t('empty.shellBody')
        return h('div', {
          className: 'ps-root',
          'data-ps-root': pendingCall ? 'pending' : loading ? 'loading' : 'waiting',
          'data-ps-pending': String(pendingCall),
          ref: rootRef,
        },
          h('div', { className: 'ps-bar' },
            h('span', { className: 'ps-dot' + (live ? ' live' : '') }),
            h('span', { className: 'ps-title' }, t('tab.title')),
            h('span', { className: 'ps-meta' }, waiting),
            h('span', { style: { flex: 1 } }),
            h('button', {
              className: 'ps-btn ghost',
              type: 'button',
              'data-ps-btn': 'refresh-empty',
              title: t('actions.refresh'),
              onClick: () => { void loadDecks() },
            }, '↻'),
            tab ? h('button', { className: 'ps-btn ghost', type: 'button', 'data-ps-btn': 'close-empty', title: t('actions.close'), onClick: closeTab }, '✕') : null,
          ),
          taskBanner,
          h('div', { className: 'ps-steps', 'data-ps-steps': '1', 'data-pending': '1' },
            STAGE_ORDER.map((key, index) => h('span', {
              key,
              className: 'ps-step',
              'data-ps-step': key,
              'data-active': 'false',
              'data-done': 'false',
              'aria-disabled': 'true',
            },
              h('span', { className: 'n' }, String(index)),
              h('span', null, t(stageStepKey(key)))))),
          h('div', { className: 'ps-editor', 'data-ps-editor': '1', 'data-waiting': '1' },
            h('div', { className: 'ps-editor-head' },
              h('span', { className: 'ps-zone-label' }, t('zone.editor')),
              h('span', { className: 'ps-zone-sub' }, t('zone.waiting')),
              h('span', { style: { flex: 1 } }),
              h('span', { className: 'ps-meta' }, loading ? t('state.loading') : '')),
            h('div', { className: 'ps-thumbs', 'data-ps-thumbs': '1', 'data-empty': 'true' },
              [0, 1, 2].map((slot) => h('span', { key: `ghost${slot}`, className: 'ps-thumb ghost', 'aria-hidden': 'true' }))),
            h('div', { className: 'ps-stage', 'data-ps-stage': '1', 'data-waiting': '1' },
              h('div', { className: 'ps-skeleton', 'data-ps-skeleton': '1' },
                h('div', { className: 'ps-skeleton-paper' },
                  h('span', { className: 'sk-line', style: { width: '46%', height: '7%' } }),
                  h('span', { className: 'sk-line', style: { width: '72%', height: '4%', marginTop: '5%' } }),
                  h('span', { className: 'sk-block', style: { width: '100%', height: '34%', marginTop: '6%' } }),
                  h('span', { className: 'sk-line', style: { width: '58%', height: '4%', marginTop: '5%' } })),
                h('div', { className: 'ps-skeleton-hint' }, waitingHint)))),
          h('div', { className: 'ps-split', 'data-ps-split': '1' },
            h('span', { className: 'ps-split-grip' }, t('zone.split'))),
          h('div', { className: 'ps-controls', 'data-ps-controls': '1' },
            h('div', { className: 'ps-controls-head' },
              h('span', { className: 'ps-zone-label' }, t('zone.controls')),
              h('span', { className: 'ps-zone-sub' }, waiting),
              h('span', { style: { flex: 1 } })),
            h('div', { className: 'ps-foot ps-controls-body' },
              h('div', { className: 'ps-row' },
                h('span', { className: 'ps-status', 'data-ps-status': '1' }, waitingHint)),
              h('div', { className: 'ps-row' },
                h('button', { className: 'ps-btn', type: 'button', 'data-ps-btn': 'empty-refresh', onClick: () => { void loadDecks() } }, t('empty.refresh')),
                h('span', { style: { flex: 1 } }),
                error ? h('span', { className: 'ps-err' }, `${t('state.error')}: ${error}`) : null))))
      }

      // ── thumbnail rail ─────────────────────────────────────────────────────

      const thumbWidth = 88
      const thumbHeight = Math.round((thumbWidth * slideH / slideW) * 100) / 100
      const thumbBackground = (slide) => (slide.background && /^#[0-9A-Fa-f]{6}$/.test(String(slide.background.color || ''))
        ? slide.background.color
        : theme.colors.background)
      const thumb = (slide, index) => h('button', {
        key: slide.id || `s${index}`,
        type: 'button',
        className: 'ps-thumb',
        'data-ps-thumb': String(index),
        'data-active': String(index === safeIndex),
        'data-hidden': String(Boolean(slide.hidden)),
        title: slide.title || `第 ${index + 1} 页`,
        style: { width: `${thumbWidth}px`, height: `${thumbHeight}px` },
        onClick: () => { setSlideIndex(index); setFollow(false); setView('slides'); setSelection(null); setMultiSelection(null); setEditing(null) },
      },
        renderThumb(slide, {
          deckId,
          slideW,
          slideH,
          rev,
          theme,
          approx: t('approx'),
          background: thumbBackground(slide),
          thumbW: thumbWidth,
          designZoom: THUMB_DESIGN_ZOOM,
        }),
        h('span', { className: 'n' }, String(index + 1)),
        Number.isInteger(slide.templateSlide)
          ? h('span', {
              className: 'ps-tpl-badge',
              'data-ps-template-badge': String(slide.templateSlide),
              title: fmt(t('template.fromSlide'), { n: slide.templateSlide }),
            }, `T${slide.templateSlide}`)
          : null,
        slide.hidden ? h('span', { className: 'h', title: t('slide.hidden') }, '⊘') : null,
        slide.transition && slide.transition.type && slide.transition.type !== 'none'
          ? h('span', { className: 'fx', title: slide.transition.type }, '✦')
          : null,
        index === slides.length - 1 ? h('span', { className: 'up' }) : null,
      )

      // ── selection overlay + handles ────────────────────────────────────────

      let overlay = null
      if (multiSelection && selectedIds.length > 0 && current) {
        overlay = h('div', { 'data-ps-multi-frames': '1' }, selectedIds.map((id) => {
          const el = (current.elements || []).find((item) => item.id === id)
          if (!el) return null
          return h('div', {
            key: id,
            className: 'ps-frame ps-frame-multi',
            style: placement(boxOf(el, slideW, slideH), zoom),
          })
        }))
      } else if (selectedElement && current) {
        const base = boxOf(selectedElement, slideW, slideH)
        const box = preview && preview.elementId === selectedElement.id ? preview.box : base
        overlay = h('div', { className: 'ps-frame', style: placement(box, zoom) },
          ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'].map((handle) => h('div', {
            key: handle,
            className: 'ps-handle',
            'data-h': handle,
            onPointerDown: (event) => onHandlePointerDown(event, handle),
            onPointerMove: onHandlePointerMove,
            onPointerUp: onElementPointerUp,
          })))
      }
      if (selection?.kind === 'region' && current) {
        overlay = h('div', {
          className: 'ps-region',
          style: placement(selection.rect, zoom),
        })
      }

      // ── inline text editor ─────────────────────────────────────────────────

      let inlineEditor = null
      if (editing && editing.mode === 'inline' && current) {
        const at = elementIndexById(current, editing.elementId)
        const target = at >= 0 ? current.elements[at] : null
        if (target) {
          const box = boxOf(target, slideW, slideH)
          const style = target.style || {}
          inlineEditor = h('textarea', {
            className: 'ps-inline',
            'data-ps-inline': editing.elementId,
            autoFocus: true,
            value: editing.value,
            style: {
              ...placement(box, zoom),
              fontSize: `${Math.max(6, (Number(style.size) || 18) * pxPerPt(zoom))}px`,
              fontFamily: style.font ? `${style.font}, "Microsoft YaHei", sans-serif` : '"Microsoft YaHei", sans-serif',
              fontWeight: style.bold ? 700 : 400,
              textAlign: style.align || 'left',
              lineHeight: '1.2',
              padding: '0',
            },
            onChange: (event) => setEditing({ elementId: editing.elementId, value: event.target.value, mode: 'inline' }),
            onBlur: commitTextEdit,
            onKeyDown: (event) => {
              if (event.key === 'Escape') {
                event.preventDefault()
                setEditing(null)
              } else if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) {
                event.preventDefault()
                commitTextEdit()
              }
            },
          })
        }
      }

      // ── inspector ──────────────────────────────────────────────────────────

      const colorInput = (value, onPick, key) => h('div', { className: 'ps-swatches', key }, SWATCHES.map((swatch) => h('button', {
        key: swatch,
        type: 'button',
        className: 'ps-swatch',
        'data-active': String(String(value || '').toLowerCase() === swatch.toLowerCase()),
        title: swatch,
        style: { background: swatch },
        onClick: () => onPick(swatch),
      })))

      let inspector = null
      if (selectedElement && current) {
        const style = selectedElement.style || {}
        const box = boxOf(selectedElement, slideW, slideH)
        const setStyle = (patch) => patchElement({ style: patch })
        const setBox = (index, value) => {
          const next = box.slice()
          next[index] = round3(clamp(Number(value) || 0, index < 2 ? -2 : MIN_EDGE, index < 2 ? 40 : 40))
          patchElement({ box: next })
        }
        const anim = selectedElement.anim || { effect: 'none', direction: 'left', duration: 500, delay: 0, order: 0 }
        const setAnim = (patch) => {
          const next = { ...anim, ...patch }
          patchElement({ anim: next.effect === 'none' ? null : next })
        }
        inspector = h('div', { className: 'ps-insp', 'data-ps-inspector': selectedElement.id },
          h('div', { className: 'ps-row' },
            h('span', { className: 'ps-meta' }, `${selectedElement.type} · ${selectedElement.id}`),
            h('span', { style: { flex: 1 } }),
            h('button', { className: 'ps-btn icon', type: 'button', title: t('actions.duplicate'), onClick: duplicateSelected }, '⧉'),
            h('button', { className: 'ps-btn icon', type: 'button', title: t('actions.front'), onClick: () => reorderSelected('front') }, '⤒'),
            h('button', { className: 'ps-btn icon', type: 'button', title: t('actions.back'), onClick: () => reorderSelected('back') }, '⤓'),
            h('button', { className: 'ps-btn icon danger', type: 'button', title: t('actions.delete'), onClick: deleteSelected }, '✕'),
          ),
          (selectedElement.type === 'text' || selectedElement.text !== undefined)
            ? h('textarea', {
                className: 'ps-textarea',
                'data-ps-text': selectedElement.id,
                value: editing && editing.elementId === selectedElement.id
                  ? editing.value
                  : String(selectedElement.text ?? ''),
                placeholder: t('edit.text'),
                onFocus: () => { if (!editing || editing.elementId !== selectedElement.id) beginInspectorEdit(selectedElement) },
                onChange: (event) => setEditing({ elementId: selectedElement.id, value: event.target.value, mode: 'inspector' }),
                onBlur: commitTextEdit,
              })
            : null,
          h('div', { className: 'ps-row' },
            h('div', { className: 'ps-field', style: { width: 64 } },
              h('label', null, t('edit.size')),
              h('input', {
                className: 'ps-input',
                type: 'number',
                min: 5,
                max: 200,
                value: String(Number(style.size) || 18),
                onChange: (event) => setStyle({ size: clamp(Number(event.target.value) || 18, 5, 200) }),
              })),
            h('button', { className: 'ps-btn', type: 'button', 'data-active': String(style.bold === true), title: t('edit.bold'), onClick: () => setStyle({ bold: style.bold !== true }) }, 'B'),
            h('button', { className: 'ps-btn', type: 'button', 'data-active': String(style.italic === true), title: t('edit.italic'), onClick: () => setStyle({ italic: style.italic !== true }) }, 'I'),
            h('button', { className: 'ps-btn', type: 'button', 'data-ps-btn': 'underline', 'data-active': String(style.underline === true), title: t('edit.underline'), onClick: () => setStyle({ underline: style.underline !== true }) }, 'U'),
            h('div', { className: 'ps-seg' },
              ['left', 'center', 'right'].map((align) => h('button', {
                key: align,
                className: 'ps-btn',
                type: 'button',
                'data-active': String((style.align || 'left') === align),
                onClick: () => setStyle({ align }),
              }, align === 'left' ? '⯇' : align === 'center' ? '≡' : '⯈')),
            ),
          ),
          h('div', { className: 'ps-row' },
            h('span', { className: 'ps-meta' }, t('edit.color')),
            colorInput(style.color || '#0F172A', (value) => setStyle({ color: value }), 'color'),
          ),
          (selectedElement.type === 'rect' || selectedElement.type === 'text')
            ? h('div', { className: 'ps-row' },
                h('span', { className: 'ps-meta' }, t('edit.fill')),
                colorInput(selectedElement.fill || '', (value) => patchElement({ fill: value }), 'fill'),
                h('button', { className: 'ps-btn ghost', type: 'button', onClick: () => patchElement({ fill: null }) }, '⌀'),
              )
            : null,
          h('div', { className: 'ps-row' },
            h('span', { className: 'ps-meta' }, t('edit.anim')),
            h('select', {
              className: 'ps-btn ps-select',
              'data-ps-field': 'anim-effect',
              value: anim.effect || 'none',
              onChange: (event) => setAnim({ effect: event.target.value }),
            }, ANIM_EFFECTS.map((effect) => h('option', { key: effect, value: effect }, t(`edit.anim${effect.charAt(0).toUpperCase()}${effect.slice(1)}`)))),
            anim.effect === 'wipe'
              ? h('select', {
                  className: 'ps-btn ps-select',
                  'data-ps-field': 'anim-direction',
                  value: anim.direction || 'left',
                  onChange: (event) => setAnim({ direction: event.target.value }),
                }, ['left', 'right', 'up', 'down'].map((dir) => h('option', { key: dir, value: dir }, t(`dir.${dir}`))))
              : null,
            anim.effect && anim.effect !== 'none'
              ? h('input', {
                  className: 'ps-input',
                  'data-ps-field': 'anim-duration',
                  style: { width: 60 },
                  type: 'number',
                  min: 100,
                  max: 10000,
                  step: 100,
                  title: t('edit.anim'),
                  value: String(anim.duration || 500),
                  onChange: (event) => setAnim({ duration: clamp(Number(event.target.value) || 500, 100, 10000) }),
                })
              : null,
            anim.effect && anim.effect !== 'none'
              ? h('button', { className: 'ps-btn ghost', type: 'button', 'data-ps-btn': 'anim-preview', onClick: () => playFx({ kind: 'element', elementId: selectedElement.id }) }, t('edit.animPlay'))
              : null,
            h('button', {
              className: 'ps-btn',
              type: 'button',
              'data-ps-btn': 'shadow',
              'data-active': String(Boolean(selectedElement.shadow)),
              title: t('edit.shadow'),
              onClick: () => patchElement({ shadow: selectedElement.shadow ? null : true }),
            }, t('edit.shadow')),
          ),
          h('div', { className: 'ps-grid' },
            ['X', 'Y', 'W', 'H'].map((labelText, index) => h('div', { className: 'ps-field', key: labelText },
              h('label', null, `${labelText} (in)`),
              h('input', {
                className: 'ps-input',
                type: 'number',
                step: 0.05,
                value: String(box[index]),
                onChange: (event) => setBox(index, event.target.value),
              }))))
        )
      }

      // ── workflow panels ────────────────────────────────────────────────────

      const previewBytes = htmlPreview && typeof htmlPreview.html === 'string'
        ? (typeof TextEncoder === 'function' ? new TextEncoder().encode(htmlPreview.html).length : htmlPreview.html.length)
        : 0
      const stepSub = {
        template: pipelineTemplate ? (templateDecomposed ? fmt(t('template.pagesShort'), { n: templatePages }) : t('stage.current')) : '',
        research: research && Array.isArray(research.outline) && research.outline.length > 0 ? `${research.outline.length} 页大纲` : '',
        preview: htmlPreview ? fmt(t('preview.size'), { kb: Math.max(1, Math.round(previewBytes / 1024)) }) : '',
        build: `${slides.length} 页`,
      }
      const legacyBuild = stage === 'build' && !research && !htmlPreview && slides.length > 0
      const steps = h('div', { className: 'ps-steps', 'data-ps-steps': '1' },
        STAGE_ORDER.map((key, index) => {
          // ⓪ template: decomposed = done, still required = in progress.
          const done = key === 'template'
            ? templateDecomposed
            : legacyBuild || (key === 'research' ? Boolean(research) : key === 'preview' ? Boolean(htmlPreview) : slides.length > 0)
          const active = key === 'build' ? activeView === 'slides' : activeView === key
          return h('button', {
            key,
            type: 'button',
            className: 'ps-step',
            'data-ps-step': key,
            'data-active': String(active),
            'data-done': String(done),
            title: t(stageTitleKey(key)),
            onClick: () => { setView(key === 'build' ? 'slides' : key); setSelection(null); setMultiSelection(null); setEditing(null) },
          },
            h('span', { className: 'n' }, String(index)),
            h('span', null, t(stageStepKey(key))),
            stepSub[key] ? h('span', { className: 'sub' }, stepSub[key]) : null)
        }))

      const stageNoteBox = (placeholder) => h('div', { className: 'ps-stage-note' },
        h('textarea', {
          className: 'ps-note',
          'data-ps-stage-note': '1',
          value: stageNote,
          placeholder,
          onChange: (event) => setStageNote(event.target.value),
          onKeyDown: (event) => {
            if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) {
              event.preventDefault()
              void sendStageFeedback()
            }
          },
        }),
        h('div', { className: 'ps-row' },
          h('button', {
            className: 'ps-btn primary',
            type: 'button',
            'data-ps-btn': 'stage-send',
            disabled: stageSending,
            onClick: () => { void sendStageFeedback() },
          }, stageSending ? t('stage.sending') : t('stage.send')),
          feedbackHint ? h('span', { className: 'ps-status' }, feedbackHint) : null))

      const researchBody = research
        ? h('div', { className: 'ps-brief' },
            h('div', { className: 'card' },
              h('h5', null, t('research.summary')),
              h('p', null, research.summary || '—')),
            h('div', { className: 'card' },
              h('h5', null, t('research.audience')),
              h('div', { className: 'ps-kv' },
                [['research.audience', research.audience], ['research.goal', research.goal], ['research.tone', research.tone], ['research.constraints', research.constraints]]
                  .map(([key, value]) => h('div', { key }, h('div', { className: 'k' }, t(key)), h('div', { className: 'v' }, value || '—'))))),
            Array.isArray(research.outline) && research.outline.length > 0
              ? h('div', { className: 'card' },
                  h('h5', null, t('research.outline')),
                  h('ol', { className: 'ps-list' }, research.outline.map((item, index) => h('li', { key: `o${index}` },
                    h('div', null, item.title || t('research.untitled')),
                    item.intent ? h('div', { className: 'intent' }, item.intent) : null,
                    item.source ? h('div', { className: 'intent' }, `来源：${item.source}`) : null))))
              : null,
            Array.isArray(research.facts) && research.facts.length > 0
              ? h('div', { className: 'card' },
                  h('h5', null, t('research.facts')),
                  h('ul', { className: 'ps-list' }, research.facts.map((item, index) => h('li', { key: `f${index}` },
                    h('div', null, item.claim),
                    item.source ? h('div', { className: 'intent' }, `来源：${item.source}`) : null))))
              : null,
            Array.isArray(research.sources) && research.sources.length > 0
              ? h('div', { className: 'card' },
                  h('h5', null, t('research.sources')),
                  research.sources.map((item, index) => h('div', { className: 'ps-source', key: `s${index}` },
                    item.url
                      ? h('a', { href: item.url, target: '_blank', rel: 'noreferrer noopener' }, item.title || item.url)
                      : h('span', null, item.title || '—'),
                    item.note ? h('span', { className: 'ps-hint' }, item.note) : null)))
              : null,
            (Array.isArray(research.palette) && research.palette.length > 0) || (research.fonts && (research.fonts.latin || research.fonts.heading))
              ? h('div', { className: 'card' },
                  h('h5', null, t('research.palette')),
                  Array.isArray(research.palette) && research.palette.length > 0
                    ? h('div', { className: 'ps-palette' }, research.palette.map((color) => h('span', { key: color, title: color, style: { background: cssColor(color, '#fff') } })))
                    : null,
                  research.fonts && (research.fonts.latin || research.fonts.heading)
                    ? h('div', { className: 'ps-hint', style: { marginTop: 4 } },
                        `${t('research.fonts')}：${[research.fonts.latin, research.fonts.heading].filter(Boolean).join(' / ')}`)
                    : null)
              : null)
        : h('div', { className: 'ps-empty' },
            h('h4', null, t('research.emptyTitle')),
            h('p', null, t('research.emptyBody')))

      const researchPanel = h('div', { className: 'ps-panel', 'data-ps-panel': 'research' },
        h('div', { className: 'ps-panel-head' },
          h('h4', null, t('stage.research')),
          h('span', { className: 'ps-badge' + (research ? ' ok' : '') }, research ? t('stage.done') : t('stage.pending')),
          h('span', { style: { flex: 1 } }),
          htmlPreview ? h('button', { className: 'ps-btn', type: 'button', 'data-ps-btn': 'goto-preview', onClick: () => setView('preview') }, t('stage.stepPreview')) : null),
        researchBody,
        stageNoteBox(t('research.ask')))

      const previewBody = htmlPreview
        ? h('div', { style: { display: 'flex', flexDirection: 'column', gap: 6, minHeight: 0 } },
            h('div', {
              className: 'ps-preview-wrap',
              'data-ps-preview-wrap': '1',
              'data-ps-preview-scale': String(Math.round(previewScale * 1000) / 1000),
              ref: previewRef,
              style: { height: `${previewDisplayHeight}px` },
            },
              h('div', {
                style: {
                  position: 'absolute',
                  left: '50%',
                  top: 0,
                  width: `${Math.round(previewSize.width * previewScale)}px`,
                  height: `${previewDisplayHeight}px`,
                  transform: 'translateX(-50%)',
                },
              },
                h('iframe', {
                  key: `${htmlPreview.updatedAt || ''}:${previewNonce}`,
                  className: 'ps-preview-frame',
                  'data-ps-preview': '1',
                  src: previewUrl,
                  title: htmlPreview.title || t('stage.preview'),
                  sandbox: '',
                  referrerPolicy: 'no-referrer',
                  style: {
                    width: `${previewSize.width}px`,
                    height: `${previewSize.height}px`,
                    transform: `scale(${previewScale})`,
                    transformOrigin: 'top left',
                  },
                }))),
            h('div', { className: 'ps-hint' }, t('preview.sandboxHint')),
            htmlPreview.notes ? h('div', { className: 'card' }, h('h5', null, t('preview.notes')), h('p', null, htmlPreview.notes)) : null)
        : h('div', { className: 'ps-empty' },
            h('h4', null, t('preview.emptyTitle')),
            h('p', null, t('preview.emptyBody')))

      const previewPanel = h('div', { className: 'ps-panel', 'data-ps-panel': 'preview' },
        h('div', { className: 'ps-panel-head' },
          h('h4', null, t('stage.preview')),
          h('span', { className: 'ps-badge' + (htmlPreview ? ' ok' : '') }, htmlPreview ? t('stage.done') : t('stage.pending')),
          h('span', { style: { flex: 1 } }),
          htmlPreview ? h('button', { className: 'ps-btn ghost', type: 'button', 'data-ps-btn': 'preview-reload', onClick: () => setPreviewNonce((n) => n + 1) }, t('preview.reload')) : null,
          htmlPreview ? h('button', { className: 'ps-btn', type: 'button', 'data-ps-btn': 'preview-open', onClick: openPreview }, t('preview.open')) : null,
          htmlPreview ? h('button', { className: 'ps-btn ghost', type: 'button', 'data-ps-btn': 'preview-download', onClick: downloadPreview }, t('preview.download')) : null),
        previewBody,
        stageNoteBox(t('preview.ask')))

      // ── template panel (⓪ before research) ────────────────────────────────

      /**
       * Template walkthrough: page plates with their role and fillable fields,
       * the template palette, and which pages this deck actually mapped. The
       * data comes from `GET /api/template`, never from guesswork.
       */
      const templatePanel = (() => {
        const info = templateInfo
        const pages = info && Array.isArray(info.pages) ? info.pages : []
        const mapping = info && Array.isArray(info.mapping) ? info.mapping : []
        const summary = (info && info.summary) || {}
        const design = (info && info.design) || {}
        const palette = Array.isArray(design.palette) ? design.palette : []
        const fonts = design.fonts || {}
        const fontUsage = Array.isArray(design.fontUsage) ? design.fontUsage : []
        const roleLabel = (role) => {
          const key = `template.role.${role}`
          const text = t(key)
          return text === key ? String(role || '') : text
        }
        const mappedPage = (pageIndex) => mapping.find((item) => Number(item && item.templateSlide) === Number(pageIndex)) || null
        const summaryBits = [
          summary.engine ? `${t('template.engine')} ${summary.engine}` : '',
          templatePages > 0 ? fmt(t('template.pagesShort'), { n: templatePages }) : '',
          Number(summary.fields) > 0 ? fmt(t('template.fieldsShort'), { n: summary.fields }) : '',
          info && info.size && Number(info.size.w) > 0 ? `${info.size.w} × ${info.size.h} in` : '',
          summary.textHidden ? t('template.textHidden') : '',
        ].filter(Boolean)
        const designBlock = palette.length > 0 || fonts.title || fonts.body || fontUsage.length > 0
          ? h('div', { className: 'ps-brief' },
              h('div', { className: 'card' },
                h('h5', null, t('template.design')),
                palette.length > 0
                  ? h('div', { className: 'ps-palette', 'data-ps-template-palette': '1' },
                      palette.map((color, index) => h('span', {
                        key: `${color.hex}:${index}`,
                        title: `${color.hex} · ${Math.round((Number(color.share) || 0) * 100)}%`,
                        'data-ps-palette-hex': String(color.hex),
                        style: { background: cssColor(color.hex, 'transparent') },
                      })))
                  : null,
                fonts.title || fonts.body
                  ? h('div', { className: 'ps-hint', style: { marginTop: 4 } },
                      `${t('template.fonts')}：${[fonts.title, fonts.body].filter(Boolean).join(' / ')}`)
                  : null,
                fontUsage.length > 0
                  ? h('div', { className: 'ps-hint' }, fontUsage.slice(0, 4).map((item) => `${item.name}×${item.count}`).join(' · '))
                  : null))
          : null
        let body
        if (!pipelineTemplate) {
          body = h('div', { className: 'ps-empty' },
            h('h4', null, t('template.emptyTitle')),
            h('p', null, t('template.emptyBody')))
        } else if (!templateDecomposed) {
          body = h('div', { className: 'ps-empty' },
            h('h4', null, t('template.decomposingTitle')),
            h('p', null, t('template.decomposingBody')),
            pipelineTemplate.path ? h('div', { className: 'ps-hint', 'data-ps-template-path': '1' }, pipelineTemplate.path) : null)
        } else if (!info) {
          body = h('div', { className: 'ps-empty' },
            h('h4', null, t('template.unavailableTitle')),
            h('p', null, templateError ? fmt(t('template.unavailableBody'), { error: templateError }) : t('template.unavailableBody')))
        } else {
          body = h('div', { className: 'ps-tpl' },
            summaryBits.length > 0
              ? h('div', { className: 'ps-tpl-summary', 'data-ps-template-summary': '1' }, summaryBits.map((bit, index) => h('span', { key: `${index}:${bit}`, className: 'ps-hint' }, bit)))
              : null,
            designBlock,
            pages.length > 0
              ? h('div', { className: 'ps-tpl-grid', 'data-ps-template-grid': '1' }, pages.map((page) => {
                  const pageIndex = Number(page.page)
                  const mapped = mappedPage(pageIndex)
                  const filled = mapped && mapped.fields && typeof mapped.fields === 'object' ? Object.keys(mapped.fields) : []
                  return h('div', {
                    key: `tpl${pageIndex}`,
                    className: 'ps-tpl-page' + (mapped ? ' mapped' : ''),
                    'data-ps-template-page': String(pageIndex),
                    'data-ps-template-role': String(page.role || ''),
                    'data-ps-template-mapped': String(Boolean(mapped)),
                  },
                    h('div', { className: 'ps-tpl-thumb' },
                      h('img', {
                        src: assetUrl(deckId, `template/thumbs/s${pageIndex}.jpg`, rev),
                        alt: `${t('stage.template')} T${pageIndex}`,
                      }),
                      h('span', { className: 'ps-tpl-no' }, `T${pageIndex}`),
                      h('span', { className: 'ps-tpl-role' }, roleLabel(page.role))),
                    h('div', { className: 'ps-tpl-meta' },
                      mapped
                        ? h('span', { className: 'ps-tpl-mapped', 'data-ps-template-deck-page': String(mapping.indexOf(mapped) + 1) }, fmt(t('template.mapped'), { n: mapping.indexOf(mapped) + 1 }))
                        : null,
                      page.family ? h('span', { className: 'ps-hint' }, `family ${page.family}`) : null),
                    Array.isArray(page.fields) && page.fields.length > 0
                      ? h('div', { className: 'ps-tpl-fields', 'data-ps-template-fields': String(page.fields.length) },
                          page.fields.map((field) => h('div', { key: `${pageIndex}:${field.id}`, className: 'ps-tpl-field', 'data-ps-template-field': String(field.id) },
                            h('code', null, field.id),
                            field.sample ? h('span', { className: 'sample', title: String(field.sample) }, String(field.sample).slice(0, 18)) : null,
                            Number(field.maxCjk) > 0
                              ? h('span', { className: 'cap', 'data-ps-field-cap': String(field.maxCjk) }, fmt(t('template.fieldCap'), { n: field.maxCjk }))
                              : null)))
                      : h('div', { className: 'ps-hint' }, t('template.noFields')),
                    filled.length > 0
                      ? h('div', { className: 'ps-hint', 'data-ps-template-fill': String(mapping.indexOf(mapped)) }, `${t('template.mappedFields')}：${filled.join('、')}`)
                      : null,
                    mapped && mapped.title ? h('div', { className: 'ps-tpl-title', title: mapped.title }, mapped.title) : null,
                    page.guidance ? h('div', { className: 'ps-hint' }, page.guidance) : null)
                }))
              : h('div', { className: 'ps-hint' }, t('template.noPages')),
            pages.length > 0 && mapping.length === 0
              ? h('div', { className: 'ps-hint', 'data-ps-template-mapping': 'empty' }, t('template.mappingEmpty'))
              : null)
        }
        return h('div', { className: 'ps-panel', 'data-ps-panel': 'template' },
          h('div', { className: 'ps-panel-head' },
            h('h4', null, t('stage.template')),
            h('span', {
              className: 'ps-badge' + (templateDecomposed ? ' ok' : ''),
              'data-ps-template-status': templateDecomposed ? 'done' : (pipelineTemplate ? 'current' : 'pending'),
            }, templateDecomposed ? t('stage.done') : pipelineTemplate ? t('stage.current') : t('stage.pending')),
            h('span', { style: { flex: 1 } }),
            templateDecomposed
              ? h('button', { className: 'ps-btn ghost', type: 'button', 'data-ps-btn': 'template-reload', onClick: () => setTemplateNonce((n) => n + 1) }, t('template.reload'))
              : null,
            templateDecomposed && !research
              ? h('button', { className: 'ps-btn', type: 'button', 'data-ps-btn': 'goto-research', onClick: () => setView('research') }, t('stage.stepResearch'))
              : null),
          body)
      })()


      const currentTransition = (current && current.transition) || null
      const transitionType = currentTransition && currentTransition.type ? currentTransition.type : 'none'
      const directed = DIRECTED_TRANSITIONS.has(transitionType)
      const bgColor = current && current.background && /^#[0-9A-Fa-f]{6}$/.test(String(current.background.color || ''))
        ? current.background.color
        : theme.colors.background
      const settingsPanel = settingsOpen && current
        ? h('div', { className: 'ps-settings', 'data-ps-settings': '1' },
            h('div', { className: 'ps-row' },
              h('b', null, t('slide.settings')),
              h('span', { style: { flex: 1 } }),
              h('button', { className: 'ps-btn ghost', type: 'button', onClick: () => setSettingsOpen(false) }, '×')),
            h('div', { className: 'ps-grid three' },
              h('div', { className: 'ps-field' },
                h('label', null, t('slide.transition')),
                h('select', {
                  className: 'ps-input',
                  'data-ps-field': 'transition',
                  value: transitionType,
                  onChange: (event) => setTransitionField('type', event.target.value),
                }, TRANSITION_TYPES.map((type) => h('option', { key: type, value: type }, type === 'none' ? t('slide.noneTransition') : t(`transition.${type}`))))),
              h('div', { className: 'ps-field' },
                h('label', null, t('slide.speed')),
                h('select', {
                  className: 'ps-input',
                  'data-ps-field': 'speed',
                  disabled: transitionType === 'none',
                  value: currentTransition?.speed || 'medium',
                  onChange: (event) => setTransitionField('speed', event.target.value),
                }, ['slow', 'medium', 'fast'].map((speed) => h('option', { key: speed, value: speed }, t(`speed.${speed}`))))),
              directed
                ? h('div', { className: 'ps-field' },
                    h('label', null, t('slide.direction')),
                    h('select', {
                      className: 'ps-input',
                      'data-ps-field': 'direction',
                      value: currentTransition?.direction || 'left',
                      onChange: (event) => setTransitionField('direction', event.target.value),
                    }, ['left', 'right', 'up', 'down'].map((dir) => h('option', { key: dir, value: dir }, t(`dir.${dir}`)))))
                : null),
            h('div', { className: 'ps-row' },
              h('label', { className: 'ps-check' },
                h('input', {
                  type: 'checkbox',
                  'data-ps-field': 'advance-click',
                  checked: currentTransition?.advanceOnClick !== false,
                  onChange: (event) => setTransitionField('advanceOnClick', event.target.checked),
                }),
                t('slide.advanceClick')),
              h('div', { className: 'ps-field', style: { width: 96 } },
                h('label', null, t('slide.advanceAfter')),
                h('input', {
                  className: 'ps-input',
                  type: 'number',
                  min: 0,
                  step: 250,
                  'data-ps-field': 'advance-ms',
                  value: String(currentTransition?.advanceAfterMs || 0),
                  onChange: (event) => setTransitionField('advanceAfterMs', Math.max(0, Number(event.target.value) || 0)),
                })),
              h('button', {
                className: 'ps-btn',
                type: 'button',
                'data-ps-btn': 'play-transition',
                disabled: transitionType === 'none',
                onClick: () => playFx({ kind: 'transition', type: transitionType, direction: currentTransition?.direction || 'left' }),
              }, t('slide.play'))),
            h('div', { className: 'ps-row' },
              h('span', { className: 'ps-meta' }, t('slide.background')),
              h('input', {
                className: 'ps-color',
                type: 'color',
                'data-ps-field': 'background',
                value: cssColor(bgColor, '#FFFFFF'),
                onChange: (event) => setBackground(event.target.value),
              }),
              h('button', { className: 'ps-btn ghost', type: 'button', onClick: () => setBackground('') }, '⌀'),
              h('span', { style: { flex: 1 } }),
              h('label', { className: 'ps-check' },
                h('input', {
                  type: 'checkbox',
                  'data-ps-field': 'hidden',
                  checked: Boolean(current.hidden),
                  onChange: (event) => toggleHidden(event.target.checked),
                }),
                t('slide.hidden'))),
            Number.isInteger(current.templateSlide)
              ? h('div', { className: 'ps-row', 'data-ps-slide-template': String(current.templateSlide) },
                  h('span', { className: 'ps-meta' }, fmt(t('template.fromSlide'), { n: current.templateSlide })),
                  h('span', { style: { flex: 1 } }),
                  h('button', {
                    className: 'ps-btn ghost',
                    type: 'button',
                    'data-ps-btn': 'goto-template',
                    onClick: () => setView('template'),
                  }, t('stage.template')))
              : null,
            template && Array.isArray(template.layouts) && template.layouts.length > 0
              ? h('div', { className: 'ps-field' },
                  h('label', null, t('slide.layout')),
                  h('select', {
                    className: 'ps-input',
                    'data-ps-field': 'layout',
                    value: String(current.layoutIndex ?? 0),
                    onChange: (event) => commitLayout(event.target.value),
                  }, template.layouts.map((layout, index) => h('option', { key: index, value: String(index) }, layout.name || `Layout ${index + 1}`))))
              : null,
            h('div', { className: 'ps-field' },
              h('label', null, t('slide.notes')),
              h('textarea', {
                className: 'ps-notes',
                'data-ps-field': 'notes',
                value: notesDraft,
                placeholder: t('slide.notes'),
                onChange: (event) => setNotesDraft(event.target.value),
                onBlur: commitNotes,
              })))
        : null

      const themePanel = themeOpen
        ? h('div', { className: 'ps-settings', 'data-ps-theme': '1' },
            h('div', { className: 'ps-row' },
              h('b', null, t('theme.title')),
              h('span', { style: { flex: 1 } }),
              h('button', { className: 'ps-btn ghost', type: 'button', onClick: () => setThemeOpen(false) }, '×')),
            h('div', { className: 'ps-grid two' },
              ['ink', 'muted', 'accent', 'surface', 'background'].map((key) => h('div', { className: 'ps-field', key },
                h('label', null, `${t('theme.colors')} · ${key}`),
                h('input', {
                  className: 'ps-color',
                  type: 'color',
                  'data-ps-theme-color': key,
                  value: cssColor(theme.colors[key], '#FFFFFF'),
                  onChange: (event) => commitTheme({ colors: { [key]: event.target.value } }),
                })))),
            h('div', { className: 'ps-grid two' },
              h('div', { className: 'ps-field' },
                h('label', null, `${t('theme.font')} · latin`),
                h('input', {
                  className: 'ps-input',
                  'data-ps-theme-font': 'latin',
                  defaultValue: theme.fonts.latin,
                  onBlur: (event) => commitTheme({ fonts: { latin: event.target.value } }),
                })),
              h('div', { className: 'ps-field' },
                h('label', null, `${t('theme.font')} · heading`),
                h('input', {
                  className: 'ps-input',
                  'data-ps-theme-font': 'heading',
                  defaultValue: theme.fonts.heading,
                  onBlur: (event) => commitTheme({ fonts: { heading: event.target.value } }),
                }))),
            h('div', { className: 'ps-field' },
              h('label', null, t('theme.size')),
              h('select', {
                className: 'ps-input',
                'data-ps-field': 'size',
                value: sizePresetOf(deck.size),
                onChange: (event) => commitSize(event.target.value),
              }, SIZE_PRESETS.map((preset) => h('option', { key: preset.id, value: preset.id }, t(preset.label))).concat(
                sizePresetOf(deck.size) === 'custom' ? [h('option', { key: 'custom', value: 'custom' }, 'custom')] : []))))
        : null

      const isTransitionFx = Boolean(fx && fx.kind === 'transition' && fx.type && fx.type !== 'none')
      // The editing surface keeps a real slide-shaped placeholder even before a
      // single slide exists, so the area never collapses into a text notice and
      // then "pops" into a full deck.
      const skeletonStage = h('div', { className: 'ps-skeleton', 'data-ps-skeleton': '1' },
        h('div', { className: 'ps-skeleton-paper' },
          h('span', { className: 'sk-line', style: { width: '46%', height: '7%' } }),
          h('span', { className: 'sk-line', style: { width: '72%', height: '4%', marginTop: '5%' } }),
          h('span', { className: 'sk-block', style: { width: '100%', height: '34%', marginTop: '6%' } }),
          h('span', { className: 'sk-line', style: { width: '58%', height: '4%', marginTop: '5%' } })),
        h('div', { className: 'ps-skeleton-hint' }, t('empty.shellBody')))
      // The HUD belongs to the page on screen; if the user has navigated away
      // (follow is off) the status bar still reports the build, but the stage
      // does not pretend to be drawing a page it is not showing.
      const drawingNow = Boolean(draw && draw.slideIndex === safeIndex + 1)
      const drawPercent = draw && draw.count > 0 ? Math.round((Math.min(draw.done, draw.count) / draw.count) * 100) : 0
      const drawHud = draw && drawingNow
        ? h('div', { className: 'ps-draw', 'data-ps-draw': '1', 'data-count': String(draw.count), 'data-done': String(draw.done) },
            h('div', { className: 'ps-draw-top' },
              h('span', { className: 'ps-draw-pen', 'aria-hidden': 'true' }, '✎'),
              h('span', { className: 'ps-draw-text' }, fmt(t('status.elements'), { index: draw.slideIndex, done: Math.min(draw.done, draw.count), count: draw.count })),
              h('span', { className: 'ps-draw-count' }, `${Math.min(draw.done, draw.count)}/${draw.count}`)),
            h('div', { className: 'ps-draw-bar' }, h('span', { style: { width: `${drawPercent}%` } })))
        : null
      const slideStage = h('div', {
        className: 'ps-stage',
        ref: stageRef,
        'data-ps-stage': '1',
        'data-drawing': String(drawingNow),
      },
        h('div', { className: 'ps-stage-inner' },
        current
          ? h('div', {
              className: 'ps-canvas',
              ref: canvasRef,
              'data-ps-canvas': deckId,
              'data-marquee': String(Boolean(marquee)),
              style: {
                width: `${slideW * PX_PER_IN * zoom}px`,
                height: `${slideH * PX_PER_IN * zoom}px`,
                background: cssColor(bgColor, '#FFFFFF'),
              },
            },
              h('div', {
                className: 'ps-slide',
                key: `${current.id}:${isTransitionFx ? fx.nonce : ''}`,
                ...(isTransitionFx ? { 'data-fx': fx.type, 'data-fx-dir': fx.direction || 'left' } : {}),
                onPointerDown: onCanvasPointerDown,
                onPointerMove: onCanvasPointerMove,
                onPointerUp: onCanvasPointerUp,
              },
                (current.elements || []).map((el) => renderElement(el, {
                  zoom,
                  deckId,
                  slideW,
                  slideH,
                  rev,
                  theme,
                  background: bgColor,
                  approx: t('approx'),
                  interactive: true,
                  selected: Boolean(selection && selection.kind === 'element' && selection.elementId === el.id),
                  fresh: Boolean(freshEls[el.id]),
                  flash: Boolean(flashEls[el.id]),
                  sketch: Boolean(sketchEls[el.id]),
                  animPlay: isTransitionFx || Boolean(fx && fx.kind === 'element' && fx.elementId === el.id),
                  overrideBox: preview && preview.elementId === el.id ? preview.box : null,
                  onPointerDown: onElementPointerDown,
                  onPointerMove: onElementPointerMove,
                  onPointerUp: onElementPointerUp,
                  onDoubleClick: beginTextEdit,
                })),
                overlay,
                inlineEditor,
                marqueeBox ? h('div', { className: 'ps-region', style: placement([marqueeBox.left, marqueeBox.top, marqueeBox.width, marqueeBox.height], zoom) }) : null,
              ),
            )
          : skeletonStage,
        drawHud),
        dragOver ? h('div', { className: 'ps-drop' }, t('actions.addImage')) : null,
      )

      let multiInspector = null
      if (multiSelection && selectedIds.length > 0 && current) {
        multiInspector = h('div', { className: 'ps-insp', 'data-ps-multi': '1' },
          h('div', { className: 'ps-row' },
            h('b', null, fmt(t('select.multi'), { n: selectedIds.length })),
            h('span', { style: { flex: 1 } }),
            h('button', { className: 'ps-btn', type: 'button', 'data-ps-btn': 'multi-copy', onClick: copySelection }, t('shortcut.copy')),
            h('button', { className: 'ps-btn', type: 'button', 'data-ps-btn': 'multi-cut', onClick: cutSelection }, t('shortcut.cut')),
            h('button', { className: 'ps-btn danger', type: 'button', 'data-ps-btn': 'multi-delete', onClick: deleteSelected }, t('actions.delete')),
            h('button', { className: 'ps-btn ghost', type: 'button', onClick: () => setMultiSelection(null) }, t('select.clear'))),
          h('div', { className: 'ps-hint' }, t('shortcut.hint')))
      }

      const shortcutRows = [
        ['Ctrl+S', 'shortcut.save'],
        ['Ctrl+Shift+S', 'shortcut.saveAs'],
        ['Ctrl+Z', 'shortcut.undo'],
        ['Ctrl+Shift+Z / Ctrl+Y', 'shortcut.redo'],
        ['Ctrl+X', 'shortcut.cut'],
        ['Ctrl+C', 'shortcut.copy'],
        ['Ctrl+V', 'shortcut.paste'],
        ['Ctrl+D', 'shortcut.duplicate'],
        ['Ctrl+M', 'shortcut.newSlide'],
        ['Ctrl+A', 'shortcut.selectAll'],
        ['Delete', 'shortcut.delete'],
        ['← ↑ → ↓', 'shortcut.nudge'],
        ['Ctrl+0 / Ctrl+±', 'shortcut.zoom'],
        ['Esc', 'shortcut.deselect'],
        ['?', 'shortcut.help'],
      ]
      const shortcutsPanel = h('div', { className: 'ps-shortcuts', 'data-ps-shortcuts': '1' },
        h('div', { className: 'ps-row' },
          h('b', null, t('shortcuts.title')),
          h('span', { style: { flex: 1 } }),
          h('button', { className: 'ps-btn ghost', type: 'button', onClick: () => setShortcutsOpen(false) }, '×')),
        h('div', { className: 'ps-hint' }, t('shortcuts.hint')),
        shortcutRows.map(([keys, label]) => h('div', { className: 'ps-shortcut', key: keys },
          h('kbd', null, keys),
          h('span', null, t(label)))))

      // ── render ─────────────────────────────────────────────────────────────

      return h('div', {
        className: 'ps-root',
        'data-ps-root': 'deck',
        tabIndex: 0,
        ref: rootRef,
        onKeyDown: handleShortcut,
        onDragOver: (event) => {
          if (!current) return
          if (event.dataTransfer && Array.from(event.dataTransfer.types || []).includes('Files')) {
            event.preventDefault()
            setDragOver(true)
          }
        },
        onDragLeave: () => setDragOver(false),
        onDrop: (event) => {
          if (!current) return
          const files = (event.dataTransfer && event.dataTransfer.files) || []
          if (files.length === 0) return
          event.preventDefault()
          setDragOver(false)
          const at = pointOf(event.clientX, event.clientY)
          void insertImageFile(files[0], at)
        },
      },
        // Bar: live dot, title (inline editable), deck picker, primary actions.
        h('div', { className: 'ps-bar' },
          h('span', { className: 'ps-dot' + (live ? ' live' : ''), title: live ? t('live.on') : t('live.off') }),
          titleEditing
            ? h('input', {
                className: 'ps-title-input',
                autoFocus: true,
                value: titleDraft,
                onChange: (event) => setTitleDraft(event.target.value),
                onBlur: commitTitle,
                onKeyDown: (event) => {
                  if (event.key === 'Enter') { event.preventDefault(); commitTitle() }
                  if (event.key === 'Escape') { event.preventDefault(); setTitleEditing(false) }
                },
              })
            : h('span', {
                className: 'ps-title',
                title: t('edit.title'),
                onDoubleClick: () => { setTitleDraft(String(deck.title || '')); setTitleEditing(true) },
              }, deck.title || t('tab.title')),
          h('span', { className: 'ps-meta' }, fmt('{n} 页', { n: slides.length })),
          decks.length > 1
            ? h('select', {
                className: 'ps-btn ps-select',
                value: deckId,
                title: t('deck.picker'),
                onChange: (event) => { setDeckId(event.target.value); setSlideIndex(0); setSelection(null); setMultiSelection(null); setEditing(null) },
              }, decks.map((item) => h('option', { key: item.id, value: item.id }, `${item.title || item.id} (${item.slides})`)))
            : null,
          deckId
            ? h('button', {
                className: 'ps-btn' + (confirmDelete ? ' danger' : ' ghost'),
                type: 'button',
                'data-ps-btn': 'deck-delete',
                title: confirmDelete ? t('deck.confirm') : t('deck.delete'),
                onClick: () => {
                  if (!confirmDelete) {
                    setConfirmDelete(true)
                    return
                  }
                  void removeDeck()
                },
              }, confirmDelete ? t('deck.confirm') : '×')
            : null,
          h('button', {
            className: 'ps-btn primary',
            type: 'button',
            'data-ps-btn': 'export',
            title: saved?.path ? `Ctrl+S · ${saved.path}` : `Ctrl+S · ${t('actions.saveHint')}`,
            disabled: busy === 'export' || slides.length === 0,
            onClick: () => { void saveDeck() },
          }, busy === 'export' ? t('actions.exporting') : t('actions.save')),
          h('button', {
            className: 'ps-btn',
            type: 'button',
            'data-ps-btn': 'download',
            title: `Ctrl+Shift+S · ${t('actions.downloadHint')}`,
            disabled: !deckId || busy !== '' || slides.length === 0,
            onClick: () => { void downloadDeck() },
          }, busy === 'download' ? t('actions.downloading') : t('actions.download')),
        ),

        // Workflow steps: ⓪ template → ① research → ② HTML preview → ③ slides.
        steps,

        // The running step, its deck and its elapsed time (live progress that
        // does not depend on the user finding the tool card in the chat).
        taskBanner,

        // Editing zone: the slide/panel surface is a visually separate card.
        h('div', { className: 'ps-editor', 'data-ps-editor': '1' },
          h('div', { className: 'ps-editor-head' },
            h('span', { className: 'ps-zone-label' }, t('zone.editor')),
            h('span', { className: 'ps-zone-sub' }, activeView === 'slides' ? t('zone.slide') : t(stageTitleKey(activeView === 'slides' ? 'build' : activeView))),
            current && Number.isInteger(current.templateSlide)
              ? h('span', {
                  className: 'ps-tpl-chip',
                  'data-ps-stage-template': String(current.templateSlide),
                  title: fmt(t('template.fromSlide'), { n: current.templateSlide }),
                }, `T${current.templateSlide}`)
              : null,
            h('span', { style: { flex: 1 } }),
            draw
              ? h('span', { className: 'ps-draw-badge', 'data-ps-drawing': '1', title: fmt(t('status.elements'), { index: draw.slideIndex, done: Math.min(draw.done, draw.count), count: draw.count }) },
                  h('span', { className: 'ps-draw-dot', 'aria-hidden': 'true' }),
                  t('zone.drawing'))
              : null,
            activeView === 'slides' ? h('span', { className: 'ps-meta' }, `${safeIndex + 1}/${slides.length}`) : null,
            activeView === 'slides' ? h('span', { className: 'ps-meta' }, fmt(t('zoom.label'), { pct: Math.round(zoom * 100) })) : null),
          // The rail is always mounted in slide view: an empty rail keeps the
          // stage at the same height as a full one, so the first slide does not
          // shove the canvas around.
          activeView === 'slides'
            ? h('div', { className: 'ps-thumbs', 'data-ps-thumbs': '1', 'data-empty': String(slides.length === 0) },
                slides.length > 0
                  ? slides.map(thumb)
                  : [0, 1, 2].map((slot) => h('span', {
                      key: `ghost${slot}`,
                      className: 'ps-thumb ghost',
                      'aria-hidden': 'true',
                      style: { width: `${thumbWidth}px`, height: `${thumbHeight}px` },
                    })))
            : null,
          activeView === 'template' ? templatePanel : activeView === 'research' ? researchPanel : activeView === 'preview' ? previewPanel : slideStage),

        // A visible divider carries the "editor / controls" boundary.
        h('div', { className: 'ps-split', 'data-ps-split': '1' },
          h('span', { className: 'ps-split-grip' }, t('zone.split'))),

        // Control dock: inputs, inspector, effects, notes and shortcut help.
        h('div', { className: 'ps-controls', 'data-ps-controls': '1' },
          h('div', { className: 'ps-controls-head' },
            h('span', { className: 'ps-zone-label' }, t('zone.controls')),
            h('span', { className: 'ps-zone-sub' }, selectionLabel || (multiSelection ? fmt(t('select.multi'), { n: selectedIds.length }) : '')),
            h('span', { style: { flex: 1 } }),
            h('button', {
              className: 'ps-btn ghost',
              type: 'button',
              'data-ps-btn': 'shortcuts',
              'data-active': String(shortcutsOpen),
              title: t('shortcuts.title'),
              onClick: () => setShortcutsOpen((open) => !open),
            }, t('actions.shortcuts'))),
          h('div', { className: 'ps-foot ps-controls-body' },
            shortcutsOpen ? shortcutsPanel : null,
          h('div', { className: 'ps-row' },
            h('span', { className: 'ps-status', 'data-ps-status': '1', title: statusText }, statusText),
            dirty > 0 ? h('span', { className: 'ps-meta' }, t('edit.unsaved')) : h('span', { className: 'ps-ok' }, t('edit.applied')),
            h('span', { className: 'ps-meta', 'data-ps-zoom': String(Math.round(zoom * 100)) }, fmt(t('zoom.label'), { pct: Math.round(zoom * 100) })),
          ),
          activeView === 'slides'
            ? h('div', { className: 'ps-row' },
                h('div', { className: 'ps-seg' },
                  h('button', { className: 'ps-btn', type: 'button', title: t('actions.zoomOut'), onClick: () => setZoomMode(clamp(zoom - 0.1, ZOOM_MIN, ZOOM_MAX)) }, '−'),
                  h('button', { className: 'ps-btn', type: 'button', 'data-active': String(zoomMode === 'fit'), onClick: () => setZoomMode('fit') }, t('actions.fit')),
                  h('button', { className: 'ps-btn', type: 'button', 'data-active': String(zoomMode === 1), onClick: () => setZoomMode(1) }, t('actions.actual')),
                  h('button', { className: 'ps-btn', type: 'button', title: t('actions.zoomIn'), onClick: () => setZoomMode(clamp(zoom + 0.1, ZOOM_MIN, ZOOM_MAX)) }, '+'),
                ),
                h('button', { className: 'ps-btn ghost', type: 'button', 'data-ps-btn': 'undo', disabled: history.undo === 0, title: t('actions.undo'), onClick: undo }, '↶'),
                h('button', { className: 'ps-btn ghost', type: 'button', 'data-ps-btn': 'redo', disabled: history.redo === 0, title: t('actions.redo'), onClick: redo }, '↷'),
                h('span', { style: { flex: 1 } }),
                h('button', { className: 'ps-btn', type: 'button', 'data-ps-btn': 'settings', 'data-active': String(settingsOpen), title: t('slide.settings'), onClick: () => { setSettingsOpen((open) => !open); setThemeOpen(false) } }, t('actions.settings')),
                h('button', { className: 'ps-btn', type: 'button', 'data-ps-btn': 'theme', 'data-active': String(themeOpen), title: t('theme.title'), onClick: () => { setThemeOpen((open) => !open); setSettingsOpen(false) } }, t('actions.theme')),
                h('button', { className: 'ps-btn', type: 'button', 'data-ps-btn': 'add-slide', title: t('actions.addSlide'), onClick: addSlide }, '+ 页'),
                h('button', { className: 'ps-btn ghost', type: 'button', title: t('actions.slideLeft'), disabled: safeIndex <= 0, onClick: () => moveSlide(-1) }, '←'),
                h('button', { className: 'ps-btn ghost', type: 'button', title: t('actions.slideRight'), disabled: safeIndex >= slides.length - 1, onClick: () => moveSlide(1) }, '→'),
                h('button', { className: 'ps-btn ghost danger', type: 'button', title: t('actions.removeSlide'), disabled: slides.length <= 1, onClick: removeSlide }, '🗑'),
              )
            : h('div', { className: 'ps-row' },
                h('button', { className: 'ps-btn', type: 'button', 'data-ps-btn': 'goto-slides', onClick: () => setView('slides') }, t('stage.stepSlides')),
                h('span', { style: { flex: 1 } }),
                h('label', { className: 'ps-meta', style: { display: 'inline-flex', alignItems: 'center', gap: 4 } },
                  h('input', { type: 'checkbox', checked: follow, onChange: (event) => setFollow(event.target.checked) }),
                  t('status.follow'),
                ),
                h('button', { className: 'ps-btn ghost', type: 'button', title: t('actions.refresh'), onClick: () => { void loadState(deckId) } }, '↻'),
                tab ? h('button', { className: 'ps-btn ghost', type: 'button', title: t('actions.close'), onClick: closeTab }, '✕') : null,
              ),
          activeView === 'slides'
            ? h('div', { className: 'ps-row' },
                h('button', { className: 'ps-btn', type: 'button', 'data-ps-btn': 'add-text', disabled: !current, onClick: addTextBox }, t('actions.addText')),
                h('button', { className: 'ps-btn', type: 'button', 'data-ps-btn': 'add-image', disabled: !current || busy === 'upload', onClick: () => { try { fileRef.current?.click() } catch { /* ignore */ } } },
                  busy === 'upload' ? t('image.uploading') : t('actions.addImage')),
                h('button', { className: 'ps-btn', type: 'button', 'data-ps-btn': 'add-rect', disabled: !current, onClick: addRect }, t('actions.addRect')),
                h('button', { className: 'ps-btn', type: 'button', 'data-ps-btn': 'add-table', disabled: !current, onClick: addTable }, t('actions.addTable')),
                h('button', { className: 'ps-btn', type: 'button', 'data-ps-btn': 'add-chart', disabled: !current, onClick: addChart }, t('actions.addChart')),
                h('input', {
                  ref: fileRef,
                  type: 'file',
                  accept: 'image/*',
                  style: { display: 'none' },
                  onChange: (event) => {
                    const file = event.target.files && event.target.files[0]
                    if (file) void insertImageFile(file, null)
                    try { event.target.value = '' } catch { /* ignore */ }
                  },
                }),
                h('span', { style: { flex: 1 } }),
                h('label', { className: 'ps-meta', style: { display: 'inline-flex', alignItems: 'center', gap: 4 } },
                  h('input', { type: 'checkbox', checked: follow, onChange: (event) => setFollow(event.target.checked) }),
                  t('status.follow'),
                ),
                h('button', { className: 'ps-btn ghost', type: 'button', title: t('actions.refresh'), onClick: () => { void loadState(deckId) } }, '↻'),
                tab ? h('button', { className: 'ps-btn ghost', type: 'button', title: t('actions.close'), onClick: closeTab }, '✕') : null,
              )
            : null,

          activeView === 'slides' ? settingsPanel : null,
          activeView === 'slides' ? themePanel : null,
          activeView === 'slides' ? (multiInspector || inspector) : null,

          activeView === 'slides'
            ? (multiSelection && selectedIds.length > 0
                ? h('div', { className: 'ps-selecting' },
                    h('b', null, fmt(t('select.multi'), { n: selectedIds.length })),
                    h('span', null, t('shortcut.hint')),
                    h('button', { className: 'ps-btn ghost', type: 'button', onClick: () => setMultiSelection(null) }, t('select.clear')),
                  )
                : selection
                  ? h('div', { className: 'ps-selecting' },
                      h('b', null, selection.kind === 'element' ? t('select.element') : t('select.region')),
                      h('span', null, selectionLabel),
                      h('button', { className: 'ps-btn ghost', type: 'button', onClick: () => setSelection(null) }, t('select.clear')),
                    )
                  : h('div', { className: 'ps-selecting' }, h('span', null, t('select.hint'))))
            : null,

          activeView === 'slides' && selection
            ? h('div', null,
                h('textarea', {
                  className: 'ps-note',
                  'data-ps-note': '1',
                  value: note,
                  placeholder: t('note.placeholder'),
                  onChange: (event) => setNote(event.target.value),
                  onKeyDown: (event) => {
                    if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) {
                      event.preventDefault()
                      void sendFeedback()
                    }
                  },
                }),
                h('div', { className: 'ps-chips', style: { marginTop: 4 } },
                  tList(t, 'note.chips').map((chip) => h('button', {
                    key: chip,
                    className: 'ps-chip',
                    type: 'button',
                    onClick: () => setNote((prev) => (prev.trim() === '' ? chip : `${prev}；${chip}`)),
                  }, chip)),
                ),
                h('div', { className: 'ps-row', style: { marginTop: 6 } },
                  h('button', {
                    className: 'ps-btn primary',
                    type: 'button',
                    'data-ps-btn': 'send',
                    disabled: sending,
                    onClick: () => { void sendFeedback() },
                  }, sending ? t('actions.sending') : t('actions.send')),
                  h('button', { className: 'ps-btn', type: 'button', onClick: () => { void copyFeedback() } }, t('actions.copy')),
                  feedbackHint ? h('span', { className: 'ps-status' }, feedbackHint) : null,
                ),
              )
            : null,
          error ? h('div', { className: 'ps-err' }, error) : null,
          ),
        ),
      )
    }

    /**
     * Call the slot's `useTabInfo` hook without letting a missing or older
     * contract throw (it is a hook, so this must run unconditionally).
     * @param {any} useTabInfo - slot hook.
     * @returns {any} snapshot or null.
     */
    function readTabInfo(useTabInfo) {
      try {
        if (typeof useTabInfo === 'function') return useTabInfo()
      } catch {
        return null
      }
      return null
    }

    /**
     * Read an image's natural size.
     * @param {string} dataUrl - image data URL.
     * @returns {Promise<number[]>} `[width, height]` (or `[0, 0]`).
     */
    function imageSize(dataUrl) {
      return new Promise((resolvePromise) => {
        let settled = false
        /** @param {number[]} value - `[width, height]`. */
        const finish = (value) => {
          if (settled) return
          settled = true
          resolvePromise(value)
        }
        // A decoder that never fires (odd formats, blocked data URLs) must not
        // wedge the upload: fall back to a default aspect after a moment.
        const timer = setTimeout(() => finish([0, 0]), 800)
        if (typeof timer.unref === 'function') timer.unref()
        try {
          if (typeof Image !== 'function') {
            finish([0, 0])
            return
          }
          const image = new Image()
          image.onload = () => {
            clearTimeout(timer)
            finish([image.naturalWidth || image.width || 0, image.naturalHeight || image.height || 0])
          }
          image.onerror = () => {
            clearTimeout(timer)
            finish([0, 0])
          }
          image.src = dataUrl
        } catch {
          clearTimeout(timer)
          finish([0, 0])
        }
      })
    }

    /**
     * Read a File as a data URL.
     * @param {File} file - file.
     * @returns {Promise<string>} data URL.
     */
    function fileToDataUrl(file) {
      return new Promise((resolvePromise, reject) => {
        const reader = new FileReader()
        reader.onload = () => resolvePromise(String(reader.result || ''))
        reader.onerror = () => reject(reader.error || new Error('read failed'))
        reader.readAsDataURL(file)
      })
    }

    /**
     * Diff two deck models into the minimal op list the host can replay.
     * @param {any} before - previous model.
     * @param {any} after - desired model.
     * @returns {Record<string, any>[]} ops.
     */
    function diffDecks(before, after) {
      const ops = []
      if (!before || !after) return ops
      if (before.title !== after.title) ops.push({ kind: 'deck.title', title: after.title })
      if (JSON.stringify(before.theme || {}) !== JSON.stringify(after.theme || {})) {
        ops.push({ kind: 'deck.theme', theme: after.theme || {} })
      }
      if (before.size?.w !== after.size?.w || before.size?.h !== after.size?.h) {
        ops.push({ kind: 'deck.size', width: after.size?.w, height: after.size?.h })
      }
      const beforeSlides = before.slides || []
      const afterSlides = after.slides || []
      const beforeById = new Map(beforeSlides.map((slide) => [slide.id, slide]))
      for (const slide of beforeSlides) {
        if (!afterSlides.some((item) => item.id === slide.id)) ops.push({ kind: 'slide.remove', slideId: slide.id })
      }
      /** Slide-level keys that travel through `slide.patch`. */
      const metaKeys = ['title', 'notes', 'background', 'transition', 'hidden', 'layoutIndex', 'layout', 'intent']
      afterSlides.forEach((slide, index) => {
        const old = beforeById.get(slide.id)
        if (!old) {
          ops.push({
            kind: 'slide.add',
            index,
            slide: {
              id: slide.id,
              title: slide.title,
              elements: slide.elements || [],
              ...(slide.transition ? { transition: slide.transition } : {}),
              ...(slide.background ? { background: slide.background } : {}),
              ...(typeof slide.notes === 'string' ? { notes: slide.notes } : {}),
              ...(slide.hidden ? { hidden: true } : {}),
              ...(slide.layoutIndex !== undefined ? { layoutIndex: slide.layoutIndex } : {}),
            },
          })
          return
        }
        const oldIndex = beforeSlides.indexOf(old)
        if (oldIndex !== index) ops.push({ kind: 'slide.move', slideId: slide.id, index })
        /** @type {Record<string, unknown>} */
        const meta = {}
        for (const key of metaKeys) {
          if (JSON.stringify(old[key] ?? null) !== JSON.stringify(slide[key] ?? null)) {
            meta[key] = slide[key] === undefined ? null : slide[key]
          }
        }
        if (Object.keys(meta).length > 0) ops.push({ kind: 'slide.patch', slideId: slide.id, patch: meta })
        const oldEls = new Map((old.elements || []).map((el) => [el.id, el]))
        for (const el of old.elements || []) {
          if (!(slide.elements || []).some((item) => item.id === el.id)) {
            ops.push({ kind: 'element.patch', slideId: slide.id, elementId: el.id, remove: true })
          }
        }
        for (const el of slide.elements || []) {
          const previous = oldEls.get(el.id)
          if (!previous) ops.push({ kind: 'element.add', slideId: slide.id, element: el })
          else if (JSON.stringify(previous) !== JSON.stringify(el)) {
            ops.push({ kind: 'element.replace', slideId: slide.id, elementId: el.id, element: el })
          }
        }
      })
      return ops
    }

    /**
     * Tab title chip. It carries a live dot while a pptx step is running, so
     * the tab itself says the work is still going on.
     * @param {any} props - slot props.
     * @returns {import('react').ReactNode} node.
     */
    function TabTitle(props) {
      const t = typeof props.t === 'function' ? props.t : (key) => key
      const task = useSessionTask(props.sessionId)
      const running = Boolean(task && task.state === 'running')
      return h('span', { className: 'ps-scope', 'data-ps-tab-task': running ? String(task.tool || 'busy') : '' },
        t('tab.title'),
        running ? h('span', { className: 'ps-tab-dot', 'aria-hidden': 'true' }) : null)
    }

    /** Every `pptx_studio_*` tool gets a card (whoever runs, the canvas follows). */
    const TOOL_CARD_KEYS = [
      'pptx_studio_open',
      'pptx_studio_template',
      'pptx_studio_research',
      'pptx_studio_preview',
      'pptx_studio_deck',
      'pptx_studio_plan',
      'pptx_studio_slide',
      'pptx_studio_add_element',
      'pptx_studio_patch',
      'pptx_studio_inspect',
      'pptx_studio_verify',
      'pptx_studio_export',
    ]

    /**
     * Tool-call card for every `pptx_studio_*` tool: it switches the right-hand
     * canvas to the deck this call writes (and focuses it) the moment the call
     * shows up, and reports the step's progress next to the result.
     *
     * The regression it guards: opening the canvas only for `pptx_studio_open`,
     * or only once the call returned, leaves the user staring at the previous
     * build (or at the conversation) while the agent draws the new deck.
     *
     * @param {any} props - tool view props.
     * @returns {import('react').ReactNode} node.
     */
    function StudioCard(props) {
      const t = typeof props.t === 'function' ? props.t : (key, values) => fmt(key, values)
      const phase = props.phase
      const toolName = typeof props.toolName === 'string' && props.toolName !== '' ? props.toolName : 'pptx_studio_open'
      const args = props.block && props.block.args
      const readArg = (key) => {
        try {
          const value = args && typeof args.text === 'function' ? args.text(key) : ''
          return typeof value === 'string' ? value : ''
        } catch {
          return ''
        }
      }
      const title = readArg('title')
      const template = readArg('template')
      const argDeckId = readArg('deck') || readArg('deckId')
      const resultValue = (props.block && props.block.result && props.block.result.value) || null
      const resultDeckId = resultValue && typeof resultValue.deckId === 'string' ? resultValue.deckId : ''
      // `pptx_studio_open` only learns its deck from the receipt; every other
      // step names the deck in its arguments, so the tab can switch before it runs.
      const deckId = resultDeckId || argDeckId
      const slideCount = resultValue && typeof resultValue.slides === 'number' ? resultValue.slides : null
      const failedResult = resultValue && resultValue.ok === false ? String(resultValue.message || resultValue.code || '') : ''
      const [opened, setOpened] = useState(false)
      const [failed, setFailed] = useState('')
      const running = phase === 'preparing' || phase === 'start'
      const autoRef = useRef(false)
      /** Deck the tab was last pointed at by this card (`''` = pending shell). */
      const pointedRef = useRef('')

      const open = useCallback((params) => {
        try {
          const success = (props.openStudio ? props.openStudio(params) : false) !== false
          setOpened(success)
          setFailed(success ? '' : t('card.noTab'))
          return success
        } catch {
          setFailed(t('card.noTab'))
          return false
        }
      }, [props, t])

      // Open — and focus — the canvas the moment the call shows up, pointing it
      // at the deck this step writes. A brand-new `pptx_studio_open` has no deck
      // yet, so it opens "pending": the canvas shows the preparing shell for
      // this call instead of the previous build's deck.
      useEffect(() => {
        if (!running) return
        if (deckId !== '') {
          if (pointedRef.current === deckId) return
          pointedRef.current = deckId
          autoRef.current = true
          open({ deckId, at: Date.now() })
          return
        }
        if (autoRef.current) return
        autoRef.current = true
        pointedRef.current = ''
        open(toolName === 'pptx_studio_open' ? { pending: '1', nonce: Date.now() } : {})
      }, [deckId, open, running, toolName])

      // The deck only exists once the call ran (`pptx_studio_open` returns it):
      // navigate the tab that opened with the call to it. A card that mounts on
      // an old result must not steal focus — only a card that opened the tab does.
      useEffect(() => {
        if (phase !== 'result' || deckId === '' || !autoRef.current) return
        if (pointedRef.current === deckId) return
        pointedRef.current = deckId
        open({ deckId, at: Date.now() })
      }, [deckId, open, phase])

      const hint = running
        ? (opened ? t('card.live') : t('card.opening'))
        : (deckId !== '' ? t('card.following') : t('card.noDeckYet'))
      return h('div', {
        className: 'ps-card ps-scope',
        'data-pptx-studio-card': props.callId || '',
        'data-ps-card-tool': toolName,
        'data-ps-card-deck': deckId,
        'data-opened': String(opened),
      },
        h('div', { className: 'hd' }, toolLabel(t, toolName),
          h('span', { className: 'badge' }, t(running ? 'card.badgeRunning' : 'card.badgeDone'))),
        h('div', { className: 'sub' },
          title || (deckId !== '' ? deckId : '(untitled)'),
          template ? ` · template: ${template}` : '',
          slideCount !== null ? ` · ${fmt(t('card.slides'), { count: slideCount })}` : '',
        ),
        h('div', { className: 'sub', 'data-ps-card-hint': opened ? 'live' : 'opening' }, hint),
        failedResult ? h('div', { className: 'ps-err', 'data-ps-card-failed': '1' }, fmt(t('card.failed'), { message: failedResult })) : null,
        h('div', { className: 'ps-row' },
          h('button', {
            className: 'ps-btn primary',
            type: 'button',
            'data-ps-btn': 'card-open',
            onClick: () => open(deckId !== '' ? { deckId, at: Date.now() } : (toolName === 'pptx_studio_open' ? { pending: '1', nonce: Date.now() } : {})),
          }, t('card.open')),
          opened ? h('span', { className: 'sub' }, t('card.opened')) : null,
          failed ? h('span', { className: 'ps-err' }, failed) : null,
        ),
      )
    }

    return {
      inject: ['slots', 'locale', 'sidebarRightTabs', 'sidebarRight'],
      apply(ctx) {
        hostCtx = ctx
        // One stylesheet for every surface this plugin owns.
        ctx.effect(() => {
          const element = document.createElement('style')
          element.setAttribute('data-dsh-plugin', 'dsh-pptx-studio')
          element.textContent = CSS
          document.head.appendChild(element)
          return () => element.remove()
        }, 'pptx-studio: stylesheet')

        if (typeof ctx.locale?.register === 'function') {
          for (const [id, dict] of [['zh', zh], ['en', en]]) {
            ctx.effect(() => ctx.locale.register(NS, id, dict), `pptx-studio: ${id} dictionary`)
          }
        }
        const bound = typeof ctx.locale?.bind === 'function' ? ctx.locale.bind(NS) : null
        const fallback = String(globalThis.navigator?.language ?? 'zh').toLowerCase().startsWith('zh') ? zh : en
        const t = bound ?? ((key, values) => fmt(fallback[key] ?? key, values))

        // ① tab type (static declaration; body/title register under its id).
        ctx.effect(() => ctx.sidebarRightTabs.register({
          id: TAB_ID,
          kind: TAB_KIND,
          priority: 'extension',
          title: () => t('tab.title'),
          guide: [{
            id: 'pptx-studio',
            order: 40,
            title: () => t('tab.title'),
            description: () => t('guide.description'),
          }],
        }), 'pptx-studio: tab type')

        // ② body + title.
        ctx.slots.inject('sidebar.right.pane.tab', () => ctx.slots.register(
          { name: 'sidebar.right.pane.tab', key: TAB_ID, locale: NS },
          CanvasBody,
        ))
        ctx.slots.inject('sidebar.right.pane.tab.title', () => ctx.slots.register(
          { name: 'sidebar.right.pane.tab.title', key: TAB_ID, locale: NS },
          TabTitle,
        ))

        /**
         * Focus the canvas tab this session already has (in any pane) so the
         * open below lands on it instead of opening a second canvas beside it.
         * @param {any} controller - sidebarRight controller.
         * @param {string} sessionId - session that owns the tab.
         * @returns {void}
         */
        const focusExistingTab = (controller, sessionId) => {
          try {
            const active = typeof controller.active === 'function' ? controller.active() : undefined
            if (active && active.kind === TAB_KIND) return
            const tabs = typeof controller.tabsIn === 'function' ? controller.tabsIn(sessionId) : null
            const mine = Array.isArray(tabs) ? tabs.find((record) => record && record.kind === TAB_KIND) : null
            if (mine && typeof controller.focus === 'function') controller.focus(mine.id)
          } catch {
            /* older sidebars: `openTab` alone still opens/reveals the canvas */
          }
        }

        /**
         * Open the canvas for one session (called from the tool cards).
         *
         * `openTab` needs a session on screen; `openTabIn` works for adopted
         * sessions; when neither applies we first bring the session up and
         * retry once. Returns false only when no route exists at all.
         * @param {string} sessionId - session that owns this card.
         * @returns {(params?: Record<string, unknown>) => boolean} opener.
         */
        const openerFor = (sessionId) => (params) => {
          const controller = ctx.get('sidebarRight')
          if (!controller) return false
          const options = params ? { params } : {}
          try {
            const mounted = controller.mounted?.getSnapshot?.()
            if (mounted === sessionId && typeof controller.openTab === 'function') {
              // A split sidebar can hold the canvas in a second pane; focus it
              // first so `openTab` reveals that one instead of adding another.
              focusExistingTab(controller, sessionId)
              controller.openTab(TAB_KIND, options)
              return true
            }
            if (typeof controller.openTabIn === 'function' && typeof controller.actionsFor === 'function' && controller.actionsFor(sessionId)) {
              controller.openTabIn(sessionId, TAB_KIND, options)
              return true
            }
          } catch {
            return false
          }
          // Not adopted yet: show the session, then retry once.
          try {
            const workspace = ctx.get('uiWorkspace')
            if (workspace && typeof workspace.openSession === 'function') {
              workspace.openSession(sessionId)
              setTimeout(() => {
                try {
                  const retry = ctx.get('sidebarRight')
                  retry?.openTabIn?.(sessionId, TAB_KIND, options)
                } catch { /* ignore */ }
              }, 160)
              return true
            }
          } catch { /* ignore */ }
          return false
        }

        // ③ tool cards: one per pptx tool, so *any* step switches the right-hand
        // canvas to its deck and shows the progress (not just the open call).
        for (const key of TOOL_CARD_KEYS) {
          ctx.slots.inject('tool.call.toolview', () => ctx.slots.register({
            name: 'tool.call.toolview',
            key,
            locale: NS,
            inject: (sessionId) => ({ openStudio: openerFor(sessionId) }),
          }, StudioCard))
        }
      },
    }
  },
})
