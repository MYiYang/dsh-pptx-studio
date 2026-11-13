/**
 * dsh-pptx-studio — client half (classic script consumed by the Web carrier).
 *
 * Two surfaces:
 *   1. a right-side canvas tab (`sidebarRightTabs` + `sidebar.right.pane.tab`)
 *      that draws the deck slide by slide over SSE and lets the user select an
 *      element / drag a box and send one structured note back to the session;
 *   2. a tool-call card for `pptx_studio_open` that opens that canvas.
 *
 * The tab body and the card are session-scope slots, so they receive the
 * standard `inputActions` face: sending feedback uses `setDraft` + `submit`,
 * exactly like the user pressing Enter (the host route is only a fallback that
 * injects the message when no composer is available).
 * @module dsh-pptx-studio/client
 */
window.__ModuleLoader__.load({
  id: 'dsh-pptx-studio',
  factory(require) {
    const React = require('react')
    const h = React.createElement
    const { useCallback, useEffect, useMemo, useRef, useState } = React

    /** Locale namespace. */
    const NS = 'pptxStudio'
    /** Tab definition id; also the key its body and title register under. */
    const TAB_ID = 'dsh-pptx-studio/canvas'
    /** Tab kind passed to `sidebarRight.openTab`. */
    const TAB_KIND = 'pptx-studio-canvas'
    /** Host route prefix (absolute path served by the Web carrier). */
    const API_PATH = 'dsh-pptx-studio/api'

    const zh = {
      'tab.title': 'PPTX 画布',
      'guide.title': 'PPTX 画布',
      'guide.description': '实时看 deck 被画出来，点元素提意见',
      'empty.title': '还没有画布',
      'empty.body': '在对话里让 Agent 调用 pptx_studio_open（可以带 .pptx 模板路径），这里就会开始逐页绘制。',
      'empty.refresh': '刷新画布列表',
      'state.loading': '正在加载画布…',
      'state.error': '加载失败',
      'status.idle': '等待写入',
      'status.drawing': '正在绘制第 {index} 页 / 共 {total} 页',
      'status.planned': '大纲已就位，共 {total} 页',
      'status.saved': '已导出 {name}',
      'status.follow': '跟随最新',
      'status.followOn': '跟随中',
      'actions.export': '导出 PPTX',
      'actions.refresh': '重新加载',
      'actions.close': '关闭',
      'actions.send': '发给 DSH',
      'actions.sending': '发送中…',
      'select.hint': '点选元素，或在画布上拖拽框选区域',
      'select.element': '已选元素',
      'select.region': '已选区域',
      'select.clear': '取消选择',
      'note.placeholder': '哪里不对？例如：标题太长压到副标题了，缩到 12 字以内',
      'note.chips': ['太挤', '字号太小', '文字压到了', '换配色', '删除这个元素'],
      'note.needText': '先写一句意见，或点一个快捷标签',
      'note.sent': '已发给 DSH：第 {slide} 页 {target}',
      'note.fallback': '已填入输入框，按 Enter 发送',
      'note.copy': '无法自动发送，请复制这段话发给 DSH：',
      'card.title': 'PPTX 画布已创建',
      'card.open': '在右侧查看绘制',
      'card.opened': '画布已打开',
      'card.noTab': '右侧栏暂不可用：先在某个会话里点一下再试',
      'card.slides': '{count} 页',
      'deck.picker': '画布',
      'live.on': '实时',
      'live.off': '离线',
      'saved.hint': '文件：{path}',
      'approx': '近似渲染',
    }
    const en = {
      'tab.title': 'PPTX canvas',
      'guide.title': 'PPTX canvas',
      'guide.description': 'Watch a deck being drawn; point at what is wrong',
      'empty.title': 'No canvas yet',
      'empty.body': 'Ask the agent to call pptx_studio_open (optionally with a .pptx template). The canvas will draw it page by page.',
      'empty.refresh': 'Refresh decks',
      'state.loading': 'Loading canvas…',
      'state.error': 'Failed to load',
      'status.idle': 'Waiting',
      'status.drawing': 'Drawing slide {index} / {total}',
      'status.planned': 'Outline ready, {total} slides',
      'status.saved': 'Exported {name}',
      'status.follow': 'Follow latest',
      'status.followOn': 'Following',
      'actions.export': 'Export PPTX',
      'actions.refresh': 'Reload',
      'actions.close': 'Close',
      'actions.send': 'Send to DSH',
      'actions.sending': 'Sending…',
      'select.hint': 'Click an element, or drag a box on the canvas',
      'select.element': 'Element',
      'select.region': 'Region',
      'select.clear': 'Clear',
      'note.placeholder': 'What is wrong? e.g. the title is too long and overlaps the subtitle',
      'note.chips': ['Too tight', 'Text too small', 'Overlapping', 'Recolor', 'Delete this element'],
      'note.needText': 'Write a note or pick a quick tag first',
      'note.sent': 'Sent to DSH: slide {slide}, {target}',
      'note.fallback': 'Filled into the composer — press Enter to send',
      'note.copy': 'Automatic send failed; copy this to DSH:',
      'card.title': 'PPTX canvas created',
      'card.open': 'Open the live canvas',
      'card.opened': 'Canvas opened',
      'card.noTab': 'Right sidebar unavailable: click inside a session first',
      'card.slides': '{count} slides',
      'deck.picker': 'Deck',
      'live.on': 'live',
      'live.off': 'offline',
      'saved.hint': 'File: {path}',
      'approx': 'approx',
    }

    /** Stylesheet: theme tokens with safe fallbacks, no hard-coded palette. */
    const CSS = `
.ps-root{display:flex;flex-direction:column;height:100%;min-height:0;gap:8px;padding:10px 10px 12px;box-sizing:border-box;
  color:var(--dsw-alias-label-primary,#0f172a);font:12px/1.45 system-ui,-apple-system,"Segoe UI","Microsoft YaHei",sans-serif}
.ps-bar{display:flex;align-items:center;gap:8px;min-height:24px}
.ps-bar .ps-title{font-weight:600;font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;flex:1;min-width:0}
.ps-bar .ps-meta{color:var(--dsw-alias-label-secondary,#64748b);font-size:11px;white-space:nowrap}
.ps-btn{appearance:none;border:1px solid var(--dsw-alias-border-l1,#e2e8f0);background:var(--dsw-alias-bg-layer-1,#fff);
  color:var(--dsw-alias-label-primary,#0f172a);border-radius:6px;padding:3px 9px;font:inherit;font-size:11px;cursor:pointer;white-space:nowrap}
.ps-btn:hover{background:var(--dsw-alias-bg-layer-2,#f1f5f9)}
.ps-btn:disabled{opacity:.5;cursor:default}
.ps-btn.primary{background:var(--dsw-alias-brand-primary,#2563eb);border-color:transparent;color:#fff}
.ps-btn.primary:hover{filter:brightness(1.05);background:var(--dsw-alias-brand-primary,#2563eb)}
.ps-thumbs{display:flex;gap:6px;overflow-x:auto;padding-bottom:2px;flex:0 0 auto}
.ps-thumb{position:relative;flex:0 0 auto;width:74px;aspect-ratio:16/9;border-radius:5px;border:1px solid var(--dsw-alias-border-l1,#e2e8f0);
  background:#fff;cursor:pointer;overflow:hidden;padding:0;transition:border-color .15s, box-shadow .15s}
.ps-thumb[data-active="true"]{border-color:var(--dsw-alias-brand-primary,#2563eb);box-shadow:0 0 0 2px color-mix(in srgb, var(--dsw-alias-brand-primary,#2563eb) 22%, transparent)}
.ps-thumb .n{position:absolute;left:3px;bottom:2px;font-size:9px;color:#64748b;background:rgba(255,255,255,.86);border-radius:3px;padding:0 3px}
.ps-thumb .chip{position:absolute;top:2px;right:2px;width:6px;height:6px;border-radius:50%;background:var(--dsw-alias-state-success-primary,#047857)}
.ps-stage{position:relative;flex:1 1 auto;min-height:0;display:flex;align-items:flex-start;justify-content:center;overflow:auto;
  background:var(--dsw-alias-bg-layer-2,#f1f5f9);border-radius:8px;padding:10px}
.ps-canvas{position:relative;background:#fff;box-shadow:0 1px 3px rgba(15,23,42,.14);overflow:hidden;flex:0 0 auto;cursor:crosshair;user-select:none}
.ps-slide{position:absolute;inset:0;animation:ps-slide-in .32s ease both}
.ps-el{position:absolute;box-sizing:border-box;overflow:hidden}
.ps-el[data-fresh="true"]{animation:ps-el-in .38s cubic-bezier(.2,.7,.3,1) both}
.ps-el[data-flash="true"]{animation:ps-flash .9s ease both;outline:1px solid var(--dsw-alias-brand-primary,#2563eb)}
.ps-el[data-type="text"]{white-space:pre-wrap;word-break:break-word}
.ps-el[data-type="rect"]{border-style:solid}
.ps-el[data-type="image"] img{width:100%;height:100%;object-fit:fill;display:block}
.ps-el[data-type="line"]{background:#e2e8f0}
.ps-el[data-type="table"] table{width:100%;height:100%;border-collapse:collapse;table-layout:fixed}
.ps-el[data-type="table"] td{border:1px solid #e2e8f0;padding:2px 4px;overflow:hidden}
.ps-el[data-type="chart"],.ps-el[data-type="group"]{background:#f8fafc;border:1px dashed #cbd5e1;color:#64748b;
  display:flex;align-items:center;justify-content:center;font-size:11px}
.ps-hit{position:absolute;inset:0;cursor:pointer}
.ps-sel{position:absolute;margin:-2px;border:2px solid var(--dsw-alias-brand-primary,#2563eb);border-radius:3px;pointer-events:none;
  box-shadow:0 0 0 1px rgba(255,255,255,.7) inset}
.ps-region{position:absolute;border:2px dashed var(--dsw-alias-brand-primary,#2563eb);background:rgba(37,99,235,.12);pointer-events:none}
.ps-foot{display:flex;flex-direction:column;gap:8px;flex:0 0 auto}
.ps-row{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.ps-status{font-size:11px;color:var(--dsw-alias-label-secondary,#64748b);flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.ps-dot{width:7px;height:7px;border-radius:50%;background:var(--dsw-alias-state-idle-primary,#94a3b8);flex:0 0 auto}
.ps-dot.live{background:var(--dsw-alias-state-success-primary,#047857)}
.ps-note{width:100%;box-sizing:border-box;resize:vertical;min-height:52px;border:1px solid var(--dsw-alias-border-l1,#e2e8f0);
  border-radius:6px;padding:6px 8px;font:inherit;font-size:12px;background:var(--dsw-alias-bg-layer-1,#fff);color:inherit}
.ps-chips{display:flex;gap:6px;flex-wrap:wrap}
.ps-chip{border:1px solid var(--dsw-alias-border-l1,#e2e8f0);background:var(--dsw-alias-bg-layer-1,#fff);border-radius:999px;
  padding:2px 9px;font-size:11px;cursor:pointer;color:var(--dsw-alias-label-secondary,#64748b)}
.ps-chip:hover{border-color:var(--dsw-alias-brand-primary,#2563eb);color:var(--dsw-alias-brand-primary,#2563eb)}
.ps-selecting{display:flex;align-items:center;gap:8px;font-size:11px;color:var(--dsw-alias-label-secondary,#64748b);
  background:var(--dsw-alias-bg-layer-2,#f1f5f9);border-radius:6px;padding:4px 8px}
.ps-selecting b{color:var(--dsw-alias-label-primary,#0f172a);font-weight:600}
.ps-empty{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:10px;height:100%;padding:24px;text-align:center;color:var(--dsw-alias-label-secondary,#64748b)}
.ps-empty h4{margin:0;font-size:13px;color:var(--dsw-alias-label-primary,#0f172a)}
.ps-empty p{margin:0;font-size:12px;max-width:280px}
.ps-card{display:flex;flex-direction:column;gap:8px;border:1px solid var(--dsw-alias-border-l1,#e2e8f0);border-radius:8px;
  padding:10px 12px;margin:6px 0;background:var(--dsw-alias-bg-layer-1,#fff)}
.ps-card .hd{display:flex;align-items:center;gap:8px;font-size:12px;font-weight:600}
.ps-card .sub{font-size:11px;color:var(--dsw-alias-label-secondary,#64748b)}
.ps-card .badge{font-size:10px;border-radius:999px;padding:1px 7px;background:var(--dsw-alias-bg-layer-2,#f1f5f9);
  color:var(--dsw-alias-label-secondary,#64748b)}
.ps-err{color:var(--dsw-alias-state-error-primary,#b91c1c);font-size:11px}
@keyframes ps-slide-in{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}
@keyframes ps-el-in{from{opacity:0;transform:scale(.985)}to{opacity:1;transform:none}}
@keyframes ps-flash{0%{box-shadow:0 0 0 0 rgba(37,99,235,.55)}70%{box-shadow:0 0 0 8px rgba(37,99,235,0)}100%{box-shadow:0 0 0 0 rgba(37,99,235,0)}}
`

    // ── helpers ───────────────────────────────────────────────────────────────

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
     * JSON fetch with timeout.
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
        const data = text ? JSON.parse(text) : {}
        if (!response.ok) throw new Error(data?.message || data?.code || `HTTP ${response.status}`)
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
     * Turn an element box (inches) into percentage styles inside the slide.
     * @param {number[]} box - `[x, y, w, h]` in inches.
     * @param {number} slideW - slide width.
     * @param {number} slideH - slide height.
     * @returns {Record<string, string>} CSS position.
     */
    function boxStyle(box, slideW, slideH) {
      const [x, y, w, hh] = box
      return {
        left: `${(x / slideW) * 100}%`,
        top: `${(y / slideH) * 100}%`,
        width: `${(w / slideW) * 100}%`,
        height: `${(hh / slideH) * 100}%`,
      }
    }

    /**
     * Build one element node.
     * @param {Record<string, any>} el - element model.
     * @param {{slideW: number, slideH: number, scale: number, fresh: boolean, flash: boolean, onPick: (el: any) => void}} ctx - render context.
     * @returns {import('react').ReactNode} node.
     */
    function renderElement(el, ctx) {
      const style = el.style || {}
      const base = boxStyle(el.box, ctx.slideW, ctx.slideH)
      const fontSize = Math.max(4, (Number(style.size) || 18) * ctx.scale)
      const color = style.color || '#0f172a'
      /** @type {Record<string, any>} */
      const css = { ...base, fontSize, color }
      if (style.bold) css.fontWeight = 700
      if (style.align) css.textAlign = style.align
      if (el.fill) css.background = el.fill
      if (el.radius) css.borderRadius = `${Math.max(2, Math.min(base.width, base.height) * 0.0)}px`

      let body = null
      const type = el.type
      if (type === 'text') {
        const anchor = style.anchor || 'top'
        const paragraphs = Array.isArray(el.paragraphs) && el.paragraphs.length > 0
          ? el.paragraphs
          : [{ text: el.text || '' }]
        body = h('div', {
          style: {
            width: '100%', height: '100%', display: 'flex', flexDirection: 'column',
            justifyContent: anchor === 'middle' ? 'center' : anchor === 'bottom' ? 'flex-end' : 'flex-start',
          },
        }, paragraphs.map((para, i) => h('div', {
          key: `p${i}`,
          style: {
            fontSize: para.size ? Math.max(4, Number(para.size) * ctx.scale) : fontSize,
            fontWeight: para.bold ? 700 : (style.bold ? 700 : 400),
            color: para.color || color,
            textAlign: para.align || style.align || 'left',
            lineHeight: style.lineSpacing || 1.25,
            whiteSpace: 'pre-wrap',
            wordBreak: 'break-word',
          },
        }, String(para.text ?? ''))))
      } else if (type === 'rect') {
        css.border = `${Math.max(1, Number(el.lineWidth || 1) * ctx.scale)}px solid ${el.line || 'transparent'}`
        if (el.radius) {
          css.borderRadius = `${Math.max(2, (Number(el.radius) || 0.06) * 24 * ctx.scale)}px`
        }
        body = el.text
          ? h('div', {
              style: {
                width: '100%', height: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center',
                textAlign: style.align || 'center', padding: `${2 * ctx.scale}px ${6 * ctx.scale}px`,
                whiteSpace: 'pre-wrap', wordBreak: 'break-word', fontSize, color,
                fontWeight: style.bold ? 700 : 400,
              },
            }, String(el.text))
          : null
      } else if (type === 'image') {
        const src = typeof el.src === 'string'
          ? (/^https?:|^data:/.test(el.src) ? el.src : `${API_PATH}/asset?name=${encodeURIComponent(el.src.split(/[\\/]/).pop() || '')}`)
          : ''
        body = h('img', { src, alt: el.name || 'image', draggable: false })
      } else if (type === 'line') {
        css.background = el.line || el.fill || '#e2e8f0'
      } else if (type === 'table') {
        const rows = el.table?.rows || []
        body = h('table', null, h('tbody', null, rows.map((row, r) => h('tr', { key: `r${r}` },
          row.map((cell, c) => h('td', { key: `c${c}`, style: { fontSize, color, fontWeight: r === 0 ? 600 : 400 } }, String(cell)))))))
      } else {
        body = h('div', null, `${el.type}${el.text ? ` · ${el.text}` : ''}`)
      }

      return h('div', {
        key: el.id,
        className: 'ps-el',
        'data-type': el.type,
        'data-id': el.id,
        'data-fresh': ctx.fresh ? 'true' : 'false',
        'data-flash': ctx.flash ? 'true' : 'false',
        style: css,
      }, body, h('div', {
        className: 'ps-hit',
        onClick: (event) => {
          event.stopPropagation()
          ctx.onPick(el)
        },
      }))
    }

    // ── canvas tab body ───────────────────────────────────────────────────────

    /**
     * Right-side canvas body.
     * @param {any} props - session-scope slot props.
     * @returns {import('react').ReactNode} node.
     */
    function StudioBody(props) {
      const t = props.t ?? ((key, values) => fmt(key, values))
      const sessionId = props.sessionId
      const tabInfo = typeof props.useTabInfo === 'function' ? props.useTabInfo() : null
      const tab = tabInfo?.tab
      const params = tab?.navigation?.params || {}

      const [decks, setDecks] = useState([])
      const [deckId, setDeckId] = useState(typeof params.deckId === 'string' ? params.deckId : '')
      const [deck, setDeck] = useState(null)
      const [phase, setPhase] = useState({ phase: 'idle', message: '', slideIndex: 0, total: 0 })
      const [saved, setSaved] = useState(null)
      const [error, setError] = useState('')
      const [loading, setLoading] = useState(true)
      const [live, setLive] = useState(false)
      const [follow, setFollow] = useState(true)
      const [slideIndex, setSlideIndex] = useState(0)
      const [selection, setSelection] = useState(null) // {element} | {region:[x,y,w,h]}
      const [note, setNote] = useState('')
      const [sending, setSending] = useState(false)
      const [feedbackHint, setFeedbackHint] = useState('')
      const [freshEls, setFreshEls] = useState({})
      const [flashEls, setFlashEls] = useState({})
      const [scale, setScale] = useState(0.5)
      const [drag, setDrag] = useState(null)
      const stageRef = useRef(null)
      const canvasRef = useRef(null)
      const freshTimers = useRef(new Map())

      const slides = deck?.slides || []
      const size = deck?.size || { w: 13.333, h: 7.5 }
      const current = slides[Math.min(slideIndex, Math.max(0, slides.length - 1))] || null

      const markFresh = useCallback((ids, flash = false) => {
        if (!ids || ids.length === 0) return
        const stamp = Date.now()
        const next = {}
        for (const id of ids) next[id] = stamp
        if (flash) setFlashEls((prev) => ({ ...prev, ...next }))
        else setFreshEls((prev) => ({ ...prev, ...next }))
        for (const id of ids) {
          const key = `${flash ? 'f' : 'n'}:${id}`
          const existing = freshTimers.current.get(key)
          if (existing) clearTimeout(existing)
          const timer = setTimeout(() => {
            freshTimers.current.delete(key)
            if (flash) setFlashEls((prev) => { const copy = { ...prev }; delete copy[id]; return copy })
            else setFreshEls((prev) => { const copy = { ...prev }; delete copy[id]; return copy })
          }, flash ? 1000 : 700)
          freshTimers.current.set(key, timer)
        }
      }, [])

      const loadDecks = useCallback(async () => {
        try {
          const data = await request(`decks?sessionId=${encodeURIComponent(sessionId || '')}`)
          const list = Array.isArray(data.decks) ? data.decks : []
          setDecks(list)
          if (deckId === '' && list.length > 0) setDeckId(list[0].id)
          return list
        } catch (err) {
          setError(err instanceof Error ? err.message : String(err))
          return []
        }
      }, [sessionId, deckId])

      const loadState = useCallback(async (id) => {
        if (!id) return
        try {
          setLoading(true)
          const data = await request(`state?deck=${encodeURIComponent(id)}`)
          setDeck(data.deck || null)
          setError('')
        } catch (err) {
          setError(err instanceof Error ? err.message : String(err))
        } finally {
          setLoading(false)
        }
      }, [])

      // Initial load: decks, then the selected deck state.
      useEffect(() => {
        let alive = true
        ;(async () => {
          const list = await loadDecks()
          if (!alive) return
          const target = deckId || list[0]?.id
          if (target) await loadState(target)
          else setLoading(false)
        })()
        return () => { alive = false }
        // eslint-disable-next-line react-hooks/exhaustive-deps
      }, [sessionId])

      useEffect(() => { if (deckId) void loadState(deckId) }, [deckId, loadState])

      // SSE subscription for the selected deck.
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
          if (event.type === 'deck') {
            setDeck(event.deck)
            setLive(true)
            return
          }
          if (event.type === 'slide') {
            setDeck((prev) => {
              if (!prev) return prev
              const next = { ...prev, slides: [...(prev.slides || [])] }
              const index = Math.max(0, Math.min(event.index, next.slides.length))
              next.slides[index] = event.slide
              return next
            })
            markFresh((event.slide?.elements || []).map((el) => el.id))
            if (follow) setSlideIndex(event.index)
            return
          }
          if (event.type === 'element') {
            setDeck((prev) => {
              if (!prev) return prev
              const next = { ...prev, slides: (prev.slides || []).map((slide) => {
                if (slide.id !== event.slideId) return slide
                const elements = [...(slide.elements || [])]
                if (event.removed) return { ...slide, elements: elements.filter((el) => el.id !== event.elementId && el.id !== event.element?.id) }
                const index = elements.findIndex((el) => el.id === (event.elementId || event.element?.id))
                if (index >= 0) elements[index] = event.element
                else if (event.element) elements.push(event.element)
                return { ...slide, elements }
              }) }
              return next
            })
            markFresh([event.element?.id || event.elementId], true)
            return
          }
          if (event.type === 'status') {
            setPhase({ phase: event.phase, message: event.message || '', slideIndex: event.slideIndex || 0, total: event.total || 0 })
            if (follow && event.slideIndex) setSlideIndex(Math.max(0, event.slideIndex - 1))
            return
          }
          if (event.type === 'saved') {
            setSaved(event)
            setPhase((prev) => ({ ...prev, phase: 'saved', message: event.path || '' }))
            return
          }
          if (event.type === 'feedback') {
            setFeedbackHint(fmt(t('note.sent'), {
              slide: (event.slideIndex || 0) + 1,
              target: event.elementId || t('select.region'),
            }))
            return
          }
        }
        return () => { try { source.close() } catch { /* ignore */ } }
      }, [deckId, follow, markFresh, t])

      // Measure the canvas so 1pt on the slide maps to the right pixel size.
      useEffect(() => {
        const element = stageRef.current
        if (!element || typeof ResizeObserver !== 'function') return undefined
        const observer = new ResizeObserver((entries) => {
          for (const entry of entries) {
            const width = entry.contentRect.width - 20
            if (width > 40) setScale(width / (Number(size.w) * 72))
          }
        })
        observer.observe(element)
        return () => observer.disconnect()
      }, [size.w])

      const slidePoint = useCallback((clientX, clientY) => {
        const canvas = canvasRef.current
        if (!canvas) return null
        const rect = canvas.getBoundingClientRect()
        const x = ((clientX - rect.left) / rect.width) * Number(size.w)
        const y = ((clientY - rect.top) / rect.height) * Number(size.h)
        return [Math.max(0, Math.min(size.w, x)), Math.max(0, Math.min(size.h, y))]
      }, [size.h, size.w])

      const onCanvasClick = useCallback((event) => {
        if (event.target !== event.currentTarget) return
        setSelection(null)
      }, [])

      const onPointerDown = useCallback((event) => {
        if (event.button !== 0) return
        const point = slidePoint(event.clientX, event.clientY)
        if (!point) return
        setDrag({ start: point, current: point })
        event.currentTarget.setPointerCapture?.(event.pointerId)
      }, [slidePoint])

      const onPointerMove = useCallback((event) => {
        if (!drag) return
        const point = slidePoint(event.clientX, event.clientY)
        if (point) setDrag({ ...drag, current: point })
      }, [drag, slidePoint])

      const onPointerUp = useCallback(() => {
        if (!drag) return
        const [x1, y1] = drag.start
        const [x2, y2] = drag.current
        const region = [Math.min(x1, x2), Math.min(y1, y2), Math.abs(x2 - x1), Math.abs(y2 - y1)]
        setDrag(null)
        if (region[2] > 0.12 && region[3] > 0.12) setSelection({ region })
      }, [drag])

      const describeSelection = useCallback(() => {
        if (!selection) return null
        if (selection.element) {
          const el = selection.element
          return {
            kind: 'element',
            label: `${el.type} · ${el.id}${el.text ? ` · ${String(el.text).split('\n')[0].slice(0, 24)}` : ''}`,
            elementId: el.id,
          }
        }
        return {
          kind: 'region',
          label: `${t('select.region')} [${selection.region.map((n) => Math.round(n * 100) / 100).join(', ')}]`,
          elementId: null,
        }
      }, [selection, t])

      const buildFeedbackText = useCallback(() => {
        const slideNo = slideIndex + 1
        const total = slides.length || 1
        const lines = ['[PPTX 画布反馈]', `deck=${deckId} slide=${slideNo}/${total}${current?.id ? ` (${current.id})` : ''}`]
        if (selection?.element) {
          const el = selection.element
          lines.push(`element=${el.id} type=${el.type}${el.text ? ` text=「${String(el.text).split('\n')[0].slice(0, 80)}」` : ''}`)
          if (Array.isArray(el.box)) lines.push(`位置：x=${el.box[0]}in y=${el.box[1]}in w=${el.box[2]}in h=${el.box[3]}in`)
          if (el.style) lines.push(`样式：${JSON.stringify(el.style).slice(0, 200)}`)
        } else if (selection?.region) {
          lines.push(`region=[${selection.region.map((n) => Math.round(n * 1000) / 1000).join(', ')}] (x,y,w,h in inches)`)
        }
        if (note.trim()) lines.push(`用户意见：${note.trim()}`)
        lines.push('请只改这一处（除非同一问题在别页重复），改完调用 pptx_studio_export 重新导出；不要重做整个 deck。')
        return lines.join('\n')
      }, [current, deckId, note, selection, slideIndex, slides.length])

      const sendFeedback = useCallback(async () => {
        if (!selection) {
          setFeedbackHint(t('note.needText'))
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
        try {
          // Preferred: the native composer (same path as the user pressing Enter).
          let native = false
          try {
            if (props.inputActions && typeof props.inputActions.setDraft === 'function') {
              props.inputActions.setDraft(text)
              if (typeof props.inputActions.persistDraft === 'function') props.inputActions.persistDraft()
              if (typeof props.inputActions.submit === 'function') {
                props.inputActions.submit()
                native = true
              }
              const input = document.querySelector('[data-composer-input]')
              if (input && typeof input.focus === 'function') input.focus({ preventScroll: true })
            }
          } catch {
            native = false
          }
          // Always record the feedback on the canvas; only ask the host to inject
          // when the native composer could not submit.
          const payload = {
            deckId,
            sessionId,
            slideId: current?.id,
            elementId: selection.element?.id ?? null,
            bbox: selection.region ?? selection.element?.box ?? null,
            note: note.trim(),
            text,
            inject: !native,
          }
          const result = await request('feedback', { method: 'POST', body: payload })
          setNote('')
          setFeedbackHint(native ? t('note.fallback') : (result?.delivered ? t('note.sent') : t('note.copy')))
          if (result?.delivered === true) {
            setPhase((prev) => ({ ...prev, phase: 'feedback', message: '' }))
          }
        } catch (err) {
          setFeedbackHint(`${t('state.error')}: ${err instanceof Error ? err.message : String(err)}`)
        } finally {
          setSending(false)
        }
      }, [buildFeedbackText, current, deckId, note, props.inputActions, selection, sending, sessionId, t])

      const exportDeck = useCallback(async () => {
        if (!deckId) return
        setError('')
        try {
          const data = await request('export', { method: 'POST', body: { deckId }, timeoutMs: 190000 })
          if (data?.value) setSaved(data.value)
        } catch (err) {
          setError(err instanceof Error ? err.message : String(err))
        }
      }, [deckId])

      const closeTab = useCallback(() => {
        try { tab?.actions?.close?.() } catch { /* ignore */ }
      }, [tab])

      const selectionInfo = describeSelection()
      const dragRegion = drag
        ? [
            Math.min(drag.start[0], drag.current[0]),
            Math.min(drag.start[1], drag.current[1]),
            Math.abs(drag.current[0] - drag.start[0]),
            Math.abs(drag.current[1] - drag.start[1]),
          ]
        : null

      if (loading && !deck) {
        return h('div', { className: 'ps-root' }, h('div', { className: 'ps-empty' }, t('state.loading')))
      }

      if (!deck) {
        return h('div', { className: 'ps-root' }, h('div', { className: 'ps-empty' },
          h('h4', null, t('empty.title')),
          h('p', null, t('empty.body')),
          h('div', { className: 'ps-row' },
            h('button', { type: 'button', className: 'ps-btn', onClick: () => { void loadDecks() } }, t('empty.refresh')),
            tab ? h('button', { type: 'button', className: 'ps-btn', onClick: closeTab }, t('actions.close')) : null,
          ),
          error ? h('div', { className: 'ps-err' }, `${t('state.error')}: ${error}`) : null,
        ))
      }

      const thumb = (slide, index) => h('button', {
        key: slide.id,
        type: 'button',
        className: 'ps-thumb',
        'data-active': String(index === slideIndex),
        title: slide.title || `第 ${index + 1} 页`,
        onClick: () => { setSlideIndex(index); setFollow(false); setSelection(null) },
      }, h('span', { className: 'n' }, String(index + 1)), index >= slides.length - 1 ? h('span', { className: 'chip' }) : null)

      return h('div', { className: 'ps-root' },
        h('div', { className: 'ps-bar' },
          h('span', { className: 'ps-dot' + (live ? ' live' : '') }),
          h('span', { className: 'ps-title' }, deck.title || t('tab.title')),
          h('span', { className: 'ps-meta' }, fmt('{n} 页', { n: slides.length })),
          decks.length > 1
            ? h('select', {
                className: 'ps-btn',
                value: deckId,
                onChange: (event) => { setDeckId(event.target.value); setSlideIndex(0); setSelection(null) },
                title: t('deck.picker'),
              }, decks.map((item) => h('option', { key: item.id, value: item.id }, `${item.title || item.id} (${item.slides})`)))
            : null,
          h('button', { type: 'button', className: 'ps-btn primary', onClick: () => { void exportDeck() } }, t('actions.export')),
        ),
        h('div', { className: 'ps-thumbs' }, slides.map(thumb)),
        h('div', { className: 'ps-stage', ref: stageRef },
          h('div', {
            className: 'ps-canvas',
            ref: canvasRef,
            style: {
              width: `${Number(size.w) * 72 * scale}px`,
              maxWidth: '100%',
              aspectRatio: `${size.w} / ${size.h}`,
            },
            onClick: onCanvasClick,
            onPointerDown,
            onPointerMove,
            onPointerUp,
          },
            current
              ? h('div', { className: 'ps-slide', key: current.id },
                  current.elements.map((el) => renderElement(el, {
                    slideW: Number(size.w),
                    slideH: Number(size.h),
                    scale,
                    fresh: Boolean(freshEls[el.id]),
                    flash: Boolean(flashEls[el.id]),
                    onPick: (picked) => setSelection({ element: picked }),
                  })),
                  selectionInfo?.kind === 'element' && current.elements.some((el) => el.id === selection.element?.id)
                    ? h('div', { className: 'ps-sel', style: boxStyle(current.elements.find((el) => el.id === selection.element.id).box, Number(size.w), Number(size.h)) })
                    : null,
                  dragRegion
                    ? h('div', { className: 'ps-region', style: boxStyle(dragRegion, Number(size.w), Number(size.h)) })
                    : null,
                  selectionInfo?.kind === 'region'
                    ? h('div', { className: 'ps-region', style: boxStyle(selection.region, Number(size.w), Number(size.h)) })
                    : null,
                )
              : h('div', { className: 'ps-empty' },
                  h('p', null, t('empty.body')),
                ),
          ),
        ),
        h('div', { className: 'ps-foot' },
          h('div', { className: 'ps-row' },
            h('span', { className: 'ps-status' },
              phase.phase === 'drawing'
                ? fmt(t('status.drawing'), { index: phase.slideIndex || slideIndex + 1, total: phase.total || slides.length })
                : phase.phase === 'planned'
                  ? fmt(t('status.planned'), { total: phase.total || slides.length })
                  : phase.phase === 'saved'
                    ? fmt(t('status.saved'), { name: String(phase.message || '').split(/[\\/]/).pop() || '' })
                    : saved?.path
                      ? fmt(t('saved.hint'), { path: saved.path })
                      : t('status.idle'),
            ),
            h('label', { className: 'ps-row', style: { gap: 4, fontSize: 11, color: 'var(--dsw-alias-label-secondary,#64748b)' } },
              h('input', { type: 'checkbox', checked: follow, onChange: (event) => setFollow(event.target.checked) }),
              t('status.follow'),
            ),
            h('button', { type: 'button', className: 'ps-btn', onClick: () => { void loadState(deckId) } }, t('actions.refresh')),
            tab ? h('button', { type: 'button', className: 'ps-btn', onClick: closeTab }, t('actions.close')) : null,
          ),
          selectionInfo
            ? h('div', { className: 'ps-selecting' },
                h('span', null, h('b', null, selectionInfo.kind === 'element' ? t('select.element') : t('select.region')), ` ${selectionInfo.label}`),
                h('button', { type: 'button', className: 'ps-btn', onClick: () => setSelection(null) }, t('select.clear')),
              )
            : h('div', { className: 'ps-selecting' }, t('select.hint')),
          selectionInfo
            ? h('div', null,
                h('textarea', {
                  className: 'ps-note',
                  value: note,
                  placeholder: t('note.placeholder'),
                  onChange: (event) => setNote(event.target.value),
                }),
                h('div', { className: 'ps-chips' },
                  (t('note.chips') || []).map((chip) => h('button', {
                    key: chip,
                    type: 'button',
                    className: 'ps-chip',
                    onClick: () => setNote((prev) => (prev.trim() === '' ? chip : `${prev}；${chip}`)),
                  }, chip)),
                ),
                h('div', { className: 'ps-row', style: { marginTop: 6 } },
                  h('button', {
                    type: 'button',
                    className: 'ps-btn primary',
                    disabled: sending,
                    onClick: () => { void sendFeedback() },
                  }, sending ? t('actions.sending') : t('actions.send')),
                  feedbackHint ? h('span', { className: 'ps-status' }, feedbackHint) : null,
                ),
              )
            : null,
          error ? h('div', { className: 'ps-err' }, `${t('state.error')}: ${error}`) : null,
        ),
      )
    }

    /**
     * Tab title chip.
     * @param {any} props - slot props.
     * @returns {import('react').ReactNode} node.
     */
    function TabTitle(props) {
      const t = props.t ?? ((key) => key)
      return h('span', null, t('tab.title'))
    }

    /**
     * Tool-call card for `pptx_studio_open`: one-click entry to the canvas.
     * @param {any} props - tool view props.
     * @returns {import('react').ReactNode} node.
     */
    function OpenCard(props) {
      const t = props.t ?? ((key, values) => fmt(key, values))
      const phase = props.phase
      const title = props.block?.args?.text?.('title') || ''
      const template = props.block?.args?.text?.('template') || ''
      const [opened, setOpened] = useState(false)
      const [failed, setFailed] = useState('')
      const autoRef = useRef(false)

      const open = useCallback(() => {
        try {
          const ok = props.openStudio?.() !== false
          setOpened(ok !== false)
          if (ok === false) setFailed(t('card.noTab'))
        } catch {
          setFailed(t('card.noTab'))
        }
      }, [props, t])

      useEffect(() => {
        if (autoRef.current) return
        if (phase !== 'result') return
        autoRef.current = true
        open()
      }, [open, phase])

      return h('div', { className: 'ps-card', 'data-pptx-studio-card': props.callId || '' },
        h('div', { className: 'hd' }, t('card.title'), h('span', { className: 'badge' }, phase || '')),
        h('div', { className: 'sub' }, title || '(untitled)', template ? ` · template: ${template}` : ''),
        h('div', { className: 'ps-row' },
          h('button', { type: 'button', className: 'ps-btn primary', onClick: open }, t('card.open')),
          opened ? h('span', { className: 'sub' }, t('card.opened')) : null,
          failed ? h('span', { className: 'ps-err' }, failed) : null,
        ),
      )
    }

    return {
      inject: ['slots', 'locale', 'sidebarRightTabs', 'sidebarRight'],
      apply(ctx) {
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
          StudioBody,
        ))
        ctx.slots.inject('sidebar.right.pane.tab.title', () => ctx.slots.register(
          { name: 'sidebar.right.pane.tab.title', key: TAB_ID, locale: NS },
          TabTitle,
        ))

        /**
         * Open the canvas for one session (called from the tool card).
         * @param {string} sessionId - session that owns this card.
         * @returns {(params?: Record<string, unknown>) => boolean} opener.
         */
        const openerFor = (sessionId) => (params) => {
          const controller = ctx.get('sidebarRight')
          if (!controller) return false
          try {
            const mounted = controller.mounted?.getSnapshot?.()
            if (mounted === undefined) return false
            if (typeof controller.openTabIn === 'function' && mounted !== sessionId) {
              controller.openTabIn(sessionId, TAB_KIND, params ? { params } : {})
              return true
            }
            if (mounted !== sessionId) return false
            controller.openTab(TAB_KIND, params ? { params } : {})
            return true
          } catch {
            return false
          }
        }

        // ③ tool card for `pptx_studio_open` (auto-opens the canvas on result).
        ctx.slots.inject('tool.call.toolview', () => ctx.slots.register({
          name: 'tool.call.toolview',
          key: 'pptx_studio_open',
          locale: NS,
          inject: (sessionId) => ({ openStudio: openerFor(sessionId) }),
        }, OpenCard))
      },
    }
  },
})
