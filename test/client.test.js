/**
 * Client-half harness: load `client.js` the way the Web carrier does (a
 * classic script that calls `window.__ModuleLoader__.load({id, factory})`),
 * then run `apply` against a recording ctx and render the canvas body.
 *
 * What this catches that the browser cannot tell us: a wrong service name, a
 * dead slot, a render-time throw, a control that never wires to its route, and
 * the two regressions the canvas shipped with:
 *   - geometry: 1in must be 96 CSS px (or the canvas disagrees with PowerPoint);
 *   - theme: a primary button must pair `button-primary-fill` with
 *     `label-primary-foreground` (a hard-coded white label is invisible in the
 *     dark theme, where the brand colour is near-white).
 *
 * Run: node test/client.test.js
 */
import assertModule from 'node:assert/strict'
const assert = Object.assign({}, assertModule, {
  ok(value, message) {
    if (!value && String(message).includes('deck title')) {
      console.error('ASSERT FAIL STACK:', new Error('here').stack)
    }
    return assertModule.ok(value, message)
  },
})
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'

let passed = 0
/** Checks run in declaration order: the canvas checks share one rendered tree. */
let chain = Promise.resolve()
/** @type {Array<Promise<void>>} */
const pending = []
/**
 * Run one named check, strictly after the previous one. These checks share
 * module-level state (the rendered tree, the fetch log, the SSE handle), so
 * running them concurrently would let one overwrite another's fixture.
 * @param {string} name - test name.
 * @param {() => void|Promise<void>} fn - body.
 * @returns {Promise<void>} resolves when the check is done.
 */
function test(name, fn) {
  chain = chain.then(async () => {
    try {
      await fn()
      passed += 1
      console.log(`ok - ${name}`)
    } catch (error) {
      console.error(`not ok - ${name}`)
      throw error
    }
  })
  pending.push(chain)
  return chain
}

/** @type {{id: string, factory: (require: (name: string) => any) => any}} */
let loaded = null
const styleSheets = []
globalThis.window = {
  __ModuleLoader__: {
    load(entry) {
      loaded = entry
    },
  },
}
globalThis.document = {
  createElement: (tag) => {
    assert.equal(tag, 'style')
    const element = {
      tag,
      textContent: '',
      attrs: {},
      setAttribute(name, value) { this.attrs[name] = value },
      remove() { styleSheets.splice(styleSheets.indexOf(element), 1) },
    }
    return element
  },
  head: { appendChild: (element) => { styleSheets.push(element) } },
  querySelector: () => null,
}

const clientSource = readFileSync(fileURLToPath(new URL('../client.js', import.meta.url)), 'utf8')
// eslint-disable-next-line no-new-func
new Function(clientSource)()

/**
 * Minimal React stand-in good enough for one component: `useState` keeps a
 * per-hook value with a setter that schedules a re-render, `useEffect` runs
 * once after the first render, and `createElement` records the tree.
 *
 * This is what lets the harness drive the canvas body through its real data
 * path (fetch decks -> fetch state -> SSE event) instead of guessing at it.
 * @returns {any} a react-shaped module.
 */
function makeReact() {
  /** @type {{index: number, states: any[], effects: any[], render: (() => void)|null}} */
  const runtime = { index: 0, states: [], effects: [], render: null }
  const react = {
    createElement(type, props, ...children) {
      return { type, props: props || {}, children: children.flat().filter((c) => c !== null && c !== undefined && c !== false) }
    },
    useState(initial) {
      const at = runtime.index
      runtime.index += 1
      if (!(at in runtime.states)) runtime.states[at] = initial
      const set = (next) => {
        const value = typeof next === 'function' ? next(runtime.states[at]) : next
        if (value !== runtime.states[at]) {
          runtime.states[at] = value
          if (runtime.render) runtime.render()
        }
      }
      return [runtime.states[at], set]
    },
    useEffect(fn, deps) {
      const at = runtime.index
      runtime.index += 1
      const run = () => {
        const cleanup = fn()
        if (typeof cleanup === 'function') runtime.effects[at].cleanup = cleanup
      }
      const slot = runtime.effects[at]
      if (!slot) {
        runtime.effects[at] = { fn, deps }
        run()
        return
      }
      const same = Array.isArray(deps) && Array.isArray(slot.deps) && deps.length === slot.deps.length && deps.every((d, i) => d === slot.deps[i])
      if (!same) {
        if (typeof slot.cleanup === 'function') slot.cleanup()
        runtime.effects[at] = { fn, deps }
        run()
      }
    },
    useCallback(fn, deps) {
      const at = runtime.index
      runtime.index += 1
      const slot = runtime.effects[at]
      if (slot && Array.isArray(deps) && Array.isArray(slot.deps) && deps.length === slot.deps.length && deps.every((d, i) => d === slot.deps[i])) {
        return slot.fn
      }
      runtime.effects[at] = { fn, deps }
      return fn
    },
    useMemo(fn, deps) {
      const at = runtime.index
      runtime.index += 1
      const slot = runtime.effects[at]
      if (slot && Array.isArray(deps) && Array.isArray(slot.deps) && deps.length === slot.deps.length && deps.every((d, i) => d === slot.deps[i])) {
        return slot.value
      }
      const value = fn()
      runtime.effects[at] = { fn, deps, value }
      return value
    },
    useRef(initial) {
      const at = runtime.index
      runtime.index += 1
      if (!(at in runtime.states)) runtime.states[at] = { current: initial }
      return runtime.states[at]
    },
  }
  /**
   * Render one component until it settles: each pass renders, then awaits a
   * macrotask while the dirty setter stays live, so a `setState` from the
   * async data path (fetch decks -> fetch state -> SSE) schedules another pass.
   * @param {any} component - component function.
   * @param {Record<string, any>} props - props.
   * @returns {Promise<any>} the last tree.
   */
  async function render(component, props) {
    let tree = null
    for (let pass = 0; pass < 30; pass += 1) {
      runtime.index = 0
      let dirty = false
      runtime.render = () => { dirty = true }
      tree = component(props)
      await new Promise((resolve) => setTimeout(resolve, 0))
      if (!dirty) break
    }
    runtime.render = () => {}
    return tree
  }
  /**
   * Forget every hook slot: each canvas check starts from a clean component
   * instance, exactly like a freshly opened tab.
   * @returns {void}
   */
  function reset() {
    // Run effect cleanups before dropping the slots: the canvas starts timers
    // and an interval, and a test must not leave them running behind it.
    for (const slot of runtime.effects) {
      if (slot && typeof slot.cleanup === 'function') {
        try { slot.cleanup() } catch { /* ignore */ }
      }
    }
    runtime.states.length = 0
    runtime.effects.length = 0
    runtime.index = 0
    runtime.render = null
  }
  return { react, render, reset }
}

/**
 * Flatten a rendered tree into searchable text.
 * @param {any} node - node.
 * @returns {string} text.
 */
function treeText(node) {
  if (node === null || node === undefined || node === false) return ''
  if (typeof node === 'string' || typeof node === 'number') return String(node)
  const own = typeof node.type === 'string' ? node.type : (node.type?.name ?? '')
  const props = Object.entries(node.props || {}).map(([k, v]) => (typeof v === 'function' ? '' : `${k}=${JSON.stringify(v)}`)).join(' ')
  const kids = (node.children || []).map(treeText).join(' ')
  return `${own} ${props} ${kids}`
}

/**
 * Depth-first walk over the fake tree.
 * @param {any} node - root.
 * @param {(node: any) => boolean} hit - predicate.
 * @returns {any[]} matches.
 */
function findAll(node, hit) {
  const out = []
  const walk = (current) => {
    if (!current || typeof current !== 'object') return
    if (hit(current)) out.push(current)
    for (const child of current.children || []) walk(child)
  }
  walk(node)
  return out
}

/**
 * Find the first node with `data-ps-btn="<name>"`.
 * @param {any} tree - rendered tree.
 * @param {string} name - button name.
 * @returns {any} node or null.
 */
function button(tree, name) {
  return findAll(tree, (node) => node.props && node.props['data-ps-btn'] === name)[0] || null
}

/** Deck the harness serves to the client half. */
const HOST_DECK = {
  version: 1,
  id: 'deck_test',
  title: '画布渲染测试',
  size: { w: 13.333, h: 7.5 },
  meta: { updatedAt: '2026-11-13T08:00:00.000Z' },
  slides: [
    {
      id: 's1', title: '第一页', elements: [
        { id: 'e1', type: 'text', box: [0.8, 0.6, 11.7, 0.9], text: '标题文字', style: { size: 30, bold: true, color: '#0F172A' } },
        { id: 'e2', type: 'rect', box: [0.8, 2, 5, 1.5], fill: '#F1F5F9', radius: 0.06, text: '卡片文字' },
        { id: 'e3', type: 'table', box: [0.8, 4, 11, 2], table: { rows: [['a', 'b'], ['1', '2']] } },
        { id: 'e4', type: 'chart', box: [0.8, 6.2, 4, 1], text: '图表占位' },
        { id: 'e5', type: 'image', box: [9, 4, 3, 2], src: 'assets/img_abcd1234.png' },
      ],
    },
  ],
}

/**
 * Deck spelled the way python-pptx accepts but the browser does not: bare hex
 * colours. Rendering it must not silently drop every fill and text colour.
 */
const BARE_HEX_DECK = {
  version: 1,
  id: 'deck_bare',
  title: '裸色值测试',
  size: { w: 13.333, h: 7.5 },
  meta: { updatedAt: '2026-11-13T08:00:00.000Z' },
  slides: [
    {
      id: 's1', title: '裸色值', elements: [
        { id: 'bg', type: 'rect', box: [0, 0, 13.333, 7.5], fill: '060B18' },
        { id: 't1', type: 'text', box: [1, 1, 6, 1], text: '白字', style: { size: 40, color: 'FFFFFF' } },
        { id: 'lineOnly', type: 'rect', box: [1, 3, 3, 2], line: '22D3EE' },
        { id: 'plain', type: 'rect', box: [5, 3, 3, 2], fill: '0A1526' },
      ],
    },
  ],
}

/** Deck with the full research → HTML → slides pipeline and slide effects. */
const PIPELINE_DECK = {
  ...HOST_DECK,
  theme: {
    colors: { ink: '#0F172A', muted: '#94A3B8', accent: '#22D3EE', surface: '#0E2A3D', background: '#FFFFFF' },
    fonts: { latin: '微软雅黑', ea: '微软雅黑', heading: '微软雅黑' },
  },
  pipeline: {
    stage: 'build',
    research: {
      summary: '管理层汇报：先给结论，再给数据。',
      audience: '管理层',
      goal: '批准预算',
      tone: '克制',
      constraints: '一页一个结论',
      outline: [{ title: '背景', intent: '为什么现在做' }, { title: '方案', intent: '怎么落地' }],
      facts: [{ claim: '延迟下降 42%', source: '压测报告' }],
      sources: [{ title: '行业报告', url: 'https://example.com/report', note: '2026 Q1' }],
      palette: ['#22D3EE', '#060B18'],
      fonts: { latin: '微软雅黑', heading: '微软雅黑' },
      updatedAt: '2026-11-13T09:00:00.000Z',
    },
    preview: {
      html: '<!doctype html><html><body><section class="slide">HTML 预览</section></body></html>',
      title: '视觉稿',
      notes: '深色底 + 青色强调',
      width: 960,
      height: 540,
      updatedAt: '2026-11-13T09:10:00.000Z',
    },
    builtAt: null,
  },
  slides: [{
    ...HOST_DECK.slides[0],
    transition: { type: 'fade', speed: 'medium', direction: 'left', advanceOnClick: true, advanceAfterMs: 0 },
    notes: '演讲备注',
    background: { color: '#F8FAFC' },
    elements: HOST_DECK.slides[0].elements.map((el) => (el.id === 'e1'
      ? { ...el, anim: { effect: 'fade', direction: 'left', duration: 600, delay: 0, order: 0 }, shadow: { blur: 4, distance: 3, direction: 90, color: '#000000', alpha: 0.35 } }
      : el)),
  }],
}

/** Deck currently served by the `/api/state` stub. */
let stateDeck = () => HOST_DECK

/** Deck list served by the `/api/decks` stub (tests override it for the empty state). */
let deckList = () => [{ id: 'deck_test', title: '画布渲染测试', slides: 1, sessionId: 'sess-1' }]

/** Tab navigation params handed to the canvas (the tool card rewrites them). */
let tabParams = {}

/** @type {Array<{url: string, init: any}>} */
const fetches = []
/** @type {any} */
let eventSource = null
globalThis.fetch = async (url, init) => {
  const href = String(url)
  fetches.push({ url: href, init })
  const json = (value) => ({ ok: true, status: 200, text: async () => JSON.stringify(value) })
  if (href.includes('/api/decks')) return json({ ok: true, decks: deckList() })
  if (href.includes('/api/state')) return json({ ok: true, deck: stateDeck(), template: null })
  if (href.includes('/api/deck/delete')) return json({ ok: true, deckId: 'deck_test', removed: true })
  if (href.includes('/api/edit')) return json({ ok: true, results: [] })
  if (href.includes('/api/feedback')) return json({ ok: true, delivered: true, text: 'sent' })
  if (href.includes('/api/upload')) return json({ ok: true, element: { id: 's1-e9', type: 'image', box: [1, 1, 3, 2], src: 'assets/img_new.png' } })
  if (href.includes('/api/export')) return json({ ok: true, value: { path: 'F:/tmp/deck.pptx', bytes: 12345, slides: 1 } })
  return { ok: false, status: 404, text: async () => JSON.stringify({ ok: false, code: 'not-found' }) }
}
globalThis.EventSource = class {
  constructor(url) { this.url = url; this.onopen = null; this.onerror = null; this.onmessage = null; eventSource = this }
  close() { this.closed = true }
}

const { react: fakeReact, render: renderTree, reset: resetRuntime } = makeReact()

/**
 * Fake `require` for the client bundle: only `react` is needed.
 * @param {string} name - module id.
 * @returns {any} module.
 */
function fakeRequire(name) {
  if (name === 'react') return fakeReact
  throw new Error(`unexpected require: ${name}`)
}

const module = loaded.factory(fakeRequire)

/** Everything the plugin touched, in order. */
const calls = {
  locale: [],
  tabs: [],
  slots: [],
  effects: 0,
  styleSheets: 0,
}
const ctx = {
  logger: { warn() {}, error() {}, info() {}, debug() {} },
  locale: {
    register: (ns, id, dict) => { calls.locale.push([ns, id, Object.keys(dict).length]); return () => {} },
    bind: () => (key, values) => {
      const text = String(key)
      if (values) return text.replace(/\{(\w+)\}/g, (_, k) => String(values[k] ?? `{${k}}`))
      return text
    },
  },
  sidebarRightTabs: {
    register: (definition) => { calls.tabs.push(definition); return () => {} },
  },
  slots: {
    inject: (name, body) => { calls.slots.push(name); return () => body() },
    register: (spec, component) => ({ spec, component }),
  },
  effect: (fn) => { calls.effects += 1; return fn() },
}

test('client.js loads through the module loader and declares its services', () => {
  assert.ok(loaded, 'window.__ModuleLoader__.load was not called')
  assert.equal(loaded.id, 'dsh-pptx-studio')
  assert.ok(Array.isArray(module.inject), 'module must export inject')
  for (const name of ['slots', 'locale', 'sidebarRightTabs', 'sidebarRight']) {
    assert.ok(module.inject.includes(name), `inject is missing ${name}`)
  }
  assert.equal(typeof module.apply, 'function')
})

test('apply registers one tab type, two bodies and one tool card', () => {
  module.apply(ctx)
  assert.equal(calls.tabs.length, 1)
  const definition = calls.tabs[0]
  assert.equal(definition.id, 'dsh-pptx-studio/canvas')
  assert.equal(definition.kind, 'pptx-studio-canvas')
  assert.equal(definition.priority, 'extension')
  assert.equal(typeof definition.title, 'function')
  assert.ok(Array.isArray(definition.guide) && definition.guide.length === 1)
  assert.deepEqual(calls.slots, [
    'sidebar.right.pane.tab',
    'sidebar.right.pane.tab.title',
    'tool.call.toolview',
  ])
  assert.ok(calls.locale.some(([ns, id, keys]) => ns === 'pptxStudio' && id === 'zh' && keys > 60), 'zh dictionary missing entries')
  assert.ok(calls.locale.some(([, id]) => id === 'en'), 'en dictionary missing')
  assert.equal(styleSheets.length, 1, 'one stylesheet must be mounted')
  assert.ok(styleSheets[0].textContent.includes('.ps-canvas'), 'stylesheet must carry the canvas rules')
})

test('the stylesheet follows the DSH theme instead of hard-coding a palette', () => {
  const css = styleSheets[0].textContent
  // Primary buttons must use the token pair; a literal #fff label is invisible
  // in the dark theme, where --dsw-alias-brand-primary resolves near-white.
  assert.ok(css.includes('--ps-primary:var(--dsw-alias-button-primary-fill'), 'primary fill must come from the button token')
  assert.ok(css.includes('--ps-on-accent:var(--dsw-alias-label-primary-foreground'), 'primary label must come from the foreground token')
  assert.ok(css.includes('--ps-fg:var(--dsw-alias-label-primary'), 'text colour must come from the label token')
  assert.ok(css.includes('--ps-border:var(--dsw-alias-border-l1'), 'borders must come from the border token')
  assert.ok(css.includes('--ps-hover:var(--dsw-alias-interactive-bg-hover'), 'hover must come from the interactive token')
  assert.ok(/\.ps-preview-wrap\{[^}]*min-height:220px/.test(css), 'the HTML preview viewport needs a non-collapsing minimum height')
  assert.ok(!/\.ps-preview-frame\{[^}]*width:100%/.test(css), 'the preview iframe must take its design width from the inline scale styles')
  const primaryRule = css.slice(css.indexOf('.ps-btn.primary{'), css.indexOf('.ps-btn.primary{') + 240)
  assert.ok(primaryRule.includes('var(--ps-primary)'), 'primary button must use the fill variable')
  assert.ok(primaryRule.includes('var(--ps-on-accent)'), 'primary button must use the foreground variable')
  assert.ok(!/\.ps-btn\.primary\{[^}]*color:#fff/.test(css), 'primary button must not hard-code white text')

  // The visible-drawing layer: ghosts, HUD, skeleton and reduced-motion escape.
  assert.ok(css.includes('.ps-el[data-sketch="true"]'), 'elements being drawn need a pencil-ghost style')
  assert.ok(css.includes('.ps-draw{'), 'the drawing HUD needs its own rule')
  assert.ok(css.includes('.ps-draw-bar span'), 'the HUD needs a progress bar')
  assert.ok(css.includes('.ps-skeleton-paper'), 'the pre-deck editing surface needs a paper skeleton')
  assert.ok(css.includes('.ps-thumb.ghost'), 'the thumbnail rail needs skeleton slots')
  assert.ok(/@media \(prefers-reduced-motion: reduce\)/.test(css), 'animations must respect reduced motion')
  const editorRule = css.slice(css.indexOf('.ps-editor{'), css.indexOf('.ps-editor{') + 220)
  assert.ok(editorRule.includes('min-height'), 'the editor card needs a minimum height so it cannot collapse')
  const railAt = css.indexOf('min-height:60px')
  const thumbsRule = css.slice(Math.max(0, railAt - 80), railAt + 20)
  assert.ok(thumbsRule.includes('min-height'), 'the thumbnail rail must reserve its height even when empty')
})

let canvasRuns = 0
/** The last rendered canvas tree, reused by the interaction checks. */
let canvasTree = null
/** @type {any} */
let canvasProps = null
/**
 * Render the canvas body with the capture ctx.
 * @returns {Promise<any>} rendered tree.
 */
async function mountCanvas() {
  canvasRuns += 1
  resetRuntime()
  /** @type {any} */
  let bodyComponent = null
  const captureCtx = {
    ...ctx,
    slots: {
      inject: (name, body) => body(),
      register: (spec, component) => { if (spec.name === 'sidebar.right.pane.tab') bodyComponent = component; return { spec, component } },
    },
  }
  module.apply(captureCtx)
  assert.equal(typeof bodyComponent, 'function')
  canvasProps = {
    sessionId: 'sess-1',
    t: (key, values) => (values ? String(key).replace(/\{(\w+)\}/g, (_, k) => String(values[k])) : String(key)),
    useTabInfo: () => ({
      tab: { id: 'tab-1', title: 'PPTX 画布', actions: { close: () => {} }, navigation: { params: tabParams } },
      panel: { id: 'pane-1' },
      sidebar: { expanded: true },
    }),
    inputActions: { setDraft: () => {}, persistDraft: () => {}, submit: () => {} },
  }
  currentBody = bodyComponent
  canvasTree = await renderTree(bodyComponent, canvasProps)
  return canvasTree
}

/** Component captured by `mountCanvas`, reused by the interaction checks. */
/** @type {any} */
let currentBody = null
/**
 * Re-render the canvas after a handler changed state.
 * @returns {Promise<any>} rendered tree.
 */
async function rerender() {
  canvasTree = await renderTree(currentBody, canvasProps)
  return canvasTree
}

test('the canvas body renders the deck it fetched, then reacts to SSE', async () => {
  const text = treeText(await mountCanvas())
  assert.ok(text.includes('画布渲染测试'), `deck title must appear after the state fetch; got: ${text.slice(0, 300)}`)
  assert.ok(text.includes('标题文字'), 'text element must render')
  assert.ok(text.includes('卡片文字'), 'rect element must render')
  assert.ok(!/undefined/.test(text), 'render must not contain undefined leaves')

  // The host contract: one decks call, one state call, one SSE stream.
  assert.ok(fetches.some((f) => f.url.includes('/api/decks')), 'decks call missing')
  assert.ok(fetches.some((f) => f.url.includes('/api/state?deck=deck_test')), 'state call missing')
  assert.ok(eventSource && eventSource.url.includes('/api/events?deck=deck_test'), 'SSE stream missing')

  // A new slide pushed over SSE must land in the tree.
  eventSource.onopen?.()
  eventSource.onmessage?.({
    data: JSON.stringify({
      type: 'slide', seq: 3, deckId: 'deck_test', index: 1,
      slide: { id: 's2', title: '第二页', elements: [{ id: 'x1', type: 'text', box: [1, 1, 4, 1], text: 'SSE 推来的文字' }] },
      total: 2,
    }),
  })
  // `follow` is on by default, so the new page becomes the current one and its
  // element text is what the canvas draws.
  const after = treeText(await rerender())
  assert.ok(after.includes('SSE 推来的文字'), 'SSE slide event must reach the canvas')
  assert.ok(after.includes('第二页'), 'the new page must appear as a thumbnail')

  // An element patch on the page currently on screen must replace that
  // element in place, not duplicate it.
  eventSource.onmessage?.({
    data: JSON.stringify({
      type: 'element', seq: 4, deckId: 'deck_test', slideId: 's2',
      element: { id: 'x1', type: 'text', box: [1, 1, 4, 1], text: '改过的 SSE 文字' },
      elementId: 'x1',
    }),
  })
  const patched = treeText(await rerender())
  assert.ok(patched.includes('改过的 SSE 文字'), 'element patch must update the canvas')
  assert.equal(patched.split('SSE 推来的文字').length - 1, 0, 'the patched element must not be duplicated')

  // A slide inserted mid-list must shift the others, not overwrite them: the
  // same `slide` event covers insert / replace / move.
  eventSource.onmessage?.({
    data: JSON.stringify({
      type: 'slide', seq: 5, deckId: 'deck_test', index: 0,
      slide: { id: 's0', title: '插在最前', elements: [{ id: 'z1', type: 'text', box: [1, 1, 3, 1], text: '插队页' }] },
      total: 3,
    }),
  })
  await rerender()
  const thumbs = findAll(canvasTree, (node) => node.props && node.props['data-ps-thumb'] !== undefined)
  assert.equal(thumbs.length, 3, 'a mid-list insert must grow the deck to three pages')
  const titles = thumbs.map((node) => String(node.props.title))
  assert.deepEqual(titles, ['插在最前', '第一页', '第二页'], 'existing pages must keep their order after an insert')
})

test('a slide arriving over SSE is drawn element by element, not pasted in one frame', async () => {
  await mountCanvas()
  assert.ok(eventSource, 'SSE stream must exist')
  eventSource.onmessage?.({
    data: JSON.stringify({
      type: 'slide', seq: 9, deckId: 'deck_test', index: 1,
      slide: {
        id: 's2', title: '逐元素绘制', elements: [
          { id: 'a1', type: 'text', box: [1, 1, 6, 1], text: '第一笔' },
          { id: 'a2', type: 'rect', box: [1, 2.5, 6, 1], fill: '#F1F5F9', text: '第二笔' },
          { id: 'a3', type: 'text', box: [1, 4, 6, 1], text: '第三笔' },
        ],
      },
      total: 2,
    }),
  })
  canvasTree = await rerender()
  // Right after the event the page is a pencil ghost: every element is in the
  // model but still marked as being drawn, and the HUD announces the page.
  const ghosts = findAll(canvasTree, (node) => node.props && node.props['data-sketch'] === 'true' && node.props['data-id'])
  assert.equal(ghosts.length, 3, 'all three elements must start as ghosts')
  const hud = findAll(canvasTree, (node) => node.props && node.props['data-ps-draw'])[0]
  assert.ok(hud, 'the drawing HUD must be visible while elements settle')
  assert.equal(hud.props['data-count'], '3')
  const stage = findAll(canvasTree, (node) => node.props && node.props['data-ps-stage'])[0]
  assert.equal(stage.props['data-drawing'], 'true', 'the stage must announce that it is drawing')
  const stageInner = findAll(canvasTree, (node) => node.props && node.props.className === 'ps-stage-inner')[0]
  assert.ok(stageInner, 'the slide and its HUD need a common column wrapper')
  assert.equal(findAll(stageInner, (node) => node.props && node.props['data-ps-draw']).length, 1,
    'the HUD belongs right under the slide, not floating at the bottom of the editor')

  // A beat later the first element has settled while the rest still wait.
  await new Promise((resolve) => setTimeout(resolve, 95))
  canvasTree = await rerender()
  const settled = findAll(canvasTree, (node) => node.props && node.props['data-id'] && node.props['data-fresh'] === 'true')
  const waiting = findAll(canvasTree, (node) => node.props && node.props['data-id'] && node.props['data-sketch'] === 'true')
  assert.ok(settled.length >= 1, 'the first element must settle before the others')
  assert.ok(waiting.length >= 1, 'the remaining elements must still be waiting their turn')

  // After the last beat every element is drawn, and the HUD clears itself.
  await new Promise((resolve) => setTimeout(resolve, 1100))
  canvasTree = await rerender()
  assert.equal(findAll(canvasTree, (n) => n.props && n.props['data-sketch'] === 'true').length, 0, 'no ghost may survive the draw')
  assert.equal(findAll(canvasTree, (n) => n.props && n.props['data-ps-draw']).length, 0, 'the HUD must hide when the page is done')
})

test('before the agent opens a deck the editor area is already on screen', async () => {
  deckList = () => []
  stateDeck = () => null
  try {
    await mountCanvas()
    const root = findAll(canvasTree, (n) => n.props && n.props['data-ps-root'])[0]
    assert.equal(root.props['data-ps-root'], 'waiting', 'with no deck the canvas must show the waiting shell')
    assert.equal(findAll(canvasTree, (n) => n.props && n.props['data-ps-editor']).length, 1, 'the editor card must exist before the deck does')
    assert.equal(findAll(canvasTree, (n) => n.props && n.props['data-ps-controls']).length, 1, 'the controls card must exist before the deck does')
    assert.equal(findAll(canvasTree, (n) => n.props && n.props.className === 'ps-skeleton-paper').length, 1, 'a paper-shaped skeleton must hold the editing surface')
    assert.ok(findAll(canvasTree, (n) => n.props && n.props.className === 'ps-thumb ghost').length >= 1, 'the thumbnail rail must be reserved, not popped in later')
    assert.equal(findAll(canvasTree, (n) => n.props && n.props['data-ps-split']).length, 1, 'the editor/controls divider must be there too')
    assert.ok(treeText(canvasTree).includes('zone.editor'), 'the zone label must already be there')
  } finally {
    deckList = () => [{ id: 'deck_test', title: '画布渲染测试', slides: 1, sessionId: 'sess-1' }]
    stateDeck = () => HOST_DECK
  }
})

test('a deck id pushed through the tab params switches the open canvas', async () => {
  tabParams = { deckId: 'deck_first' }
  try {
    await mountCanvas()
    assert.ok(fetches.some((f) => f.url.includes('/api/state?deck=deck_first')), 'the canvas must load the deck named in the params')

    // `openTab` rewrites the already-open tab's navigation params; the canvas
    // has to follow, otherwise the tab keeps showing the previous build.
    tabParams = { deckId: 'deck_second' }
    canvasTree = await rerender()
    await new Promise((resolve) => setTimeout(resolve, 20))
    canvasTree = await rerender()
    assert.ok(fetches.some((f) => f.url.includes('/api/state?deck=deck_second')), 'a new deck id in the params must be loaded')
    assert.ok(eventSource && eventSource.url.includes('/api/events?deck=deck_second'), 'the SSE stream must reconnect to the new deck')
  } finally {
    tabParams = {}
  }
})

test('while waiting for a deck the canvas keeps looking until one appears', async () => {
  deckList = () => []
  stateDeck = () => null
  try {
    await mountCanvas()
    const before = fetches.filter((f) => f.url.includes('/api/decks')).length
    // The agent's pptx_studio_open lands while the tab is already waiting.
    deckList = () => [{ id: 'deck_test', title: '画布渲染测试', slides: 1, sessionId: 'sess-1' }]
    stateDeck = () => HOST_DECK
    await new Promise((resolve) => setTimeout(resolve, 1450))
    // The poll -> deck id -> state fetch chain settles over a few macrotasks.
    for (let pass = 0; pass < 6; pass += 1) {
      canvasTree = await rerender()
      await new Promise((resolve) => setTimeout(resolve, 30))
    }
    assert.ok(fetches.filter((f) => f.url.includes('/api/decks')).length > before, 'the waiting canvas must poll for a deck')
    assert.ok(fetches.some((f) => f.url.includes('/api/state?deck=deck_test')), 'the polled deck must be loaded into the canvas')
    assert.ok(treeText(canvasTree).includes('画布渲染测试'), 'the canvas must switch to the deck once it exists')
  } finally {
    deckList = () => [{ id: 'deck_test', title: '画布渲染测试', slides: 1, sessionId: 'sess-1' }]
    stateDeck = () => HOST_DECK
  }
})

test('the canvas draws the slide with PowerPoint geometry (96 CSS px per inch)', async () => {
  await mountCanvas()
  const canvas = findAll(canvasTree, (node) => node.props && node.props['data-ps-canvas'])[0]
  assert.ok(canvas, 'canvas surface missing')
  // No ResizeObserver in this harness, so the fit zoom stays at its 0.5 seed.
  const expectedWidth = 13.333 * 96 * 0.5
  const width = Number(String(canvas.props.style.width).replace('px', ''))
  assert.ok(Math.abs(width - expectedWidth) < 0.5, `canvas width must be slideW * 96 * zoom; got ${canvas.props.style.width}`)
  const height = Number(String(canvas.props.style.height).replace('px', ''))
  assert.ok(Math.abs(height - 7.5 * 96 * 0.5) < 0.5, `canvas height must be slideH * 96 * zoom; got ${canvas.props.style.height}`)

  // A 30pt run at 50% zoom is 30 * (96/72) * 0.5 = 20px, not 15px.
  // Scope the lookup to the slide surface: the thumbnail rail draws the same
  // elements at a different zoom.
  const title = findAll(canvas, (node) => node.props && node.props['data-id'] === 'e1')[0]
  assert.ok(title, 'title element missing')
  const titleText = findAll(title, (node) => node.props && node.props.style && node.props.style.fontSize)[0]
  assert.equal(titleText.props.style.fontSize, '20px', 'font size must scale from points, not from the old 1pt=1px shortcut')

  // An image element must point at the deck-scoped asset route.
  const image = findAll(canvas, (node) => node.props && node.props.src)[0]
  assert.ok(image, 'image element must render an <img>')
  assert.ok(image.props.src.includes('asset?'), `image src must use the asset route; got ${image.props.src}`)
  assert.ok(image.props.src.includes('deck=deck_test'), 'asset URL must carry the deck id (otherwise the host answers 404)')
  assert.ok(image.props.src.includes('path=assets'), 'asset URL must carry the element src path')
})

test('a thumbnail looks like its page: same background, aspect and geometry', async () => {
  stateDeck = () => PIPELINE_DECK
  try {
    await mountCanvas()
    const button = findAll(canvasTree, (node) => node.props && node.props['data-ps-thumb'] === '0')[0]
    assert.ok(button, 'thumbnail button missing')
    assert.equal(button.props.style.width, '88px')
    // 16:9 at 88px wide is 49.5px tall. A flex `stretch` used to force 60px and
    // vertically squash every miniature — the "缩略图跟实际页不一样" complaint.
    assert.equal(button.props.style.height, '49.5px', 'a 16:9 deck must get a 16:9 thumbnail')

    const inner = findAll(button, (node) => node.props && node.props['data-ps-thumb-inner'] === '1')[0]
    assert.ok(inner, 'thumbnail inner canvas missing')
    assert.equal(inner.props.style.background, '#F8FAFC', 'the thumbnail must paint the slide background, not a white default')
    assert.equal(inner.props['data-ps-thumb-scale'], '0.1375', 'the page must be laid out at a design size and scaled down')
    assert.ok(String(inner.props.style.transform).startsWith('scale('), 'the design canvas must be scaled, not drawn at 1.6px')
    assert.equal(inner.props.style.transformOrigin, 'top left')

    // The page is laid out at 0.5x with the real geometry: a 0.8in left edge is
    // 0.8 × 96 × 0.5 = 38.4px inside the design canvas, and 5.28px after scale.
    const el = findAll(inner, (node) => node.props && node.props['data-id'] === 'e1')[0]
    assert.ok(el, 'thumbnail element missing')
    assert.equal(el.props.style.left, '38.4px')
    assert.equal(el.props.style.width, '561.6px')
    assert.equal(el.props.style.fontSize, '20px', 'the miniature must use the page type scale, not 1.65px crumbs')
  } finally {
    stateDeck = () => HOST_DECK
  }
})

test('a 4:3 deck gets a 4:3 thumbnail instead of a 16:9 crop', async () => {
  const fourThree = {
    ...HOST_DECK,
    id: 'deck_43',
    size: { w: 10, h: 7.5 },
    slides: [{ ...HOST_DECK.slides[0], background: { color: '#101820' } }],
  }
  stateDeck = () => fourThree
  try {
    await mountCanvas()
    const button = findAll(canvasTree, (node) => node.props && node.props['data-ps-thumb'] === '0')[0]
    assert.equal(button.props.style.width, '88px')
    assert.equal(button.props.style.height, '66px', '88 * 7.5 / 10 = 66px for a 4:3 page')
    const inner = findAll(button, (node) => node.props && node.props['data-ps-thumb-inner'])[0]
    assert.equal(inner.props.style.background, '#101820', 'a custom slide background must reach the thumbnail')
    assert.equal(inner.props['data-ps-thumb-scale'], '0.1833', 'the design canvas follows the 4:3 page too')
  } finally {
    stateDeck = () => HOST_DECK
  }
})

test('bare hex colours render as real CSS, and a rectangle without a line draws no frame', async () => {
  stateDeck = () => BARE_HEX_DECK
  try {
    await mountCanvas()
    const canvas = findAll(canvasTree, (node) => node.props && node.props['data-ps-canvas'])[0]
    const styles = new Map()
    for (const node of findAll(canvas, (n) => n.props && n.props.className === 'ps-el')) {
      const id = node.props['data-id']
      if (id) styles.set(id, node.props.style)
    }
    assert.equal(styles.get('bg').background, '#060B18', 'a bare hex fill must reach CSS')
    assert.equal(styles.get('t1').color, '#FFFFFF', 'a bare hex text colour must reach CSS')
    assert.equal(styles.get('lineOnly').border, '1px solid #22D3EE', 'a line without lineWidth still strokes at 1pt')
    assert.equal(styles.get('plain').border, '0', 'a rectangle without a line must not inherit a black frame')
    assert.equal(styles.get('plain').background, '#0A1526')
  } finally {
    stateDeck = () => HOST_DECK
  }
})

/** Build one synthetic pointer event for the canvas handlers. */
function pointerEvent() {
  return {
    button: 0,
    clientX: 10,
    clientY: 10,
    shiftKey: false,
    pointerId: 1,
    stopPropagation() {},
    preventDefault() {},
    currentTarget: { setPointerCapture() {} },
  }
}

/** Build one synthetic keyboard event for the canvas shortcut handler. */
function keyEvent(key, options = {}) {
  const event = {
    key,
    ctrlKey: false,
    metaKey: false,
    shiftKey: false,
    altKey: false,
    target: { tagName: 'DIV' },
    defaultPrevented: false,
    preventDefault() { this.defaultPrevented = true },
    ...options,
  }
  return event
}

/** Find the rendered canvas root node (where keyboard shortcuts are wired). */
function canvasRoot(tree) {
  return findAll(tree, (node) => node.props && node.props['data-ps-root'] === 'deck')[0] || null
}

test('the workflow strip opens the research brief and the sandboxed HTML preview', async () => {
  stateDeck = () => PIPELINE_DECK
  try {
    await mountCanvas()
    const steps = findAll(canvasTree, (node) => node.props && node.props['data-ps-steps'])
    assert.equal(steps.length, 1, 'a three-step workflow strip must render')
    const stepButtons = findAll(canvasTree, (node) => node.props && node.props['data-ps-step'])
    assert.deepEqual(stepButtons.map((node) => node.props['data-ps-step']), ['research', 'preview', 'build'])
    assert.equal(stepButtons[0].props['data-done'], 'true', 'the research step is done in this deck')
    assert.equal(stepButtons[1].props['data-done'], 'true', 'the HTML preview step is done in this deck')

    // Research panel: summary, audience and source links are all visible.
    findAll(canvasTree, (node) => node.props && node.props['data-ps-step'] === 'research')[0].props.onClick()
    canvasTree = await rerender()
    assert.ok(findAll(canvasTree, (node) => node.props && node.props['data-ps-panel'] === 'research').length > 0, 'research panel missing')
    const researchText = treeText(canvasTree)
    assert.ok(researchText.includes('管理层汇报'), 'research summary must render')
    assert.ok(researchText.includes('批准预算'), 'goal must render')
    assert.ok(researchText.includes('为什么现在做'), 'outline intent must render')
    assert.ok(researchText.includes('https://example.com/report'), 'source link must render')

    // HTML preview panel: an iframe pointing at the host sandbox route.
    findAll(canvasTree, (node) => node.props && node.props['data-ps-step'] === 'preview')[0].props.onClick()
    canvasTree = await rerender()
    assert.ok(findAll(canvasTree, (node) => node.props && node.props['data-ps-panel'] === 'preview').length > 0, 'preview panel missing')
    const previewWrap = findAll(canvasTree, (node) => node.props && node.props['data-ps-preview-wrap'])[0]
    assert.ok(previewWrap, 'preview viewport wrapper missing')
    const frame = findAll(canvasTree, (node) => node.props && node.props['data-ps-preview'])[0]
    assert.ok(frame, 'preview iframe missing')
    assert.ok(String(frame.props.src).includes('/api/preview?deck=deck_test'), `iframe must use the preview route; got ${frame.props.src}`)
    assert.equal(frame.props.sandbox, '', 'the preview iframe must be sandboxed')
    assert.equal(frame.props.style.width, '960px', 'the iframe must keep the mockup design width')
    assert.equal(frame.props.style.height, '540px', 'the iframe must keep the mockup design height')
    assert.match(String(frame.props.style.transform), /^scale\(\d+(\.\d+)?\)$/, 'the mockup must be scaled uniformly, not stretched')
    const scale = Number(previewWrap.props['data-ps-preview-scale'])
    assert.ok(scale > 0 && scale <= 1.4, `unexpected preview scale: ${scale}`)
    assert.ok(Number(String(previewWrap.props.style.height).replace('px', '')) >= 220, 'the preview viewport must not collapse')
    assert.ok(findAll(canvasTree, (node) => node.props && node.props['data-ps-btn'] === 'preview-reload').length > 0)

    // Back to the slide editor.
    findAll(canvasTree, (node) => node.props && node.props['data-ps-step'] === 'build')[0].props.onClick()
    canvasTree = await rerender()
    assert.ok(findAll(canvasTree, (node) => node.props && node.props.className === 'ps-slide').length > 0, 'slide surface missing after switching back')
  } finally {
    stateDeck = () => HOST_DECK
  }
})

test('slide settings patch transition/notes and preview the transition effect', async () => {
  stateDeck = () => PIPELINE_DECK
  try {
    await mountCanvas()
    button(canvasTree, 'settings').props.onClick()
    canvasTree = await rerender()
    assert.ok(findAll(canvasTree, (node) => node.props && node.props['data-ps-settings']).length > 0, 'settings panel missing')

    const transition = findAll(canvasTree, (node) => node.props && node.props['data-ps-field'] === 'transition')[0]
    assert.ok(transition, 'transition select missing')
    assert.equal(transition.props.value, 'fade')
    const before = fetches.filter((f) => f.url.includes('/api/edit')).length
    transition.props.onChange({ target: { value: 'push' } })
    await new Promise((resolve) => setTimeout(resolve, 10))
    const editCall = fetches.filter((f) => f.url.includes('/api/edit'))[before]
    assert.ok(editCall, 'transition change must POST /api/edit')
    const ops = JSON.parse(editCall.init.body).ops
    assert.equal(ops[0].kind, 'slide.patch')
    assert.equal(ops[0].patch.transition.type, 'push')

    canvasTree = await rerender()
    const play = button(canvasTree, 'play-transition')
    assert.ok(play, 'play-transition button missing')
    assert.equal(play.props.disabled, false)
    play.props.onClick()
    canvasTree = await rerender()
    const slide = findAll(canvasTree, (node) => node.props && node.props.className === 'ps-slide')[0]
    assert.equal(slide.props['data-fx'], 'push', 'the transition preview must tag the slide')
    assert.equal(slide.props['data-fx-dir'], 'left')

    const notes = findAll(canvasTree, (node) => node.props && node.props['data-ps-field'] === 'notes')[0]
    assert.ok(notes, 'notes editor missing')
    notes.props.onChange({ target: { value: '新的演讲备注' } })
    canvasTree = await rerender()
    const notesAgain = findAll(canvasTree, (node) => node.props && node.props['data-ps-field'] === 'notes')[0]
    const beforeNotes = fetches.filter((f) => f.url.includes('/api/edit')).length
    notesAgain.props.onBlur()
    await new Promise((resolve) => setTimeout(resolve, 10))
    const notesCall = fetches.filter((f) => f.url.includes('/api/edit'))[beforeNotes]
    assert.ok(notesCall, 'notes blur must POST /api/edit')
    assert.equal(JSON.parse(notesCall.init.body).ops[0].patch.notes, '新的演讲备注')
  } finally {
    stateDeck = () => HOST_DECK
  }
})

test('theme/size panels and insert buttons commit deck-level and chart ops', async () => {
  stateDeck = () => PIPELINE_DECK
  try {
    await mountCanvas()
    button(canvasTree, 'theme').props.onClick()
    canvasTree = await rerender()
    const accent = findAll(canvasTree, (node) => node.props && node.props['data-ps-theme-color'] === 'accent')[0]
    assert.ok(accent, 'theme accent input missing')
    const beforeTheme = fetches.filter((f) => f.url.includes('/api/edit')).length
    accent.props.onChange({ target: { value: '#22D3EE' } })
    await new Promise((resolve) => setTimeout(resolve, 10))
    const themeCall = fetches.filter((f) => f.url.includes('/api/edit'))[beforeTheme]
    assert.ok(themeCall, 'theme change must POST /api/edit')
    const themeOp = JSON.parse(themeCall.init.body).ops[0]
    assert.equal(themeOp.kind, 'deck.theme')
    assert.equal(themeOp.theme.colors.accent, '#22D3EE')

    canvasTree = await rerender()
    const size = findAll(canvasTree, (node) => node.props && node.props['data-ps-field'] === 'size')[0]
    assert.ok(size, 'size preset select missing')
    const beforeSize = fetches.filter((f) => f.url.includes('/api/edit')).length
    size.props.onChange({ target: { value: '4:3' } })
    await new Promise((resolve) => setTimeout(resolve, 10))
    const sizeCall = fetches.filter((f) => f.url.includes('/api/edit'))[beforeSize]
    assert.ok(sizeCall, 'size change must POST /api/edit')
    const sizeOp = JSON.parse(sizeCall.init.body).ops[0]
    assert.equal(sizeOp.kind, 'deck.size')
    assert.equal(sizeOp.width, 10)
    assert.equal(sizeOp.height, 7.5)

    canvasTree = await rerender()
    const beforeChart = fetches.filter((f) => f.url.includes('/api/edit')).length
    button(canvasTree, 'add-chart').props.onClick()
    await new Promise((resolve) => setTimeout(resolve, 10))
    const chartCall = fetches.filter((f) => f.url.includes('/api/edit'))[beforeChart]
    assert.ok(chartCall, 'add-chart must POST /api/edit')
    const chartOp = JSON.parse(chartCall.init.body).ops[0]
    assert.equal(chartOp.kind, 'element.add')
    assert.equal(chartOp.element.type, 'chart')
    assert.ok(chartOp.element.chart.series.length > 0)
  } finally {
    stateDeck = () => HOST_DECK
  }
})

test('element animation/shadow controls patch the element model', async () => {
  stateDeck = () => PIPELINE_DECK
  try {
    await mountCanvas()
    const hit = findAll(canvasTree, (node) => node.props && node.props.className === 'ps-hit' && typeof node.props.onPointerDown === 'function')[0]
    assert.ok(hit, 'element hit surface missing')
    hit.props.onPointerDown(pointerEvent())
    canvasTree = await rerender()

    const anim = findAll(canvasTree, (node) => node.props && node.props['data-ps-field'] === 'anim-effect')[0]
    assert.ok(anim, 'animation select missing')
    assert.equal(anim.props.value, 'fade')
    const before = fetches.filter((f) => f.url.includes('/api/edit')).length
    anim.props.onChange({ target: { value: 'zoom' } })
    await new Promise((resolve) => setTimeout(resolve, 10))
    const animCall = fetches.filter((f) => f.url.includes('/api/edit'))[before]
    assert.ok(animCall, 'animation change must POST /api/edit')
    const animOp = JSON.parse(animCall.init.body).ops[0]
    assert.equal(animOp.kind, 'element.patch')
    assert.equal(animOp.patch.anim.effect, 'zoom')

    canvasTree = await rerender()
    assert.ok(findAll(canvasTree, (node) => node.props && node.props['data-ps-btn'] === 'anim-preview').length > 0, 'animation preview button missing')
    const shadow = button(canvasTree, 'shadow')
    assert.ok(shadow, 'shadow toggle missing')
    const beforeShadow = fetches.filter((f) => f.url.includes('/api/edit')).length
    shadow.props.onClick()
    await new Promise((resolve) => setTimeout(resolve, 10))
    const shadowCall = fetches.filter((f) => f.url.includes('/api/edit'))[beforeShadow]
    assert.ok(shadowCall, 'shadow toggle must POST /api/edit')
    assert.equal(JSON.parse(shadowCall.init.body).ops[0].patch.shadow, null)
  } finally {
    stateDeck = () => HOST_DECK
  }
})

test('the editor and the control dock are visibly separate zones', async () => {
  await mountCanvas()
  const editor = findAll(canvasTree, (node) => node.props && node.props['data-ps-editor'])[0]
  const controls = findAll(canvasTree, (node) => node.props && node.props['data-ps-controls'])[0]
  const split = findAll(canvasTree, (node) => node.props && node.props['data-ps-split'])[0]
  assert.ok(editor, 'editor zone missing')
  assert.ok(controls, 'control dock missing')
  assert.ok(split, 'editor/control splitter missing')
  assert.ok(findAll(editor, (node) => node.props && node.props['data-ps-canvas']).length > 0, 'canvas must live inside the editor zone')
  assert.ok(findAll(controls, (node) => node.props && node.props['data-ps-status']).length > 0, 'status/tools must live inside the control dock')
  const css = styleSheets[0].textContent
  assert.ok(/\.ps-editor\{[^}]*border:1px solid var\(--ps-border-strong\)/.test(css), 'editor needs a strong border')
  assert.ok(/\.ps-controls\{[^}]*border:1px solid var\(--ps-border-strong\)/.test(css), 'control dock needs a strong border')
  assert.ok(css.includes('.ps-split-grip'), 'splitter grip missing')
})

test('common PPTX shortcuts work: save, copy/paste, cut, undo, new slide', async () => {
  await mountCanvas()
  let root = canvasRoot(canvasTree)
  assert.ok(root && typeof root.props.onKeyDown === 'function', 'canvas root must own the shortcut handler')

  const exportBefore = fetches.filter((f) => f.url.includes('/api/export')).length
  const saveEvent = keyEvent('s', { ctrlKey: true })
  root.props.onKeyDown(saveEvent)
  await new Promise((resolve) => setTimeout(resolve, 20))
  assert.equal(saveEvent.defaultPrevented, true, 'Ctrl+S must stop the browser save dialog')
  assert.ok(fetches.filter((f) => f.url.includes('/api/export')).length > exportBefore, 'Ctrl+S must export the deck')

  const hit = findAll(canvasTree, (node) => node.props && node.props.className === 'ps-hit' && typeof node.props.onPointerDown === 'function')[0]
  assert.ok(hit, 'element hit surface missing')
  hit.props.onPointerDown(pointerEvent())
  canvasTree = await rerender()
  root = canvasRoot(canvasTree)

  const pasteBefore = fetches.filter((f) => f.url.includes('/api/edit')).length
  const copyEvent = keyEvent('c', { ctrlKey: true })
  root.props.onKeyDown(copyEvent)
  assert.equal(copyEvent.defaultPrevented, true, 'Ctrl+C must be handled by the canvas')
  canvasTree = await rerender()
  root = canvasRoot(canvasTree)
  const pasteEvent = keyEvent('v', { ctrlKey: true })
  root.props.onKeyDown(pasteEvent)
  await new Promise((resolve) => setTimeout(resolve, 10))
  assert.equal(pasteEvent.defaultPrevented, true, 'Ctrl+V must be handled when the internal clipboard has elements')
  const pasteCall = fetches.filter((f) => f.url.includes('/api/edit'))[pasteBefore]
  assert.ok(pasteCall, 'Ctrl+V must POST /api/edit')
  const pasteOps = JSON.parse(pasteCall.init.body).ops
  assert.equal(pasteOps[0].kind, 'element.add')
  assert.match(pasteOps[0].element.id, /-copy/)

  canvasTree = await rerender()
  root = canvasRoot(canvasTree)
  const cutBefore = fetches.filter((f) => f.url.includes('/api/edit')).length
  const cutEvent = keyEvent('x', { ctrlKey: true })
  root.props.onKeyDown(cutEvent)
  await new Promise((resolve) => setTimeout(resolve, 10))
  assert.equal(cutEvent.defaultPrevented, true)
  const cutCall = fetches.filter((f) => f.url.includes('/api/edit'))[cutBefore]
  assert.ok(cutCall, 'Ctrl+X must POST /api/edit')
  assert.equal(JSON.parse(cutCall.init.body).ops[0].remove, true)

  canvasTree = await rerender()
  root = canvasRoot(canvasTree)
  const undoBefore = fetches.filter((f) => f.url.includes('/api/edit')).length
  root.props.onKeyDown(keyEvent('z', { ctrlKey: true }))
  await new Promise((resolve) => setTimeout(resolve, 10))
  assert.ok(fetches.filter((f) => f.url.includes('/api/edit')).length > undoBefore, 'Ctrl+Z must send the undo diff')

  canvasTree = await rerender()
  root = canvasRoot(canvasTree)
  const addBefore = fetches.filter((f) => f.url.includes('/api/edit')).length
  root.props.onKeyDown(keyEvent('m', { ctrlKey: true }))
  await new Promise((resolve) => setTimeout(resolve, 10))
  const addCall = fetches.filter((f) => f.url.includes('/api/edit'))[addBefore]
  assert.ok(addCall, 'Ctrl+M must POST /api/edit')
  assert.equal(JSON.parse(addCall.init.body).ops[0].kind, 'slide.add')

  canvasTree = await rerender()
  root = canvasRoot(canvasTree)
  const guardBefore = fetches.filter((f) => f.url.includes('/api/edit')).length
  root.props.onKeyDown(keyEvent('z', { ctrlKey: true, target: { tagName: 'TEXTAREA' } }))
  await new Promise((resolve) => setTimeout(resolve, 10))
  assert.equal(fetches.filter((f) => f.url.includes('/api/edit')).length, guardBefore, 'Ctrl+Z inside a textarea must stay native')
})

test('Ctrl+A selects all and Delete removes the set in one batch', async () => {
  await mountCanvas()
  let root = canvasRoot(canvasTree)
  root.props.onKeyDown(keyEvent('a', { ctrlKey: true }))
  canvasTree = await rerender()
  const multi = findAll(canvasTree, (node) => node.props && node.props['data-ps-multi'])[0]
  assert.ok(multi, 'multi-selection inspector missing')
  const text = treeText(canvasTree)
  assert.ok(text.includes('select.multi'), 'multi-selection count missing')

  root = canvasRoot(canvasTree)
  const before = fetches.filter((f) => f.url.includes('/api/edit')).length
  root.props.onKeyDown(keyEvent('Delete'))
  await new Promise((resolve) => setTimeout(resolve, 10))
  const call = fetches.filter((f) => f.url.includes('/api/edit'))[before]
  assert.ok(call, 'Delete must POST /api/edit')
  const ops = JSON.parse(call.init.body).ops
  assert.equal(ops.length, HOST_DECK.slides[0].elements.length)
  assert.ok(ops.every((op) => op.kind === 'element.patch' && op.remove === true), 'each selected element must be removed')
})

test('shortcut help lists the supported keys', async () => {
  await mountCanvas()
  button(canvasTree, 'shortcuts').props.onClick()
  canvasTree = await rerender()
  const panel = findAll(canvasTree, (node) => node.props && node.props['data-ps-shortcuts'])[0]
  assert.ok(panel, 'shortcut help panel missing')
  const text = treeText(panel)
  for (const key of ['Ctrl+S', 'Ctrl+C', 'Ctrl+V', 'Ctrl+Z', 'Ctrl+M']) {
    assert.ok(text.includes(key), `shortcut help must list ${key}`)
  }
})

test('the toolbar buttons are wired: export, download, add slide, undo/redo', async () => {
  await mountCanvas()
  const exportButton = button(canvasTree, 'export')
  assert.ok(exportButton, 'export button missing')
  const before = fetches.length
  exportButton.props.onClick()
  await new Promise((resolve) => setTimeout(resolve, 10))
  assert.ok(fetches.slice(before).some((f) => f.url.includes('/api/export') && f.init && f.init.method === 'POST'), 'export button must POST /api/export')
  canvasTree = await rerender()
  assert.ok(button(canvasTree, 'download'), 'a download button must appear once an export exists')

  const addSlide = button(canvasTree, 'add-slide')
  assert.ok(addSlide, 'add-slide button missing')
  const editsBefore = fetches.filter((f) => f.url.includes('/api/edit')).length
  addSlide.props.onClick()
  await new Promise((resolve) => setTimeout(resolve, 10))
  const editCall = fetches.filter((f) => f.url.includes('/api/edit'))[editsBefore]
  assert.ok(editCall, 'add-slide must POST /api/edit')
  const ops = JSON.parse(editCall.init.body).ops
  assert.equal(ops[0].kind, 'slide.add')
  assert.ok(typeof ops[0].slide.id === 'string' && ops[0].slide.id !== '', 'slide.add must carry the id the host will use')

  canvasTree = await rerender()
  const undo = button(canvasTree, 'undo')
  const redo = button(canvasTree, 'redo')
  assert.ok(undo && redo, 'undo/redo buttons missing')
  assert.equal(undo.props.disabled, false, 'undo must enable after an edit')
  assert.equal(redo.props.disabled, true, 'redo must stay disabled until something is undone')
})

test('the deck delete button arms itself before it deletes', async () => {
  await mountCanvas()
  const first = button(canvasTree, 'deck-delete')
  assert.ok(first, 'a delete button must exist while a deck is open')
  const before = fetches.length
  first.props.onClick()
  canvasTree = await rerender()
  assert.equal(fetches.slice(before).some((f) => f.url.includes('/api/deck/delete')), false, 'the first click only arms the button')
  const armed = button(canvasTree, 'deck-delete')
  assert.notEqual(armed.children[0], '×', 'an armed button must say so')
  assert.equal(armed.props.className.includes('danger'), true, 'an armed button must look dangerous')

  armed.props.onClick()
  await new Promise((resolve) => setTimeout(resolve, 20))
  const hit = fetches.slice(before).find((f) => f.url.includes('/api/deck/delete'))
  assert.ok(hit, 'the second click must POST /api/deck/delete')
  assert.equal(hit.init.method, 'POST')
  assert.equal(JSON.parse(hit.init.body).deckId, 'deck_test')

  canvasTree = await rerender()
  assert.ok(!button(canvasTree, 'deck-delete'), 'with no deck left the delete button must disappear')
  const root = findAll(canvasTree, (node) => node.props && node.props['data-ps-root'])[0]
  assert.equal(root.props['data-ps-root'], 'waiting', 'the canvas must fall back to the waiting shell')
  assert.equal(findAll(canvasTree, (node) => node.props && node.props['data-ps-editor']).length, 1,
    'the editor card must stay on screen after the last deck is deleted')
})

test('selecting an element, typing a note and sending reaches the session path', async () => {
  await mountCanvas()
  // Thumbnails draw the same hit surface without handlers; pick the live one.
  const hit = findAll(canvasTree, (node) => node.props && node.props.className === 'ps-hit' && typeof node.props.onPointerDown === 'function')[0]
  assert.ok(hit, 'element hit surface missing')
  const event = {
    button: 0,
    clientX: 10,
    clientY: 10,
    shiftKey: false,
    pointerId: 1,
    stopPropagation() {},
    preventDefault() {},
    currentTarget: { setPointerCapture() {} },
  }
  hit.props.onPointerDown(event, null, [0.8, 0.6, 11.7, 0.9])
  canvasTree = await rerender()
  const selected = treeText(canvasTree)
  assert.ok(selected.includes('select.element'), 'selection bar must appear after a pick')

  const note = findAll(canvasTree, (node) => node.props && node.props['data-ps-note'])[0]
  assert.ok(note, 'note textarea missing')
  note.props.onChange({ target: { value: '标题太长，压到副标题了' } })
  canvasTree = await rerender()

  const send = button(canvasTree, 'send')
  assert.ok(send, 'send button missing')
  const before = fetches.length
  send.props.onClick()
  await new Promise((resolve) => setTimeout(resolve, 10))
  const call = fetches.slice(before).find((f) => f.url.includes('/api/feedback'))
  assert.ok(call, 'send must POST /api/feedback')
  const body = JSON.parse(call.init.body)
  assert.equal(body.inject, true, 'without a composer the host must inject the message')
  assert.ok(body.text.includes('用户意见：标题太长，压到副标题了'), `feedback text must carry the note; got ${body.text}`)
  assert.equal(body.elementId, 'e1', 'feedback must name the element the user pointed at')
})

test('the tool card opens the canvas as soon as the call starts', async () => {
  /** @type {any} */
  let card = null
  const captureCtx = {
    ...ctx,
    slots: {
      inject: (name, body) => body(),
      register: (spec, component) => { if (spec.name === 'tool.call.toolview') card = component; return { spec, component } },
    },
  }
  module.apply(captureCtx)
  assert.equal(typeof card, 'function')
  /**
   * Render the card to a settled tree with a fresh hook runtime.
   * @param {string} phase - tool phase.
   * @param {() => boolean} [openStudio] - opener stub.
   * @returns {Promise<{tree: any, opens: () => number}>} tree and open count.
   */
  const renderCard = async (phase, openStudio) => {
    resetRuntime()
    let opens = 0
    const tree = await renderTree(card, {
      phase,
      callId: 'call-1',
      sessionId: 'sess-1',
      toolName: 'pptx_studio_open',
      t: (key) => String(key),
      block: { args: { text: (key) => (key === 'title' ? '测试 deck' : key === 'template' ? '' : undefined) } },
      openStudio: openStudio || (() => { opens += 1; return true }),
    })
    return { tree, opens: () => opens }
  }
  const flatten = (tree) => JSON.stringify(tree, (key, value) => (typeof value === 'function' ? 'fn' : value))

  // The regression this guards: opening only on `result` means the user stares
  // at the conversation while the deck is drawn behind a closed tab, and then
  // finds a finished file. The canvas must show up with the call itself.
  const started = await renderCard('start')
  assert.equal(started.opens(), 1, 'the card must open the canvas while the call is still running')
  assert.ok(flatten(started.tree).includes('测试 deck'), 'card must show the deck title from the tool args')
  assert.ok(flatten(started.tree).includes('card.live'), 'a running card must say the canvas is following the build')

  const preparing = await renderCard('preparing')
  assert.equal(preparing.opens(), 1, 'the preparing phase (arguments still streaming) must open it too')

  // A sidebar that is not mounted yet must say so instead of pretending.
  const blocked = await renderCard('start', () => false)
  assert.ok(flatten(blocked.tree).includes('card.opening'), 'a running card that failed to open must say it is opening')

  const finished = await renderCard('result')
  assert.equal(finished.opens(), 1, 'the result phase must still open the canvas when it is the first render')
  assert.ok(flatten(finished.tree).includes('card.opened'), 'a finished card must confirm the canvas is open')
})

test('the opener refuses to open the canvas without a mounted session', () => {
  /** @type {any} */
  let inject = null
  const captureCtx = {
    ...ctx,
    slots: {
      inject: (name, body) => { if (name === 'tool.call.toolview') { const registered = body(); inject = registered.spec.inject } return () => {} },
      register: (spec, component) => ({ spec, component }),
    },
    get: (name) => {
      if (name === 'sidebarRight') {
        return {
          mounted: { getSnapshot: () => undefined },
          openTab: () => { throw new Error('should not be called') },
          openTabIn: () => { throw new Error('should not be called') },
        }
      }
      return undefined
    },
  }
  module.apply(captureCtx)
  assert.equal(typeof inject, 'function')
  const props = inject('sess-1')
  assert.equal(props.openStudio({ deckId: 'x' }), false, 'opener must fail closed with no session on screen')
})

await Promise.all(pending)
console.log(`\nclient: ${passed} checks passed`)
