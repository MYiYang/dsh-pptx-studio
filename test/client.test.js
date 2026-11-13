/**
 * Client-half harness: load `client.js` the way the Web carrier does (a
 * classic script that calls `window.__ModuleLoader__.load({id, factory})`),
 * then run `apply` against a recording ctx and render the two surfaces.
 *
 * This is what catches "the client half silently eats itself": a wrong service
 * name, a missing slot, or a render-time throw all fail here instead of in the
 * browser with no console access.
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
/** @type {Array<Promise<void>>} */
const pending = []
/**
 * Run one named check. Async bodies are collected so the summary waits for
 * them: without that, a late failure would surface as an unhandled rejection
 * and the pass count would be printed before the checks finish.
 * @param {string} name - test name.
 * @param {() => void|Promise<void>} fn - body.
 * @returns {Promise<void>} resolves when the check is done.
 */
function test(name, fn) {
  const run = (async () => {
    try {
      await fn()
      passed += 1
      console.log(`ok - ${name}`)
    } catch (error) {
      console.error(`not ok - ${name}`)
      throw error
    }
  })()
  pending.push(run)
  return run
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
      const slot = runtime.effects[at]
      if (!slot) {
        runtime.effects[at] = { fn, deps }
        fn()
        return
      }
      const same = Array.isArray(deps) && Array.isArray(slot.deps) && deps.length === slot.deps.length && deps.every((d, i) => d === slot.deps[i])
      if (!same) {
        runtime.effects[at] = { fn, deps }
        fn()
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
  return { react, render }
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

/** Deck the harness serves to the client half. */
const HOST_DECK = {
  version: 1,
  id: 'deck_test',
  title: '画布渲染测试',
  size: { w: 13.333, h: 7.5 },
  slides: [
    {
      id: 's1', title: '第一页', elements: [
        { id: 'e1', type: 'text', box: [0.8, 0.6, 11.7, 0.9], text: '标题文字', style: { size: 30, bold: true, color: '#0F172A' } },
        { id: 'e2', type: 'rect', box: [0.8, 2, 5, 1.5], fill: '#F1F5F9', radius: 0.06, text: '卡片文字' },
        { id: 'e3', type: 'table', box: [0.8, 4, 11, 2], table: { rows: [['a', 'b'], ['1', '2']] } },
        { id: 'e4', type: 'chart', box: [0.8, 6.2, 4, 1], text: '图表占位' },
      ],
    },
  ],
}

/** @type {Array<{url: string, init: any}>} */
const fetches = []
/** @type {any} */
let eventSource = null
globalThis.fetch = async (url, init) => {
  const href = String(url)
  fetches.push({ url: href, init })
  const json = (value) => ({ ok: true, status: 200, text: async () => JSON.stringify(value) })
  if (href.includes('/api/decks')) return json({ ok: true, decks: [{ id: 'deck_test', title: '画布渲染测试', slides: 1 }] })
  if (href.includes('/api/state')) return json({ ok: true, deck: HOST_DECK, template: null })
  if (href.includes('/api/feedback')) return json({ ok: true, delivered: true, text: 'sent' })
  if (href.includes('/api/export')) return json({ ok: true, value: { path: 'F:/tmp/deck.pptx', bytes: 12345, slides: 1 } })
  return { ok: false, status: 404, text: async () => JSON.stringify({ ok: false, code: 'not-found' }) }
}
globalThis.EventSource = class {
  constructor(url) { this.url = url; this.onopen = null; this.onerror = null; this.onmessage = null; eventSource = this }
  close() { this.closed = true }
}

const { react: fakeReact, render: renderTree } = makeReact()

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
  assert.ok(calls.locale.some(([ns, id, keys]) => ns === 'pptxStudio' && id === 'zh' && keys > 10), 'zh dictionary missing')
  assert.ok(calls.locale.some(([, id]) => id === 'en'), 'en dictionary missing')
  assert.equal(styleSheets.length, 1, 'one stylesheet must be mounted')
  assert.ok(styleSheets[0].textContent.includes('.ps-canvas'), 'stylesheet must carry the canvas rules')
})

let canvasRuns = 0
test('the canvas body renders the deck it fetched, then reacts to SSE', async () => {
  canvasRuns += 1
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

  const props = {
    sessionId: 'sess-1',
    t: (key, values) => (values ? String(key).replace(/\{(\w+)\}/g, (_, k) => String(values[k])) : String(key)),
    useTabInfo: () => ({
      tab: { id: 'tab-1', title: 'PPTX 画布', actions: { close: () => {} }, navigation: { params: {} } },
      panel: { id: 'pane-1' },
      sidebar: { expanded: true },
    }),
    inputActions: { setDraft: () => {}, persistDraft: () => {}, submit: () => {} },
  }
  const text = treeText(await renderTree(bodyComponent, props))
  assert.ok(text.includes('画布渲染测试'), `deck title must appear after the state fetch; got: ${text.slice(0, 300)}`)
  assert.ok(text.includes('标题文字'), 'text element must render')
  assert.ok(text.includes('卡片文字'), 'rect element must render')
  assert.ok(!/undefined/.test(text), 'render must not contain undefined leaves')

  // The host contract: one decks call, one state call, one SSE stream.
  assert.ok(fetches.some((f) => f.url.includes('/api/decks?sessionId=sess-1')), 'decks call missing')
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
  // A new slide lands in the model: the page count grows and its thumbnail
  // appears (the canvas shows one page at a time, so the element text of the
  // new page is only visible once that page is current).
  // `follow` is on by default, so the new page becomes the current one and its
  // element text is what the canvas draws.
  const after = treeText(await renderTree(bodyComponent, props))
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
  const patched = treeText(await renderTree(bodyComponent, props))
  assert.ok(patched.includes('改过的 SSE 文字'), 'element patch must update the canvas')
  assert.equal(patched.split('SSE 推来的文字').length - 1, 0, 'the patched element must not be duplicated')
})

test('the tool card renders for pptx_studio_open', () => {
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
  const tree = card({
    phase: 'result',
    callId: 'call-1',
    sessionId: 'sess-1',
    toolName: 'pptx_studio_open',
    t: (key) => String(key),
    block: { args: { text: (key) => (key === 'title' ? '测试 deck' : key === 'template' ? '' : undefined) } },
    openStudio: () => true,
  })
  assert.ok(tree && tree.type === 'div')
  const flat = JSON.stringify(tree, (key, value) => (typeof value === 'function' ? 'fn' : value))
  assert.ok(flat.includes('测试 deck'), 'card must show the deck title from the tool args')
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
