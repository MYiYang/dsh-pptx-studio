/**
 * End-to-end harness: apply the real plugin against a fake ctx, run the
 * tool chain (open → plan → slide → add_element → patch → inspect → export),
 * and assert the results survive the registry's lossless-JSON snapshot rule.
 *
 * Run: node test/e2e.test.js
 */
import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { Readable } from 'node:stream'

import { PythonRunner, runProcess } from '../lib/python.js'

const plugin = await import('../index.js')

let passed = 0
/**
 * Run one named check.
 * @param {string} name - test name.
 * @param {() => Promise<void>|void} fn - body.
 * @returns {Promise<void>} resolves when done.
 */
async function test(name, fn) {
  try {
    await fn()
    passed += 1
    console.log(`ok - ${name}`)
  } catch (error) {
    console.error(`not ok - ${name}`)
    throw error
  }
}

/**
 * The exact rule DSH applies to a tool result (`snapshotJsonValue` in
 * `@deepseek-ai/dsh-util-values`): plain objects/arrays only, no `undefined`,
 * no NaN, no -0, no class instances, no non-string keys.
 * @param {unknown} value - candidate.
 * @returns {boolean} whether the value survives.
 */
function lossless(value) {
  if (value === null) return true
  const kind = typeof value
  if (kind === 'boolean' || kind === 'string') return true
  if (kind === 'number') return Number.isFinite(value) && !Object.is(value, -0)
  if (kind !== 'object') return false
  if (Array.isArray(value)) {
    if (Reflect.ownKeys(value).length !== value.length + 1) return false
    return value.every(lossless)
  }
  const proto = Object.getPrototypeOf(value)
  if (proto !== null && proto !== Object.prototype) return false
  const keys = Reflect.ownKeys(value)
  if (keys.some((key) => typeof key !== 'string')) return false
  return keys.every((key) => lossless(value[key]))
}

const workspace = mkdtempSync(join(tmpdir(), 'pptx-studio-e2e-'))

/** Captured web route definition (registered by the plugin). */
let route = null
/** @type {any[]} */
const tools = []
/** @type {Map<string, Set<(event: Record<string, unknown>) => void>>} */
const events = new Map()
/** Skills the plugin registered with the host registry. @type {any[]} */
const skills = []
const ctx = {
  logger: { info() {}, warn() {}, debug() {}, error() {} },
  tools: { register: (tool) => { tools.push(tool); return () => {} } },
  get(name) {
    if (name === 'tools') return ctx.tools
    return undefined
  },
  inject(wanted, callback) {
    if (!Array.isArray(wanted) || !wanted.includes('skills')) return
    callback({
      skills: { register: (skill) => { skills.push(skill); return () => {} } },
      effect: (fn) => { fn(); return () => {} },
    })
  },
  loader: {
    import: async (name) => {
      if (name === '@deepseek-ai/dsh-tools') return { defineTool: (definition) => definition }
      throw new Error(`unexpected import: ${name}`)
    },
  },
  webServer: { register: (definition) => { route = definition; return () => {} } },
  connection: { requestRejection: () => undefined },
  effect: (fn) => { fn(); return () => {} },
}

/**
 * Build a fake Node request that streams one JSON body.
 * @param {string} method - HTTP method.
 * @param {string} url - request URL.
 * @param {unknown} [body] - JSON body (omitted for GET).
 * @returns {any} fake IncomingMessage (a Readable with `.on`).
 */
function fakeRequest(method, url, body) {
  const text = body === undefined ? '' : JSON.stringify(body)
  const req = Readable.from(text === '' ? [] : [Buffer.from(text, 'utf8')])
  req.method = method
  req.url = url
  req.headers = {}
  req.destroy = () => {}
  return req
}

/**
 * Build a fake Node response that captures status, headers and body bytes.
 * @returns {any} fake ServerResponse.
 */
function fakeResponse() {
  const res = {
    statusCode: 0,
    headers: {},
    chunks: [],
    writeHead(code, headers) {
      res.statusCode = code
      res.headers = { ...headers }
      return res
    },
    write(chunk) {
      if (chunk !== undefined) res.chunks.push(Buffer.from(chunk))
      return true
    },
    end(chunk) {
      if (chunk !== undefined) res.chunks.push(Buffer.from(chunk))
    },
    // An SSE route hangs its cleanup on the response's close event.
    on() { return res },
  }
  Object.defineProperty(res, 'body', { get: () => Buffer.concat(res.chunks) })
  Object.defineProperty(res, 'text', { get: () => res.body.toString('utf8') })
  Object.defineProperty(res, 'json', { get: () => JSON.parse(res.text || '{}') })
  return res
}

/**
 * Call one host web route through the captured handler.
 * @param {string} method - HTTP method.
 * @param {string} pathWithQuery - path after the API prefix, including `?`.
 * @param {unknown} [body] - JSON body (omitted for GET).
 * @returns {Promise<any>} fake response (statusCode/headers/body/text/json).
 */
async function call(method, pathWithQuery, body) {
  assert.ok(route, 'the plugin registered a web route')
  const req = fakeRequest(method, `/dsh-pptx-studio/api${pathWithQuery}`, body)
  const res = fakeResponse()
  // SSE routes keep the response open; hand the request back so a test can
  // close the stream when it is done with it.
  res.request = req
  await route.handler(req, res)
  return res
}

const exec = { agent: { session: { id: 'sess-e2e', cwd: workspace } } }
/**
 * Look up one registered tool.
 * @param {string} name - tool name.
 * @returns {any} tool definition.
 */
const tool = (name) => tools.find((t) => t.name === name)

await test('apply registers twelve pptx_studio tools and the web route', async () => {
  await plugin.apply(ctx, {})
  assert.equal(tools.length, 12)
  for (const name of [
    'pptx_studio_open', 'pptx_studio_template', 'pptx_studio_research', 'pptx_studio_preview',
    'pptx_studio_deck', 'pptx_studio_plan', 'pptx_studio_slide', 'pptx_studio_add_element',
    'pptx_studio_patch', 'pptx_studio_inspect', 'pptx_studio_verify', 'pptx_studio_export',
  ]) {
    assert.ok(tool(name), `missing ${name}`)
    assert.equal(typeof tool(name).execute, 'function')
    assert.equal(typeof tool(name).output.render, 'function', 'output.render is required by the registry')
  }
  assert.ok(route, 'web route registered')
  assert.equal(route.kind, 'prefix')
  assert.equal(route.path, '/dsh-pptx-studio/api')
  assert.equal(typeof route.handler, 'function')
  // The plugin's own skill is the constraint: it must arrive with `apply`.
  assert.equal(skills.length, 1, 'apply registers exactly one bundled skill')
  assert.equal(skills[0].name, 'pptx-studio')
  assert.match(skills[0].description, /唯一通道/)
  assert.match(skills[0].description, /不要用 python-pptx/)
  assert.match(skills[0].content, /以本技能为准/)
})

let deckId = ''
/** Deck directory (`<workspace>/.pptx-studio/<deckId>`) derived from the export receipt. */
let deckDir = ''
/** Element created through `POST /edit` and reused by later ops. */
let addedElementId = ''
/** Slide inserted through `POST /edit`. */
let insertedSlideId = ''
/** Second deck with a single slide (last-slide guard + lazy download). */
let singleDeckId = ''
/** Deck-relative src of the uploaded image. */
let uploadedSrc = ''
/** Deck id of the template-mode deck. */
let tplDeckId = ''
/** Its directory on disk. */
let tplDeckDir = ''
/** Requirement decomposition used by the template scenario. */
let tplOutline = []
await test('research → HTML preview gates the flow, then slides draw three pages', async () => {
  const opened = await tool('pptx_studio_open').execute({ title: 'E2E deck' }, exec)
  assert.equal(opened.ok, true)
  assert.equal(opened.stage, 'research')
  deckId = opened.deckId
  assert.match(deckId, /^deck_\d{8}_[0-9a-f]{6}$/)

  // A fresh deck cannot skip the analysis step.
  const blocked = await tool('pptx_studio_plan').execute({ deck: deckId, slides: [{ title: 'P1' }] }, exec)
  assert.equal(blocked.ok, false)
  assert.equal(blocked.code, 'stage-research')
  assert.match(blocked.hint, /pptx_studio_research/)
  assert.ok(lossless(blocked))
  // A refusal has to reach the model with its reasons, not just a code.
  const blockedCard = tool('pptx_studio_plan').output.render({}, blocked)
  assert.match(blockedCard[0].text, /stage-research/)
  assert.match(blockedCard[0].text, /pptx_studio_research/)

  const researched = await tool('pptx_studio_research').execute({
    deck: deckId,
    brief: {
      summary: '受众是管理层，先给结论再给数据。',
      audience: '管理层',
      goal: '批准预算',
      outline: [{ title: '背景', intent: '为什么现在做', source: '内部访谈' }, { title: '方案', intent: '怎么落地' }],
      facts: [{ claim: '延迟下降 42%', source: '压测报告' }],
      sources: [{ title: '行业报告', url: 'https://example.com/report', note: '2026 Q1' }],
      palette: ['#22D3EE', '#060B18'],
      fonts: { latin: '微软雅黑', heading: '微软雅黑' },
    },
  }, exec)
  assert.equal(researched.ok, true, researched.hint || 'research failed')
  assert.equal(researched.stage, 'preview')
  assert.equal(researched.outline, 2)
  assert.equal(researched.sources, 1)
  assert.ok(lossless(researched))

  // Research alone is not enough: the HTML visual must exist too.
  const blockedAgain = await tool('pptx_studio_plan').execute({ deck: deckId, slides: [{ title: 'P1' }] }, exec)
  assert.equal(blockedAgain.code, 'stage-preview')

  const previewed = await tool('pptx_studio_preview').execute({
    deck: deckId,
    title: 'E2E 视觉稿',
    html: '<!doctype html><html><head><style>.slide{width:960px;height:540px;background:#060B18;color:#fff}</style></head><body><section class="slide">E2E 视觉稿</section></body></html>',
    notes: '深色底 + 青色强调',
    width: 960,
    height: 540,
  }, exec)
  assert.equal(previewed.ok, true, previewed.hint || 'preview failed')
  assert.equal(previewed.stage, 'build')
  assert.ok(previewed.previewBytes > 0)
  assert.ok(existsSync(previewed.path), 'preview.html must be written next to deck.json')
  assert.ok(lossless(previewed))

  // The preview is served as an opaque, script-free sandbox.
  const served = await call('GET', `/preview?deck=${encodeURIComponent(deckId)}`)
  assert.equal(served.statusCode, 200)
  assert.match(String(served.headers['content-type']), /^text\/html/)
  assert.match(String(served.headers['content-security-policy']), /sandbox/)
  assert.ok(served.text.includes('E2E 视觉稿'))

  // The theme tool changes deck-level settings the canvas follows.
  const deckSettings = await tool('pptx_studio_deck').execute({
    deck: deckId,
    theme: { colors: { accent: '#22D3EE' }, fonts: { latin: '微软雅黑' } },
  }, exec)
  assert.equal(deckSettings.ok, true)
  assert.equal(deckSettings.theme.colors.accent, '#22D3EE')

  const planned = await tool('pptx_studio_plan').execute({ deck: deckId, slides: [{ title: 'P1' }, { title: 'P2' }, { title: 'P3' }] }, exec)
  assert.equal(planned.total, 3)
  assert.match(tool('pptx_studio_plan').output.render({}, planned)[0].text, /共 3 页/, 'plan receipt must state the page count')

  for (let index = 1; index <= 3; index += 1) {
    const written = await tool('pptx_studio_slide').execute({
      deck: deckId,
      index,
      slide: {
        transition: { type: 'fade', speed: 'fast' },
        notes: `第 ${index} 页的演讲备注`,
        elements: [{ id: 'title', type: 'text', box: [0.8, 0.6, 11.7, 0.9], text: `第 ${index} 页`, style: { size: 30, bold: true }, anim: { effect: 'fade' } }],
      },
    }, exec)
    assert.equal(written.ok, true)
    assert.equal(written.index, index)
    assert.equal(written.total, 3)
    assert.ok(lossless(written), 'slide result must be lossless JSON')
  }

  const inspected = await tool('pptx_studio_inspect').execute({ deck: deckId }, exec)
  assert.equal(inspected.deck.pipeline.stage, 'build')
  assert.equal(inspected.deck.pipeline.research, true)
  assert.equal(inspected.deck.pipeline.preview, true)
})

await test('every page is announced before it is written, so the canvas can draw it', async () => {
  const paced = await tool('pptx_studio_open').execute({ title: 'Pacing deck' }, exec)
  assert.equal(paced.ok, true)
  const pacedDeck = paced.deckId

  // Subscribe before the writes: the SSE stream is the user's view of a build.
  const stream = fakeResponse()
  stream.on = () => {}
  const req = fakeRequest('GET', `/dsh-pptx-studio/api/events?deck=${pacedDeck}`)
  await route.handler(req, stream)
  const lookAgain = await tool('pptx_studio_open').execute({ deckId: pacedDeck, title: 'Pacing deck' }, exec)
  assert.equal(lookAgain.ok, true)

  await tool('pptx_studio_plan').execute({ deck: pacedDeck, slides: [{ title: 'A' }, { title: 'B' }], force: true }, exec)
  const startedAt = Date.now()
  await tool('pptx_studio_slide').execute({
    deck: pacedDeck,
    index: 1,
    slide: { title: 'A', elements: [{ type: 'text', box: [1, 1, 4, 1], text: 'hello' }] },
    force: true,
  }, exec)
  const slideMs = Date.now() - startedAt
  await tool('pptx_studio_add_element').execute({
    deck: pacedDeck,
    slideId: 's1',
    element: { type: 'text', box: [1, 3, 4, 1], text: 'more' },
  }, exec)

  /** @type {any[]} */
  const sent = []
  for (const chunk of stream.chunks) {
    const line = String(Buffer.from(chunk)).split('\n').find((row) => row.startsWith('data: '))
    if (line) sent.push(JSON.parse(line.slice(6)))
  }
  const statusIndex = (phase, extra) => sent.findIndex((event) => event.type === 'status' && event.phase === phase && Object.entries(extra).every(([key, value]) => event[key] === value))
  const slideIndex = (phase, index) => sent.findIndex((event) => event.type === 'slide' && event.phase === phase && event.index === index)

  assert.ok(statusIndex('ready', {}) >= 0, 'opening a deck must announce that the canvas is ready')
  const planning = statusIndex('planning', { begin: true })
  assert.ok(planning >= 0, 'plan must announce that it is laying out the skeleton pages')
  assert.ok(planning < slideIndex('skeleton', 0), 'the outline announcement must come before the first skeleton page')
  assert.equal(sent.filter((event) => event.type === 'slide' && event.phase === 'skeleton').length, 2, 'both skeleton pages must be streamed')

  const begin = statusIndex('drawing', { begin: true, op: 'slide', slideIndex: 1 })
  const content = slideIndex('content', 0)
  assert.ok(begin >= 0, 'a page write must start with a begin status')
  assert.ok(begin < content, 'the begin status must precede the page itself — otherwise it lands in one frame')
  assert.ok(slideMs >= 180, `the page write must hold the pencil beat (took ${slideMs}ms)`)

  const addBegin = statusIndex('drawing', { begin: true, op: 'element', slideIndex: 1 })
  assert.ok(addBegin >= 0, 'adding one element must announce it too')
  assert.ok(addBegin < sent.findIndex((event) => event.type === 'element' && event.slideId === 's1'), 'the element announcement must come first')
  assert.ok(sent.some((event) => event.type === 'element' && event.slideId === 's1'), 'the element event still has to arrive')
})

await test('add_element appends and patch merges', async () => {
  const added = await tool('pptx_studio_add_element').execute({
    deck: deckId,
    slideId: 's2',
    element: { type: 'rect', box: [1, 2, 4, 1.2], fill: '#F1F5F9', text: '卡片' },
  }, exec)
  assert.equal(added.ok, true)
  assert.equal(added.elements, 2)

  const patched = await tool('pptx_studio_patch').execute({
    deck: deckId,
    slideId: 's2',
    elementId: 's2-e2',
    patch: { text: '改过的卡片', style: { size: 20 } },
  }, exec)
  assert.equal(patched.ok, true)
  assert.equal(patched.element.text, '改过的卡片')
  assert.equal(patched.element.style.size, 20)
  assert.ok(lossless(patched))
})

await test('data errors refuse the write; repairable warnings still draw', async () => {
  // Error level (unsupported type, five-number box): refused, deck untouched.
  const refused = await tool('pptx_studio_slide').execute({
    deck: deckId,
    index: 3,
    slide: { elements: [{ type: 'video', box: [1, 1, 2, 0.8, 'oops'], text: 'bad' }] },
  }, exec)
  assert.equal(refused.ok, false, 'an error-level payload must not be written')
  assert.equal(refused.code, 'invalid-data')
  assert.match(refused.hint, /type=/, 'the receipt names the offending field')
  assert.match(refused.hint, /box 必须是 4 个数字/)
  assert.ok(lossless(refused), 'the refusal must survive the snapshot rule')
  const after = await tool('pptx_studio_inspect').execute({ deck: deckId }, exec)
  assert.equal(after.deck.slideCount, 3, 'a refused write must not add the page')
  assert.ok(!after.deck.slides.some((s) => s.elements.some((el) => el.text === 'bad')), 'a refused element must not land anywhere')

  // Warning level (unknown field): repaired, drawn, and named.
  const drawn = await tool('pptx_studio_slide').execute({
    deck: deckId,
    index: 3,
    slide: { elements: [{ id: 't', type: 'text', box: [0.8, 0.6, 8, 0.9], text: 'ok', style: { size: 24 }, bogus: 1 }] },
  }, exec)
  assert.equal(drawn.ok, true, 'a repairable payload must still draw')
  assert.equal(drawn.elements, 1)
  assert.ok(Array.isArray(drawn.issues) && drawn.issues.some((line) => /未知字段 "bogus"/.test(line)), `issues: ${JSON.stringify(drawn.issues)}`)
  assert.ok(lossless(drawn), 'issues must survive the snapshot rule')
  const rendered = tool('pptx_studio_slide').output.render({}, drawn)
  assert.match(rendered[0].text, /参数已修正/)

  const clean = await tool('pptx_studio_slide').execute({
    deck: deckId,
    index: 3,
    slide: { elements: [{ id: 't', type: 'text', box: [0.8, 0.6, 8, 0.9], text: 'ok', style: { size: 24 } }] },
  }, exec)
  assert.equal('issues' in clean, false, 'a clean payload must not carry an issues key')
})

await test('inspect reports the deck model without undefined values', async () => {
  const result = await tool('pptx_studio_inspect').execute({ deck: deckId }, exec)
  assert.equal(result.ok, true)
  assert.equal(result.deck.slideCount, 3)
  assert.equal(result.deck.slides[0].title, '第 1 页')
  assert.ok(lossless(result), 'inspect result must be lossless JSON')
})

await test('export writes a real .pptx with transitions, animations and notes', async () => {
  // The canvas-edit checks cleared s1's transition on purpose; restore one so
  // this export covers every slide.
  const restored = await call('POST', '/edit', {
    deckId,
    ops: [
      { kind: 'slide.patch', slideId: 's1', patch: { transition: { type: 'fade', speed: 'fast' }, notes: '第 1 页的演讲备注' } },
      { kind: 'slide.patch', slideId: 's3', patch: { transition: { type: 'fade', speed: 'fast' }, notes: '第 3 页的演讲备注' } },
      { kind: 'element.patch', slideId: 's3', elementId: 't', patch: { anim: { effect: 'fade', duration: 600 } } },
    ],
  })
  assert.equal(restored.statusCode, 200)

  const result = await tool('pptx_studio_export').execute({ deck: deckId }, exec)
  assert.equal(result.ok, true, result.hint || 'export failed')
  assert.ok(existsSync(result.path), `missing ${result.path}`)
  assert.ok(result.bytes > 5000)
  assert.equal(result.slides, 3)
  assert.equal(result.features.transitions, 3, 'every slide carries the fade transition')
  assert.ok(result.features.animations >= 3, 'every title carries a fade entrance animation')
  assert.equal(result.features.notes, 3, 'notes are written into the notes slides')
  assert.ok(result.pipeline.builtAt, 'a successful export marks the pipeline as built')
  assert.ok(lossless(result), 'export receipt must survive the snapshot rule')
  deckDir = dirname(result.path)
})

await test('failure receipts are model-actionable and lossless', async () => {
  const missing = await tool('pptx_studio_inspect').execute({ deck: 'deck_nope' }, exec)
  assert.equal(missing.ok, false)
  assert.equal(missing.code, 'deck-not-found')
  assert.ok(missing.hint.length > 0)
  assert.ok(lossless(missing))
})

await test('POST /edit applies patch/add/slide.add/title and persists them', async () => {
  const response = await call('POST', '/edit', {
    deckId,
    ops: [
      { kind: 'element.patch', slideId: 's2', elementId: 's2-e2', patch: { text: 'route edited' } },
      { kind: 'element.add', slideId: 's1', element: { type: 'rect', box: [1, 1, 3, 1], text: 'added via route', fill: '#eeeeee' } },
      { kind: 'slide.add', index: 1, slide: { title: '被插入的页', elements: [] } },
      { kind: 'deck.title', title: 'Route edited deck' },
    ],
  })
  assert.equal(response.statusCode, 200)
  const data = response.json
  assert.equal(data.ok, true)
  assert.equal(data.results.length, 4)
  for (const result of data.results) assert.equal(result.ok, true)

  assert.equal(data.results[0].kind, 'element.patch')
  assert.equal(data.results[0].element.text, 'route edited')
  assert.equal(data.results[1].element.type, 'rect')
  assert.equal(data.results[1].element.text, 'added via route')
  assert.equal(data.results[1].index, 1)
  assert.match(data.results[1].elementId, /^s1-e\d+$/)
  addedElementId = data.results[1].elementId
  assert.match(data.results[2].slideId, /^s\d+$/)
  assert.equal(data.results[2].index, 1)
  assert.equal(data.results[2].total, 4)
  insertedSlideId = data.results[2].slideId
  assert.equal(data.results[3].title, 'Route edited deck')
  assert.equal(data.deckTitle, 'Route edited deck')

  // `emit()` schedules a debounced deck.json write; wait for it to land.
  await new Promise((resolvePromise) => setTimeout(resolvePromise, 700))
  const persisted = JSON.parse(readFileSync(join(deckDir, 'deck.json'), 'utf8'))
  assert.equal(persisted.title, 'Route edited deck')
  assert.equal(persisted.slides.length, 4)
  assert.equal(persisted.slides[1].id, insertedSlideId)
  const patchedSlide = persisted.slides.find((slide) => slide.id === 's2')
  assert.equal(patchedSlide.elements.find((element) => element.id === 's2-e2').text, 'route edited')
})

await test('POST /edit supports replace/duplicate/reorder/remove and slide move/remove', async () => {
  const response = await call('POST', '/edit', {
    deckId,
    ops: [
      { kind: 'element.replace', slideId: 's1', elementId: addedElementId, element: { type: 'rect', box: [2, 2, 2, 1], text: 'replaced', fill: '#ffffff' } },
      { kind: 'element.duplicate', slideId: 's1', elementId: addedElementId, offset: [0.5, 0.3] },
      { kind: 'element.reorder', slideId: 's1', elementId: `${addedElementId}-copy`, to: 'back' },
      { kind: 'element.reorder', slideId: 's1', elementId: `${addedElementId}-copy`, to: 'front' },
      { kind: 'element.patch', slideId: 's1', elementId: `${addedElementId}-copy`, remove: true },
      { kind: 'slide.move', slideId: insertedSlideId, index: 3 },
      { kind: 'slide.remove', slideId: insertedSlideId },
    ],
  })
  assert.equal(response.statusCode, 200)
  const { results } = response.json
  assert.equal(results.length, 7)
  for (const result of results) assert.equal(result.ok, true)

  assert.equal(results[0].elementId, addedElementId)
  assert.equal(results[0].element.text, 'replaced')
  assert.deepEqual(results[0].element.box, [2, 2, 2, 1])

  assert.equal(results[1].elementId, `${addedElementId}-copy`)
  assert.equal(results[1].element.id, `${addedElementId}-copy`)
  assert.equal(results[1].index, 2, 'the copy lands right after its source')
  assert.deepEqual(results[1].element.box, [2.5, 2.3, 2, 1])

  assert.equal(results[2].index, 0, 'reorder back puts the element first')
  assert.equal(results[3].index, 2, 'reorder front puts the element last')

  assert.equal(results[4].removed, true)
  assert.equal(results[5].index, 3)
  assert.equal(results[5].total, 4)
  assert.equal(results[6].removed, true)
  assert.equal(results[6].total, 3)

  await new Promise((resolvePromise) => setTimeout(resolvePromise, 700))
  const persisted = JSON.parse(readFileSync(join(deckDir, 'deck.json'), 'utf8'))
  assert.equal(persisted.slides.length, 3)
  assert.equal(persisted.slides.some((slide) => slide.id === insertedSlideId), false)
  const target = persisted.slides.find((slide) => slide.id === 's1')
  assert.equal(target.elements.length, 2)
  assert.equal(target.elements.find((element) => element.id === addedElementId).text, 'replaced')
})

await test('POST /edit patches slide effects/notes/background and deck theme/size', async () => {
  const response = await call('POST', '/edit', {
    deckId,
    ops: [
      {
        kind: 'slide.patch',
        slideId: 's1',
        patch: {
          transition: { type: 'push', speed: 'slow', direction: 'up', advanceOnClick: false, advanceAfterMs: 1500 },
          notes: '这段是画布上改的备注',
          background: { color: '#060B18' },
          hidden: true,
        },
      },
      { kind: 'deck.theme', theme: { colors: { accent: '#22D3EE' }, fonts: { heading: '微软雅黑' } } },
      { kind: 'deck.size', width: 10, height: 7.5 },
    ],
  })
  assert.equal(response.statusCode, 200)
  const { results } = response.json
  assert.equal(results.length, 3)
  assert.equal(results[0].slide.transition.type, 'push')
  assert.equal(results[0].slide.transition.advanceAfterMs, 1500)
  assert.equal(results[0].slide.notes, '这段是画布上改的备注')
  assert.deepEqual(results[0].slide.background, { color: '#060B18' })
  assert.equal(results[0].slide.hidden, true)
  assert.equal(results[1].theme.colors.accent, '#22D3EE')
  assert.equal(results[2].size.w, 10)
  assert.ok(lossless(results), 'edit results must be lossless')

  await new Promise((resolvePromise) => setTimeout(resolvePromise, 700))
  const persisted = JSON.parse(readFileSync(join(deckDir, 'deck.json'), 'utf8'))
  const slide = persisted.slides.find((candidate) => candidate.id === 's1')
  assert.equal(slide.transition.type, 'push')
  assert.equal(slide.notes, '这段是画布上改的备注')
  assert.equal(persisted.theme.colors.accent, '#22D3EE')

  // Clearing a transition must remove the key, not leave `null` behind.
  const cleared = await call('POST', '/edit', {
    deckId,
    ops: [{ kind: 'slide.patch', slideId: 's1', patch: { transition: null } }],
  })
  assert.equal(cleared.statusCode, 200)
  assert.equal('transition' in cleared.json.results[0].slide, false)

  // Restore the 16:9 canvas for the remaining geometry-sensitive checks.
  const restored = await call('POST', '/edit', {
    deckId,
    ops: [{ kind: 'deck.size', width: 13.333, height: 7.5 }],
  })
  assert.equal(restored.statusCode, 200)
  assert.equal(restored.json.results[0].size.w, 13.333)
})

await test('preview is optional until submitted, then served as a sandbox', async () => {
  const fresh = await tool('pptx_studio_open').execute({ title: '无预览的 deck' }, exec)
  const missing = await call('GET', `/preview?deck=${encodeURIComponent(fresh.deckId)}`)
  assert.equal(missing.statusCode, 404)
  assert.equal(missing.json.code, 'preview-missing')

  const empty = await tool('pptx_studio_preview').execute({ deck: fresh.deckId, html: '   ' }, exec)
  assert.equal(empty.ok, false)
  assert.equal(empty.code, 'preview-required')

  const forced = await tool('pptx_studio_plan').execute({ deck: fresh.deckId, slides: [{ title: '快速草稿' }], force: true }, exec)
  assert.equal(forced.ok, true)
  assert.equal(forced.total, 1)
})

await test('POST /edit refuses the last slide and unknown ops', async () => {
  const opened = await tool('pptx_studio_open').execute({ title: 'Single slide deck' }, exec)
  assert.equal(opened.ok, true)
  singleDeckId = opened.deckId

  const added = await call('POST', '/edit', { deckId: singleDeckId, ops: [{ kind: 'slide.add', index: 0 }] })
  assert.equal(added.statusCode, 200)
  assert.equal(added.json.results[0].slideId, 's1')
  assert.equal(added.json.results[0].total, 1)

  const refused = await call('POST', '/edit', { deckId: singleDeckId, ops: [{ kind: 'slide.remove', slideId: 's1' }] })
  assert.equal(refused.statusCode, 400)
  assert.equal(refused.json.ok, false)
  assert.equal(refused.json.code, 'edit-failed')
  assert.equal(refused.json.results.at(-1).code, 'last-slide')

  const unknown = await call('POST', '/edit', { deckId: singleDeckId, ops: [{ kind: 'element.explode' }] })
  assert.equal(unknown.statusCode, 400)
  assert.equal(unknown.json.ok, false)
  assert.equal(unknown.json.code, 'unknown-op')

  const missing = await call('POST', '/edit', { deckId: 'deck_missing', ops: [] })
  assert.equal(missing.statusCode, 404)
  assert.equal(missing.json.code, 'deck-not-found')
})

/** A real 1x1 transparent PNG. */
const PNG_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
const PNG_BYTES = Buffer.from(PNG_BASE64, 'base64')

await test('POST /upload stores one deduped asset and adds an image element', async () => {
  const dataUrl = `data:image/png;base64,${PNG_BASE64}`
  const first = await call('POST', '/upload', { deckId, slideId: 's1', name: 'dot.png', dataUrl, box: [1, 1, 2, 2] })
  assert.equal(first.statusCode, 200)
  const data = first.json
  assert.equal(data.ok, true)
  assert.equal(data.slideId, 's1')
  assert.equal(data.element.type, 'image')
  assert.equal(data.element.name, 'dot.png')
  assert.equal(data.element.src, data.src)
  assert.deepEqual(data.element.box, [1, 1, 2, 2])
  assert.match(data.element.id, /^s1-e\d+$/)
  uploadedSrc = data.src

  const hash = createHash('sha1').update(PNG_BYTES).digest('hex').slice(0, 8)
  assert.equal(uploadedSrc, `assets/img_${hash}.png`)
  const assetsDir = join(deckDir, 'assets')
  const file = join(assetsDir, `img_${hash}.png`)
  assert.ok(existsSync(file), `missing ${file}`)
  assert.deepEqual(readFileSync(file), PNG_BYTES)

  const before = readdirSync(assetsDir).length
  const second = await call('POST', '/upload', { deckId, slideId: 's1', dataUrl })
  assert.equal(second.statusCode, 200)
  assert.equal(second.json.src, uploadedSrc)
  assert.equal(readdirSync(assetsDir).length, before, 'the same bytes must reuse one file')

  await new Promise((resolvePromise) => setTimeout(resolvePromise, 700))
  const persisted = JSON.parse(readFileSync(join(deckDir, 'deck.json'), 'utf8'))
  const slide = persisted.slides.find((candidate) => candidate.id === 's1')
  assert.ok(slide.elements.some((element) => element.type === 'image' && element.src === uploadedSrc))
})

await test('POST /upload rejects non-images and oversized payloads', async () => {
  const notAnImage = await call('POST', '/upload', { deckId, dataUrl: 'data:text/plain;base64,aGVsbG8=' })
  assert.equal(notAnImage.statusCode, 400)
  assert.equal(notAnImage.json.code, 'bad-upload')

  const brokenBase64 = await call('POST', '/upload', { deckId, dataUrl: 'data:image/png;base64,!!!!' })
  assert.equal(brokenBase64.statusCode, 400)
  assert.equal(brokenBase64.json.code, 'bad-upload')

  const huge = Buffer.alloc(20 * 1024 * 1024 + 3).toString('base64')
  const tooLarge = await call('POST', '/upload', { deckId, dataUrl: `data:image/png;base64,${huge}` })
  assert.equal(tooLarge.statusCode, 413)
  assert.equal(tooLarge.json.code, 'too-large')
})

await test('GET /asset serves deck assets by name and workspace files by path', async () => {
  const name = uploadedSrc.split('/').pop()
  const byName = await call('GET', `/asset?deck=${encodeURIComponent(deckId)}&name=${encodeURIComponent(name)}`)
  assert.equal(byName.statusCode, 200)
  assert.equal(byName.headers['content-type'], 'image/png')
  assert.equal(byName.headers['cache-control'], 'no-store')
  assert.equal(Number(byName.headers['content-length']), byName.body.length)
  assert.deepEqual(byName.body, PNG_BYTES)

  // `src` form (`assets/<file>`) resolves against the deck directory. A
  // workspace-level `assets/` folder with the same name must not shadow it.
  mkdirSync(join(workspace, 'assets'), { recursive: true })
  writeFileSync(join(workspace, 'assets', name), 'decoy')
  const bySrc = await call('GET', `/asset?deck=${encodeURIComponent(deckId)}&path=${encodeURIComponent(uploadedSrc)}`)
  assert.equal(bySrc.statusCode, 200)
  assert.deepEqual(bySrc.body, PNG_BYTES)

  writeFileSync(join(workspace, 'note.txt'), 'workspace note')
  const byPath = await call('GET', `/asset?deck=${encodeURIComponent(deckId)}&path=note.txt`)
  assert.equal(byPath.statusCode, 200)
  assert.equal(byPath.text, 'workspace note')

  const absolute = await call('GET', `/asset?deck=${encodeURIComponent(deckId)}&path=${encodeURIComponent(join(workspace, 'note.txt'))}`)
  assert.equal(absolute.statusCode, 200)
  assert.equal(absolute.text, 'workspace note')

  const traversalName = await call('GET', `/asset?deck=${encodeURIComponent(deckId)}&name=${encodeURIComponent('../deck.json')}`)
  assert.equal(traversalName.statusCode, 400)
  assert.equal(traversalName.json.code, 'bad-asset')

  const escape = await call('GET', `/asset?deck=${encodeURIComponent(deckId)}&path=${encodeURIComponent('../../outside.txt')}`)
  assert.equal(escape.statusCode, 403)
  assert.equal(escape.json.code, 'asset-forbidden')

  const missing = await call('GET', `/asset?deck=${encodeURIComponent(deckId)}&name=nope.png`)
  assert.equal(missing.statusCode, 404)
  assert.equal(missing.json.code, 'asset-not-found')
})

await test('GET /state reports the workspace cwd and deck assets dir', async () => {
  const response = await call('GET', `/state?deck=${encodeURIComponent(deckId)}`)
  assert.equal(response.statusCode, 200)
  const data = response.json
  assert.equal(data.ok, true)
  assert.equal(data.deck.id, deckId)
  assert.equal(data.cwd, workspace)
  assert.equal(data.assets, join(deckDir, 'assets'))
  assert.equal(data.preview, join(deckDir, 'preview.html'))
  assert.equal(data.previewExists, true)
  assert.equal(data.deck.pipeline.stage, 'build')
  assert.equal(data.deck.pipeline.preview.width, 960)
  assert.equal(data.deck.pipeline.preview.height, 540)
  assert.ok(Array.isArray(data.deck.slides))
})

await test('POST /export and GET /download serve pptx bytes (lazy export included)', async () => {
  const exported = await call('POST', '/export', { deckId })
  assert.equal(exported.statusCode, 200)
  assert.equal(exported.json.ok, true)
  assert.ok(exported.json.value.bytes > 5000)
  assert.equal(exported.json.value.download, `/dsh-pptx-studio/api/download?deck=${deckId}`)

  const download = await call('GET', `/download?deck=${encodeURIComponent(deckId)}`)
  assert.equal(download.statusCode, 200)
  assert.equal(download.headers['content-type'], 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
  assert.match(String(download.headers['content-disposition']), /^attachment; filename=".*\.pptx"/)
  assert.equal(download.headers['cache-control'], 'no-store')
  assert.equal(Number(download.headers['content-length']), download.body.length)
  assert.equal(download.body.subarray(0, 2).toString('latin1'), 'PK')

  // This deck was never exported and holds an empty page: the export gate
  // must refuse it (a blank page in the deliverable is a data mistake), then
  // the same deck must export once the page has content.
  const gated = await call('GET', `/download?deck=${encodeURIComponent(singleDeckId)}`)
  assert.equal(gated.statusCode, 500, 'an empty page must not reach the .pptx')
  assert.equal(gated.json.code, 'invalid-data')
  const filled = await call('POST', '/edit', { deckId: singleDeckId, ops: [{ kind: 'element.add', slideId: 's1', element: { type: 'text', box: [1, 1, 4, 1], text: 'lazy export' } }] })
  assert.equal(filled.statusCode, 200)
  const lazy = await call('GET', `/download?deck=${encodeURIComponent(singleDeckId)}`)
  assert.equal(lazy.statusCode, 200)
  assert.equal(lazy.body.subarray(0, 2).toString('latin1'), 'PK')

  const missing = await call('GET', '/download?deck=deck_missing')
  assert.equal(missing.statusCode, 404)
  assert.equal(missing.json.code, 'deck-not-found')
})

await test('export accepts a folder the user picked and never overwrites a file', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'pptx-studio-out-'))
  try {
    const first = await call('POST', '/export', { deckId, dir })
    assert.equal(first.statusCode, 200)
    assert.equal(first.json.ok, true)
    assert.equal(dirname(first.json.value.path), dir, 'the file must land in the picked folder')
    assert.match(first.json.value.path, /\.pptx$/)
    assert.ok(existsSync(first.json.value.path), 'the export must actually write the file')

    const second = await call('POST', '/export', { deckId, dir })
    assert.equal(second.statusCode, 200)
    assert.notEqual(second.json.value.path, first.json.value.path, 'a second save must not eat the first file')
    assert.match(second.json.value.path, / \(1\)\.pptx$/)

    const relative = await call('POST', '/export', { deckId, dir: 'relative/folder' })
    assert.equal(relative.statusCode, 400)
    assert.equal(relative.json.code, 'bad-directory')

    const absent = await call('POST', '/export', { deckId, dir: join(dir, 'nope') })
    assert.equal(absent.statusCode, 400)
    assert.equal(absent.json.code, 'bad-directory')
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
})

await test('a deck written to disk is still listed and loadable after a plugin reload', async () => {
  // Simulate the state a restart leaves behind: deck.json on disk, nothing in
  // the (per-apply) in-memory store.
  const onDiskId = 'deck_ondisk_1'
  const onDiskDir = join(workspace, '.pptx-studio', onDiskId)
  mkdirSync(onDiskDir, { recursive: true })
  writeFileSync(join(onDiskDir, 'deck.json'), JSON.stringify({
    version: 1,
    id: onDiskId,
    title: '磁盘上的旧画布',
    size: { w: 13.333, h: 7.5 },
    slides: [{ id: 's1', elements: [{ id: 'e1', type: 'text', box: [1, 1, 4, 1], text: '重启前写的' }] }],
  }, null, 2), 'utf8')

  const listed = await call('GET', `/decks?sessionId=sess-e2e&cwd=${encodeURIComponent(workspace)}`)
  assert.equal(listed.statusCode, 200)
  const summary = listed.json.decks.find((item) => item.id === onDiskId)
  assert.ok(summary, `the deck on disk must be listed; got ${listed.json.decks.map((item) => item.id).join(',')}`)
  assert.equal(summary.fromDisk, true, 'an in-memory miss must be flagged as fromDisk')
  assert.equal(summary.slides, 1)
  assert.equal(summary.title, '磁盘上的旧画布')

  // A second apply() starts with an empty store: exactly what a reload leaves.
  /** @type {any} */
  let routeB = null
  /** @type {any[]} */
  const toolsB = []
  const ctxB = {
    ...ctx,
    tools: { register: (definition) => { toolsB.push(definition); return () => {} } },
    webServer: { register: (definition) => { routeB = definition; return () => {} } },
  }
  await plugin.apply(ctxB, {})
  const req = fakeRequest('GET', `/dsh-pptx-studio/api/state?deck=${onDiskId}&cwd=${encodeURIComponent(workspace)}`)
  const res = fakeResponse()
  await routeB.handler(req, res)
  assert.equal(res.statusCode, 200, 'state must reopen the deck from deck.json')
  assert.equal(res.json.deck.id, onDiskId)
  assert.equal(res.json.deck.slides[0].elements[0].text, '重启前写的')
  // A pre-pipeline deck with slides is grandfathered as already built.
  assert.equal(res.json.deck.pipeline.stage, 'build')
  assert.ok(res.json.deck.pipeline.builtAt)

  // Tools must also reopen a forgotten deck from disk (a host reload forgets
  // the in-memory map but not deck.json).
  const inspectB = toolsB.find((definition) => definition.name === 'pptx_studio_inspect')
  assert.ok(inspectB, 'the reloaded plugin must register inspect')
  const inspected = await inspectB.execute({ deck: onDiskId }, exec)
  assert.equal(inspected.ok, true, inspected.hint || 'inspect after reload failed')
  assert.equal(inspected.deck.id, onDiskId)

  // An unknown id still 404s instead of pretending to open something.
  const missing = await call('GET', `/state?deck=deck_nope_404&cwd=${encodeURIComponent(workspace)}`)
  assert.equal(missing.statusCode, 404)
})

await test('deck/delete drops memory + disk and refuses junk ids', async () => {
  // Disk-only deck: the directory exists, the store has never seen it.
  const victim = 'deck_20260101_abcdef'
  const dir = join(workspace, '.pptx-studio', victim)
  mkdirSync(dir, { recursive: true })
  writeFileSync(join(dir, 'deck.json'), JSON.stringify({
    version: 1, id: victim, title: '待删画布', size: { w: 13.333, h: 7.5 }, slides: [{ id: 's1', elements: [] }],
  }))

  const traversal = await call('POST', '/deck/delete', { deckId: '../../etc', cwd: workspace })
  assert.equal(traversal.statusCode, 400, 'a path-shaped id must be refused')

  const removed = await call('POST', '/deck/delete', { deckId: victim, cwd: workspace })
  assert.equal(removed.statusCode, 200)
  assert.equal(removed.json.removed, true)
  assert.ok(!existsSync(dir), 'the deck directory must be gone')

  const listed = await call('GET', `/decks?cwd=${encodeURIComponent(workspace)}`)
  assert.ok(!listed.json.decks.some((item) => item.id === victim), 'a deleted deck must leave the list')

  const again = await call('POST', '/deck/delete', { deckId: victim, cwd: workspace })
  assert.equal(again.statusCode, 404)

  // In-memory deck: the store must drop it too, or a later persist resurrects it.
  const opened = await tool('pptx_studio_open').execute({ title: '内存里的待删画布' }, exec)
  const liveId = opened.deckId
  const liveDir = join(workspace, '.pptx-studio', liveId)
  assert.ok(existsSync(liveDir), 'open() writes the directory')
  const gone = await call('POST', `/deck/delete?sessionId=sess-e2e`, { deckId: liveId })
  assert.equal(gone.statusCode, 200)
  assert.ok(!existsSync(liveDir), 'in-memory delete removes the directory')
  const after = await call('GET', '/decks?sessionId=sess-e2e')
  assert.ok(!after.json.decks.some((item) => item.id === liveId), 'in-memory delete removes the entry')
})

// ── template-first workflow ────────────────────────────────────────────────
// Built on a synthetic template (no PowerPoint needed: decomposition runs with
// render:false, which is the documented degraded mode). Every assertion here is
// about the guarantee the feature exists for: copy is mapped once, and the HTML
// sheet, the slides and the exported file cannot drift apart.

const assetsDir = new URL('../assets', import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1')
const pythonRunner = new PythonRunner({ assetsDir, log: () => {} })

/**
 * Whether this machine can build a synthetic template (needs python-pptx).
 * @returns {Promise<string|null>} interpreter path, or null when unavailable.
 */
async function templateInterpreter() {
  let python = null
  try {
    python = await pythonRunner.interpreter()
  } catch {
    return null
  }
  const probe = await runProcess(python, ['-c', 'import pptx; print(pptx.__version__)'], { timeoutMs: 60_000 })
  return probe.code === 0 ? python : null
}

/**
 * Write a 3-page template (cover / content / ending) into the workspace.
 * @param {string} python - interpreter with python-pptx.
 * @param {string} target - .pptx path.
 * @returns {Promise<void>} resolves once written.
 */
async function writeSyntheticTemplate(python, target) {
  const script = `
from pptx import Presentation
from pptx.util import Inches, Pt
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]

s0 = prs.slides.add_slide(blank)
box = s0.shapes.add_textbox(Inches(1.2), Inches(2.6), Inches(10.9), Inches(1.6))
tf = box.text_frame
tf.text = "封面标题"
tf.paragraphs[0].runs[0].font.size = Pt(44)

s1 = prs.slides.add_slide(blank)
t1 = s1.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.3), Inches(1.0))
t1.text_frame.text = "页标题"
t1.text_frame.paragraphs[0].runs[0].font.size = Pt(32)
t2 = s1.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.3), Inches(2.4))
t2.text_frame.word_wrap = True
t2.text_frame.text = "这里是正文内容示例，用来占位说明这一页要讲的信息。"
t2.text_frame.paragraphs[0].runs[0].font.size = Pt(16)

s2 = prs.slides.add_slide(blank)
t3 = s2.shapes.add_textbox(Inches(4.2), Inches(3.2), Inches(5.0), Inches(1.2))
t3.text_frame.text = "谢谢观看"
t3.text_frame.paragraphs[0].runs[0].font.size = Pt(40)
prs.save(r"${target.replace(/\\/g, '\\\\')}")
`
  const result = await runProcess(python, ['-c', script], { timeoutMs: 120_000 })
  assert.equal(result.code, 0, `template build failed: ${result.stderr}`)
}

const tplPython = await templateInterpreter()

if (!tplPython) {
  console.log('skip - no python-pptx on this machine (template e2e stays unproven)')
} else {
  await test('a template deck cannot skip decomposition', async () => {
    const templatePath = join(workspace, 'brand-template.pptx')
    await writeSyntheticTemplate(tplPython, templatePath)
    assert.ok(statSync(templatePath).size > 0, 'the synthetic template exists')

    const opened = await tool('pptx_studio_open').execute({ title: '模板流程 deck', template: 'brand-template.pptx' }, exec)
    assert.equal(opened.ok, true)
    assert.equal(opened.mode, 'template')
    assert.equal(opened.stage, 'template', 'a template deck starts at the template stage')
    assert.ok(lossless(opened))
    tplDeckId = opened.deckId
    tplDeckDir = join(workspace, '.pptx-studio', tplDeckId)

    // Neither research nor slides may come first.
    const earlyResearch = await tool('pptx_studio_research').execute({ deck: tplDeckId, brief: { summary: 'x', outline: [{ title: 'A' }] } }, exec)
    assert.equal(earlyResearch.ok, false)
    assert.equal(earlyResearch.code, 'stage-template')
    assert.match(earlyResearch.hint, /pptx_studio_template/)
    const earlyPlan = await tool('pptx_studio_plan').execute({ deck: tplDeckId, slides: [{ title: 'A' }] }, exec)
    assert.equal(earlyPlan.ok, false)
    assert.equal(earlyPlan.code, 'stage-template')
  })

  await test('a template outside the workspace is imported instead of refused', async () => {
    // Users point at ~/Downloads; the plugin must copy it in, not read ambient
    // paths later (and the old containment check let a cross-drive path through).
    const outside = join(tmpdir(), `pptx-outside-${Date.now()}.pptx`)
    await writeSyntheticTemplate(tplPython, outside)
    const opened = await tool('pptx_studio_open').execute({ title: '外部模板', template: outside }, exec)
    assert.equal(opened.ok, true, JSON.stringify(opened))
    assert.equal(opened.stage, 'template')
    const recorded = (await tool('pptx_studio_inspect').execute({ deck: opened.deckId }, exec)).deck.template
    assert.ok(recorded.startsWith(workspace), `the recorded template must live in the workspace: ${recorded}`)
    assert.ok(recorded.includes('_templates'), 'the imported copy lives in the plugin store')
    assert.ok(existsSync(recorded), 'the imported copy exists on disk')
    rmSync(recorded, { force: true })
  })

  await test('decomposition reports pages, roles and fillable fields', async () => {
    const decomposed = await tool('pptx_studio_template').execute({ deck: tplDeckId, pages: 3, render: false }, exec)
    assert.equal(decomposed.ok, true, JSON.stringify(decomposed))
    assert.equal(decomposed.template.pages, 3)
    assert.equal(decomposed.template.engine, 'none', 'render:false is the structural mode')
    assert.deepEqual(Object.keys(decomposed.template.roles).sort(), ['content', 'cover', 'ending'])
    assert.equal(decomposed.pages.length, 3)
    for (const page of decomposed.pages) assert.match(page.fields, /≤\d+字/)
    assert.ok(lossless(decomposed))
    assert.ok(existsSync(join(tplDeckDir, 'template', 'template.json')), 'the spec lands on disk')
    const pipeline = (await tool('pptx_studio_inspect').execute({ deck: tplDeckId }, exec)).deck.pipeline
    assert.equal(pipeline.stage, 'research', 'decomposition unlocks the analysis step')
    assert.equal(pipeline.template.decomposed, true)
  })

  await test('the canvas can read the template catalogue and design tokens', async () => {
    const catalogue = await call('GET', `/template?deck=${tplDeckId}&cwd=${encodeURIComponent(workspace)}`)
    assert.equal(catalogue.statusCode, 200)
    assert.equal(catalogue.json.pages.length, 3)
    assert.ok(Array.isArray(catalogue.json.design.palette))
    assert.ok(catalogue.json.pages[0].fields.length >= 1)
    assert.equal(catalogue.json.mapping, null, 'nothing mapped yet')
  })

  await test('requirement decomposition is refused when it does not fit the template', async () => {
    const wrongField = await tool('pptx_studio_research').execute({
      deck: tplDeckId,
      brief: { summary: '需求', outline: [{ title: '封面', templateSlide: 0, fields: { nope: 'x' } }] },
    }, exec)
    assert.equal(wrongField.ok, false)
    assert.equal(wrongField.code, 'template-mapping-invalid')
    assert.match(wrongField.hint, /没有这个字段/)

    const noMapping = await tool('pptx_studio_research').execute({
      deck: tplDeckId,
      brief: { summary: '需求', outline: [{ title: '封面' }] },
    }, exec)
    assert.equal(noMapping.ok, false)
    assert.equal(noMapping.code, 'template-mapping-required')
  })

  await test('research maps requirements onto template pages, then HTML is generated from it', async () => {
    const catalogue = (await call('GET', `/template?deck=${tplDeckId}&cwd=${encodeURIComponent(workspace)}`)).json.pages
    // Every field id comes from the template, and copy is sized to the slot so
    // the mapping check has nothing to complain about.
    tplOutline = catalogue.map((page, index) => ({
      title: index === 0 ? '敦煌夜游' : index === 1 ? '活动目的' : '谢谢',
      templateSlide: page.page,
      fields: Object.fromEntries(page.fields.map((field) => [
        field.id,
        (index === 0 && field.id === 'title' ? '敦煌夜游' : `文案${index + 1}`).slice(0, Math.max(1, field.maxCjk)),
      ])),
    }))
    const researched = await tool('pptx_studio_research').execute({
      deck: tplDeckId,
      brief: {
        summary: '把敦煌文化做成一场城市夜游。',
        audience: '城市青年',
        goal: '记住三个敦煌符号',
        outline: tplOutline,
        sources: [{ title: '敦煌研究院', url: 'https://www.e-dunhuang.com/', note: '素材来源' }],
      },
    }, exec)
    assert.equal(researched.ok, true, JSON.stringify(researched))
    assert.deepEqual(researched.templateMapping.templateSlides, [0, 1, 2])
    assert.ok(lossless(researched))

    const previewed = await tool('pptx_studio_preview').execute({ deck: tplDeckId, title: 'HTML 模版' }, exec)
    assert.equal(previewed.ok, true, JSON.stringify(previewed))
    assert.equal(previewed.mode, 'template')
    assert.equal(previewed.pages, 3)
    assert.deepEqual(previewed.templateSlides, [0, 1, 2])
    assert.ok(lossless(previewed))
  })

  await test('GET /preview renders the HTML template from the same mapping', async () => {
    const preview = await call('GET', `/preview?deck=${tplDeckId}&cwd=${encodeURIComponent(workspace)}`)
    assert.equal(preview.statusCode, 200)
    assert.equal((preview.text.match(/class="slide"/g) || []).length, 3)
    assert.match(preview.headers['content-security-policy'], /sandbox/)
    assert.match(preview.text, /data-field="title"/)
    assert.match(preview.text, /敦煌夜游/, 'the mapped copy is in the sheet')
    // No plates in structural mode: the page keeps the template background.
    assert.match(preview.text, /background:#/)
  })

  await test('slides built from the template inherit its geometry and provenance', async () => {
    const planned = await tool('pptx_studio_plan').execute({
      deck: tplDeckId,
      slides: tplOutline.map((item) => ({ title: item.title, templateSlide: item.templateSlide })),
    }, exec)
    assert.equal(planned.ok, true)
    assert.equal(planned.total, 3)

    for (const [index, item] of tplOutline.entries()) {
      const written = await tool('pptx_studio_slide').execute({
        deck: tplDeckId,
        index: index + 1,
        templateSlide: item.templateSlide,
        fields: item.fields,
      }, exec)
      assert.equal(written.ok, true, JSON.stringify(written))
      assert.equal(written.templateSlide, item.templateSlide)
    }
    const inspected = (await tool('pptx_studio_inspect').execute({ deck: tplDeckId }, exec)).deck
    assert.equal(inspected.slideCount, 3)
    assert.deepEqual(inspected.slides.map((slide) => slide.templateSlide), [0, 1, 2])
    assert.ok(inspected.slides[0].elements.some((el) => el.field === 'title'), 'field provenance is exposed to the model')

    // A slide write that names a field the page does not have never lands
    // silently: with nothing left to draw it refuses and says why.
    const bad = await tool('pptx_studio_slide').execute({
      deck: tplDeckId, index: 2, templateSlide: 1, fields: { notAField: 'x' },
    }, exec)
    assert.equal(bad.ok, false)
    assert.equal(bad.code, 'template-slide-empty')
    assert.match(bad.hint, /没有字段 "notAField"/)
    assert.match(bad.hint, /title/, 'the receipt lists the field ids the page does have')
    const kept = (await tool('pptx_studio_inspect').execute({ deck: tplDeckId }, exec)).deck.slides[1]
    assert.ok(kept.elements.some((el) => el.field === 'title'), 'a refused write leaves the previous page alone')
  })

  await test('verify passes, and export refuses a deck that drifted from its HTML', async () => {
    const verified = await tool('pptx_studio_verify').execute({ deck: tplDeckId }, exec)
    assert.equal(verified.ok, true, JSON.stringify(verified))
    assert.equal(verified.consistent, true)
    assert.equal(verified.templateMode, true)

    // Rewrite one slide's copy without touching the HTML sheet: that is exactly
    // the drift this feature exists to catch.
    const first = (await tool('pptx_studio_inspect').execute({ deck: tplDeckId }, exec)).deck.slides[0]
    const titleEl = first.elements.find((el) => el.field === 'title')
    const patched = await tool('pptx_studio_patch').execute({
      deck: tplDeckId, slideId: first.id, elementId: titleEl.id, patch: { text: '偷偷改掉的标题' },
    }, exec)
    assert.equal(patched.ok, true)

    const inconsistent = await tool('pptx_studio_verify').execute({ deck: tplDeckId }, exec)
    assert.equal(inconsistent.ok, false)
    assert.ok(inconsistent.issues.some((line) => /文字不一致/.test(line)), JSON.stringify(inconsistent.issues))
    const refused = await tool('pptx_studio_export').execute({ deck: tplDeckId }, exec)
    assert.equal(refused.ok, false)
    assert.equal(refused.code, 'template-inconsistent')
    assert.ok(lossless(refused))
    const refusedCard = tool('pptx_studio_export').output.render({}, refused)
    assert.match(refusedCard[0].text, /template-inconsistent/)
    assert.match(refusedCard[0].text, /文字不一致/, 'the rendered card carries the issue list')

    // Re-sync the sheet to the deck, then the export goes through.
    const resync = await tool('pptx_studio_preview').execute({
      deck: tplDeckId,
      pages: tplOutline.map((item, index) => ({
        templateSlide: item.templateSlide,
        title: item.title,
        fields: index === 0 ? { ...item.fields, title: '偷偷改掉的标题' } : item.fields,
      })),
    }, exec)
    assert.equal(resync.ok, true, JSON.stringify(resync))
    const clean = await tool('pptx_studio_verify').execute({ deck: tplDeckId }, exec)
    assert.equal(clean.ok, true, JSON.stringify(clean.issues))
    const exported = await tool('pptx_studio_export').execute({ deck: tplDeckId }, exec)
    assert.equal(exported.ok, true, JSON.stringify(exported))
    assert.equal(exported.slides, 3)
    assert.ok(existsSync(exported.path))
    assert.ok(existsSync(join(tplDeckDir, 'template', 'template.json')))
  })
}

await test('the session feed names the running step and the deck it writes', async () => {
  const quiet = await call('GET', '/session?sessionId=sess-quiet')
  assert.equal(quiet.json.ok, true)
  assert.equal(quiet.json.task, null, 'a session that never ran a pptx tool has no task')

  // The canvas opens *before* the deck exists, so the feed has to push: a
  // snapshot answered on demand would arrive too late to switch the tab.
  const stream = await call('GET', '/session/events?sessionId=sess-feed')
  const frames = () => stream.text.split('\n\n')
    .filter((row) => row.startsWith('data: '))
    .map((row) => JSON.parse(row.slice(6)))
  const opening = frames()
  assert.ok(opening.length >= 1, 'the stream must open with a snapshot')
  assert.equal(opening[0].type, 'task')
  assert.equal(opening[0].task, null, 'nothing runs in a fresh session yet')

  const feedExec = { agent: { session: { id: 'sess-feed', cwd: workspace } } }
  const opened = await tool('pptx_studio_open').execute({ title: 'Feed deck' }, feedExec)
  assert.equal(opened.ok, true)
  const events = frames().slice(opening.length).map((frame) => frame.task)
  assert.ok(events.length >= 2, 'a call must publish its start and its outcome')
  const started = events[0]
  assert.equal(started.tool, 'pptx_studio_open')
  assert.equal(started.state, 'running')
  assert.equal(started.label, '打开画布', 'the canvas needs a human label, not the wire name')
  assert.equal(started.deckId, '', 'the open call has no deck while it runs: the canvas must wait, not show the old one')
  const finished = events[events.length - 1]
  assert.equal(finished.state, 'done')
  assert.equal(finished.deckId, opened.deckId, 'the outcome carries the deck the canvas must switch to')
  assert.ok(finished.seq > started.seq, 'seq only grows, so a reconnecting canvas can order snapshots')

  // A step that writes an existing deck names it up front (from its arguments),
  // so the canvas switches before the tool returns.
  const before = frames().length
  const inspected = await tool('pptx_studio_inspect').execute({ deck: opened.deckId }, feedExec)
  assert.equal(inspected.ok, true)
  const inspectEvents = frames().slice(before).map((frame) => frame.task)
  assert.equal(inspectEvents[0].deckId, opened.deckId, 'a step that writes a deck must announce that deck at call time')
  assert.equal(inspectEvents[inspectEvents.length - 1].deckId, opened.deckId, 'and keep it after the call')

  // The on-demand snapshot a reconnecting tab asks for is the last task.
  const snapshot = await call('GET', '/session?sessionId=sess-feed')
  assert.equal(snapshot.json.task.tool, 'pptx_studio_inspect')
  assert.equal(snapshot.json.task.state, 'done')

  // A failing call must not leave the canvas claiming work is still running.
  const failed = await tool('pptx_studio_export').execute({ deck: 'deck_missing' }, feedExec)
  assert.equal(failed.ok, false)
  const failedSnapshot = await call('GET', '/session?sessionId=sess-feed')
  assert.equal(failedSnapshot.json.task.state, 'failed')
  assert.ok(failedSnapshot.json.task.message.length > 0, 'the failure carries the host reason')

  stream.request.emit('close')
})

rmSync(join(workspace, '.pptx-studio'), { recursive: true, force: true })
console.log(`\ne2e: ${passed} checks passed`)
