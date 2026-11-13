/**
 * End-to-end harness: apply the real plugin against a fake ctx, run the
 * tool chain (open → plan → slide → add_element → patch → inspect → export),
 * and assert the results survive the registry's lossless-JSON snapshot rule.
 *
 * Run: node test/e2e.test.js
 */
import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { Readable } from 'node:stream'

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

await test('apply registers ten pptx_studio tools and the web route', async () => {
  await plugin.apply(ctx, {})
  assert.equal(tools.length, 10)
  for (const name of [
    'pptx_studio_open', 'pptx_studio_research', 'pptx_studio_preview', 'pptx_studio_deck',
    'pptx_studio_plan', 'pptx_studio_slide', 'pptx_studio_add_element', 'pptx_studio_patch',
    'pptx_studio_inspect', 'pptx_studio_export',
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

await test('a malformed slide payload draws but reports what was fixed', async () => {
  const written = await tool('pptx_studio_slide').execute({
    deck: deckId,
    index: 3,
    slide: { elements: [{ type: 'video', box: [1, 1, 2, 0.8, 'oops'], text: 'bad' }] },
  }, exec)
  assert.equal(written.ok, true, 'a recoverable payload must still draw')
  assert.equal(written.elements, 1)
  assert.ok(Array.isArray(written.issues) && written.issues.length >= 2, `issues: ${JSON.stringify(written.issues)}`)
  assert.ok(lossless(written), 'issues must survive the snapshot rule')
  const rendered = tool('pptx_studio_slide').output.render({}, written)
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

  // This deck was never exported: /download must run the exporter first.
  const lazy = await call('GET', `/download?deck=${encodeURIComponent(singleDeckId)}`)
  assert.equal(lazy.statusCode, 200)
  assert.equal(lazy.body.subarray(0, 2).toString('latin1'), 'PK')

  const missing = await call('GET', '/download?deck=deck_missing')
  assert.equal(missing.statusCode, 404)
  assert.equal(missing.json.code, 'deck-not-found')
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

rmSync(join(workspace, '.pptx-studio'), { recursive: true, force: true })
console.log(`\ne2e: ${passed} checks passed`)
