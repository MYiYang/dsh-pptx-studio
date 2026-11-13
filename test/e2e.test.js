/**
 * End-to-end harness: apply the real plugin against a fake ctx, run the
 * tool chain (open → plan → slide → add_element → patch → inspect → export),
 * and assert the results survive the registry's lossless-JSON snapshot rule.
 *
 * Run: node test/e2e.test.js
 */
import assert from 'node:assert/strict'
import { mkdtempSync, existsSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

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

/** @type {any[]} */
const tools = []
/** @type {Map<string, Set<(event: Record<string, unknown>) => void>>} */
const events = new Map()
const ctx = {
  logger: { info() {}, warn() {}, debug() {}, error() {} },
  tools: { register: (tool) => { tools.push(tool); return () => {} } },
  get(name) {
    if (name === 'tools') return ctx.tools
    return undefined
  },
  loader: {
    import: async (name) => {
      if (name === '@deepseek-ai/dsh-tools') return { defineTool: (definition) => definition }
      throw new Error(`unexpected import: ${name}`)
    },
  },
  webServer: { register: () => () => {} },
  connection: { requestRejection: () => undefined },
  effect: (fn) => { fn(); return () => {} },
}

const exec = { agent: { session: { id: 'sess-e2e', cwd: workspace } } }
/**
 * Look up one registered tool.
 * @param {string} name - tool name.
 * @returns {any} tool definition.
 */
const tool = (name) => tools.find((t) => t.name === name)

await test('apply registers seven pptx_studio tools', async () => {
  await plugin.apply(ctx, {})
  assert.equal(tools.length, 7)
  for (const name of ['pptx_studio_open', 'pptx_studio_plan', 'pptx_studio_slide', 'pptx_studio_add_element', 'pptx_studio_patch', 'pptx_studio_inspect', 'pptx_studio_export']) {
    assert.ok(tool(name), `missing ${name}`)
    assert.equal(typeof tool(name).execute, 'function')
    assert.equal(typeof tool(name).output.render, 'function', 'output.render is required by the registry')
  }
})

let deckId = ''
await test('open → plan → slide draws three pages and returns lossless JSON', async () => {
  const opened = await tool('pptx_studio_open').execute({ title: 'E2E deck' }, exec)
  assert.equal(opened.ok, true)
  deckId = opened.deckId
  assert.match(deckId, /^deck_\d{8}_[0-9a-f]{6}$/)

  const planned = await tool('pptx_studio_plan').execute({ deck: deckId, slides: [{ title: 'P1' }, { title: 'P2' }, { title: 'P3' }] }, exec)
  assert.equal(planned.total, 3)

  for (let index = 1; index <= 3; index += 1) {
    const written = await tool('pptx_studio_slide').execute({
      deck: deckId,
      index,
      slide: { elements: [{ id: 'title', type: 'text', box: [0.8, 0.6, 11.7, 0.9], text: `第 ${index} 页`, style: { size: 30, bold: true } }] },
    }, exec)
    assert.equal(written.ok, true)
    assert.equal(written.index, index)
    assert.equal(written.total, 3)
    assert.ok(lossless(written), 'slide result must be lossless JSON')
  }
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

await test('inspect reports the deck model without undefined values', async () => {
  const result = await tool('pptx_studio_inspect').execute({ deck: deckId }, exec)
  assert.equal(result.ok, true)
  assert.equal(result.deck.slideCount, 3)
  assert.equal(result.deck.slides[0].title, '第 1 页')
  assert.ok(lossless(result), 'inspect result must be lossless JSON')
})

await test('export writes a real .pptx and reports bytes', async () => {
  const result = await tool('pptx_studio_export').execute({ deck: deckId }, exec)
  assert.equal(result.ok, true, result.hint || 'export failed')
  assert.ok(existsSync(result.path), `missing ${result.path}`)
  assert.ok(result.bytes > 5000)
  assert.equal(result.slides, 3)
})

await test('failure receipts are model-actionable and lossless', async () => {
  const missing = await tool('pptx_studio_inspect').execute({ deck: 'deck_nope' }, exec)
  assert.equal(missing.ok, false)
  assert.equal(missing.code, 'deck-not-found')
  assert.ok(missing.hint.length > 0)
  assert.ok(lossless(missing))
})

rmSync(join(workspace, '.pptx-studio'), { recursive: true, force: true })
console.log(`\ne2e: ${passed} checks passed`)
