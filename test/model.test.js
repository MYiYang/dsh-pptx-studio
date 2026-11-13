/**
 * Unit tests for the deck model (no DSH runtime needed).
 * Run: node test/model.test.js
 */
import assert from 'node:assert/strict'

import {
  applyElementPatch,
  auditElement,
  auditSlide,
  createDeck,
  normalizeAnimation,
  normalizeChart,
  normalizeColor,
  normalizeDeck,
  normalizePipeline,
  normalizePreview,
  normalizeShadow,
  normalizeSlide,
  normalizeTheme,
  normalizeTransition,
  pipelineSummary,
  summarizeDeck,
} from '../lib/model.js'

let passed = 0
/**
 * Run one named check.
 * @param {string} name - test name.
 * @param {() => void} fn - body.
 * @returns {void}
 */
function test(name, fn) {
  try {
    fn()
    passed += 1
    console.log(`ok - ${name}`)
  } catch (error) {
    console.error(`not ok - ${name}`)
    throw error
  }
}

test('createDeck defaults to 16:9 and no slides', () => {
  const deck = createDeck({ id: 'deck_x' })
  assert.equal(deck.size.w, 13.333)
  assert.equal(deck.size.h, 7.5)
  assert.deepEqual(deck.slides, [])
})

test('normalizeSlide assigns ids and clamps boxes', () => {
  const slide = normalizeSlide({ elements: [{ type: 'text', box: [-99, 0, 999, 0.5], text: 'x' }] }, { index: 0, slideW: 13.333, slideH: 7.5 })
  assert.equal(slide.id, 's1')
  assert.equal(slide.elements[0].id, 's1-e1')
  const [x, y, w, h] = slide.elements[0].box
  assert.ok(x >= -2 && y >= 0 && w > 0 && h > 0, 'box is clamped')
})

test('normalizeElement rejects unknown types but keeps text', () => {
  const slide = normalizeSlide({ elements: [{ id: 'a', type: 'video', text: 'hello' }] }, { index: 0, slideW: 13.333, slideH: 7.5 })
  assert.equal(slide.elements[0].type, 'text')
  assert.equal(slide.elements[0].text, 'hello')
})

test('normalizeColor makes every hex spelling CSS-safe', () => {
  assert.equal(normalizeColor('060B18'), '#060B18')
  assert.equal(normalizeColor('#060b18'), '#060B18')
  assert.equal(normalizeColor('abc'), '#AABBCC')
  assert.equal(normalizeColor('#ABC'), '#AABBCC')
  assert.equal(normalizeColor('rgba(0, 0, 0, .4)'), 'rgba(0, 0, 0, .4)')
  assert.equal(normalizeColor('  transparent '), 'transparent')
  assert.equal(normalizeColor(''), '')
  assert.equal(normalizeColor(null), '')
})

test('slide ingestion canonicalises fill, line and text colours', () => {
  const slide = normalizeSlide({
    elements: [
      { type: 'rect', box: [0, 0, 4, 2], fill: '060B18', line: '22d3ee' },
      { type: 'text', box: [0, 0, 4, 1], text: 'x', style: { color: 'FFFFFF' }, paragraphs: [{ text: 'y', color: '9fb2cb' }] },
    ],
  }, { index: 0, slideW: 13.333, slideH: 7.5 })
  assert.equal(slide.elements[0].fill, '#060B18')
  assert.equal(slide.elements[0].line, '#22D3EE')
  assert.equal(slide.elements[1].style.color, '#FFFFFF')
  assert.equal(slide.elements[1].paragraphs[0].color, '#9FB2CB')
})

test('auditSlide flags a box with the wrong arity', () => {
  const issues = auditSlide({ elements: [{ type: 'text', box: [1, 2, 3, 4, ''], text: 'x' }] })
  assert.equal(issues.length, 1)
  assert.match(issues[0], /box 必须是 4 个数字/)
})

test('auditSlide flags an empty field name (the broken-JSON shape)', () => {
  const issues = auditSlide({ elements: [{ type: 'text', box: [1, 2, 3, 1], text: 'x', '': null }] })
  assert.ok(issues.some((issue) => /空字段名/.test(issue)), `unexpected: ${issues.join(' | ')}`)
})

test('auditSlide stays quiet on a clean payload', () => {
  const issues = auditSlide({
    title: 'p',
    elements: [
      { id: 'a', type: 'text', box: [0.8, 0.6, 4, 0.9], text: 'hi', style: { size: 20, bold: true } },
      { type: 'rect', box: [0, 0, 13.333, 7.5], fill: '060B18' },
      { type: 'image', box: [1, 1, 2, 2], src: 'assets/a.png' },
    ],
  })
  assert.deepEqual(issues, [])
})

test('auditSlide reports missing, non-array and empty elements', () => {
  assert.ok(auditSlide({}).some((issue) => /缺少 elements/.test(issue)))
  assert.deepEqual(auditSlide({ elements: {} }), ['elements 必须是数组'])
  assert.deepEqual(auditSlide({ elements: [] }), ['elements 是空数组：这一页会是空白页'])
})

test('auditElement names the offender and clamps the wording to 1-based index', () => {
  const issues = auditElement({ type: 'video', box: [1, 1, 0, 1], text: 'x' }, 2)
  assert.ok(issues.some((issue) => /不受支持/.test(issue)))
  assert.ok(issues.some((issue) => /宽高必须大于 0/.test(issue)))
  assert.ok(issues.every((issue) => issue.startsWith('第 3 个元素')), issues.join(' | '))
})

test('auditSlide caps the issue list so a receipt stays readable', () => {
  const elements = Array.from({ length: 20 }, () => ({ type: 'video', box: 'nope', text: 'x' }))
  const issues = auditSlide({ elements })
  assert.ok(issues.length <= 13, `${issues.length} issues`)
  assert.match(issues[issues.length - 1], /另有 \d+ 条参数问题/)
})

test('applyElementPatch merges style and box', () => {
  const deck = normalizeDeck({
    id: 'd1',
    size: { w: 13.333, h: 7.5 },
    slides: [{ id: 's1', elements: [{ id: 'e1', type: 'text', box: [1, 1, 2, 0.5], text: 'old', style: { size: 20, bold: true } }] }],
  }, 'd1')
  const result = applyElementPatch(deck, {
    slideId: 's1',
    elementId: 'e1',
    patch: { text: 'new', box: [2, 2, 3, 0.6], style: { size: 30 } },
  })
  assert.equal(result.ok, true)
  const el = /** @type {any} */ (deck).slides[0].elements[0]
  assert.equal(el.text, 'new')
  assert.deepEqual(el.box, [2, 2, 3, 0.6])
  assert.equal(el.style.size, 30)
  assert.equal(el.style.bold, true, 'style merges instead of replacing')
})

test('applyElementPatch appends when the element is new', () => {
  const deck = normalizeDeck({ id: 'd1', slides: [{ id: 's1', elements: [] }] }, 'd1')
  const result = applyElementPatch(deck, { slideId: 's1', elementId: 'e9', patch: { type: 'rect', box: [1, 1, 1, 1] } })
  assert.equal(result.ok, true)
  assert.equal(/** @type {any} */ (deck).slides[0].elements.length, 1)
})

test('applyElementPatch removes and reports missing targets', () => {
  const deck = normalizeDeck({ id: 'd1', slides: [{ id: 's1', elements: [{ id: 'e1', type: 'text', box: [1, 1, 1, 1] }] }] }, 'd1')
  const removed = applyElementPatch(deck, { slideId: 's1', elementId: 'e1', remove: true, patch: {} })
  assert.equal(removed.removed, true)
  const missing = applyElementPatch(deck, { slideId: 'nope', elementId: 'e1', patch: {} })
  assert.equal(missing.ok, false)
})

test('summarizeDeck reports slide titles and element ids', () => {
  const deck = normalizeDeck({
    id: 'd1',
    slides: [{ id: 's1', elements: [{ id: 'e1', type: 'text', box: [1, 1, 2, 1], text: '标题一' }] }],
  }, 'd1')
  const summary = summarizeDeck(deck)
  assert.equal(summary.slideCount, 1)
  assert.equal(summary.slides[0].title, '标题一')
  assert.equal(summary.slides[0].elements[0].id, 'e1')
})

test('normalizeDeck survives junk input', () => {
  assert.equal(normalizeDeck(null, 'x'), null)
  const deck = normalizeDeck({ slides: 'nope' }, 'fallback')
  assert.equal(deck.id, 'fallback')
  assert.deepEqual(deck.slides, [])
})

test('theme defaults merge with partial overrides', () => {
  const theme = normalizeTheme({ colors: { accent: '22d3ee' }, fonts: { latin: '思源黑体' } })
  assert.equal(theme.colors.accent, '#22D3EE')
  assert.equal(theme.colors.ink, '#0F172A', 'missing colours keep the default')
  assert.equal(theme.fonts.latin, '思源黑体')
  assert.equal(theme.fonts.heading, '微软雅黑')
})

test('transition / animation / shadow normalisation clamps and defaults', () => {
  const transition = normalizeTransition({ type: 'push', speed: 'warp', direction: 'up', advanceAfterMs: -5 })
  assert.equal(transition.type, 'push')
  assert.equal(transition.speed, 'medium')
  assert.equal(transition.direction, 'up')
  assert.equal(transition.advanceAfterMs, 0)
  assert.equal(normalizeTransition({ type: 'nope' }).type, 'fade')
  assert.equal(normalizeTransition(null), null)

  const anim = normalizeAnimation({ effect: 'wipe', direction: 'right', duration: 99999, delay: -2, order: 4 })
  assert.equal(anim.effect, 'wipe')
  assert.equal(anim.direction, 'right')
  assert.equal(anim.duration, 10000)
  assert.equal(anim.delay, 0)
  assert.equal(normalizeAnimation({ effect: 'bogus' }).effect, 'fade')

  const shadow = normalizeShadow(true)
  assert.equal(shadow.alpha, 0.35)
  assert.equal(shadow.color, '#000000')
  assert.equal(normalizeShadow({ blur: 999, alpha: 2 }).blur, 100)
  assert.equal(normalizeShadow({ alpha: 2 }).alpha, 1)
  assert.equal(normalizeShadow(null), null)
})

test('slides carry transitions, backgrounds, notes, hidden and native charts', () => {
  const slide = normalizeSlide({
    transition: { type: 'wipe', direction: 'down' },
    background: '#060B18',
    notes: '备注',
    hidden: true,
    elements: [
      {
        type: 'chart',
        box: [1, 1, 6, 4],
        chart: { kind: 'pie', title: '占比', categories: ['A', 'B'], series: [{ name: 's', values: [1, -2], color: '22D3EE' }] },
        anim: { effect: 'zoom', order: 1 },
        shadow: { blur: 5 },
        paragraphs: [{ text: 'p', runs: [{ text: 'r', bold: true, color: 'ffffff' }] }],
      },
    ],
  }, { index: 0, slideW: 13.333, slideH: 7.5 })
  assert.deepEqual(slide.background, { color: '#060B18' })
  assert.equal(slide.transition.type, 'wipe')
  assert.equal(slide.transition.direction, 'down')
  assert.equal(slide.notes, '备注')
  assert.equal(slide.hidden, true)
  const chart = slide.elements[0].chart
  assert.equal(chart.kind, 'pie')
  assert.equal(chart.series[0].color, '#22D3EE')
  assert.equal(chart.series[0].values[1], -2)
  assert.equal(slide.elements[0].anim.effect, 'zoom')
  assert.equal(slide.elements[0].shadow.blur, 5)
  assert.equal(slide.elements[0].paragraphs[0].runs[0].color, '#FFFFFF')
})

test('normalizePreview keeps optional design-canvas dimensions', () => {
  const preview = normalizePreview({ html: '<section>x</section>', width: 960, height: 540 })
  assert.equal(preview.width, 960)
  assert.equal(preview.height, 540)
  const inferred = normalizePreview({ html: '<section>x</section>' })
  assert.equal('width' in inferred, false)
  assert.equal(normalizePreview({ html: '   ' }), null)
})

test('normalizeChart drops a chart without series and defaults the kind', () => {
  assert.equal(normalizeChart({ categories: ['A'], series: [] }), null)
  const chart = normalizeChart({ categories: ['A'], series: [{ name: 'x', values: [1] }] })
  assert.equal(chart.kind, 'column')
  assert.equal(chart.series[0].name, 'x')
})

test('pipeline stage is derived from research/preview/builtAt', () => {
  const empty = normalizePipeline(null)
  assert.equal(empty.stage, 'research')
  const researched = normalizePipeline({ research: { summary: 's', outline: [{ title: 'a' }] } })
  assert.equal(researched.stage, 'preview')
  const previewed = normalizePipeline({ research: { summary: 's' }, preview: { html: '<html></html>' } })
  assert.equal(previewed.stage, 'build')
  const built = normalizePipeline({ builtAt: '2026-01-01T00:00:00.000Z' })
  assert.equal(built.stage, 'build')
  const summary = pipelineSummary(normalizeDeck({ id: 'p', slides: [{ id: 's1', elements: [] }] }, 'p'))
  assert.equal(summary.stage, 'build', 'a deck with slides is grandfathered')
  assert.ok(summary.builtAt)
})

test('summarizeDeck exposes pipeline and per-slide effects', () => {
  const deck = normalizeDeck({
    id: 'd1',
    pipeline: { research: { summary: 's' }, preview: { html: '<section>x</section>' } },
    slides: [{
      id: 's1',
      transition: { type: 'fade' },
      notes: 'notes',
      elements: [{ id: 'e1', type: 'text', box: [1, 1, 2, 1], text: 'x', anim: { effect: 'fade' } }],
    }],
  }, 'd1')
  const summary = summarizeDeck(deck)
  assert.equal(summary.pipeline.stage, 'build')
  assert.equal(summary.pipeline.preview, true)
  assert.equal(summary.slides[0].transition.type, 'fade')
  assert.equal(summary.slides[0].notes, 'notes')
  assert.equal(summary.slides[0].elements[0].anim.effect, 'fade')
})

console.log(`\nmodel: ${passed} checks passed`)
