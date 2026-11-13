/**
 * Unit tests for the deck model (no DSH runtime needed).
 * Run: node test/model.test.js
 */
import assert from 'node:assert/strict'

import {
  applyElementPatch,
  createDeck,
  normalizeDeck,
  normalizeSlide,
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

console.log(`\nmodel: ${passed} checks passed`)
