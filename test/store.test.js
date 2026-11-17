/**
 * Unit tests for the deck store (filesystem only).
 * Run: node test/store.test.js
 */
import assert from 'node:assert/strict'
import { existsSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { DeckStore, STORE_DIR, ensureInside } from '../lib/store.js'

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

const root = mkdtempSync(join(tmpdir(), 'pptx-studio-'))
const store = new DeckStore({ log: () => {} })

test('ensureInside keeps children and rejects escapes', () => {
  const inside = ensureInside(root, join(root, 'a', 'b'))
  assert.ok(inside.endsWith('b'))
  assert.throws(() => ensureInside(root, join(root, '..', 'evil')), /escapes/)
  // A path on another drive is absolute after `relative()`, not `..`-prefixed:
  // it must still be refused (this used to be accepted on Windows).
  if (process.platform === 'win32') {
    assert.throws(() => ensureInside(root, 'C:\\Users\\someone\\file.pptx'), /escapes/)
  }
})

test('open creates a deck directory and persists deck.json', () => {
  const entry = store.open({ cwd: root, sessionId: 'sess-1', title: '测试 deck' })
  assert.ok(entry.model.id.startsWith('deck_'))
  entry.model.slides.push({ id: 's1', elements: [] })
  entry.persist()
  const file = join(root, STORE_DIR, entry.model.id, 'deck.json')
  assert.equal(existsSync(file), true)
  const saved = JSON.parse(readFileSync(file, 'utf8'))
  assert.equal(saved.title, '测试 deck')
  assert.equal(saved.slides.length, 1)
})

test('reopening an existing deck resumes the saved model', () => {
  const first = store.open({ cwd: root, sessionId: 'sess-1', title: '恢复测试' })
  first.model.slides.push({ id: 's1', elements: [{ id: 'e1', type: 'text', box: [1, 1, 1, 1], text: 'keep' }] })
  first.persist()
  const reopenedStore = new DeckStore({ log: () => {} })
  const resumed = reopenedStore.open({ cwd: root, deckId: first.model.id, sessionId: 'sess-1' })
  assert.equal(resumed.model.slides.length, 1)
  assert.equal(resumed.model.slides[0].elements[0].text, 'keep')
})

test('list() filters by session', () => {
  const other = store.open({ cwd: root, sessionId: 'sess-2', title: '别人的' })
  const mine = store.list('sess-1').map((entry) => entry.model.id)
  assert.ok(mine.length >= 1)
  assert.equal(mine.includes(other.model.id), false)
})

test('emit increments the sequence and notifies subscribers', () => {
  const entry = store.open({ cwd: root, sessionId: 'sess-3', title: '事件' })
  const seen = []
  const off = entry.subscribe((event) => seen.push(event))
  store.changed(entry, { type: 'status', phase: 'drawing' })
  store.changed(entry, { type: 'status', phase: 'planned' })
  off()
  store.changed(entry, { type: 'status', phase: 'done' })
  assert.equal(seen.length, 2)
  assert.equal(seen[0].seq, 1)
  assert.equal(seen[1].seq, 2)
  assert.equal(seen[0].type, 'status')
})

test('corrupt deck.json falls back to a fresh model', () => {
  const entry = store.open({ cwd: root, sessionId: 'sess-4', title: '坏文件' })
  entry.persist()
  const file = join(root, STORE_DIR, entry.model.id, 'deck.json')
  writeFileSync(file, '{ this is not json', 'utf8')
  const fresh = new DeckStore({ log: () => {} }).open({ cwd: root, deckId: entry.model.id, sessionId: 'sess-4' })
  assert.deepEqual(fresh.model.slides, [])
})

test('assetsDir creates the directory inside the deck', () => {
  const entry = store.open({ cwd: root, sessionId: 'sess-5', title: '资产' })
  const dir = store.assetsDir(entry)
  assert.equal(existsSync(dir), true)
  assert.ok(dir.includes(entry.model.id))
})

console.log(`\nstore: ${passed} checks passed`)
