/**
 * Unit tests for the python bridge.
 * Run: node test/python.test.js
 *
 * The python-pptx part skips itself when no interpreter with python-pptx is
 * reachable, so the test suite still passes on a machine without Python.
 */
import assert from 'node:assert/strict'
import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { PythonRunner, parseJsonOutput, pythonCandidates, runProcess } from '../lib/python.js'

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

await test('parseJsonOutput ignores noise and keeps the last JSON line', () => {
  const parsed = parseJsonOutput('warning: x\n{"ok":true,"a":1}\ntrailing text\n')
  assert.deepEqual(parsed, { ok: true, a: 1 })
  assert.throws(() => parseJsonOutput('no json here'), /JSON/)
})

await test('pythonCandidates puts an explicit interpreter first and dedupes', () => {
  const list = pythonCandidates('C:/custom/python.exe')
  assert.equal(list[0], 'C:/custom/python.exe')
  assert.equal(new Set(list).size, list.length)
})

await test('runProcess captures stdout, stderr and exit code', async () => {
  const result = await runProcess(process.execPath, ['-e', 'console.log("out"); console.error("err"); process.exit(3)'])
  assert.equal(result.code, 3)
  assert.match(result.stdout, /out/)
  assert.match(result.stderr, /err/)
})

await test('PythonRunner resolves a python with python-pptx (or degrades)', async () => {
  const runner = new PythonRunner({ assetsDir: new URL('../assets', import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1'), log: () => {} })
  const python = await runner.interpreter()
  if (python === null) {
    console.log('skip - no python-pptx on this machine (export stays unavailable)')
    return
  }
  const dir = mkdtempSync(join(tmpdir(), 'pptx-studio-py-'))
  const deck = {
    version: 1,
    id: 'deck_py',
    title: 'python bridge',
    size: { w: 13.333, h: 7.5 },
    slides: [{ id: 's1', elements: [{ id: 'e1', type: 'text', box: [1, 1, 4, 1], text: 'hello' }] }],
  }
  const deckFile = join(dir, 'deck.json')
  writeFileSync(deckFile, JSON.stringify(deck), 'utf8')
  const out = join(dir, 'deck.pptx')
  const receipt = await runner.run('pptx_build.py', [deckFile, out])
  assert.equal(receipt.ok, true)
  assert.equal(receipt.slides, 1)
  assert.ok(receipt.bytes > 1000)
  const info = await runner.run('pptx_inspect.py', [out])
  assert.equal(info.ok, true)
  assert.equal(info.slideCount, 1)
})

console.log(`\npython: ${passed} checks passed`)
