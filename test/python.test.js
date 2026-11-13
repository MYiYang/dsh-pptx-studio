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

  // The writer must produce real OOXML for the PPTX features, not just a
  // plausible model. Read the zip back with python's own zipfile.
  const rich = {
    version: 2,
    id: 'deck_rich',
    title: 'effects',
    size: { w: 13.333, h: 7.5 },
    theme: { colors: { accent: '22D3EE' }, fonts: { latin: '微软雅黑' } },
    slides: [{
      id: 's1',
      notes: '特效测试备注',
      hidden: true,
      background: { color: '#060B18' },
      transition: { type: 'push', speed: 'slow', direction: 'up', advanceOnClick: false, advanceAfterMs: 2000 },
      elements: [
        { id: 't', type: 'text', box: [1, 1, 4, 1], text: 'hello', shadow: true, anim: { effect: 'fade', duration: 600 } },
        { id: 'c', type: 'chart', box: [6, 2, 5, 3], chart: { kind: 'pie', categories: ['A', 'B'], series: [{ name: 's', values: [1, 2] }] } },
      ],
    }],
  }
  const richFile = join(dir, 'rich.json')
  writeFileSync(richFile, JSON.stringify(rich), 'utf8')
  const richOut = join(dir, 'rich.pptx')
  const built = await runner.run('pptx_build.py', [richFile, richOut])
  assert.equal(built.features.transitions, 1)
  assert.equal(built.features.timings, 1)
  assert.equal(built.features.animations, 1)
  assert.equal(built.features.notes, 1)
  assert.equal(built.features.hidden, 1)
  assert.equal(built.features.charts, 1)

  const probeScript = [
    'import json, sys, zipfile',
    'import xml.etree.ElementTree as ET',
    'p = sys.argv[1]',
    'z = zipfile.ZipFile(p)',
    'slide = z.read("ppt/slides/slide1.xml").decode("utf-8")',
    'notes = ""',
    'if "ppt/notesSlides/notesSlide1.xml" in z.namelist():',
    '    notes = z.read("ppt/notesSlides/notesSlide1.xml").decode("utf-8")',
    'root = ET.fromstring(slide)',
    'children = [c.tag.split("}")[-1] for c in root]',
    'print(json.dumps({',
    '  "children": children,',
    '  "transition": bool([c for c in children if c == "transition"]),',
    '  "timing": bool([c for c in children if c == "timing"]),',
    '  "hidden": root.get("show") == "0",',
    '  "shadow": "outerShdw" in slide,',
    '  "chart": any(n.startswith("ppt/charts/chart") for n in z.namelist()),',
    '  "notes": "特效测试备注" in notes,',
    '}, ensure_ascii=False))',
  ].join('\n')
  const probe = await runProcess(await runner.interpreter(), ['-c', probeScript, richOut], { timeoutMs: 60_000 })
  assert.equal(probe.code, 0, probe.stderr)
  const facts = parseJsonOutput(probe.stdout)
  assert.deepEqual(facts.children, ['cSld', 'clrMapOvr', 'transition', 'timing'], 'OOXML child order must be transition then timing')
  assert.equal(facts.transition, true)
  assert.equal(facts.timing, true)
  assert.equal(facts.hidden, true)
  assert.equal(facts.shadow, true)
  assert.equal(facts.chart, true)
  assert.equal(facts.notes, true)

  // `--no-effects` must be able to strip transitions/timing for compatibility.
  const plainOut = join(dir, 'plain.pptx')
  const plain = await runner.run('pptx_build.py', [richFile, plainOut, '--no-effects'])
  assert.equal(plain.features.transitions, 0)
  assert.equal(plain.features.timings, 0)
})

console.log(`\npython: ${passed} checks passed`)
