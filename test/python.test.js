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

await test('a cartesian chart pins its value axis to the canvas scale', async () => {
  const runner = new PythonRunner({ assetsDir: new URL('../assets', import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1'), log: () => {} })
  const python = await runner.interpreter()
  if (python === null) {
    console.log('skip - no python-pptx on this machine (export stays unavailable)')
    return
  }
  const dir = mkdtempSync(join(tmpdir(), 'pptx-studio-axis-'))
  const deck = {
    version: 2,
    id: 'deck_axis',
    title: 'axis',
    size: { w: 13.333, h: 7.5 },
    theme: { colors: { accent: '2563EB' } },
    slides: [{
      id: 's1',
      elements: [
        { id: 'c1', type: 'chart', box: [1, 1, 5, 3], chart: { kind: 'column', categories: ['a', 'b', 'c'], series: [{ name: 's', values: [10, 25, 18] }] } },
        { id: 'c2', type: 'chart', box: [7, 1, 5, 3], chart: { kind: 'line', categories: ['a', 'b', 'c'], series: [{ name: 's', values: [40, 20, 30] }] } },
      ],
    }],
  }
  const deckFile = join(dir, 'deck.json')
  writeFileSync(deckFile, JSON.stringify(deck), 'utf8')
  const out = join(dir, 'deck.pptx')
  const built = await runner.run('pptx_build.py', [deckFile, out])
  assert.equal(built.features.charts, 2)
  const probe = await runProcess(python, ['-c', [
    'import json, sys',
    'from pptx import Presentation',
    'from pptx.enum.chart import XL_CHART_TYPE',
    'p = Presentation(sys.argv[1])',
    'axes = []',
    'for s in p.slides:',
    '    for sh in s.shapes:',
    '        if sh.has_chart and sh.chart.chart_type not in (XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT):',
    '            va = sh.chart.value_axis',
    '            axes.append([va.minimum_scale, va.maximum_scale, va.major_unit])',
    'print(json.dumps({"axes": axes}))',
  ].join('\n'), out], { timeoutMs: 60_000 })
  assert.equal(probe.code, 0, probe.stderr)
  const axes = parseJsonOutput(probe.stdout).axes
  // Same scale client.js#axisScale picks: [10,25,18] -> 0/30/10; [40,20,30] -> 0/40/10.
  assert.deepEqual(axes[0], [0, 30, 10], 'the exported axis must match the canvas scale (0/30/10)')
  assert.deepEqual(axes[1], [0, 40, 10], 'the exported axis must match the canvas scale (0/40/10)')
})

await test('the design-sheet chart style reaches the export (bare bars, gradient, labels)', async () => {
  const runner = new PythonRunner({ assetsDir: new URL('../assets', import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1'), log: () => {} })
  const python = await runner.interpreter()
  if (python === null) {
    console.log('skip - no python-pptx on this machine (export stays unavailable)')
    return
  }
  const dir = mkdtempSync(join(tmpdir(), 'pptx-studio-style-'))
  const deck = {
    version: 2,
    id: 'deck_style',
    title: 'style',
    size: { w: 13.333, h: 7.5 },
    theme: { colors: { accent: '22D3EE', surface: 'F1F5F9' } },
    slides: [{
      id: 's1',
      background: { color: '#0F172A' },
      elements: [
        { id: 'bare', type: 'chart', box: [1, 1, 7, 4], chart: { kind: 'column', categories: ['a', 'b', 'c'], axis: 'none', gradient: ['#22D3EE', '#0E7490'], dataLabels: true, corner: 0.1, series: [{ name: 's', values: [10, 25, 18], color: '#22D3EE' }] } },
      ],
    }],
  }
  const deckFile = join(dir, 'deck.json')
  writeFileSync(deckFile, JSON.stringify(deck), 'utf8')
  const out = join(dir, 'deck.pptx')
  const built = await runner.run('pptx_build.py', [deckFile, out])
  assert.equal(built.features.charts, 1)
  const probe = await runProcess(python, ['-c', [
    'import json, re, sys, zipfile',
    'z = zipfile.ZipFile(sys.argv[1])',
    'name = [n for n in z.namelist() if n.startswith("ppt/charts/chart")][0]',
    'xml = z.read(name).decode("utf-8")',
    '# python-pptx writes a bare <c:delete/> (val defaults to true) for a hidden axis.',
    'deleted = [t for t in re.findall(r"<c:delete[^>]*>", xml) if \'val="0"\' not in t]',
    'print(json.dumps({',
    '  "axesDeleted": len(deleted) >= 2,',
    '  "gridlinesOff": "<c:majorGridlines/>" not in xml,',
    '  "gradient": "<a:gradFill" in xml and "22D3EE" in xml and "0E7490" in xml,',
    '  "dataLabels": "<c:dLbls>" in xml,',
    '  "autoTitleOff": \'<c:autoTitleDeleted val="1"/>\' in xml,',
    '  "roundRect": \'prst="roundRect"\' in xml,',
    '}))',
  ].join('\n'), out], { timeoutMs: 60_000 })
  assert.equal(probe.code, 0, probe.stderr)
  const facts = parseJsonOutput(probe.stdout)
  assert.equal(facts.axesDeleted, true, 'axis:"none" must delete both axes in the export')
  assert.equal(facts.gridlinesOff, true, 'axis:"none" must drop the gridlines')
  assert.equal(facts.gradient, true, 'the gradient stops must land in the chart XML')
  assert.equal(facts.dataLabels, true, 'dataLabels must emit c:dLbls')
  assert.equal(facts.autoTitleOff, true, 'the single-series auto title must be suppressed')
  assert.equal(facts.roundRect, true, 'corner must become a roundRect geometry')
})

await test('contrast picking reads the plate under each field', async () => {
  const assetsPath = new URL('../assets', import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1')
  const runner2 = new PythonRunner({ assetsDir: assetsPath, log: () => {} })
  let python = null
  try {
    python = await runner2.interpreter()
  } catch {
    console.log('skip - no python-pptx on this machine (export stays unavailable)')
    return
  }
  const probeDir = mkdtempSync(join(tmpdir(), 'pptx-contrast-'))
  const script = [
    'import json, sys',
    'from PIL import Image',
    'sys.path.insert(0, ' + JSON.stringify(assetsPath) + ')',
    'import pptx_template',
    'out = sys.argv[1]',
    'import os',
    'os.makedirs(os.path.join(out, "plates"), exist_ok=True)',
    'img = Image.new("RGB", (1600, 900), (250, 245, 235))',
    'img.paste((40, 18, 8), (0, 0, 1600, 300))',
    'img.save(os.path.join(out, "plates", "s0.jpg"))',
    'fields = [',
    '    {"id": "onDark", "box": [0.5, 0.4, 4, 0.8], "style": {}},',
    '    {"id": "onLight", "box": [0.5, 5.0, 6, 1.2], "style": {}},',
    '    {"id": "fixed", "box": [0.5, 5.0, 6, 1.2], "style": {"color": "#123456"}},',
    ']',
    'pptx_template.apply_contrast_colors([{"plate": "plates/s0.jpg", "fields": fields}], out, 13.333, 7.5)',
    'print(json.dumps({f["id"]: f["style"].get("color") for f in fields}))',
  ].join('\n')
  const result = await runProcess(python, ['-c', script, probeDir], { timeoutMs: 60_000 })
  assert.equal(result.code, 0, result.stderr)
  const picked = parseJsonOutput(result.stdout)
  assert.ok(!picked.onDark.startsWith('#5F2311'), 'copy on a dark band must not stay dark')
  assert.ok(['#FFFFFF', '#FFF8EC'].includes(picked.onDark), `dark band picked ${picked.onDark}`)
  assert.ok(['#2B1206', '#4A2410'].includes(picked.onLight), `light band picked ${picked.onLight}`)
  assert.equal(picked.fixed, '#123456', 'an explicit colour is never overridden')
})

console.log(`\npython: ${passed} checks passed`)
