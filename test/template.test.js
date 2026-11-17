/**
 * Unit tests for the template-first workflow (lib/template.js + the model
 * fields it depends on). No Python, no PowerPoint: these are the pure
 * transforms that decide whether copy lands on the template or scrambles it.
 *
 * Run: node test/template.test.js
 */
import assert from 'node:assert/strict'

import {
  buildSlideFromTemplate,
  checkMapping,
  escapeHtml,
  normalizeFill,
  normalizeTemplateSpec,
  pagesFromDeck,
  previewConsistency,
  renderTemplateHtml,
  templateCatalogue,
  templatePage,
  templateSummary,
} from '../lib/template.js'
import { createDeck, normalizeDeck, normalizePipeline, normalizePreview, normalizeSlide, pipelineStage, pipelineSummary } from '../lib/model.js'

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
 * A spec shaped exactly like `assets/pptx_template.py` output: a cover with a
 * per-character title, a body page, and one page whose text inherited its
 * colour from the layout.
 * @returns {Record<string, any>} raw python output.
 */
function rawSpec() {
  return {
    ok: true,
    version: 2,
    path: 'templates/brand.pptx',
    size: { w: 13.333, h: 7.5 },
    slideCount: 3,
    decomposedSlides: 3,
    engine: 'powerpoint',
    render: { ok: true, engine: 'powerpoint', width: 1600, height: 900, textHidden: true },
    theme: { colors: { accent1: '#FF0000' }, fonts: {} },
    design: {
      palette: [{ hex: '#FEC68D', share: 0.4 }, { hex: '#5F2311', share: 0.2 }],
      background: '#FEC68D',
      fonts: { title: '字魂73号-江南手书', body: '微软雅黑' },
      fontUsage: [{ name: '微软雅黑', count: 9 }, { name: '字魂73号-江南手书', count: 4 }],
      textColors: [{ hex: '#5F2311', count: 12 }],
    },
    families: [
      { id: 'f1', role: 'cover', slides: [0] },
      { id: 'f2', role: 'content', slides: [1] },
    ],
    slides: [
      {
        index: 0,
        role: 'cover',
        family: 'f1',
        layoutName: '标题幻灯片',
        plate: 'plates/s0.png',
        reference: 'reference/s0.png',
        capacity: { fields: 3, maxCjk: 30, bodyMaxCjk: 12 },
        guidance: 'cover页：title=敦煌魅力；body=中国风活动策划',
        fields: [
          {
            id: 'title',
            slots: ['title', 'title2', 'title3', 'title4'],
            box: [3, 1.77, 7.32, 2.49],
            sample: '敦煌魅力',
            perChar: true,
            fillAs: 'line',
            style: { size: 120, bold: false, color: '#5F2311', font: '字魂73号-江南手书', align: 'ctr', anchor: 'top' },
            fit: { lines: 1, cjkPerLine: 4, latinPerLine: 7, maxCjk: 4, maxLatin: 7 },
          },
          {
            id: 'body',
            slots: ['body'],
            box: [4.08, 4.32, 5.17, 0.44],
            sample: '中国风活动策划',
            perChar: false,
            fillAs: 'slots',
            style: { size: 20, bold: true, color: '#5F2311', font: '微软雅黑', align: 'l', anchor: 'top' },
            fit: { lines: 1, cjkPerLine: 18, latinPerLine: 26, maxCjk: 18, maxLatin: 26 },
          },
        ],
        slots: [],
      },
      {
        index: 1,
        role: 'content',
        family: 'f2',
        layoutName: 'Blank',
        plate: 'plates/s1.png',
        reference: 'reference/s1.png',
        capacity: { fields: 2, maxCjk: 100, bodyMaxCjk: 90 },
        guidance: 'content页：title=活动目的；body=……',
        fields: [
          {
            id: 'title',
            slots: ['title'],
            box: [3.81, 3.19, 5.71, 1.11],
            sample: '活动目的',
            perChar: false,
            fillAs: 'slots',
            style: { size: 60, bold: true, color: '#5F2311', font: '思源宋体 CN Heavy', align: 'ctr', anchor: 'top' },
            fit: { lines: 1, cjkPerLine: 6, latinPerLine: 11, maxCjk: 6, maxLatin: 11 },
          },
          {
            id: 'body',
            slots: ['body'],
            box: [1, 4.6, 5, 1.4],
            sample: '敦煌是一个美丽的梦。',
            perChar: false,
            fillAs: 'slots',
            style: { size: 16, bold: false, font: '微软雅黑', align: 'l', anchor: 'top' },
            fit: { lines: 3, cjkPerLine: 20, latinPerLine: 37, maxCjk: 60, maxLatin: 110 },
          },
        ],
        slots: [],
      },
      {
        index: 2,
        role: 'ending',
        family: 'f2',
        plate: null,
        reference: null,
        capacity: { fields: 1, maxCjk: 10, bodyMaxCjk: 0 },
        guidance: 'ending页',
        fields: [],
        slots: [],
      },
    ],
  }
}

/** @type {Record<string, any>} */
let spec

await test('normalizeTemplateSpec keeps the model and points assets under template/', () => {
  spec = normalizeTemplateSpec(rawSpec(), { w: 10, h: 7.5 })
  assert.ok(spec)
  assert.deepEqual(spec.size, { w: 13.333, h: 7.5 }, 'the template size wins over the deck size')
  assert.equal(spec.pages.length, 3)
  assert.equal(spec.pages[0].plate, 'template/plates/s0.png', 'asset paths are deck-relative')
  assert.equal(spec.pages[0].reference, 'template/reference/s0.png')
  assert.equal(spec.pages[0].fields.length, 2)
  assert.equal(spec.pages[0].fields[0].perChar, true)
  assert.equal(spec.pages[0].fields[0].fit.maxCjk, 4)
  assert.equal(spec.render.textHidden, true)
  assert.equal(spec.design.palette.length, 2)
})

await test('normalizeTemplateSpec inherits the template text colour when a run had none', () => {
  const page = spec.pages[1]
  const body = page.fields.find((/** @type {any} */ field) => field.id === 'body')
  assert.equal(body.style.color, '#5F2311', 'no explicit fill -> the template dominant text colour')
})

await test('normalizeTemplateSpec rejects junk instead of trusting it', () => {
  assert.equal(normalizeTemplateSpec(null), null)
  assert.equal(normalizeTemplateSpec({ ok: false, error: 'boom' }), null)
  assert.equal(normalizeTemplateSpec({ ok: true, slides: [] }), null, 'no pages means nothing to build from')
  const dirty = normalizeTemplateSpec({
    ok: true,
    size: { w: 'nope', h: 7.5 },
    slides: [{ index: 0, role: 'nonsense', fields: [{ id: '', box: [0, 0, 0, 0] }, { id: 'ok', box: [1, 1, 2, 1], style: { size: 9999 } }] }],
  })
  assert.ok(dirty)
  assert.equal(dirty.size.w, 13.333, 'bad width falls back to the default deck size')
  assert.equal(dirty.pages[0].role, 'content', 'unknown role degrades to content')
  assert.equal(dirty.pages[0].fields.length, 1, 'a zero-size box is dropped')
  assert.equal(dirty.pages[0].fields[0].style.size, 400, 'font size is clamped to the model cap')
})

await test('normalizeTemplateSpec round-trips the spec it persisted', () => {
  // The host writes the *normalized* spec to disk; reading it back must not
  // produce null (that silently disabled template mode after a reload).
  const once = normalizeTemplateSpec(rawSpec())
  const twice = normalizeTemplateSpec(once)
  assert.ok(twice, 'a persisted spec must load again')
  assert.equal(twice.pages.length, 3)
  assert.equal(twice.pages[0].plate, 'template/plates/s0.png', 'asset paths stay deck-relative, not double-prefixed')
  assert.equal(twice.pages[0].fields.length, 2)
  assert.equal(twice.pages[0].fields[0].perChar, true)
  assert.equal(twice.design.palette.length, 2)
  assert.equal(twice.render.textHidden, true)
  // A hand-edited file still goes through the same validation.
  const tampered = normalizeTemplateSpec({ ...once, pages: [{ index: 0, role: 'nonsense', fields: [{ id: 'x', box: [0, 0, 0, 0] }] }] })
  assert.equal(tampered.pages[0].role, 'content')
  assert.equal(tampered.pages[0].fields.length, 0)
})

await test('templateSummary and templateCatalogue are what the model sees', () => {
  const summary = templateSummary(spec)
  assert.equal(summary.pages, 3)
  assert.deepEqual(summary.roles, { cover: 1, content: 1, ending: 1 })
  assert.equal(summary.fields, 4)
  const catalogue = templateCatalogue(spec)
  assert.equal(catalogue.length, 3)
  assert.deepEqual(catalogue[0].fields.map((/** @type {any} */ f) => f.id), ['title', 'body'])
  assert.equal(catalogue[0].fields[0].maxCjk, 4)
  assert.equal(templatePage(spec, 2).role, 'ending')
  assert.equal(templatePage(spec, 99), null)
})

await test('buildSlideFromTemplate takes geometry from the template, never from the caller', () => {
  const built = buildSlideFromTemplate(spec, 1, { title: '活动目的', body: '一句话说清楚目的。' })
  assert.deepEqual(built.issues, [])
  assert.deepEqual(built.matched, ['title', 'body'])
  assert.equal(built.slide.templateSlide, 1)
  const plate = built.slide.elements.find((/** @type {any} */ el) => el.field === '__plate')
  assert.ok(plate, 'a plate element covers the page')
  assert.deepEqual(plate.box, [0, 0, 13.333, 7.5])
  assert.equal(plate.src, 'template/plates/s1.png')
  const title = built.slide.elements.find((/** @type {any} */ el) => el.field === 'title')
  assert.deepEqual(title.box, [3.81, 3.19, 5.71, 1.11], 'box comes from the template field')
  assert.equal(title.style.size, 60)
  assert.equal(title.style.align, 'center')
  assert.equal(title.text, '活动目的')
})

await test('buildSlideFromTemplate collapses a per-character title into one centred line', () => {
  const built = buildSlideFromTemplate(spec, 0, { title: '敦煌夜游', body: '中国风活动策划' })
  const title = built.slide.elements.find((/** @type {any} */ el) => el.field === 'title')
  assert.equal(title.text, '敦煌夜游', 'the four original per-character boxes do not get one char each')
  assert.deepEqual(title.box, [3, 1.77, 7.32, 2.49], 'the merged union box is used')
  assert.equal(title.style.align, 'center')
})

await test('buildSlideFromTemplate names unknown fields, overflow and unpainted slots', () => {
  const built = buildSlideFromTemplate(spec, 1, { title: '活动目的', bogus: 'x', body: '这一句明显太长了，会被判定为超出模板这一格能放下的字数上限。' .repeat(3) })
  assert.ok(built.issues.some((line) => /没有字段 "bogus"/.test(line)))
  assert.ok(built.issues.some((line) => /约能放 60 字/.test(line)))
  assert.ok(!built.slide.elements.some((/** @type {any} */ el) => el.field === 'bogus'), 'an unknown field writes nothing')
})

await test('buildSlideFromTemplate reports the fields it left empty', () => {
  const built = buildSlideFromTemplate(spec, 1, { title: '活动目的' })
  assert.ok(built.issues.some((line) => /没填，会留空：body/.test(line)))
})

await test('buildSlideFromTemplate works without a render (no PowerPoint degradation)', () => {
  const structural = normalizeTemplateSpec({
    ...rawSpec(),
    render: { ok: false, engine: 'none' },
    slides: [{ ...rawSpec().slides[1], plate: null }],
  })
  const built = buildSlideFromTemplate(/** @type {any} */ (structural), 0, { title: '活动目的', body: '纯色底也要能建页。' })
  assert.deepEqual(built.issues, [], 'no plate is a degradation, not a build error')
  assert.ok(!built.slide.elements.some((/** @type {any} */ el) => el.field === '__plate'))
})

await test('buildSlideFromTemplate keeps extra elements and hands out unique ids', () => {
  const built = buildSlideFromTemplate(spec, 1, { title: 'A' }, {
    slideId: 's9',
    extra: [{ type: 'rect', box: [1, 1, 2, 1], fill: '#22D3EE' }],
  })
  const ids = built.slide.elements.map((/** @type {any} */ el) => el.id)
  assert.equal(new Set(ids).size, ids.length, 'ids stay unique')
  assert.ok(ids.every((id) => String(id).startsWith('s9-')))
})

await test('checkMapping refuses a page that is not in the template', () => {
  const checked = checkMapping(spec, [{ title: 'A', templateSlide: 9 }])
  assert.ok(checked.issues.some((line) => /超出模板范围/.test(line)))
})

await test('checkMapping refuses a field the template page does not have', () => {
  const checked = checkMapping(spec, [{ title: 'A', templateSlide: 1, fields: { subtitle: 'x' } }])
  assert.ok(checked.issues.some((line) => /没有这个字段/.test(line)))
  assert.match(checked.issues[0], /可填：title\/body|可用：title\/body/)
})

await test('checkMapping refuses copy that cannot fit and warns on copy that barely does', () => {
  const tooLong = checkMapping(spec, [{ title: 'A', templateSlide: 1, fields: { title: '一二三四五六七八九十' } }])
  assert.ok(tooLong.issues.some((line) => /会溢出/.test(line)))
  const barely = checkMapping(spec, [{ title: 'A', templateSlide: 1, fields: { body: '字'.repeat(70) } }])
  assert.ok(barely.warnings.some((line) => /略超/.test(line)))
})

await test('checkMapping marks the same template page twice as a copy/paste risk', () => {
  const checked = checkMapping(spec, [
    { title: 'A', templateSlide: 1, fields: { title: '一' } },
    { title: 'B', templateSlide: 1, fields: { title: '二' } },
  ])
  assert.ok(checked.warnings.some((line) => /同一个模板页/.test(line)))
})

await test('checkMapping requires a mapping at all', () => {
  assert.equal(checkMapping(spec, []).issues.length, 1)
  assert.equal(checkMapping(spec, [{ title: 'A' }]).issues.length, 1, 'a page without templateSlide cannot be built')
})

await test('renderTemplateHtml draws one .slide per page at the template geometry', () => {
  const sheet = renderTemplateHtml(spec, [
    { templateSlide: 0, fields: { title: '敦煌夜游', body: '中国风活动策划' } },
    { templateSlide: 1, fields: { title: '活动目的', body: '一句话。' } },
  ], {
    designWidth: 1280,
    readImage: () => 'data:image/jpeg;base64,AAAA',
  })
  assert.equal((sheet.html.match(/class="slide"/g) || []).length, 2)
  assert.equal(sheet.width, 1280)
  assert.equal(sheet.height, 720)
  assert.equal((sheet.html.match(/data:image\/jpeg;base64,AAAA/g) || []).length, 2, 'the plate is inlined per page')
  // 3in at 1280/13.333 scale = 288px; 120pt = 160px.
  assert.match(sheet.html, /left:288px;top:170px;width:703px;height:239px/)
  assert.match(sheet.html, /font-size:160px/)
  assert.match(sheet.html, /data-field="title"/)
  assert.match(sheet.html, /敦煌夜游/)
})

await test('renderTemplateHtml escapes copy and reports pages the template does not have', () => {
  const sheet = renderTemplateHtml(spec, [{ templateSlide: 1, fields: { title: '<script>alert(1)</script>' } }], { readImage: () => null })
  assert.ok(!sheet.html.includes('<script>alert(1)'), 'copy is escaped')
  assert.ok(sheet.html.includes('&lt;script&gt;'))
  const missing = renderTemplateHtml(spec, [{ templateSlide: 42, fields: {} }], { readImage: () => null })
  assert.deepEqual(missing.missing, [42])
  assert.match(missing.html, /模板没有第 42 页/)
})

await test('renderTemplateHtml falls back to the template background colour without a plate', () => {
  const sheet = renderTemplateHtml(spec, [{ templateSlide: 1, fields: { title: 'A' } }], { readImage: () => null })
  assert.match(sheet.html, /background:#FEC68D/)
})

await test('escapeHtml covers the five interesting characters', () => {
  assert.equal(escapeHtml(`<&>"'`), '&lt;&amp;&gt;&quot;&#39;')
})

await test('normalizeFill keeps strings/numbers and drops the rest', () => {
  assert.deepEqual(normalizeFill({ a: 'x', b: 3, c: null, d: { x: 1 }, '': 'y' }), { a: 'x', b: '3' })
})

await test('pagesFromDeck rebuilds the mapping from element provenance', () => {
  const deck = createDeck({ id: 'deck_1', width: 13.333, height: 7.5 })
  deck.slides = [
    normalizeSlide({
      id: 's1',
      templateSlide: 0,
      elements: [
        { type: 'image', box: [0, 0, 13, 7], src: 'template/plates/s0.jpg', field: '__plate' },
        { type: 'text', box: [1, 1, 2, 1], text: '标题', field: 'title' },
      ],
    }, { index: 0, slideW: 13.333, slideH: 7.5 }),
  ]
  const pages = pagesFromDeck(deck)
  assert.equal(pages.length, 1)
  assert.equal(pages[0].page, 0)
  assert.deepEqual(pages[0].fields, { title: '标题' }, 'the plate is not copy')
})

await test('previewConsistency catches HTML/slide drift in both directions', () => {
  const deck = createDeck({ id: 'deck_2', width: 13.333, height: 7.5 })
  deck.slides = [
    normalizeSlide({ id: 's1', templateSlide: 0, elements: [{ type: 'text', box: [1, 1, 2, 1], text: '标题', field: 'title' }] }, { index: 0, slideW: 13.333, slideH: 7.5 }),
    normalizeSlide({ id: 's2', templateSlide: 1, elements: [{ type: 'text', box: [1, 1, 2, 1], text: '第二页', field: 'title' }] }, { index: 1, slideW: 13.333, slideH: 7.5 }),
  ]
  const aligned = previewConsistency(deck, {
    mode: 'template',
    pages: [
      { templateSlide: 0, fields: { title: '标题' } },
      { templateSlide: 1, fields: { title: '第二页' } },
    ],
  })
  assert.deepEqual(aligned, [])

  const drifted = previewConsistency(deck, {
    mode: 'template',
    pages: [
      { templateSlide: 0, fields: { title: '标题被改了' } },
      { templateSlide: 2, fields: { title: '第二页', body: 'HTML 多写的' } },
    ],
  })
  assert.ok(drifted.some((line) => /文字不一致/.test(line)))
  assert.ok(drifted.some((line) => /底版不一致/.test(line)))
  assert.ok(drifted.some((line) => /幻灯片里没有这个字段/.test(line)))

  const shortSheet = previewConsistency(deck, { mode: 'template', pages: [{ templateSlide: 0, fields: {} }] })
  assert.ok(shortSheet.some((line) => /排了 1 页，幻灯片有 2 页/.test(line)))
})

await test('pipeline gains a template stage before research when a template is required', () => {
  const pipeline = normalizePipeline({ template: { path: 'a.pptx', required: true } })
  assert.equal(pipeline.stage, 'template')
  assert.equal(pipeline.template.decomposedAt, null)
  const decomposed = normalizePipeline({ template: { path: 'a.pptx', required: true, decomposedAt: '2026-01-01T00:00:00.000Z', pages: 12 } })
  assert.equal(decomposed.stage, 'research')
  assert.equal(decomposed.template.pages, 12)
  assert.equal(pipelineStage(null), 'research')
})

await test('a deck opened with a template reports the template stage and inherits it from disk', () => {
  const deck = createDeck({ id: 'deck_3', template: 'templates/brand.pptx' })
  assert.equal(deck.pipeline.stage, 'template')
  assert.equal(deck.pipeline.template.path, 'templates/brand.pptx')
  assert.equal(deck.pipeline.template.required, true)
  // Decks written before the field existed still come back in template mode…
  const legacy = normalizeDeck({ id: 'deck_4', template: { path: 't.pptx' }, slides: [] }, 'deck_4')
  assert.equal(legacy.pipeline.stage, 'template')
  assert.equal(legacy.pipeline.template.required, true)
  // …but a legacy deck that already has slides is grandfathered as built.
  const built = normalizeDeck({ id: 'deck_5', template: { path: 't.pptx' }, slides: [{ id: 's1', elements: [{ type: 'text', box: [1, 1, 2, 1], text: 'x' }] }] }, 'deck_5')
  assert.equal(built.pipeline.stage, 'build')
  assert.equal(built.slides[0].templateSlide, undefined)
})

await test('pipelineSummary exposes template state and the template preview mode', () => {
  const deck = createDeck({ id: 'deck_6', template: 't.pptx' })
  deck.pipeline.template.decomposedAt = '2026-01-01T00:00:00.000Z'
  deck.pipeline.template.pages = 5
  deck.pipeline.template.engine = 'powerpoint'
  deck.pipeline.preview = { mode: 'template', html: '', pages: [{ templateSlide: 0, fields: {} }], title: 'x', notes: '', updatedAt: 'x' }
  const summary = pipelineSummary(deck)
  assert.equal(summary.stage, 'build')
  assert.equal(summary.template.decomposed, true)
  assert.equal(summary.template.pages, 5)
  assert.equal(summary.previewMode, 'template')
  assert.equal(summary.previewPages, 1)
})

await test('normalizePreview accepts a template mapping with no hand-written HTML', () => {
  const preview = normalizePreview({ mode: 'template', pages: [{ templateSlide: 2, fields: { title: 'x' } }] })
  assert.ok(preview)
  assert.equal(preview.mode, 'template')
  assert.equal(preview.pages.length, 1)
  assert.equal(normalizePreview({ mode: 'template', pages: [] }), null, 'a template preview needs pages')
  assert.equal(normalizePreview({ html: '' }), null, 'the HTML mode still needs HTML')
})

console.log(`\ntemplate: ${passed} checks passed`)
