/**
 * Template-first workflow: decomposed `.pptx` template → HTML 模版 → slides.
 *
 * The whole point of this module is that the three artefacts (the template
 * spec, the HTML design sheet and the editable slides) are all derived from one
 * model instead of being written three times by hand. A page's copy is decided
 * once — in the preview's `pages` mapping — and the HTML and the slide are both
 * generated from it, so "HTML 和幻灯片不一致" cannot happen by drift.
 *
 * @module lib/template
 */

import { clampText, normalizeColor, num, round } from './model.js'

/** Version of the spec `assets/pptx_template.py` writes. */
export const TEMPLATE_VERSION = 2

/** Page roles the decomposer can emit. */
export const TEMPLATE_ROLES = new Set(['cover', 'toc', 'section', 'content', 'ending'])

/** Roles that read as a divider: they carry one heading, not a paragraph. */
const DIVIDER_ROLES = new Set(['cover', 'section', 'ending'])

/**
 * CSS pixels per inch (matches the canvas: 1in = 96px).
 * @type {number}
 */
const PX_PER_IN = 96

/**
 * Points per inch (the unit pptx font sizes are authored in).
 * @type {number}
 */
const PX_PER_PT = PX_PER_IN / 72

/**
 * Normalise one field from the decomposer.
 * @param {unknown} raw - candidate field.
 * @param {number} index - 0-based position.
 * @param {number} slideW - slide width in inches.
 * @param {number} slideH - slide height in inches.
 * @returns {Record<string, unknown>|null} field or null.
 */
function normalizeField(raw, index, slideW, slideH) {
  if (raw === null || typeof raw !== 'object') return null
  const src = /** @type {Record<string, any>} */ (raw)
  const box = Array.isArray(src.box) && src.box.length === 4
    ? src.box.map((n) => num(n, 0, 0, slideW + slideH))
    : null
  if (!box || box[2] <= 0 || box[3] <= 0) return null
  const style = src.style !== null && typeof src.style === 'object' ? src.style : {}
  const id = typeof src.id === 'string' && src.id.trim() !== '' ? src.id.trim().slice(0, 40) : `field${index + 1}`
  return {
    id,
    slots: Array.isArray(src.slots) ? src.slots.filter((s) => typeof s === 'string').slice(0, 40) : [],
    box: box.map((n) => round(n)),
    sample: clampText(src.sample, 500),
    perChar: src.perChar === true,
    fillAs: src.fillAs === 'line' ? 'line' : 'slots',
    style: {
      size: num(style.size, 18, 6, 400),
      bold: style.bold === true,
      italic: style.italic === true,
      underline: style.underline === true,
      color: normalizeColor(style.color) || '',
      font: clampText(style.font, 80),
      align: ['l', 'ctr', 'r', 'just', 'center', 'left', 'right'].includes(String(style.align)) ? String(style.align) : 'l',
      anchor: ['top', 'middle', 'bottom'].includes(String(style.anchor)) ? String(style.anchor) : 'top',
    },
    fit: src.fit !== null && typeof src.fit === 'object'
      ? {
        lines: num(src.fit.lines, 1, 1, 60),
        cjkPerLine: num(src.fit.cjkPerLine, 1, 1, 200),
        latinPerLine: num(src.fit.latinPerLine, 1, 1, 400),
        maxCjk: num(src.fit.maxCjk, 1, 1, 4000),
        maxLatin: num(src.fit.maxLatin, 1, 1, 12000),
      }
      : { lines: 1, cjkPerLine: 1, latinPerLine: 1, maxCjk: 1, maxLatin: 1 },
  }
}

/**
 * Normalise one template page.
 * @param {unknown} raw - candidate page.
 * @param {number} index - 0-based position.
 * @param {number} slideW - slide width in inches.
 * @param {number} slideH - slide height in inches.
 * @returns {Record<string, unknown>} page.
 */
function normalizePage(raw, index, slideW, slideH) {
  const src = raw !== null && typeof raw === 'object' ? /** @type {Record<string, any>} */ (raw) : {}
  const role = TEMPLATE_ROLES.has(String(src.role)) ? String(src.role) : 'content'
  const fields = Array.isArray(src.fields)
    ? src.fields.map((f, i) => normalizeField(f, i, slideW, slideH)).filter(Boolean)
    : []
  return {
    index: num(src.index, index, 0, 200),
    role,
    family: typeof src.family === 'string' ? src.family.slice(0, 16) : null,
    layoutName: clampText(src.layoutName, 120),
    plate: typeof src.plate === 'string' ? src.plate.slice(0, 300) : null,
    reference: typeof src.reference === 'string' ? src.reference.slice(0, 300) : null,
    fields,
    capacity: src.capacity !== null && typeof src.capacity === 'object'
      ? {
        fields: num(src.capacity.fields, fields.length, 0, 100),
        maxCjk: num(src.capacity.maxCjk, 0, 0, 20000),
        bodyMaxCjk: num(src.capacity.bodyMaxCjk, 0, 0, 20000),
      }
      : { fields: fields.length, maxCjk: 0, bodyMaxCjk: 0 },
    guidance: clampText(src.guidance, 400),
    slots: Array.isArray(src.slots) ? src.slots : [],
  }
}

/**
 * Normalise the whole decomposed-template spec.
 *
 * Forgiving on purpose: the spec is produced by a Python child process, and a
 * missing optional section must not stop a deck from being built. Anything the
 * builder cannot use is dropped rather than trusted.
 * @param {unknown} raw - python output.
 * @param {{w?: number, h?: number}} [fallbackSize] - deck size when the spec has none.
 * @returns {Record<string, unknown>|null} spec, or null when unusable.
 */
export function normalizeTemplateSpec(raw, fallbackSize = {}) {
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) return null
  const src = /** @type {Record<string, any>} */ (raw)
  if (src.ok === false) return null
  // What this function returns is also what the host persists, so a spec read
  // back from disk arrives in the normalized shape (`pages`, no `slides`).
  // Re-shape it into the python contract and normalise again, which keeps one
  // code path (and keeps a hand-edited file from smuggling junk past it).
  if (Array.isArray(src.pages) && src.pages.length > 0) {
    const { pages: _normalizedPages, ...rest } = src
    return normalizeTemplateSpec({
      ...rest,
      ok: true,
      slides: src.pages.map((/** @type {any} */ page) => ({
        index: page.index,
        role: page.role,
        family: page.family ?? undefined,
        layoutName: page.layoutName,
        plate: page.plate,
        reference: page.reference,
        capacity: page.capacity,
        guidance: page.guidance,
        fields: Array.isArray(page.fields)
          ? page.fields.map((/** @type {any} */ field) => ({
            id: field.id,
            slots: field.slots,
            box: field.box,
            sample: field.sample,
            perChar: field.perChar,
            fillAs: field.fillAs,
            style: field.style,
            fit: field.fit,
          }))
          : [],
        slots: [],
      })),
    }, fallbackSize)
  }
  const sizeSrc = src.size !== null && typeof src.size === 'object' ? src.size : {}
  const w = num(sizeSrc.w, num(fallbackSize.w, 13.333, 2, 40), 2, 40)
  const h = num(sizeSrc.h, num(fallbackSize.h, 7.5, 2, 40), 2, 40)
  const slidesRaw = Array.isArray(src.slides) ? src.slides : []
  const pages = slidesRaw.slice(0, 200).map((page, i) => normalizePage(page, i, w, h))
  if (pages.length === 0) return null
  const designSrc = src.design !== null && typeof src.design === 'object' ? src.design : {}
  const palette = Array.isArray(designSrc.palette)
    ? designSrc.palette.slice(0, 10).map((entry) => {
      const one = entry !== null && typeof entry === 'object' ? entry : {}
      const hex = normalizeColor(one.hex)
      return hex ? { hex, share: num(one.share, 0, 0, 1) } : null
    }).filter(Boolean)
    : []
  const fontUsage = Array.isArray(designSrc.fontUsage)
    ? designSrc.fontUsage.slice(0, 10).map((entry) => {
      const one = entry !== null && typeof entry === 'object' ? entry : {}
      return { name: clampText(one.name, 80), count: num(one.count, 0, 0, 10000) }
    }).filter((one) => one.name !== '')
    : []
  const textColors = Array.isArray(designSrc.textColors)
    ? designSrc.textColors.slice(0, 10).map((entry) => {
      const one = entry !== null && typeof entry === 'object' ? entry : {}
      const hex = normalizeColor(one.hex)
      return hex ? { hex, count: num(one.count, 0, 0, 10000) } : null
    }).filter(Boolean)
    : []
  const families = Array.isArray(src.families)
    ? src.families.slice(0, 100).map((entry) => {
      const one = entry !== null && typeof entry === 'object' ? entry : {}
      return {
        id: clampText(one.id, 16) || 'f',
        role: TEMPLATE_ROLES.has(String(one.role)) ? String(one.role) : 'content',
        slides: Array.isArray(one.slides) ? one.slides.map((n) => num(n, 0, 0, 200)).slice(0, 100) : [],
      }
    })
    : []
  const render = src.render !== null && typeof src.render === 'object' ? src.render : {}
  // Asset paths arrive relative to the spec directory (`plates/s0.jpg`); the
  // deck owns them under `template/`, which is what the canvas, the HTML sheet
  // and the exporter resolve against.
  const underTemplate = (value) => {
    const text = typeof value === 'string' ? value.trim() : ''
    if (text === '') return null
    if (text.startsWith('template/') || text.startsWith('./template/')) return text.replace(/^\.\//, '')
    return `template/${text.replace(/^\.\//, '')}`
  }
  for (const page of pages) {
    page.plate = underTemplate(page.plate)
    page.reference = underTemplate(page.reference)
  }
  // Text that inherited its colour from the layout has no explicit fill in
  // OOXML. Falling back to the template's dominant text colour keeps copy off
  // the default-ink path, which reads as "the theme changed" on a warm plate.
  const inherited = textColors[0]?.hex
  if (inherited) {
    for (const page of pages) {
      for (const field of page.fields) {
        if (!field.style.color) field.style.color = inherited
      }
    }
  }
  return {
    version: TEMPLATE_VERSION,
    path: clampText(src.path, 1024),
    size: { w: round(w), h: round(h) },
    slideCount: num(src.slideCount, pages.length, 0, 500),
    decomposedSlides: num(src.decomposedSlides, pages.length, 0, 500),
    engine: clampText(src.engine, 40),
    render: {
      ok: render.ok === true,
      engine: clampText(render.engine, 40),
      error: clampText(render.error, 400),
      textHidden: render.textHidden === true,
      width: num(render.width, 0, 0, 8000),
      height: num(render.height, 0, 0, 8000),
    },
    theme: src.theme !== null && typeof src.theme === 'object' ? src.theme : { colors: {}, fonts: {} },
    design: {
      palette,
      background: normalizeColor(designSrc.background) || '',
      fonts: designSrc.fonts !== null && typeof designSrc.fonts === 'object'
        ? {
          title: clampText(/** @type {any} */ (designSrc.fonts).title, 80),
          body: clampText(/** @type {any} */ (designSrc.fonts).body, 80),
        }
        : { title: '', body: '' },
      fontUsage,
      textColors,
    },
    families,
    pages,
  }
}

/**
 * Compact spec summary for tool receipts and the canvas.
 * @param {Record<string, any>|null|undefined} spec - template spec.
 * @returns {Record<string, unknown>|null} summary.
 */
export function templateSummary(spec) {
  if (!spec || !Array.isArray(spec.pages)) return null
  /** @type {Record<string, number>} */
  const roles = {}
  for (const page of spec.pages) roles[page.role] = (roles[page.role] || 0) + 1
  return {
    pages: spec.pages.length,
    size: spec.size,
    engine: spec.engine,
    textHidden: Boolean(spec.render?.textHidden),
    roles,
    families: Array.isArray(spec.families) ? spec.families.length : 0,
    palette: (spec.design?.palette || []).map((/** @type {any} */ e) => e.hex),
    fonts: spec.design?.fonts || null,
    fields: spec.pages.reduce((/** @type {number} */ sum, /** @type {any} */ page) => sum + page.fields.length, 0),
  }
}

/**
 * The page catalogue the model picks from during requirement decomposition.
 *
 * Two lines per template page is the whole cost of letting a model see the
 * template; it is enough to map "第 3 页讲活动目的" onto "用 t3 这一页".
 * @param {Record<string, any>} spec - template spec.
 * @param {{limit?: number}} [options] - paging for very long templates.
 * @returns {Array<Record<string, unknown>>} page catalogue.
 */
export function templateCatalogue(spec, options = {}) {
  const limit = num(options.limit, 200, 1, 200)
  return (spec.pages || []).slice(0, limit).map((/** @type {any} */ page) => ({
    page: page.index,
    role: page.role,
    family: page.family,
    fields: page.fields.map((/** @type {any} */ field) => ({
      id: field.id,
      sample: field.sample,
      maxCjk: field.fit.maxCjk,
      box: field.box,
      perChar: field.perChar,
    })),
    capacity: page.capacity,
    guidance: page.guidance,
  }))
}

/**
 * Find one template page by index.
 * @param {Record<string, any>} spec - template spec.
 * @param {number} index - page index.
 * @returns {Record<string, any>|null} page.
 */
export function templatePage(spec, index) {
  if (!spec || !Array.isArray(spec.pages)) return null
  return spec.pages.find((page) => page.index === index) || spec.pages[index] || null
}

/**
 * @typedef {{slot?: string, field: string, text: string}} FieldFill
 */

/**
 * Turn one template page plus filled copy into deck elements.
 *
 * The plate (template background with its own text removed) becomes a
 * full-bleed image; each field becomes a text element at the *template's* own
 * box and style. Because the box comes from the template rather than from the
 * caller, a slide can never drift away from the page it was built on.
 * @param {Record<string, any>} spec - template spec.
 * @param {number} pageIndex - template page index.
 * @param {Record<string, unknown>} fill - `{fieldId: text}` map.
 * @param {{slideId?: string, extra?: any[]}} [options] - id and extra elements.
 * @returns {{slide: Record<string, unknown>, issues: string[], matched: string[]}} built slide.
 */
export function buildSlideFromTemplate(spec, pageIndex, fill, options = {}) {
  const page = templatePage(spec, pageIndex)
  if (!page) {
    return { slide: { elements: [] }, issues: [`模板没有第 ${pageIndex} 页`], matched: [] }
  }
  const slideId = typeof options.slideId === 'string' && options.slideId !== '' ? options.slideId : `s${(page.index || 0) + 1}`
  /** @type {Record<string, unknown>[]} */
  const elements = []
  const issues = []
  /** @type {string[]} */
  const matched = []
  if (page.plate) {
    elements.push({
      id: `${slideId}-e1`,
      type: 'image',
      box: [0, 0, spec.size.w, spec.size.h],
      src: page.plate,
      name: '模板底图',
      field: '__plate',
    })
  } else if (spec.render?.ok === true) {
    issues.push('模板这一页没有可用的底图（分解时渲染失败），导出会只有文字')
  }
  // `render.ok === false` means the template was decomposed structurally only
  // (no PowerPoint). That is a degraded but valid mode: pages get the flat
  // template background colour instead of a plate, so it is not an error.
  const byId = new Map(page.fields.map((/** @type {any} */ field) => [field.id, field]))
  let seq = 2
  for (const [rawField, rawText] of Object.entries(fill || {})) {
    const field = byId.get(rawField)
    if (!field) {
      issues.push(`模板第 ${page.index} 页没有字段 "${rawField}"；可用：${[...byId.keys()].join('/')}`)
      continue
    }
    const text = typeof rawText === 'string' ? rawText : String(rawText ?? '')
    const trimmed = text.trim()
    if (trimmed === '') continue
    matched.push(rawField)
    const limit = field.fit.maxCjk
    if (trimmed.length > limit) {
      issues.push(`字段 ${rawField} 写了 ${trimmed.length} 字，模板这一格约能放 ${limit} 字，可能溢出（已照写，请检查）`)
    }
    elements.push(textElementForField(field, trimmed, `${slideId}-e${seq++}`))
  }
  const extra = Array.isArray(options.extra) ? options.extra : []
  for (const one of extra.slice(0, 40)) {
    const el = one !== null && typeof one === 'object' ? { .../** @type {any} */ (one) } : null
    if (!el) continue
    if (typeof el.id !== 'string' || el.id === '') el.id = `${slideId}-e${seq++}`
    elements.push(el)
  }
  const missing = page.fields
    .filter((/** @type {any} */ field) => !matched.includes(field.id) && field.sample.trim() !== '')
    .map((/** @type {any} */ field) => field.id)
  if (missing.length > 0) {
    issues.push(`这些字段没填，会留空：${missing.join('/')}`)
  }
  return {
    slide: { id: slideId, templateSlide: page.index, elements },
    issues,
    matched,
  }
}

/**
 * One field → one text element. Per-character fields collapse into a single
 * text box over the union of their boxes; that is what keeps a re-typed title
 * readable instead of scattering characters into old kern positions.
 * @param {Record<string, any>} field - template field.
 * @param {string} text - copy.
 * @param {string} id - element id.
 * @returns {Record<string, unknown>} element.
 */
function textElementForField(field, text, id) {
  const style = field.style
  /** @type {Record<string, unknown>} */
  const element = {
    id,
    type: 'text',
    box: field.perChar ? field.box : field.box,
    text,
    field: field.id,
  }
  /** @type {Record<string, unknown>} */
  const elStyle = {
    size: style.size,
    bold: style.bold,
    color: style.color || undefined,
    align: style.align === 'ctr' ? 'center' : style.align === 'r' ? 'right' : style.align === 'just' ? 'justify' : 'left',
  }
  if (style.font) elStyle.font = style.font
  if (style.anchor === 'middle') elStyle.anchor = 'middle'
  if (style.anchor === 'bottom') elStyle.anchor = 'bottom'
  Object.keys(elStyle).forEach((key) => elStyle[key] === undefined && delete elStyle[key])
  element.style = elStyle
  return element
}

/**
 * Coerce a caller-provided field map into `{fieldId: string}`.
 * @param {unknown} raw - candidate map.
 * @returns {Record<string, string>} clean map.
 */
export function normalizeFill(raw) {
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) return {}
  /** @type {Record<string, string>} */
  const out = {}
  for (const [key, value] of Object.entries(/** @type {Record<string, unknown>} */ (raw))) {
    if (typeof key !== 'string' || key.trim() === '') continue
    if (value === null || value === undefined) continue
    if (typeof value !== 'string' && typeof value !== 'number' && typeof value !== 'boolean') continue
    out[key.trim().slice(0, 40)] = String(value).slice(0, 4000)
  }
  return out
}

/**
 * Rebuild `{page, fields}` for each slide, from the deck itself.
 *
 * This is the consistency anchor: the HTML sheet is generated from this, and
 * the slide's own text elements should reproduce the same map. When they do not,
 * something edited the slide without going through the template.
 * @param {Record<string, any>} deck - deck model.
 * @returns {Array<{page: number, fields: Record<string, string>, slideId: string}>} mapping.
 */
export function pagesFromDeck(deck) {
  const slides = Array.isArray(deck?.slides) ? deck.slides : []
  /** @type {Array<{page: number, fields: Record<string, string>, slideId: string}>} */
  const out = []
  for (const slide of slides) {
    const page = Number.isInteger(slide.templateSlide) ? slide.templateSlide : null
    /** @type {Record<string, string>} */
    const fields = {}
    for (const element of Array.isArray(slide.elements) ? slide.elements : []) {
      if (typeof element?.field === 'string' && element.field !== '__plate' && typeof element.text === 'string') {
        fields[element.field] = element.text
      }
    }
    out.push({ slideId: String(slide.id ?? ''), page, fields })
  }
  return out
}

/**
 * Compare the HTML sheet with the slides it is supposed to describe.
 *
 * The failure this exists for is the quiet one: someone patches a slide (or the
 * model writes HTML by hand), the PPTX and the design sheet drift apart, and
 * nobody notices until the file is opened in PowerPoint.
 * @param {Record<string, any>} deck - deck model.
 * @param {Record<string, any>|null} [preview] - preview record (`pipeline.preview`).
 * @returns {string[]} human-readable inconsistencies (empty when aligned).
 */
export function previewConsistency(deck, preview = deck?.pipeline?.preview) {
  const issues = []
  if (!preview) return issues
  const slides = Array.isArray(deck?.slides) ? deck.slides : []
  const pages = Array.isArray(preview.pages) ? preview.pages : []
  if (preview.mode === 'template') {
    if (pages.length !== slides.length) {
      issues.push(`HTML 模版排了 ${pages.length} 页，幻灯片有 ${slides.length} 页——先让两者页数一致再导出`)
    }
    const bySlide = pagesFromDeck(deck)
    for (const [i, page] of pages.entries()) {
      const slide = slides[i]
      if (!slide) continue
      if (page.templateSlide !== slide.templateSlide) {
        issues.push(`第 ${i + 1} 页底版不一致：HTML 用模板第 ${page.templateSlide} 页，幻灯片用第 ${slide.templateSlide} 页`)
      }
      const fields = normalizeFill(page.fields)
      for (const [fieldId, text] of Object.entries(fields)) {
        const actual = bySlide[i]?.fields?.[fieldId]
        if (actual === undefined) {
          issues.push(`第 ${i + 1} 页：HTML 写了字段 ${fieldId}「${shorten(text)}」，幻灯片里没有这个字段`)
        } else if (actual.trim() !== text.trim()) {
          issues.push(`第 ${i + 1} 页字段 ${fieldId} 文字不一致：HTML「${shorten(text)}」/ 幻灯片「${shorten(actual)}」`)
        }
      }
      for (const fieldId of Object.keys(bySlide[i]?.fields || {})) {
        if (!(fieldId in fields)) {
          issues.push(`第 ${i + 1} 页：幻灯片有字段 ${fieldId}，HTML 里没有——先更新 preview 再导出`)
        }
      }
    }
  }
  return issues
}

/**
 * Shorten a string for an issue message.
 * @param {string} text - source text.
 * @param {number} [max] - max length.
 * @returns {string} shortened text.
 */
function shorten(text, max = 16) {
  const value = String(text ?? '')
  return value.length > max ? `${value.slice(0, max)}…` : value
}

/**
 * Escape text for HTML text nodes and attributes.
 * @param {unknown} value - raw value.
 * @returns {string} escaped string.
 */
export function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
}

/**
 * CSS `align-items`/`text-align` for a field style.
 * @param {Record<string, any>} style - field style.
 * @returns {{items: string, text: string}} flexbox and text alignment.
 */
function alignmentOf(style) {
  const items = style.anchor === 'middle' ? 'center' : style.anchor === 'bottom' ? 'flex-end' : 'flex-start'
  const text = style.align === 'ctr' ? 'center' : style.align === 'r' ? 'right' : style.align === 'just' ? 'justify' : 'left'
  return { items, text }
}

/**
 * Render the HTML 模版: one `.slide` per mapped template page, the plate as the
 * background and the copy laid out at the template's own geometry.
 *
 * Images are inlined as data URLs because the preview iframe runs under a
 * `sandbox` CSP that allows `img-src data:` only.
 * @param {Record<string, any>} spec - template spec.
 * @param {Array<{templateSlide: number, fields: Record<string, string>, title?: string}>} pages - page mapping.
 * @param {{designWidth?: number, readImage?: (page: Record<string, any>) => string|null, title?: string}} [options] - render options.
 * @returns {{html: string, width: number, height: number, missing: number[]}} rendered sheet.
 */
export function renderTemplateHtml(spec, pages, options = {}) {
  const designWidth = num(options.designWidth, 1280, 320, 3840)
  const scaleX = designWidth / spec.size.w
  const height = Math.round(spec.size.h * scaleX)
  const { w } = spec.size
  const readImage = typeof options.readImage === 'function' ? options.readImage : () => null
  /** @type {string[]} */
  const blocks = []
  /** @type {number[]} */
  const missing = []
  for (const [i, page] of pages.entries()) {
    const tpl = templatePage(spec, page.templateSlide)
    if (!tpl) {
      missing.push(page.templateSlide)
      blocks.push(`<section class="slide" data-page="${i}" data-template="${page.templateSlide}"><div class="empty">模板没有第 ${escapeHtml(page.templateSlide)} 页</div></section>`)
      continue
    }
    const image = tpl.plate ? readImage(tpl) : null
    const background = image
      ? `background-image:url('${image}')`
      : `background:${escapeHtml(spec.design.background || '#FFFFFF')}`
    const fieldHtml = tpl.fields.map((/** @type {any} */ field) => {
      const text = page.fields?.[field.id]
      if (typeof text !== 'string' || text.trim() === '') return ''
      const style = field.style
      const align = alignmentOf(style)
      // Whole pixels: sub-pixel offsets make the sheet blurry and the numbers
      // unreadable when someone diffs the HTML by hand.
      const left = Math.round(field.box[0] * scaleX)
      const top = Math.round(field.box[1] * scaleX)
      const width = Math.round(field.box[2] * scaleX)
      const boxHeight = Math.round(field.box[3] * scaleX)
      const fontSize = Math.round(style.size * PX_PER_PT * (scaleX / PX_PER_IN))
      const color = style.color || '#1F2937'
      const fontWeight = style.bold ? 700 : 400
      const fontStyle = style.italic ? 'italic' : 'normal'
      const fontFamily = style.font ? `${style.font}, "Microsoft YaHei", sans-serif` : '"Microsoft YaHei", sans-serif'
      return `<div class="slot" data-field="${escapeHtml(field.id)}" style="left:${left}px;top:${top}px;width:${width}px;height:${boxHeight}px;` +
        `font-size:${fontSize}px;color:${escapeHtml(color)};font-weight:${fontWeight};font-style:${fontStyle};` +
        `font-family:${escapeHtml(fontFamily)};text-align:${align.text};align-items:${align.items}">` +
        `<span>${escapeHtml(text).replace(/\n/g, '<br>')}</span></div>`
    }).filter(Boolean).join('')
    blocks.push(`<section class="slide" data-page="${i}" data-template="${page.templateSlide}" data-role="${escapeHtml(tpl.role)}" style="${background};background-size:100% 100%">${fieldHtml}</section>`)
  }
  const html = `<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8">
<title>${escapeHtml(options.title || 'HTML 模版')}</title>
<style>
  html,body{margin:0;padding:0;background:#0F172A;}
  body{display:flex;flex-direction:column;align-items:center;gap:18px;padding:18px;}
  .slide{position:relative;width:${designWidth}px;height:${height}px;overflow:hidden;flex:none;background-color:#fff;box-shadow:0 8px 28px rgba(0,0,0,.35);}
  .slot{position:absolute;display:flex;flex-direction:column;box-sizing:border-box;line-height:1.22;white-space:pre-wrap;word-break:break-word;overflow:hidden;}
  .slot span{display:block;width:100%;}
  .empty{display:flex;align-items:center;justify-content:center;height:100%;font:16px/1.4 "Microsoft YaHei",sans-serif;color:#94A3B8;}
</style></head><body>
${blocks.join('\n')}
</body></html>`
  return { html, width: designWidth, height, missing }
}

/**
 * Requirement decomposition: validate that the outline actually maps onto the
 * template it was written for.
 *
 * The model owns the words; this function only checks the shape of the mapping
 * (unknown page, divider page carrying paragraphs, copy longer than the slot,
 * duplicated template page for two consecutive pages) so a mismatch is caught
 * before anything is built.
 * @param {Record<string, any>} spec - template spec.
 * @param {Array<Record<string, any>>} outline - `research.outline` entries with `templateSlide` + `fields`.
 * @returns {{issues: string[], warnings: string[], pages: Array<{templateSlide: number, fields: Record<string,string>, title: string}>}} result.
 */
export function checkMapping(spec, outline) {
  /** @type {string[]} */
  const issues = []
  /** @type {string[]} */
  const warnings = []
  /** @type {Array<{templateSlide: number, fields: Record<string,string>, title: string}>} */
  const pages = []
  const items = Array.isArray(outline) ? outline : []
  if (items.length === 0) return { issues: ['outline 为空：先给出每一页要讲什么，以及用模板的哪一页做底版'], warnings, pages }
  for (const [index, item] of items.entries()) {
    const label = `outline[${index}]`
    const pageIndex = item.templateSlide
    if (!Number.isInteger(pageIndex)) {
      issues.push(`${label} 缺 templateSlide：这一页打算用模板的哪一页做底版？（用 pptx_studio_template 返回的 page 编号）`)
      continue
    }
    const page = templatePage(spec, pageIndex)
    if (!page) {
      issues.push(`${label} templateSlide=${pageIndex} 超出模板范围（模板共 ${spec.pages.length} 页，编号 0-${spec.pages.length - 1}）`)
      continue
    }
    const fields = normalizeFill(item.fields)
    const known = new Set(page.fields.map((/** @type {any} */ field) => field.id))
    for (const fieldId of Object.keys(fields)) {
      if (!known.has(fieldId)) {
        issues.push(`${label} 填了字段 ${fieldId}，但模板第 ${pageIndex} 页没有这个字段；可用：${[...known].join('/')}`)
      }
    }
    for (const [fieldId, text] of Object.entries(fields)) {
      const field = page.fields.find((/** @type {any} */ one) => one.id === fieldId)
      if (!field) continue
      if (text.length > field.fit.maxCjk * 1.35) {
        issues.push(`${label}.${fieldId} 有 ${text.length} 字，模板该格约 ${field.fit.maxCjk} 字——会溢出，先压短`)
      } else if (text.length > field.fit.maxCjk) {
        warnings.push(`${label}.${fieldId} 有 ${text.length} 字，略超该格 ${field.fit.maxCjk} 字，导出时可能换行`)
      }
    }
    if (DIVIDER_ROLES.has(page.role) && Object.keys(fields).length > 3) {
      warnings.push(`${label} 用了 ${page.role} 页（模板第 ${pageIndex} 页），这类页一般只放标题，塞了 ${Object.keys(fields).length} 个字段`)
    }
    pages.push({ templateSlide: pageIndex, fields, title: clampText(item.title, 200) })
  }
  // The same template page twice in a row reads as a copy/paste slip.
  for (let i = 1; i < pages.length; i += 1) {
    if (pages[i].templateSlide === pages[i - 1].templateSlide) {
      warnings.push(`第 ${i} / ${i + 1} 页用了同一个模板页（${pages[i].templateSlide}）`)
    }
  }
  return { issues, warnings, pages }
}
