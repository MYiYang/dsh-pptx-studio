/**
 * Deck model: validation, normalisation and patch application.
 *
 * The model is the single source of truth shared by three consumers:
 *   1. the agent tools in `index.js`,
 *   2. the right-side canvas (`client.js`),
 *   3. `assets/pptx_build.py`, which turns it into a real .pptx.
 *
 * Everything is expressed in inches so the canvas can scale by `px / inch`
 * without a second unit system. See REQUIREMENTS.md for the schema.
 * @module dsh-pptx-studio/lib/model
 */

export const DECK_VERSION = 1
export const MAX_SLIDES = 60
export const MAX_ELEMENTS_PER_SLIDE = 200
export const MAX_TEXT = 8000
export const MAX_ELEMENTS_TOTAL = 2000
export const ELEMENT_TYPES = new Set(['text', 'rect', 'image', 'line', 'table', 'chart', 'group'])

/**
 * Clamp one number into a safe range.
 * @param {unknown} value - candidate.
 * @param {number} fallback - used when the candidate is not a finite number.
 * @param {number} min - lower bound.
 * @param {number} max - upper bound.
 * @returns {number} the clamped number.
 */
export function num(value, fallback, min, max) {
  const n = typeof value === 'number' ? value : Number(value)
  if (!Number.isFinite(n)) return fallback
  return Math.min(max, Math.max(min, n))
}

/**
 * Normalise a box `[x, y, w, h]` (inches) into a fresh array.
 * @param {unknown} box - candidate box.
 * @param {number} slideW - slide width in inches.
 * @param {number} slideH - slide height in inches.
 * @returns {number[]} `[x, y, w, h]`.
 */
export function normalizeBox(box, slideW, slideH) {
  const raw = Array.isArray(box) ? box : []
  const x = num(raw[0], 0, -2, slideW + 2)
  const y = num(raw[1], 0, -2, slideH + 2)
  const w = num(raw[2], Math.max(0.2, slideW - 2 * Math.max(0, x)), 0.01, slideW * 2)
  const h = num(raw[3], Math.max(0.2, slideH - 2 * Math.max(0, y)), 0.01, slideH * 2)
  return [round(x), round(y), round(w), round(h)]
}

/**
 * Round to 3 decimals (0.001in ~ 0.025mm) to keep JSON and diffs small.
 * @param {number} n - number.
 * @returns {number} rounded number.
 */
export function round(n) {
  return Math.round(n * 1000) / 1000
}

/**
 * Clamp one string, including the length cap.
 * @param {unknown} value - candidate.
 * @param {number} max - maximum length.
 * @returns {string} string value.
 */
export function clampText(value, max = MAX_TEXT) {
  const s = typeof value === 'string' ? value : value == null ? '' : String(value)
  return s.length > max ? s.slice(0, max) : s
}

/**
 * Normalise a style object to the keys the canvas and python builder know.
 * @param {unknown} style - candidate.
 * @returns {Record<string, unknown>} normalised style (empty object when absent).
 */
export function normalizeStyle(style) {
  if (style === null || typeof style !== 'object') return {}
  const src = /** @type {Record<string, unknown>} */ (style)
  /** @type {Record<string, unknown>} */
  const out = {}
  if (src.size !== undefined) out.size = num(src.size, 18, 5, 200)
  if (src.bold !== undefined) out.bold = src.bold === true
  if (src.italic !== undefined) out.italic = src.italic === true
  if (typeof src.color === 'string') out.color = src.color.slice(0, 32)
  if (typeof src.font === 'string') out.font = src.font.slice(0, 64)
  if (typeof src.align === 'string' && ['left', 'center', 'right', 'justify'].includes(src.align)) out.align = src.align
  if (typeof src.anchor === 'string' && ['top', 'middle', 'bottom'].includes(src.anchor)) out.anchor = src.anchor
  if (src.lineSpacing !== undefined) out.lineSpacing = num(src.lineSpacing, 1.15, 0.6, 3)
  if (src.spaceAfter !== undefined) out.spaceAfter = num(src.spaceAfter, 0, 0, 200)
  return out
}

/**
 * Normalise one element.
 * @param {unknown} raw - candidate element.
 * @param {{index: number, slideW: number, slideH: number, prefix: string}} ctx - id/size context.
 * @returns {Record<string, unknown>} normalised element.
 */
export function normalizeElement(raw, ctx) {
  const src = raw !== null && typeof raw === 'object' ? /** @type {Record<string, unknown>} */ (raw) : {}
  const type = ELEMENT_TYPES.has(String(src.type)) ? String(src.type) : 'text'
  /** @type {Record<string, unknown>} */
  const el = {
    id: typeof src.id === 'string' && src.id.trim() !== '' ? src.id.slice(0, 64) : `${ctx.prefix}e${ctx.index + 1}`,
    type,
    box: normalizeBox(src.box, ctx.slideW, ctx.slideH),
  }
  if (src.text !== undefined) el.text = clampText(src.text)
  if (src.paragraphs !== undefined && Array.isArray(src.paragraphs)) {
    el.paragraphs = src.paragraphs.slice(0, 40).map((p) => {
      const para = p !== null && typeof p === 'object' ? /** @type {Record<string, unknown>} */ (p) : {}
      /** @type {Record<string, unknown>} */
      const out = { text: clampText(para.text, 2000) }
      if (para.size !== undefined) out.size = num(para.size, 18, 5, 200)
      if (para.bold !== undefined) out.bold = para.bold === true
      if (typeof para.color === 'string') out.color = para.color.slice(0, 32)
      if (typeof para.align === 'string') out.align = para.align
      return out
    })
  }
  if (src.style !== undefined) el.style = normalizeStyle(src.style)
  for (const key of ['fill', 'line', 'src', 'radius', 'lineWidth', 'name']) {
    if (src[key] !== undefined) {
      if (key === 'radius' || key === 'lineWidth') el[key] = num(src[key], 1, 0, 20)
      else el[key] = typeof src[key] === 'string' ? String(src[key]).slice(0, 512) : src[key]
    }
  }
  if (src.table !== undefined && src.table !== null && typeof src.table === 'object') {
    const rows = Array.isArray(/** @type {any} */ (src.table).rows) ? /** @type {any} */ (src.table).rows : []
    el.table = {
      rows: rows.slice(0, 30).map((row) =>
        (Array.isArray(row) ? row : []).slice(0, 12).map((cell) => clampText(cell, 500)),
      ),
      header: /** @type {any} */ (src.table).header === true,
    }
  }
  return el
}

/**
 * Normalise one slide.
 * @param {unknown} raw - candidate slide.
 * @param {{index: number, slideW: number, slideH: number}} ctx - index/size context.
 * @returns {Record<string, unknown>} normalised slide.
 */
export function normalizeSlide(raw, ctx) {
  const src = raw !== null && typeof raw === 'object' ? /** @type {Record<string, unknown>} */ (raw) : {}
  const slideId = typeof src.id === 'string' && src.id.trim() !== '' ? src.id.slice(0, 64) : `s${ctx.index + 1}`
  const elementsRaw = Array.isArray(src.elements) ? src.elements.slice(0, MAX_ELEMENTS_PER_SLIDE) : []
  /** @type {Record<string, unknown>} */
  const slide = {
    id: slideId,
    elements: elementsRaw.map((el, i) => normalizeElement(el, {
      index: i, slideW: ctx.slideW, slideH: ctx.slideH, prefix: `${slideId}-`,
    })),
  }
  if (typeof src.layout === 'string') slide.layout = slice(src.layout, 80)
  if (src.layoutIndex !== undefined) slide.layoutIndex = num(src.layoutIndex, 0, 0, 200)
  if (typeof src.title === 'string') slide.title = slice(src.title, 200)
  if (typeof src.notes === 'string') slide.notes = clampText(src.notes, 4000)
  return slide
}

/**
 * Slice a string helper.
 * @param {unknown} value - candidate.
 * @param {number} max - maximum length.
 * @returns {string} sliced string.
 */
function slice(value, max) {
  return String(value ?? '').slice(0, max)
}

/**
 * Create a fresh deck model.
 * @param {{id: string, title?: string, width?: number, height?: number, template?: string|null}} options - deck options.
 * @returns {Record<string, unknown>} deck model.
 */
export function createDeck(options) {
  return {
    version: DECK_VERSION,
    id: options.id,
    title: clampText(options.title || '未命名演示文稿', 200),
    size: {
      w: num(options.width, 13.333, 2, 40),
      h: num(options.height, 7.5, 2, 40),
    },
    theme: { colors: {}, fonts: { latin: '微软雅黑', ea: '微软雅黑' } },
    template: options.template ? { path: options.template } : null,
    slides: [],
    meta: { createdAt: new Date().toISOString(), updatedAt: new Date().toISOString() },
  }
}

/**
 * Normalise an entire deck model (e.g. loaded from disk).
 * @param {unknown} raw - candidate deck.
 * @param {string} fallbackId - id when the model has none.
 * @returns {Record<string, unknown>} normalised deck, or `null` when unusable.
 */
export function normalizeDeck(raw, fallbackId) {
  if (raw === null || typeof raw !== 'object') return null
  const src = /** @type {Record<string, unknown>} */ (raw)
  const size = src.size !== null && typeof src.size === 'object' ? /** @type {any} */ (src.size) : {}
  const deck = createDeck({
    id: typeof src.id === 'string' && src.id !== '' ? src.id : fallbackId,
    title: typeof src.title === 'string' ? src.title : '未命名演示文稿',
    width: size.w,
    height: size.h,
    template: null,
  })
  if (src.template !== null && typeof src.template === 'object' && typeof /** @type {any} */ (src.template).path === 'string') {
    deck.template = { .../** @type {any} */ (src.template) }
    deck.template.path = String(deck.template.path).slice(0, 1024)
  }
  if (src.theme !== null && typeof src.theme === 'object') deck.theme = src.theme
  if (src.meta !== null && typeof src.meta === 'object') deck.meta = src.meta
  const slides = Array.isArray(src.slides) ? src.slides.slice(0, MAX_SLIDES) : []
  const { w, h } = /** @type {any} */ (deck.size)
  deck.slides = slides.map((slide, i) => normalizeSlide(slide, { index: i, slideW: w, slideH: h }))
  return deck
}

/**
 * Deep-ish merge used by `patch`: objects merge one level, everything else replaces.
 * @param {Record<string, unknown>} target - element fields.
 * @param {Record<string, unknown>} patch - incoming fields.
 * @returns {Record<string, unknown>} merged element.
 */
function mergeElement(target, patch) {
  /** @type {Record<string, unknown>} */
  const out = { ...target }
  for (const [key, value] of Object.entries(patch)) {
    if (value === null) {
      delete out[key]
    } else if (value !== null && typeof value === 'object' && !Array.isArray(value) && out[key] !== null && typeof out[key] === 'object' && !Array.isArray(out[key])) {
      out[key] = { .../** @type {Record<string, unknown>} */ (out[key]), .../** @type {Record<string, unknown>} */ (value) }
    } else {
      out[key] = value
    }
  }
  return out
}

/**
 * Apply one patch to an element in place.
 * @param {Record<string, unknown>} deck - deck model (mutated).
 * @param {{slideId: string, elementId: string, patch: Record<string, unknown>, remove?: boolean}} change - change request.
 * @returns {{ok: boolean, error?: string, element?: Record<string, unknown>, removed?: boolean}} result.
 */
export function applyElementPatch(deck, change) {
  const slides = /** @type {any[]} */ (deck.slides)
  const slide = slides.find((s) => s.id === change.slideId)
  if (!slide) return { ok: false, error: `slide not found: ${change.slideId}` }
  const elements = slide.elements
  const index = elements.findIndex((el) => el.id === change.elementId)
  if (change.remove === true) {
    if (index < 0) return { ok: false, error: `element not found: ${change.elementId}` }
    elements.splice(index, 1)
    return { ok: true, removed: true }
  }
  const { w, h } = /** @type {any} */ (deck.size)
  const normalised = normalizeElement(mergeElement(index >= 0 ? elements[index] : { id: change.elementId }, change.patch), {
    index: index >= 0 ? index : elements.length,
    slideW: w,
    slideH: h,
    prefix: `${slide.id}-`,
  })
  normalised.id = change.elementId
  if (index >= 0) elements[index] = normalised
  else elements.push(normalised)
  return { ok: true, element: normalised }
}

/**
 * Compact summary of a deck for the model's `inspect` tool / cards.
 * @param {Record<string, unknown>} deck - deck model.
 * @param {string|null} slideId - optional slide filter.
 * @returns {Record<string, unknown>} summary.
 */
export function summarizeDeck(deck, slideId = null) {
  const slides = /** @type {any[]} */ (deck.slides)
  return {
    id: deck.id,
    title: deck.title,
    size: deck.size,
    template: deck.template ? deck.template.path : null,
    slideCount: slides.length,
    slides: slides
      .filter((slide) => slideId === null || slide.id === slideId)
      .map((slide) => ({
        index: slides.indexOf(slide),
        id: slide.id,
        title: slide.title || textOf(slide),
        elementCount: slide.elements.length,
        elements: slide.elements.map((el) => {
          /** @type {Record<string, unknown>} */
          const row = { id: el.id, type: el.type, box: el.box }
          if (typeof el.text === 'string') row.text = el.text.slice(0, 120)
          return row
        }),
      })),
  }
}

/**
 * First text of a slide, used as its display title.
 * @param {any} slide - slide model.
 * @returns {string} text.
 */
function textOf(slide) {
  const first = (slide.elements || []).find((el) => el.type === 'text' && typeof el.text === 'string' && el.text.trim() !== '')
  return first ? String(first.text).split('\n')[0].slice(0, 60) : ''
}
