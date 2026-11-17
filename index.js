/**
 * dsh-pptx-studio — host half.
 *
 * One live deck document per workspace entry, plus:
 *   - model-facing tools to open / plan / write / patch / inspect / export a deck,
 *   - an SSE stream so the right-side canvas can draw slides as they arrive,
 *   - a feedback route that turns "this slide, this element, this note" into a
 *     real user message in the owning session.
 *
 * The deck model lives in `lib/model.js`, persistence in `lib/store.js`, and
 * the actual .pptx writer (python-pptx) is spawned per export through
 * `lib/python.js`.
 * @module dsh-pptx-studio
 */
import { createHash } from 'node:crypto'
import { copyFileSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { dirname, extname, isAbsolute, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

import { DeckStore, STORE_DIR, ensureInside, makeDeckId } from './lib/store.js'
import {
  MAX_ELEMENTS_PER_SLIDE,
  MAX_PREVIEW_HTML,
  MAX_SLIDES,
  applyElementPatch,
  auditDeck,
  auditElement,
  auditSlide,
  clampText,
  collectElementIssues,
  collectSlideIssues,
  mergeElement,
  normalizeBox,
  normalizeElement,
  normalizePipeline,
  normalizePreview,
  normalizeResearch,
  normalizeSlide,
  normalizeTheme,
  pipelineStage,
  pipelineSummary,
  round,
  summarizeDeck,
} from './lib/model.js'
import { PythonRunner } from './lib/python.js'
import { registerBundledSkills } from './lib/skills.js'
import {
  buildSlideFromTemplate,
  checkMapping,
  normalizeFill,
  normalizeTemplateSpec,
  pagesFromDeck,
  previewConsistency,
  renderTemplateHtml,
  templateCatalogue,
  templatePage,
  templateSummary,
} from './lib/template.js'

export const inject = ['tools', 'webServer', 'connection']

/** Route prefix served by the Web carrier. */
const API_PREFIX = '/dsh-pptx-studio/api'
/** Bundled python helpers. */
const ASSETS_DIR = fileURLToPath(new URL('./assets', import.meta.url))
/** Default slide geometry (16:9, inches). */
const DEFAULT_SIZE = { w: 13.333, h: 7.5 }
/** SSE heartbeat. */
const HEARTBEAT_MS = 15_000
/** Decoded image cap for `POST /upload`. */
const MAX_UPLOAD_BYTES = 20 * 1024 * 1024
/** Request-body cap for `POST /upload` (base64 is ~4/3 of the decoded size). */
const UPLOAD_BODY_LIMIT = 32 * 1024 * 1024
/** Deck-relative prefix used by image elements. */
const ASSET_SRC_PREFIX = 'assets/'
/** Extension → MIME type map shared by `/asset` and `/upload`. */
const ASSET_TYPES = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
  '.svg': 'image/svg+xml',
}
/** Content type of a built .pptx. */
const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
/** File name of the HTML design preview inside the deck directory. */
const PREVIEW_FILE = 'preview.html'

/** Decomposed-template spec, relative to a deck directory. */
const TEMPLATE_SPEC_FILE = 'template/template.json'

/** Where the decomposer writes plates / thumbnails, relative to a deck directory. */
const TEMPLATE_DIR = 'template'
/** Accepted `<img>` upload data URLs. */
const IMAGE_DATA_URL = /^data:image\/(png|jpe?g|gif|webp|svg\+xml);base64,([A-Za-z0-9+/=\s]+)$/
/** Every op kind `POST /edit` accepts. */
const EDIT_KINDS = new Set([
  'element.patch',
  'element.replace',
  'element.add',
  'element.duplicate',
  'element.reorder',
  'slide.add',
  'slide.remove',
  'slide.move',
  'slide.patch',
  'deck.title',
  'deck.theme',
  'deck.size',
])

const JSON_HEADERS = {
  'content-type': 'application/json; charset=utf-8',
  'cache-control': 'no-store',
}

/**
 * Small helper: await a delay (used to pace slide-by-slide drawing).
 *
 * The timer is deliberately *ref'd*: this promise is awaited, so letting the
 * loop consider itself idle would let Node exit mid-draw (and the e2e harness
 * would hang as "unsettled top-level await").
 * @param {number} ms - milliseconds.
 * @returns {Promise<void>} resolved after the delay.
 */
function sleep(ms) {
  return new Promise((resolvePromise) => {
    setTimeout(resolvePromise, ms)
  })
}

/**
 * How long the canvas is told "this page is being prepared" before the page
 * content is written. Without this beat a fast tool call would land as one
 * frame; with it the user sees the pencil arrive, then the page.
 */
const DRAW_BEAT_MS = 220

/**
 * Pace the skeleton pages of `pptx_studio_plan` so they appear one by one
 * instead of all at once.
 */
const PLAN_BEAT_MS = 160

/**
 * Write a JSON response.
 * @param {import('node:http').ServerResponse} res - response.
 * @param {number} status - status code.
 * @param {unknown} value - JSON body.
 * @returns {void}
 */
function writeJson(res, status, value) {
  const body = JSON.stringify(value)
  res.writeHead(status, { ...JSON_HEADERS, 'content-length': Buffer.byteLength(body) })
  res.end(body)
}

/**
 * Read a JSON body with a size cap.
 * @param {import('node:http').IncomingMessage} req - request.
 * @param {number} [limit] - max bytes.
 * @returns {Promise<Record<string, unknown>>} parsed body.
 */
function readJsonBody(req, limit = 2 * 1024 * 1024) {
  return new Promise((resolvePromise, reject) => {
    /** @type {Buffer[]} */
    const chunks = []
    let total = 0
    let settled = false
    req.on('data', (chunk) => {
      if (settled) return
      total += chunk.length
      if (total > limit) {
        settled = true
        chunks.length = 0
        reject(new Error('request body too large'))
        return
      }
      chunks.push(chunk)
    })
    req.on('end', () => {
      if (settled) return
      settled = true
      const text = Buffer.concat(chunks).toString('utf8').trim()
      if (text === '') return resolvePromise({})
      try {
        const parsed = JSON.parse(text)
        resolvePromise(parsed !== null && typeof parsed === 'object' ? parsed : {})
      } catch (error) {
        reject(new Error(`invalid JSON body: ${error instanceof Error ? error.message : String(error)}`))
      }
    })
    req.on('error', (error) => {
      if (settled) return
      settled = true
      reject(error)
    })
  })
}

/** Output schema shared by every pptx-studio tool (loose object). */
const OUTPUT_SCHEMA = { type: 'object', additionalProperties: true }

/**
 * Element schema shared by the write tools.
 *
 * Declared as data rather than prose so a caller that emits the wrong shape is
 * told which field is wrong; `normalizeElement` is forgiving (it also re-reads
 * decks from disk and takes canvas patches), and that forgiveness is exactly
 * what makes a bad payload look like a plausible slide.
 */
const ELEMENT_SCHEMA = {
  type: 'object',
  additionalProperties: true,
  description: '元素：{type, box:[x,y,w,h], text?, style?, fill?, src?, table?}。box 必须是 4 个数字（英寸），宽高 > 0。',
  properties: {
    // Types the canvas *and* the pptx builder fully draw. `group` stays
    // accepted by the model for compatibility but renders as a placeholder.
    type: { type: 'string', enum: ['text', 'rect', 'image', 'line', 'table', 'chart'], description: '元素类型；缺省按 text 处理。' },
    box: { type: 'array', items: { type: 'number' }, description: '英寸坐标 [x, y, w, h]，恰好 4 个数字。' },
    text: { type: 'string', description: '文本内容（type=text 必填）。' },
    style: { type: 'object', additionalProperties: true, description: '{size, bold, italic, underline, color, font, align, anchor, lineSpacing, spaceAfter}。' },
    fill: { type: 'string', description: '填充色，如 22D3EE 或 #22D3EE。' },
    line: { type: 'string', description: '描边色。' },
    src: { type: 'string', description: 'type=image 时的图片路径（assets/... 或绝对路径）。' },
    radius: { type: 'number', description: '圆角（英寸）。' },
    lineWidth: { type: 'number', description: '描边宽度（pt）。' },
    table: { type: 'object', additionalProperties: true, description: '{rows: [[cell, ...], ...], header?: boolean}。' },
    chart: {
      type: 'object',
      additionalProperties: true,
      description: 'type=chart 时必填：{kind: column/bar/line/pie/doughnut/area, title?, categories: [...], series: [{name, values: [...], color?}]}。',
    },
    shadow: { type: 'object', additionalProperties: true, description: '阴影：true 或 {blur, distance, direction, color, alpha}。' },
    anim: { type: 'object', additionalProperties: true, description: '入场动画：{effect: appear/fade/wipe/zoom/none, direction, duration, delay, order}。' },
    paragraphs: { type: 'array', description: '[{text, size?, bold?, italic?, underline?, color?, align?, runs?: [{text, size?, bold?, color?, font?}]}]，多段文本。' },
    id: { type: 'string', description: '元素 id；缺省自动生成。' },
  },
}

/**
 * Pick an id that is not used yet, appending `-2`, `-3`, ... on collision.
 * @param {Set<string>} taken - ids already in use.
 * @param {string} base - preferred id.
 * @returns {string} a free id.
 */
function uniqueId(taken, base) {
  const preferred = String(base)
  if (!taken.has(preferred)) return preferred
  let suffix = 2
  while (taken.has(`${preferred}-${suffix}`)) suffix += 1
  return `${preferred}-${suffix}`
}

/**
 * Find a slide by id.
 * @param {Record<string, unknown>} deck - deck model.
 * @param {string} slideId - slide id.
 * @returns {Record<string, any>|null} slide model or null.
 */
function slideById(deck, slideId) {
  const slides = /** @type {any[]} */ (deck.slides)
  return slides.find((slide) => slide.id === slideId) ?? null
}

/**
 * Clamp a 0-based insertion index into `[0, max]`.
 * @param {unknown} value - candidate index (omitted → append).
 * @param {number} max - highest allowed index (array length).
 * @returns {number} integer index.
 */
function clampIndex(value, max) {
  const n = typeof value === 'number' ? value : Number(value)
  if (!Number.isFinite(n)) return max
  return Math.max(0, Math.min(max, Math.floor(n)))
}

/**
 * Fit a box inside the slide bounds (clamps both position and size).
 * @param {unknown} box - candidate `[x, y, w, h]` in inches.
 * @param {number} slideW - slide width in inches.
 * @param {number} slideH - slide height in inches.
 * @returns {number[]} box that stays within the slide.
 */
function insideBox(box, slideW, slideH) {
  const [x, y, w, h] = normalizeBox(box, slideW, slideH)
  const width = Math.min(w, slideW)
  const height = Math.min(h, slideH)
  const left = Math.min(Math.max(0, x), Math.max(0, slideW - width))
  const top = Math.min(Math.max(0, y), Math.max(0, slideH - height))
  return [round(left), round(top), round(width), round(height)]
}

/**
 * Pick a slide id: the supplied one when free, otherwise `s<max+1>`.
 * @param {any[]} slides - current slides.
 * @param {unknown} candidate - supplied slide id, if any.
 * @returns {string} free slide id.
 */
function slideIdFor(slides, candidate) {
  const taken = new Set(slides.map((slide) => String(slide.id)))
  const preferred = typeof candidate === 'string' && candidate.trim() !== '' ? candidate.trim().slice(0, 64) : ''
  if (preferred !== '' && !taken.has(preferred)) return preferred
  let max = 0
  for (const id of taken) {
    const match = /^s(\d+)$/.exec(id)
    if (match) max = Math.max(max, Number(match[1]))
  }
  let next = max + 1
  while (taken.has(`s${next}`)) next += 1
  return `s${next}`
}

/**
 * Decode one base64 image data URL.
 * @param {unknown} value - `data:image/...;base64,...` string.
 * @returns {{ok: true, ext: string, bytes: Buffer}|{ok: false}} decoded bytes.
 */
function decodeImageDataUrl(value) {
  if (typeof value !== 'string') return { ok: false }
  const match = IMAGE_DATA_URL.exec(value.trim())
  if (!match) return { ok: false }
  const base64 = match[2].replace(/\s+/g, '')
  if (base64 === '' || base64.length % 4 !== 0 || !/^[A-Za-z0-9+/]+={0,2}$/.test(base64)) return { ok: false }
  let bytes
  try {
    bytes = Buffer.from(base64, 'base64')
  } catch {
    return { ok: false }
  }
  if (bytes.length === 0) return { ok: false }
  const subtype = match[1].toLowerCase()
  const ext = subtype === 'svg+xml' ? 'svg' : subtype === 'jpeg' ? 'jpg' : subtype
  return { ok: true, ext, bytes }
}

/**
 * Content type for one served file, by extension.
 * @param {string} file - absolute file path.
 * @returns {string} MIME type.
 */
function contentTypeOf(file) {
  return ASSET_TYPES[extname(file).toLowerCase()] || 'application/octet-stream'
}

/**
 * Build a safe `Content-Disposition` filename pair for one deck title.
 * @param {unknown} title - deck title.
 * @param {string} [extension] - file extension without the dot.
 * @returns {{ascii: string, encoded: string}} plain and RFC 5987 file names.
 */
function downloadFilename(title, extension = 'pptx') {
  const raw = deckFileStem(title)
  const ascii = raw.replace(/[^\x20-\x7e]/g, '_') || 'deck'
  return { ascii: `${ascii}.${extension}`, encoded: encodeURIComponent(`${raw}.${extension}`) }
}

/**
 * Turn a deck title into a file name stem that every common filesystem accepts:
 * no separators, no reserved characters, no trailing dots or spaces.
 * @param {unknown} title - deck title.
 * @returns {string} file name without extension.
 */
function deckFileStem(title) {
  return clampText(String(title ?? ''), 200)
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[. ]+$/, '')
    .slice(0, 80) || 'deck'
}

/**
 * Resolve one user-picked folder into the `.pptx` path it should receive.
 * A save must never eat a file the user already has, so the name steps aside
 * (`deck (1).pptx`) instead of overwriting.
 * @param {string} dir - absolute directory the user picked.
 * @param {unknown} title - deck title.
 * @returns {string} absolute output path that does not exist yet.
 */
function uniquePptxTarget(dir, title) {
  const stem = deckFileStem(title)
  for (let index = 0; index < 1000; index += 1) {
    const candidate = join(dir, index === 0 ? `${stem}.pptx` : `${stem} (${index}).pptx`)
    if (!existsSync(candidate)) return candidate
  }
  return join(dir, `${stem}-${Date.now()}.pptx`)
}

/**
 * Whether `candidate` stays inside `root` (never throws).
 * @param {string} root - allowed root directory.
 * @param {string} candidate - resolved candidate path.
 * @returns {boolean} true when admitted.
 */
function isInside(root, candidate) {
  try {
    ensureInside(root, candidate)
    return true
  } catch {
    return false
  }
}

/**
 * Build the SSE payload for one element change (the same shape the patch tool
 * emits, plus the element's index when it still exists).
 * @param {Record<string, any>} slide - slide model.
 * @param {string} elementId - element id.
 * @param {Record<string, any>|null} element - element model (null when removed).
 * @param {number} index - current index in `slide.elements`, or -1.
 * @returns {Record<string, unknown>} event payload.
 */
function elementEvent(slide, elementId, element, index) {
  /** @type {Record<string, unknown>} */
  const event = { type: 'element', slideId: slide.id, elementId, element: element ?? null, removed: element === null }
  if (index >= 0) {
    event.index = index
    event.total = slide.elements.length
  }
  return event
}

/**
 * Apply one canvas edit op in place and describe the change to broadcast.
 *
 * The op vocabulary mirrors what the canvas needs for direct manipulation:
 * element patch/replace/add/duplicate/reorder and slide add/remove/move plus
 * the deck title. Every successful outcome carries exactly one SSE event; a
 * failure carries a machine code and never throws.
 * @param {import('./lib/store.js').DeckEntry} entry - deck entry (model mutated in place).
 * @param {Record<string, any>} op - one op of `POST /edit`.
 * @returns {{ok: boolean, code?: string, message?: string, result?: Record<string, unknown>, event?: Record<string, unknown>}} outcome.
 */
function applyEditOp(entry, op) {
  const deck = entry.model
  const slides = /** @type {any[]} */ (deck.slides)
  const { w: slideW, h: slideH } = /** @type {any} */ (deck.size)
  const kind = String(op.kind)
  const slideId = typeof op.slideId === 'string' ? op.slideId : ''
  /** @param {string} code - machine code. @param {string} message - human reason. */
  const failOp = (code, message) => ({ ok: false, code, message })
  /** @param {string} reason - missing slide message. */
  const noSlide = (reason) => failOp('slide-not-found', reason)

  if (kind === 'element.patch') {
    const elementId = typeof op.elementId === 'string' ? op.elementId : ''
    const change = {
      slideId,
      elementId,
      patch: op.patch !== null && typeof op.patch === 'object' ? op.patch : {},
      remove: op.remove === true,
    }
    const patched = applyElementPatch(deck, change)
    if (!patched.ok) return failOp('patch-failed', patched.error || 'patch failed')
    const slide = slideById(deck, slideId)
    const index = slide && patched.element ? slide.elements.findIndex((el) => el.id === elementId) : -1
    return {
      ok: true,
      result: { ok: true, kind, slideId, elementId, element: patched.element ?? null, removed: patched.removed === true },
      event: elementEvent(slide, elementId, patched.element ?? null, index),
    }
  }

  if (kind === 'element.replace') {
    const slide = slideById(deck, slideId)
    if (!slide) return noSlide(`slide not found: ${slideId}`)
    if (op.element === null || typeof op.element !== 'object') return failOp('edit-failed', 'element required')
    const source = /** @type {Record<string, any>} */ (op.element)
    const requested = typeof op.elementId === 'string' && op.elementId !== ''
      ? op.elementId
      : typeof source.id === 'string' && source.id !== '' ? source.id : ''
    const index = requested === '' ? -1 : slide.elements.findIndex((el) => el.id === requested)
    const id = requested !== ''
      ? requested
      : uniqueId(new Set(slide.elements.map((el) => el.id)), `${slide.id}-e${slide.elements.length + 1}`)
    const normalized = normalizeElement({ ...source, id }, {
      index: index >= 0 ? index : slide.elements.length,
      slideW,
      slideH,
      prefix: `${slide.id}-`,
    })
    normalized.id = id
    if (index >= 0) slide.elements[index] = normalized
    else slide.elements.push(normalized)
    const at = index >= 0 ? index : slide.elements.length - 1
    return {
      ok: true,
      result: { ok: true, kind, slideId, elementId: id, element: normalized, index: at },
      event: elementEvent(slide, id, normalized, at),
    }
  }

  if (kind === 'element.add') {
    const slide = slideById(deck, slideId)
    if (!slide) return noSlide(`slide not found: ${slideId}`)
    if (slide.elements.length >= MAX_ELEMENTS_PER_SLIDE) return failOp('element-limit', `每页最多 ${MAX_ELEMENTS_PER_SLIDE} 个元素`)
    const raw = op.element !== null && typeof op.element === 'object' ? op.element : {}
    const at = clampIndex(op.index, slide.elements.length)
    const normalized = normalizeElement(raw, { index: at, slideW, slideH, prefix: `${slide.id}-` })
    normalized.id = uniqueId(new Set(slide.elements.map((el) => el.id)), normalized.id)
    if (at >= slide.elements.length) slide.elements.push(normalized)
    else slide.elements.splice(at, 0, normalized)
    const index = slide.elements.indexOf(normalized)
    return {
      ok: true,
      result: { ok: true, kind, slideId, elementId: normalized.id, element: normalized, index },
      event: elementEvent(slide, normalized.id, normalized, index),
    }
  }

  if (kind === 'element.duplicate') {
    const slide = slideById(deck, slideId)
    if (!slide) return noSlide(`slide not found: ${slideId}`)
    const elementId = typeof op.elementId === 'string' ? op.elementId : ''
    const source = slide.elements.find((el) => el.id === elementId)
    if (!source) return failOp('element-not-found', `element not found: ${elementId}`)
    if (slide.elements.length >= MAX_ELEMENTS_PER_SLIDE) return failOp('element-limit', `每页最多 ${MAX_ELEMENTS_PER_SLIDE} 个元素`)
    const offset = Array.isArray(op.offset) && op.offset.length >= 2
      ? [Number(op.offset[0]) || 0, Number(op.offset[1]) || 0]
      : [0.2, 0.2]
    const box = source.box
    const shifted = Array.isArray(box)
      ? [Number(box[0]) + offset[0], Number(box[1]) + offset[1], box[2], box[3]]
      : box
    const copyId = uniqueId(new Set(slide.elements.map((el) => el.id)), `${elementId}-copy`)
    const copy = normalizeElement({ ...source, id: copyId, box: insideBox(shifted, slideW, slideH) }, {
      index: slide.elements.length,
      slideW,
      slideH,
      prefix: `${slide.id}-`,
    })
    copy.id = copyId
    // Keep the copy next to its source: the canvas applies the same op locally
    // before the SSE echo arrives, so both sides must land on the same order.
    const index = slide.elements.indexOf(source) + 1
    slide.elements.splice(index, 0, copy)
    return {
      ok: true,
      result: { ok: true, kind, slideId, elementId: copyId, element: copy, index },
      event: elementEvent(slide, copyId, copy, index),
    }
  }

  if (kind === 'element.reorder') {
    const slide = slideById(deck, slideId)
    if (!slide) return noSlide(`slide not found: ${slideId}`)
    const elementId = typeof op.elementId === 'string' ? op.elementId : ''
    const from = slide.elements.findIndex((el) => el.id === elementId)
    if (from < 0) return failOp('element-not-found', `element not found: ${elementId}`)
    let to
    if (op.to === 'front') to = slide.elements.length - 1
    else if (op.to === 'back') to = 0
    else {
      const n = Number(op.to)
      if (!Number.isFinite(n)) return failOp('edit-failed', `bad reorder target: ${String(op.to)}`)
      to = Math.max(0, Math.min(slide.elements.length - 1, Math.floor(n)))
    }
    const [element] = slide.elements.splice(from, 1)
    slide.elements.splice(to, 0, element)
    return {
      ok: true,
      result: { ok: true, kind, slideId, elementId, element, index: to },
      event: elementEvent(slide, elementId, element, to),
    }
  }

  if (kind === 'slide.add') {
    if (slides.length >= MAX_SLIDES) return failOp('too-many-slides', `最多 ${MAX_SLIDES} 页`)
    const at = clampIndex(op.index, slides.length)
    const raw = op.slide !== null && typeof op.slide === 'object' ? op.slide : {}
    const id = slideIdFor(slides, raw.id)
    const slide = normalizeSlide({ ...raw, id }, { index: at, slideW, slideH })
    slides.splice(at, 0, slide)
    return {
      ok: true,
      result: { ok: true, kind, slideId: slide.id, slide, index: at, total: slides.length },
      event: { type: 'slide', slide, index: at, total: slides.length },
    }
  }

  if (kind === 'slide.remove') {
    const index = slides.findIndex((slide) => slide.id === slideId)
    if (index < 0) return noSlide(`slide not found: ${slideId}`)
    if (slides.length <= 1) return failOp('last-slide', '不能删除最后一页幻灯片')
    slides.splice(index, 1)
    return {
      ok: true,
      result: { ok: true, kind, slideId, removed: true, index, total: slides.length },
      event: { type: 'deck', deck: entry.model, mode: 'snapshot' },
    }
  }

  if (kind === 'slide.move') {
    const from = slides.findIndex((slide) => slide.id === slideId)
    if (from < 0) return noSlide(`slide not found: ${slideId}`)
    const n = Number(op.index)
    if (!Number.isFinite(n)) return failOp('edit-failed', `bad slide index: ${String(op.index)}`)
    const to = Math.max(0, Math.min(slides.length - 1, Math.floor(n)))
    const [slide] = slides.splice(from, 1)
    slides.splice(to, 0, slide)
    return {
      ok: true,
      result: { ok: true, kind, slideId, index: to, total: slides.length },
      event: { type: 'slide', slide, index: to, total: slides.length },
    }
  }

  if (kind === 'slide.patch') {
    const slide = slideById(deck, slideId)
    if (!slide) return noSlide(`slide not found: ${slideId}`)
    const index = slides.indexOf(slide)
    const patch = op.patch !== null && typeof op.patch === 'object' ? /** @type {Record<string, unknown>} */ (op.patch) : {}
    /** @type {Record<string, unknown>} */
    const merged = { ...slide }
    for (const [key, value] of Object.entries(patch)) {
      if (value === null) delete merged[key]
      else merged[key] = value
    }
    const normalized = normalizeSlide({ ...merged, id: slide.id }, { index, slideW, slideH })
    slides[index] = normalized
    return {
      ok: true,
      result: { ok: true, kind, slideId, slide: normalized, index, total: slides.length },
      event: { type: 'slide', slide: normalized, index, total: slides.length },
    }
  }

  if (kind === 'deck.theme') {
    const incoming = op.theme !== null && typeof op.theme === 'object' ? /** @type {Record<string, any>} */ (op.theme) : {}
    const current = /** @type {any} */ (deck.theme) || normalizeTheme(null)
    deck.theme = normalizeTheme({
      colors: { ...current.colors, ...(incoming.colors || {}) },
      fonts: { ...current.fonts, ...(incoming.fonts || {}) },
    })
    return {
      ok: true,
      result: { ok: true, kind, theme: deck.theme },
      event: { type: 'deck', deck: entry.model },
    }
  }

  if (kind === 'deck.size') {
    const width = Number(op.width)
    const height = Number(op.height)
    if (!Number.isFinite(width) || !Number.isFinite(height) || width < 2 || width > 40 || height < 2 || height > 40) {
      return failOp('edit-failed', 'deck.size 需要 2..40 英寸的 width/height')
    }
    deck.size = { w: round(width), h: round(height) }
    const { w, h } = /** @type {any} */ (deck.size)
    deck.slides = slides.map((slide, index) => normalizeSlide(slide, { index, slideW: w, slideH: h }))
    return {
      ok: true,
      result: { ok: true, kind, size: deck.size, total: /** @type {any[]} */ (deck.slides).length },
      event: { type: 'deck', deck: entry.model },
    }
  }

  if (kind === 'deck.title') {
    const title = clampText(op.title ?? '', 200)
    deck.title = title.trim() === '' ? '未命名演示文稿' : title
    return {
      ok: true,
      result: { ok: true, kind, title: deck.title },
      event: { type: 'deck', deck: entry.model },
    }
  }

  return failOp('unknown-op', `unsupported op: ${kind}`)
}

/** Human labels for the three workflow stages. */
const STAGE_LABELS = { research: '① 调研分析', preview: '② HTML 效果', build: '③ 幻灯片/导出' }

/**
 * Render one tool receipt back to the model: a short status line, never the
 * whole deck (the canvas already has it; echoing it would double the tokens).
 * @param {Record<string, any>} _args - tool args (unused).
 * @param {Record<string, any>} value - execute() result.
 * @returns {Array<{type: string, text: string}>} content blocks.
 */
function renderReceipt(_args, value) {
  if (!value || value.ok === false) {
    const lines = [`pptx-studio 未完成：${value?.code ?? 'error'}${value?.hint ? ` — ${value.hint}` : ''}`]
    // A refusal is only actionable if it names what is wrong; the list used to
    // live in the JSON payload the model never sees.
    if (Array.isArray(value?.issues) && value.issues.length > 0) {
      lines.push(`问题 ${value.issues.length} 处：${value.issues.slice(0, 8).join('；')}`)
    }
    if (Array.isArray(value?.warnings) && value.warnings.length > 0) {
      lines.push(`提醒 ${value.warnings.length} 条：${value.warnings.slice(0, 5).join('；')}`)
    }
    return [{ type: 'text', text: lines.join('\n') }]
  }
  const parts = []
  // `inspect` answers with the deck model itself; summarise it for the model
  // instead of echoing the whole structure (the canvas already shows it).
  if (value.deck && typeof value.deck === 'object') {
    const deck = /** @type {any} */ (value.deck)
    if (deck.id) parts.push(`deck=${deck.id}`)
    if (deck.pipeline?.stage) parts.push(`阶段=${STAGE_LABELS[deck.pipeline.stage] || deck.pipeline.stage}`)
    if (typeof deck.slideCount === 'number') parts.push(`${deck.slideCount} 页`)
    if (Array.isArray(deck.slides)) {
      parts.push(deck.slides.map((slide) => `${slide.index + 1}:${slide.id}(${slide.elementCount}元素)`).join(' '))
    }
    return [{ type: 'text', text: parts.join('；') }]
  }
  if (value.deckId) parts.push(`deck=${value.deckId}`)
  if (value.stage) parts.push(`阶段=${STAGE_LABELS[value.stage] || value.stage}`)
  else if (value.pipeline?.stage) parts.push(`阶段=${STAGE_LABELS[value.pipeline.stage] || value.pipeline.stage}`)
  if (value.mode) parts.push(value.mode === 'template' ? '模板模式' : '空白模式')
  if (value.total !== undefined && typeof value.slides !== 'number') parts.push(`共 ${value.total} 页`)
  if (value.slides !== undefined && typeof value.slides === 'number') parts.push(`${value.slides} 页`)
  if (value.index !== undefined) parts.push(`第 ${value.index} 页`)
  if (value.elements !== undefined && typeof value.elements === 'number') parts.push(`${value.elements} 个元素`)
  if (value.outline !== undefined && typeof value.outline === 'number') parts.push(`大纲 ${value.outline} 页`)
  if (value.sources !== undefined && typeof value.sources === 'number') parts.push(`来源 ${value.sources} 条`)
  if (value.previewBytes !== undefined && typeof value.previewBytes === 'number' && !value.path) {
    parts.push(`HTML 预览 ${Math.round(value.previewBytes / 1024)}KB`)
  }
  if (value.features && typeof value.features === 'object') {
    const f = /** @type {any} */ (value.features)
    const bits = [
      f.transitions ? `切换 ${f.transitions} 页` : '',
      f.animations ? `动画 ${f.animations} 处` : '',
      f.notes ? `备注 ${f.notes} 页` : '',
      f.charts ? `图表 ${f.charts} 个` : '',
      f.hidden ? `隐藏 ${f.hidden} 页` : '',
    ].filter(Boolean)
    if (bits.length > 0) parts.push(bits.join('/'))
  }
  if (value.path) {
    if (value.previewBytes !== undefined) {
      parts.push(`HTML 预览已写入 ${value.path}（${Math.round(Number(value.previewBytes) / 1024)}KB）`)
    } else {
      parts.push(`已导出 ${value.path}（${value.bytes ?? '?'} 字节）`)
    }
  }
  const text = parts.length > 0 ? parts.join('；') : 'ok'
  const lines = [text]
  // Normalisation is forgiving on purpose, so a bad payload would otherwise be
  // "fixed" silently and the model would never learn its arguments were wrong.
  if (Array.isArray(value.issues) && value.issues.length > 0) {
    lines.push(`参数已修正 ${value.issues.length} 处：${value.issues.join('；')}`)
  }
  if (value.hint) lines.push(String(value.hint))
  return [{ type: 'text', text: lines.join('\n') }]
}

/**
 * Answer with one failure object the model can act on.
 * @param {string} code - machine code.
 * @param {string} hint - one concrete next step.
 * @returns {Record<string, unknown>} failure receipt.
 */
function fail(code, hint) {
  return { ok: false, code, hint }
}

/**
 * Normalise `{{deckId}}` style arguments (tools accept `deck` or `deckId`).
 * @param {Record<string, unknown>} args - tool args.
 * @returns {string} deck id.
 */
function deckArg(args) {
  const raw = args.deck ?? args.deckId ?? args.studio
  return typeof raw === 'string' ? raw.trim() : ''
}

/**
 * Activate the plugin.
 * @param {object} ctx - Cordis host context.
 * @param {Record<string, unknown>} [config] - row config.
 * @returns {Promise<void>} resolved once everything is registered.
 */
/**
 * Load the in-box `defineTool` through the loader seam.
 *
 * It validates model-generated arguments against the flat parameter DSL and
 * adds presentation hooks. When the loader cannot provide it, the raw
 * definition still registers: one safety net degrades instead of the plugin
 * breaking.
 * @param {any} ctx - Cordis host context.
 * @param {(level: string, message: string) => void} log - logger sink.
 * @returns {Promise<(definition: any) => any>} the factory, or identity.
 */
async function loadDefineTool(ctx, log) {
  try {
    const pkg = await ctx.loader.import('@deepseek-ai/dsh-tools')
    const factory = pkg?.defineTool ?? pkg?.default?.defineTool
    if (typeof factory === 'function') return factory
    log('warn', 'pptx-studio: @deepseek-ai/dsh-tools 没有 defineTool，退化为无 schema 校验')
  } catch (error) {
    log('warn', `pptx-studio: 无法导入 @deepseek-ai/dsh-tools（${error instanceof Error ? error.message : String(error)}），退化为无 schema 校验`)
  }
  return (definition) => definition
}

export async function apply(ctx, config = {}) {
  const log = (level, message) => {
    const sink = ctx.logger?.[level]
    if (typeof sink === 'function') sink.call(ctx.logger, message)
  }
  const store = new DeckStore({ log })
  const python = new PythonRunner({
    assetsDir: ASSETS_DIR,
    log,
    configured: typeof config.python === 'string' ? config.python : undefined,
  })
  /** Template summaries are small; keeping them in memory is enough. */
  const templateInfo = new Map()
  /** Decomposed template specs, keyed by deck id (disk copy is the truth). */
  const templateSpecs = new Map()
  /** Plate file -> data URL, so the HTML sheet does not re-encode on every open. */
  const plateCache = new Map()
  /** deck id -> latest generated HTML sheet (keyed by the mapping it came from). */
  const templateSheetCache = new Map()

  // ── helpers ────────────────────────────────────────────────────────────────

  /**
   * Resolve the calling session: id + working directory.
   * @param {any} exec - tool execution input (`exec.agent.session`).
   * @returns {{id: string|null, cwd: string}} session info.
   */
  function sessionOf(exec) {
    const session = exec?.agent?.session
    const id = typeof session?.id === 'string' && session.id !== '' ? session.id : null
    let cwd = typeof session?.cwd === 'string' && session.cwd !== '' ? session.cwd : ''
    if (cwd === '' && id !== null) {
      try {
        const workingDirectory = ctx.get('workingDirectory')
        const sessionObject = ctx.get('sessions')?.get?.(id)
        if (workingDirectory && sessionObject) cwd = workingDirectory.get(sessionObject) || ''
      } catch {
        /* fall back to process.cwd() */
      }
    }
    return { id, cwd: cwd || process.cwd() }
  }

  /**
   * Working directory of a live session, used by the HTTP routes (which only
   * get a session id, not the tool `exec` object).
   * @param {string|null} sessionId - session id.
   * @returns {string} cwd, or '' when it cannot be resolved.
   */
  function sessionCwd(sessionId) {
    if (!sessionId) return ''
    try {
      const workingDirectory = ctx.get('workingDirectory')
      const sessionObject = ctx.get('sessions')?.get?.(sessionId)
      if (workingDirectory && sessionObject) return workingDirectory.get(sessionObject) || ''
    } catch {
      /* fall through */
    }
    return ''
  }

  /**
   * Decks that exist on disk for one workspace but are not in memory (the
   * process restarted, or the plugin row was reloaded: `deck.json` survives,
   * the in-memory map does not).
   * @param {string} cwd - workspace directory.
   * @returns {Array<{id: string, title: string, slides: number, updatedAt: string|null}>} summaries.
   */
  function decksOnDisk(cwd) {
    if (!cwd) return []
    const root = join(store.rootOf(cwd), STORE_DIR)
    /** @type {Array<{id: string, title: string, slides: number, updatedAt: string|null}>} */
    const out = []
    let names = []
    try {
      names = readdirSync(root, { withFileTypes: true }).filter((item) => item.isDirectory()).map((item) => item.name)
    } catch {
      return out
    }
    for (const name of names) {
      const file = join(root, name, 'deck.json')
      if (!existsSync(file)) continue
      try {
        const model = JSON.parse(readFileSync(file, 'utf8'))
        out.push({
          id: typeof model.id === 'string' && model.id !== '' ? model.id : name,
          title: typeof model.title === 'string' ? model.title : name,
          slides: Array.isArray(model.slides) ? model.slides.length : 0,
          updatedAt: model.meta && typeof model.meta.updatedAt === 'string' ? model.meta.updatedAt : null,
        })
      } catch {
        /* one broken deck.json must not hide the others */
      }
    }
    return out
  }

  /**
   * Resolve a deck from memory, reopening it from disk when the process has
   * forgotten it (page refresh after a plugin reload must still see the deck).
   * @param {string} deckId - deck id.
   * @param {{sessionId?: string, cwd?: string}} [hint] - where to look on disk.
   * @returns {import('./lib/store.js').DeckEntry|null} entry.
   */
  function entryFor(deckId, hint = {}) {
    const known = store.get(deckId)
    if (known) return known
    if (!/^[A-Za-z0-9_-]{4,80}$/.test(deckId)) return null
    const cwd = hint.cwd || sessionCwd(hint.sessionId ?? '')
    if (!cwd) return null
    const dir = join(store.rootOf(cwd), STORE_DIR, deckId)
    if (!existsSync(join(dir, 'deck.json'))) return null
    try {
      return store.open({ deckId, cwd, sessionId: hint.sessionId ?? null })
    } catch (error) {
      log('warn', `pptx-studio: 从磁盘恢复 deck 失败 ${deckId}: ${error instanceof Error ? error.message : String(error)}`)
      return null
    }
  }

  /**
   * Resolve a workspace-relative path and keep it inside the workspace.
   * @param {string} cwd - workspace directory.
   * @param {string|undefined} candidate - user/model supplied path.
   * @param {string} fallback - relative fallback.
   * @returns {string} absolute path.
   */
  function workspacePath(cwd, candidate, fallback) {
    const raw = typeof candidate === 'string' && candidate.trim() !== '' ? candidate.trim() : fallback
    const absolute = isAbsolute(raw) ? raw : resolve(cwd, raw)
    return ensureInside(cwd, absolute)
  }

  /**
   * Resolve the template path a caller passed.
   *
   * Workspace paths are used as-is. A path outside the workspace (the common
   * case: the user points at `~/Downloads/xxx.pptx`) is copied into the plugin
   * store first, so every later step — decomposition, rendering, export — reads
   * a file that lives inside the workspace and the sandbox boundary stays
   * intact. The copy is content-addressed, so re-opening the same template does
   * not duplicate it.
   * @param {string} cwd - workspace root.
   * @param {string} raw - caller path.
   * @returns {{path: string}|{error: string}} resolved path or a reason.
   */
  function resolveTemplateInput(cwd, raw) {
    const value = String(raw ?? '').trim()
    if (value === '') return { error: '模板路径为空' }
    const absolute = isAbsolute(value) ? value : resolve(cwd, value)
    if (!existsSync(absolute)) return { error: `模板不存在：${absolute}` }
    try {
      if (statSync(absolute).isFile() && /^[\s\S]*\.pptx$/i.test(absolute)) {
        // `ensureInside` throws when the path escapes the workspace.
        try {
          const inside = ensureInside(cwd, absolute)
          return { path: inside }
        } catch {
          /* outside: import below */
        }
        const target = join(store.rootOf(cwd), STORE_DIR, '_templates')
        mkdirSync(target, { recursive: true })
        const name = `${createHash('sha1').update(absolute).update(String(statSync(absolute).size)).digest('hex').slice(0, 12)}.pptx`
        const copy = join(target, name)
        if (!existsSync(copy)) copyFileSync(absolute, copy)
        log('info', `pptx-studio: 模板来自工作区外，已导入 ${copy}`)
        return { path: copy, importedFrom: absolute }
      }
    } catch (error) {
      return { error: error instanceof Error ? error.message : String(error) }
    }
    return { error: `模板不是 .pptx 文件：${absolute}` }
  }

  /**
   * Load a deck or answer with a model-actionable failure.
   * @param {Record<string, unknown>} args - tool args.
   * @returns {{entry: import('./lib/store.js').DeckEntry}|{error: Record<string, unknown>}} result.
   */
  function requireDeck(args, exec) {
    const id = deckArg(args)
    if (id === '') return { error: fail('deck-required', '先调用 pptx_studio_open 创建画布，并把这个工具返回的 deckId 传进 deck 参数。') }
    let entry = store.get(id)
    // A host reload forgets the in-memory map; the deck.json on disk is still
    // the truth. Reopen it from the calling session's workspace when possible.
    if (!entry && exec) {
      const { cwd } = sessionOf(exec)
      entry = entryFor(id, { cwd }) ?? undefined
    }
    if (!entry) return { error: fail('deck-not-found', `没有这个 deck：${id}。用 pptx_studio_open 重新打开。`) }
    return { entry }
  }

  /**
   * Keep `pipeline.stage` derived from its contents.
   * @param {import('./lib/store.js').DeckEntry} entry - deck entry.
   * @returns {Record<string, unknown>} pipeline.
   */
  function syncPipeline(entry) {
    if (!entry.model.pipeline) entry.model.pipeline = normalizePipeline(null)
    const pipeline = entry.model.pipeline
    pipeline.stage = pipelineStage(pipeline)
    return /** @type {Record<string, unknown>} */ (pipeline)
  }

  /**
   * Enforce the analysis → HTML preview → PPTX workflow for model-facing tools.
   *
   * The workflow is the point of the plugin, so a fresh deck cannot jump
   * straight to `plan`/`slide`/`export`. Two escapes keep it practical:
   * `force:true` for a deliberate shortcut, and decks that already contain
   * slides (legacy decks or canvas-authored work) are grandfathered.
   * @param {import('./lib/store.js').DeckEntry} entry - deck entry.
   * @param {boolean} force - explicit bypass from the tool args.
   * @returns {Record<string, unknown>|null} failure receipt or null.
   */
  function stageGate(entry, force) {
    if (force === true) return null
    const pipeline = entry.model.pipeline || normalizePipeline(null)
    if (/** @type {any} */ (pipeline).builtAt) return null
    const slides = /** @type {any[]} */ (entry.model.slides) || []
    if (slides.length > 0) return null
    const template = /** @type {any} */ (pipeline).template
    // Template mode: the requirement cannot be decomposed before we know what
    // the template actually offers, so decomposition comes first.
    if (template && template.required === true && !template.decomposedAt) {
      return fail('stage-template', '先调用 pptx_studio_template 把模板分解成「可复用版式 + 文字槽」（右侧画布会显示每页的角色和能填哪些字段），再写需求分析；确实要跳过就传 force:true。')
    }
    if (!/** @type {any} */ (pipeline).research) {
      return fail('stage-research', '先用 pptx_studio_research 提交需求分析/调研（画布会显示这一阶段）；确实要跳过流程就传 force:true。')
    }
    if (!/** @type {any} */ (pipeline).preview) {
      return fail('stage-preview', '先用 pptx_studio_preview 提交 HTML 效果预览（画布会用 iframe 显示）；确实要跳过流程就传 force:true。')
    }
    return null
  }

  /**
   * Absolute path of the HTML preview file for one deck.
   * @param {import('./lib/store.js').DeckEntry} entry - deck entry.
   * @returns {string} path.
   */
  function previewFile(entry) {
    return join(entry.dir, PREVIEW_FILE)
  }

  /**
   * SSE event for a pipeline change.
   * @param {import('./lib/store.js').DeckEntry} entry - deck entry.
   * @returns {Record<string, unknown>} event.
   */
  function pipelineEvent(entry) {
    const pipeline = /** @type {any} */ (entry.model.pipeline || normalizePipeline(null))
    return { type: 'pipeline', pipeline, stage: pipeline.stage ?? 'research' }
  }

  /**
   * Read the template layout info (cached per deck, persisted as template.json).
   * @param {import('./lib/store.js').DeckEntry} entry - deck entry.
   * @returns {Record<string, unknown>|null} template summary.
   */
  function templateOf(entry) {
    const cached = templateInfo.get(entry.model.id)
    if (cached) return cached
    const file = join(entry.dir, 'template.json')
    if (existsSync(file)) {
      try {
        const parsed = JSON.parse(readFileSync(file, 'utf8'))
        templateInfo.set(entry.model.id, parsed)
        return parsed
      } catch {
        /* re-inspect below */
      }
    }
    return null
  }

  /**
   * Load the decomposed template spec (memory first, then the deck on disk).
   * @param {import('./lib/store.js').DeckEntry} entry - deck entry.
   * @returns {Record<string, any>|null} spec.
   */
  function templateSpecOf(entry) {
    const cached = templateSpecs.get(entry.model.id)
    if (cached) return cached
    const file = join(entry.dir, TEMPLATE_SPEC_FILE)
    if (!existsSync(file)) return null
    try {
      const spec = normalizeTemplateSpec(JSON.parse(readFileSync(file, 'utf8')), /** @type {any} */ (entry.model.size))
      if (spec) {
        templateSpecs.set(entry.model.id, spec)
        return spec
      }
      log('warn', `pptx-studio: ${TEMPLATE_SPEC_FILE} 无法解析，忽略`)
    } catch (error) {
      log('warn', `pptx-studio: ${TEMPLATE_SPEC_FILE} 读取失败：${error instanceof Error ? error.message : String(error)}`)
    }
    return null
  }

  /**
   * Persist a decomposed spec next to the deck.
   * @param {import('./lib/store.js').DeckEntry} entry - deck entry.
   * @param {Record<string, any>|null} spec - spec or null to forget.
   * @returns {void}
   */
  function storeTemplateSpec(entry, spec) {
    if (!spec) templateSpecs.delete(entry.model.id)
    else templateSpecs.set(entry.model.id, spec)
    templateSheetCache.delete(entry.model.id)
    const dir = join(entry.dir, TEMPLATE_DIR)
    try {
      if (spec) {
        mkdirSync(dir, { recursive: true })
        writeFileSync(join(dir, 'template.json'), JSON.stringify(spec, null, 2), 'utf8')
      }
    } catch (error) {
      log('warn', `pptx-studio: 模板 spec 落盘失败：${error instanceof Error ? error.message : String(error)}`)
    }
  }

  /**
   * Data URL for one template plate, for the sandboxed HTML sheet.
   *
   * The preview iframe runs under `default-src 'none'` with `img-src data:` only,
   * so images have to travel inside the document; thumbnails (not the full-size
   * plates) keep that document a few hundred KB instead of tens of MB.
   * @param {import('./lib/store.js').DeckEntry} entry - deck entry.
   * @param {Record<string, any>} page - template page.
   * @returns {string|null} data URL.
   */
  function plateDataUrl(entry, page) {
    const candidates = []
    if (typeof page.plate === 'string' && page.plate !== '') {
      const name = page.plate.replace(/\.[A-Za-z0-9]+$/, '')
      candidates.push(`${name}-thumb.jpg`, `${name}.jpg`, page.plate)
    }
    for (const relative of candidates) {
      const file = join(entry.dir, relative)
      if (!existsSync(file)) continue
      const hit = plateCache.get(file)
      if (hit) return hit
      try {
        const bytes = readFileSync(file)
        // A 1600px plate is ~200KB; the sheet only needs the thumbnail.
        if (bytes.length > 400 * 1024) continue
        const mime = /\.jpe?g$/i.test(relative) ? 'image/jpeg' : 'image/png'
        const url = `data:${mime};base64,${bytes.toString('base64')}`
        plateCache.set(file, url)
        return url
      } catch {
        /* try the next candidate */
      }
    }
    return null
  }

  /**
   * Render (and cache) the HTML sheet for a template-mode preview.
   * @param {import('./lib/store.js').DeckEntry} entry - deck entry.
   * @param {Record<string, any>} spec - template spec.
   * @param {Array<Record<string, any>>} pages - page mapping.
   * @param {string} [title] - sheet title.
   * @returns {{html: string, width: number, height: number, missing: number[]}} sheet.
   */
  function templateSheet(entry, spec, pages, title) {
    const key = createHash('sha1').update(JSON.stringify(pages)).update(String(spec.version || '')).digest('hex')
    const cached = templateSheetCache.get(entry.model.id)
    if (cached && cached.key === key && cached.html && (title === undefined || cached.title === title)) return cached
    const sheet = renderTemplateHtml(spec, pages, {
      designWidth: 1280,
      title: title ?? 'HTML 模版',
      readImage: (page) => plateDataUrl(entry, page),
    })
    templateSheetCache.set(entry.model.id, { key, html: sheet.html, width: sheet.width, height: sheet.height, title })
    return sheet
  }

  /**
   * The page→copy mapping for a template-mode deck: the explicit preview
   * mapping when there is one, otherwise the outline the research step wrote.
   * @param {import('./lib/store.js').DeckEntry} entry - deck entry.
   * @returns {Array<Record<string, any>>|null} mapping.
   */
  function mappingOf(entry) {
    const preview = /** @type {any} */ (entry.model.pipeline || {}).preview
    if (preview && preview.mode === 'template' && Array.isArray(preview.pages) && preview.pages.length > 0) {
      return preview.pages
    }
    const research = /** @type {any} */ (entry.model.pipeline || {}).research
    const outline = Array.isArray(research?.outline) ? research.outline : []
    const mapped = outline.filter((item) => Number.isInteger(item.templateSlide))
    if (mapped.length === 0) return null
    return mapped.map((item) => ({ templateSlide: item.templateSlide, title: item.title || '', fields: item.fields || {} }))
  }

  /**
   * Decompose `deck.template.path` into the reusable spec, rendering plates.
   * @param {import('./lib/store.js').DeckEntry} entry - deck entry.
   * @param {{pages?: number|null, render?: boolean}} [options] - limits.
   * @returns {Promise<{spec: Record<string, any>|null, error: string|null, raw: Record<string, any>|null}>} result.
   */
  async function decomposeTemplate(entry, options = {}) {
    const templatePath = /** @type {any} */ (entry.model.template)?.path
    if (typeof templatePath !== 'string' || templatePath === '') {
      return { spec: null, error: '这个 deck 没有模板：先用 pptx_studio_open 传 template 建画布。', raw: null }
    }
    if (!existsSync(templatePath)) {
      return { spec: null, error: `模板文件不在了：${templatePath}`, raw: null }
    }
    const outDir = join(entry.dir, TEMPLATE_DIR)
    mkdirSync(outDir, { recursive: true })
    const args = [templatePath, outDir]
    const pages = Number.isFinite(options.pages) && /** @type {number} */ (options.pages) > 0 ? Math.floor(/** @type {number} */ (options.pages)) : 0
    if (pages > 0) args.push('--max-slides', String(pages))
    if (options.render === false) args.push('--no-render')
    const raw = await python.run('pptx_template.py', args, { timeoutMs: 360_000 })
    if (!raw || raw.ok !== true) {
      const message = raw && typeof raw.error === 'string' ? raw.error : '模板分解未返回结果'
      return { spec: null, error: message, raw: /** @type {any} */ (raw) }
    }
    const spec = normalizeTemplateSpec(raw, /** @type {any} */ (entry.model.size))
    if (!spec) return { spec: null, error: '模板分解结果不可用（没有解析出任何页）', raw: /** @type {any} */ (raw) }
    return { spec: /** @type {any} */ (spec), error: null, raw: /** @type {any} */ (raw) }
  }

  /**
   * Trim a sample string for a one-line catalogue.
   * @param {string} value - sample text.
   * @returns {string} shortened sample.
   */
  function shortSample(value) {
    const text = String(value ?? '').replace(/\s+/g, ' ').trim()
    return text.length > 14 ? `${text.slice(0, 14)}…` : text
  }

  /**
   * Default geometry for a title element, using the template's title
   * placeholder when one exists.
   * @param {Record<string, unknown>|null} info - template summary.
   * @param {Record<string, unknown>} deck - deck model.
   * @param {string} text - title text.
   * @returns {Record<string, unknown>} text element.
   */
  function titleElement(info, deck, text) {
    const size = /** @type {any} */ (deck.size) || DEFAULT_SIZE
    let box = [0.8, 0.6, size.w - 1.6, 0.9]
    let style = { size: 30, bold: true, color: '#0F172A' }
    const layouts = /** @type {any[]} */ (info?.layouts || [])
    const layout = layouts[1] || layouts[0]
    const title = layout?.placeholders?.find?.((ph) => String(ph.type).startsWith('title') || String(ph.type).startsWith('center_title'))
    if (title && Array.isArray(title.box) && title.box.length === 4) {
      const [x, y, w, h] = title.box
      box = [x, y, w, h]
      style = { size: 28, bold: true, color: '#0F172A' }
    }
    return { id: 'title', type: 'text', box, text, style }
  }

  /**
   * Broadcast and persist one change.
   * @param {import('./lib/store.js').DeckEntry} entry - deck entry.
   * @param {Record<string, unknown>} event - event payload.
   * @returns {Record<string, unknown>} emitted event.
   */
  function emit(entry, event) {
    return store.changed(entry, event)
  }

  // ── session activity feed ─────────────────────────────────────────────────

  /** Human label of one tool, shown on the canvas progress bar and its card. */
  const TOOL_LABELS = {
    pptx_studio_open: '打开画布',
    pptx_studio_template: '拆解模板',
    pptx_studio_research: '分析与调研',
    pptx_studio_preview: '生成 HTML 效果稿',
    pptx_studio_deck: '更新演示文稿设置',
    pptx_studio_plan: '规划页面骨架',
    pptx_studio_slide: '绘制幻灯片',
    pptx_studio_add_element: '添加元素',
    pptx_studio_patch: '修改画布',
    pptx_studio_inspect: '检查画布',
    pptx_studio_verify: '一致性校验',
    pptx_studio_export: '导出 PPTX',
  }

  /**
   * Per-session feed of the `pptx_studio_*` call in flight. The canvas tab
   * subscribes to it for one reason: a tab that is showing the previous build
   * has to learn, the moment a call starts, *which* deck that call writes — and
   * switch to it — instead of leaving the old pptx on screen while the agent
   * draws the new one somewhere else.
   *
   * @type {Map<string, {task: Record<string, any>|null, listeners: Set<(task: Record<string, any>) => void>}>}
   */
  const sessionFeeds = new Map()

  /**
   * The feed of one session (created on first use).
   * @param {string} sessionId - session id (`''` for a stateless feed).
   * @returns {{task: Record<string, any>|null, listeners: Set<(task: Record<string, any>) => void>}} feed.
   */
  function feedOf(sessionId) {
    const key = typeof sessionId === 'string' ? sessionId : ''
    let feed = sessionFeeds.get(key)
    if (!feed) {
      feed = { task: null, listeners: new Set() }
      if (key !== '') sessionFeeds.set(key, feed)
    }
    return feed
  }

  /**
   * Publish the current call of one session. `seq` only grows, so a canvas that
   * reconnects can tell an old snapshot from a new one.
   * @param {string} sessionId - session id.
   * @param {Record<string, any>} patch - fields to merge.
   * @returns {Record<string, any>|null} the published task (null without a session).
   */
  function publishTask(sessionId, patch) {
    if (typeof sessionId !== 'string' || sessionId === '') return null
    const feed = feedOf(sessionId)
    const previous = feed.task
    const now = new Date().toISOString()
    const next = {
      sessionId,
      tool: String(patch.tool ?? previous?.tool ?? ''),
      label: String(patch.label ?? previous?.label ?? ''),
      deckId: patch.deckId !== undefined ? String(patch.deckId ?? '') : String(previous?.deckId ?? ''),
      state: String(patch.state ?? 'running'),
      message: String(patch.message ?? ''),
      startedAt: String(patch.startedAt ?? previous?.startedAt ?? now),
      updatedAt: now,
      seq: Number(previous?.seq ?? 0) + 1,
    }
    feed.task = next
    for (const listener of [...feed.listeners]) {
      try {
        listener(next)
      } catch {
        /* a broken subscriber must never break the tool */
      }
    }
    return next
  }

  /**
   * Subscribe to one session's feed.
   * @param {string} sessionId - session id.
   * @param {(task: Record<string, any>) => void} listener - sink.
   * @returns {() => void} unsubscribe.
   */
  function subscribeTask(sessionId, listener) {
    const feed = feedOf(sessionId)
    feed.listeners.add(listener)
    return () => feed.listeners.delete(listener)
  }

  /**
   * Wrap one tool so every call reports its start, its deck and its outcome to
   * the session feed. Wrapping here (instead of inside each tool) keeps the
   * progress truthful for all twelve steps, including the ones that only patch
   * an existing deck.
   * @param {Record<string, any>} tool - tool definition.
   * @returns {Record<string, any>} wrapped definition.
   */
  function trackedTool(tool) {
    const label = TOOL_LABELS[tool.name] ?? tool.name
    return {
      ...tool,
      /**
       * @param {Record<string, any>} args - tool args.
       * @param {any} exec - execution context.
       * @returns {Promise<Record<string, unknown>>} receipt.
       */
      async execute(args, exec) {
        const sessionId = String(sessionOf(exec).id ?? '')
        const hinted = typeof args?.deck === 'string' && args.deck.trim() !== ''
          ? args.deck.trim()
          : (typeof args?.deckId === 'string' ? args.deckId.trim() : '')
        const startedAt = new Date().toISOString()
        // A new call starts a new task: the deck it names (or a blank one for
        // `open`, whose deck does not exist yet) is what the canvas must show.
        publishTask(sessionId, { tool: tool.name, label, deckId: hinted, state: 'running', message: '', startedAt })
        let result
        try {
          result = await tool.execute(args, exec)
        } catch (error) {
          publishTask(sessionId, {
            tool: tool.name,
            label,
            state: 'failed',
            message: error instanceof Error ? error.message : String(error),
            startedAt,
          })
          throw error
        }
        const resultDeckId = result && typeof result.deckId === 'string' && result.deckId !== '' ? result.deckId : ''
        const ok = !(result && result.ok === false)
        publishTask(sessionId, {
          tool: tool.name,
          label,
          ...(resultDeckId !== '' || hinted !== '' ? { deckId: resultDeckId || hinted } : {}),
          state: ok ? 'done' : 'failed',
          message: ok ? '' : String(result?.message ?? result?.code ?? ''),
          startedAt,
        })
        return result
      },
    }
  }

  /**
   * Stream one session's activity feed (`/api/session/events`).
   * @param {import('node:http').IncomingMessage} req - request.
   * @param {import('node:http').ServerResponse} res - response.
   * @param {string} sessionId - session id.
   * @returns {void}
   */
  function serveSessionEvents(req, res, sessionId) {
    res.writeHead(200, {
      'content-type': 'text/event-stream; charset=utf-8',
      'cache-control': 'no-cache, no-transform',
      connection: 'keep-alive',
      'x-accel-buffering': 'no',
    })
    const send = (event) => {
      try {
        res.write(`data: ${JSON.stringify(event)}\n\n`)
      } catch {
        /* client vanished; the close handler cleans up */
      }
    }
    const sendTask = (task) => send({ type: 'task', task, seq: Number(task?.seq ?? 0) })
    res.write('retry: 2000\n\n')
    sendTask(feedOf(sessionId).task)
    const unsubscribe = subscribeTask(sessionId, sendTask)
    const heartbeat = setInterval(() => {
      try {
        res.write(`: ping ${Date.now()}\n\n`)
      } catch {
        /* handled by close */
      }
    }, HEARTBEAT_MS)
    if (typeof heartbeat.unref === 'function') heartbeat.unref()
    const cleanup = () => {
      clearInterval(heartbeat)
      unsubscribe()
    }
    req.on('close', cleanup)
    req.on('error', cleanup)
    res.on('close', cleanup)
  }

  // ── tools ──────────────────────────────────────────────────────────────────

  const toolLog = (name, entry, detail) => {
    log('debug', `pptx-studio: ${name} deck=${entry.model.id} ${detail}`)
  }

  /**
   * `pptx_studio_open`: create or resume the canvas.
   */
  const openTool = {
    name: 'pptx_studio_open',
    output: { schema: OUTPUT_SCHEMA, render: renderReceipt },
    description:
      '打开一块 PPTX 实时画布（DSH 右侧栏会显示绘制过程）。给模板路径即进入模板模式（沿用模板的母版/版式/主题色），不给则从空白 16:9 开始。返回 deckId，后续所有 pptx_studio_* 工具都要带上它。',
    parameters: {
      title: { type: 'string', description: '演示文稿标题。' },
      deckId: { type: 'string', description: '续写已有画布时传上次返回的 deckId。' },
      template: { type: 'string', description: '工作区内的 .pptx 模板路径（模板模式）。' },
      width: { type: 'number', description: '画布宽度（英寸），默认 13.333（16:9）。' },
      height: { type: 'number', description: '画布高度（英寸），默认 7.5。' },
    },
    /**
     * @param {Record<string, any>} args - tool args.
     * @param {any} exec - execution context.
     * @returns {Promise<Record<string, unknown>>} receipt.
     */
    async execute(args, exec) {
      const { id: sessionId, cwd } = sessionOf(exec)
      const root = store.rootOf(cwd)
      let templatePath = null
      let info = null
      if (typeof args.template === 'string' && args.template.trim() !== '') {
        const resolved = resolveTemplateInput(root, args.template)
        if (/** @type {any} */ (resolved).error) {
          return fail('template-missing', `${/** @type {any} */ (resolved).error}。给一个存在的 .pptx 路径（工作区内或绝对路径都行）。`)
        }
        templatePath = /** @type {any} */ (resolved).path
        try {
          info = await python.run('pptx_inspect.py', [templatePath])
        } catch (error) {
          return fail('template-unreadable', `模板解析失败：${error instanceof Error ? error.message : String(error)}；可先不传 template 用空白画布。`)
        }
      }
      const entry = store.open({
        deckId: typeof args.deckId === 'string' ? args.deckId.trim() : undefined,
        cwd: root,
        sessionId,
        title: typeof args.title === 'string' ? args.title : undefined,
        width: args.width ?? info?.size?.w ?? DEFAULT_SIZE.w,
        height: args.height ?? info?.size?.h ?? DEFAULT_SIZE.h,
        template: templatePath ?? undefined,
      })
      if (info) {
        templateInfo.set(entry.model.id, info)
        try {
          writeFileSync(join(entry.dir, 'template.json'), JSON.stringify(info, null, 2), 'utf8')
        } catch (error) {
          log('warn', `pptx-studio: template.json 写入失败：${error instanceof Error ? error.message : String(error)}`)
        }
      }
      // Template mode always reports where it is: the first job is decomposing
      // the template, not writing slides.
      if (templatePath) {
        const pipeline = /** @type {any} */ (entry.model.pipeline || normalizePipeline(null))
        if (!pipeline.template || pipeline.template.path === '') {
          pipeline.template = { path: templatePath, required: true, decomposedAt: null, pages: 0, engine: '', textHidden: false, design: null }
        }
        pipeline.template.required = true
        entry.model.pipeline = pipeline
        entry.model.template = { ...(entry.model.template || {}), path: templatePath }
      }
      syncPipeline(entry)
      entry.persist()
      // A resumed deck keeps its template: report the mode from the model, not
      // from whether this particular call supplied one.
      const templateMode = Boolean(/** @type {any} */ (entry.model.template)?.path)
      emit(entry, { type: 'deck', deck: entry.model, mode: templateMode ? 'template' : 'blank' })
      emit(entry, {
        type: 'status',
        phase: 'ready',
        total: /** @type {any[]} */ (entry.model.slides).length,
        message: '画布已就绪',
      })
      toolLog('open', entry, `template=${templateMode ? 'yes' : 'no'} slides=${/** @type {any[]} */ (entry.model.slides).length}`)
      return {
        ok: true,
        deckId: entry.model.id,
        dir: entry.dir,
        mode: templateMode ? 'template' : 'blank',
        size: entry.model.size,
        slides: /** @type {any[]} */ (entry.model.slides).length,
        template: info
          ? {
              path: templatePath,
              layouts: /** @type {any[]} */ (info.layouts || []).length,
              colors: Object.keys(/** @type {any} */ (info.theme?.colors || {})).length,
            }
          : null,
        stage: pipelineSummary(entry.model).stage,
        pipeline: pipelineSummary(entry.model),
        hint: templateMode
          ? '模板模式。按流程走：⓪ pptx_studio_template 分解模板（得到每页版式/角色/可填字段）→ ① pptx_studio_research 做需求分解（outline 每项写 templateSlide + fields）→ ② pptx_studio_preview 自动按同一份映射生成 HTML 模版 → ③ pptx_studio_plan / pptx_studio_slide 传 templateSlide + fields 落成幻灯片 → pptx_studio_verify 校验一致 → pptx_studio_export。'
          : '右侧画布已打开。按流程走：① pptx_studio_research 交需求分析与调研 → ② pptx_studio_preview 交 HTML 效果预览 → ③ pptx_studio_plan / pptx_studio_slide 落成幻灯片 → pptx_studio_export。',
      }
    },
  }

  /**
   * `pptx_studio_template`: step 0 in template mode — decompose the template.
   */
  const templateTool = {
    name: 'pptx_studio_template',
    output: { schema: OUTPUT_SCHEMA, render: renderReceipt },
    description:
      '模板模式的第一步：把用户给的 .pptx 模板分解成「可复用版式」。结果包含每页的角色（封面/目录/章节/内容/结尾）、去掉原文的干净底图（plate）、以及每页可填的文字字段（id、示例、能放多少字）。这一步也是需求分解的依据：拿到 pages 之后，pptx_studio_research 的 outline 每一项要写 templateSlide 与 fields，把需求落到模板页上。',
    parameters: {
      deck: { type: 'string', required: true, description: 'pptx_studio_open 返回的 deckId（建画布时传了 template）。' },
      pages: { type: 'integer', description: '只分解前 N 页（模板很长时先看一部分）。省略=全部。' },
      render: { type: 'boolean', description: '是否用 PowerPoint 渲染底图（默认 true）。false 只解析结构，快，但每页没有干净底图。' },
      list: { type: 'boolean', description: '只读已分解结果，不重新渲染。' },
    },
    /**
     * @param {Record<string, any>} args - tool args.
     * @param {any} exec - execution context.
     * @returns {Promise<Record<string, unknown>>} receipt.
     */
    async execute(args, exec) {
      const found = requireDeck(args, exec)
      if (/** @type {any} */ (found).error) return /** @type {any} */ (found).error
      const entry = /** @type {any} */ (found).entry
      const deck = entry.model
      /** @type {Record<string, any>|null} */
      let spec = templateSpecOf(entry)
      const reading = args.list === true
      if (!spec && reading) {
        return fail('template-not-decomposed', '还没有分解过模板：先调用 pptx_studio_template（不带 list）分解一次。')
      }
      if (!reading) {
        emit(entry, {
          type: 'status',
          phase: 'template',
          op: 'template',
          begin: true,
          message: '正在分解模板：解析版式、提取文字槽、渲染干净底图',
        })
        const result = await decomposeTemplate(entry, {
          pages: Number.isFinite(args.pages) ? Number(args.pages) : null,
          render: args.render !== false,
        })
        if (result.error || !result.spec) {
          emit(entry, { type: 'status', phase: 'ready', message: '模板分解失败' })
          return fail('template-decompose-failed', `模板分解失败：${result.error}。可以传 render:false 只解析结构，或检查模板文件是否可被 PowerPoint 打开。`)
        }
        spec = result.spec
        storeTemplateSpec(entry, spec)
        const pipeline = /** @type {any} */ (syncPipeline(entry))
        pipeline.template = {
          path: spec.path || deck.template?.path || '',
          required: true,
          decomposedAt: new Date().toISOString(),
          pages: spec.pages.length,
          engine: spec.engine,
          textHidden: Boolean(spec.render?.textHidden),
          design: {
            palette: (spec.design?.palette || []).map((/** @type {any} */ e) => e.hex),
            background: spec.design?.background || '',
            fonts: spec.design?.fonts || null,
          },
        }
        deck.template = { ...(deck.template || {}), path: spec.path || deck.template?.path, pages: spec.pages.length, decomposedAt: pipeline.template.decomposedAt }
        // Adopt the template's size and visual language so later slides do not
        // have to be told the palette or the page geometry again.
        if (spec.size && (spec.size.w !== deck.size.w || spec.size.h !== deck.size.h)) {
          deck.size = { w: spec.size.w, h: spec.size.h }
        }
        if (spec.design?.background) {
          const theme = normalizeTheme(deck.theme)
          theme.colors.background = spec.design.background
          if (spec.design.palette?.[0]?.hex) theme.colors.accent = spec.design.palette[0].hex
          if (spec.design.fonts?.body) theme.fonts.ea = spec.design.fonts.body
          if (spec.design.fonts?.body) theme.fonts.latin = spec.design.fonts.latin || spec.design.fonts.body
          if (spec.design.fonts?.title) theme.fonts.heading = spec.design.fonts.title
          deck.theme = theme
        }
        deck.pipeline = pipeline
        emit(entry, pipelineEvent(entry))
        emit(entry, { type: 'deck', deck, mode: 'template' })
        emit(entry, { type: 'status', phase: 'ready', message: `模板已分解：${spec.pages.length} 页` })
        toolLog('template', entry, `pages=${spec.pages.length} engine=${spec.engine}`)
      }
      const summary = templateSummary(/** @type {any} */ (spec))
      const catalogue = templateCatalogue(/** @type {any} */ (spec))
      /** @type {Record<string, unknown>} */
      const receipt = {
        ok: true,
        deckId: deck.id,
        stage: /** @type {any} */ (syncPipeline(entry)).stage,
        template: summary,
        // One compact line per page is what the model needs to map requirements
        // onto the template; the full spec (boxes/styles) stays in the canvas.
        pages: catalogue.map((page) => ({
          page: page.page,
          role: page.role,
          fields: page.fields.map((field) => `${field.id}=${shortSample(field.sample)}（≤${field.maxCjk}字）`).join(' '),
        })),
        hint: '下一步 pptx_studio_research：brief.outline 每项写 { title, templateSlide: <上面的 page>, fields: { 字段id: 文案 } }，把需求直接落到模板页上。',
      }
      if (!spec?.render?.textHidden) {
        receipt.warnings = ['底图不是用 PowerPoint 渲染的（没装 PowerPoint 或渲染失败），页面上原有的文字可能没被去掉，成品需要人工核对。']
      }
      return receipt
    },
  }

  /**
   * `pptx_studio_research`: step 1 — analysis / research brief.
   */
  const researchTool = {
    name: 'pptx_studio_research',
    output: { schema: OUTPUT_SCHEMA, render: renderReceipt },
    description:
      '生成 PPTX 的第一步（必须先做）：把需求分析、调研结论与叙事大纲交给画布。可用 web_search 等工具先调研，再填 brief；画布会显示调研阶段，用户可以在做 HTML/PPT 前先审这份分析。',
    parameters: {
      deck: { type: 'string', required: true, description: 'pptx_studio_open 返回的 deckId。' },
      brief: {
        type: 'object',
        required: true,
        additionalProperties: true,
        description:
          '{summary, audience, goal, tone, constraints, outline:[{title,intent,source}], facts:[{claim,source}], sources:[{title,url,note}], palette:["#RRGGBB"], fonts:{latin,ea,heading}}。outline 的 title 必填。',
      },
    },
    /**
     * @param {Record<string, any>} args - tool args.
     * @returns {Promise<Record<string, unknown>>} receipt.
     */
    async execute(args, exec) {
      const found = requireDeck(args, exec)
      if (/** @type {any} */ (found).error) return /** @type {any} */ (found).error
      const entry = /** @type {any} */ (found).entry
      // Template mode: the requirement cannot be broken down before the
      // template has been decomposed, because the mapping needs its pages.
      const tplPipeline = /** @type {any} */ (entry.model.pipeline || {})
      if (tplPipeline.template?.required === true && !tplPipeline.template.decomposedAt) {
        return fail('stage-template', '先调用 pptx_studio_template 分解模板（得到每页角色与可填字段），再把需求映射到模板页上；确实要跳过就换成不带模板的画布。')
      }
      const brief = normalizeResearch(args.brief)
      if (!brief) return fail('brief-required', 'brief 必须是一个对象；至少给 summary 或 outline。')
      if (brief.summary === '' && brief.outline.length === 0) {
        return fail('brief-empty', 'brief.summary 与 brief.outline 都为空。先写一段分析摘要，并给出每页要讲什么。')
      }
      if (!entry.model.pipeline) entry.model.pipeline = normalizePipeline(null)
      // Template mode: the brief is also the requirement decomposition, so the
      // outline has to say which template page each slide reuses. Checking it
      // here (against the real fields) is what stops a plan from being written
      // for a template that cannot hold it.
      const spec = templateSpecOf(entry)
      if (spec && /** @type {any} */ (entry.model.pipeline).template?.required === true) {
        const mapped = brief.outline.filter((item) => Number.isInteger(item.templateSlide))
        if (mapped.length === 0) {
          return fail('template-mapping-required', '模板模式下 outline 每一项都要写 templateSlide（用模板的哪一页做底版）与 fields（这一页各字段的文案）。先看 pptx_studio_template 返回的 pages，再逐页映射；否则需求没有落到模板上。')
        }
        const checked = checkMapping(spec, brief.outline)
        if (checked.issues.length > 0) {
          return fail('template-mapping-invalid', `需求分解与模板对不上（${checked.issues.length} 处），先修好再继续：` + checked.issues.slice(0, 8).map((issue) => `\n- ${issue}`).join(''))
        }
        brief.outline = brief.outline.map((item, index) => {
          const page = checked.pages[index]
          return page ? { ...item, templateSlide: page.templateSlide, fields: page.fields } : item
        })
        if (checked.warnings.length > 0) {
          log('info', `pptx-studio: 需求分解提醒 ${checked.warnings.length} 条：${checked.warnings[0]}`)
        }
      }
      entry.model.pipeline.research = brief
      syncPipeline(entry)
      emit(entry, pipelineEvent(entry))
      toolLog('research', entry, `outline=${brief.outline.length} sources=${brief.sources.length}`)
      /** @type {Record<string, unknown>} */
      const receipt = {
        ok: true,
        deckId: entry.model.id,
        stage: /** @type {any} */ (entry.model.pipeline).stage,
        outline: brief.outline.length,
        facts: brief.facts.length,
        sources: brief.sources.length,
        hint: spec
          ? '第二步：pptx_studio_preview 只需传 pages:false 或直接调它——模板模式下 HTML 模版会按 outline 的 templateSlide/fields 自动生成（同一份映射，幻灯片会照着它建）。'
          : '第二步：用 pptx_studio_preview 提交自包含 HTML 效果预览（每页一个 .slide、内联 CSS；画布会直接用 iframe 显示）。用户确认视觉方向后，再用 pptx_studio_plan / pptx_studio_slide 落成 PPTX。',
      }
      if (spec) {
        const checked = checkMapping(spec, brief.outline)
        receipt.templateMapping = {
          pages: brief.outline.length,
          templateSlides: brief.outline.map((item) => item.templateSlide ?? null),
        }
        if (checked.warnings.length > 0) receipt.warnings = checked.warnings.slice(0, 10)
      }
      return receipt
    },
  }

  /**
   * `pptx_studio_preview`: step 2 — an HTML design mockup the user can see.
   */
  const previewTool = {
    name: 'pptx_studio_preview',
    output: { schema: OUTPUT_SCHEMA, render: renderReceipt },
    description:
      '生成 PPTX 的第二步：提交一份自包含 HTML 视觉稿（每页一个 .slide，尺寸与 deck 一致，内联 CSS，不依赖外部脚本）。画布会用沙箱 iframe 把它渲染出来，并按 width/height 缩放适配侧栏，让用户在生成 PPTX 前先确认效果。提交后流程进入「幻灯片」阶段。',
    parameters: {
      deck: { type: 'string', required: true, description: 'deckId。' },
      html: { type: 'string', description: '完整的 HTML 文档（含 <style>），内联样式优先；不要用外部脚本。模板模式下不需要（传了就以它为准）。' },
      pages: {
        type: 'array',
        description: '模板模式：逐页映射 [{templateSlide, title?, fields:{字段id:文案}}]。省略时用 pptx_studio_research 里 outline 的 templateSlide/fields（推荐，两边同一份数据）。',
        items: { type: 'object', additionalProperties: true },
      },
      fromDeck: { type: 'boolean', description: '模板模式：用幻灯片现在的文字重建映射（画布上手工改过文字后，用它把 HTML 模版同步过来）。' },
      title: { type: 'string', description: '这份视觉稿的标题。' },
      notes: { type: 'string', description: '设计说明：配色、字体、版式意图。' },
      width: { type: 'integer', description: '视觉稿设计画布宽度（px），例如 960 或 1280；省略时画布从 HTML 或 deck 尺寸推断。' },
      height: { type: 'integer', description: '视觉稿设计画布高度（px），例如 540 或 720；省略时按 deck 比例推断。' },
    },
    /**
     * @param {Record<string, any>} args - tool args.
     * @returns {Promise<Record<string, unknown>>} receipt.
     */
    async execute(args, exec) {
      const found = requireDeck(args, exec)
      if (/** @type {any} */ (found).error) return /** @type {any} */ (found).error
      const entry = /** @type {any} */ (found).entry
      const deck = entry.model
      const rawHtml = typeof args.html === 'string' ? args.html : ''
      const spec = templateSpecOf(entry)
      if (!entry.model.pipeline) entry.model.pipeline = normalizePipeline(null)
      // ── template mode: the sheet is generated from the same mapping the
      // slides will be built from, so HTML and PPTX cannot drift apart. ──────
      if (spec) {
        // `fromDeck` is the "I edited the slides, sync the sheet" switch: the
        // mapping is read back out of the deck's own element provenance.
        const deckPages = args.fromDeck === true
          ? pagesFromDeck(deck).map((page) => ({ templateSlide: page.page, title: '', fields: page.fields }))
          : null
        const provided = deckPages && deckPages.length > 0
          ? deckPages
          : (Array.isArray(args.pages) && args.pages.length > 0
            ? args.pages.map((item) => {
              const one = item !== null && typeof item === 'object' ? /** @type {Record<string, any>} */ (item) : {}
              return { templateSlide: Number(one.templateSlide) || 0, title: typeof one.title === 'string' ? one.title : '', fields: normalizeFill(one.fields) }
            })
            : mappingOf(entry))
        if (!provided || provided.length === 0) {
          return fail('preview-mapping-required', '模板模式需要 pages（或用 pptx_studio_research 的 outline 自带 templateSlide/fields）。先做需求分解：每页指定用模板的哪一页、各字段写什么。')
        }
        const bad = provided.filter((page) => !templatePage(spec, page.templateSlide))
        if (bad.length > 0) {
          return fail('preview-bad-page', `这些页引用了模板里不存在的页：${bad.map((page) => page.templateSlide).join(',')}（模板共 ${spec.pages.length} 页）。`)
        }
        const sheet = templateSheet(entry, spec, provided, typeof args.title === 'string' ? args.title : '')
        if (sheet.missing.length > 0) {
          return fail('preview-missing-plate', `模板第 ${sheet.missing.join(',')} 页没有底图（分解时没渲染成功）；传 render:true 重新跑 pptx_studio_template。`)
        }
        const preview = normalizePreview({
          mode: 'template',
          pages: provided,
          title: typeof args.title === 'string' ? args.title : 'HTML 模版',
          notes: args.notes,
          width: sheet.width,
          height: sheet.height,
        })
        if (!preview) return fail('preview-required', '模板映射为空。')
        entry.model.pipeline.preview = preview
        syncPipeline(entry)
        if (preview.title && entry.model.title === '未命名演示文稿') entry.model.title = preview.title
        try {
          writeFileSync(previewFile(entry), sheet.html, 'utf8')
        } catch (error) {
          return fail('preview-write-failed', `预览已存进模型，但写 preview.html 失败：${error instanceof Error ? error.message : String(error)}`)
        }
        emit(entry, pipelineEvent(entry))
        toolLog('preview', entry, `mode=template pages=${provided.length}`)
        return {
          ok: true,
          deckId: entry.model.id,
          mode: 'template',
          stage: /** @type {any} */ (entry.model.pipeline).stage,
          pages: provided.length,
          templateSlides: provided.map((page) => page.templateSlide),
          previewBytes: Buffer.byteLength(sheet.html, 'utf8'),
          path: previewFile(entry),
          hint: 'HTML 模版已按模板底图 + 同一份字段映射生成（画布「② HTML」可直接看）。第三步：pptx_studio_plan 用同一顺序列页，再 pptx_studio_slide 传 templateSlide + fields 逐页落成幻灯片；最后 pptx_studio_verify 校验一致性后 export。',
        }
      }
      const html = rawHtml
      if (html.trim() === '') return fail('preview-required', 'html 不能为空；给一份完整的 HTML 文档。')
      const bytes = Buffer.byteLength(html, 'utf8')
      if (bytes > MAX_PREVIEW_HTML) {
        return fail('preview-too-large', `HTML 预览 ${bytes} 字节，超过上限 ${MAX_PREVIEW_HTML}；请精简（预览是静态沙箱，大图请改用小尺寸 data URL 或省略）。`)
      }
      const preview = normalizePreview({ html, title: args.title, notes: args.notes, width: args.width, height: args.height })
      if (!preview) return fail('preview-required', 'html 不能为空。')
      entry.model.pipeline.preview = preview
      syncPipeline(entry)
      if (preview.title && entry.model.title === '未命名演示文稿') entry.model.title = preview.title
      try {
        writeFileSync(previewFile(entry), preview.html, 'utf8')
      } catch (error) {
        return fail('preview-write-failed', `预览已存进模型，但写 preview.html 失败：${error instanceof Error ? error.message : String(error)}`)
      }
      emit(entry, pipelineEvent(entry))
      toolLog('preview', entry, `bytes=${bytes}`)
      return {
        ok: true,
        deckId: entry.model.id,
        stage: /** @type {any} */ (entry.model.pipeline).stage,
        previewBytes: bytes,
        path: previewFile(entry),
        hint: '第三步：现在用 pptx_studio_plan 给大纲，再 pptx_studio_slide 逐页把 HTML 设计落成可编辑 PPTX；完成后 pptx_studio_export。',
      }
    },
  }

  /**
   * `pptx_studio_deck`: deck-level settings (title, size, theme).
   */
  const deckTool = {
    name: 'pptx_studio_deck',
    output: { schema: OUTPUT_SCHEMA, render: renderReceipt },
    description:
      '设置整份 deck 的标题、画布尺寸与主题（颜色/字体）。改尺寸会按新画布重新归一化所有元素；主题影响画布与导出的默认字体/色板。',
    parameters: {
      deck: { type: 'string', required: true, description: 'deckId。' },
      title: { type: 'string', description: '演示文稿标题。' },
      width: { type: 'number', description: '画布宽度（英寸），常用 13.333（16:9）/10（4:3）。' },
      height: { type: 'number', description: '画布高度（英寸），常用 7.5。' },
      theme: {
        type: 'object',
        additionalProperties: true,
        description: '{colors:{ink,muted,accent,surface,background}, fonts:{latin,ea,heading}}；只传要改的键。',
      },
    },
    /**
     * @param {Record<string, any>} args - tool args.
     * @returns {Promise<Record<string, unknown>>} receipt.
     */
    async execute(args, exec) {
      const found = requireDeck(args, exec)
      if (/** @type {any} */ (found).error) return /** @type {any} */ (found).error
      const entry = /** @type {any} */ (found).entry
      const deck = entry.model
      if (typeof args.title === 'string' && args.title.trim() !== '') deck.title = clampText(args.title.trim(), 200)
      const width = Number(args.width)
      const height = Number(args.height)
      if (Number.isFinite(width) && width >= 2 && width <= 40) deck.size.w = round(width)
      if (Number.isFinite(height) && height >= 2 && height <= 40) deck.size.h = round(height)
      if (Number.isFinite(width) || Number.isFinite(height)) {
        const { w, h } = /** @type {any} */ (deck.size)
        deck.slides = /** @type {any[]} */ (deck.slides).map((slide, index) => normalizeSlide(slide, { index, slideW: w, slideH: h }))
      }
      if (args.theme !== null && typeof args.theme === 'object') {
        const incoming = /** @type {Record<string, any>} */ (args.theme)
        deck.theme = normalizeTheme({
          colors: { .../** @type {any} */ (deck.theme).colors, ...(incoming.colors || {}) },
          fonts: { .../** @type {any} */ (deck.theme).fonts, ...(incoming.fonts || {}) },
        })
      }
      emit(entry, { type: 'deck', deck })
      toolLog('deck', entry, `size=${deck.size.w}x${deck.size.h}`)
      return {
        ok: true,
        deckId: deck.id,
        title: deck.title,
        size: deck.size,
        theme: deck.theme,
        slides: /** @type {any[]} */ (deck.slides).length,
        hint: 'deck 设置已更新；画布会立刻跟随。',
      }
    },
  }

  /**
   * `pptx_studio_plan`: outline first, so the canvas can draw page skeletons.
   */
  const planTool = {
    name: 'pptx_studio_plan',
    output: { schema: OUTPUT_SCHEMA, render: renderReceipt },
    description:
      '先把页大纲交给画布：每页会出现一张带标题的骨架卡片（用户能看到逐页绘制）。之后用 pptx_studio_slide 按 index 填内容，index 从 1 开始。',
    parameters: {
      deck: { type: 'string', required: true, description: 'pptx_studio_open 返回的 deckId。' },
      slides: {
        type: 'array',
        required: true,
        description: '页大纲，按顺序。',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            title: { type: 'string', required: true, description: '这页的标题/主题。' },
            intent: { type: 'string', description: '这页要讲什么（只用于画布提示，不进 PPT 正文）。' },
            templateSlide: { type: 'integer', description: '模板模式：这页用模板第几页做底版（通常与 research outline 的映射一致）。' },
          },
        },
      },
      force: { type: 'boolean', description: '跳过「调研 → HTML 预览」两步直接写页（不建议；只用于用户明确要求的快速草稿）。' },
    },
    /**
     * @param {Record<string, any>} args - tool args.
     * @param {any} exec - execution context.
     * @returns {Promise<Record<string, unknown>>} receipt.
     */
    async execute(args, exec) {
      const found = requireDeck(args, exec)
      if (/** @type {any} */ (found).error) return /** @type {any} */ (found).error
      const entry = /** @type {any} */ (found).entry
      const gate = stageGate(entry, args.force === true)
      if (gate) return gate
      const planned = Array.isArray(args.slides) ? args.slides.slice(0, MAX_SLIDES) : []
      if (planned.length === 0) return fail('empty-plan', 'slides 至少要有一页 {title}。')
      const deck = entry.model
      const info = templateOf(entry)
      const planSpec = templateSpecOf(entry)
      /** @type {any[]} */
      const slides = []
      // Announce the outline first: the canvas shows "laying out N pages", then
      // each skeleton page lands with its own beat.
      emit(entry, {
        type: 'status',
        phase: 'planning',
        op: 'plan',
        begin: true,
        total: planned.length,
        message: `正在摆放 ${planned.length} 页骨架`,
      })
      for (const [index, item] of planned.entries()) {
        const title = String(item?.title ?? `第 ${index + 1} 页`)
        // Template mode: a skeleton is a placeholder for a template page, so it
        // gets the template page's provenance and *no* invented title box — an
        // off-template element here would have to be cleaned up later and would
        // show up as an inconsistency.
        const templateSlide = Number.isInteger(item?.templateSlide) ? item.templateSlide : mappingOf(entry)?.[index]?.templateSlide
        const slide = planSpec
          ? normalizeSlide(
            { id: `s${index + 1}`, title, templateSlide, elements: [] },
            { index, slideW: deck.size.w, slideH: deck.size.h },
          )
          : normalizeSlide(
            { id: `s${index + 1}`, title, elements: [titleElement(info, deck, title)] },
            { index, slideW: deck.size.w, slideH: deck.size.h },
          )
        if (item?.intent) slide.intent = String(item.intent).slice(0, 300)
        slides.push(slide)
        deck.slides = slides.slice()
        emit(entry, { type: 'slide', slide, index, total: slides.length, phase: 'skeleton' })
        if (index < planned.length - 1) await sleep(PLAN_BEAT_MS)
      }
      emit(entry, { type: 'status', phase: 'planned', slideIndex: 0, total: slides.length, message: `大纲已就位，共 ${slides.length} 页` })
      toolLog('plan', entry, `slides=${slides.length}`)
      return {
        ok: true,
        deckId: deck.id,
        total: slides.length,
        slides: slides.map((slide, index) => ({ index: index + 1, id: slide.id, title: slide.title })),
        hint: '现在用 pptx_studio_slide 按 index 逐页写入内容；每写一页用户就会看到它被画出来。',
      }
    },
  }

  /**
   * `pptx_studio_slide`: replace one slide wholesale.
   */
  const slideTool = {
    name: 'pptx_studio_slide',
    output: { schema: OUTPUT_SCHEMA, render: renderReceipt },
    description:
      '写入/替换一页幻灯片（1-based index）。slide.elements 里每个元素用英寸坐标 box=[x,y,w,h]；type 支持 text/rect/image/line/table/chart。可在页上设置转场 transition、背景 background、备注 notes、隐藏 hidden；元素可带 anim 入场动画与 shadow 阴影。写入后画布会以动画绘制这一页。',
    parameters: {
      deck: { type: 'string', required: true, description: 'deckId。' },
      index: { type: 'number', required: true, description: '页码，从 1 开始；超出当前页数则追加。' },
      templateSlide: { type: 'integer', description: '模板模式：这一页用模板第几页做底版（pptx_studio_template 返回的 page 编号）。传了就按模板字段建页，不需要自己写 elements。' },
      fields: {
        type: 'object',
        additionalProperties: true,
        description: '模板模式：字段文案 {字段id: 文本}。字段 id 来自 pptx_studio_template 的 pages（如 title/subtitle/body1）；位置和字号由模板决定，不会跑版。',
      },
      extra: {
        type: 'array',
        description: '模板模式：在模板字段之外额外加的图元（图表/图片/表格等），坐标照常写英寸。',
        items: ELEMENT_SCHEMA,
      },
      slide: {
        type: 'object',
        additionalProperties: true,
        description: '整页模型：{ title?, layoutIndex?, notes?, hidden?, background?:{color}, transition?:{type,speed,direction,advanceOnClick,advanceAfterMs}, elements:[...] }，元素用英寸坐标 box=[x,y,w,h]。模板模式下可只写 notes/transition/hidden 等页面设置。',
        properties: {
          title: { type: 'string', description: '页标题（画布用）。' },
          notes: { type: 'string', description: '演讲者备注。' },
          layoutIndex: { type: 'integer', description: '模板模式下使用哪个版式（0-based）。' },
          hidden: { type: 'boolean', description: '隐藏这一页（导出后幻灯片仍存在，但放映时跳过）。' },
          background: { type: 'object', additionalProperties: true, description: '页背景：{color:"#RRGGBB"}。' },
          transition: {
            type: 'object',
            additionalProperties: true,
            description: '幻灯片切换：{type: fade/push/wipe/split/cover/pull/cut/zoom/dissolve/circle/random, speed: slow/medium/fast, direction: left/right/up/down, advanceOnClick, advanceAfterMs}。',
          },
          elements: { type: 'array', description: '元素列表，按数组顺序从底到顶叠放。', items: ELEMENT_SCHEMA },
        },
      },
      force: { type: 'boolean', description: '跳过「调研 → HTML 预览」流程（不建议）。' },
    },
    /**
     * @param {Record<string, any>} args - tool args.
     * @param {any} exec - execution context.
     * @returns {Promise<Record<string, unknown>>} receipt.
     */
    async execute(args, exec) {
      const found = requireDeck(args, exec)
      if (/** @type {any} */ (found).error) return /** @type {any} */ (found).error
      const entry = /** @type {any} */ (found).entry
      const gate = stageGate(entry, args.force === true)
      if (gate) return gate
      const deck = entry.model
      const slides = /** @type {any[]} */ (deck.slides)
      const oneBased = Math.max(1, Math.floor(Number(args.index) || 1))
      const index = oneBased - 1
      if (oneBased > MAX_SLIDES) return fail('too-many-slides', `最多 ${MAX_SLIDES} 页。`)
      const raw = args.slide !== null && typeof args.slide === 'object' ? args.slide : {}
      // ── template mode: geometry comes from the template page, copy comes
      // from `fields`. The caller cannot specify a box, so a slide built here
      // cannot drift from the page it claims to be. ───────────────────────────
      const spec = templateSpecOf(entry)
      const templateSlideIndex = Number.isInteger(args.templateSlide) ? Number(args.templateSlide) : null
      /** @type {Record<string, any>} */
      let slideSource = raw
      /** @type {string[]} */
      let templateIssues = []
      if (spec && (templateSlideIndex !== null || args.fields)) {
        if (templateSlideIndex === null) {
          return fail('template-slide-required', '模板模式下传了 fields 就要同时传 templateSlide（用模板的哪一页做底版）。')
        }
        const built = buildSlideFromTemplate(spec, templateSlideIndex, normalizeFill(args.fields), {
          slideId: slides[index]?.id ?? `s${oneBased}`,
          extra: Array.isArray(args.extra) ? args.extra : [],
        })
        if (built.slide.elements.length === 0) {
          const reasons = built.issues.length > 0
            ? `原因：` + built.issues.slice(0, 6).map((issue) => `
- ${issue}`).join('')
            : '先确认 pptx_studio_template 分解成功（页有底图或字段）。'
          return fail('template-slide-empty', `模板第 ${templateSlideIndex} 页没有生成任何元素，这一页不会被写入。${reasons}`)
        }
        slideSource = {
          ...raw,
          elements: built.slide.elements,
          templateSlide: built.slide.templateSlide,
        }
        if (raw.elements !== undefined) {
          log('warn', 'pptx-studio: 模板模式下忽略了手写的 slide.elements（位置由模板决定）；要额外加图元请用 extra。')
        }
        if (built.issues.length > 0) {
          log('info', `pptx-studio: 模板第 ${templateSlideIndex} 页填页提醒 ${built.issues.length} 条：${built.issues[0]}`)
        }
        templateIssues = built.issues
      }
      // Fail-closed on data errors: a malformed slide would otherwise be
      // silently repaired into a plausible-looking-but-wrong page.
      const graded = collectSlideIssues(slideSource, { w: deck.size.w, h: deck.size.h })
      if (graded.errors.length > 0) {
        return fail('invalid-data', `第 ${index + 1} 页有 ${graded.errors.length} 处数据错误，已拒绝写入（deck 未被改动）：\n- ${graded.errors.slice(0, 8).join('\n- ')}${graded.errors.length > 8 ? `\n……另有 ${graded.errors.length - 8} 条` : ''}`)
      }
      const issues = graded.warnings
      const incoming = normalizeSlide(
        { ...slideSource, id: slides[index]?.id ?? `s${oneBased}` },
        { index, slideW: deck.size.w, slideH: deck.size.h },
      )
      const total = slides[index] ? slides.length : slides.length + 1
      // Tell the canvas a page is being prepared before the writes land, then
      // give the pencil beat a moment. This is what turns "the deck was already
      // there" into "the deck is being drawn".
      emit(entry, {
        type: 'status',
        phase: 'drawing',
        op: 'slide',
        begin: true,
        slideIndex: index + 1,
        total,
        elements: incoming.elements.length,
        message: `正在准备第 ${index + 1} 页`,
      })
      await sleep(DRAW_BEAT_MS)
      if (slides[index]) slides[index] = incoming
      else slides.push(incoming)
      emit(entry, { type: 'slide', slide: incoming, index, total: slides.length, phase: 'content' })
      emit(entry, { type: 'status', phase: 'drawing', slideIndex: index + 1, total: slides.length, message: `已绘制第 ${index + 1} 页，共 ${slides.length} 页` })
      toolLog('slide', entry, `index=${index + 1} elements=${incoming.elements.length}`)
      /** @type {Record<string, unknown>} */
      const receipt = {
        ok: true,
        deckId: deck.id,
        index: index + 1,
        slideId: incoming.id,
        elements: incoming.elements.length,
        total: slides.length,
      }
      // Only present when non-empty: the registry's snapshot rule rejects `undefined`.
      const notes = [...issues, ...templateIssues]
      if (notes.length > 0) receipt.issues = notes
      if (incoming.templateSlide !== undefined) receipt.templateSlide = incoming.templateSlide
      return receipt
    },
  }

  /**
   * `pptx_studio_add_element`: append one element (fine-grained drawing).
   */
  const addElementTool = {
    name: 'pptx_studio_add_element',
    output: { schema: OUTPUT_SCHEMA, render: renderReceipt },
    description:
      '往某一页追加一个元素（比整页写入更细，用户能看到元素逐个画出来）。element 用 {type, box, text?, style?, fill?, src?, chart?, anim?, shadow?}。',
    parameters: {
      deck: { type: 'string', required: true, description: 'deckId。' },
      slideId: { type: 'string', required: true, description: '页 id（pptx_studio_inspect 里能看到）。' },
      element: ELEMENT_SCHEMA,
      force: { type: 'boolean', description: '跳过「调研 → HTML 预览」流程（不建议）。' },
    },
    /**
     * @param {Record<string, any>} args - tool args.
     * @returns {Promise<Record<string, unknown>>} receipt.
     */
    async execute(args, exec) {
      const found = requireDeck(args, exec)
      if (/** @type {any} */ (found).error) return /** @type {any} */ (found).error
      const entry = /** @type {any} */ (found).entry
      const gate = stageGate(entry, args.force === true)
      if (gate) return gate
      const deck = entry.model
      const slide = /** @type {any[]} */ (deck.slides).find((s) => s.id === args.slideId)
      if (!slide) return fail('slide-not-found', `没有这一页：${args.slideId}；先用 pptx_studio_inspect 看现有 slideId。`)
      const graded = collectElementIssues(args.element, slide.elements.length, { w: deck.size.w, h: deck.size.h })
      if (graded.errors.length > 0) {
        return fail('invalid-data', `新元素有 ${graded.errors.length} 处数据错误，已拒绝追加（deck 未被改动）：\n- ${graded.errors.slice(0, 8).join('\n- ')}${graded.errors.length > 8 ? `\n……另有 ${graded.errors.length - 8} 条` : ''}`)
      }
      const issues = graded.warnings
      const el = normalizeElement(args.element, {
        index: slide.elements.length,
        slideW: deck.size.w,
        slideH: deck.size.h,
        prefix: `${slide.id}-`,
      })
      const slideIndex = /** @type {any[]} */ (deck.slides).findIndex((item) => item.id === slide.id)
      // One element at a time is also a drawing beat: the canvas pencils the
      // element in, so fine-grained builds are visibly incremental too.
      emit(entry, {
        type: 'status',
        phase: 'drawing',
        op: 'element',
        begin: true,
        slideIndex: slideIndex + 1,
        total: /** @type {any[]} */ (deck.slides).length,
        elements: slide.elements.length + 1,
        message: `正在往第 ${slideIndex + 1} 页加元素`,
      })
      await sleep(Math.round(DRAW_BEAT_MS / 2))
      slide.elements.push(el)
      emit(entry, { type: 'element', slideId: slide.id, element: el, index: slide.elements.length - 1, total: slide.elements.length })
      /** @type {Record<string, unknown>} */
      const receipt = {
        ok: true,
        deckId: deck.id,
        slideId: slide.id,
        elementId: el.id,
        elements: slide.elements.length,
      }
      if (issues.length > 0) receipt.issues = issues
      return receipt
    },
  }

  /**
   * `pptx_studio_patch`: change one element (the correction loop).
   */
  const patchTool = {
    name: 'pptx_studio_patch',
    output: { schema: OUTPUT_SCHEMA, render: renderReceipt },
    description:
      '改一个元素（局部修正）：只传要改的字段，例如 {text:"新标题"}、{box:[x,y,w,h]}、{style:{size:24}}。也可以 remove:true 删除它。用户指着画布提意见时，优先用这个工具。',
    parameters: {
      deck: { type: 'string', required: true, description: 'deckId。' },
      slideId: { type: 'string', required: true, description: '页 id。' },
      elementId: { type: 'string', required: true, description: '元素 id。' },
      patch: { type: 'object', additionalProperties: true, description: '要合并进元素的字段，例如 {text:"新标题"}、{box:[x,y,w,h]}、{style:{size:24}}。' },
      remove: { type: 'boolean', description: '为 true 时删除这个元素。' },
    },
    /**
     * @param {Record<string, any>} args - tool args.
     * @returns {Promise<Record<string, unknown>>} receipt.
     */
    async execute(args, exec) {
      const found = requireDeck(args, exec)
      if (/** @type {any} */ (found).error) return /** @type {any} */ (found).error
      const entry = /** @type {any} */ (found).entry
      const deck = entry.model
      const change = {
        slideId: String(args.slideId ?? ''),
        elementId: String(args.elementId ?? ''),
        patch: args.patch !== null && typeof args.patch === 'object' ? args.patch : {},
        remove: args.remove === true,
      }
      // Pre-audit the merged element: patch merges one level deep (chart is
      // replaced wholesale), so a mistyped chart field would land silently.
      if (!change.remove) {
        const slide = /** @type {any[]} */ (deck.slides).find((item) => item.id === change.slideId)
        const at = slide ? slide.elements.findIndex((item) => item.id === change.elementId) : -1
        if (slide && at >= 0) {
          const merged = mergeElement(slide.elements[at], change.patch)
          const graded = collectElementIssues(merged, at, { w: deck.size.w, h: deck.size.h })
          if (graded.errors.length > 0) {
            return fail('invalid-data', `这次 patch 有 ${graded.errors.length} 处数据错误，已拒绝（deck 未被改动）：\n- ${graded.errors.slice(0, 8).join('\n- ')}${graded.errors.length > 8 ? `\n……另有 ${graded.errors.length - 8} 条` : ''}`)
          }
        }
      }
      const result = applyElementPatch(deck, change)
      if (!result.ok) return fail('patch-failed', result.error || 'patch failed')
      emit(entry, {
        type: 'element',
        slideId: change.slideId,
        element: result.element ?? null,
        elementId: change.elementId,
        removed: result.removed === true,
      })
      toolLog('patch', entry, `el=${change.elementId} removed=${result.removed === true}`)
      /** @type {Record<string, unknown>} */
      const patchReceipt = { ok: true, deckId: deck.id, slideId: change.slideId, elementId: change.elementId, removed: result.removed === true, element: result.element ?? null }
      // Name what the merge repaired, so the caller can see it happened.
      const after = result.element
      if (after && !change.remove) {
        const slide = /** @type {any[]} */ (deck.slides).find((item) => item.id === change.slideId)
        const at = slide ? slide.elements.findIndex((item) => item.id === change.elementId) : -1
        const graded = collectElementIssues(after, at >= 0 ? at : 0, { w: deck.size.w, h: deck.size.h })
        if (graded.warnings.length > 0) patchReceipt.issues = graded.warnings
      }
      return patchReceipt
    },
  }

  /**
   * `pptx_studio_inspect`: read the deck back for self-checks / continuation.
   */
  const inspectTool = {
    name: 'pptx_studio_inspect',
    output: { schema: OUTPUT_SCHEMA, render: renderReceipt },
    description:
      '读回当前画布的紧凑模型（页 id、每页元素 id/类型/box/前 120 字文本）。改稿前、续写时先调用它，避免瞎猜 id。',
    parameters: {
      deck: { type: 'string', required: true, description: 'deckId。' },
      slideId: { type: 'string', description: '只看某一页时传页 id。' },
    },
    /**
     * @param {Record<string, any>} args - tool args.
     * @returns {Promise<Record<string, unknown>>} receipt.
     */
    async execute(args, exec) {
      const found = requireDeck(args, exec)
      if (/** @type {any} */ (found).error) return /** @type {any} */ (found).error
      const entry = /** @type {any} */ (found).entry
      return { ok: true, deck: summarizeDeck(entry.model, typeof args.slideId === 'string' ? args.slideId : null) }
    },
  }

  /**
   * Export the model to a real .pptx.
   * @param {import('./lib/store.js').DeckEntry} entry - deck entry.
   * @param {string|undefined} target - optional output path (workspace relative or absolute inside workspace).
   * @param {string|null} templatePath - template to inherit master/layouts.
   * @param {string|undefined} directory - optional absolute folder the user picked in the Host's own directory picker; when set it wins over `target` and the file is named after the deck (never overwriting).
   * @returns {Promise<Record<string, unknown>>} export receipt.
   */
  async function exportDeck(entry, target, templatePath, directory) {
    // Gate the export on a clean deck: the .pptx is the deliverable, so it is
    // the last place a data mistake is allowed to slip through.
    const audit = auditDeck(entry.model)
    if (audit.errors.length > 0) {
      return fail('invalid-data', `deck 有 ${audit.errors.length} 处数据错误，拒绝导出（修好再导，文件未被写出）：\n- ${audit.errors.slice(0, 8).join('\n- ')}${audit.errors.length > 8 ? `\n……另有 ${audit.errors.length - 8} 条` : ''}`)
    }
    // Template mode: the .pptx is supposed to look like the template and say
    // what the HTML sheet says. Name the drift instead of shipping a file that
    // quietly disagrees with what the user approved on screen.
    const specForExport = templateSpecOf(entry)
    const consistency = specForExport ? previewConsistency(entry.model, /** @type {any} */ (entry.model.pipeline?.preview)) : []
    if (consistency.length > 0) {
      return fail('template-inconsistent', `幻灯片与 HTML 模版/模板有 ${consistency.length} 处不一致，拒绝导出（文件未被写出）。先 pptx_studio_verify 看全部问题：` + consistency.slice(0, 6).map((issue) => `
- ${issue}`).join(''))
    }
    entry.persist()
    const deckFile = join(entry.dir, 'deck.json')
    let out = join(entry.dir, 'deck.pptx')
    if (typeof directory === 'string' && directory.trim() !== '') {
      // A folder the user chose in the Host picker: this is the one export path
      // that deliberately leaves the workspace (and the store) behind.
      out = uniquePptxTarget(directory.trim(), entry.model.title)
    } else if (typeof target === 'string' && target.trim() !== '') {
      out = workspacePath(entry.cwd, target, '')
      if (extname(out).toLowerCase() !== '.pptx') out = `${out}.pptx`
    }
    const args = [deckFile, out]
    if (templatePath && existsSync(templatePath)) args.push('--template', templatePath)
    const info = await python.run('pptx_build.py', args, { timeoutMs: 180_000 })
    if (!entry.model.pipeline) entry.model.pipeline = normalizePipeline(null)
    const pipeline = /** @type {any} */ (entry.model.pipeline)
    pipeline.builtAt = new Date().toISOString()
    syncPipeline(entry)
    const event = emit(entry, {
      type: 'saved',
      path: info.path,
      bytes: info.bytes,
      slides: info.slides,
      shapes: info.shapes,
      features: info.features ?? {},
    })
    // Anything the python side failed to render, plus anything the audit
    // merely warned about, is named here instead of vanishing into stderr.
    const warnings = [...audit.warnings, ...(Array.isArray(info.warnings) ? info.warnings : [])]
    if (warnings.length > 0) log('warn', `pptx-studio: 导出带 ${warnings.length} 条警告：${warnings[0]}`)
    log('info', `pptx-studio: 导出 ${info.path}（${info.slides} 页 / ${info.bytes} 字节）`)
    /** @type {Record<string, unknown>} */
    const receipt = {
      ok: true,
      deckId: entry.model.id,
      path: info.path,
      bytes: info.bytes,
      slides: info.slides,
      shapes: info.shapes,
      features: info.features ?? {},
      pipeline: pipelineSummary(entry.model),
      seq: event.seq,
      hint: warnings.length > 0
        ? '已导出，但有警告（见 warnings）：画布与此文件同源。'
        : '用 present 或让用户在文件卡片里打开这个 .pptx；画布与此文件同源。',
    }
    if (warnings.length > 0) receipt.warnings = warnings.slice(0, 20)
    return receipt
  }

  /**
   * `pptx_studio_verify`: does the deck still match its HTML sheet and template?
   */
  const verifyTool = {
    name: 'pptx_studio_verify',
    output: { schema: OUTPUT_SCHEMA, render: renderReceipt },
    description:
      '一致性校验：检查每一页是否还贴在它声称的模板页上（底图、字段位置），以及幻灯片文字是否与「② HTML 模版」里写的完全一致。模板模式下导出前建议先跑它；返回 ok:false 时按 issues 逐条修（改 preview 或重写对应页），不要直接导出。',
    parameters: {
      deck: { type: 'string', required: true, description: 'deckId。' },
    },
    /**
     * @param {Record<string, any>} args - tool args.
     * @param {any} exec - execution context.
     * @returns {Promise<Record<string, unknown>>} receipt.
     */
    async execute(args, exec) {
      const found = requireDeck(args, exec)
      if (/** @type {any} */ (found).error) return /** @type {any} */ (found).error
      const entry = /** @type {any} */ (found).entry
      const deck = entry.model
      const spec = templateSpecOf(entry)
      const slides = /** @type {any[]} */ (deck.slides)
      /** @type {string[]} */
      const issues = []
      /** @type {string[]} */
      const warnings = []
      if (slides.length === 0) {
        return fail('nothing-to-verify', '还没有任何幻灯片可以校验：先 pptx_studio_plan + pptx_studio_slide。')
      }
      if (spec) {
        issues.push(...previewConsistency(deck, /** @type {any} */ (deck.pipeline?.preview)))
        for (const [index, slide] of slides.entries()) {
          if (!Number.isInteger(slide.templateSlide)) {
            issues.push(`第 ${index + 1} 页没有 templateSlide：这一页不是从模板建的，和模板长相会不一致（模板模式下每页都应用 pptx_studio_slide 传 templateSlide 建）`)
            continue
          }
          const page = templatePage(spec, slide.templateSlide)
          if (!page) {
            issues.push(`第 ${index + 1} 页引用模板第 ${slide.templateSlide} 页，但模板里没有这一页`)
            continue
          }
          const plate = slide.elements.find((/** @type {any} */ el) => el.field === '__plate')
          const renderAvailable = spec.render?.ok === true
          if (!plate) {
            // Without a PowerPoint render there is no plate to inherit; that is
            // a degradation the receipt already reported, not a mistake here.
            const message = `第 ${index + 1} 页没有模板底图元素（field=__plate）——导出后这一页会缺少背景与装饰`
            if (renderAvailable) issues.push(message)
            else warnings.push(`第 ${index + 1} 页用的是纯色底：模板没有渲染底图（没装 PowerPoint），成品背景是近似`)
          } else if (page.plate && plate.src !== page.plate) {
            issues.push(`第 ${index + 1} 页底图用的是 ${plate.src}，模板第 ${slide.templateSlide} 页的底图是 ${page.plate}`)
          }
          // Field boxes must come from the template, not from a hand edit.
          const byField = new Map(page.fields.map((/** @type {any} */ field) => [field.id, field]))
          for (const element of slide.elements) {
            if (typeof element.field !== 'string' || element.field === '__plate') continue
            const field = byField.get(element.field)
            if (!field) {
              issues.push(`第 ${index + 1} 页有字段 ${element.field}，模板第 ${slide.templateSlide} 页没有这个字段`)
              continue
            }
            const drift = Math.max(...element.box.map((v, i) => Math.abs(Number(v) - Number(field.box[i]))))
            if (drift > 0.06) {
              warnings.push(`第 ${index + 1} 页字段 ${element.field} 的位置被改过（偏离模板 ${Math.round(drift * 100) / 100}in）`)
            }
            if (element.style?.size && Math.abs(element.style.size - field.style.size) > 1) {
              warnings.push(`第 ${index + 1} 页字段 ${element.field} 字号 ${element.style.size}pt，模板是 ${field.style.size}pt`)
            }
          }
          const emptyRequired = page.fields.filter((/** @type {any} */ field) => field.sample.trim() !== ''
            && !slide.elements.some((/** @type {any} */ el) => el.field === field.id))
          if (emptyRequired.length > 0) {
            warnings.push(`第 ${index + 1} 页这些字段没填（会留空）：${emptyRequired.map((/** @type {any} */ f) => f.id).join('/')}`)
          }
        }
      } else if (deck.pipeline?.preview?.mode === 'template') {
        issues.push('deck 记录了模板模式的预览，但读不到模板 spec（template/template.json 丢失）；重新跑 pptx_studio_template。')
      }
      const consistent = issues.length === 0
      toolLog('verify', entry, `consistent=${consistent} issues=${issues.length}`)
      /** @type {Record<string, unknown>} */
      const receipt = {
        ok: consistent,
        deckId: deck.id,
        consistent,
        slides: slides.length,
        templateMode: Boolean(spec),
        issues: issues.slice(0, 30),
        warnings: warnings.slice(0, 20),
        hint: consistent
          ? '一致性没问题，可以 pptx_studio_export。'
          : '按 issues 修：文字不一致就先 pptx_studio_preview 更新映射再重写那页；位置/底图偏了就 pptx_studio_slide 用 templateSlide 重建该页。',
      }
      if (consistent) delete receipt.issues
      if (warnings.length === 0) delete receipt.warnings
      return receipt
    },
  }

  /**
   * `pptx_studio_export`: write the .pptx now.
   */
  const exportTool = {
    name: 'pptx_studio_export',
    output: { schema: OUTPUT_SCHEMA, render: renderReceipt },
    description: '把当前画布导出成真实的 .pptx（python-pptx 生成，原生可编辑，含转场/动画/备注/背景/图表）。返回文件绝对路径与页数/字节数/特效统计。',
    parameters: {
      deck: { type: 'string', required: true, description: 'deckId。' },
      path: { type: 'string', description: '可选输出路径（工作区内），默认 .pptx-studio/<deckId>/deck.pptx。' },
      force: { type: 'boolean', description: '跳过「调研 → HTML 预览」流程（不建议）。' },
    },
    /**
     * @param {Record<string, any>} args - tool args.
     * @returns {Promise<Record<string, unknown>>} receipt.
     */
    async execute(args, exec) {
      const found = requireDeck(args, exec)
      if (/** @type {any} */ (found).error) return /** @type {any} */ (found).error
      const entry = /** @type {any} */ (found).entry
      const gate = stageGate(entry, args.force === true)
      if (gate) return gate
      if (!Array.isArray(entry.model.slides) || entry.model.slides.length === 0) {
        return fail('empty-deck', '画布还没有页。先用 pptx_studio_plan 给大纲、pptx_studio_slide 写内容，再导出。')
      }
      const templatePath = entry.model.template?.path ?? null
      try {
        return await exportDeck(entry, args.path, templatePath)
      } catch (error) {
        return fail('export-failed', `导出失败：${error instanceof Error ? error.message : String(error)}。先修好再重试，画布内容没有丢。`)
      }
    },
  }

  /**
   * Register every tool exactly once, disposed with the plugin.
   */
  const defineTool = await loadDefineTool(ctx, log)
  const toolDefs = [openTool, templateTool, researchTool, previewTool, deckTool, planTool, slideTool, addElementTool, patchTool, inspectTool, verifyTool, exportTool]
  for (const tool of toolDefs) {
    // Every call is tracked: the canvas subscribes to the feed so it can switch
    // to the deck a running call writes (and say which step is running).
    const defined = defineTool(trackedTool(tool))
    ctx.effect(() => ctx.tools.register(defined), `pptx-studio: ${tool.name}`)
  }

  // ── bundled skill ──────────────────────────────────────────────────────────
  // The skill *is* the constraint, not a reminder: it tells the model that a
  // .pptx is produced through these tools and that a skill teaching python-pptx
  // loses to it. Registered through `ctx.inject`, so disabling the plugin
  // withdraws the rule with it — no copy is left in the user's home.
  const skillState = config.skill === false
    ? { registered: false, skills: [] }
    : registerBundledSkills(ctx, { log })

  // ── web routes ─────────────────────────────────────────────────────────────

  /**
   * Only the local, authenticated page may call these routes.
   * @param {import('node:http').IncomingMessage} req - request.
   * @returns {boolean} whether the caller is admitted.
   */
  function admits(req) {
    try {
      return ctx.connection.requestRejection(req) === undefined
    } catch {
      return false
    }
  }

  /**
   * Refuse a save folder before the export tries to write into it.
   * @param {string} directory - absolute folder the caller asked for.
   * @returns {string|null} the reason it cannot be used, or null when it can.
   */
  function directoryRefusal(directory) {
    if (!isAbsolute(directory)) return `保存目录必须是绝对路径：${directory}`
    try {
      if (!statSync(directory).isDirectory()) return `保存位置不是目录：${directory}`
    } catch (error) {
      return `保存目录不存在或无法访问：${directory}（${error instanceof Error ? error.message : String(error)}）`
    }
    return null
  }

  /**
   * Compose the user-facing feedback message for one canvas note.
   * @param {import('./lib/store.js').DeckEntry} entry - deck entry.
   * @param {Record<string, any>} body - request body.
   * @returns {{text: string, meta: Record<string, unknown>}} message text + metadata.
   */
  function composeFeedback(entry, body) {
    const deck = entry.model
    const slides = /** @type {any[]} */ (deck.slides)
    const slideId = typeof body.slideId === 'string' ? body.slideId : ''
    const slideIndex = Math.max(0, slides.findIndex((s) => s.id === slideId))
    const slide = slides[slideIndex]
    const elementId = typeof body.elementId === 'string' ? body.elementId : ''
    const element = slide && elementId ? slide.elements.find((el) => el.id === elementId) : null
    let box = Array.isArray(body.bbox) ? body.bbox.map((n) => Math.round(Number(n) * 1000) / 1000) : null
    if (!box && element) box = element.box
    const note = String(body.note ?? '').trim().slice(0, 1000)
    const action = String(body.action ?? '').trim().slice(0, 300)
    const stage = ['research', 'preview', 'build'].includes(String(body.stage))
      ? String(body.stage)
      : (/** @type {any} */ (deck.pipeline)?.stage || 'build')
    const onSlide = stage === 'build' && slides.length > 0
    const lines = ['[PPTX 画布反馈]']
    lines.push(onSlide
      ? `deck=${deck.id} slide=${slideIndex + 1}/${slides.length}${slideId ? ` (${slideId})` : ''}`
      : `deck=${deck.id} slide=-`)
    lines.push(`阶段=${stage === 'research' ? '调研分析' : stage === 'preview' ? 'HTML 效果预览' : '幻灯片'}`)
    if (stage === 'research') {
      lines.push('用户正在看调研/方案页：先判断意见属于分析结论、叙事大纲还是设计方向，必要时重新调研并更新 pptx_studio_research。')
    } else if (stage === 'preview') {
      lines.push('用户正在看 HTML 效果预览：先改 pptx_studio_preview 的视觉稿，再把同一套设计落到幻灯片（pptx_studio_slide），两边不要脱节。')
    }
    if (onSlide && elementId) {
      const current = element
        ? `type=${element.type}${typeof element.text === 'string' ? ` text=「${element.text.slice(0, 80)}」` : ''}`
        : '(元素已不在模型里)'
      lines.push(`element=${elementId} ${current}`)
    }
    if (onSlide && box) lines.push(`位置：x=${box[0]}in y=${box[1]}in w=${box[2]}in h=${box[3]}in`)
    if (onSlide && element && element.style) lines.push(`样式：${JSON.stringify(element.style).slice(0, 200)}`)
    if (note) lines.push(`用户意见：${note}`)
    if (action) lines.push(`期望动作：${action}`)
    lines.push('请只改这一处（除非同一问题在别页重复），改完调用 pptx_studio_export 重新导出；不要重做整个 deck。')
    return {
      text: lines.join('\n'),
      meta: { deckId: deck.id, slideId, elementId, bbox: box, slideIndex: slideIndex + 1, total: slides.length, stage },
    }
  }

  /**
   * Inject one user message into the owning session.
   * @param {import('./lib/store.js').DeckEntry} entry - deck entry.
   * @param {string} sessionId - target session id.
   * @param {string} text - message text.
   * @returns {Promise<{ok: boolean, mode?: string, reason?: string}>} injection result.
   */
  async function injectUserMessage(entry, sessionId, text) {
    const target = sessionId || entry.sessionId
    if (!target) return { ok: false, reason: 'session-unknown' }
    const agents = ctx.get('agents')
    const agent = agents?.get?.(target)
    if (!agent) return { ok: false, reason: 'agent-not-live' }
    let createUserMessage
    try {
      const llm = await ctx.loader.import('@deepseek-ai/dsh-llm')
      createUserMessage = llm?.createUserMessage
    } catch (error) {
      return { ok: false, reason: `llm-import-failed: ${error instanceof Error ? error.message : String(error)}` }
    }
    if (typeof createUserMessage !== 'function') return { ok: false, reason: 'api-missing' }
    const message = createUserMessage({
      content: [{ type: 'text', text }],
      source: { kind: 'plugin:dsh-pptx-studio', form: 'canvas-feedback' },
    })
    try {
      if (agent.status === 'running' && typeof agent.steer === 'function') {
        agent.steer(message)
        return { ok: true, mode: 'steer' }
      }
      if (typeof agent.followup === 'function') {
        agent.followup(message)
        return { ok: true, mode: 'followup' }
      }
      return { ok: false, reason: 'no-message-api' }
    } catch (error) {
      return { ok: false, reason: error instanceof Error ? error.message : String(error) }
    }
  }

  /**
   * Serve the SSE stream for one deck.
   * @param {import('node:http').IncomingMessage} req - request.
   * @param {import('node:http').ServerResponse} res - response.
   * @param {import('./lib/store.js').DeckEntry} entry - deck entry.
   * @returns {void}
   */
  function serveEvents(req, res, entry) {
    res.writeHead(200, {
      'content-type': 'text/event-stream; charset=utf-8',
      'cache-control': 'no-cache, no-transform',
      connection: 'keep-alive',
      'x-accel-buffering': 'no',
    })
    const send = (event) => {
      try {
        res.write(`data: ${JSON.stringify(event)}\n\n`)
      } catch {
        /* client vanished; close handler cleans up */
      }
    }
    res.write('retry: 2000\n\n')
    send({ type: 'deck', deck: entry.model, seq: entry.seq, mode: 'snapshot' })
    const unsubscribe = entry.subscribe(send)
    const heartbeat = setInterval(() => {
      try {
        res.write(`: ping ${Date.now()}\n\n`)
      } catch {
        /* handled by close */
      }
    }, HEARTBEAT_MS)
    if (typeof heartbeat.unref === 'function') heartbeat.unref()
    const cleanup = () => {
      clearInterval(heartbeat)
      unsubscribe()
    }
    req.on('close', cleanup)
    req.on('error', cleanup)
    res.on('close', cleanup)
  }

  /**
   * Stream one existing file back with cache and type headers.
   * @param {import('node:http').ServerResponse} res - response.
   * @param {string} file - absolute path of an admitted file.
   * @returns {void}
   */
  function serveFile(res, file) {
    try {
      const body = readFileSync(file)
      res.writeHead(200, {
        'content-type': contentTypeOf(file),
        'cache-control': 'no-store',
        'content-length': body.length,
      })
      res.end(body)
    } catch (error) {
      writeJson(res, 500, { ok: false, code: 'asset-read-failed', message: error instanceof Error ? error.message : String(error) })
    }
  }

  /**
   * Serve one image from the deck's assets directory (`?name=` form).
   * @param {import('node:http').ServerResponse} res - response.
   * @param {import('./lib/store.js').DeckEntry} entry - deck entry.
   * @param {string} name - file name (no path segments).
   * @returns {void}
   */
  function serveAsset(res, entry, name) {
    if (!/^[A-Za-z0-9._-]{1,200}$/.test(name) || name.includes('..')) {
      writeJson(res, 400, { ok: false, code: 'bad-asset' })
      return
    }
    const file = join(entry.dir, 'assets', name)
    if (!existsSync(file)) {
      writeJson(res, 404, { ok: false, code: 'asset-not-found' })
      return
    }
    serveFile(res, file)
  }

  /**
   * Serve one file by workspace-relative or absolute path (`?path=` form).
   *
   * `assets/<file>` is resolved against the deck directory first, so the model's
   * deck-relative `src` works unchanged; anything outside `entry.cwd` and
   * `entry.dir` is refused with 403 instead of leaking file existence.
   * @param {import('node:http').ServerResponse} res - response.
   * @param {import('./lib/store.js').DeckEntry} entry - deck entry.
   * @param {string} rawPath - requested path.
   * @returns {void}
   */
  function serveAssetPath(res, entry, rawPath) {
    const value = String(rawPath ?? '').trim()
    if (value === '') {
      writeJson(res, 400, { ok: false, code: 'bad-asset' })
      return
    }
    /** @type {string[]} */
    const candidates = []
    if (isAbsolute(value)) {
      candidates.push(resolve(value))
    } else if (value.startsWith('assets/') || value.startsWith('./assets/')) {
      // The model's `src` is deck-relative: try the deck directory first so a
      // workspace `assets/` folder cannot shadow the deck's own image.
      candidates.push(resolve(entry.dir, value))
      candidates.push(resolve(entry.cwd, value))
    } else if (value.startsWith('template/') || value.startsWith('./template/')) {
      // Template plates live under `<deck>/template/…` (this deck's own copy),
      // so that prefix resolves against the deck directory first.
      candidates.push(resolve(entry.dir, value))
      candidates.push(resolve(entry.cwd, value))
    } else {
      candidates.push(resolve(entry.cwd, value))
    }
    let inside = false
    for (const candidate of candidates) {
      if (!isInside(entry.cwd, candidate) && !isInside(entry.dir, candidate)) continue
      inside = true
      if (existsSync(candidate) && statSync(candidate).isFile()) {
        serveFile(res, candidate)
        return
      }
    }
    if (!inside) {
      writeJson(res, 403, { ok: false, code: 'asset-forbidden' })
      return
    }
    writeJson(res, 404, { ok: false, code: 'asset-not-found' })
  }

  /**
   * Apply a list of canvas edits to one deck, broadcasting every change.
   *
   * Successful ops are emitted as they land; the first failing op stops the
   * batch and answers 400 with the receipts collected so far.
   * @param {import('node:http').ServerResponse} res - response.
   * @param {import('./lib/store.js').DeckEntry} entry - deck entry.
   * @param {unknown} value - request `ops` value.
   * @returns {void}
   */
  function serveEdit(res, entry, value) {
    const ops = Array.isArray(value) ? value : null
    if (ops === null) {
      writeJson(res, 400, { ok: false, code: 'bad-request', message: 'ops must be an array' })
      return
    }
    const badIndex = ops.findIndex((op) => op === null || typeof op !== 'object' || !EDIT_KINDS.has(String(op.kind)))
    if (badIndex >= 0) {
      const op = ops[badIndex]
      const kind = op !== null && typeof op === 'object' ? String(op.kind) : String(op)
      writeJson(res, 400, { ok: false, code: 'unknown-op', message: `unsupported op: ${kind}` })
      return
    }
    /** @type {Record<string, unknown>[]} */
    const results = []
    /** @type {string|undefined} */
    let deckTitle
    for (const op of ops) {
      const outcome = applyEditOp(entry, op)
      if (!outcome.ok) {
        results.push({ ok: false, kind: op.kind, code: outcome.code ?? 'edit-failed', message: outcome.message ?? 'edit failed' })
        writeJson(res, 400, { ok: false, code: 'edit-failed', message: outcome.message ?? 'edit failed', results })
        return
      }
      if (outcome.event) emit(entry, outcome.event)
      results.push(/** @type {Record<string, unknown>} */ (outcome.result))
      if (outcome.result?.kind === 'deck.title') deckTitle = String(outcome.result.title)
    }
    writeJson(res, 200, deckTitle === undefined ? { ok: true, results } : { ok: true, results, deckTitle })
  }

  /**
   * Store one pasted/uploaded image and add it to a slide.
   * @param {import('node:http').ServerResponse} res - response.
   * @param {import('./lib/store.js').DeckEntry} entry - deck entry.
   * @param {Record<string, any>} body - request body.
   * @returns {void}
   */
  function serveUpload(res, entry, body) {
    const deck = entry.model
    const slides = /** @type {any[]} */ (deck.slides)
    const wanted = typeof body.slideId === 'string' && body.slideId !== '' ? body.slideId : ''
    const slide = wanted === '' ? slides[0] : slides.find((candidate) => candidate.id === wanted)
    if (!slide) {
      writeJson(res, 404, { ok: false, code: 'slide-not-found' })
      return
    }
    const decoded = decodeImageDataUrl(body.dataUrl)
    if (!decoded.ok) {
      writeJson(res, 400, { ok: false, code: 'bad-upload', message: 'dataUrl must be a base64 image (png/jpeg/gif/webp/svg)' })
      return
    }
    if (decoded.bytes.length > MAX_UPLOAD_BYTES) {
      writeJson(res, 413, { ok: false, code: 'too-large', message: `图片超过 ${MAX_UPLOAD_BYTES} 字节` })
      return
    }
    if (slide.elements.length >= MAX_ELEMENTS_PER_SLIDE) {
      writeJson(res, 400, { ok: false, code: 'element-limit', message: `每页最多 ${MAX_ELEMENTS_PER_SLIDE} 个元素` })
      return
    }
    const hash = createHash('sha1').update(decoded.bytes).digest('hex').slice(0, 8)
    const fileName = `img_${hash}.${decoded.ext}`
    const file = join(store.assetsDir(entry), fileName)
    if (!existsSync(file)) writeFileSync(file, decoded.bytes)
    const { w: slideW, h: slideH } = /** @type {any} */ (deck.size)
    const defaultW = slideW * 0.4
    const defaultH = defaultW * 0.6
    const fallback = [(slideW - defaultW) / 2, (slideH - defaultH) / 2, defaultW, defaultH]
    const box = insideBox(Array.isArray(body.box) ? body.box : fallback, slideW, slideH)
    const src = `${ASSET_SRC_PREFIX}${fileName}`
    const name = typeof body.name === 'string' && body.name.trim() !== '' ? body.name.trim().slice(0, 200) : fileName
    const elementId = uniqueId(new Set(slide.elements.map((el) => el.id)), `${slide.id}-e${slide.elements.length + 1}`)
    const element = normalizeElement({ id: elementId, type: 'image', box, src, name }, {
      index: slide.elements.length,
      slideW,
      slideH,
      prefix: `${slide.id}-`,
    })
    slide.elements.push(element)
    const index = slide.elements.length - 1
    emit(entry, { type: 'element', slideId: slide.id, elementId: element.id, element, removed: false, index, total: slide.elements.length })
    writeJson(res, 200, { ok: true, slideId: slide.id, element, src })
  }

  /**
   * Serve the exported .pptx, exporting on demand when the default file is
   * still missing.
   * @param {import('node:http').ServerResponse} res - response.
   * @param {import('./lib/store.js').DeckEntry} entry - deck entry.
   * @returns {Promise<void>} resolved once the response is written.
   */
  async function serveDownload(res, entry) {
    const file = join(entry.dir, 'deck.pptx')
    if (!existsSync(file)) {
      let receipt = null
      try {
        receipt = await exportDeck(entry, undefined, entry.model.template?.path ?? null)
      } catch (error) {
        writeJson(res, 500, { ok: false, code: 'export-failed', message: error instanceof Error ? error.message : String(error) })
        return
      }
      // The gate refuses bad data with a receipt, not a throw: surface it.
      if (!receipt || receipt.ok !== true) {
        writeJson(res, 500, {
          ok: false,
          code: String(receipt?.code || 'export-failed'),
          message: String(receipt?.hint || '导出被数据校验拒绝'),
        })
        return
      }
    }
    if (!existsSync(file)) {
      writeJson(res, 404, { ok: false, code: 'export-missing' })
      return
    }
    try {
      const body = readFileSync(file)
      const names = downloadFilename(entry.model.title)
      res.writeHead(200, {
        'content-type': PPTX_MIME,
        'content-length': body.length,
        'content-disposition': `attachment; filename="${names.ascii}"; filename*=UTF-8''${names.encoded}`,
        'cache-control': 'no-store',
      })
      res.end(body)
    } catch (error) {
      writeJson(res, 500, { ok: false, code: 'download-failed', message: error instanceof Error ? error.message : String(error) })
    }
  }

  ctx.effect(
    () =>
      ctx.webServer.register({
        kind: 'prefix',
        path: API_PREFIX,
        handler: async (req, res) => {
          if (!admits(req)) {
            writeJson(res, 403, { ok: false, code: 'forbidden' })
            return
          }
          const url = new URL(req.url ?? '/', 'http://dsh.internal')
          if (!url.pathname.startsWith(`${API_PREFIX}/`)) {
            writeJson(res, 404, { ok: false, code: 'not-found' })
            return
          }
          const parts = url.pathname.slice(API_PREFIX.length + 1).split('/').filter(Boolean)
          const head = parts[0] ?? ''
          try {
            if (head === 'health' && req.method === 'GET') {
              writeJson(res, 200, { ok: true, name: 'dsh-pptx-studio', decks: store.decks.size, skills: skillState.skills })
              return
            }
            // The canvas opens before the deck exists. This feed tells it which
            // deck the running call writes (and which step it is), so the tab
            // switches instead of showing the previous build.
            if (head === 'session' && parts[1] === undefined && req.method === 'GET') {
              const sessionId = url.searchParams.get('sessionId') ?? ''
              writeJson(res, 200, { ok: true, sessionId, task: feedOf(sessionId).task })
              return
            }
            if (head === 'session' && parts[1] === 'events' && req.method === 'GET') {
              serveSessionEvents(req, res, url.searchParams.get('sessionId') ?? '')
              return
            }
            if (head === 'template' && req.method === 'GET') {
              const entry = entryFor(url.searchParams.get('deck') ?? '', {
                sessionId: url.searchParams.get('sessionId') ?? '',
                cwd: url.searchParams.get('cwd') ?? '',
              })
              if (!entry) {
                writeJson(res, 404, { ok: false, code: 'deck-not-found' })
                return
              }
              const spec = templateSpecOf(entry)
              if (!spec) {
                writeJson(res, 404, { ok: false, code: 'template-missing', hint: '先跑 pptx_studio_template 分解模板。' })
                return
              }
              // The canvas renders the template catalogue (roles + fields) and
              // the page plates; it never needs the full spec with styles.
              writeJson(res, 200, {
                ok: true,
                summary: templateSummary(spec),
                size: spec.size,
                design: spec.design,
                families: spec.families,
                pages: templateCatalogue(spec),
                mapping: mappingOf(entry),
              })
              return
            }
            if (head === 'decks' && req.method === 'GET') {
              const sessionId = url.searchParams.get('sessionId')
              const decks = store.list(sessionId).map((entry) => ({
                id: entry.model.id,
                title: entry.model.title,
                slides: /** @type {any[]} */ (entry.model.slides).length,
                sessionId: entry.sessionId,
                updatedAt: entry.model.meta?.updatedAt ?? null,
              }))
              // Decks that live on disk but are not in memory yet (first load
              // after a restart / plugin reload) still have to show up.
              const cwd = url.searchParams.get('cwd') || sessionCwd(sessionId ?? '')
              const seen = new Set(decks.map((item) => item.id))
              for (const item of decksOnDisk(cwd)) {
                if (seen.has(item.id)) continue
                seen.add(item.id)
                decks.push({ ...item, sessionId: sessionId ?? null, fromDisk: true })
              }
              writeJson(res, 200, { ok: true, decks })
              return
            }
            if (head === 'deck' && parts[1] === 'delete' && req.method === 'POST') {
              /** @type {Record<string, any>} */
              let body
              try {
                body = await readJsonBody(req, 64 * 1024)
              } catch (error) {
                writeJson(res, 400, { ok: false, code: 'bad-request', message: error instanceof Error ? error.message : String(error) })
                return
              }
              const deckId = typeof body.deckId === 'string' ? body.deckId.trim() : ''
              // Deck ids are generated by `makeDeckId`; anything else would be a
              // path traversal attempt dressed up as an id.
              if (!/^deck_[A-Za-z0-9_]{1,60}$/.test(deckId)) {
                writeJson(res, 400, { ok: false, code: 'bad-deck-id', hint: 'deckId 形如 deck_20261113_244a4f' })
                return
              }
              // The canvas knows only its session id; a deck the host still holds
              // in memory also knows its own workspace, which is the reliable
              // fallback when the session lookup is unavailable.
              const known = store.get(deckId)
              const cwd = typeof body.cwd === 'string' && body.cwd.trim() !== ''
                ? body.cwd.trim()
                : (url.searchParams.get('cwd') || sessionCwd(url.searchParams.get('sessionId') ?? '') || known?.cwd || '')
              if (!cwd) {
                writeJson(res, 400, { ok: false, code: 'cwd-unknown', hint: '带上 cwd 或 sessionId 才能定位工作区' })
                return
              }
              const root = store.rootOf(cwd)
              /** @type {string} */
              let dir
              try {
                dir = ensureInside(root, join(root, STORE_DIR, deckId))
              } catch {
                writeJson(res, 400, { ok: false, code: 'bad-deck-id' })
                return
              }
              if (!existsSync(join(dir, 'deck.json'))) {
                writeJson(res, 404, { ok: false, code: 'deck-not-found' })
                return
              }
              // Drop the in-memory copy first so its debounce cannot rewrite the
              // directory we are about to remove.
              const entry = store.remove(deckId)
              entry?.emit({ type: 'deck-removed', deckId })
              try {
                rmSync(dir, { recursive: true, force: true })
              } catch (error) {
                writeJson(res, 500, { ok: false, code: 'delete-failed', message: error instanceof Error ? error.message : String(error) })
                return
              }
              log('info', `pptx-studio: 已删除 deck ${deckId}`)
              writeJson(res, 200, { ok: true, deckId, removed: true })
              return
            }
            if (head === 'state' && req.method === 'GET') {
              const entry = entryFor(url.searchParams.get('deck') ?? '', {
                sessionId: url.searchParams.get('sessionId') ?? '',
                cwd: url.searchParams.get('cwd') ?? '',
              })
              if (!entry) {
                writeJson(res, 404, { ok: false, code: 'deck-not-found' })
                return
              }
              writeJson(res, 200, {
                ok: true,
                deck: entry.model,
                template: templateOf(entry),
                dir: entry.dir,
                cwd: entry.cwd,
                assets: join(entry.dir, 'assets'),
                preview: previewFile(entry),
                previewExists: existsSync(previewFile(entry)),
                sessionId: entry.sessionId,
              })
              return
            }
            if (head === 'preview' && req.method === 'GET') {
              const entry = entryFor(url.searchParams.get('deck') ?? '', {
                sessionId: url.searchParams.get('sessionId') ?? '',
                cwd: url.searchParams.get('cwd') ?? '',
              })
              if (!entry) {
                writeJson(res, 404, { ok: false, code: 'deck-not-found' })
                return
              }
              const preview = /** @type {any} */ (entry.model.pipeline)?.preview
              // Template mode stores only the page→copy mapping; the sheet is
              // generated from the template spec on demand, so it can never be
              // stale relative to the mapping that slides are built from.
              let previewHtml = preview && typeof preview.html === 'string' ? preview.html : ''
              if (preview && preview.mode === 'template') {
                const spec = templateSpecOf(entry)
                if (!spec || !Array.isArray(preview.pages)) {
                  writeJson(res, 404, { ok: false, code: 'template-missing', hint: '模板 spec 不见了，重新跑 pptx_studio_template。' })
                  return
                }
                previewHtml = templateSheet(entry, spec, preview.pages, preview.title).html
              }
              if (!preview || previewHtml.trim() === '') {
                writeJson(res, 404, { ok: false, code: 'preview-missing', hint: '先调用 pptx_studio_preview 提交 HTML 效果。' })
                return
              }
              const body = Buffer.from(previewHtml, 'utf8')
              const download = url.searchParams.get('download') === '1'
              /** @type {Record<string, string>} */
              const headers = {
                'content-type': 'text/html; charset=utf-8',
                'content-length': String(body.length),
                'cache-control': 'no-store',
                'x-content-type-options': 'nosniff',
                // The preview is model-authored markup: render it as an opaque
                // sandbox without scripts, and require inline styles. Images
                // must be inline data URLs, so a mockup cannot phone home.
                'content-security-policy': "sandbox; default-src 'none'; style-src 'unsafe-inline'; img-src data: blob:; font-src data:; media-src data:",
                'referrer-policy': 'no-referrer',
              }
              if (download) {
                const names = downloadFilename(preview.title || entry.model.title, 'html')
                headers['content-disposition'] = `attachment; filename="${names.ascii}"; filename*=UTF-8''${names.encoded}`
              }
              res.writeHead(200, headers)
              res.end(body)
              return
            }
            if (head === 'events' && req.method === 'GET') {
              const entry = entryFor(url.searchParams.get('deck') ?? '', {
                sessionId: url.searchParams.get('sessionId') ?? '',
              })
              if (!entry) {
                writeJson(res, 404, { ok: false, code: 'deck-not-found' })
                return
              }
              serveEvents(req, res, entry)
              return
            }
            if (head === 'asset' && req.method === 'GET') {
              const entry = entryFor(url.searchParams.get('deck') ?? '', {
                sessionId: url.searchParams.get('sessionId') ?? '',
              })
              if (!entry) {
                writeJson(res, 404, { ok: false, code: 'deck-not-found' })
                return
              }
              const name = url.searchParams.get('name')
              if (typeof name === 'string' && name !== '') {
                serveAsset(res, entry, name)
                return
              }
              const path = url.searchParams.get('path')
              if (typeof path === 'string' && path !== '') {
                serveAssetPath(res, entry, path)
                return
              }
              writeJson(res, 400, { ok: false, code: 'bad-asset', message: 'name or path required' })
              return
            }
            if (head === 'edit' && req.method === 'POST') {
              /** @type {Record<string, any>} */
              let body
              try {
                body = await readJsonBody(req)
              } catch (error) {
                writeJson(res, 400, { ok: false, code: 'bad-request', message: error instanceof Error ? error.message : String(error) })
                return
              }
              const entry = store.get(typeof body.deckId === 'string' ? body.deckId : '')
              if (!entry) {
                writeJson(res, 404, { ok: false, code: 'deck-not-found' })
                return
              }
              serveEdit(res, entry, body.ops)
              return
            }
            if (head === 'upload' && req.method === 'POST') {
              /** @type {Record<string, any>} */
              let body
              try {
                body = await readJsonBody(req, UPLOAD_BODY_LIMIT)
              } catch (error) {
                const message = error instanceof Error ? error.message : String(error)
                if (message.includes('too large')) {
                  writeJson(res, 413, { ok: false, code: 'too-large', message })
                  return
                }
                writeJson(res, 400, { ok: false, code: 'bad-upload', message })
                return
              }
              const entry = store.get(typeof body.deckId === 'string' ? body.deckId : '')
              if (!entry) {
                writeJson(res, 404, { ok: false, code: 'deck-not-found' })
                return
              }
              serveUpload(res, entry, body)
              return
            }
            if (head === 'feedback' && req.method === 'POST') {
              const body = await readJsonBody(req)
              const entry = store.get(typeof body.deckId === 'string' ? body.deckId : '')
              if (!entry) {
                writeJson(res, 404, { ok: false, code: 'deck-not-found' })
                return
              }
              const composed = composeFeedback(entry, body)
              const sessionId = typeof body.sessionId === 'string' ? body.sessionId : ''
              // The canvas prefers the native composer (`setDraft` + `submit`); in
              // that case it passes inject:false and this route only records the
              // note. When no composer face exists, it asks us to inject instead.
              const injected = body.inject === false
                ? { ok: false, reason: 'client-native' }
                : await injectUserMessage(entry, sessionId, composed.text)
              emit(entry, {
                type: 'feedback',
                slideId: composed.meta.slideId,
                elementId: composed.meta.elementId,
                stage: composed.meta.stage,
                note: String(body.note ?? ''),
                by: 'user',
                delivered: injected.ok === true,
                mode: injected.mode ?? injected.reason ?? null,
              })
              writeJson(res, 200, {
                ok: true,
                delivered: injected.ok === true,
                mode: injected.mode ?? null,
                reason: injected.reason ?? null,
                text: composed.text,
                meta: composed.meta,
              })
              return
            }
            if (head === 'export' && req.method === 'POST') {
              const body = await readJsonBody(req)
              const entry = store.get(typeof body.deckId === 'string' ? body.deckId : '')
              if (!entry) {
                writeJson(res, 404, { ok: false, code: 'deck-not-found' })
                return
              }
              // The canvas asks for a folder after its own directory picker, so
              // the user has already consented to this location by hand.
              const directory = typeof body.dir === 'string' ? body.dir.trim() : ''
              if (directory !== '') {
                const refusal = directoryRefusal(directory)
                if (refusal !== null) {
                  writeJson(res, 400, { ok: false, code: 'bad-directory', message: refusal })
                  return
                }
              }
              try {
                const receipt = await exportDeck(
                  entry,
                  typeof body.path === 'string' ? body.path : undefined,
                  entry.model.template?.path ?? null,
                  directory === '' ? undefined : directory,
                )
                writeJson(res, 200, {
                  ok: true,
                  value: { ...receipt, download: `${API_PREFIX}/download?deck=${encodeURIComponent(entry.model.id)}` },
                })
              } catch (error) {
                writeJson(res, 500, { ok: false, code: 'export-failed', message: error instanceof Error ? error.message : String(error) })
              }
              return
            }
            if (head === 'download' && req.method === 'GET') {
              const entry = store.get(url.searchParams.get('deck') ?? '')
              if (!entry) {
                writeJson(res, 404, { ok: false, code: 'deck-not-found' })
                return
              }
              await serveDownload(res, entry)
              return
            }
            writeJson(res, 404, { ok: false, code: 'not-found' })
          } catch (error) {
            writeJson(res, 500, { ok: false, code: 'internal', message: error instanceof Error ? error.message : String(error) })
          }
        },
      }),
    'pptx-studio: routes',
  )

  // Warm up the python probe in the background so the first export is fast.
  python.interpreter().catch(() => {})

  log('info', `pptx-studio ready (store=${STORE_DIR}, routes=${API_PREFIX}, tools=${toolDefs.length})`)
}
