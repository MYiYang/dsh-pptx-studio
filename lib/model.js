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

export const DECK_VERSION = 2
export const MAX_SLIDES = 60
export const MAX_ELEMENTS_PER_SLIDE = 200
export const MAX_TEXT = 8000
export const MAX_ELEMENTS_TOTAL = 2000
export const MAX_PREVIEW_HTML = 400_000
export const MAX_RESEARCH_TEXT = 8000
export const ELEMENT_TYPES = new Set(['text', 'rect', 'image', 'line', 'table', 'chart', 'group'])

/** Slide transition kinds the canvas can preview and python-pptx can write. */
export const TRANSITION_TYPES = new Set([
  'none', 'fade', 'push', 'wipe', 'split', 'cover', 'pull', 'cut', 'zoom', 'dissolve', 'circle', 'random',
])
/** Transition speed buckets (OOXML `spd`). */
export const TRANSITION_SPEEDS = new Set(['slow', 'medium', 'fast'])
/** Directions shared by transitions and entrance animations. */
export const DIRECTIONS = new Set(['left', 'right', 'up', 'down'])
/** Entrance effects the .pptx timing writer supports without motion paths. */
export const ANIM_EFFECTS = new Set(['none', 'appear', 'fade', 'wipe', 'zoom'])
/** Chart kinds rendered natively by python-pptx and approximated on canvas. */
export const CHART_KINDS = new Set(['column', 'bar', 'line', 'pie', 'doughnut', 'area'])
/** Common slide-size presets (inches), used by tooling and the canvas. */
export const SIZE_PRESETS = {
  '16:9': { w: 13.333, h: 7.5 },
  '4:3': { w: 10, h: 7.5 },
  '16:10': { w: 12.8, h: 8 },
  'a4-landscape': { w: 11.693, h: 8.268 },
}
/** Default deck theme. */
export const DEFAULT_THEME = {
  colors: {
    ink: '#0F172A',
    muted: '#64748B',
    accent: '#2563EB',
    surface: '#F1F5F9',
    background: '#FFFFFF',
  },
  fonts: { latin: '微软雅黑', ea: '微软雅黑', heading: '微软雅黑' },
}

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
 * Canonicalise one colour into a CSS-safe string.
 *
 * The canvas writes colours straight into inline CSS, where a bare `060B18` is
 * an invalid value that the browser silently drops (transparent fill, inherited
 * text colour). python-pptx tolerates both spellings, so the mismatch only ever
 * showed up on the canvas. Accept `RGB`, `#RGB`, `RRGGBB`, `#RRGGBB`, 8-digit
 * alpha forms, and pass anything else (`rgba(...)`, `transparent`, named
 * colours) through untouched.
 * @param {unknown} value - candidate colour.
 * @returns {string} `#RRGGBB` (or the trimmed original), empty when unusable.
 */
export function normalizeColor(value) {
  if (typeof value !== 'string') return ''
  const raw = value.trim()
  if (raw === '') return ''
  const hex = /^#?([0-9a-fA-F]{3}|[0-9a-fA-F]{4}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/.exec(raw)
  if (!hex) return raw.slice(0, 32)
  const digits = hex[1]
  const expanded = digits.length <= 4 ? [...digits].map((d) => d + d).join('') : digits
  return `#${expanded.toUpperCase()}`.slice(0, 32)
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
  if (src.underline !== undefined) out.underline = src.underline === true
  if (typeof src.color === 'string') {
    const color = normalizeColor(src.color)
    if (color !== '') out.color = color
  }
  if (typeof src.font === 'string') out.font = src.font.slice(0, 64)
  if (typeof src.align === 'string' && ['left', 'center', 'right', 'justify'].includes(src.align)) out.align = src.align
  if (typeof src.anchor === 'string' && ['top', 'middle', 'bottom'].includes(src.anchor)) out.anchor = src.anchor
  if (src.lineSpacing !== undefined) out.lineSpacing = num(src.lineSpacing, 1.15, 0.6, 3)
  if (src.spaceAfter !== undefined) out.spaceAfter = num(src.spaceAfter, 0, 0, 200)
  return out
}

/**
 * Merge a theme with the deck defaults.
 * @param {unknown} raw - candidate theme (`{colors, fonts}`).
 * @returns {{colors: Record<string, string>, fonts: Record<string, string>}} theme.
 */
export function normalizeTheme(raw) {
  const src = raw !== null && typeof raw === 'object' ? /** @type {Record<string, any>} */ (raw) : {}
  const colors = src.colors !== null && typeof src.colors === 'object' ? src.colors : {}
  const fonts = src.fonts !== null && typeof src.fonts === 'object' ? src.fonts : {}
  /** @param {string} key - colour slot. */
  const color = (key) => normalizeColor(colors[key]) || DEFAULT_THEME.colors[key]
  /** @param {string} key - font slot. */
  const font = (key) => {
    const value = fonts[key]
    return typeof value === 'string' && value.trim() !== '' ? value.trim().slice(0, 64) : DEFAULT_THEME.fonts[key]
  }
  return {
    colors: {
      ink: color('ink'),
      muted: color('muted'),
      accent: color('accent'),
      surface: color('surface'),
      background: color('background'),
    },
    fonts: { latin: font('latin'), ea: font('ea'), heading: font('heading') },
  }
}

/**
 * Normalise a slide background (v1: a solid colour).
 * @param {unknown} raw - `"#RRGGBB"` or `{color}`.
 * @returns {{color: string}|null} background or null.
 */
export function normalizeBackground(raw) {
  if (typeof raw === 'string') {
    const color = normalizeColor(raw)
    return color === '' ? null : { color }
  }
  if (raw !== null && typeof raw === 'object') {
    const color = normalizeColor(/** @type {any} */ (raw).color)
    return color === '' ? null : { color }
  }
  return null
}

/**
 * Normalise one slide transition.
 * @param {unknown} raw - candidate transition.
 * @returns {Record<string, unknown>|null} transition or null.
 */
export function normalizeTransition(raw) {
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) return null
  const src = /** @type {Record<string, any>} */ (raw)
  return {
    type: TRANSITION_TYPES.has(String(src.type)) ? String(src.type) : 'fade',
    speed: TRANSITION_SPEEDS.has(String(src.speed)) ? String(src.speed) : 'medium',
    direction: DIRECTIONS.has(String(src.direction)) ? String(src.direction) : 'left',
    advanceOnClick: src.advanceOnClick !== false,
    advanceAfterMs: Math.round(num(src.advanceAfterMs, 0, 0, 3_600_000)),
  }
}

/**
 * Normalise one entrance animation.
 * @param {unknown} raw - candidate animation.
 * @returns {Record<string, unknown>|null} animation or null.
 */
export function normalizeAnimation(raw) {
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) return null
  const src = /** @type {Record<string, any>} */ (raw)
  return {
    effect: ANIM_EFFECTS.has(String(src.effect)) ? String(src.effect) : 'fade',
    direction: DIRECTIONS.has(String(src.direction)) ? String(src.direction) : 'left',
    duration: Math.round(num(src.duration, 500, 100, 10_000)),
    delay: Math.round(num(src.delay, 0, 0, 30_000)),
    order: Math.round(num(src.order, 0, 0, 999)),
  }
}

/**
 * Normalise an element shadow.
 * @param {unknown} raw - `true` or `{blur, distance, direction, color, alpha}`.
 * @returns {Record<string, unknown>|null} shadow or null.
 */
export function normalizeShadow(raw) {
  if (raw === true) raw = {}
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) return null
  const src = /** @type {Record<string, any>} */ (raw)
  return {
    blur: num(src.blur, 4, 0, 100),
    distance: num(src.distance, 3, 0, 100),
    direction: num(src.direction, 90, 0, 360),
    color: normalizeColor(src.color) || '#000000',
    alpha: num(src.alpha, 0.35, 0, 1),
  }
}

/**
 * Normalise inline runs inside one paragraph.
 * @param {unknown} raw - candidate runs array.
 * @returns {Array<Record<string, unknown>>|null} runs or null.
 */
function normalizeRuns(raw) {
  if (!Array.isArray(raw)) return null
  const runs = raw.slice(0, 40).map((item) => {
    const src = item !== null && typeof item === 'object' ? /** @type {Record<string, any>} */ (item) : {}
    /** @type {Record<string, unknown>} */
    const run = { text: clampText(src.text, 2000) }
    if (src.size !== undefined) run.size = num(src.size, 18, 5, 200)
    if (src.bold !== undefined) run.bold = src.bold === true
    if (src.italic !== undefined) run.italic = src.italic === true
    if (src.underline !== undefined) run.underline = src.underline === true
    if (typeof src.color === 'string') {
      const color = normalizeColor(src.color)
      if (color !== '') run.color = color
    }
    if (typeof src.font === 'string') run.font = src.font.slice(0, 64)
    return run
  })
  return runs.length > 0 ? runs : null
}

/**
 * Normalise a native chart model.
 * @param {unknown} raw - candidate chart.
 * @returns {Record<string, unknown>|null} chart or null.
 */
export function normalizeChart(raw) {
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) return null
  const src = /** @type {Record<string, any>} */ (raw)
  let categories = Array.isArray(src.categories) ? src.categories.slice(0, 24).map((item) => clampText(item, 80)) : []
  const seriesRaw = Array.isArray(src.series) ? src.series : []
  const series = seriesRaw.slice(0, 6).map((item, index) => {
    const one = item !== null && typeof item === 'object' ? /** @type {Record<string, any>} */ (item) : {}
    const values = Array.isArray(one.values) ? one.values.slice(0, 24).map((value) => num(value, 0, -1e9, 1e9)) : []
    /** @type {Record<string, unknown>} */
    const out = { name: clampText(one.name || `系列 ${index + 1}`, 80), values }
    const color = normalizeColor(one.color)
    if (color !== '') out.color = color
    return out
  })
  if (series.length === 0) return null
  // Keep the category axis aligned with the longest series so the canvas and
  // python-pptx draw the same bars (python fills missing labels itself).
  const longest = series.reduce((max, item) => Math.max(max, item.values.length), 0)
  if (categories.length < longest) {
    categories = categories.concat(Array.from({ length: longest - categories.length }, (_, index) => String(categories.length + index + 1)))
  }
  return {
    kind: CHART_KINDS.has(String(src.kind)) ? String(src.kind) : 'column',
    title: clampText(src.title, 200),
    categories,
    series,
  }
}

/**
 * Normalise the research/analysis brief.
 * @param {unknown} raw - candidate brief.
 * @returns {Record<string, unknown>|null} brief or null.
 */
export function normalizeResearch(raw) {
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) return null
  const src = /** @type {Record<string, any>} */ (raw)
  const outline = Array.isArray(src.outline) ? src.outline.slice(0, 80).map((item) => {
    const one = item !== null && typeof item === 'object' ? /** @type {Record<string, any>} */ (item) : {}
    return {
      title: clampText(one.title, 200),
      intent: clampText(one.intent, 500),
      source: clampText(one.source, 300),
    }
  }).filter((item) => item.title !== '') : []
  const facts = Array.isArray(src.facts) ? src.facts.slice(0, 80).map((item) => {
    const one = item !== null && typeof item === 'object' ? /** @type {Record<string, any>} */ (item) : {}
    return { claim: clampText(one.claim, 1000), source: clampText(one.source, 300) }
  }).filter((item) => item.claim !== '') : []
  const sources = Array.isArray(src.sources) ? src.sources.slice(0, 60).map((item) => {
    const one = item !== null && typeof item === 'object' ? /** @type {Record<string, any>} */ (item) : {}
    const url = typeof one.url === 'string' && /^https?:\/\//i.test(one.url.trim()) ? one.url.trim().slice(0, 1000) : ''
    return { title: clampText(one.title, 300), url, note: clampText(one.note, 500) }
  }).filter((item) => item.title !== '' || item.url !== '') : []
  const palette = Array.isArray(src.palette)
    ? src.palette.slice(0, 8).map((color) => normalizeColor(color)).filter(Boolean)
    : []
  const fonts = src.fonts !== null && typeof src.fonts === 'object' ? src.fonts : {}
  return {
    summary: clampText(src.summary, MAX_RESEARCH_TEXT),
    audience: clampText(src.audience, 500),
    goal: clampText(src.goal, 1000),
    tone: clampText(src.tone, 500),
    constraints: clampText(src.constraints, 2000),
    outline,
    facts,
    sources,
    palette,
    fonts: {
      latin: clampText(fonts.latin, 64),
      ea: clampText(fonts.ea, 64),
      heading: clampText(fonts.heading, 64),
    },
    updatedAt: typeof src.updatedAt === 'string' && src.updatedAt !== '' ? src.updatedAt.slice(0, 40) : new Date().toISOString(),
  }
}

/**
 * Normalise one HTML design preview.
 * @param {unknown} raw - candidate `{html, title, notes}`.
 * @returns {Record<string, unknown>|null} preview or null.
 */
export function normalizePreview(raw) {
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) return null
  const src = /** @type {Record<string, any>} */ (raw)
  const html = typeof src.html === 'string' ? src.html.slice(0, MAX_PREVIEW_HTML) : ''
  if (html.trim() === '') return null
  /** @type {Record<string, unknown>} */
  const preview = {
    html,
    title: clampText(src.title, 200),
    notes: clampText(src.notes, 2000),
    updatedAt: typeof src.updatedAt === 'string' && src.updatedAt !== '' ? src.updatedAt.slice(0, 40) : new Date().toISOString(),
  }
  // Optional design-canvas pixels. The canvas uses these to scale the iframe
  // to the panel instead of clipping or stretching a fixed-size mockup.
  const width = Math.round(num(src.width, 0, 0, 3840))
  const height = Math.round(num(src.height, 0, 0, 2160))
  if (width > 0) preview.width = width
  if (height > 0) preview.height = height
  return preview
}

/**
 * Derive the workflow stage from what the pipeline holds.
 * @param {Record<string, any>|null|undefined} pipeline - pipeline model.
 * @returns {'research'|'preview'|'build'} stage.
 */
export function pipelineStage(pipeline) {
  if (!pipeline) return 'research'
  if (pipeline.preview) return 'build'
  if (pipeline.research) return 'preview'
  return pipeline.builtAt ? 'build' : 'research'
}

/**
 * Normalise the workflow pipeline. `fallbackBuiltAt` grandfathers decks that
 * already have slides but were written before the pipeline existed.
 * @param {unknown} raw - candidate pipeline.
 * @param {string|null} [fallbackBuiltAt] - mark as already built when no evidence exists.
 * @returns {Record<string, unknown>} pipeline.
 */
export function normalizePipeline(raw, fallbackBuiltAt = null) {
  const src = raw !== null && typeof raw === 'object' ? /** @type {Record<string, any>} */ (raw) : {}
  const research = normalizeResearch(src.research)
  const preview = normalizePreview(src.preview)
  const pipeline = {
    research,
    preview,
    builtAt: typeof src.builtAt === 'string' && src.builtAt !== '' ? src.builtAt.slice(0, 40) : null,
  }
  if (fallbackBuiltAt && !research && !preview) pipeline.builtAt = fallbackBuiltAt
  pipeline.stage = pipelineStage(pipeline)
  return pipeline
}

/**
 * Compact pipeline status for tool receipts and the canvas.
 * @param {Record<string, any>|null|undefined} deck - deck model.
 * @returns {Record<string, unknown>} summary.
 */
export function pipelineSummary(deck) {
  const pipeline = deck?.pipeline || normalizePipeline(null)
  const slides = Array.isArray(deck?.slides) ? deck.slides.length : 0
  const stage = pipeline.preview
    ? 'build'
    : pipeline.research
      ? 'preview'
      : (pipeline.builtAt || slides > 0 ? 'build' : 'research')
  return {
    stage,
    research: Boolean(pipeline.research),
    preview: Boolean(pipeline.preview),
    builtAt: pipeline.builtAt || null,
    outline: Array.isArray(pipeline.research?.outline) ? pipeline.research.outline.length : 0,
    sources: Array.isArray(pipeline.research?.sources) ? pipeline.research.sources.length : 0,
    previewBytes: pipeline.preview && typeof pipeline.preview.html === 'string'
      ? Buffer.byteLength(pipeline.preview.html, 'utf8')
      : 0,
  }
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
      if (para.italic !== undefined) out.italic = para.italic === true
      if (para.underline !== undefined) out.underline = para.underline === true
      if (typeof para.color === 'string') {
        const color = normalizeColor(para.color)
        if (color !== '') out.color = color
      }
      if (typeof para.align === 'string') out.align = para.align
      const runs = normalizeRuns(para.runs)
      if (runs) out.runs = runs
      return out
    })
  }
  if (src.style !== undefined) el.style = normalizeStyle(src.style)
  for (const key of ['fill', 'line', 'src', 'radius', 'lineWidth', 'name']) {
    if (src[key] !== undefined) {
      if (key === 'radius' || key === 'lineWidth') el[key] = num(src[key], 1, 0, 20)
      else if (key === 'fill' || key === 'line') {
        const color = normalizeColor(src[key])
        if (color !== '') el[key] = color
      } else el[key] = typeof src[key] === 'string' ? String(src[key]).slice(0, 512) : src[key]
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
  if (src.chart !== undefined) {
    const chart = normalizeChart(src.chart)
    if (chart) el.chart = chart
  }
  if (src.shadow !== undefined) {
    const shadow = normalizeShadow(src.shadow)
    if (shadow) el.shadow = shadow
  }
  if (src.anim !== undefined) {
    const anim = normalizeAnimation(src.anim)
    if (anim) el.anim = anim
  }
  return el
}

/**
 * Element fields the canvas and the python builder actually read.
 * Anything else is dropped by `normalizeElement`, so an unexpected key is a
 * sign the caller sent the wrong shape.
 */
export const ELEMENT_KEYS = new Set([
  'id', 'type', 'box', 'text', 'paragraphs', 'style', 'fill', 'line', 'src',
  'radius', 'lineWidth', 'name', 'table', 'chart', 'shadow', 'anim', 'elements',
])

/**
 * Describe one element's problems without rejecting it.
 *
 * `normalizeElement` has to be forgiving: it also re-reads decks from disk and
 * receives canvas patches. But the tool path should not swallow a malformed
 * payload silently — a box of five numbers or an unknown `type` would become a
 * plausible-looking-but-wrong element. These strings are returned in the tool
 * receipt so the model can fix its own arguments.
 * @param {unknown} raw - candidate element.
 * @param {number} index - 0-based position.
 * @returns {string[]} problems, in caller-facing wording (empty when clean).
 */
export function auditElement(raw, index) {
  const at = index + 1
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) {
    return [`第 ${at} 个元素不是对象（收到 ${Array.isArray(raw) ? 'array' : raw === null ? 'null' : typeof raw}）`]
  }
  const src = /** @type {Record<string, unknown>} */ (raw)
  const issues = []
  const type = String(src.type)
  const known = ELEMENT_TYPES.has(type)
  if (!known) {
    issues.push(`第 ${at} 个元素 type=${JSON.stringify(src.type ?? null)} 不受支持（可用 ${[...ELEMENT_TYPES].join('/')}），已按 text 处理`)
  }
  const box = src.box
  const numeric = Array.isArray(box) && box.length === 4 && box.every((n) => typeof n === 'number' && Number.isFinite(n))
  if (!numeric) {
    const seen = Array.isArray(box) ? `${box.length} 项` : box === undefined ? '缺失' : typeof box
    issues.push(`第 ${at} 个元素 box 必须是 4 个数字 [x, y, w, h]（收到${seen}），已回退到默认位置`)
  } else if (box[2] <= 0 || box[3] <= 0) {
    issues.push(`第 ${at} 个元素 box 的宽高必须大于 0（收到 w=${box[2]} h=${box[3]}）`)
  }
  for (const key of Object.keys(src)) {
    if (key === '') issues.push(`第 ${at} 个元素含空字段名 ""，这不是合法元素字段，已忽略`)
    else if (!ELEMENT_KEYS.has(key)) issues.push(`第 ${at} 个元素含未知字段 "${key}"，已忽略`)
  }
  if (type === 'group' && known) {
    issues.push(`第 ${at} 个元素 type=group 只会画成虚线占位框（canvas 与 pptx 都不展开内容）`)
  }
  if (type === 'chart' && known && normalizeChart(src.chart) === null) {
    issues.push(`第 ${at} 个元素 type=chart 需要 chart.series（[{name, values}, …]）；现在只会画成占位框`)
  }
  if (type === 'text' && known && typeof src.text !== 'string' && !Array.isArray(src.paragraphs)) {
    issues.push(`第 ${at} 个元素是 text 但没有 text 字符串（也可以用 paragraphs）`)
  }
  const anim = src.anim
  if (anim !== undefined && anim !== null) {
    const effect = String(/** @type {any} */ (anim).effect ?? '')
    if (effect !== '' && !ANIM_EFFECTS.has(effect)) {
      issues.push(`第 ${at} 个元素的 anim.effect=${JSON.stringify(effect)} 不受支持（可用 ${[...ANIM_EFFECTS].join('/')}）`)
    }
  }
  return issues
}

/**
 * Describe one slide payload's problems without rejecting it.
 * @param {unknown} raw - candidate slide (`{ title?, elements }`).
 * @returns {string[]} problems, capped so the receipt stays readable.
 */
export function auditSlide(raw) {
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) return ['slide 必须是对象 { title?, elements: [...] }']
  const src = /** @type {Record<string, unknown>} */ (raw)
  const issues = []
  if (src.elements === undefined) issues.push('缺少 elements 数组：这一页会是空白页')
  else if (!Array.isArray(src.elements)) issues.push('elements 必须是数组')
  else if (src.elements.length === 0) issues.push('elements 是空数组：这一页会是空白页')
  else src.elements.forEach((el, i) => { issues.push(...auditElement(el, i)) })
  if (issues.length > 12) {
    const rest = issues.length - 12
    return issues.slice(0, 12).concat(`……另有 ${rest} 条参数问题`)
  }
  return issues
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
  if (typeof src.intent === 'string') slide.intent = slice(src.intent, 300)
  if (typeof src.notes === 'string') slide.notes = clampText(src.notes, 4000)
  if (src.hidden === true) slide.hidden = true
  if (src.background !== undefined) {
    const background = normalizeBackground(src.background)
    if (background) slide.background = background
  }
  if (src.transition !== undefined) {
    const transition = normalizeTransition(src.transition)
    if (transition) slide.transition = transition
  }
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
    theme: normalizeTheme(null),
    template: options.template ? { path: options.template } : null,
    pipeline: normalizePipeline(null),
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
  deck.theme = normalizeTheme(src.theme)
  if (src.meta !== null && typeof src.meta === 'object') deck.meta = src.meta
  const slides = Array.isArray(src.slides) ? src.slides.slice(0, MAX_SLIDES) : []
  const { w, h } = /** @type {any} */ (deck.size)
  deck.slides = slides.map((slide, i) => normalizeSlide(slide, { index: i, slideW: w, slideH: h }))
  // Decks written before the research/preview pipeline existed keep working:
  // content on disk means they were already built, so do not gate their exports.
  const builtFallback = deck.slides.length > 0
    ? (deck.meta && typeof deck.meta.updatedAt === 'string' ? deck.meta.updatedAt : new Date().toISOString())
    : null
  deck.pipeline = normalizePipeline(src.pipeline, builtFallback)
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
    theme: deck.theme,
    template: deck.template ? deck.template.path : null,
    pipeline: pipelineSummary(deck),
    slideCount: slides.length,
    slides: slides
      .filter((slide) => slideId === null || slide.id === slideId)
      .map((slide) => ({
        index: slides.indexOf(slide),
        id: slide.id,
        title: slide.title || textOf(slide),
        elementCount: slide.elements.length,
        ...(slide.transition ? { transition: slide.transition } : {}),
        ...(slide.notes ? { notes: String(slide.notes).slice(0, 120) } : {}),
        ...(slide.hidden ? { hidden: true } : {}),
        ...(slide.background ? { background: slide.background } : {}),
        elements: slide.elements.map((el) => {
          /** @type {Record<string, unknown>} */
          const row = { id: el.id, type: el.type, box: el.box }
          if (typeof el.text === 'string') row.text = el.text.slice(0, 120)
          if (el.anim && el.anim.effect && el.anim.effect !== 'none') row.anim = el.anim
          if (el.shadow) row.shadow = true
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
