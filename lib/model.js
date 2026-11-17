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

/** Hex colour spellings every consumer (canvas CSS and python-pptx) accepts. */
const HEX_COLOR = /^#?([0-9a-fA-F]{3}|[0-9a-fA-F]{4}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/

/**
 * Whether a colour is written in a form that survives the export.
 *
 * `normalizeColor` deliberately passes CSS named colours and `rgba(...)`
 * through for the canvas, but python-pptx only understands hex — anything else
 * silently falls back to the theme ink. The audit uses this to tell the two
 * apart, so a "blue" can never quietly become near-black in the .pptx.
 * @param {unknown} value - candidate colour.
 * @returns {boolean} true when the value is a hex spelling.
 */
function isHexColor(value) {
  return typeof value === 'string' && HEX_COLOR.test(value.trim())
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
  if (src.size !== undefined) out.size = num(src.size, 18, 5, 400)
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
    if (src.size !== undefined) run.size = num(src.size, 18, 5, 400)
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
  /** @type {Record<string, unknown>} */
  const out = {
    kind: CHART_KINDS.has(String(src.kind)) ? String(src.kind) : 'column',
    title: clampText(src.title, 200),
    categories,
    series,
  }
  // Chart styling that the HTML design sheet can ask for and both the canvas
  // and the export must honour, so the three views never drift apart.
  if (src.axis === 'none') out.axis = 'none'
  if (src.dataLabels === true) out.dataLabels = true
  if (Array.isArray(src.gradient) && src.gradient.length >= 2) {
    const from = normalizeColor(src.gradient[0])
    const to = normalizeColor(src.gradient[1])
    if (from !== '' && to !== '') out.gradient = [from, to]
  }
  const corner = num(src.corner, 0, 0, 0.5)
  if (corner > 0) out.corner = corner
  return out
}

/** Chart keys the canvas and the python builder actually read. */
const CHART_KEYS = new Set([
  'kind', 'title', 'categories', 'series', 'axis', 'dataLabels', 'gradient', 'corner',
])

/**
 * Describe one chart payload's problems without rejecting it.
 *
 * `normalizeChart` has to stay forgiving (it re-reads decks from disk), so it
 * silently repairs a lot: missing categories get numbered, values clamp to 0,
 * an unknown `axis` falls back to the automatic ruler. Every one of those
 * silent repairs is a place where a mistyped argument turns into a
 * plausible-looking-but-wrong chart, so each is named here instead.
 * @param {unknown} raw - candidate chart.
 * @param {number} at - 1-based element position for caller-facing wording.
 * @returns {{errors: string[], warnings: string[]}} graded problems.
 */
export function auditChart(raw, at) {
  /** @type {string[]} */
  const errors = []
  /** @type {string[]} */
  const warnings = []
  const at_ = `第 ${at} 个元素`
  if (raw === null || raw === undefined || typeof raw !== 'object' || Array.isArray(raw)) {
    errors.push(`${at_} chart 必须是对象 { kind, categories, series, … }（收到 ${Array.isArray(raw) ? 'array' : raw === null ? 'null' : typeof raw}）`)
    return { errors, warnings }
  }
  const src = /** @type {Record<string, any>} */ (raw)
  for (const key of Object.keys(src)) {
    if (key === '') errors.push(`${at_} chart 含空字段名 ""，已忽略`)
    else if (!CHART_KEYS.has(key)) errors.push(`${at_} chart 含未知字段 "${key}"（可用 ${[...CHART_KEYS].join('/')}），已忽略——多半是字段名写错了`)
  }
  const kind = String(src.kind ?? 'column')
  if (!CHART_KINDS.has(kind)) {
    errors.push(`${at_} chart.kind=${JSON.stringify(src.kind ?? null)} 不受支持（可用 ${[...CHART_KINDS].join('/')}），已按 column 处理`)
  }
  if (src.title !== undefined && typeof src.title !== 'string') {
    errors.push(`${at_} chart.title 必须是字符串，已忽略`)
  }
  if (src.axis !== undefined && src.axis !== 'none' && src.axis !== 'auto') {
    errors.push(`${at_} chart.axis=${JSON.stringify(src.axis)} 不受支持（可用 auto/none），已按 auto 处理`)
  }
  if (src.dataLabels !== undefined && typeof src.dataLabels !== 'boolean') {
    errors.push(`${at_} chart.dataLabels 必须是 true/false，已忽略`)
  }
  if (src.gradient !== undefined) {
    const g = src.gradient
    if (!Array.isArray(g) || g.length < 2 || !isHexColor(g[0]) || !isHexColor(g[1])) {
      errors.push(`${at_} chart.gradient 必须是两个十六进制颜色 [起, 止]（如 ["#22D3EE", "#0E7490"]），已忽略`)
    }
  }
  if (src.corner !== undefined && (typeof src.corner !== 'number' || !Number.isFinite(src.corner) || src.corner < 0 || src.corner > 0.5)) {
    errors.push(`${at_} chart.corner 必须是 0–0.5 的数字（柱顶圆角占柱宽比例），已忽略`)
  }
  const categories = src.categories
  if (categories !== undefined && !Array.isArray(categories)) {
    errors.push(`${at_} chart.categories 必须是字符串数组，已按自动编号处理`)
  }
  const seriesRaw = src.series
  if (!Array.isArray(seriesRaw) || seriesRaw.length === 0) {
    errors.push(`${at_} chart.series 至少需要一个系列 [{ name, values }, …]；现在只会画成占位框`)
    return { errors, warnings }
  }
  if (seriesRaw.length > 6) warnings.push(`${at_} chart.series 超过 6 个，只画前 6 个`)
  let longest = 0
  seriesRaw.forEach((one, index) => {
    const label = `${at_} chart.series[${index}]`
    if (one === null || typeof one !== 'object' || Array.isArray(one)) {
      errors.push(`${label} 必须是对象 { name, values, color? }`)
      return
    }
    const item = /** @type {Record<string, any>} */ (one)
    if (!item.name || typeof item.name !== 'string') {
      warnings.push(`${label} 缺 name，已用「系列 ${index + 1}」顶上`)
    }
    if (item.color !== undefined && !isHexColor(item.color)) {
      errors.push(`${label} color=${JSON.stringify(item.color)} 不是十六进制颜色（#RRGGBB 或 #RRGGBBAA；命名色如 blue 导出时会静默变默认色），已按主题色处理`)
    }
    const values = item.values
    if (!Array.isArray(values) || values.length === 0) {
      errors.push(`${label} 缺 values 数组（至少一个数字）`)
      return
    }
    longest = Math.max(longest, values.length)
    values.forEach((value, vi) => {
      if (typeof value !== 'number' || !Number.isFinite(value)) {
        errors.push(`${label}.values[${vi}]=${JSON.stringify(value)} 不是有限数字，已按 0 处理`)
      }
    })
  })
  const catLen = Array.isArray(categories) ? categories.length : 0
  if (catLen !== longest) {
    errors.push(`${at_} chart.categories 有 ${catLen} 个，而最长的系列有 ${longest} 个值——两者必须一样长（少的会补编号、多的被截断，图表就和你想画的对不上了）`)
  }
  return { errors, warnings }
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
    /** @type {Record<string, unknown>} */
    const entry = {
      title: clampText(one.title, 200),
      intent: clampText(one.intent, 500),
      source: clampText(one.source, 300),
    }
    // Template mode: this page is built on a specific template page, and the
    // copy for that page is decided here — the same map the HTML sheet and the
    // slide are both generated from.
    if (Number.isInteger(one.templateSlide) && one.templateSlide >= 0) {
      entry.templateSlide = one.templateSlide
    }
    if (one.fields !== null && typeof one.fields === 'object' && !Array.isArray(one.fields)) {
      /** @type {Record<string, string>} */
      const fields = {}
      for (const [key, value] of Object.entries(/** @type {Record<string, any>} */ (one.fields))) {
        if (typeof key !== 'string' || key.trim() === '') continue
        if (value === null || value === undefined) continue
        if (typeof value !== 'string' && typeof value !== 'number') continue
        fields[key.trim().slice(0, 40)] = String(value).slice(0, 2000)
      }
      entry.fields = fields
    }
    return entry
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
  // Template mode keeps the page→copy mapping and generates the HTML from it,
  // so the sheet can never say something the slides do not.
  const templatePages = Array.isArray(src.pages)
    ? src.pages.slice(0, 80).map((item) => {
      const one = item !== null && typeof item === 'object' ? /** @type {Record<string, any>} */ (item) : {}
      /** @type {Record<string, string>} */
      const fields = {}
      if (one.fields !== null && typeof one.fields === 'object' && !Array.isArray(one.fields)) {
        for (const [key, value] of Object.entries(one.fields)) {
          if (typeof key !== 'string' || key.trim() === '') continue
          if (typeof value !== 'string' && typeof value !== 'number') continue
          fields[key.trim().slice(0, 40)] = String(value).slice(0, 2000)
        }
      }
      return {
        templateSlide: Number.isInteger(one.templateSlide) ? one.templateSlide : 0,
        title: clampText(one.title, 200),
        fields,
      }
    })
    : []
  const mode = src.mode === 'template' && templatePages.length > 0 ? 'template' : 'html'
  const html = typeof src.html === 'string' ? src.html.slice(0, MAX_PREVIEW_HTML) : ''
  if (mode === 'html' && html.trim() === '') return null
  /** @type {Record<string, unknown>} */
  const preview = {
    mode,
    html,
    title: clampText(src.title, 200),
    notes: clampText(src.notes, 2000),
    updatedAt: typeof src.updatedAt === 'string' && src.updatedAt !== '' ? src.updatedAt.slice(0, 40) : new Date().toISOString(),
  }
  if (mode === 'template') preview.pages = templatePages
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
  if (pipeline.builtAt) return 'build'
  // A template that has not been decomposed yet is its own first step: the
  // requirement cannot be broken down before we know what the template offers.
  if (pipeline.template && pipeline.template.required === true && !pipeline.template.decomposedAt) return 'template'
  return 'research'
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
  const tplSrc = src.template !== null && typeof src.template === 'object' ? /** @type {Record<string, any>} */ (src.template) : null
  const template = tplSrc
    ? {
      path: clampText(tplSrc.path, 1024),
      required: tplSrc.required === true,
      decomposedAt: typeof tplSrc.decomposedAt === 'string' && tplSrc.decomposedAt !== '' ? tplSrc.decomposedAt.slice(0, 40) : null,
      pages: num(tplSrc.pages, 0, 0, 500),
      engine: clampText(tplSrc.engine, 40),
      textHidden: tplSrc.textHidden === true,
      design: tplSrc.design !== null && typeof tplSrc.design === 'object' ? tplSrc.design : null,
    }
    : null
  const pipeline = {
    research,
    preview,
    template,
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
      : (pipeline.builtAt || slides > 0
        ? 'build'
        : (pipeline.template && pipeline.template.required === true && !pipeline.template.decomposedAt ? 'template' : 'research'))
  return {
    stage,
    research: Boolean(pipeline.research),
    preview: Boolean(pipeline.preview),
    template: pipeline.template
      ? {
        path: pipeline.template.path,
        decomposed: Boolean(pipeline.template.decomposedAt),
        pages: pipeline.template.pages,
        engine: pipeline.template.engine,
        textHidden: pipeline.template.textHidden,
      }
      : null,
    previewMode: pipeline.preview?.mode === 'template' ? 'template' : 'html',
    previewPages: Array.isArray(pipeline.preview?.pages) ? pipeline.preview.pages.length : 0,
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
      if (para.size !== undefined) out.size = num(para.size, 18, 5, 400)
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
  // Which template field this text element fills (template mode provenance).
  if (typeof src.field === 'string' && src.field !== '') el.field = src.field.slice(0, 40)
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
  'field',
])

/**
 * Describe one element's problems, graded into errors (would produce a wrong
 * .pptx — the tool path refuses these) and warnings (repaired, but the caller
 * should know).
 *
 * `normalizeElement` has to be forgiving: it also re-reads decks from disk and
 * receives canvas patches. But the tool path must not swallow a malformed
 * payload silently — a box of five numbers, a colour that is not a colour or a
 * chart whose categories disagree with its values would become a
 * plausible-looking-but-wrong element.
 * @param {unknown} raw - candidate element.
 * @param {number} index - 0-based position.
 * @param {{w: number, h: number}|null} [size] - canvas size in inches, when known.
 * @returns {{errors: string[], warnings: string[]}} graded problems.
 */
export function collectElementIssues(raw, index, size = null) {
  /** @type {string[]} */
  const errors = []
  /** @type {string[]} */
  const warnings = []
  const at = index + 1
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) {
    errors.push(`第 ${at} 个元素不是对象（收到 ${Array.isArray(raw) ? 'array' : raw === null ? 'null' : typeof raw}）`)
    return { errors, warnings }
  }
  const src = /** @type {Record<string, any>} */ (raw)
  const type = String(src.type)
  const known = ELEMENT_TYPES.has(type)
  if (!known) {
    errors.push(`第 ${at} 个元素 type=${JSON.stringify(src.type ?? null)} 不受支持（可用 ${[...ELEMENT_TYPES].join('/')}），已按 text 处理`)
  }
  const box = src.box
  const numeric = Array.isArray(box) && box.length === 4 && box.every((n) => typeof n === 'number' && Number.isFinite(n))
  if (!numeric) {
    const seen = Array.isArray(box) ? `${box.length} 项` : box === undefined ? '缺失' : typeof box
    errors.push(`第 ${at} 个元素 box 必须是 4 个数字 [x, y, w, h]（收到${seen}），已回退到默认位置`)
  } else if (box[2] <= 0 || box[3] <= 0) {
    errors.push(`第 ${at} 个元素 box 的宽高必须大于 0（收到 w=${box[2]} h=${box[3]}）`)
  } else if (size && (box[0] < -0.01 || box[1] < -0.01 || box[0] + box[2] > size.w + 0.01 || box[1] + box[3] > size.h + 0.01)) {
    errors.push(`第 ${at} 个元素 box 超出画布：x+w=${Math.round((box[0] + box[2]) * 1000) / 1000} > ${size.w} 或 y+h=${Math.round((box[1] + box[3]) * 1000) / 1000} > ${size.h}（页就那么大，超出的部分在 pptx 里会被裁掉或跑到页外）`)
  }
  for (const key of Object.keys(src)) {
    if (key === '') warnings.push(`第 ${at} 个元素含空字段名 ""，这不是合法元素字段，已忽略`)
    else if (!ELEMENT_KEYS.has(key)) warnings.push(`第 ${at} 个元素含未知字段 "${key}"，已忽略`)
  }
  if (type === 'group' && known) {
    warnings.push(`第 ${at} 个元素 type=group 只会画成虚线占位框（canvas 与 pptx 都不展开内容）`)
  }
  const style = src.style
  if (style !== undefined && style !== null && typeof style === 'object') {
    const one = /** @type{Record<string, any>} */ (style)
    if (one.size !== undefined && (typeof one.size !== 'number' || !Number.isFinite(one.size) || one.size <= 0 || one.size > 400)) {
      errors.push(`第 ${at} 个元素 style.size=${JSON.stringify(one.size)} 不是 0–400 之间的磅值，已回退到默认字号`)
    }
    if (one.color !== undefined && !isHexColor(one.color)) {
      errors.push(`第 ${at} 个元素 style.color=${JSON.stringify(one.color)} 不是十六进制颜色（#RRGGBB 或 #RRGGBBAA；命名色如 blue 画布能认、pptx 里会静默变默认色），文字会变成看不见的默认色`)
    }
    if (one.font !== undefined && typeof one.font !== 'string') {
      warnings.push(`第 ${at} 个元素 style.font 必须是字符串，已忽略`)
    }
  }
  for (const key of ['fill', 'line']) {
    if (src[key] !== undefined && src[key] !== null && !isHexColor(src[key])) {
      errors.push(`第 ${at} 个元素 ${key}=${JSON.stringify(src[key])} 不是十六进制颜色（#RRGGBB 或 #RRGGBBAA；命名色导出时会静默按无${key === 'fill' ? '填充' : '描边'}处理）`)
    }
  }
  if (type === 'chart' && known) {
    const graded = auditChart(src.chart, at)
    errors.push(...graded.errors)
    warnings.push(...graded.warnings)
  }
  if (type === 'text' && known && typeof src.text !== 'string' && !Array.isArray(src.paragraphs)) {
    errors.push(`第 ${at} 个元素是 text 但没有 text 字符串（也可以用 paragraphs）`)
  }
  const anim = src.anim
  if (anim !== undefined && anim !== null) {
    const effect = String(/** @type{any} */ (anim).effect ?? '')
    if (effect !== '' && !ANIM_EFFECTS.has(effect)) {
      errors.push(`第 ${at} 个元素的 anim.effect=${JSON.stringify(effect)} 不受支持（可用 ${[...ANIM_EFFECTS].join('/')}），这个动画不会出现`)
    }
  }
  return { errors, warnings }
}

/**
 * Flat, ordered problem list for one element (errors first, then warnings).
 * Kept for callers that only need a list; the tool path uses the graded form.
 * @param {unknown} raw - candidate element.
 * @param {number} index - 0-based position.
 * @param {{w: number, h: number}|null} [size] - canvas size in inches, when known.
 * @returns {string[]} problems, in caller-facing wording (empty when clean).
 */
export function auditElement(raw, index, size = null) {
  const { errors, warnings } = collectElementIssues(raw, index, size)
  return [...errors, ...warnings]
}

/**
 * Describe one slide payload's problems without rejecting it.
 * @param {unknown} raw - candidate slide (`{ title?, elements }`).
 * @param {{w: number, h: number}|null} [size] - canvas size in inches, when known.
 * @returns {string[]} problems, capped so the receipt stays readable.
 */
export function auditSlide(raw, size = null) {
  const { errors, warnings } = collectSlideIssues(raw, size)
  const issues = [...errors, ...warnings]
  if (issues.length > 12) {
    const rest = issues.length - 12
    return issues.slice(0, 12).concat(`……另有 ${rest} 条参数问题`)
  }
  return issues
}

/**
 * Graded problems for one slide payload.
 * @param {unknown} raw - candidate slide.
 * @param {{w: number, h: number}|null} [size] - canvas size in inches, when known.
 * @returns {{errors: string[], warnings: string[]}} graded problems.
 */
export function collectSlideIssues(raw, size = null) {
  /** @type {string[]} */
  const errors = []
  /** @type {string[]} */
  const warnings = []
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) {
    errors.push('slide 必须是对象 { title?, elements: [...] }')
    return { errors, warnings }
  }
  const src = /** @type{Record<string, any>} */ (raw)
  if (src.elements === undefined) errors.push('缺少 elements 数组：这一页会是空白页')
  else if (!Array.isArray(src.elements)) errors.push('elements 必须是数组')
  else if (src.elements.length === 0) errors.push('elements 是空数组：这一页会是空白页')
  else if (src.elements.length > MAX_ELEMENTS_PER_SLIDE) errors.push(`elements 超过 ${MAX_ELEMENTS_PER_SLIDE} 个，只画前 ${MAX_ELEMENTS_PER_SLIDE} 个`)
  else src.elements.forEach((el, i) => {
    const graded = collectElementIssues(el, i, size)
    errors.push(...graded.errors)
    warnings.push(...graded.warnings)
  })
  if (src.transition !== undefined && src.transition !== null) {
    const one = /** @type{Record<string, any>} */ (src.transition)
    if (one.type !== undefined && !TRANSITION_TYPES.has(String(one.type))) {
      errors.push(`transition.type=${JSON.stringify(one.type)} 不受支持（可用 ${[...TRANSITION_TYPES].join('/')}），这个转场不会出现`)
    }
  }
  return { errors, warnings }
}

/**
 * Whole-deck audit, run before any write lands and before every export.
 *
 * This is the gate that makes "the data was quietly wrong" impossible: errors
 * refuse the write/export (fail-closed, with the exact field to fix), warnings
 * pass but are named in the receipt.
 * @param {unknown} deck - candidate deck model.
 * @returns {{errors: string[], warnings: string[]}} graded problems.
 */
export function auditDeck(deck) {
  /** @type {string[]} */
  const errors = []
  /** @type {string[]} */
  const warnings = []
  if (deck === null || typeof deck !== 'object' || Array.isArray(deck)) {
    return { errors: ['deck 必须是对象'], warnings }
  }
  const src = /** @type{Record<string, any>} */ (deck)
  const size = src.size && typeof src.size === 'object'
    ? { w: Number(src.size.w), h: Number(src.size.h) }
    : null
  if (!size || !Number.isFinite(size.w) || !Number.isFinite(size.h) || size.w <= 0 || size.h <= 0) {
    errors.push('deck.size 必须是 { w, h }（英寸，正数），画布与 pptx 都按它排版')
  }
  const slides = src.slides
  if (!Array.isArray(slides) || slides.length === 0) {
    errors.push('deck.slides 至少需要一页')
  } else {
    if (slides.length > MAX_SLIDES) errors.push(`deck.slides 超过 ${MAX_SLIDES} 页，只导出前 ${MAX_SLIDES} 页`)
    let total = 0
    slides.forEach((one, i) => {
      total += Array.isArray(one?.elements) ? one.elements.length : 0
      const graded = collectSlideIssues(one, size)
      errors.push(...graded.errors.map((line) => `第 ${i + 1} 页：${line}`))
      warnings.push(...graded.warnings.map((line) => `第 ${i + 1} 页：${line}`))
    })
    if (total > MAX_ELEMENTS_TOTAL) errors.push(`全 deck 元素超过 ${MAX_ELEMENTS_TOTAL} 个，只导出前 ${MAX_ELEMENTS_TOTAL} 个`)
  }
  // The design sheet is the source of truth for page count: a deck whose
  // slides drifted away from it is exactly the "HTML 与幻灯片不一致" failure.
  const preview = src.pipeline && src.pipeline.preview
  if (preview && typeof preview.html === 'string' && Array.isArray(slides)) {
    const sheetPages = (preview.html.match(/class="slide"/g) || []).length
    if (sheetPages > 0 && sheetPages !== slides.length) {
      warnings.push(`HTML 设计稿有 ${sheetPages} 页，而幻灯片有 ${slides.length} 页——两者页数不一致，导出前请对齐`)
    }
  }
  return { errors, warnings }
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
  // Provenance: which template page this slide was built on.
  if (Number.isInteger(src.templateSlide) && src.templateSlide >= 0) slide.templateSlide = src.templateSlide
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
    pipeline: normalizePipeline(options.template ? { template: { path: options.template, required: true } } : null),
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
  // A deck opened with a template always reports template mode, even when the
  // pipeline written to disk predates this field.
  if (deck.template && deck.template.path && (!deck.pipeline.template || deck.pipeline.template.path === '')) {
    deck.pipeline.template = {
      path: String(deck.template.path),
      required: true,
      decomposedAt: deck.template.decomposedAt ?? null,
      pages: num(deck.template.pages, 0, 0, 500),
      engine: '',
      textHidden: false,
      design: null,
    }
  }
  deck.pipeline.stage = pipelineStage(deck.pipeline)
  return deck
}

/**
 * Deep-ish merge used by `patch`: objects merge one level, everything else replaces.
 * Exported so the tool path can pre-audit a patched element before it lands.
 * @param {Record<string, unknown>} target - element fields.
 * @param {Record<string, unknown>} patch - incoming fields.
 * @returns {Record<string, unknown>} merged element.
 */
export function mergeElement(target, patch) {
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
        ...(Number.isInteger(slide.templateSlide) ? { templateSlide: slide.templateSlide } : {}),
        ...(slide.transition ? { transition: slide.transition } : {}),
        ...(slide.notes ? { notes: String(slide.notes).slice(0, 120) } : {}),
        ...(slide.hidden ? { hidden: true } : {}),
        ...(slide.background ? { background: slide.background } : {}),
        elements: slide.elements.map((el) => {
          /** @type {Record<string, unknown>} */
          const row = { id: el.id, type: el.type, box: el.box }
          if (typeof el.text === 'string') row.text = el.text.slice(0, 120)
          if (typeof el.field === 'string') row.field = el.field
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
