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
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, extname, isAbsolute, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

import { DeckStore, STORE_DIR, ensureInside, makeDeckId } from './lib/store.js'
import {
  MAX_SLIDES,
  applyElementPatch,
  normalizeElement,
  normalizeSlide,
  summarizeDeck,
} from './lib/model.js'
import { PythonRunner } from './lib/python.js'

export const inject = ['tools', 'webServer', 'connection']

/** Route prefix served by the Web carrier. */
const API_PREFIX = '/dsh-pptx-studio/api'
/** Bundled python helpers. */
const ASSETS_DIR = fileURLToPath(new URL('./assets', import.meta.url))
/** Default slide geometry (16:9, inches). */
const DEFAULT_SIZE = { w: 13.333, h: 7.5 }
/** SSE heartbeat. */
const HEARTBEAT_MS = 15_000

const JSON_HEADERS = {
  'content-type': 'application/json; charset=utf-8',
  'cache-control': 'no-store',
}

/**
 * Small helper: await a delay (used to pace slide-by-slide drawing).
 * @param {number} ms - milliseconds.
 * @returns {Promise<void>} resolved after the delay.
 */
function sleep(ms) {
  return new Promise((resolvePromise) => {
    const timer = setTimeout(resolvePromise, ms)
    if (typeof timer.unref === 'function') timer.unref()
  })
}

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
    req.on('data', (chunk) => {
      total += chunk.length
      if (total > limit) {
        reject(new Error('request body too large'))
        req.destroy()
        return
      }
      chunks.push(chunk)
    })
    req.on('end', () => {
      const text = Buffer.concat(chunks).toString('utf8').trim()
      if (text === '') return resolvePromise({})
      try {
        const parsed = JSON.parse(text)
        resolvePromise(parsed !== null && typeof parsed === 'object' ? parsed : {})
      } catch (error) {
        reject(new Error(`invalid JSON body: ${error instanceof Error ? error.message : String(error)}`))
      }
    })
    req.on('error', reject)
  })
}

/** Output schema shared by every pptx-studio tool (loose object). */
const OUTPUT_SCHEMA = { type: 'object', additionalProperties: true }

/**
 * Render one tool receipt back to the model: a short status line, never the
 * whole deck (the canvas already has it; echoing it would double the tokens).
 * @param {Record<string, any>} _args - tool args (unused).
 * @param {Record<string, any>} value - execute() result.
 * @returns {Array<{type: string, text: string}>} content blocks.
 */
function renderReceipt(_args, value) {
  if (!value || value.ok === false) {
    return [{ type: 'text', text: `pptx-studio 未完成：${value?.code ?? 'error'}${value?.hint ? ` — ${value.hint}` : ''}` }]
  }
  const parts = []
  // `inspect` answers with the deck model itself; summarise it for the model
  // instead of echoing the whole structure (the canvas already shows it).
  if (value.deck && typeof value.deck === 'object') {
    const deck = /** @type {any} */ (value.deck)
    if (deck.id) parts.push(`deck=${deck.id}`)
    if (typeof deck.slideCount === 'number') parts.push(`${deck.slideCount} 页`)
    if (Array.isArray(deck.slides)) {
      parts.push(deck.slides.map((slide) => `${slide.index + 1}:${slide.id}(${slide.elementCount}元素)`).join(' '))
    }
    return [{ type: 'text', text: parts.join('；') }]
  }
  if (value.deckId) parts.push(`deck=${value.deckId}`)
  if (value.mode) parts.push(value.mode === 'template' ? '模板模式' : '空白模式')
  if (value.total !== undefined && value.slides === undefined) parts.push(`共 ${value.total} 页`)
  if (value.slides !== undefined && typeof value.slides === 'number') parts.push(`${value.slides} 页`)
  if (value.index !== undefined) parts.push(`第 ${value.index} 页`)
  if (value.elements !== undefined && typeof value.elements === 'number') parts.push(`${value.elements} 个元素`)
  if (value.path) parts.push(`已导出 ${value.path}（${value.bytes ?? '?'} 字节）`)
  const text = parts.length > 0 ? parts.join('；') : 'ok'
  return [{ type: 'text', text: value.hint ? `${text}
${value.hint}` : text }]
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
   * Load a deck or answer with a model-actionable failure.
   * @param {Record<string, unknown>} args - tool args.
   * @returns {{entry: import('./lib/store.js').DeckEntry}|{error: Record<string, unknown>}} result.
   */
  function requireDeck(args) {
    const id = deckArg(args)
    if (id === '') return { error: fail('deck-required', '先调用 pptx_studio_open 创建画布，并把这个工具返回的 deckId 传进 deck 参数。') }
    const entry = store.get(id)
    if (!entry) return { error: fail('deck-not-found', `没有这个 deck：${id}。用 pptx_studio_open 重新打开。`) }
    return { entry }
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
        try {
          templatePath = workspacePath(root, args.template, '')
          if (!existsSync(templatePath)) {
            return fail('template-missing', `模板不存在：${templatePath}。把模板放到工作区目录后再试。`)
          }
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
      entry.persist()
      emit(entry, { type: 'deck', deck: entry.model, mode: templatePath ? 'template' : 'blank' })
      toolLog('open', entry, `template=${templatePath ? 'yes' : 'no'} slides=${/** @type {any[]} */ (entry.model.slides).length}`)
      return {
        ok: true,
        deckId: entry.model.id,
        dir: entry.dir,
        mode: templatePath ? 'template' : 'blank',
        size: entry.model.size,
        slides: /** @type {any[]} */ (entry.model.slides).length,
        template: info
          ? {
              path: templatePath,
              layouts: /** @type {any[]} */ (info.layouts || []).length,
              colors: Object.keys(/** @type {any} */ (info.theme?.colors || {})).length,
            }
          : null,
        hint: '右侧画布已打开。接着用 pptx_studio_plan 给出页大纲，再用 pptx_studio_slide 逐页写入。',
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
          },
        },
      },
    },
    /**
     * @param {Record<string, any>} args - tool args.
     * @param {any} exec - execution context.
     * @returns {Promise<Record<string, unknown>>} receipt.
     */
    async execute(args, exec) {
      const found = requireDeck(args)
      if (/** @type {any} */ (found).error) return /** @type {any} */ (found).error
      const entry = /** @type {any} */ (found).entry
      const planned = Array.isArray(args.slides) ? args.slides.slice(0, MAX_SLIDES) : []
      if (planned.length === 0) return fail('empty-plan', 'slides 至少要有一页 {title}。')
      const deck = entry.model
      const info = templateOf(entry)
      /** @type {any[]} */
      const slides = []
      for (const [index, item] of planned.entries()) {
        const title = String(item?.title ?? `第 ${index + 1} 页`)
        const slide = normalizeSlide(
          { id: `s${index + 1}`, title, elements: [titleElement(info, deck, title)] },
          { index, slideW: deck.size.w, slideH: deck.size.h },
        )
        if (item?.intent) slide.intent = String(item.intent).slice(0, 300)
        slides.push(slide)
        deck.slides = slides.slice()
        emit(entry, { type: 'slide', slide, index, total: slides.length, phase: 'skeleton' })
        if (index < planned.length - 1) await sleep(140)
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
      '写入/替换一页幻灯片（1-based index）。slide.elements 里每个元素用英寸坐标 box=[x,y,w,h]；type 支持 text/rect/image/line/table/chart。写入后画布会以动画绘制这一页。',
    parameters: {
      deck: { type: 'string', required: true, description: 'deckId。' },
      index: { type: 'number', required: true, description: '页码，从 1 开始；超出当前页数则追加。' },
      slide: {
        type: 'object',
        required: true,
        additionalProperties: true,
        description: '整页模型：{ title?, layoutIndex?, notes?, elements:[...] }，元素用英寸坐标 box=[x,y,w,h]。',
        properties: {
          title: { type: 'string', description: '页标题（画布用）。' },
          notes: { type: 'string', description: '演讲者备注。' },
          layoutIndex: { type: 'integer', description: '模板模式下使用哪个版式（0-based）。' },
          elements: { type: 'array', description: '元素列表。', items: { type: 'object', additionalProperties: true, description: '元素：{type,box,text,style,fill,src}。' } },
        },
      },
    },
    /**
     * @param {Record<string, any>} args - tool args.
     * @param {any} exec - execution context.
     * @returns {Promise<Record<string, unknown>>} receipt.
     */
    async execute(args) {
      const found = requireDeck(args)
      if (/** @type {any} */ (found).error) return /** @type {any} */ (found).error
      const entry = /** @type {any} */ (found).entry
      const deck = entry.model
      const slides = /** @type {any[]} */ (deck.slides)
      const oneBased = Math.max(1, Math.floor(Number(args.index) || 1))
      const index = oneBased - 1
      if (oneBased > MAX_SLIDES) return fail('too-many-slides', `最多 ${MAX_SLIDES} 页。`)
      const raw = args.slide !== null && typeof args.slide === 'object' ? args.slide : {}
      const incoming = normalizeSlide(
        { ...raw, id: slides[index]?.id ?? `s${oneBased}` },
        { index, slideW: deck.size.w, slideH: deck.size.h },
      )
      if (slides[index]) slides[index] = incoming
      else slides.push(incoming)
      emit(entry, { type: 'slide', slide: incoming, index, total: slides.length, phase: 'content' })
      emit(entry, { type: 'status', phase: 'drawing', slideIndex: index + 1, total: slides.length, message: `已绘制第 ${index + 1} 页，共 ${slides.length} 页` })
      toolLog('slide', entry, `index=${index + 1} elements=${incoming.elements.length}`)
      return {
        ok: true,
        deckId: deck.id,
        index: index + 1,
        slideId: incoming.id,
        elements: incoming.elements.length,
        total: slides.length,
      }
    },
  }

  /**
   * `pptx_studio_add_element`: append one element (fine-grained drawing).
   */
  const addElementTool = {
    name: 'pptx_studio_add_element',
    output: { schema: OUTPUT_SCHEMA, render: renderReceipt },
    description:
      '往某一页追加一个元素（比整页写入更细，用户能看到元素逐个画出来）。element 用 {type, box, text?, style?, fill?, src?}。',
    parameters: {
      deck: { type: 'string', required: true, description: 'deckId。' },
      slideId: { type: 'string', required: true, description: '页 id（pptx_studio_inspect 里能看到）。' },
      element: { type: 'object', required: true, additionalProperties: true, description: '元素模型：{type, box, text?, style?, fill?, src?}。' },
    },
    /**
     * @param {Record<string, any>} args - tool args.
     * @returns {Promise<Record<string, unknown>>} receipt.
     */
    async execute(args) {
      const found = requireDeck(args)
      if (/** @type {any} */ (found).error) return /** @type {any} */ (found).error
      const entry = /** @type {any} */ (found).entry
      const deck = entry.model
      const slide = /** @type {any[]} */ (deck.slides).find((s) => s.id === args.slideId)
      if (!slide) return fail('slide-not-found', `没有这一页：${args.slideId}；先用 pptx_studio_inspect 看现有 slideId。`)
      const el = normalizeElement(args.element, {
        index: slide.elements.length,
        slideW: deck.size.w,
        slideH: deck.size.h,
        prefix: `${slide.id}-`,
      })
      slide.elements.push(el)
      emit(entry, { type: 'element', slideId: slide.id, element: el, index: slide.elements.length - 1, total: slide.elements.length })
      return { ok: true, deckId: deck.id, slideId: slide.id, elementId: el.id, elements: slide.elements.length }
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
    async execute(args) {
      const found = requireDeck(args)
      if (/** @type {any} */ (found).error) return /** @type {any} */ (found).error
      const entry = /** @type {any} */ (found).entry
      const deck = entry.model
      const change = {
        slideId: String(args.slideId ?? ''),
        elementId: String(args.elementId ?? ''),
        patch: args.patch !== null && typeof args.patch === 'object' ? args.patch : {},
        remove: args.remove === true,
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
      return { ok: true, deckId: deck.id, slideId: change.slideId, elementId: change.elementId, removed: result.removed === true, element: result.element ?? null }
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
    async execute(args) {
      const found = requireDeck(args)
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
   * @returns {Promise<Record<string, unknown>>} export receipt.
   */
  async function exportDeck(entry, target, templatePath) {
    entry.persist()
    const deckFile = join(entry.dir, 'deck.json')
    let out = join(entry.dir, 'deck.pptx')
    if (typeof target === 'string' && target.trim() !== '') {
      out = workspacePath(entry.cwd, target, '')
      if (extname(out).toLowerCase() !== '.pptx') out = `${out}.pptx`
    }
    const args = [deckFile, out]
    if (templatePath && existsSync(templatePath)) args.push('--template', templatePath)
    const info = await python.run('pptx_build.py', args, { timeoutMs: 180_000 })
    const event = emit(entry, {
      type: 'saved',
      path: info.path,
      bytes: info.bytes,
      slides: info.slides,
      shapes: info.shapes,
    })
    log('info', `pptx-studio: 导出 ${info.path}（${info.slides} 页 / ${info.bytes} 字节）`)
    return {
      ok: true,
      deckId: entry.model.id,
      path: info.path,
      bytes: info.bytes,
      slides: info.slides,
      shapes: info.shapes,
      seq: event.seq,
      hint: '用 present 或让用户在文件卡片里打开这个 .pptx；画布与此文件同源。',
    }
  }

  /**
   * `pptx_studio_export`: write the .pptx now.
   */
  const exportTool = {
    name: 'pptx_studio_export',
    output: { schema: OUTPUT_SCHEMA, render: renderReceipt },
    description: '把当前画布导出成真实的 .pptx（python-pptx 生成，原生可编辑）。返回文件绝对路径与页数/字节数。',
    parameters: {
      deck: { type: 'string', required: true, description: 'deckId。' },
      path: { type: 'string', description: '可选输出路径（工作区内），默认 .pptx-studio/<deckId>/deck.pptx。' },
    },
    /**
     * @param {Record<string, any>} args - tool args.
     * @returns {Promise<Record<string, unknown>>} receipt.
     */
    async execute(args) {
      const found = requireDeck(args)
      if (/** @type {any} */ (found).error) return /** @type {any} */ (found).error
      const entry = /** @type {any} */ (found).entry
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
  const toolDefs = [openTool, planTool, slideTool, addElementTool, patchTool, inspectTool, exportTool]
  for (const tool of toolDefs) {
    const defined = defineTool(tool)
    ctx.effect(() => ctx.tools.register(defined), `pptx-studio: ${tool.name}`)
  }

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
    const lines = ['[PPTX 画布反馈]']
    lines.push(`deck=${deck.id} slide=${slideIndex + 1}/${slides.length}${slideId ? ` (${slideId})` : ''}`)
    if (elementId) {
      const current = element
        ? `type=${element.type}${typeof element.text === 'string' ? ` text=「${element.text.slice(0, 80)}」` : ''}`
        : '(元素已不在模型里)'
      lines.push(`element=${elementId} ${current}`)
    }
    if (box) lines.push(`位置：x=${box[0]}in y=${box[1]}in w=${box[2]}in h=${box[3]}in`)
    if (element && element.style) lines.push(`样式：${JSON.stringify(element.style).slice(0, 200)}`)
    if (note) lines.push(`用户意见：${note}`)
    if (action) lines.push(`期望动作：${action}`)
    lines.push('请只改这一处（除非同一问题在别页重复），改完调用 pptx_studio_export 重新导出；不要重做整个 deck。')
    return {
      text: lines.join('\n'),
      meta: { deckId: deck.id, slideId, elementId, bbox: box, slideIndex: slideIndex + 1, total: slides.length },
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
   * Serve one image from the deck's assets directory.
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
    try {
      const body = readFileSync(file)
      const ext = extname(file).toLowerCase()
      const types = { '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.gif': 'image/gif', '.webp': 'image/webp', '.svg': 'image/svg+xml' }
      res.writeHead(200, { 'content-type': types[ext] || 'application/octet-stream', 'cache-control': 'no-store', 'content-length': body.length })
      res.end(body)
    } catch (error) {
      writeJson(res, 500, { ok: false, code: 'asset-read-failed', message: error instanceof Error ? error.message : String(error) })
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
              writeJson(res, 200, { ok: true, name: 'dsh-pptx-studio', decks: store.decks.size })
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
              writeJson(res, 200, { ok: true, decks })
              return
            }
            if (head === 'state' && req.method === 'GET') {
              const entry = store.get(url.searchParams.get('deck') ?? '')
              if (!entry) {
                writeJson(res, 404, { ok: false, code: 'deck-not-found' })
                return
              }
              writeJson(res, 200, {
                ok: true,
                deck: entry.model,
                template: templateOf(entry),
                dir: entry.dir,
                sessionId: entry.sessionId,
              })
              return
            }
            if (head === 'events' && req.method === 'GET') {
              const entry = store.get(url.searchParams.get('deck') ?? '')
              if (!entry) {
                writeJson(res, 404, { ok: false, code: 'deck-not-found' })
                return
              }
              serveEvents(req, res, entry)
              return
            }
            if (head === 'asset' && req.method === 'GET') {
              const entry = store.get(url.searchParams.get('deck') ?? '')
              if (!entry) {
                writeJson(res, 404, { ok: false, code: 'deck-not-found' })
                return
              }
              serveAsset(res, entry, url.searchParams.get('name') ?? '')
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
              try {
                const receipt = await exportDeck(entry, typeof body.path === 'string' ? body.path : undefined, entry.model.template?.path ?? null)
                writeJson(res, 200, { ok: true, value: receipt })
              } catch (error) {
                writeJson(res, 500, { ok: false, code: 'export-failed', message: error instanceof Error ? error.message : String(error) })
              }
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
