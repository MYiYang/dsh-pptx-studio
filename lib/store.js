/**
 * Deck store: in-memory models, per-deck event stream, debounced disk
 * persistence and the ``.pptx-studio`` workspace layout.
 *
 * Layout (inside the session's working directory):
 *   .pptx-studio/<deckId>/deck.json     the model (source of truth on disk)
 *   .pptx-studio/<deckId>/deck.pptx     last export
 *   .pptx-studio/<deckId>/assets/       images referenced by elements
 *   .pptx-studio/<deckId>/template.json parsed template summary
 * @module dsh-pptx-studio/lib/store
 */
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from 'node:fs'
import { join, relative, resolve } from 'node:path'

import { createDeck, normalizeDeck } from './model.js'

/** Root directory name inside the workspace. */
export const STORE_DIR = '.pptx-studio'
/** Debounce before writing deck.json. */
const PERSIST_DELAY_MS = 400

/**
 * Check that `target` stays inside `root` (after resolving both).
 * @param {string} root - workspace directory.
 * @param {string} target - candidate path.
 * @returns {string} resolved target.
 * @throws {Error} when the target escapes the root.
 */
export function ensureInside(root, target) {
  const rootResolved = resolve(root)
  const targetResolved = resolve(target)
  const rel = relative(rootResolved, targetResolved)
  if (rel === '' || (!rel.startsWith('..') && !rel.includes(`..${process.platform === 'win32' ? '\\' : '/'}`) && !rel.startsWith(process.platform === 'win32' ? '\\' : '/'))) {
    if (!rel.startsWith('..')) return targetResolved
  }
  throw new Error(`path escapes the workspace: ${target}`)
}

/**
 * Make a filesystem-safe deck id.
 * @param {string} seed - title or explicit id.
 * @returns {string} id such as `deck_20261113_1a2b3c`.
 */
export function makeDeckId(seed) {
  const now = new Date()
  const stamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`
  const hash = createHash('sha1').update(`${seed}:${Date.now()}:${Math.random()}`).digest('hex').slice(0, 6)
  return `deck_${stamp}_${hash}`
}

/**
 * One deck entry.
 */
class DeckEntry {
  /**
   * @param {object} options - entry options.
   * @param {Record<string, unknown>} options.model - deck model.
   * @param {string} options.dir - deck directory.
   * @param {string|null} options.sessionId - owning session.
   * @param {string} options.cwd - workspace directory.
   */
  constructor({ model, dir, sessionId, cwd }) {
    this.model = model
    this.dir = dir
    this.sessionId = sessionId
    this.cwd = cwd
    this.seq = 0
    /** @type {Set<(event: Record<string, unknown>) => void>} */
    this.listeners = new Set()
    /** @type {NodeJS.Timeout|null} */
    this.persistTimer = null
    this.dirty = false
  }

  /**
   * Emit one event to every subscriber and bump the deck sequence.
   * @param {Record<string, unknown>} event - event payload (without `seq`).
   * @returns {Record<string, unknown>} the emitted event.
   */
  emit(event) {
    this.seq += 1
    const payload = { ...event, seq: this.seq, deckId: this.model.id }
    for (const listener of this.listeners) {
      try {
        listener(payload)
      } catch {
        /* one broken subscriber must not stop the others */
      }
    }
    return payload
  }

  /**
   * Subscribe to this deck's events.
   * @param {(event: Record<string, unknown>) => void} listener - event sink.
   * @returns {() => void} unsubscribe.
   */
  subscribe(listener) {
    this.listeners.add(listener)
    return () => this.listeners.delete(listener)
  }

  /**
   * Schedule a persisted write (400ms debounce, write-once semantics).
   * @param {(error: Error) => void} [onError] - error sink.
   * @returns {void}
   */
  schedulePersist(onError) {
    this.dirty = true
    if (this.persistTimer !== null) return
    this.persistTimer = setTimeout(() => {
      this.persistTimer = null
      if (!this.dirty) return
      this.dirty = false
      try {
        this.persist()
      } catch (error) {
        onError?.(error instanceof Error ? error : new Error(String(error)))
      }
    }, PERSIST_DELAY_MS)
    if (typeof this.persistTimer.unref === 'function') this.persistTimer.unref()
  }

  /**
   * Write deck.json atomically (tmp + rename).
   * @returns {void}
   */
  persist() {
    mkdirSync(this.dir, { recursive: true })
    const target = join(this.dir, 'deck.json')
    const tmp = `${target}.tmp`
    writeFileSync(tmp, JSON.stringify(this.model, null, 2), 'utf8')
    renameSync(tmp, target)
  }

  /**
   * Mark the model as updated.
   * @returns {void}
   */
  touch() {
    const meta = /** @type {any} */ (this.model.meta) || {}
    meta.updatedAt = new Date().toISOString()
    this.model.meta = meta
  }
}

/**
 * Store for all decks the current DSH process knows about.
 */
export class DeckStore {
  /**
   * @param {object} options - store options.
   * @param {(level: string, message: string) => void} options.log - logger sink.
   */
  constructor({ log }) {
    this.log = log
    /** @type {Map<string, DeckEntry>} */
    this.decks = new Map()
  }

  /**
   * Pick the workspace root for a session.
   * @param {string} cwd - session working directory.
   * @returns {string} resolved root.
   */
  rootOf(cwd) {
    return resolve(cwd || process.cwd())
  }

  /**
   * Open (or resume) a deck in a workspace.
   * @param {object} options - open options.
   * @param {string} [options.deckId] - resume an existing deck.
   * @param {string} options.cwd - workspace directory.
   * @param {string|null} options.sessionId - owning session id.
   * @param {string} [options.title] - deck title.
   * @param {number} [options.width] - slide width in inches.
   * @param {number} [options.height] - slide height in inches.
   * @param {string} [options.template] - template .pptx path (already resolved).
   * @returns {DeckEntry} the deck entry.
   */
  open({ deckId, cwd, sessionId, title, width, height, template }) {
    const existing = deckId ? this.decks.get(deckId) : undefined
    if (existing) {
      if (sessionId) existing.sessionId = sessionId
      return existing
    }
    const root = this.rootOf(cwd)
    const id = deckId && /^[A-Za-z0-9_-]{4,80}$/.test(deckId) ? deckId : makeDeckId(title || 'deck')
    const dir = ensureInside(root, join(root, STORE_DIR, id))
    const modelFile = join(dir, 'deck.json')
    /** @type {Record<string, unknown>|null} */
    let model = null
    if (existsSync(modelFile)) {
      try {
        model = normalizeDeck(JSON.parse(readFileSync(modelFile, 'utf8')), id)
      } catch (error) {
        this.log('warn', `pptx-studio: deck.json 读取失败 ${modelFile}: ${error instanceof Error ? error.message : String(error)}`)
      }
    }
    if (model === null) {
      model = createDeck({ id, title, width, height, template })
      if (template) model.template = { path: template }
    }
    if (title && typeof model.title === 'string' && model.title === '未命名演示文稿') model.title = title
    const entry = new DeckEntry({ model, dir, sessionId: sessionId ?? null, cwd: root })
    this.decks.set(id, entry)
    return entry
  }

  /**
   * Find a deck by id.
   * @param {string} deckId - deck id.
   * @returns {DeckEntry|undefined} entry.
   */
  get(deckId) {
    return this.decks.get(deckId)
  }

  /**
   * Forget one deck: cancel its pending write and drop it from the map. The
   * caller removes the directory afterwards; cancelling the debounce first
   * keeps a stale timer from recreating the deck.json we just deleted.
   * @param {string} deckId - deck id.
   * @returns {DeckEntry|null} the entry that was dropped, if any.
   */
  remove(deckId) {
    const entry = this.decks.get(deckId)
    if (!entry) return null
    if (entry.persistTimer !== null) {
      clearTimeout(entry.persistTimer)
      entry.persistTimer = null
    }
    entry.dirty = false
    this.decks.delete(deckId)
    return entry
  }

  /**
   * Decks visible to one session (or all when no session id is given).
   * @param {string|null} sessionId - session id.
   * @returns {DeckEntry[]} entries.
   */
  list(sessionId) {
    const all = [...this.decks.values()]
    return sessionId ? all.filter((entry) => entry.sessionId === sessionId) : all
  }

  /**
   * Directory of one deck's assets.
   * @param {DeckEntry} entry - deck entry.
   * @returns {string} assets directory.
   */
  assetsDir(entry) {
    const dir = join(entry.dir, 'assets')
    mkdirSync(dir, { recursive: true })
    return dir
  }

  /**
   * Mark the deck dirty and broadcast one event.
   * @param {DeckEntry} entry - deck entry.
   * @param {Record<string, unknown>} event - event payload.
   * @returns {Record<string, unknown>} emitted event.
   */
  changed(entry, event) {
    entry.touch()
    entry.schedulePersist((error) => this.log('warn', `pptx-studio: 落盘失败：${error.message}`))
    return entry.emit(event)
  }
}
