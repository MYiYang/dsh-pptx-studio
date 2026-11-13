/**
 * Bundled skill: expose `skills/<name>/SKILL.md` to the host skill registry.
 *
 * The skill is this plugin's *contract*, not documentation: it tells the model
 * that `.pptx` is produced through `pptx_studio_*`, and that a skill teaching
 * python-pptx loses to it. Registering it as a provider (instead of copying
 * files into `~/.dsh/skills`) keeps the package self-contained: install the
 * plugin and the rule arrives, uninstall it and the rule leaves with it.
 *
 * Rank/source: the registry files this under the calling context's layer as a
 * `runtime` skill — below a project copy in the same layer, which is exactly
 * what a workspace that wants its own house rules should be able to do.
 *
 * @module dsh-pptx-studio/skills
 */
import { existsSync, readdirSync, readFileSync, statSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

/** Provider name reported to the catalog. */
export const SKILL_PROVIDER = 'dsh-pptx-studio'
/** Same grammar the host registry enforces, so a bad directory name is skipped here too. */
const SKILL_NAME = /^[a-z0-9]+(?:-[a-z0-9]+)*$/

/**
 * Absolute path of this package's `skills/` directory.
 * @returns {string} directory path.
 */
export function bundledSkillsDir() {
  return join(dirname(dirname(fileURLToPath(import.meta.url))), 'skills')
}

/**
 * Strip one YAML frontmatter block from a SKILL.md.
 *
 * Deliberately tiny: `key: value` scalars plus folded (`>`) / literal (`|`)
 * blocks, which is everything a skill header needs. Unknown nesting is ignored
 * rather than guessed at.
 * @param {string} raw - file contents.
 * @returns {{ meta: Record<string, string>, body: string }} header fields + body.
 */
export function parseFrontmatter(raw) {
  const text = String(raw ?? '').replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n')
  const empty = { meta: {}, body: text.trim() }
  if (!text.startsWith('---\n')) return empty
  const end = text.indexOf('\n---', 3)
  if (end === -1) return empty
  const header = text.slice(4, end)
  const bodyStart = text.indexOf('\n', end + 1)
  const body = bodyStart === -1 ? '' : text.slice(bodyStart + 1)

  /** @type {Record<string, string>} */
  const meta = {}
  const lines = header.split('\n')
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i]
    if (line.trim() === '' || line.startsWith('#')) continue
    if (/^\s/.test(line)) continue // nested / continuation lines are not supported
    const colon = line.indexOf(':')
    if (colon <= 0) continue
    const key = line.slice(0, colon).trim()
    let value = line.slice(colon + 1).trim()
    if (value === '>' || value === '|' || value === '>-' || value === '|-') {
      const folded = value.startsWith('>')
      const collected = []
      while (i + 1 < lines.length && (/^\s+\S/.test(lines[i + 1]) || lines[i + 1].trim() === '')) {
        collected.push(lines[i + 1].trim())
        i += 1
      }
      meta[key] = folded ? collected.filter(Boolean).join(' ') : collected.join('\n').trim()
      continue
    }
    if (
      value.length >= 2 &&
      ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'")))
    ) {
      value = value.slice(1, -1)
    }
    meta[key] = value
  }
  return { meta, body: body.trim() }
}

/**
 * Read one skill directory.
 * @param {string} dir - skill directory (contains SKILL.md).
 * @param {string} [fallbackName] - name to use when frontmatter omits one.
 * @returns {{ name: string, description: string, whenToUse?: string, path: string, content: string } | null} skill, or null when unusable.
 */
export function readSkill(dir, fallbackName) {
  const file = join(dir, 'SKILL.md')
  let raw
  try {
    raw = readFileSync(file, 'utf8')
  } catch {
    return null
  }
  const { meta, body } = parseFrontmatter(raw)
  const name = (meta.name || fallbackName || '').trim()
  const description = (meta.description || '').trim()
  if (!SKILL_NAME.test(name) || description === '' || body === '') return null
  return {
    name,
    description,
    whenToUse: meta.when_to_use?.trim() || meta.whenToUse?.trim() || undefined,
    path: file,
    content: body,
  }
}

/**
 * Load every valid skill under one `skills/` directory.
 * Never throws: a missing or unreadable pack degrades to "no bundled skills".
 * @param {{ dir?: string, log?: (level: string, message: string) => void }} [options] - loader options.
 * @returns {{ skills: Array<{name: string, description: string, whenToUse?: string, path: string, content: string}>, skipped: string[] }} loaded skills + skipped directories.
 */
export function loadBundledSkills(options = {}) {
  const log = options.log ?? (() => {})
  const dir = options.dir ?? bundledSkillsDir()
  /** @type {Array<{name: string, description: string, whenToUse?: string, path: string, content: string}>} */
  const skills = []
  /** @type {string[]} */
  const skipped = []
  try {
    if (!existsSync(dir) || !statSync(dir).isDirectory()) {
      log('warn', `pptx-studio: bundled skills directory not found: ${dir}`)
      return { skills, skipped }
    }
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      if (!entry.isDirectory()) continue
      const full = join(dir, entry.name)
      if (!existsSync(join(full, 'SKILL.md'))) continue
      const skill = readSkill(full, entry.name)
      if (skill === null) {
        skipped.push(entry.name)
        continue
      }
      if (skills.some((existing) => existing.name === skill.name)) {
        skipped.push(entry.name)
        continue
      }
      skills.push(skill)
    }
  } catch (error) {
    log('warn', `pptx-studio: bundled skills unreadable (${dir}): ${error instanceof Error ? error.message : String(error)}`)
  }
  if (skipped.length > 0) log('warn', `pptx-studio: bundled skills skipped (bad name/description/body): ${skipped.join(', ')}`)
  return { skills, skipped }
}

/**
 * Put one loaded skill into the shape `ctx.skills.register()` accepts.
 *
 * `source` is not optional: the registry validates it as a string at load time
 * and rejects the whole skill (`loaded skill "x" source must be a string`) —
 * only `invocation` and `provider` receive defaults. `runtime` is the bucket
 * this plug-in actually is.
 * @param {{ name: string, description: string, whenToUse?: string, path: string, content: string }} skill - loaded skill.
 * @returns {Record<string, unknown>} registration payload.
 */
export function skillRegistration(skill) {
  return {
    name: skill.name,
    description: skill.description,
    ...(skill.whenToUse === undefined ? {} : { whenToUse: skill.whenToUse }),
    content: skill.content,
    source: 'runtime',
    invocation: { modelInvocable: true, userInvocable: true },
    resourceBase: { kind: 'directory', path: dirname(skill.path) },
    path: skill.path,
  }
}

/**
 * Register the package's bundled skills once the `skills` service is mounted.
 *
 * Degrades to a log line when the host has no skill registry (older dsh, or a
 * headless profile that does not load the skill packages): a decoration must
 * never be able to stop the plugin from loading.
 * @param {any} ctx - cordis context of the plugin.
 * @param {{ dir?: string, log?: (level: string, message: string) => void }} [options] - registration options.
 * @returns {{ registered: boolean, skills: string[] }} what happened.
 */
export function registerBundledSkills(ctx, options = {}) {
  const log = options.log ?? (() => {})
  const loaded = loadBundledSkills({ dir: options.dir, log })
  const names = loaded.skills.map((skill) => skill.name)
  if (loaded.skills.length === 0) return { registered: false, skills: names }
  if (typeof ctx?.inject !== 'function') {
    log('warn', `pptx-studio: host exposes no ctx.inject(); ${names.length} bundled skill(s) stay unregistered`)
    return { registered: false, skills: names }
  }
  try {
    // `ctx.inject` ties the registration to this plugin's fiber: unloading the
    // plugin unregisters the skill, so a disabled plugin leaves no rule behind.
    ctx.inject(['skills'], (scope) => {
      try {
        const registry = scope?.skills
        if (registry === undefined || typeof registry.register !== 'function') {
          log('warn', 'pptx-studio: skills registry has no register(); bundled skill skipped')
          return
        }
        for (const skill of loaded.skills) {
          const dispose = registry.register(skillRegistration(skill))
          if (typeof dispose === 'function' && typeof scope.effect === 'function') {
            scope.effect(() => dispose, `pptx-studio: skill ${skill.name}`)
          }
        }
        log('debug', `pptx-studio: registered bundled skill(s): ${names.join(', ')}`)
      } catch (error) {
        log('warn', `pptx-studio: bundled skill registration failed: ${error instanceof Error ? error.message : String(error)}`)
      }
    })
  } catch (error) {
    log('warn', `pptx-studio: bundled skill injection failed: ${error instanceof Error ? error.message : String(error)}`)
    return { registered: false, skills: names }
  }
  return { registered: true, skills: names }
}
