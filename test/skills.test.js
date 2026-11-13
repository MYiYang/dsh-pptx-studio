/**
 * Bundled-skill tests: the skill is the plugin's contract, so what it says and
 * how it is registered are both worth pinning down.
 *
 * Run: node test/skills.test.js
 */
import assert from 'node:assert/strict'
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'

import {
  bundledSkillsDir,
  loadBundledSkills,
  parseFrontmatter,
  readSkill,
  registerBundledSkills,
  skillRegistration,
} from '../lib/skills.js'

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

await test('parseFrontmatter reads scalars and strips the header', () => {
  const { meta, body } = parseFrontmatter('---\nname: demo\ndescription: 一句话\n---\n\n# 标题\n正文\n')
  assert.equal(meta.name, 'demo')
  assert.equal(meta.description, '一句话')
  assert.equal(body, '# 标题\n正文')
})

await test('parseFrontmatter unquotes quoted values and folds >- blocks', () => {
  const { meta } = parseFrontmatter('---\nname: "demo"\ndescription: >-\n  第一段\n  第二段\n---\nbody\n')
  assert.equal(meta.name, 'demo')
  assert.equal(meta.description, '第一段 第二段')
})

await test('parseFrontmatter leaves a header-less file alone', () => {
  const { meta, body } = parseFrontmatter('# 没有 frontmatter\n')
  assert.deepEqual(meta, {})
  assert.equal(body, '# 没有 frontmatter')
})

await test('parseFrontmatter handles CRLF line endings', () => {
  const { meta, body } = parseFrontmatter('---\r\nname: demo\r\ndescription: d\r\n---\r\nbody\r\n')
  assert.equal(meta.name, 'demo')
  assert.equal(body, 'body')
})

await test('readSkill loads the packaged pptx-studio skill', () => {
  const skill = readSkill(join(bundledSkillsDir(), 'pptx-studio'), 'pptx-studio')
  assert.ok(skill, 'packaged skill must load')
  assert.equal(skill.name, 'pptx-studio')
  assert.match(skill.description, /唯一通道/)
  assert.match(skill.description, /不要用 python-pptx/)
  // Body keeps the workflow the model needs, and the hard rule.
  assert.match(skill.content, /pptx_studio_open/)
  assert.match(skill.content, /pptx_studio_export/)
  assert.match(skill.content, /以本技能为准/)
  // Frontmatter must not leak into the body.
  assert.ok(!skill.content.startsWith('---'))
})

await test('readSkill rejects a skill without a description', () => {
  const dir = mkdtempSync(join(tmpdir(), 'pptx-studio-skill-bad-'))
  try {
    writeFileSync(join(dir, 'SKILL.md'), '---\nname: bad\n---\nbody\n', 'utf8')
    assert.equal(readSkill(dir, 'bad'), null)
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
})

await test('loadBundledSkills finds exactly one skill and skips broken packs', () => {
  const root = mkdtempSync(join(tmpdir(), 'pptx-studio-skillpack-'))
  try {
    mkdirSync(join(root, 'good'))
    writeFileSync(join(root, 'good', 'SKILL.md'), '---\nname: good\ndescription: fine\n---\nbody\n', 'utf8')
    mkdirSync(join(root, 'Bad_Name'))
    writeFileSync(join(root, 'Bad_Name', 'SKILL.md'), '---\nname: Bad_Name\ndescription: nope\n---\nbody\n', 'utf8')
    const { skills, skipped } = loadBundledSkills({ dir: root })
    assert.equal(skills.length, 1)
    assert.equal(skills[0].name, 'good')
    assert.deepEqual(skipped, ['Bad_Name'])
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
})

await test('loadBundledSkills degrades to an empty pack when the dir is missing', () => {
  const missing = join(tmpdir(), 'pptx-studio-definitely-missing-dir')
  const { skills } = loadBundledSkills({ dir: missing })
  assert.deepEqual(skills, [])
})

await test('skillRegistration carries invocation, resource base and body', () => {
  const skill = readSkill(join(bundledSkillsDir(), 'pptx-studio'), 'pptx-studio')
  assert.ok(skill)
  const registration = skillRegistration(skill)
  assert.equal(registration.name, 'pptx-studio')
  assert.equal(registration.path, skill.path)
  // `source` is validated as a string by the host when the skill is loaded;
  // omitting it makes `skill` fail with "loaded skill ... source must be a string".
  assert.equal(registration.source, 'runtime')
  assert.deepEqual(registration.invocation, { modelInvocable: true, userInvocable: true })
  assert.deepEqual(registration.resourceBase, { kind: 'directory', path: dirname(skill.path) })
  assert.equal(typeof registration.content, 'string')
  assert.ok(/** @type {string} */ (registration.content).length > 200)
})

await test('skillRegistration would survive the host load validator', () => {
  const skill = readSkill(join(bundledSkillsDir(), 'pptx-studio'), 'pptx-studio')
  assert.ok(skill)
  const registration = skillRegistration(skill)
  // Mirror of the guards the registry applies before a skill can be invoked.
  assert.match(String(registration.name), /^[a-z0-9]+(?:-[a-z0-9]+)*$/)
  assert.equal(typeof registration.description, 'string')
  assert.notEqual(String(registration.description).trim(), '')
  assert.equal(typeof registration.content, 'string')
  assert.notEqual(String(registration.content).trim(), '')
  assert.equal(typeof registration.source, 'string')
  const resourceBase = /** @type {any} */ (registration.resourceBase)
  assert.equal(resourceBase?.kind, 'directory')
  assert.equal(typeof resourceBase?.path, 'string')
})

await test('registerBundledSkills needs the skills service and registers once', () => {
  /** @type {string[]} */
  const deps = []
  /** @type {any[]} */
  const registrations = []
  /** @type {Array<() => void>} */
  const effects = []
  const ctx = {
    inject(wanted, callback) {
      deps.push(...wanted)
      callback({
        skills: {
          register(skill) {
            registrations.push(skill)
            return () => {}
          },
        },
        effect(fn) {
          effects.push(fn)
          return () => {}
        },
      })
    },
  }
  const result = registerBundledSkills(ctx)
  assert.deepEqual(deps, ['skills'])
  assert.deepEqual(result.skills, ['pptx-studio'])
  assert.equal(result.registered, true)
  assert.equal(registrations.length, 1)
  assert.equal(registrations[0].name, 'pptx-studio')
  // The registration is wrapped in an effect, so unloading withdraws the rule.
  assert.equal(effects.length, 1)
})

await test('registerBundledSkills survives a host without ctx.inject', () => {
  /** @type {string[]} */
  const warnings = []
  const result = registerBundledSkills({}, { log: (level, message) => warnings.push(`${level}:${message}`) })
  assert.equal(result.registered, false)
  assert.deepEqual(result.skills, ['pptx-studio'])
  assert.ok(warnings.some((line) => line.includes('ctx.inject')))
})

await test('registerBundledSkills survives a registry without register()', () => {
  /** @type {string[]} */
  const warnings = []
  registerBundledSkills(
    { inject: (deps, callback) => callback({ skills: {} }) },
    { log: (level, message) => warnings.push(`${level}:${message}`) },
  )
  assert.ok(warnings.some((line) => line.includes('no register()')))
})

await test('registerBundledSkills reports what it registered, by name', () => {
  /** @type {string[]} */
  const debug = []
  const result = registerBundledSkills(
    { inject: (deps, callback) => callback({ skills: { register: () => () => {} } }) },
    { log: (level, message) => { if (level === 'debug') debug.push(message) } },
  )
  assert.equal(result.registered, true)
  assert.ok(debug.some((line) => line.includes('pptx-studio')))
})

console.log(`\n${passed} checks passed`)
