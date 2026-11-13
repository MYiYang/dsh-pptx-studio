/**
 * Python bridge: resolve a usable interpreter (python-pptx present) once, then
 * run the bundled `assets/*.py` helpers and parse their JSON on stdout.
 *
 * The plugin never assumes `python` is on PATH: DSH ships its own runtimes
 * under `$DSH_HOME/dsh-runtimes/<runtime>/dependencies/python`, and a profile config
 * can pin an interpreter explicitly.
 * @module dsh-pptx-studio/lib/python
 */
import { spawn } from 'node:child_process'
import { existsSync, readdirSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'

/** Timeout for the interpreter probe. */
const PROBE_TIMEOUT_MS = 20_000

/**
 * Run one command to completion.
 * @param {string} command - executable.
 * @param {string[]} args - arguments.
 * @param {{cwd?: string, timeoutMs?: number, env?: Record<string, string>}} [options] - spawn options.
 * @returns {Promise<{code: number|null, stdout: string, stderr: string, timedOut: boolean}>} result.
 */
export function runProcess(command, args, options = {}) {
  const timeoutMs = options.timeoutMs ?? 120_000
  return new Promise((resolvePromise) => {
    /** @type {import('node:child_process').ChildProcessWithoutNullStreams} */
    let child
    try {
      child = spawn(command, args, {
        cwd: options.cwd,
        windowsHide: true,
        env: { ...process.env, PYTHONIOENCODING: 'utf-8', ...(options.env || {}) },
      })
    } catch (error) {
      resolvePromise({ code: null, stdout: '', stderr: String(error), timedOut: false })
      return
    }
    let stdout = ''
    let stderr = ''
    let timedOut = false
    const timer = setTimeout(() => {
      timedOut = true
      try {
        child.kill()
      } catch {
        /* already gone */
      }
    }, timeoutMs)
    child.stdout.on('data', (chunk) => {
      stdout += String(chunk)
      if (stdout.length > 4_000_000) stdout = stdout.slice(-2_000_000)
    })
    child.stderr.on('data', (chunk) => {
      stderr += String(chunk)
      if (stderr.length > 200_000) stderr = stderr.slice(-100_000)
    })
    child.on('error', (error) => {
      clearTimeout(timer)
      resolvePromise({ code: null, stdout, stderr: `${stderr}${error.message}`, timedOut })
    })
    child.on('close', (code) => {
      clearTimeout(timer)
      resolvePromise({ code, stdout, stderr, timedOut })
    })
  })
}

/**
 * Parse the last JSON line of a helper's stdout.
 * @param {string} stdout - process output.
 * @returns {Record<string, unknown>} parsed object.
 * @throws {Error} when no JSON line is present.
 */
export function parseJsonOutput(stdout) {
  const lines = String(stdout).trim().split(/\r?\n/).filter((line) => line.trim() !== '')
  for (let i = lines.length - 1; i >= 0; i -= 1) {
    const line = lines[i].trim()
    if (!line.startsWith('{') || !line.endsWith('}')) continue
    try {
      return JSON.parse(line)
    } catch {
      /* keep scanning */
    }
  }
  throw new Error(`helper 没有输出 JSON：${lines.slice(-3).join(' | ').slice(0, 500) || '(empty)'}`)
}

/**
 * Locate python interpreters that might carry python-pptx.
 * @param {string|undefined} configured - explicit config value.
 * @returns {string[]} candidate commands.
 */
export function pythonCandidates(configured) {
  const list = []
  if (configured && String(configured).trim() !== '') list.push(String(configured).trim())
  if (process.env.PPTX_STUDIO_PYTHON) list.push(process.env.PPTX_STUDIO_PYTHON)
  const dshHome = process.env.DSH_HOME || join(homedir(), '.dsh')
  const runtimes = join(dshHome, 'dsh-runtimes')
  try {
    for (const dir of readdirSync(runtimes)) {
      const exe = join(runtimes, dir, 'dependencies', 'python', process.platform === 'win32' ? 'python.exe' : 'bin/python3')
      if (existsSync(exe)) list.push(exe)
    }
  } catch {
    /* no DSH runtime directory: fall back to PATH */
  }
  list.push('python', 'python3', 'py')
  return [...new Set(list)]
}

/**
 * Bridge around the bundled helpers.
 */
export class PythonRunner {
  /**
   * @param {object} options - runner options.
   * @param {string} options.assetsDir - directory that holds the .py helpers.
   * @param {(level: string, message: string) => void} options.log - logger sink.
   * @param {string} [options.configured] - explicit interpreter from config.
   */
  constructor({ assetsDir, log, configured }) {
    this.assetsDir = assetsDir
    this.log = log
    this.configured = configured
    /** @type {string|null|undefined} - undefined = not probed yet */
    this.resolved = undefined
    /** @type {Promise<string|null>|null} */
    this.probe = null
  }

  /**
   * Resolve one interpreter that can `import pptx`; cached after the first call.
   * @returns {Promise<string|null>} command, or null when python-pptx is missing.
   */
  async interpreter() {
    if (this.resolved !== undefined) return this.resolved
    if (this.probe !== null) return this.probe
    this.probe = (async () => {
      for (const candidate of pythonCandidates(this.configured)) {
        const result = await runProcess(candidate, ['-c', 'import pptx, sys; sys.stdout.write(pptx.__version__)'], {
          timeoutMs: PROBE_TIMEOUT_MS,
        })
        if (result.code === 0 && /^\d/.test(result.stdout.trim())) {
          this.log('info', `pptx-studio: python = ${candidate} (python-pptx ${result.stdout.trim()})`)
          this.resolved = candidate
          return candidate
        }
      }
      this.log('warn', 'pptx-studio: 找不到带 python-pptx 的解释器；导出功能不可用（可在插件 config.python 里指定）')
      this.resolved = null
      return null
    })().finally(() => {
      this.probe = null
    })
    return this.probe
  }

  /**
   * Run one bundled helper.
   * @param {string} script - file name inside assetsDir.
   * @param {string[]} args - script arguments.
   * @param {{timeoutMs?: number}} [options] - run options.
   * @returns {Promise<Record<string, unknown>>} parsed JSON output.
   */
  async run(script, args, options = {}) {
    const python = await this.interpreter()
    if (python === null) {
      const error = new Error('python-pptx 不可用：请安装 python + python-pptx，或在插件配置里设置 python 路径')
      /** @type {any} */ (error).code = 'python-unavailable'
      throw error
    }
    const scriptPath = join(this.assetsDir, script)
    if (!existsSync(scriptPath)) throw new Error(`helper 缺失：${scriptPath}`)
    const result = await runProcess(python, [scriptPath, ...args], {
      timeoutMs: options.timeoutMs ?? 120_000,
      cwd: this.assetsDir,
    })
    const stdout = result.stdout.trim()
    if (result.code !== 0 && stdout === '') {
      const detail = (result.stderr || '').trim().split(/\r?\n/).slice(-6).join(' | ')
      throw new Error(`${script} 失败（exit ${result.code}${result.timedOut ? ', timeout' : ''}）：${detail || 'no output'}`)
    }
    const parsed = parseJsonOutput(stdout)
    if (parsed.ok === false) {
      const error = new Error(`${script}：${String(parsed.error || 'unknown error')}`)
      /** @type {any} */ (error).detail = parsed
      throw error
    }
    return parsed
  }
}
