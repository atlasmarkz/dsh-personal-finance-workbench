/**
 * dsh-personal-finance-workbench · Host 侧入口
 *
 * 职责：暴露本机 HTTP 接口，供工作台面板读取 Excel/CSV 个人流水、做数据体检、
 * 跑收支分析、导出 xlsx/docx 交付物。分析引擎是一个独立的 Python 脚本
 * （engine/finance.py），使用 DSH Desktop 自带的 office-runtime Python 运行，
 * 数据不出本机。
 */

import { spawn } from 'node:child_process';
import { readFile, writeFile, readdir, mkdir, stat, access } from 'node:fs/promises';
import { constants as fsConstants } from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath } from 'node:url';

export const name = 'dsh-personal-finance-workbench';

/** 需要 connection 服务来挂载本机接口。 */
export const inject = ['connection'];

const API = '/api/personal-finance-workbench';
const ENGINE = fileURLToPath(new URL('./engine/finance.py', import.meta.url));
const DATA_EXTENSIONS = new Set(['.xlsx', '.xlsm', '.csv', '.xls']);

function resolveRoot(applyConfig) {
  const configured = applyConfig && typeof applyConfig.root === 'string' ? applyConfig.root.trim() : '';
  if (configured) return path.resolve(configured);
  const home = process.env.DSH_HOME || path.join(os.homedir(), 'AppData', 'Roaming', 'dsh-desktop', 'harness');
  return path.join(home, 'workbenches', 'personal-finance-workbench');
}

function pythonCandidates(applyConfig) {
  const list = [];
  if (applyConfig && typeof applyConfig.python === 'string' && applyConfig.python.trim()) list.push(applyConfig.python.trim());
  if (process.env.DSH_PERSONAL_FINANCE_PYTHON && process.env.DSH_PERSONAL_FINANCE_PYTHON.trim()) list.push(process.env.DSH_PERSONAL_FINANCE_PYTHON.trim());
  try {
    const exeDir = path.dirname(process.execPath);
    list.push(path.join(exeDir, 'resources', 'office-runtime', 'primary-runtime', 'dependencies', 'python', 'python.exe'));
    list.push(path.join(exeDir, 'resources', 'office-runtime', 'primary-runtime', 'dependencies', 'python', 'bin', 'python3'));
  } catch {}
  list.push('python3', 'python');
  return list;
}

async function firstExisting(candidates) {
  for (const candidate of candidates) {
    if (!candidate.includes(path.sep)) return candidate;
    try {
      await access(candidate, fsConstants.X_OK);
      return candidate;
    } catch {}
  }
  return null;
}

function runEngine(python, command, payload, timeoutMs = 180000) {
  return new Promise((resolve, reject) => {
    const child = spawn(python, [ENGINE, command], {
      windowsHide: true,
      stdio: ['pipe', 'pipe', 'pipe'],
      env: { ...process.env, PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1' },
    });
    const stdout = [];
    const stderr = [];
    let settled = false;
    const timer = setTimeout(() => {
      if (settled) return;
      settled = true;
      try { child.kill(); } catch {}
      reject(new Error('分析超时（超过 3 分钟），请缩小数据量后重试。'));
    }, timeoutMs);
    child.stdout.on('data', (chunk) => stdout.push(chunk));
    child.stderr.on('data', (chunk) => stderr.push(chunk));
    child.on('error', (error) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      reject(new Error(`无法启动分析引擎：${error.message}`));
    });
    child.on('close', (code) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      const out = Buffer.concat(stdout).toString('utf8').trim();
      const err = Buffer.concat(stderr).toString('utf8').trim();
      if (!out) {
        reject(new Error(`分析引擎没有返回结果（退出码 ${code}）。${err ? ' ' + err.slice(-400) : ''}`));
        return;
      }
      let data;
      try {
        data = JSON.parse(out);
      } catch (error) {
        reject(new Error(`分析引擎返回了无法解析的内容：${out.slice(0, 200)}`));
        return;
      }
      if (data && typeof data === 'object' && data.error) {
        reject(new Error(String(data.error)));
        return;
      }
      resolve(data);
    });
    child.stdin.end(JSON.stringify(payload || {}));
  });
}

function jsonResponse(value, status = 200) {
  return new Response(JSON.stringify(value), {
    status,
    headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' },
  });
}

function errorResponse(error, status = 400) {
  const message = error && error.message ? error.message : String(error);
  return jsonResponse({ error: message }, status);
}

async function readJsonBody(request) {
  const text = await request.text();
  if (!text) return {};
  try {
    return JSON.parse(text);
  } catch {
    throw new Error('请求内容不是合法 JSON。');
  }
}

function safeInside(root, candidate) {
  const resolved = path.resolve(candidate);
  const relative = path.relative(root, resolved);
  return relative && !relative.startsWith('..') && !path.isAbsolute(relative) ? relative : null;
}

export function apply(ctx, applyConfig) {
  const root = resolveRoot(applyConfig);
  const exportsDir = path.join(root, 'exports');
  const samplesDir = path.join(root, 'samples');
  const historyFile = path.join(exportsDir, 'history.json');

  async function ensureDirs() {
    await mkdir(exportsDir, { recursive: true });
    await mkdir(samplesDir, { recursive: true });
  }

  async function readHistory() {
    try {
      const raw = await readFile(historyFile, 'utf8');
      const parsed = JSON.parse(raw.replace(/^\uFEFF/, ''));
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  }

  async function writeHistory(items) {
    await ensureDirs();
    await writeFile(historyFile, JSON.stringify(items.slice(-200), null, 2), 'utf8');
  }

  async function engine(command, payload) {
    const python = await firstExisting(pythonCandidates(applyConfig));
    if (!python) throw new Error('没有找到可用的 Python 运行环境，无法执行分析。');
    return runEngine(python, command, payload);
  }

  async function toPublicFile(file) {
    const rel = safeInside(root, file.path);
    return {
      kind: file.kind,
      name: path.basename(file.path),
      bytes: file.bytes,
      path: file.path,
      url: rel ? `${API}/download?rel=${encodeURIComponent(rel)}` : null,
    };
  }

  ctx.connection.fetch.register({
    path: `${API}/info`,
    methods: ['GET'],
    requestBody: 'buffered',
    async fetch() {
      const history = await readHistory();
      return jsonResponse({
        name: '个人财务管理工作台',
        version: '0.1.0',
        root,
        exportsDir,
        samplesDir,
        history: history.slice(-8).reverse(),
      });
    },
  });

  ctx.connection.fetch.register({
    path: `${API}/list`,
    methods: ['POST'],
    requestBody: 'buffered',
    async fetch(request) {
      try {
        const body = await readJsonBody(request);
        const dir = typeof body.dir === 'string' ? body.dir.trim() : '';
        if (!dir) return errorResponse(new Error('缺少目录路径。'));
        const info = await stat(dir);
        if (!info.isDirectory()) return errorResponse(new Error('所选路径不是文件夹。'));
        const entries = await readdir(dir, { withFileTypes: true });
        const files = [];
        for (const entry of entries) {
          if (!entry.isFile()) continue;
          const extension = path.extname(entry.name).toLowerCase();
          if (!DATA_EXTENSIONS.has(extension)) continue;
          const full = path.join(dir, entry.name);
          let bytes = 0;
          let modified = null;
          try {
            const fileInfo = await stat(full);
            bytes = fileInfo.size;
            modified = fileInfo.mtime.toISOString();
          } catch {}
          files.push({ name: entry.name, path: full, bytes, modified });
        }
        files.sort((a, b) => String(b.modified || '').localeCompare(String(a.modified || '')));
        return jsonResponse({ dir, files });
      } catch (error) {
        return errorResponse(error);
      }
    },
  });

  ctx.connection.fetch.register({
    path: `${API}/inspect`,
    methods: ['POST'],
    requestBody: 'buffered',
    async fetch(request) {
      try {
        const body = await readJsonBody(request);
        if (typeof body.path !== 'string' || !body.path.trim()) throw new Error('请先选择数据文件。');
        const result = await engine('inspect', { path: body.path.trim() });
        return jsonResponse(result);
      } catch (error) {
        return errorResponse(error);
      }
    },
  });

  ctx.connection.fetch.register({
    path: `${API}/analyze`,
    methods: ['POST'],
    requestBody: 'buffered',
    async fetch(request) {
      try {
        const body = await readJsonBody(request);
        if (typeof body.path !== 'string' || !body.path.trim()) throw new Error('请先选择数据文件。');
        const result = await engine('analyze', {
          path: body.path.trim(),
          mapping: body.mapping || undefined,
        });
        return jsonResponse(result);
      } catch (error) {
        return errorResponse(error);
      }
    },
  });

  ctx.connection.fetch.register({
    path: `${API}/report`,
    methods: ['POST'],
    requestBody: 'buffered',
    async fetch(request) {
      try {
        const body = await readJsonBody(request);
        if (typeof body.path !== 'string' || !body.path.trim()) throw new Error('请先选择数据文件。');
        const sourcePath = path.resolve(body.path.trim());
        await ensureDirs();
        const history = await readHistory();
        const previous = [...history].reverse().find((item) => item.sourcePath === sourcePath);
        const compare = previous && previous.kpis ? { previous: { ...previous.kpis, generatedAt: previous.generatedAt } } : undefined;
        const result = await engine('report', {
          path: sourcePath,
          mapping: body.mapping || undefined,
          outDir: exportsDir,
          compare,
        });
        const files = [];
        for (const file of result.files || []) files.push(await toPublicFile(file));
        const item = {
          sourcePath,
          sourceName: path.basename(sourcePath),
          generatedAt: result.generatedAt,
          kpis: result.analysis && result.analysis.kpis ? result.analysis.kpis : null,
          files: files.map((file) => ({ kind: file.kind, name: file.name, bytes: file.bytes, path: file.path, url: file.url })),
        };
        history.push(item);
        await writeHistory(history);
        return jsonResponse({ ...result, files, compare: compare || null });
      } catch (error) {
        return errorResponse(error);
      }
    },
  });

  ctx.connection.fetch.register({
    path: `${API}/sample`,
    methods: ['POST'],
    requestBody: 'buffered',
    async fetch() {
      try {
        await ensureDirs();
        const target = path.join(samplesDir, '示例个人流水.xlsx');
        const result = await engine('sample', { out: target });
        return jsonResponse({ path: result.path || target, rows: result.rows || null });
      } catch (error) {
        return errorResponse(error);
      }
    },
  });

  ctx.connection.fetch.register({
    path: `${API}/reveal`,
    methods: ['POST'],
    requestBody: 'buffered',
    async fetch(request) {
      try {
        const body = await readJsonBody(request);
        const target = typeof body.path === 'string' ? path.resolve(body.path) : exportsDir;
        const exists = await stat(target).then(() => true).catch(() => false);
        if (!exists) throw new Error('目标文件或目录不存在。');
        spawn('explorer.exe', [target], { detached: true, windowsHide: true, stdio: 'ignore' }).unref();
        return jsonResponse({ ok: true, path: target });
      } catch (error) {
        return errorResponse(error);
      }
    },
  });

  ctx.connection.fetch.register({
    path: `${API}/download`,
    methods: ['GET'],
    requestBody: 'buffered',
    async fetch(request) {
      try {
        const url = new URL(request.url);
        const rel = url.searchParams.get('rel') || '';
        if (!rel) return errorResponse(new Error('缺少文件参数。'), 404);
        const relPath = safeInside(root, path.join(root, rel));
        if (!relPath) return errorResponse(new Error('文件不在允许的目录内。'), 403);
        const full = path.join(root, relPath);
        const info = await stat(full);
        if (!info.isFile()) return errorResponse(new Error('文件不存在。'), 404);
        const data = await readFile(full);
        const name = path.basename(full);
        const encoded = encodeURIComponent(name).replace(/['()]/g, escape);
        return new Response(data, {
          status: 200,
          headers: {
            'Content-Type': 'application/octet-stream',
            'Content-Disposition': `attachment; filename*=UTF-8''${encoded}`,
            'Content-Length': String(data.length),
            'Cache-Control': 'no-store',
          },
        });
      } catch (error) {
        return errorResponse(error, 404);
      }
    },
  });
}
