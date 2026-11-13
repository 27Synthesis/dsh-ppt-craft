// dsh-ppt-craft — 让 PPT 一次做对：卡片悬浮阴影 / 页面切换 / 组件动画 / 箭头轨道 / 闪动 / 渲染自检
import { spawn } from 'node:child_process'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

export const name = 'dsh-ppt-craft'

const CRAFT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'bin', 'craft.py')
const OUT_DIR = '/sdcard/Download/DSH-Go'
const usage = Object.create(null)

function bump(tool) {
  usage[tool] = (usage[tool] ?? 0) + 1
}

function runCraft(cmd, args, timeout = 240_000) {
  return new Promise((resolve) => {
    let child
    try {
      child = spawn('python3', [CRAFT], { stdio: ['pipe', 'pipe', 'pipe'] })
    } catch (error) {
      resolve({ ok: false, error: `无法启动 python3: ${error.message}` })
      return
    }
    let out = ''
    let err = ''
    const timer = setTimeout(() => child.kill('SIGKILL'), timeout)
    child.stdout.on('data', (d) => { out += d })
    child.stderr.on('data', (d) => { err += d })
    child.on('error', (e) => {
      clearTimeout(timer)
      resolve({ ok: false, error: `spawn 失败: ${e.message}` })
    })
    child.on('close', () => {
      clearTimeout(timer)
      try {
        resolve(JSON.parse(out))
      } catch {
        resolve({ ok: false, error: (err || out || 'python 无输出').trim().slice(0, 900) })
      }
    })
    child.stdin.end(JSON.stringify({ cmd, args }))
  })
}

// 工具返回值必须是对象；统一把结果包成 { text }
const wrap = (fn) => async (args) => {
  const r = await fn(args)
  return typeof r === 'string' ? { text: r } : r
}

function tool(spec) {
  return {
    name: spec.name,
    description: spec.description,
    parameters: spec.parameters,
    output: {
      schema: { type: 'object', properties: { text: { type: 'string' } }, additionalProperties: true },
      render: (_args, value) => [{ type: 'text', text: typeof value === 'string' ? value : JSON.stringify(value) }],
    },
    execute: spec.execute,
  }
}

const P = (props) => ({ type: 'object', properties: props, required: [], additionalProperties: false })

async function call(tool, cmd, args) {
  bump(tool)
  const res = await runCraft(cmd, args)
  if (!res.ok) return `❌ ${res.error}${res.trace ? '\n' + res.trace : ''}`
  return res.data
}

// ------------------------------------------------------------------ 格式化

function fmtAudit(d) {
  const lines = [
    `📄 ${d.slides} 页，${d.size_in[0]}×${d.size_in[1]} 英寸`,
    `🎬 有动画的页：${d.with_animation}，动画步数：${d.total_effects}`,
    '',
    '页\t卡片\t缺阴影\t切换\t动画步',
  ]
  for (const p of d.pages) {
    lines.push(`${p.page}\t${p.cards}\t${p.cards_without_shadow}\t${p.transition ? '✓' : '✗'}\t${p.effects}`)
  }
  if (d.media.length) {
    lines.push('', '视频/媒体：', ...d.media)
  }
  lines.push('', d.issues.length ? '问题：' : '✅ 没有发现问题')
  lines.push(...d.issues.slice(0, 30))
  return lines.join('\n')
}

function fmtPolish(d) {
  return [`✅ 已保存：${d.out}（${d.slides} 页）`, ...d.report].join('\n')
}

function fmtAnimate(d) {
  const head = d.warn ? `⚠️ ${d.warn}` : `✅ 已保存：${d.out}（${d.slides} 页，重新打开校验通过）`
  return [head, ...d.report].join('\n')
}

function fmtRender(d) {
  const files = d.files.slice(0, 12)
  return [
    `✅ 渲染 ${d.count} 张（${d.width}×${d.height}）到 ${path.dirname(d.files[0] || '/tmp/x')}`,
    ...files,
    d.files.length > 12 ? `…还有 ${d.files.length - 12} 张` : '',
  ].filter(Boolean).join('\n')
}

// ------------------------------------------------------------------ 参数

const PATH = { type: 'string', description: '本机 .pptx 绝对路径' }
const OUT = { type: 'string', description: '输出路径；省略则覆盖输入文件' }

const POLISH = {
  path: PATH,
  out: { type: 'string', description: '输出路径；省略覆盖原文件。改用户文件前先备份' },
  card_fills: { type: 'array', items: { type: 'string' }, description: '算作卡片的填充色（6 位 hex）；默认纸白/米色一组' },
  border: { type: 'object', additionalProperties: true, description: '填充色 → 描边色，如 {"FFFCF5":"EAE1D1"}' },
  shadow: { type: 'object', additionalProperties: true, description: '{blur,dist,alpha,color}；默认 9pt/3pt/17%/2B2A26' },
  accent_fills: { type: 'array', items: { type: 'string' }, description: '强调块填充色（只加重阴影、不加描边），默认 ["B0472C"]' },
  recolor: { type: 'object', additionalProperties: true, description: '换色，如 {"F6F3EE":"EAE0CC"}（把和背景糊在一起的灰卡换掉）' },
  min_h: { type: 'number', description: '最小高度（英寸），默认 0.55，用来排除细线/小标签' },
  min_w: { type: 'number', description: '最小宽度，默认 0.9' },
  skip_prefix: { type: 'array', items: { type: 'string' }, description: '按名字跳过的形状前缀，默认 ["ribbon","tab"]' },
  renumber: { type: 'boolean', description: '把页脚「n / N」和「NN」大号页码按当前页序重写' },
  delete_text: { type: 'array', items: { type: 'string' }, description: '按页内文字删页（不能按序号删！）' },
}

const ANIMATE = {
  path: PATH,
  out: OUT,
  spec: {
    type: 'object', additionalProperties: true,
    description: [
      '动画脚本。{transition:"fade", auto:{pages:[4,5],stagger:150}, slides:{"4":[[{"eff":"wipe","target":["ribbon1"],"dir":"left","dur":650},{"eff":"pulse","target":["Rounded Rectangle 7"]}],[...]]}, reveal:[{page:12,trigger:["问号卡"],hide:["答案卡"]}]}',
      'slides 里每个内层数组 = 一次点击；eff 可选 fade / wipe(dir:left/right/up/down) / pulse(scale) / hide。target 用形状名（shapes 列表里能看到的 name）。',
    ].join(' '),
  },
}

const RENDER = {
  path: PATH,
  outdir: { type: 'string', description: '输出目录，默认 /tmp/ppt-render' },
  pages: { type: 'array', items: { type: 'integer' }, description: '只渲染这几页（从 1 开始）' },
  scale: { type: 'integer', description: '每英寸像素，默认 130（13.33in 宽 → 1733px）' },
}

// ------------------------------------------------------------------ 插件入口

export function apply(ctx) {
  ctx.inject(['tools'], (ctx) => {
    ctx.tools.register(tool({
      name: 'ppt_render',
      description: '把 .pptx 渲染成 PNG 来核对版式：会画出圆角、描边、朝下阴影和文本换行，接近 PowerPoint 观感。做完一页就渲染一次，比盲改可靠。',
      parameters: P(RENDER),
      execute: wrap(async (args) => {
        const r = await call('ppt_render', 'render', args)
        return typeof r === 'string' ? r : fmtRender(r)
      }),
    }))

    ctx.tools.register(tool({
      name: 'ppt_audit',
      description: '体检一份 .pptx：每页卡片数、缺阴影的卡片、有没有页面切换、动画步数、动画 id 是否重复（重复会导致文件打不开）、视频比例。改完稿交付前跑一次。',
      parameters: P({ path: PATH }),
      execute: wrap(async (args) => {
        const r = await call('ppt_audit', 'audit', args)
        return typeof r === 'string' ? r : fmtAudit(r)
      }),
    }))

    ctx.tools.register(tool({
      name: 'ppt_polish',
      description: '一批给「卡片」加悬浮感：细描边 + 朝下柔和阴影（PowerPoint/WPS 都能显示），可同时换掉糊在背景里的灰卡、按文字删页、重写页码。不用逐页手改。',
      parameters: P(POLISH),
      execute: wrap(async (args) => {
        const r = await call('ppt_polish', 'polish', args)
        return typeof r === 'string' ? r : fmtPolish(r)
      }),
    }))

    ctx.tools.register(tool({
      name: 'ppt_animate',
      description: '往 .pptx 里写真正的 Office 动画：全篇淡入切换 + 点击逐步出现（淡入/擦除/脉冲闪动/隐藏），适合「点一下出一段箭头」「点问号才出答案」「卡片依次浮现」。写完会重新打开校验。',
      parameters: P(ANIMATE),
      execute: wrap(async (args) => {
        const spec = args.spec ?? {}
        const r = await call('ppt_animate', 'animate', { path: args.path, out: args.out ?? OUT_DIR + '/animated.pptx', spec })
        return typeof r === 'string' ? r : fmtAnimate(r)
      }),
    }))
  })
}