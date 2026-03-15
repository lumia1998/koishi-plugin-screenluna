import { Context, Schema, h } from 'koishi'
import { IncomingMessage } from 'http'
import { WebSocket } from 'ws'
import { } from '@koishijs/plugin-server'
import * as fs from 'fs/promises'
import * as path from 'path'

export const name = 'monitorluna'

export const usage = `
## MonitorLuna 使用说明

### 下载 Windows Agent

[点击下载 monitorluna-agent.zip](https://github.com/lumia1998/koishi-plugin-monitorluna/releases/download/v1.1.1/monitorluna-agent.zip)

解压后双击 \`start-silent.vbs\` 启动，然后访问 http://127.0.0.1:6315 配置连接信息。

### 浏览器扩展（可选）

在 \`browser-extension/\` 目录中有 Chrome/Edge 扩展，可追踪浏览器各域名的活跃时长。

### 命令

| 命令 | 说明 |
|------|------|
| \`monitor.list\` | 列出所有在线设备 |
| \`monitor.screen <设备ID>\` | 截取设备全屏截图 |
| \`monitor.window <设备ID>\` | 截取设备当前活跃窗口 |
| \`monitor.status <设备ID>\` | 查看设备 CPU/内存/GPU 状态 |
| \`monitor.analytics <设备ID>\` | 生成当天活动总结图（需要 puppeteer 插件） |
| \`monitor.browser <设备ID>\` | 查看今日浏览器域名时长排行 |

`

export const inject = {
  required: ['server', 'database'],
  optional: ['puppeteer']
}

// ── Database Tables ──
declare module 'koishi' {
  interface Tables {
    monitorluna_activity: ActivityLog
    monitorluna_screenshot: Screenshot
    monitorluna_input_stats: InputStats
    monitorluna_browser_activity: BrowserActivity
  }
}

interface ActivityLog {
  id: number
  deviceId: string
  process: string
  title: string
  timestamp: Date
  endTime: Date
  afk: boolean
}

interface Screenshot {
  id: number
  deviceId: string
  url: string
  timestamp: Date
}

interface InputStats {
  id: number
  deviceId: string
  process: string
  displayName: string
  iconBase64: string
  keyPresses: number
  leftClicks: number
  rightClicks: number
  scrollDistance: number
  timestamp: Date
}

interface BrowserActivity {
  id: number
  deviceId: string
  domain: string
  duration: number
  timestamp: Date
}

// ── Storage Backend Interface ──
interface StorageBackend {
  upload(buffer: Buffer, filename: string): Promise<{ key: string; url: string }>
  delete(key: string): Promise<void>
  init(): Promise<void>
  cleanup(days: number): Promise<void>
}

// ── Local Storage ──
class LocalStorage implements StorageBackend {
  constructor(private ctx: Context, private config: Config) { }

  async init() {
    const dir = path.join(this.ctx.baseDir, this.config.storagePath || 'data/monitorluna')
    await fs.mkdir(dir, { recursive: true })

    // Mount HTTP route
    this.ctx.server.get('/monitorluna/:filename', async (ctx) => {
      try {
        const file = path.join(dir, ctx.params.filename)
        const buf = await fs.readFile(file)
        ctx.type = 'image/jpeg'
        ctx.body = buf
      } catch {
        ctx.status = 404
      }
    })
  }

  async upload(buffer: Buffer, filename: string) {
    const dir = path.join(this.ctx.baseDir, this.config.storagePath || 'data/monitorluna')
    const filepath = path.join(dir, filename)
    await fs.writeFile(filepath, buffer)
    const baseUrl = this.config.serverPath || 'http://127.0.0.1:5140'
    return { key: filename, url: `${baseUrl}/monitorluna/${filename}` }
  }

  async delete(key: string) {
    const dir = path.join(this.ctx.baseDir, this.config.storagePath || 'data/monitorluna')
    await fs.unlink(path.join(dir, key)).catch(() => { })
  }

  async cleanup(days: number) {
    const dir = path.join(this.ctx.baseDir, this.config.storagePath || 'data/monitorluna')
    const cutoff = Date.now() - days * 86400000
    const files = await fs.readdir(dir)
    for (const file of files) {
      const stat = await fs.stat(path.join(dir, file))
      if (stat.mtimeMs < cutoff) await this.delete(file)
    }
  }
}

// ── WebDAV Storage ──
class WebDAVStorage implements StorageBackend {
  constructor(private config: Config) { }

  async init() { }

  async upload(buffer: Buffer, filename: string) {
    const url = `${this.config.webdavEndpoint}/${filename}`
    const auth = Buffer.from(`${this.config.webdavUsername}:${this.config.webdavPassword}`).toString('base64')
    const res = await fetch(url, {
      method: 'PUT',
      headers: { 'Authorization': `Basic ${auth}` },
      body: buffer as unknown as BodyInit
    })
    if (!res.ok) throw new Error(`WebDAV upload failed: ${res.status} ${res.statusText}`)
    const publicUrl = this.config.webdavPublicUrl || this.config.webdavEndpoint
    return { key: filename, url: `${publicUrl}/${filename}` }
  }

  async delete(key: string) {
    const url = `${this.config.webdavEndpoint}/${key}`
    const auth = Buffer.from(`${this.config.webdavUsername}:${this.config.webdavPassword}`).toString('base64')
    await fetch(url, { method: 'DELETE', headers: { 'Authorization': `Basic ${auth}` } }).catch(() => { })
  }

  async cleanup(days: number) {
    const cutoff = new Date(Date.now() - days * 86400000)
    if (!this.config.webdavEndpoint) return
    const auth = Buffer.from(`${this.config.webdavUsername}:${this.config.webdavPassword}`).toString('base64')

    try {
      // PROPFIND to list files
      const res = await fetch(this.config.webdavEndpoint, {
        method: 'PROPFIND',
        headers: {
          'Authorization': `Basic ${auth}`,
          'Depth': '1',
          'Content-Type': 'application/xml'
        },
        body: '<?xml version="1.0"?><propfind xmlns="DAV:"><prop><getlastmodified/></prop></propfind>'
      })
      if (!res.ok) return

      const xml = await res.text()
      // Simple regex to extract href and getlastmodified
      const hrefRegex = /<D:href>([^<]+)<\/D:href>/gi
      const modifiedRegex = /<D:getlastmodified>([^<]+)<\/D:getlastmodified>/gi

      const hrefs: string[] = []
      const modifieds: string[] = []
      let match
      while ((match = hrefRegex.exec(xml)) !== null) hrefs.push(match[1])
      while ((match = modifiedRegex.exec(xml)) !== null) modifieds.push(match[1])

      for (let i = 0; i < Math.min(hrefs.length, modifieds.length); i++) {
        const href = hrefs[i]
        const modified = new Date(modifieds[i])
        if (modified < cutoff && href.includes('.jpg')) {
          await this.delete(href.split('/').pop() || '')
        }
      }
    } catch {}
  }
}

// ── S3 Storage ──
class S3Storage implements StorageBackend {
  constructor(private config: Config) { }

  async init() { }

  async upload(buffer: Buffer, filename: string) {
    const url = this.getUrl(filename)
    const headers = await this.signRequest('PUT', filename, buffer)
    const res = await fetch(url, { method: 'PUT', headers, body: buffer as unknown as BodyInit })
    if (!res.ok) throw new Error(`S3 upload failed: ${res.status} ${res.statusText}`)
    const publicUrl = this.config.s3PublicUrl || url
    return { key: filename, url: publicUrl }
  }

  async delete(key: string) {
    const url = this.getUrl(key)
    const headers = await this.signRequest('DELETE', key)
    await fetch(url, { method: 'DELETE', headers }).catch(() => { })
  }

  async cleanup(days: number) {
    const cutoff = new Date(Date.now() - days * 86400000)
    if (!this.config.s3Endpoint) return
    try {
      const listUrl = this.getUrl('') + '?list-type=2'
      const listHeaders = await this.signRequest('GET', '')
      const res = await fetch(listUrl, { method: 'GET', headers: listHeaders })
      if (!res.ok) return

      const xml = await res.text()
      const keyRegex = /<Key>([^<]+)<\/Key>/g
      const modifiedRegex = /<LastModified>([^<]+)<\/LastModified>/g

      const keys: string[] = []
      const modifieds: string[] = []
      let match
      while ((match = keyRegex.exec(xml)) !== null) keys.push(match[1])
      while ((match = modifiedRegex.exec(xml)) !== null) modifieds.push(match[1])

      const toDelete = keys.filter((_, i) => {
        const modified = new Date(modifieds[i])
        return modified < cutoff
      })

      if (toDelete.length === 0) return

      const deleteXml = [
        '<?xml version="1.0"?><Delete>',
        ...toDelete.map(k => `<Object><Key>${k}</Key></Object>`),
        '</Delete>'
      ].join('')
      const deleteBuf = Buffer.from(deleteXml)
      const deleteUrl = this.getUrl('') + '?delete'
      const deleteHeaders = await this.signRequest('POST', '', deleteBuf)
      await fetch(deleteUrl, { method: 'POST', headers: deleteHeaders, body: deleteBuf as unknown as BodyInit })
    } catch {}
  }

  private getUrl(key: string) {
    const { s3Endpoint, s3Bucket, s3PathStyle } = this.config
    if (!s3Endpoint) throw new Error('s3Endpoint is required')
    if (s3PathStyle) return `${s3Endpoint}/${s3Bucket}/${key}`
    return `${s3Endpoint.replace('://', `://${s3Bucket}.`)}/${key}`
  }

  private async signRequest(method: string, key: string, body?: Buffer) {
    const date = new Date().toUTCString()
    const contentType = 'image/jpeg'
    const resource = `/${this.config.s3Bucket}/${key}`
    const stringToSign = `${method}\n\n${contentType}\n${date}\n${resource}`

    const crypto = await import('crypto')
    const secret = this.config.s3SecretAccessKey || ''
    const signature = crypto.createHmac('sha1', secret)
      .update(stringToSign).digest('base64')

    return {
      'Authorization': `AWS ${this.config.s3AccessKeyId}:${signature}`,
      'Date': date,
      'Content-Type': contentType
    }
  }
}

// ── Config Schema ──
export interface Config {
  token: string
  commandTimeout: number
  debug: boolean
  storageType: 'local' | 'webdav' | 's3'
  storagePath?: string
  serverPath?: string
  webdavEndpoint?: string
  webdavUsername?: string
  webdavPassword?: string
  webdavPublicUrl?: string
  s3Endpoint?: string
  s3Bucket?: string
  s3Region?: string
  s3AccessKeyId?: string
  s3SecretAccessKey?: string
  s3PublicUrl?: string
  s3PathStyle?: boolean
  imageRetentionDays: number
  dailySummaryEnabled: boolean
  dailySummaryTime: string
  dailySummaryTargets: Array<{
    deviceId: string
    channelIds: string[]
  }>
}

export const Config: Schema<Config> = Schema.intersect([
  Schema.object({
    token: Schema.string().required().description('鉴权 Token，需与 Agent 端一致'),
    commandTimeout: Schema.number().default(15000).description('指令超时时间（毫秒）'),
    debug: Schema.boolean().default(false).description('启用调试日志'),
    storageType: Schema.union(['local', 'webdav', 's3']).default('local').description('存储后端类型'),
  }),
  Schema.union([
    Schema.object({
      storageType: Schema.const('local'),
      storagePath: Schema.string().default('data/monitorluna').description('本地存储目录'),
      serverPath: Schema.string().description('Koishi 公网地址（用于生成图片 URL）'),
    }),
    Schema.object({
      storageType: Schema.const('webdav'),
      webdavEndpoint: Schema.string().required().description('WebDAV 端点'),
      webdavUsername: Schema.string().required().description('WebDAV 用户名'),
      webdavPassword: Schema.string().role('secret').required().description('WebDAV 密码'),
      webdavPublicUrl: Schema.string().description('WebDAV 公网访问地址'),
    }),
    Schema.object({
      storageType: Schema.const('s3'),
      s3Endpoint: Schema.string().required().description('S3 端点'),
      s3Bucket: Schema.string().required().description('S3 Bucket'),
      s3Region: Schema.string().default('us-east-1').description('S3 区域'),
      s3AccessKeyId: Schema.string().required().description('Access Key ID'),
      s3SecretAccessKey: Schema.string().role('secret').required().description('Secret Access Key'),
      s3PublicUrl: Schema.string().description('S3 公网访问地址'),
      s3PathStyle: Schema.boolean().default(false).description('使用路径风格 URL'),
    }),
  ]),
  Schema.object({
    imageRetentionDays: Schema.number().default(7).description('图片保存天数'),
    dailySummaryEnabled: Schema.boolean().default(false).description('启用每日总结'),
    dailySummaryTime: Schema.string().default('22:00').description('每日总结时间（HH:mm）'),
    dailySummaryTargets: Schema.array(Schema.object({
      deviceId: Schema.string().required().description('设备 ID'),
      channelIds: Schema.array(String).description('推送目标群组'),
    })).default([]).description('每日总结推送配置'),
  }),
])

// ── Interfaces ──
interface DeviceConnection {
  ws: WebSocket
  deviceId: string
  pendingCommands: Map<string, {
    resolve: (data: string) => void
    reject: (err: Error) => void
    timer: ReturnType<typeof setTimeout>
  }>
}

function generateId(): string {
  return crypto.randomUUID()
}

// ── Main Plugin ──
export function apply(ctx: Context, config: Config) {
  const devices = new Map<string, DeviceConnection>()
  let storage: StorageBackend

  // Initialize database
  ctx.model.extend('monitorluna_activity', {
    id: 'unsigned',
    deviceId: 'string',
    process: 'string',
    title: 'text',
    timestamp: 'timestamp',
    endTime: 'timestamp',
    afk: 'boolean',
  }, { autoInc: true })

  ctx.model.extend('monitorluna_screenshot', {
    id: 'unsigned',
    deviceId: 'string',
    url: 'string',
    timestamp: 'timestamp'
  }, { autoInc: true })

  ctx.model.extend('monitorluna_input_stats', {
    id: 'unsigned',
    deviceId: 'string',
    process: 'string',
    displayName: 'string',
    iconBase64: 'text',
    keyPresses: 'unsigned',
    leftClicks: 'unsigned',
    rightClicks: 'unsigned',
    scrollDistance: 'double',
    timestamp: 'timestamp'
  }, { autoInc: true })

  ctx.model.extend('monitorluna_browser_activity', {
    id: 'unsigned',
    deviceId: 'string',
    domain: 'string',
    duration: 'double',
    timestamp: 'timestamp',
  }, { autoInc: true })

  // Initialize storage
  if (config.storageType === 'local') storage = new LocalStorage(ctx, config)
  else if (config.storageType === 'webdav') storage = new WebDAVStorage(config)
  else storage = new S3Storage(config)

  ctx.on('ready', async () => {
    await storage.init()
    startCleanupTimer()
    startPeriodicScreenshot()
    if (config.dailySummaryEnabled) startDailySummary()
  })

  ctx.on('dispose', () => {
  })

  // WebSocket handler
  ctx.server.ws('/monitorluna', (ws: WebSocket, req: IncomingMessage) => {
    let deviceId: string | null = null
    let device: DeviceConnection | null = null

    ws.on('message', (raw) => {
      let msg: any
      try {
        msg = JSON.parse(raw.toString())
      } catch {
        ws.close(1008, 'invalid json')
        return
      }

      if (msg.type === 'hello') {
        if (msg.token !== config.token) {
          ws.send(JSON.stringify({ type: 'error', message: 'invalid token' }))
          ws.close(1008, 'invalid token')
          return
        }
        deviceId = String(msg.device_id || 'unknown')
        device = { ws, deviceId, pendingCommands: new Map() }
        devices.set(deviceId, device)
        ctx.logger.info(`[monitorluna] 设备上线: ${deviceId}`)
        ws.send(JSON.stringify({ type: 'hello_ack', message: 'connected' }))
        return
      }

      if (msg.type === 'activity') {
        if (!deviceId) {
          ws.close(1008, 'not authenticated')
          return
        }
        handleActivity(deviceId, msg).catch(e => ctx.logger.warn(`[monitorluna] 处理活动失败: ${e.message}`))
        return
      }

      if (msg.type === 'input_stats') {
        if (!deviceId) {
          ws.close(1008, 'not authenticated')
          return
        }
        handleInputStats(deviceId, msg).catch(e => ctx.logger.warn(`[monitorluna] 处理输入统计失败: ${e.message}`))
        return
      }

      if (msg.type === 'browser_activity') {
        // Browser extension authenticates via token in the message (not WebSocket handshake)
        if (msg.token !== config.token) {
          ws.send(JSON.stringify({ type: 'error', message: 'invalid token' }))
          ws.close(1008, 'invalid token')
          return
        }
        const browserDeviceId = String(msg.device_id || 'unknown')
        handleBrowserActivity(browserDeviceId, msg).catch(e => ctx.logger.warn(`[monitorluna] 处理浏览器活动失败: ${e.message}`))
        return
      }

      if (!device) {
        ws.close(1008, 'not authenticated')
        return
      }

      if (msg.type === 'result' && msg.id) {
        const pending = device.pendingCommands.get(msg.id)
        if (pending) {
          clearTimeout(pending.timer)
          device.pendingCommands.delete(msg.id)
          if (msg.ok) pending.resolve(msg.data)
          else pending.reject(new Error(msg.error || 'unknown error'))
        }
      }
    })

    ws.on('close', () => {
      if (deviceId) {
        devices.delete(deviceId)
        ctx.logger.info(`[monitorluna] 设备下线: ${deviceId}`)
        if (device) {
          for (const pending of device.pendingCommands.values()) {
            clearTimeout(pending.timer)
            pending.reject(new Error('设备断开连接'))
          }
          device.pendingCommands.clear()
        }
      }
    })

    ws.on('error', (err) => {
      ctx.logger.warn(`[monitorluna] WebSocket 错误: ${err.message}`)
    })
  })

  function sendCommand(deviceId: string, cmd: string): Promise<string> {
    const device = devices.get(deviceId)
    if (!device) return Promise.reject(new Error(`设备 "${deviceId}" 不在线`))
    const id = generateId()
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        device.pendingCommands.delete(id)
        reject(new Error('指令超时'))
      }, config.commandTimeout)
      device.pendingCommands.set(id, { resolve, reject, timer })
      device.ws.send(JSON.stringify({ type: 'command', id, cmd }))
    })
  }

  async function handleActivity(deviceId: string, msg: any) {
    const process = msg.process
    const title = msg.title
    const afk: boolean = msg.afk === true
    if (config.debug) ctx.logger.info(`[monitorluna][debug] activity: device=${deviceId}, process=${process}, title=${title}, afk=${afk}`)
    try {
      const now = new Date()
      const cutoff = new Date(now.getTime() - 10000)
      const recent = await ctx.database.get('monitorluna_activity', {
        deviceId,
        process,
        endTime: { $gte: cutoff }
      }, { limit: 1, sort: { endTime: 'desc' } })
      if (recent.length > 0) {
        await ctx.database.set('monitorluna_activity', recent[0].id, { endTime: now, afk })
      } else {
        await ctx.database.create('monitorluna_activity', {
          deviceId,
          process,
          title,
          timestamp: now,
          endTime: now,
          afk,
        })
      }
    } catch (e) {
      ctx.logger.warn(`[monitorluna] 记录活动失败: ${e instanceof Error ? e.message : String(e)}`)
    }
  }

  async function handleInputStats(deviceId: string, msg: any) {
    const stats = msg.stats
    if (!stats) return
    if (config.debug) ctx.logger.info(`[monitorluna][debug] input_stats: device=${deviceId}, apps=${Object.keys(stats).length}`)
    try {
      const now = new Date()
      const rows = (Object.entries(stats) as [string, any][]).map(([process, data]) => ({
        deviceId,
        process,
        displayName: data.display_name || process,
        iconBase64: data.icon_base64 || '',
        keyPresses: data.key_presses || 0,
        leftClicks: data.left_clicks || 0,
        rightClicks: data.right_clicks || 0,
        scrollDistance: data.scroll_distance || 0,
        timestamp: now
      }))
      if (rows.length > 0) {
        await ctx.database.upsert('monitorluna_input_stats', rows)
      }
    } catch (e) {
      ctx.logger.warn(`[monitorluna] 记录输入统计失败: ${e instanceof Error ? e.message : String(e)}`)
    }
  }

  async function handleBrowserActivity(browserDeviceId: string, msg: any) {
    const stats = msg.stats
    if (!stats || typeof stats !== 'object') return
    if (config.debug) ctx.logger.info(`[monitorluna][debug] browser_activity: device=${browserDeviceId}, domains=${Object.keys(stats).length}`)
    try {
      const today = new Date()
      today.setHours(0, 0, 0, 0)
      const tomorrow = new Date(today)
      tomorrow.setDate(tomorrow.getDate() + 1)

      for (const [domain, duration] of Object.entries(stats)) {
        if (typeof duration !== 'number' || duration <= 0) continue
        const existing = await ctx.database.get('monitorluna_browser_activity', {
          deviceId: browserDeviceId,
          domain,
          timestamp: { $gte: today, $lt: tomorrow }
        }, { limit: 1 })
        if (existing.length > 0) {
          await ctx.database.set('monitorluna_browser_activity', existing[0].id, {
            duration: existing[0].duration + duration
          })
        } else {
          await ctx.database.create('monitorluna_browser_activity', {
            deviceId: browserDeviceId,
            domain,
            duration,
            timestamp: today
          })
        }
      }
    } catch (e) {
      ctx.logger.warn(`[monitorluna] 记录浏览器活动失败: ${e instanceof Error ? e.message : String(e)}`)
    }
  }

  function startPeriodicScreenshot() {
    setInterval(async () => {
      for (const [deviceId] of devices) {
        try {
          const data = await sendCommand(deviceId, 'screenshot')
          const buf = Buffer.from(data, 'base64')
          const filename = `${deviceId}_${Date.now()}.jpg`
          const { url } = await storage.upload(buf, filename)
          await ctx.database.create('monitorluna_screenshot', {
            deviceId,
            url,
            timestamp: new Date()
          })
          if (config.debug) ctx.logger.info(`[monitorluna][debug] periodic screenshot ok: ${deviceId}, url: ${url}`)
        } catch (e) {
          // Device offline, skip
        }
      }
    }, 15 * 60 * 1000) // 每 15 分钟
  }

  function startCleanupTimer() {
    setInterval(() => {
      storage.cleanup(config.imageRetentionDays).catch(e =>
        ctx.logger.warn(`[monitorluna] 清理失败: ${e.message}`)
      )
    }, 86400000) // Daily
  }

  function startDailySummary() {
    const [hour, minute] = config.dailySummaryTime.split(':').map(Number)
    setInterval(() => {
      const now = new Date()
      if (now.getHours() === hour && now.getMinutes() === minute) {
        generateDailySummary()
      }
    }, 60000)
  }

  async function generateDailySummary() {
    const today = new Date()
    today.setHours(0, 0, 0, 0)
    const tomorrow = new Date(today)
    tomorrow.setDate(tomorrow.getDate() + 1)

    for (const target of config.dailySummaryTargets) {
      const records = await ctx.database.get('monitorluna_activity', {
        deviceId: target.deviceId,
        timestamp: { $gte: today, $lt: tomorrow }
      })
      if (records.length === 0) continue

      const html = await buildSummaryHtml(records, target.deviceId, today, tomorrow)
      const puppeteer = (ctx as any)['puppeteer']
      if (!puppeteer) {
        ctx.logger.warn('[monitorluna] puppeteer 服务未启用，跳过每日总结')
        continue
      }
      const page = await puppeteer.page()
      await page.setContent(html)
      const body = await page.$('body')
      const clip = body ? await body.boundingBox() : null
      const buf = await page.screenshot({ clip, type: 'jpeg', quality: 90 })
      await page.close()
      const filename = `summary_${target.deviceId}_${Date.now()}.jpg`
      const { url } = await storage.upload(buf, filename)

      for (const channelId of target.channelIds) {
        const [platform, selfId, gid] = channelId.split(':')
        await ctx.bots[`${platform}:${selfId}`]?.sendMessage(gid, h.image(url))
      }
    }
  }

  async function buildSummaryHtml(records: ActivityLog[], deviceId: string, today: Date, tomorrow: Date): Promise<string> {
    const date = new Date().toLocaleDateString('zh-CN')

    // 格式化应用名称（去除 .exe）
    const formatAppName = (process: string) => process.replace(/\.exe$/i, '')

    // 查询输入统计数据（今天0点到现在）
    const inputStats = await ctx.database.get('monitorluna_input_stats', {
      deviceId,
      timestamp: { $gte: today, $lt: new Date() }
    })

    // 查询浏览器活动数据（今天）
    const browserRecords = await ctx.database.get('monitorluna_browser_activity', {
      deviceId,
      timestamp: { $gte: today, $lt: new Date() }
    })
    // 按域名聚合（当天可能有多条记录）
    const browserDomainMap = new Map<string, number>() // domain -> total seconds
    for (const r of browserRecords) {
      browserDomainMap.set(r.domain, (browserDomainMap.get(r.domain) || 0) + r.duration)
    }
    // 浏览器进程名称集合
    const BROWSER_PROCESSES = new Set(['chrome.exe', 'msedge.exe', 'firefox.exe', 'brave.exe', 'opera.exe'])

    // 构建图标映射
    const iconMap = new Map<string, string>()
    for (const record of inputStats) {
      if (record.iconBase64 && !iconMap.has(record.process)) {
        iconMap.set(record.process, record.iconBase64)
      }
    }

    // 聚合输入统计
    const appInputStats = new Map<string, { displayName: string, icon: string, keyPresses: number, clicks: number, scrollDistance: number }>()
    for (const record of inputStats) {
      const existing = appInputStats.get(record.process) || {
        displayName: record.displayName,
        icon: record.iconBase64,
        keyPresses: 0,
        clicks: 0,
        scrollDistance: 0
      }
      existing.keyPresses += record.keyPresses
      existing.clicks += record.leftClicks + record.rightClicks
      existing.scrollDistance += record.scrollDistance
      appInputStats.set(record.process, existing)
    }

    // TOP 6（按总输入量排序）
    const topInputApps = [...appInputStats.entries()]
      .sort((a, b) => (b[1].keyPresses + b[1].clicks) - (a[1].keyPresses + a[1].clicks))
      .slice(0, 6)

    // 计算最大值用于进度条
    const maxKeys = Math.max(...topInputApps.map(([, s]) => s.keyPresses), 1)
    const maxClicks = Math.max(...topInputApps.map(([, s]) => s.clicks), 1)
    const maxScroll = Math.max(...topInputApps.map(([, s]) => s.scrollDistance), 1)

    // 过滤 AFK 记录，计算实际活跃时间
    const activeRecords = records.filter(r => !r.afk)
    const totalActiveMs = records.reduce((sum, r) => {
      if (r.afk) return sum
      const end = r.endTime ? r.endTime.getTime() : r.timestamp.getTime()
      return sum + Math.max(0, end - r.timestamp.getTime())
    }, 0)
    const totalActiveMinutes = Math.round(totalActiveMs / 1000 / 60)

    // 按时间排序
    records.sort((a, b) => a.timestamp.getTime() - b.timestamp.getTime())

    // 按小时统计每个进程的活跃时长（分钟），使用 endTime - timestamp
    const hourlyStats = new Map<number, Map<string, number>>() // hour -> process -> minutes
    for (let h = 0; h < 24; h++) hourlyStats.set(h, new Map())

    for (const record of records) {
      if (record.afk) continue
      const end = record.endTime ? record.endTime.getTime() : record.timestamp.getTime()
      const duration = Math.max(0, end - record.timestamp.getTime()) / 1000 / 60 // 分钟
      if (duration <= 0) continue
      const hour = record.timestamp.getHours()
      const processMap = hourlyStats.get(hour)!
      processMap.set(record.process, (processMap.get(record.process) || 0) + duration)
    }

    // 计算每小时的总活跃时长（分钟，用于柱状图）
    const hourlyActivity = new Map<number, number>()
    for (let h = 0; h < 24; h++) hourlyActivity.set(h, 0)
    for (const record of records) {
      if (record.afk) continue
      const end = record.endTime ? record.endTime.getTime() : record.timestamp.getTime()
      const duration = Math.max(0, end - record.timestamp.getTime()) / 1000 / 60
      const hour = record.timestamp.getHours()
      hourlyActivity.set(hour, (hourlyActivity.get(hour) || 0) + duration)
    }

    // 每小时活跃时间最长的 4 个进程
    const hourlyTop4 = new Map<number, Array<[string, number]>>()
    for (const [hour, processMap] of hourlyStats.entries()) {
      const top4 = [...processMap.entries()]
        .sort((a, b) => b[1] - a[1])
        .slice(0, 4)
      hourlyTop4.set(hour, top4)
    }

    // 全天活跃时间最长的 4 个进程
    const totalDuration = new Map<string, number>()
    for (const processMap of hourlyStats.values()) {
      for (const [process, duration] of processMap) {
        totalDuration.set(process, (totalDuration.get(process) || 0) + duration)
      }
    }
    const topDuration = [...totalDuration.entries()]
      .sort((a, b) => b[1] - a[1])
      .slice(0, 4)

    const maxActivity = Math.max(...hourlyActivity.values(), 1)
    const startHour = records.length > 0 ? Math.min(...records.map(r => r.timestamp.getHours())) : 0
    const endHour = records.length > 0 ? Math.max(...records.map(r => r.timestamp.getHours())) : 23
    const afkMinutes = Math.round(records.reduce((sum, r) => {
      if (!r.afk) return sum
      const end = r.endTime ? r.endTime.getTime() : r.timestamp.getTime()
      return sum + Math.max(0, end - r.timestamp.getTime())
    }, 0) / 1000 / 60)

    return `<!DOCTYPE html>
<html><head><meta charset="utf-8">
<link href="https://fonts.googleapis.com/css2?family=ZCOOL+KuaiLe&family=Patrick+Hand&family=Noto+Sans+SC:wght@400;500;700&display=swap" rel="stylesheet">
<style>
:root{--bg-paper:#fdfbf7;--ink-primary:#5d4037;--ink-secondary:#8d6e63;--color-yellow:#fff9c4;--color-pink:#ffccbc;--color-blue:#b3e5fc;--color-green:#c8e6c9;--accent-orange:#ff7043;--font-title:'ZCOOL KuaiLe',cursive;--font-hand:'Patrick Hand',"KaiTi",serif;--font-body:'Noto Sans SC',sans-serif}
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:var(--font-body);color:var(--ink-primary);background-color:var(--bg-paper);background-image:radial-gradient(#ddd 2px,transparent 2px);background-size:20px 20px;min-height:100vh;padding:40px 20px;line-height:1.6}
.container{max-width:900px;margin:0 auto;background:#fff;border:2px solid var(--ink-primary);border-radius:20px;padding:40px;box-shadow:8px 8px 0 var(--color-blue),16px 16px 0 var(--color-pink),0 20px 40px rgba(0,0,0,0.1)}
.header{text-align:center;margin-bottom:40px;position:relative;padding-top:20px}
.title-sticker{display:inline-block;background:#fff;padding:20px 40px;border:3px dashed var(--ink-primary);border-radius:15px;box-shadow:5px 5px 0 var(--color-blue);transform:rotate(-2deg);position:relative}
.title-sticker h1{font-family:var(--font-title);font-size:2.5rem;color:var(--accent-orange);margin:0;-webkit-text-stroke:1px var(--ink-primary)}
.date-badge{position:absolute;bottom:-15px;right:-20px;background:var(--color-yellow);padding:5px 15px;font-family:var(--font-hand);font-size:1.1rem;box-shadow:2px 2px 3px rgba(0,0,0,0.1);transform:rotate(5deg);border:1px solid var(--ink-primary)}
.tape{position:absolute;top:-15px;left:50%;transform:translateX(-50%);width:100px;height:20px;background:rgba(255,171,145,0.7);opacity:0.8}
.section{margin-bottom:35px}
.section-title{font-family:var(--font-title);font-size:1.5rem;margin-bottom:15px;color:var(--accent-orange);display:flex;align-items:center;gap:8px}
.chart-box{background:#fff;padding:20px;border:2px solid var(--ink-primary);border-radius:12px;box-shadow:4px 4px 0 var(--color-green)}
.chart{display:flex;align-items:flex-end;height:180px;gap:3px;padding:10px 0}
.bar{flex:1;background:var(--accent-orange);border-radius:4px 4px 0 0;position:relative;min-height:2px}
.bar-label{position:absolute;bottom:-22px;left:50%;transform:translateX(-50%);font-size:10px;color:var(--ink-secondary);white-space:nowrap}
.list-box{background:#fff;padding:20px;border:2px solid var(--ink-primary);border-radius:12px;box-shadow:4px 4px 0 var(--color-pink)}
.list-item{display:flex;justify-content:space-between;padding:10px 0;border-bottom:1px dashed #e0e0e0}
.list-item:last-child{border-bottom:none}
.item-name{font-weight:600;color:var(--ink-primary);flex:1}
.item-value{color:var(--accent-orange);font-weight:600;margin-left:10px}
.hour-section{margin-bottom:20px;padding:12px;border:1px dashed var(--ink-secondary);border-radius:8px;background:#fafafa}
.hour-label{font-weight:700;color:var(--ink-secondary);margin-bottom:6px;font-size:0.95rem}
.hour-list{display:flex;flex-wrap:wrap;gap:6px}
.hour-tag{background:var(--color-blue);padding:3px 0;border-radius:12px;font-size:0.85rem;color:var(--ink-primary);display:inline-flex;overflow:hidden}
.hour-tag-name{background:var(--color-blue);padding:3px 10px;color:var(--ink-primary)}
.hour-tag-time{background:#fff;padding:3px 10px;color:var(--ink-primary);font-weight:600}
.input-stats-item{padding:8px 0;border-bottom:1px dashed #e0e0e0}
.input-stats-item:last-child{border-bottom:none}
.app-row{display:grid;grid-template-columns:1fr 160px 120px 120px;align-items:center;gap:8px}
.app-info{display:flex;align-items:center;gap:6px;overflow:hidden}
.app-icon{width:16px;height:16px;border-radius:3px;flex-shrink:0}
.app-name{font-weight:600;color:var(--ink-primary);font-size:0.85rem;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.bar-col{display:flex;flex-direction:column;gap:2px}
.bar-track{background:#f0f0f0;border-radius:4px;height:16px;position:relative;overflow:hidden}
.bar-fill-keys{background:#5c9bd6;border-radius:4px;height:100%;position:absolute;left:0;top:0}
.bar-fill-clicks{background:#70b870;border-radius:4px;height:100%;position:absolute;left:0;top:0}
.bar-fill-scroll{background:#d4a843;border-radius:4px;height:100%;position:absolute;left:0;top:0}
.bar-text{position:absolute;right:4px;top:50%;transform:translateY(-50%);font-size:10px;font-weight:600;color:#fff;white-space:nowrap}
.bar-text-outside{font-size:10px;color:var(--ink-secondary);text-align:right;margin-top:1px}
.col-header{font-size:11px;font-weight:700;color:var(--ink-secondary);text-align:center}
.stats-header{display:grid;grid-template-columns:1fr 160px 120px 120px;gap:8px;margin-bottom:8px;padding-bottom:6px;border-bottom:2px solid #e0e0e0}
.footer{text-align:center;margin-top:40px;font-family:var(--font-hand);color:var(--ink-secondary);font-size:0.9rem}
</style>
</head><body><div class="container">
<div class="header">
<div class="title-sticker">
<div class="tape"></div>
<h1>今日活动总结</h1>
<div class="date-badge">${date}</div>
</div>
</div>
<div style="text-align:center;color:var(--ink-secondary);margin-bottom:30px;font-family:var(--font-hand)">设备: ${deviceId} · 统计时段: ${startHour}:00 - ${endHour}:59 · 实际活跃: ${totalActiveMinutes}分钟 · AFK: ${afkMinutes}分钟</div>
<div class="section">
<div class="section-title">📊 24H 活跃轨迹</div>
<div class="chart-box">
<div class="chart">
${Array.from(hourlyActivity.entries()).map(([hour, count]) => {
      const height = count > 0 ? (count / maxActivity * 100) : 2
      return `<div class="bar" style="height:${height}%"><div class="bar-label">${hour}</div></div>`
    }).join('')}
</div>
</div>
</div>
<div class="section">
<div class="section-title">⌨️ 输入统计 TOP 6</div>
<div class="list-box">
${topInputApps.length > 0 ? `
<div class="stats-header">
  <div class="col-header" style="text-align:left">应用</div>
  <div class="col-header">键盘</div>
  <div class="col-header">鼠标</div>
  <div class="col-header">滚轮</div>
</div>
${topInputApps.map(([process, stats], idx) => {
      const keysW = Math.round(stats.keyPresses / maxKeys * 100)
      const clicksW = Math.round(stats.clicks / maxClicks * 100)
      const scrollW = Math.round(stats.scrollDistance / maxScroll * 100)
      const keysInside = keysW > 30
      const clicksInside = clicksW > 30
      const scrollInside = scrollW > 30
      return `<div class="input-stats-item">
<div class="app-row">
  <div class="app-info">
    <span style="color:var(--ink-secondary);font-size:0.8rem;min-width:16px">${idx + 1}</span>
    ${stats.icon ? `<img src="data:image/png;base64,${stats.icon}" class="app-icon">` : '<div style="width:16px"></div>'}
    <span class="app-name">${formatAppName(process)}</span>
  </div>
  <div class="bar-col">
    <div class="bar-track"><div class="bar-fill-keys" style="width:${keysW}%">${keysInside ? `<span class="bar-text">${stats.keyPresses.toLocaleString()}</span>` : ''}</div>${!keysInside ? `<span style="position:absolute;right:4px;top:50%;transform:translateY(-50%);font-size:10px;font-weight:600;color:#888">${stats.keyPresses.toLocaleString()}</span>` : ''}</div>
  </div>
  <div class="bar-col">
    <div class="bar-track"><div class="bar-fill-clicks" style="width:${clicksW}%">${clicksInside ? `<span class="bar-text">${stats.clicks.toLocaleString()}</span>` : ''}</div>${!clicksInside ? `<span style="position:absolute;right:4px;top:50%;transform:translateY(-50%);font-size:10px;font-weight:600;color:#888">${stats.clicks.toLocaleString()}</span>` : ''}</div>
  </div>
  <div class="bar-col">
    <div class="bar-track"><div class="bar-fill-scroll" style="width:${scrollW}%">${scrollInside ? `<span class="bar-text">${Math.round(stats.scrollDistance).toLocaleString()}</span>` : ''}</div>${!scrollInside ? `<span style="position:absolute;right:4px;top:50%;transform:translateY(-50%);font-size:10px;font-weight:600;color:#888">${Math.round(stats.scrollDistance).toLocaleString()}</span>` : ''}</div>
  </div>
</div>
</div>`
    }).join('')}` : '<div style="color:var(--ink-secondary);text-align:center;padding:20px">暂无输入统计数据</div>'}
</div>
</div>
<div class="section">
<div class="section-title">🕐 每小时 TOP 4</div>
<div class="chart-box">
${Array.from(hourlyTop4.entries()).filter(([, top]) => top.length > 0).map(([hour, top]) => `<div class="hour-section">
<div class="hour-label">${hour}:00 - ${hour}:59</div>
<div class="hour-list">
${top.map(([app, dur]) => {
      const icon = iconMap.get(app)
      const iconHtml = icon ? `<img src="data:image/png;base64,${icon}" style="width:14px;height:14px;margin-right:3px;vertical-align:middle">` : ''
      return `<span class="hour-tag"><span class="hour-tag-name">${iconHtml}${formatAppName(app)}</span><span class="hour-tag-time">${Math.round(dur)}m</span></span>`
    }).join('')}
</div>
</div>`).join('')}
</div>
</div>
${browserDomainMap.size > 0 ? `<div class="section">
<div class="section-title">🌐 浏览器域名时长 TOP 5</div>
<div class="list-box">
${[...browserDomainMap.entries()]
  .sort((a, b) => b[1] - a[1])
  .slice(0, 5)
  .map(([domain, secs], i) => {
    const mins = Math.round(secs / 60)
    return `<div class="list-item"><span class="item-name">${i + 1}. ${domain}</span><span class="item-value">${mins} 分钟</span></div>`
  }).join('')}
</div>
</div>` : ''}
<div class="footer">Generated by MonitorLuna · Scrapbook Theme</div>
</div></body></html>`
  }

  // Bot commands
  const monitor = ctx.command('monitor', '远程设备监控')

  monitor.subcommand('.list', '列出所有在线设备')
    .action(() => {
      if (devices.size === 0) return '当前没有在线设备'
      return '在线设备：\n' + [...devices.keys()].map(id => `• ${id}`).join('\n')
    })

  monitor.subcommand('.screen <device:string>', '截取远程设备屏幕')
    .action(async ({ session }, device) => {
      if (!device) return '请指定设备名称'
      try {
        const data = await sendCommand(device, 'screenshot')
        const buf = Buffer.from(data, 'base64')
        return h.image(buf, 'image/jpeg')
      } catch (e) {
        return `截图失败: ${e instanceof Error ? e.message : String(e)}`
      }
    })

  monitor.subcommand('.window <device:string>', '截取远程设备当前活跃窗口')
    .action(async ({ session }, device) => {
      if (!device) return '请指定设备名称'
      try {
        const data = await sendCommand(device, 'window_screenshot')
        const buf = Buffer.from(data, 'base64')
        return h.image(buf, 'image/jpeg')
      } catch (e) {
        return `窗口截图失败: ${e instanceof Error ? e.message : String(e)}`
      }
    })

  monitor.subcommand('.status <device:string>', '查看远程设备系统状态')
    .action(async ({ session }, device) => {
      if (!device) return '请指定设备名称'
      try {
        const data = await sendCommand(device, 'system_status')
        const info = JSON.parse(data)
        let result = `CPU: ${info.cpu_percent.toFixed(1)}%\n`
        result += `内存: ${info.memory.used} MB / ${info.memory.total} MB (${info.memory.percent.toFixed(1)}%)`
        if (info.gpu && info.gpu.length > 0) {
          for (const gpu of info.gpu) {
            result += `\nGPU: ${gpu.name}`
            if (gpu.load >= 0) result += `\n  负载: ${gpu.load.toFixed(1)}%`
            if (gpu.memory_used >= 0) result += `\n  显存: ${gpu.memory_used} MB / ${gpu.memory_total} MB`
            else if (gpu.memory_total > 0) result += `\n  显存: ${gpu.memory_total} MB`
          }
        } else {
          result += '\nGPU: 未检测到独立显卡'
        }
        return result
      } catch (e) {
        return `获取系统状态失败: ${e instanceof Error ? e.message : String(e)}`
      }
    })

  monitor.subcommand('.analytics <device:string>', '生成当天活动总结')
    .action(async ({ session }, device) => {
      if (!device) return '请指定设备名称'
      const today = new Date()
      today.setHours(0, 0, 0, 0)
      const tomorrow = new Date(today)
      tomorrow.setDate(tomorrow.getDate() + 1)

      const records = await ctx.database.get('monitorluna_activity', {
        deviceId: device,
        timestamp: { $gte: today, $lt: tomorrow }
      })
      if (config.debug) ctx.logger.info(`[monitorluna][debug] analytics: device=${device}, records=${records.length}`)
      if (records.length === 0) return '今日暂无活动记录'

      const puppeteer = (ctx as any)['puppeteer']
      if (config.debug) ctx.logger.info(`[monitorluna][debug] puppeteer service: ${puppeteer ? 'available' : 'not available'}`)
      if (!puppeteer) return '需要安装 puppeteer 插件才能生成总结图'

      try {
        const html = await buildSummaryHtml(records, device, today, tomorrow)
        if (config.debug) ctx.logger.info(`[monitorluna][debug] html length: ${html.length}`)

        const page = await puppeteer.page()
        await page.setContent(html)
        const body = await page.$('body')
        const clip = body ? await body.boundingBox() : null
        const buf = await page.screenshot({ clip, type: 'jpeg', quality: 90 })
        await page.close()

        if (config.debug) ctx.logger.info(`[monitorluna][debug] screenshot ok, buf size: ${buf.length}`)
        const filename = `summary_${device}_${Date.now()}.jpg`
        const { url } = await storage.upload(buf, filename)
        if (config.debug) ctx.logger.info(`[monitorluna][debug] upload ok, url: ${url}`)
        return h.image(url)
      } catch (e) {
        return `生成总结失败: ${e instanceof Error ? e.message : String(e)}`
      }
    })

  monitor.subcommand('.browser <device:string>', '查看今日浏览器域名时长排行')
    .action(async ({ session }, device) => {
      if (!device) return '请指定设备名称'
      const today = new Date()
      today.setHours(0, 0, 0, 0)
      const tomorrow = new Date(today)
      tomorrow.setDate(tomorrow.getDate() + 1)

      const records = await ctx.database.get('monitorluna_browser_activity', {
        deviceId: device,
        timestamp: { $gte: today, $lt: tomorrow }
      })
      if (records.length === 0) return '今日暂无浏览器活动记录'

      const sorted = records.sort((a, b) => b.duration - a.duration).slice(0, 20)
      let result = `今日浏览器活动 TOP ${sorted.length}：\n`
      sorted.forEach((r, i) => {
        const mins = Math.round(r.duration / 60)
        result += `${i + 1}. ${r.domain} - ${mins}分钟\n`
      })
      return result
    })
}
