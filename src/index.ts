import { Context, Schema, h } from 'koishi'
import { IncomingMessage } from 'http'
import { WebSocket } from 'ws'
import '@koishijs/plugin-server'
import * as fs from 'fs/promises'
import * as path from 'path'
import * as crypto from 'crypto'

export const name = 'monitorluna'

export const usage = `
## MonitorLuna 使用说明

### 下载 Windows Agent

[点击下载 monitorluna-agent.zip](https://github.com/lumia1998/koishi-plugin-monitorluna/releases/download/v1.1.1/monitorluna-agent.zip)

解压后双击 \`start-silent.vbs\` 启动，然后访问 http://127.0.0.1:6315 配置连接信息。

### 命令

| 命令 | 说明 |
|------|------|
| \`monitor.list\` | 列出所有在线设备 |
| \`monitor.screen <设备ID>\` | 截取设备全屏截图 |
| \`monitor.window <设备ID>\` | 截取设备当前活跃窗口 |
| \`monitor.status <设备ID>\` | 查看设备 CPU/内存/GPU 状态 |
| \`monitor.analytics <设备ID>\` | 生成当天活动总结图（需要 puppeteer 插件） |

`

export const inject = {
  required: ['server', 'database'],
  optional: ['puppeteer']
}

// ── Database Tables ──
declare module 'koishi' {
  interface Tables {
    monitorluna_screenshot: Screenshot
  }
}

interface Screenshot {
  id: number
  deviceId: string
  url: string
  type: 'manual' | 'push' | 'daily' | 'analytics'
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
  private dir: string

  constructor(private ctx: Context, private config: Config) {
    this.dir = path.join(ctx.baseDir, config.storagePath || 'data/monitorluna')
  }

  async init() {
    await fs.mkdir(this.dir, { recursive: true })

    // Mount HTTP route
    this.ctx.server.get('/monitorluna/:filename', async (ctx) => {
      const filename = ctx.params.filename
      if (filename.includes('..') || filename.includes('/') || filename.includes('\\')) {
        ctx.status = 400
        return
      }
      try {
        const file = path.join(this.dir, filename)
        const buf = await fs.readFile(file)
        ctx.type = 'image/jpeg'
        ctx.body = buf
      } catch {
        ctx.status = 404
      }
    })
  }

  async upload(buffer: Buffer, filename: string) {
    const filepath = path.join(this.dir, filename)
    await fs.writeFile(filepath, buffer)
    const baseUrl = this.config.serverPath || this.ctx.root.config.selfUrl || 'http://127.0.0.1:5140'
    return { key: filename, url: `${baseUrl.replace(/\/$/, '')}/monitorluna/${filename}` }
  }

  async delete(key: string) {
    await fs.unlink(path.join(this.dir, key)).catch(() => { })
  }

  async cleanup(days: number) {
    const cutoff = Date.now() - days * 86400000
    try {
      const files = await fs.readdir(this.dir)
      for (const file of files) {
        if (file === '.keep') continue
        const stat = await fs.stat(path.join(this.dir, file))
        if (stat.mtimeMs < cutoff) await this.delete(file)
      }
    } catch (e) {
      this.ctx.logger.warn(`[monitorluna] 本地清理失败: ${e instanceof Error ? e.message : String(e)}`)
    }
  }
}

// ── WebDAV Storage ──
class WebDAVStorage implements StorageBackend {
  constructor(private config: Config, private ctx: Context) { }

  async init() { }

  async upload(buffer: Buffer, filename: string) {
    const endpoint = this.config.webdavEndpoint!
    const url = `${endpoint.replace(/\/$/, '')}/${filename}`
    const auth = Buffer.from(`${this.config.webdavUsername}:${this.config.webdavPassword}`).toString('base64')
    const res = await fetch(url, {
      method: 'PUT',
      headers: { 'Authorization': `Basic ${auth}` },
      body: buffer as unknown as BodyInit
    })
    if (!res.ok) throw new Error(`WebDAV upload failed: ${res.status} ${res.statusText}`)
    const publicUrl = this.config.webdavPublicUrl || endpoint
    return { key: filename, url: `${publicUrl.replace(/\/$/, '')}/${filename}` }
  }

  async delete(key: string) {
    const endpoint = this.config.webdavEndpoint!
    const url = `${endpoint.replace(/\/$/, '')}/${key}`
    const auth = Buffer.from(`${this.config.webdavUsername}:${this.config.webdavPassword}`).toString('base64')
    await fetch(url, { method: 'DELETE', headers: { 'Authorization': `Basic ${auth}` } }).catch(() => { })
  }

  async cleanup(days: number) {
    const cutoff = Date.now() - days * 86400000
    try {
      const auth = Buffer.from(`${this.config.webdavUsername}:${this.config.webdavPassword}`).toString('base64')
      const res = await fetch(this.config.webdavEndpoint!, {
        method: 'PROPFIND',
        headers: {
          'Authorization': `Basic ${auth}`,
          'Depth': '1'
        }
      })
      if (!res.ok) return

      const xml = await res.text()
      const hrefMatches = xml.matchAll(/<D:href>([^<]+)<\/D:href>/g)
      const modifiedMatches = xml.matchAll(/<D:getlastmodified>([^<]+)<\/D:getlastmodified>/g)

      const hrefs = [...hrefMatches].map(m => m[1]).slice(1) // Skip root
      const dates = [...modifiedMatches].map(m => new Date(m[1]).getTime())

      for (let i = 0; i < hrefs.length; i++) {
        if (dates[i] < cutoff) {
          const filename = hrefs[i].split('/').pop()
          if (filename) await this.delete(filename)
        }
      }
    } catch (e) {
      this.ctx.logger.warn(`[monitorluna] WebDAV 清理失败: ${e instanceof Error ? e.message : String(e)}`)
    }
  }
}

// ── S3 Storage ──
class S3Storage implements StorageBackend {
  constructor(private config: Config, private ctx: Context) { }

  async init() { }

  async upload(buffer: Buffer, filename: string) {
    const url = this.getUrl(filename)
    const headers = await this.signRequestV4('PUT', filename, buffer)
    const res = await fetch(url, { method: 'PUT', headers, body: buffer as unknown as BodyInit })
    if (!res.ok) throw new Error(`S3 upload failed: ${res.status} ${res.statusText}`)
    if (this.config.s3PublicUrl) {
      return { key: filename, url: this.config.s3PublicUrl.replace(/\/$/, '') + '/' + filename }
    }
    return { key: filename, url }
  }

  async delete(key: string) {
    const url = this.getUrl(key)
    const headers = await this.signRequestV4('DELETE', key)
    await fetch(url, { method: 'DELETE', headers }).catch(() => { })
  }

  async cleanup(days: number) {
    const cutoff = Date.now() - days * 86400000
    try {
      const listUrl = this.config.s3PathStyle
        ? `${this.config.s3Endpoint!.replace(/\/$/, '')}/${this.config.s3Bucket}?list-type=2`
        : `${this.config.s3Endpoint!.replace('://', `://${this.config.s3Bucket}.`).replace(/\/$/, '')}/?list-type=2`

      const headers = await this.signRequestV4('GET', '', undefined, { 'list-type': '2' })
      const res = await fetch(listUrl, { method: 'GET', headers })
      if (!res.ok) return

      const xml = await res.text()
      const keyMatches = xml.matchAll(/<Key>(.*?)<\/Key>/g)
      const modifiedMatches = xml.matchAll(/<LastModified>(.*?)<\/LastModified>/g)

      const keys = [...keyMatches].map(m => m[1])
      const dates = [...modifiedMatches].map(m => new Date(m[1]).getTime())

      for (let i = 0; i < keys.length; i++) {
        if (dates[i] < cutoff) await this.delete(keys[i])
      }
    } catch (e) {
      this.ctx.logger.warn(`[monitorluna] S3 清理失败: ${e instanceof Error ? e.message : String(e)}`)
    }
  }

  private getUrl(key: string) {
    const { s3Endpoint, s3Bucket, s3PathStyle } = this.config
    if (!s3Endpoint) throw new Error('s3Endpoint is required')
    if (s3PathStyle) return `${s3Endpoint.replace(/\/$/, '')}/${s3Bucket}/${key}`
    return `${s3Endpoint.replace('://', `://${s3Bucket}.`).replace(/\/$/, '')}/${key}`
  }

  private async signRequestV4(method: string, key: string, body?: Buffer, queryParams?: Record<string, string>) {
    const now = new Date()
    const amzDate = now.toISOString().replace(/[:-]|\.\d{3}/g, '')
    const dateStamp = amzDate.slice(0, 8)
    const region = this.config.s3Region || 'us-east-1'
    const service = 's3'

    const endpointUrl = new URL(this.config.s3Endpoint!)
    const host = this.config.s3PathStyle
      ? endpointUrl.host
      : `${this.config.s3Bucket}.${endpointUrl.host}`

    const canonicalUri = this.config.s3PathStyle ? `/${this.config.s3Bucket}/${key}` : `/${key}`
    const canonicalQuerystring = queryParams
      ? Object.entries(queryParams).sort().map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`).join('&')
      : ''

    const payloadHash = crypto.createHash('sha256').update(body || '').digest('hex')
    const canonicalHeaders = `host:${host}\nx-amz-content-sha256:${payloadHash}\nx-amz-date:${amzDate}\n`
    const signedHeaders = 'host;x-amz-content-sha256;x-amz-date'

    const canonicalRequest = `${method}\n${canonicalUri}\n${canonicalQuerystring}\n${canonicalHeaders}\n${signedHeaders}\n${payloadHash}`
    const canonicalRequestHash = crypto.createHash('sha256').update(canonicalRequest).digest('hex')

    const credentialScope = `${dateStamp}/${region}/${service}/aws4_request`
    const stringToSign = `AWS4-HMAC-SHA256\n${amzDate}\n${credentialScope}\n${canonicalRequestHash}`

    const kDate = crypto.createHmac('sha256', `AWS4${this.config.s3SecretAccessKey}`).update(dateStamp).digest()
    const kRegion = crypto.createHmac('sha256', kDate).update(region).digest()
    const kService = crypto.createHmac('sha256', kRegion).update(service).digest()
    const kSigning = crypto.createHmac('sha256', kService).update('aws4_request').digest()
    const signature = crypto.createHmac('sha256', kSigning).update(stringToSign).digest('hex')

    const authorization = `AWS4-HMAC-SHA256 Credential=${this.config.s3AccessKeyId}/${credentialScope}, SignedHeaders=${signedHeaders}, Signature=${signature}`

    return {
      'Authorization': authorization,
      'x-amz-date': amzDate,
      'x-amz-content-sha256': payloadHash,
      'Content-Type': 'image/jpeg'
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
  pushChannelIds: string[]
  pushPrivateIds: string[]
  pushPollInterval: number
  dailySummaryEnabled: boolean
  dailySummaryTime: string
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
  }),
  Schema.object({
    pushChannelIds: Schema.array(Schema.string()).default([]).description('推送目标群组（格式: platform:selfId:channelId）'),
    pushPrivateIds: Schema.array(Schema.string()).default([]).description('推送目标私聊（格式: platform:selfId:userId）'),
    pushPollInterval: Schema.number().default(10000).description('应用切换轮询间隔（毫秒）'),
    dailySummaryEnabled: Schema.boolean().default(false).description('启用每日总结推送'),
    dailySummaryTime: Schema.string().default('22:00').description('每日总结推送时间（HH:mm）'),
  }).description('推送设置'),
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
  ctx.model.extend('monitorluna_screenshot', {
    id: 'unsigned',
    deviceId: 'string',
    url: 'string',
    type: 'string',
    timestamp: 'timestamp'
  }, { autoInc: true })

  // Initialize storage
  if (config.storageType === 'local') storage = new LocalStorage(ctx, config)
  else if (config.storageType === 'webdav') storage = new WebDAVStorage(config, ctx)
  else storage = new S3Storage(config, ctx)

  const timers: ReturnType<typeof setInterval>[] = []

  function clearPendingCommands(device: DeviceConnection, error: Error) {
    for (const pending of device.pendingCommands.values()) {
      clearTimeout(pending.timer)
      pending.reject(error)
    }
    device.pendingCommands.clear()
  }

  async function saveScreenshotRecord(deviceId: string, url: string, type: Screenshot['type']) {
    await ctx.database.create('monitorluna_screenshot', {
      deviceId,
      url,
      type,
      timestamp: new Date()
    })
  }

  async function renderSummaryImage(puppeteer: any, html: string): Promise<Buffer> {
    const page = await puppeteer.page()
    try {
      await page.setContent(html)
      const body = await page.$('body')
      const clip = body ? await body.boundingBox() : null
      return await page.screenshot({ clip, type: 'jpeg', quality: 90 })
    } finally {
      await page.close().catch(() => { })
    }
  }

  ctx.on('ready', async () => {
    await storage.init()
    // Immediate cleanup on start
    storage.cleanup(config.imageRetentionDays).catch(() => { })
    timers.push(setInterval(() => {
      storage.cleanup(config.imageRetentionDays).catch(() => { })
    }, 86400000))

    startPushPolling()
    if (config.dailySummaryEnabled) startDailySummary()
  })

  ctx.on('dispose', () => {
    for (const timer of timers) clearInterval(timer)
    timers.length = 0
  })

  // WebSocket handler
  ctx.server.ws('/monitorluna', (ws: WebSocket, _req: IncomingMessage) => {
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
        const existing = devices.get(deviceId)
        if (existing && existing.ws !== ws) {
          clearPendingCommands(existing, new Error('设备连接已被新的会话替换'))
          existing.ws.close(1000, 'replaced by new session')
        }
        device = { ws, deviceId, pendingCommands: new Map() }
        devices.set(deviceId, device)
        ctx.logger.info(`[monitorluna] 设备上线: ${deviceId}`)
        ws.send(JSON.stringify({ type: 'hello_ack', message: 'connected' }))
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
      if (deviceId && device) {
        const isCurrentConnection = devices.get(deviceId)?.ws === ws
        if (isCurrentConnection) {
          devices.delete(deviceId)
          lastKnownApp.delete(deviceId)
          ctx.logger.info(`[monitorluna] 设备下线: ${deviceId}`)
        } else {
          ctx.logger.info(`[monitorluna] 设备旧连接关闭: ${deviceId}`)
        }
        clearPendingCommands(device, new Error('设备断开连接'))
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
      try {
        device.ws.send(JSON.stringify({ type: 'command', id, cmd }))
      } catch (error) {
        clearTimeout(timer)
        device.pendingCommands.delete(id)
        reject(error instanceof Error ? error : new Error(String(error)))
      }
    })
  }

  // ── Push Polling ──
  const lastKnownApp = new Map<string, { title: string; process: string }>()

  async function sendToTargets(content: string) {
    const targets = [...config.pushChannelIds.map(t => ({ t, type: 'channel' })), ...config.pushPrivateIds.map(t => ({ t, type: 'private' }))]
    for (const { t, type } of targets) {
      const [platform, selfId, id] = t.split(':')
      if (!platform || !selfId || !id) continue
      const bot = ctx.bots[`${platform}:${selfId}`]
      if (!bot) {
        ctx.logger.warn(`[monitorluna] 找不到机器人 ${platform}:${selfId}`)
        continue
      }
      try {
        if (type === 'channel') await bot.sendMessage(id, content)
        else await bot.sendPrivateMessage(id, content)
      } catch (e) {
        ctx.logger.warn(`[monitorluna] 发送失败到 ${t}: ${e instanceof Error ? e.message : String(e)}`)
      }
    }
  }

  function startPushPolling() {
    if (config.pushChannelIds.length === 0 && config.pushPrivateIds.length === 0) return

    const poll = async () => {
      for (const [deviceId] of devices) {
        try {
          const data = await sendCommand(deviceId, 'window_info')
          const info = JSON.parse(data)
          const prev = lastKnownApp.get(deviceId)
          const current = { title: info.title || '', process: info.process || '' }

          if (prev && (prev.process !== current.process || prev.title !== current.title)) {
            try {
              const screenshotData = await sendCommand(deviceId, 'window_screenshot')
              const buf = Buffer.from(screenshotData, 'base64')
              const filename = `push_${deviceId}_${Date.now()}.jpg`
              const { url } = await storage.upload(buf, filename)
              const processName = current.process.replace(/\.exe$/i, '')

              await saveScreenshotRecord(deviceId, url, 'push')

              await sendToTargets(`🖥️ ${deviceId} 切换到: ${processName}\n${current.title}\n${h.image(url)}`)
            } catch (e) {
              ctx.logger.debug(`[monitorluna] 推送截图失败: ${e instanceof Error ? e.message : String(e)}`)
            }
          }
          lastKnownApp.set(deviceId, current)
        } catch { }
      }
    }
    void poll()
    timers.push(setInterval(() => { void poll() }, config.pushPollInterval))
  }

  // ── Daily Summary ──
  function startDailySummary() {
    const check = async () => {
      const now = new Date()
      const currentTime = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`

      if (currentTime === config.dailySummaryTime) {
        const puppeteer = (ctx as any)['puppeteer']
        if (puppeteer) {
          const todayStart = new Date()
          todayStart.setHours(0, 0, 0, 0)
          const todayEnd = new Date()
          todayEnd.setHours(23, 59, 59, 999)

          for (const [deviceId] of devices) {
            try {
              // Persistence check via database
              const existing = await ctx.database.get('monitorluna_screenshot', {
                deviceId,
                type: 'daily',
                timestamp: { $gte: todayStart, $lte: todayEnd }
              })

              if (existing.length > 0) continue

              const dataStr = await sendCommand(deviceId, 'query_daily_data')
              const data = JSON.parse(dataStr)
              if (!data.activities || data.activities.length === 0) continue

              const html = await buildSummaryHtml(data, deviceId)
              const buf = await renderSummaryImage(puppeteer, html)

              const filename = `daily_${deviceId}_${Date.now()}.jpg`
              const { url } = await storage.upload(buf, filename)

              await saveScreenshotRecord(deviceId, url, 'daily')

              await sendToTargets(`📊 ${deviceId} 每日活动总结\n${h.image(url)}`)
            } catch (e) {
              ctx.logger.warn(`[monitorluna] 每日总结推送失败 (${deviceId}): ${e instanceof Error ? e.message : String(e)}`)
            }
          }
        }
      }
    }
    void check()
    timers.push(setInterval(() => { void check() }, 60000))
  }

  async function buildSummaryHtml(data: any, deviceId: string): Promise<string> {
    const date = new Date().toLocaleDateString('zh-CN')
    const formatAppName = (process: string) => process.replace(/\.exe$/i, '')

    const records = data.activities || []
    const inputStatsData = data.input_stats || {}
    const browserActivityData = data.browser_activity || {}
    const iconsData = data.icons || {}

    // 处理浏览器活动
    const topBrowserDomains: [string, number][] = (Object.entries(browserActivityData) as [string, number][])
      .sort((a, b) => b[1] - a[1])
      .slice(0, 8)

    // 构建图标映射
    const iconMap = new Map<string, string>()
    for (const [process, stats] of Object.entries(inputStatsData) as [string, any][]) {
      if (stats.icon_hash && iconsData[stats.icon_hash]) {
        iconMap.set(process, iconsData[stats.icon_hash])
      }
    }

    // 处理输入统计
    interface InputAppStats { displayName: string, iconHash: string, keyPresses: number, clicks: number, scrollDistance: number }
    const topInputApps: [string, InputAppStats][] = (Object.entries(inputStatsData) as [string, any][])
      .map(([process, stats]) => [process, {
        displayName: stats.display_name,
        iconHash: stats.icon_hash,
        keyPresses: stats.key_presses || 0,
        clicks: (stats.left_clicks || 0) + (stats.right_clicks || 0),
        scrollDistance: stats.scroll_distance || 0
      }] as [string, InputAppStats])
      .sort((a, b) => (b[1].keyPresses + b[1].clicks) - (a[1].keyPresses + a[1].clicks))
      .slice(0, 6)

    const maxKeys = Math.max(...topInputApps.map((a: any) => a[1].keyPresses), 1)
    const maxClicks = Math.max(...topInputApps.map((a: any) => a[1].clicks), 1)
    const maxScroll = Math.max(...topInputApps.map((a: any) => a[1].scrollDistance), 1)

    // 处理活动记录
    records.sort((a: any, b: any) => new Date(a.timestamp).getTime() - new Date(b.timestamp).getTime())

    const hourlyStats = new Map<number, Map<string, number>>()
    for (let h = 0; h < 24; h++) hourlyStats.set(h, new Map())

    for (let i = 0; i < records.length - 1; i++) {
      const curr = records[i]
      const next = records[i + 1]
      const duration = (new Date(next.timestamp).getTime() - new Date(curr.timestamp).getTime()) / 1000 / 60
      const hour = new Date(curr.timestamp).getHours()
      const processMap = hourlyStats.get(hour)!
      processMap.set(curr.process, (processMap.get(curr.process) || 0) + duration)
    }

    const hourlyActivity = new Map<number, number>()
    for (let h = 0; h < 24; h++) hourlyActivity.set(h, 0)
    records.forEach((r: any) => {
      const hour = new Date(r.timestamp).getHours()
      hourlyActivity.set(hour, (hourlyActivity.get(hour) || 0) + 1)
    })

    const hourlyTop4 = new Map<number, Array<[string, number]>>()
    for (const [hour, processMap] of hourlyStats.entries()) {
      const top4 = [...processMap.entries()].sort((a, b) => b[1] - a[1]).slice(0, 4)
      hourlyTop4.set(hour, top4)
    }

    const maxActivity = Math.max(...hourlyActivity.values(), 1)
    const startHour = records.length > 0 ? Math.min(...records.map((r: any) => new Date(r.timestamp).getHours())) : 0
    const endHour = records.length > 0 ? Math.max(...records.map((r: any) => new Date(r.timestamp).getHours())) : 23

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
<div style="text-align:center;color:var(--ink-secondary);margin-bottom:30px;font-family:var(--font-hand)">设备: ${deviceId} · 统计时段: ${startHour}:00 - ${endHour}:59</div>
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
      const iconData = iconMap.get(process)
      return `<div class="input-stats-item">
<div class="app-row">
  <div class="app-info">
    <span style="color:var(--ink-secondary);font-size:0.8rem;min-width:16px">${idx + 1}</span>
    ${iconData ? `<img src="data:image/png;base64,${iconData}" class="app-icon">` : '<div style="width:16px"></div>'}
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
${topBrowserDomains.length > 0 ? `<div class="section">
<div class="section-title">🌐 浏览器活动 TOP 8</div>
<div class="list-box">
${topBrowserDomains.map(([domain, seconds], idx) => {
      const minutes = Math.round(seconds / 60)
      const display = minutes >= 60 ? `${Math.floor(minutes / 60)}h${minutes % 60}m` : minutes > 0 ? `${minutes}m` : `${Math.round(seconds)}s`
      return `<div class="list-item"><span class="item-name">${idx + 1}. ${domain}</span><span class="item-value">${display}</span></div>`
    }).join('')}
</div>
</div>` : ''}
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
        const filename = `manual_${device}_${Date.now()}.jpg`
        const { url } = await storage.upload(buf, filename)
        await saveScreenshotRecord(device, url, 'manual')
        return h.image(url)
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
        const filename = `window_${device}_${Date.now()}.jpg`
        const { url } = await storage.upload(buf, filename)
        await saveScreenshotRecord(device, url, 'manual')
        return h.image(url)
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

      try {
        const dataStr = await sendCommand(device, 'query_daily_data')
        const data = JSON.parse(dataStr)

        if (!data.activities || data.activities.length === 0) {
          return '今日暂无活动记录'
        }

        const puppeteer = (ctx as any)['puppeteer']
        if (!puppeteer) return '需要安装 puppeteer 插件才能生成总结图'

        const html = await buildSummaryHtml(data, device)
        const buf = await renderSummaryImage(puppeteer, html)

        const filename = `manual_summary_${device}_${Date.now()}.jpg`
        const { url } = await storage.upload(buf, filename)

        await saveScreenshotRecord(device, url, 'analytics')

        return h.image(url)
      } catch (e) {
        return `生成总结失败: ${e instanceof Error ? e.message : String(e)}`
      }
    })
}
