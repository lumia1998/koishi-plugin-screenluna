# koishi-plugin-monitorluna

[![npm](https://img.shields.io/npm/v/koishi-plugin-monitorluna?style=flat-square)](https://www.npmjs.com/package/koishi-plugin-monitorluna)

远程设备监控插件，通过 WebSocket 连接 Windows 后台 Agent，实现远程截图、系统监控、窗口活动追踪和每日总结生成。

## ✨ 功能特性

- 📸 **远程截图**：全屏截图 / 当前活跃窗口截图
- 📊 **系统监控**：实时查看 CPU、内存、GPU 使用率
- 🪟 **窗口追踪**：自动记录窗口切换，每 2 秒检测前台应用变化
- 📅 **每日总结**：生成手绘风格活动总结图，统计每小时和全天的应用使用时长
- 🖼️ **定时截图**：每 15 分钟自动截图并存储
- 💾 **多存储后端**：支持本地存储 / WebDAV / S3
- 🔐 **Token 鉴权**：WebSocket 连接需要 Token 验证

## 📦 安装

### Koishi 插件安装

在 Koishi 插件市场搜索 `monitorluna` 安装，或使用命令：

```bash
npm install koishi-plugin-monitorluna
```

### Windows Agent 安装

#### 方式一：一键启动（推荐）

1. 前往 [Releases](https://github.com/lumia1998/koishi-plugin-monitorluna/releases/latest) 下载 `monitorluna-agent.zip`
2. 解压到任意目录
3. 双击运行 `start-server.bat`
4. 脚本会自动检测并安装 Python 依赖
5. 启动后会在系统托盘显示图标，右键点击"打开设置"配置连接信息

#### 方式二：手动安装

```bash
pip install websockets pyautogui pillow psutil pystray aiohttp pywin32 wmi GPUtil
python screenshot-server.py
```

## ⚙️ Agent 配置

Agent 启动后访问 http://127.0.0.1:6315 打开配置页面：

| 配置项 | 说明 | 示例 |
|--------|------|------|
| Koishi WebSocket URL | Koishi 服务的 WebSocket 地址 | `ws://127.0.0.1:5140/monitorluna` |
| Token | 与 Koishi 插件配置的 Token 一致 | `admin` |
| Device ID | 设备标识符，用于区分多台设备 | `my-pc`、`work-laptop` |

配置保存后 Agent 会自动重新连接。系统托盘图标显示当前连接状态。

## 🔧 Koishi 插件配置

### 基础配置

| 配置项 | 默认值 | 说明 |
|--------|--------|------|
| token | - | 鉴权密钥，需与 Agent 一致 |
| serverPath | - | Koishi 公网地址（本地存储时用于生成图片 URL） |
| commandTimeout | 15000 | 指令超时时间（ms） |
| debug | false | 开启调试日志 |

### 存储配置

**本地存储（默认）：**
- `storagePath`：存储目录，默认 `data/monitorluna`
- `serverPath`：Koishi 公网访问地址，用于生成图片链接

**WebDAV：**
- `webdavEndpoint`：WebDAV 端点地址
- `webdavUsername` / `webdavPassword`：认证信息
- `webdavPublicUrl`：公网访问地址

**S3：**
- `s3Endpoint` / `s3Bucket` / `s3Region`
- `s3AccessKeyId` / `s3SecretAccessKey`
- `s3PublicUrl`：公网访问地址
- `s3PathStyle`：是否使用路径风格 URL

### 每日总结配置

```yaml
monitorluna:
  dailySummaryEnabled: true
  dailySummaryTime: '22:00'
  dailySummaryTargets:
    - deviceId: my-pc
      channelIds:
        - 'onebot:123456:987654321'
```

## 🤖 命令

| 命令 | 说明 |
|------|------|
| `monitor.list` | 列出所有在线设备 |
| `monitor.screen <设备ID>` | 截取设备全屏截图 |
| `monitor.window <设备ID>` | 截取设备当前活跃窗口 |
| `monitor.status <设备ID>` | 查看设备 CPU/内存/GPU 状态 |
| `monitor.analytics <设备ID>` | 生成当天活动总结图（需要 puppeteer 插件） |

## 📊 每日总结说明

总结图包含三个部分：

1. **24H 活跃轨迹**：柱状图显示每小时的活跃度
2. **全天活跃时间 TOP 4**：全天活跃时长最长的 4 个应用
3. **每小时 TOP 4**：每个小时内活跃时长最长的 4 个应用

统计逻辑：通过记录相邻两次窗口切换的时间差来计算各应用的使用时长，从当天实际开机使用时刻开始统计。

需要安装 `@koishijs/plugin-puppeteer` 插件才能渲染总结图。

## 📁 项目结构

```
koishi-plugin-monitorluna/
├── src/
│   └── index.ts          # Koishi 插件主文件
├── screenshot-server.py  # Windows Agent
├── start-server.bat      # Windows 一键启动脚本
└── readme.md
```

## 📝 License

MIT

