# koishi-plugin-monitorluna

[![npm](https://img.shields.io/npm/v/koishi-plugin-monitorluna?style=flat-square)](https://www.npmjs.com/package/koishi-plugin-monitorluna)

远程设备监控插件，通过 WebSocket 连接本地 Agent，实现远程截图、系统状态查看、窗口活动追踪、浏览器活动统计和每日总结图片生成。

## 功能

- 远程截图：全屏截图和当前活跃窗口截图。
- 系统状态：查看 CPU、内存、GPU 使用率。
- 窗口追踪：轮询在线设备的前台窗口变化并推送截图。
- 浏览器统计：通过 `client/monitorluna.user.js` 上报域名活跃时长到本地 Agent。
- 每日总结：渲染活动总结图并支持定时推送。
- 多存储后端：支持本地存储、WebDAV 和 S3。
- 截图记录：截图 URL 和类型会写入 Koishi 数据库表 `monitorluna_screenshot`。

## 安装

### Koishi 插件

安装 `koishi-plugin-monitorluna`，并确保同时具备：

- `@koishijs/plugin-server`
- 数据库插件，例如 SQLite / MySQL / PostgreSQL
- `@koishijs/plugin-puppeteer`，仅在需要生成总结图时安装

### Windows Agent

Agent 相关文件现在统一放在 [`client/`](./client) 目录。

方式一：使用启动脚本

1. 打开 [`client/start-server.bat`](./client/start-server.bat) 启动 Agent。
2. 如果需要静默启动，运行 [`client/start-silent.vbs`](./client/start-silent.vbs)。
3. 首次启动会尝试使用 `uv` 运行 [`client/screenshot-server.py`](./client/screenshot-server.py)。

方式二：手动运行

```bash
cd client
pip install -r requirements.txt
python screenshot-server.py
```

也可以使用 [`client/pyproject.toml`](./client/pyproject.toml) 配合 `uv` 安装依赖。

## Agent 配置

启动后访问 `http://127.0.0.1:6315` 打开本地配置页面。

主要配置项：

- `Koishi WebSocket URL`：例如 `ws://127.0.0.1:5140/monitorluna`。
- `Token`：必须与 Koishi 插件配置一致。
- `Device ID`：设备唯一标识。
- `浏览器扩展密码`：给浏览器脚本连接 `ws://127.0.0.1:6315/ws/browser` 用，可留空。

Agent 会在 `client/monitorluna.db` 中保存活动记录、输入统计、浏览器统计和图标缓存。

## 浏览器统计

当前仓库内保留的是油猴脚本版本：[`client/monitorluna.user.js`](./client/monitorluna.user.js)。

使用方式：

1. 安装 Tampermonkey 或 Violentmonkey。
2. 导入 [`client/monitorluna.user.js`](./client/monitorluna.user.js)。
3. 在脚本菜单中配置本地 Agent 地址，默认是 `ws://127.0.0.1:6315/ws/browser`。
4. 如已设置浏览器扩展密码，脚本内也要填写相同 token。

## Koishi 配置

基础配置：

- `token`：必填，Agent 握手 token。
- `commandTimeout`：命令超时时间，默认 `15000`。
- `debug`：是否输出调试日志。

存储配置：

- `storageType=local`：使用 `storagePath` 保存图片，本地 URL 基于 `serverPath` 或 Koishi `selfUrl` 生成。
- `storageType=webdav`：需要 `webdavEndpoint`、`webdavUsername`、`webdavPassword`，可选 `webdavPublicUrl`。
- `storageType=s3`：需要 `s3Endpoint`、`s3Bucket`、`s3AccessKeyId`、`s3SecretAccessKey`，可选 `s3Region`、`s3PublicUrl`、`s3PathStyle`。

推送配置：

- `pushChannelIds`：群聊目标，格式 `platform:selfId:channelId`。
- `pushPrivateIds`：私聊目标，格式 `platform:selfId:userId`。
- `pushPollInterval`：窗口切换轮询间隔。
- `dailySummaryEnabled`：是否启用每日总结推送。
- `dailySummaryTime`：每日总结时间，格式 `HH:mm`。

## 命令

- `monitor.list`：列出在线设备。
- `monitor.screen <设备ID>`：截取设备全屏并返回图片。
- `monitor.window <设备ID>`：截取当前活跃窗口并返回图片。
- `monitor.status <设备ID>`：查看设备 CPU / 内存 / GPU 状态。
- `monitor.analytics <设备ID>`：生成当天活动总结图。

## 项目结构

```text
koishi-plugin-monitorluna/
├── src/
│   └── index.ts
├── client/
│   ├── screenshot-server.py
│   ├── start-server.bat
│   ├── start-silent.vbs
│   ├── requirements.txt
│   ├── pyproject.toml
│   └── monitorluna.user.js
├── .github/workflows/release.yml
├── package.json
└── readme.md
```

## 发布说明

GitHub Release workflow 会从 `client/` 目录打包 Agent。发布前请确认 `client/` 内脚本与依赖文件完整。
