# koishi-plugin-screenluna

远程屏幕截图插件

## 功能

- 截取远程设备的屏幕
- 支持多设备管理
- 通过设备名快速截图

## 使用方法

### 1. 在远程设备上安装截图服务端

**Windows 用户（推荐）：**

下载并双击运行启动脚本：
https://github.com/lumia1998/koishi-plugin-screenluna/raw/main/start-server.bat

脚本会自动检测 Python 环境、安装依赖并启动服务。

**手动安装：**

下载服务端脚本：
```bash
wget https://raw.githubusercontent.com/lumia1998/koishi-plugin-screenluna/main/screenshot-server.py
```

安装依赖并运行：
```bash
pip install flask pyautogui pillow
python screenshot-server.py
```

### 2. 配置设备

在 Koishi 插件配置页面添加设备：
- 设备名称：例如 "家里"、"单位"
- IP 地址：例如 10.1.2.200
- 端口：默认 6314

### 3. 使用命令

- `screen 家里` - 截取家里电脑的屏幕
- `screen 单位` - 截取单位电脑的屏幕
- `screen -l` - 查看已绑定的设备列表

## 注意事项

- 远程设备需要运行截图服务
- 确保网络可达
- 注意防火墙设置
