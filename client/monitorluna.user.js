// ==UserScript==
// @name         MonitorLuna Browser Tracker
// @namespace    https://github.com/lumia1998/koishi-plugin-monitorluna
// @version      2.0.0
// @description  追踪浏览器各域名活跃时长，上报到本地 MonitorLuna Agent
// @author       MonitorLuna
// @match        *://*/*
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_registerMenuCommand
// @run-at       document-start
// ==/UserScript==

(function () {
  'use strict';

  // ── 配置 ──────────────────────────────────────────────────────────────────
  const CFG_URL      = 'ml_url';
  const CFG_TOKEN    = 'ml_token';
  const CFG_SETUP_DONE = 'ml_setup_done';
  const REPORT_INTERVAL = 30000; // 30秒上报一次
  const IDLE_TIMEOUT    = 60000; // 1分钟无交互视为非活跃

  let wsUrl = GM_getValue(CFG_URL, 'ws://127.0.0.1:6315/ws/browser');
  let wsToken = GM_getValue(CFG_TOKEN, '');
  let setupDone = GM_getValue(CFG_SETUP_DONE, false);

  // ── 菜单命令（点击脚本管理器图标可配置）──────────────────────────────────
  GM_registerMenuCommand('⚙️ MonitorLuna 设置', openSettings);
  GM_registerMenuCommand('📊 查看统计状态', showStatus);

  function openSettings() {
    const curUrl = GM_getValue(CFG_URL, 'ws://127.0.0.1:6315/ws/browser');
    const curToken = GM_getValue(CFG_TOKEN, '');

    const newUrl = prompt('本地 Agent WebSocket 地址：', curUrl);
    if (newUrl === null) return;

    const trimmedUrl = newUrl.trim();
    if (!trimmedUrl.startsWith('ws://') && !trimmedUrl.startsWith('wss://')) {
      alert('地址必须以 ws:// 或 wss:// 开头');
      return;
    }

    const newToken = prompt('浏览器扩展密码（留空则不启用鉴权）：', curToken);
    if (newToken === null) return;

    GM_setValue(CFG_URL, trimmedUrl);
    GM_setValue(CFG_TOKEN, newToken.trim());
    GM_setValue(CFG_SETUP_DONE, true);
    wsUrl = trimmedUrl;
    wsToken = newToken.trim();
    setupDone = true;

    alert('✅ 保存成功！正在重新连接...');
    reconnect();
  }

  function showStatus() {
    const state = ws ? ['CONNECTING','CONNECTED','CLOSING','CLOSED'][ws.readyState] : 'NO WS';
    const domains = Object.entries(pendingStats)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5)
      .map(([d, s]) => `  ${d}: ${Math.round(s)}s`)
      .join('\n');
    alert(`MonitorLuna 状态\n服务器: ${wsUrl}\n连接状态: ${state}\n\n待上报 TOP5:\n${domains || '  (暂无数据)'}`);
  }

  // ── 时长追踪 ─────────────────────────────────────────────────────────────
  let pendingStats = {};   // domain -> seconds (本标签页待上报)
  let trackStart   = null; // 当前开始追踪的时间
  let isActive     = true; // 当前标签页是否处于活跃状态
  let lastActivity = Date.now();

  const currentDomain = location.hostname;
  if (!currentDomain) return; // about:blank 等跳过

  function startTracking() {
    if (!isActive || trackStart !== null) return;
    trackStart = Date.now();
  }

  function stopTracking() {
    if (trackStart === null) return;
    const elapsed = (Date.now() - trackStart) / 1000;
    if (elapsed > 0.5) {
      pendingStats[currentDomain] = (pendingStats[currentDomain] || 0) + elapsed;
    }
    trackStart = null;
  }

  // 页面可见性变化
  document.addEventListener('visibilitychange', () => {
    if (document.hidden) {
      stopTracking();
      isActive = false;
    } else {
      isActive = true;
      lastActivity = Date.now();
      startTracking();
    }
  });

  // 用户活动检测（防止静置标签页持续计时）
  ['mousemove', 'keydown', 'click', 'scroll', 'touchstart'].forEach(evt => {
    document.addEventListener(evt, () => {
      lastActivity = Date.now();
      if (!isActive && !document.hidden) {
        isActive = true;
        startTracking();
      }
    }, { passive: true });
  });

  // 空闲检测
  setInterval(() => {
    if (Date.now() - lastActivity > IDLE_TIMEOUT) {
      if (isActive) {
        stopTracking();
        isActive = false;
      }
    } else if (!document.hidden && !isActive) {
      isActive = true;
      startTracking();
    }
  }, 5000);

  // 页面卸载时保存
  window.addEventListener('beforeunload', () => stopTracking());

  // ── WebSocket ─────────────────────────────────────────────────────────────
  let ws = null;
  let wsReady = false;
  let reconnectTimer = null;

  function connect() {
    if (!wsUrl) return;
    try {
      ws = new WebSocket(wsUrl);
    } catch {
      scheduleReconnect();
      return;
    }

    ws.onopen = () => {
      ws.send(JSON.stringify({ type: 'hello', token: wsToken }));
    };

    ws.onmessage = (e) => {
      try {
        const msg = JSON.parse(e.data);
        if (msg.type === 'hello_ack') wsReady = true;
      } catch {}
    };

    ws.onerror = () => { wsReady = false; };

    ws.onclose = () => {
      wsReady = false;
      ws = null;
      scheduleReconnect();
    };
  }

  function scheduleReconnect() {
    if (reconnectTimer) return;
    reconnectTimer = setTimeout(() => {
      reconnectTimer = null;
      connect();
    }, 15000);
  }

  function reconnect() {
    wsReady = false;
    if (ws) { try { ws.close(); } catch {} ws = null; }
    if (reconnectTimer) { clearTimeout(reconnectTimer); reconnectTimer = null; }
    connect();
  }

  // ── 定时上报 ─────────────────────────────────────────────────────────────
  setInterval(() => {
    // 先 flush 当前计时
    if (trackStart !== null && isActive) {
      const elapsed = (Date.now() - trackStart) / 1000;
      if (elapsed > 0.5) {
        pendingStats[currentDomain] = (pendingStats[currentDomain] || 0) + elapsed;
        trackStart = Date.now(); // 重置，避免重复计算
      }
    }

    if (!wsReady || !ws || Object.keys(pendingStats).length === 0) return;

    const stats = {};
    for (const [domain, secs] of Object.entries(pendingStats)) {
      const rounded = Math.round(secs);
      if (rounded > 0) stats[domain] = rounded;
    }
    if (Object.keys(stats).length === 0) return;

    try {
      ws.send(JSON.stringify({
        type: 'browser_activity',
        stats,
      }));
      pendingStats = {};
    } catch {
      wsReady = false;
    }
  }, REPORT_INTERVAL);

  // ── 初始化 ──────────────────────────────────────────────────────────────
  if (!setupDone) {
    setTimeout(() => {
      if (confirm('MonitorLuna: 检测到尚未配置，是否现在设置？')) {
        openSettings();
      }
    }, 1000);
  } else {
    connect();
    if (!document.hidden) startTracking();
  }

})();
