// 保活 Render 免费档：Cloudflare Cron 每 5 分钟触发一次，戳健康检查接口让实例保持唤醒。
const TARGET = "https://zjgsu-paper-formatter.onrender.com/api/health";

export default {
  async scheduled(_event, _env, ctx) {
    ctx.waitUntil(fetch(TARGET, { signal: AbortSignal.timeout(60_000) }).catch(() => {}));
  },

  // 手动访问 Worker 地址时返回目标当前状态，便于随手检查保活是否正常。
  async fetch() {
    try {
      const started = Date.now();
      const res = await fetch(TARGET, { signal: AbortSignal.timeout(60_000) });
      return new Response(
        `keepwarm OK -> ${TARGET}\nstatus: ${res.status}\nlatency: ${Date.now() - started}ms\n`,
        { headers: { "content-type": "text/plain; charset=utf-8" } },
      );
    } catch (err) {
      return new Response(`keepwarm ping failed: ${err}\n`, { status: 502 });
    }
  },
};
