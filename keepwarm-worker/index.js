// 保活 Render 免费档：Cloudflare Cron 每 5 分钟触发一次，戳健康检查接口让实例保持唤醒。
const DEFAULT_TARGET = "https://zjgsu-paper-formatter.onrender.com/api/health";
const DEFAULT_REQUEST_TIMEOUT_MS = 60_000;
const MIN_REQUEST_TIMEOUT_MS = 5_000;
const MAX_REQUEST_TIMEOUT_MS = 120_000;
const DEFAULT_TARGET_SOURCE = "default";
const TEXT_HEADERS = {
  "content-type": "text/plain; charset=utf-8",
  "cache-control": "no-store",
  "pragma": "no-cache",
  "expires": "0",
  "x-content-type-options": "nosniff",
  "referrer-policy": "strict-origin-when-cross-origin",
  "x-frame-options": "DENY",
  "permissions-policy": "camera=(), microphone=(), geolocation=()",
  "content-security-policy": "default-src 'none'; base-uri 'none'; frame-ancestors 'none'",
};

function isValidHttpUrl(value) {
  if (!value) {
    return false;
  }
  try {
    const parsed = new URL(value);
    // The target is echoed in the Worker report and sent to an external
    // service.  Reject credentials so a mistyped secret-bearing URL cannot
    // leak through logs/reports, and require a real host instead of accepting
    // values such as `https:///api/health`.
    return (parsed.protocol === "http:" || parsed.protocol === "https:")
      && Boolean(parsed.hostname)
      && !parsed.username
      && !parsed.password;
  } catch (_err) {
    return false;
  }
}

function parsePositiveIntegerMs(rawValue, fallback) {
  const text = String(rawValue || "").trim();
  if (!/^\d+$/.test(text)) {
    return fallback;
  }
  const parsed = Number.parseInt(text, 10);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    return fallback;
  }
  return parsed;
}

function resolveConfig(env = {}) {
  const targetInput = (
    env.KEEPWARM_TARGET_URL
    || env.KEEPWARM_TARGET
    || ""
  ).trim();
  const timeoutInput = String(
    env.KEEPWARM_REQUEST_TIMEOUT_MS
    || env.KEEPWARM_TIMEOUT_MS
    || ""
  ).trim();

  let target = DEFAULT_TARGET;
  let targetSource = DEFAULT_TARGET_SOURCE;
  let targetWarning = "";
  if (targetInput) {
    if (isValidHttpUrl(targetInput)) {
      target = targetInput;
      targetSource = "env";
    } else {
      targetWarning = "KEEPWARM_TARGET_URL/KEEPWARM_TARGET 不是有效 http/https 链接，已回退默认值";
    }
  }

  let parsedTimeout = parsePositiveIntegerMs(timeoutInput, DEFAULT_REQUEST_TIMEOUT_MS);
  let timeoutWarning = "";
  const isTimeoutInputFiniteNumber = /^\d+$/.test(timeoutInput || "");

  if (timeoutInput && !isTimeoutInputFiniteNumber) {
    timeoutWarning = "KEEPWARM_REQUEST_TIMEOUT_MS 非法，已回退默认值 60000ms";
  }

  if (parsedTimeout < MIN_REQUEST_TIMEOUT_MS || parsedTimeout > MAX_REQUEST_TIMEOUT_MS) {
    parsedTimeout = Math.max(
      MIN_REQUEST_TIMEOUT_MS,
      Math.min(MAX_REQUEST_TIMEOUT_MS, parsedTimeout),
    );
    timeoutWarning = `KEEPWARM_REQUEST_TIMEOUT_MS 超出范围 ${MIN_REQUEST_TIMEOUT_MS}~${MAX_REQUEST_TIMEOUT_MS}ms，已裁剪为 ${parsedTimeout}ms`;
  }

  return {
    target,
    timeoutMs: parsedTimeout,
    targetSource,
    targetWarning,
    timeoutWarning,
  };
}

async function fetchHealth(config) {
  const target = config.target;
  const started = Date.now();
  const res = await fetch(target, {
    cache: "no-store",
    headers: { accept: "application/json" },
    signal: AbortSignal.timeout(config.timeoutMs),
  });
  const latencyMs = Date.now() - started;
  let health = null;

  if ((res.headers.get("content-type") || "").includes("application/json")) {
    try {
      health = await res.json();
    } catch {
      health = null;
    }
  }

  const appStatus = typeof health?.status === "string" ? health.status : "unknown";
  const storageReady = typeof health?.storage?.storage_ready === "boolean"
    ? health.storage.storage_ready
    : null;
  // Treat an incomplete health payload as degraded.  A missing storage flag
  // must not let the keepwarm cron report success while the app cannot accept
  // new uploads; the deploy gate uses the same explicit readiness contract.
  const healthy = res.ok && appStatus === "ok" && storageReady === true;

  return {
    httpStatus: res.status,
    appStatus,
    storageReady,
    healthy,
    target,
    targetSource: config.targetSource,
    timeoutMs: config.timeoutMs,
    targetWarning: config.targetWarning,
    timeoutWarning: config.timeoutWarning,
    latencyMs,
  };
}

function formatHealthReport(result) {
  const storageReady = result.storageReady === null ? "unknown" : String(result.storageReady);
  const warnings = [result.targetWarning, result.timeoutWarning].filter(Boolean);
  // Query strings are valid for some health gateways, but may contain an API
  // token.  Keep the actual URL for fetch while never echoing credentials,
  // query parameters, or fragments in the manually exposed report/logs.
  let targetForReport = result.target;
  try {
    const parsed = new URL(result.target);
    parsed.username = "";
    parsed.password = "";
    parsed.search = "";
    parsed.hash = "";
    targetForReport = parsed.toString();
  } catch {
    targetForReport = "[invalid target]";
  }
  return [
    `keepwarm ${result.healthy ? "OK" : "DEGRADED"} -> ${targetForReport}`,
    `target_source: ${result.targetSource}`,
    `timeout_ms: ${result.timeoutMs}`,
    `http_status: ${result.httpStatus}`,
    `app_status: ${result.appStatus}`,
    `storage_ready: ${storageReady}`,
    `warnings: ${warnings.length ? warnings.join(" | ") : "none"}`,
    `latency: ${result.latencyMs}ms`,
    "",
  ].join("\n");
}

function formatWorkerError(err) {
  let message = "unknown error";
  try {
    message = err instanceof Error ? err.message : String(err);
  } catch {
    message = "unknown error";
  }
  const normalized = String(message).replace(/\s+/g, " ").trim();
  // Fetch implementations may include the requested URL in an exception.
  // Strip credentials, query parameters, and fragments before exposing the
  // message from the public Worker endpoint.
  const redacted = normalized.replace(/https?:\/\/[^\s"'<>]+/gi, (raw) => {
    try {
      const parsed = new URL(raw);
      parsed.username = "";
      parsed.password = "";
      parsed.search = "";
      parsed.hash = "";
      return parsed.toString();
    } catch {
      return "[redacted-url]";
    }
  });
  return redacted ? redacted.slice(0, 240) : "unknown error";
}

function logConfigWarnings(config) {
  const warnings = [config.targetWarning, config.timeoutWarning].filter(Boolean);
  if (warnings.length === 0) {
    return;
  }
  console.warn(`keepwarm config warning: ${warnings.join(" | ")}`);
}

export default {
  async scheduled(_event, _env, ctx) {
    const config = resolveConfig(_env);
    logConfigWarnings(config);
    ctx.waitUntil(
      fetchHealth(config)
        .then((result) => {
          if (!result.healthy) {
            throw new Error(formatHealthReport(result));
          }
        })
        .catch((err) => {
          console.warn(`keepwarm ping unhealthy: ${err}`);
        }),
    );
  },

  // 手动访问 Worker 地址时返回目标当前状态，便于随手检查保活是否正常。
  async fetch(_request, env) {
    const config = resolveConfig(env);
    logConfigWarnings(config);
    try {
      const result = await fetchHealth(config);
      return new Response(formatHealthReport(result), {
        status: result.healthy ? 200 : 502,
        headers: TEXT_HEADERS,
      });
    } catch (err) {
      return new Response(`keepwarm ping failed: ${formatWorkerError(err)}\n`, {
        status: 502,
        headers: TEXT_HEADERS,
      });
    }
  },
};
