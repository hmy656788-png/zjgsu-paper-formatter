# 部署到常驻服务器（解决 Vercel 4.5MB 上传限制）

Vercel serverless 单次请求/响应体有 ~4.5MB 硬限制（`FUNCTION_PAYLOAD_TOO_LARGE`），
封面+正文较大时拼接会失败。本应用本就是为常驻服务器设计的（后台线程、内存任务、
SSE、`/tmp` 临时文件），换到常驻平台跑 `gunicorn` 即可彻底解决，上传上限提升到 50MB。

## 启动命令（核心）

```bash
gunicorn app:app --workers 1 --threads 8 --timeout 300 --bind 0.0.0.0:$PORT
```

> ⚠️ 必须 **单 worker + 多线程**：
> - 单 worker —— 任务进度与 `/tmp` 文件是进程内共享的，多 worker 会让「建任务」和
>   「查进度/下载」落到不同进程而 404。
> - 多线程 —— 支撑 SSE 长连接和并发请求。

## 方式一：Render（推荐，有免费档）

1. 把代码推到 GitHub。
2. Render 控制台 → **New → Blueprint** → 选本仓库 → 应用会读取根目录的 `render.yaml`。
3. 等构建完成，访问分配的 `*.onrender.com` 域名即可。
4. 自定义域名：Service → Settings → Custom Domains，把 `zjgsu-formatter.hmyapp.com`
   的 DNS 指过来。

免费档会在闲置后休眠，下次访问有几十秒冷启动；要常驻可升 Starter 档。

## 方式二：Railway

1. New Project → Deploy from GitHub repo。
2. Railway 自动识别 Python，并使用根目录的 `Procfile` 启动命令。
3. Settings → Networking → Generate Domain。

## 方式三：任意 VPS / Docker

```bash
pip install -r requirements.txt
gunicorn app:app --workers 1 --threads 8 --timeout 300 --bind 0.0.0.0:8000
```
前面挂 Nginx/Caddy 反代即可（注意把反代的请求体上限调到 ≥ 50MB，
例如 Nginx `client_max_body_size 50m;`）。

## 上传上限

- 后端：`app.py` 的 `MAX_CONTENT_LENGTH`（默认 50MB）。
- 前端：`static/script.js` 的 `MAX_UPLOAD_BYTES`（默认 48MB，留余量）。
- 两处需保持一致；要改上限就同时改这两个值（并确保反代的 body 限制也够大）。

## 本地验证

```bash
gunicorn app:app --workers 1 --threads 8 --bind 127.0.0.1:5057
# 另开终端：
curl -s http://127.0.0.1:5057/api/health
```
