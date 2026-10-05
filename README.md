# 浙江工商大学论文排版工具

这是一个 Flask + Gunicorn 应用，用于把未排版的 `.docx` 论文转换为浙江工商大学格式，也支持粘贴文本排版、封面与正文合并，以及两个 Word 文档的拼接。

## 本地运行

项目使用 Python 3.12（部署环境固定为 3.12.13）。在仓库根目录执行：

```bash
python -m venv .venv
source .venv/bin/activate       # Windows PowerShell：.venv\Scripts\Activate.ps1
python -m pip install --only-binary=:all: --require-hashes -r requirements.txt
gunicorn app:app --config gunicorn_config.py --workers 1 --threads 8 \
  --timeout 300 --graceful-timeout 290 --bind 127.0.0.1:5057
```

打开 <http://127.0.0.1:5057/> 即可使用。另开终端运行 `curl -s http://127.0.0.1:5057/api/health`，确认返回 HTTP 200 且 `status` 为 `ok` 后再提交文件。开发环境也可以用 `flask --app app run`，但异步任务、SSE 进度和优雅退出验证应使用上面的 Gunicorn 命令。

## 浏览器使用

1. **上传文件**：选择正文 `.docx`，可选上传封面，或勾选“自动生成模板封面”填写课程信息，然后点击“开始排版”。
2. **粘贴文字**：切换到“粘贴文字”，粘贴正文后点击“一键排版并生成文档”。纯文本接口最多接受 2MB 文本和 10,000 个有效段落。
3. **拼接文档**：切换到“拼接文档”，按顺序选择封面和正文；工具会从新页开始拼接并保留两个文档各自的排版。
4. 任务会优先通过实时进度显示结果；浏览器不支持 `EventSource` 或实时连接达到上限时会自动改用结果轮询。下载前请等待任务进入完成状态，不要关闭页面后立即重复提交。

浏览器端文件预检上限为 48MB，服务端单次请求上限为 50MB；两个 `.docx` 输入的总大小也必须落在服务端存储预算内。完整的上传、并发、临时文件和产物保留限制见 [DEPLOY.md](DEPLOY.md)。

## 部署

- 推荐使用 Render Blueprint：在 Render 控制台选择 **New → Blueprint** 并导入本仓库，配置由 [render.yaml](render.yaml) 提供。
- Railway 和普通 VPS 的启动方式见 [DEPLOY.md](DEPLOY.md)。所有常驻部署都必须使用单个 Gunicorn worker；任务状态和临时文件保存在实例内存与 `/tmp`，不适合水平扩容。
- 根目录的 [vercel.json](vercel.json) 只负责将旧 Vercel 地址以 307 重定向到 Render，不在 Vercel 上运行 Python 应用。

生产依赖通过 `requirements.txt` 哈希锁定安装。修改 [requirements.in](requirements.in) 后，请按 [DEPLOY.md](DEPLOY.md) 中的命令重新生成锁文件，并在部署前运行完整测试。

Gunicorn 默认将请求头总时限设为 10 秒、上传正文总时限设为 300 秒；Render Blueprint 会显式配置
这两个边界（`REQUEST_HEADER_TIMEOUT_SECONDS` 和 `REQUEST_BODY_TIMEOUT_SECONDS`），避免慢速请求
长期占用工作线程。正文完整接收后，正文时限会立即撤销。

## 验证

```bash
python -m compileall -q app.py format_paper.py docx_validation.py
python -m unittest discover -s tests -p 'test_*.py'
```

测试覆盖 API、异步任务恢复、前端运行时、DOCX 安全校验和部署保活脚本。
