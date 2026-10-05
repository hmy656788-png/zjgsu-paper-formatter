# 部署到常驻服务器（解决 Vercel 4.5MB 上传限制）

Vercel serverless 单次请求/响应体有 ~4.5MB 硬限制（`FUNCTION_PAYLOAD_TOO_LARGE`），
封面+正文较大时拼接会失败。本应用本就是为常驻服务器设计的（后台线程、内存任务、
SSE、`/tmp` 临时文件），换到常驻平台跑 `gunicorn` 即可彻底解决，上传上限提升到 50MB。

## Vercel 仅作为兼容跳转

根目录的 `vercel.json` 不构建 Python Function，也不在 Vercel 上运行应用代码；它只把所有路径
以临时重定向（HTTP 307）转到 `https://zjgsu-paper-formatter.onrender.com`。307 会保留 POST
方法和请求体，因此旧的 Vercel 域名仍可兼容表单与 API 地址，同时真正的任务创建、SSE、结果查询
和下载始终落到同一个 Render 服务。这里使用 Render 的直连域名，避免自定义域名重新指向 Vercel
时形成重定向循环。

不要把该跳转改为 Vercel rewrite/proxy：代理仍会让上传体积和 SSE 受 Vercel 请求链路限制。若要
撤销兼容入口，只需移除 Vercel 域名；Render 部署本身不受影响。
`.vercelignore` 使用只允许 `vercel.json` 的白名单，避免把应用源码或本地 `.env` 一并上传；
Git 也会忽略所有 `.env` / `.env.*`，仅允许提交不含密钥的 `.env.example` 模板。

## 启动命令（核心）

```bash
gunicorn app:app --config gunicorn_config.py --workers 1 --threads 8 --timeout 300 --graceful-timeout 290 --bind 0.0.0.0:$PORT
```

> ⚠️ 必须 **单 worker + 多线程**：
> - 单 worker —— 任务进度与 `/tmp` 文件是进程内共享的，多 worker 会让「建任务」和
>   「查进度/下载」落到不同进程而 404。
> - 多线程 —— 支撑 SSE 长连接和并发请求。
> - 收到 SIGTERM 后，服务会停止接收新排版任务，并等待已启动的后台任务完成。Render 配置会
>   最多等待 300 秒，Gunicorn 在 290 秒后结束仍未完成的 worker，为平台清理预留 10 秒。
> - 超过退出窗口的任务，以及实例崩溃、强制终止或底层主机故障时的任务仍会丢失；当前架构
>   不支持持久任务队列或水平扩容。
> - 异步任务状态和输出仍只存在于创建它的实例内存与 `/tmp`。零停机部署切流后，即使旧实例
>   完成了任务，后续结果或下载请求也可能落到新实例而无法取回。需要完全消除该限制时，必须把
>   任务队列、状态和产物迁移到外部持久存储；发布前仍应尽量避开活跃任务。

## 文档处理并发保护

同步与异步排版、合并和拼接共用固定数量的处理槽位，默认最多同时处理 2 个任务。超过容量的新请求
会在解析 multipart/JSON 请求体前返回 HTTP 503，不会先把大文件写入临时 spool；异步入口也不会
继续创建线程挤占内存。异步请求接受后会把入口预留的槽位移交给后台线程，直到任务结束才释放。
可通过 `MAX_CONCURRENT_PROCESSING_JOBS` 调整为 1–8；无效或越界值会回退到安全默认值 2。

`render.yaml` 已显式设置为 2。提高该值前应先根据实例内存，用接近 50MB 的真实 DOCX 做并发压测；
Gunicorn 的 `--threads 8` 是 HTTP/SSE 请求线程数，不等同于允许同时执行 8 个文档处理任务。
HTTP/1.x 请求若在解析 body 前因限流、过大、无效恢复标识、服务排空或任务幂等复用而提前返回，
响应写完后会立即停用对应 Gunicorn 连接，不再让线程用最多 5 秒排空慢速请求体。完整解析过 body
的正常请求仍可复用 keep-alive；Gunicorn 已完整接收 stream 的 HTTP/2 连接也不会被整体关闭。
Gunicorn 还会为每个 HTTP/1 请求头（包括 keep-alive 上的后续请求）执行默认 10 秒的总时限；
缓慢滴入单个字节不会重置计时。超时的不完整请求头会直接双向断连，不进入 Gunicorn 逐连接的
优雅排空；时限在头部解析完成后立即撤销，不会截断大文件正文、SSE 或下载。可用
`REQUEST_HEADER_TIMEOUT_SECONDS` 在 1–60 秒内调整；无效值回退为 10 秒。
请求头完成后，上传正文还受不可被慢速滴入重置的总时限保护，默认 300 秒；正文完整消费后该
时限立即撤销，不会限制后续排版处理、SSE 或下载。可用 `REQUEST_BODY_TIMEOUT_SECONDS` 在
1–600 秒内调整；无效值回退为 300 秒。Render Blueprint 已显式设置为 `300`，让生产部署的
慢速上传边界可审计且与应用默认值一致。
上传包在 50MB HTTP 上限之外还有解压预算：包内总量最多 64MB，单个 XML 最多 8MB、XML 总量最多
16MB。主文档还会在交给 `python-docx` 前流式检查表格结构：单行逻辑网格最多 256 列，全文最多
20,000 个逻辑单元格，避免极小 XML 通过异常 `gridSpan` 声明放大为海量 Python 对象。这些上限是按
`python-docx` 实际 XML 解析内存放大后的实例预算设置的。排版遍历按实际 `w:tc` 单元格处理，横向
或纵向合并不会重复展开、递归处理同一单元格。每个 XML 部件还限制为最多 200,000 个元素、
256 层嵌套，整份文档累计最多 500,000 个 XML 元素。所有 Word 故事部件还共享
10,000 个 `w:p` 段落和 50,000 个 `w:r` 文本运行的独立预算，阻止数万个空段落或空 run
被高压缩成数十 KB 后，在 `python-docx` 中放大为数百 MiB 对象。双文档生成产物使用
两倍预算并保留小幅封面/目录/分节余量。同时拒绝 DTD/实体声明以及带路径穿越或符号链接属性的
ZIP 成员，避免标准库或下游解析器发生实体展开、深度放大或包路径混淆。
ZIP 成员仅允许 STORE/DEFLATE；校验器会绕过可伪造的声明展开尺寸，按原始压缩数据以 64 KiB
分块重新计算真实展开长度与 CRC，并核对本地头、数据描述符和中央目录边界。因而攻击者不能用
伪造 `file_size` 或不受控的 BZIP2/LZMA 解压路径，把超大载荷伪装成很小的成员。
关系部件另设更紧的预算：单个 `.rels` 最多 1 MiB、4,096 条关系，整包最多 4 MiB、8,192 条；
生成产物按两个已验证输入的合法峰值放宽到两倍，避免高压缩关系扇出把小上传放大为数百 MiB 内存。
这些规则集中在无 Flask 依赖的 `docx_validation.py`；Web 上传、Python 直接调用与命令行入口共用
同一限制配置。直接入口还会先按 50MB 原始归档上限拒绝文件，再在同一个已验证文件句柄上调用
`python-docx`，避免“校验路径后重新打开”产生替换竞态。脚注直接格式化也复用同一个已验证句柄；
脚注改写只在暂存产物再次通过生成文档预算后才原子替换原文件，避免格式补全造成部件放大后发布。
合并入口会先验证封面，再开始正文排版。

multipart 请求会在读取正文前按 `2 × Content-Length` 预留临时空间（缺少长度时按 100MB 上界），
覆盖 Werkzeug spool 与 `/tmp/uploads` 命名副本短时并存的峰值；该预留与输出产物共用同一容量账本。
命名副本保存完成后会立即关闭全部上传流并释放增长预留，已落盘输入继续由真实磁盘余量计入预算。
因此活动上传不能侵占已承诺给输出的空间，活动输出也会阻止新的上传超额进入。

SSE 实时进度流默认最多同时保持 4 条连接，超额订阅返回 HTTP 429，前端会自动改用结果轮询。
可通过 `MAX_CONCURRENT_SSE_CONNECTIONS` 调低为 1–4；不要把它设到 Gunicorn 总线程数，
必须给健康检查、任务提交、轮询和下载预留普通请求线程。`render.yaml` 已显式设为 4。
浏览器完全不支持 `EventSource` 时也仍会走可恢复的异步任务创建接口，并在拿到结果地址后直接轮询，
不会退回无法保存任务标识和结果地址的同步长请求。

进程内任务表最多保留 128 条记录。容量不足时优先回收客户端已经请求过的终态结果；尚未请求的
完成或失败结果至少保留 10 分钟，与前端最长等待窗口一致。只有任务表被大量未读取结果占满时，
新任务才会暂时返回 HTTP 503，而不会用新请求换掉其他用户仍在等待的结果。结果接口每次收到 GET
都会刷新至少 60 秒的保护窗口，给响应中断后的重试留出完整时间；自动 HEAD 不会把结果标记为已领取。

提交异步请求前，浏览器会先生成 128 位随机任务标识，通过 `X-Job-ID` 发送，并把可预知的
`result_url`、任务模式、创建时间和“待确认”状态写入当前标签页的 `sessionStorage`。后端会按同类型
任务幂等复用该标识，因此即使任务创建成功但 POST 响应丢失，刷新后仍能从已保存的 URL 找回结果；
创建窗口内的 404 会暂按“任务尚未出现”继续轮询，而不是立即删除恢复指针；该窗口比 3 分钟的
创建请求超时额外延长 2 分钟，用来覆盖慢上传或服务端收尾的边界。相同标识的重试会在限流前直接
复用原任务，即使处理槽已满也不会重复解析请求体或启动 worker。携带该标识的请求若在上传落盘或
后台线程启动阶段失败，后端也会保留对应的终态错误，恢复入口可立即给出失败原因；尚未创建任务的
5xx 会显式标记为不可恢复，避免无效 provisional 记录阻塞新任务。结果轮询遇到没有终态标记的临时
5xx 时不会提前确认 provisional 任务，因此后续创建窗口内的 404 仍会继续等待。未携带恢复标识的
旧客户端仍会直接回收这类创建失败记录。

最终结果读取遇到网络错误、超时或 5xx 会有限退避重试；停止等待或刷新后会显示“重新获取任务结果”，
并用有界轮询继续等待尚未完成的任务。结果成功展示后仍会保留恢复记录，直到用户明确选择继续新任务
或放弃旧任务；超过创建窗口后的 404、确认 410 过期或服务器明确返回终态失败时也会清除记录。SSE
失败事件展示给用户后还会 best-effort 请求一次结果接口，避免已读错误继续占用“未领取结果”保留位。

输出目录另有独立的硬预算：最多 256 个 `.docx`、合计 256 MiB。每个开始处理的任务会先预留
64 MiB，并在 ZIP 每次写入前强制执行 64 MiB 单文件上限；超过上限会返回 HTTP 413 并清理半成品。
合并排版会为“临时排版正文 + 最终产物”预留 128 MiB，且临时正文位于受 6 小时 TTL 管理的上传
目录。系统还会要求文件系统至少保留 64 MiB 空闲，因此即使把处理并发上调到很高，也不会让多个
大文档同时无上限写满 `/tmp`。生成中的最终路径不会提供下载，避免客户端读到不完整的 ZIP。
处理器返回成功后还会在任务终点逐项验证 ZIP 成员路径与类型、重复成员、必需 OOXML 部件、全部成员
CRC、内容类型表和主文档 XML；生成产物的展开预算按最多两个输入包计算，并确认主文档声明了正确的
Word 内容类型。截断、损坏或结构不完整的产物会转为失败并立即清理，不会发布下载地址。
所有新产物至少保留 30 分钟；容量压力下只会回收超过该窗口的系统
生成文件，顺序为已请求过结果的异步产物、无任务状态的同步/遗留产物、尚未请求结果的异步产物。
正在生成的文件、未知来源文件和保留窗口内的结果不会被删除；无法安全回收时，新任务返回 HTTP
503。每次写入前会先清理过期文件再做可写探针，避免磁盘已满时因探针失败而永远无法进入清理。
健康检查除了要求当前用量未超预算，还要求至少能再预留一份标准产物；无法接收新任务时
`/api/health` 会返回 degraded/503，Render 部署门不会把“进程存活但所有新任务失败”误判为可用。

下载接口支持单段 `Range: bytes=...` 与断点续传，有效范围返回 HTTP 206 和对应的
`Content-Range`；无效范围或多段 Range 返回 JSON 格式的 HTTP 416。下载响应保持打开期间会为
产物持有活动下载租约，TTL 清理与容量回收都不会删除该文件；响应关闭并释放租约后仍会刷新至少
60 秒的重试保护窗口，避免连接中断后立即重试却遇到产物已被回收。租约会先锚定真实输出目录，
再以只读文件描述符打开并核对目标身份；路径在响应期间被替换也不会改变实际发送的文件。Range
读取使用可 seek 的流定位范围起点，下载文件尾部时不会先顺序读取整份文档。

## 方式一：Render（推荐，有免费档）

1. 把代码推到 GitHub。
2. Render 控制台 → **New → Blueprint** → 选本仓库 → 应用会读取根目录的 `render.yaml`。
3. 等构建完成，访问分配的 `*.onrender.com` 域名即可。
4. 自定义域名：Service → Settings → Custom Domains，把 `zjgsu-formatter.hmyapp.com`
   的 DNS 指过来。

免费档会在闲置后休眠，下次访问有几十秒冷启动；要常驻可升 Starter 档。

### 自动部署门禁与保活告警

GitHub Actions 会在 pull request 和 main 推送时安装生产依赖并运行完整测试、JavaScript 语法检查
与 Python 编译检查；只有全部通过的 main 推送或在 main 上手动运行才会调用
`RENDER_DEPLOY_HOOK`。工作流会把本次通过验证的 `GITHUB_SHA` 作为 Deploy Hook 的 `ref`
提交，`render.yaml` 同时显式关闭 Render 自带的提交自动部署，避免未经该门禁的版本抢跑。
该 Secret 缺失时工作流会明确失败，不会静默跳过部署。Hook 返回 200 仅代表部署开始，202 仅代表
排队；工作流会继续轮询 `/api/health`，直到响应中的 `deployment.commit` 与本次 `GITHUB_SHA`
完全一致且服务健康，才把部署标为成功。构建、启动或切流失败会保留旧 commit，并在 30 分钟后让
工作流明确失败，避免“已接受部署请求”被误报为“已经上线”。该部署身份来自 Render 提供的
`RENDER_GIT_COMMIT`，后端只公开格式合法的 40 位十六进制 SHA。

定时保活会在短暂网络错误时自动重试，但持续失败会把工作流标记为失败，避免 Render 服务异常被
`|| true` 吞掉。相同保活任务不会并行堆积。

Cloudflare keepwarm 侧与 GitHub Actions 侧都支持从配置覆盖健康检查参数，可在同一仓库变量/环境中设置，
无需改动脚本，避免每次改域名都重新提交配置：
- `KEEPWARM_TARGET_URL=https://你的域名/api/health`
- `KEEPWARM_CONNECT_TIMEOUT=10`（单位：秒）
- `KEEPWARM_MAX_TIME=40`（单位：秒）
- `KEEPWARM_RETRY_COUNT=2`
- `KEEPWARM_RETRY_DELAY=5`（单位：秒）
- `KEEPWARM_REQUEST_TIMEOUT_MS=60000`（单位：ms，可选，Worker 单独参数）

Action 侧：
`KEEPWARM_CONNECT_TIMEOUT` 用于连接超时，`KEEPWARM_MAX_TIME` 用于整次 `curl` 最长等待；`KEEPWARM_RETRY_COUNT`
和 `KEEPWARM_RETRY_DELAY` 控制短暂抖动后的重试次数与间隔。配置会做基本校验：URL 非 http(s) 地址
会回退到默认健康检查地址；空值/非法值/小于 1 的数字会回退到默认值，并保证 `KEEPWARM_MAX_TIME >=
KEEPWARM_CONNECT_TIMEOUT + 1`。为避免仓库变量误配拖垮计划任务，工作流还将连接超时、总超时、重试次数和
重试间隔分别限制为 60 秒、120 秒、5 次和 30 秒以内。

Worker 侧：
`KEEPWARM_TARGET_URL` 可替换默认健康检查地址；`KEEPWARM_REQUEST_TIMEOUT_MS`（兼容 `KEEPWARM_TIMEOUT_MS`）
默认为 60000ms，Worker 会在检测到 URL 非法时回退默认地址，超时在 5000–120000ms 外会自动裁剪；发生回退时会在
fetch 返回体里带上 `warnings` 字段。

## 方式二：Railway

1. New Project → Deploy from GitHub repo。
2. Railway 自动识别 Python，并使用根目录的 `Procfile` 启动命令。
3. Settings → Networking → Generate Domain。

## 方式三：任意 VPS / Docker

```bash
python -m pip install --only-binary=:all: --require-hashes -r requirements.txt
gunicorn app:app --config gunicorn_config.py --workers 1 --threads 8 --timeout 300 --graceful-timeout 290 --bind 0.0.0.0:8000
```
前面挂 Nginx/Caddy 反代即可（注意把反代的请求体上限调到 ≥ 50MB，
例如 Nginx `client_max_body_size 50m;`）。

## Python 与依赖锁定

Render 和 GitHub Actions 都固定使用 Python 3.12.13。`requirements.in` 只维护直接生产依赖，
`requirements.txt` 是部署锁文件：它精确固定完整传递依赖闭包，并记录 PyPI 分发包的 SHA-256。
Render、CI 与手动部署都使用 `--require-hashes` 和 `--only-binary=:all:`，因此上游发布新版本、
替换非预期文件或在不同机器上临时编译原生扩展，都不会让一次无代码变更的重建悄悄换掉运行栈。

只有在准备好运行完整回归时才更新依赖。修改 `requirements.in` 后，用项目锁文件头部记录的命令
重新生成 `requirements.txt`：

```bash
uv pip compile requirements.in --universal --python-version 3.12 --generate-hashes --output-file requirements.txt
```

随后应在干净环境中执行哈希安装、完整测试与 `pip check`，不要直接手改生成的版本或哈希。

## 上传上限

- 后端：`app.py` 的 `MAX_CONTENT_LENGTH`（默认 50MB）。
- 前端：`static/script.js` 的 `MAX_UPLOAD_BYTES`（默认 48MB，留余量）。
- 两处需保持一致；要改上限就同时改这两个值（并确保反代的 body 限制也够大）。
- 纯文本排版接口会在解析 JSON 前额外限制请求体约 6MB，并在解析后继续校验 `text` 字段不超过
  2MB、有效段落不超过 10,000 段；额外空间用于容纳 Unicode JSON 转义和少量封面字段，未使用的
  大字段会直接返回 413。段落计数将 CRLF 视为一个分隔符，单独的 CR/LF 各视为一个分隔符，保留
  中间空段但忽略末尾仅由换行产生的空段。浏览器会做同样的预检，后端校验仍是最终边界。

## 跨域访问

默认只支持同源 API 访问，不向任意 `Origin` 返回 CORS 许可。若确实要把前端和后端部署在不同域名，
设置逗号分隔的 `CORS_ALLOWED_ORIGINS`。只填写完整 origin（协议 + 域名 + 可选端口），不要填 `*`
或带路径的 URL，例如：

```bash
CORS_ALLOWED_ORIGINS=https://zjgsu-formatter.example.com,https://admin.example.com
```

## 本地验证

```bash
gunicorn app:app --config gunicorn_config.py --workers 1 --threads 8 --graceful-timeout 290 --bind 127.0.0.1:5057
# 另开终端：
curl -s http://127.0.0.1:5057/api/health
```

API 响应的 `X-Request-ID` 是服务器为当前 HTTP 请求生成的标识；同步请求出现未预期异常时，
脱敏日志会附带相同的 `request_id`，可据此定位失败。客户端传入的同名头不会覆盖它。
允许的跨域来源也可读取这个响应头。异步处理中的失败仍使用任务标识定位。

`/api/health` 只有在上传与输出目录都是真实、可写且服务未进入退出排空状态时才返回 HTTP 200；
存储不可用时返回 HTTP 503 与 `status: "degraded"`，收到 SIGTERM 后返回 HTTP 503 与
`status: "draining"`，便于探测和日志明确区分故障实例与正常退出中的实例。健康检查也会顺带执行
受租约和任务状态保护的 TTL 清理，因此仅有保活流量时，崩溃遗留的过期临时文件也不会无限占用
`/tmp`。响应里的 `jobs`
只提供当前进程的聚合计数：`active` 是排队或处理中的异步任务数，`terminal` 是已结束但仍保留结果的
任务数，`tracked` 是任务表总量，`capacity` 是任务表上限；不会暴露任务 ID、文件名或错误内容。
手动发布前可先确认 `active` 为 0，但该读数只是请求时刻的快照，不能替代退出排空机制。
`output_storage` 仅提供文件数、总字节、含活动预留的有效用量、活动上传预留的数量/字节数、是否位于
预算内以及当前能否再预留一个标准输出，不暴露文件名；对应聚合硬上限、单文件硬上限、单请求上传
预留上界与最短保留时间位于 `limits`。
