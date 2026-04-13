# 浙江工商大学论文自动排版工具

基于 `Flask + python-docx` 的论文排版服务，支持：

- 上传 `.docx` 文档并自动排版
- 直接粘贴文本后生成 `.docx`
- 课程论文模板封面生成
- 封面文档与正文合并

## 本地开发

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python3 app.py
```

默认启动地址：

- `http://127.0.0.1:5001`

## 运行测试

```bash
python3 -m unittest discover -s tests -q
```

如果缺少 `docx`、`lxml` 等模块，先确认已经在当前虚拟环境中执行过：

```bash
pip install -r requirements.txt
```

## 部署说明

### 异步任务与实时进度

项目里保留了基于内存和 SSE 的异步排版链路，但它只适合长驻进程环境。

在以下环境里，前端会自动回退到同步处理：

- Vercel
- 其他无状态、实例可能随时切换的 Serverless 环境

这是因为异步任务状态当前保存在进程内存中，输出文件也暂存在本地临时目录。

如果你明确部署在稳定的长驻进程环境中，并且希望启用实时进度，可以设置：

```bash
ENABLE_ASYNC_JOBS=1
```

如果要在真正的 Serverless 环境里长期稳定使用，下一步建议把这两部分改成持久化方案：

- 任务状态：Redis / 数据库 / 队列
- 输出文件：Blob / S3 / 对象存储

### 输出文件存储

当前输出文件默认写入本地临时目录：

- 上传文件：`/tmp/uploads`
- 输出文件：`/tmp/outputs`

这对本地运行没有问题，但不适合作为 Serverless 生产环境里的长期存储方案。
