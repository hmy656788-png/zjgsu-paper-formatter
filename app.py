#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
学术论文自动排版工具 - Web 服务端
基于 Flask 提供文件上传、排版处理、结果下载功能。
"""

import re
import uuid
import time
import logging
from contextlib import suppress
from pathlib import Path

from flask import Flask, request, jsonify, send_file, send_from_directory
from werkzeug.exceptions import HTTPException
from werkzeug.exceptions import RequestEntityTooLarge

from format_paper import (
    ParagraphType,
    format_academic_paper,
    format_academic_paper_from_text,
    merge_cover_and_body,
)

try:
    from flask_cors import CORS
except ImportError:  # pragma: no cover - optional dependency fallback
    CORS = None

# ============================================================
# 应用配置
# ============================================================
app = Flask(__name__, static_folder="static", static_url_path="/static")
if CORS is not None:
    CORS(app)  # 允许跨域请求（适应前后台分离的云部署架构）

# 在 Serverless 环境中，通常只有 /tmp 目录有写入权限
UPLOAD_FOLDER = Path("/tmp") / "uploads"
OUTPUT_FOLDER = Path("/tmp") / "outputs"
UPLOAD_FOLDER.mkdir(exist_ok=True, parents=True)
OUTPUT_FOLDER.mkdir(exist_ok=True, parents=True)

app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB 上传限制

ALLOWED_EXTENSIONS = {".docx"}
GENERATED_OUTPUT_RE = re.compile(r"^[0-9a-f]{8}_output\.docx$")
TEMP_FILE_TTL_SECONDS = 6 * 60 * 60

logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

if CORS is None:
    logger.warning("未检测到 Flask-Cors，已跳过 CORS 配置；同源部署不受影响。")


def allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def is_api_request() -> bool:
    return request.path.startswith("/api/")


def json_error(message: str, status_code: int):
    return jsonify({"success": False, "error": message}), status_code


def get_health_payload() -> dict:
    return {
        "success": True,
        "status": "ok",
        "timestamp": int(time.time()),
        "storage": {
            "upload_dir": str(UPLOAD_FOLDER),
            "output_dir": str(OUTPUT_FOLDER),
            "upload_ready": UPLOAD_FOLDER.exists() and UPLOAD_FOLDER.is_dir(),
            "output_ready": OUTPUT_FOLDER.exists() and OUTPUT_FOLDER.is_dir(),
        },
    }


def get_display_name(filename: str) -> str:
    """提取用于展示/下载的原始文件名，保留中文等 Unicode 字符。"""
    normalized = (filename or "").replace("\\", "/").strip()
    basename = Path(normalized).name
    stem = Path(basename).stem.strip()
    return stem or "document"


def cleanup_expired_files(*folders: Path, max_age_seconds: int = TEMP_FILE_TTL_SECONDS) -> None:
    """清理过期的临时 docx 文件，避免 /tmp 持续膨胀。"""
    cutoff = time.time() - max_age_seconds

    for folder in folders:
        for path in folder.glob("*.docx"):
            with suppress(OSError):
                if path.is_file() and path.stat().st_mtime < cutoff:
                    path.unlink()


def is_generated_output(filename: str) -> bool:
    """仅允许下载系统生成的输出文件。"""
    return bool(GENERATED_OUTPUT_RE.fullmatch(Path(filename).name))


def get_download_name(filename: str, fallback: str) -> str:
    """规范下载文件名，去掉路径片段并强制使用 .docx 后缀。"""
    normalized = (filename or "").replace("\\", "/").strip()
    basename = Path(normalized).name.strip()
    candidate = basename or fallback
    stem = Path(candidate).stem.strip() or get_display_name(fallback)
    return f"{stem}.docx"


def normalize_format_summary(result) -> dict:
    """兼容旧布尔返回值，统一整理为结构化摘要。"""
    if isinstance(result, dict):
        return {
            "stats": result.get("stats", {}),
            "page_setup": result.get("page_setup", {}),
            "outline": result.get("outline", []),
            "title_text": result.get("title_text", ""),
        }

    return {
        "stats": {},
        "page_setup": {},
        "outline": [],
        "title_text": "",
    }


def truncate_preview_text(text: str, max_length: int = 24) -> str:
    stripped = (text or "").strip()
    if len(stripped) <= max_length:
        return stripped

    return f"{stripped[:max_length - 1].rstrip()}…"


def describe_structure_count(count: int, label: str, unit: str = "个") -> str:
    return f"{label} {count} {unit}"


def build_preview(summary: dict) -> dict:
    """将排版摘要转换为前端可直接展示的结果预览。"""
    stats = summary.get("stats", {})
    page_setup = summary.get("page_setup", {})
    outline = summary.get("outline", [])[:8]
    title_text = (summary.get("title_text") or "").strip()

    top_margin = page_setup.get("margins_cm", {}).get("top", 2.54)
    left_margin = page_setup.get("margins_cm", {}).get("left", 3.18)
    header_text = truncate_preview_text(page_setup.get("header_text", ""))

    structure_bits = []
    for key, label, unit in (
        (ParagraphType.TITLE, "论文标题", "个"),
        (ParagraphType.HEADING_L1, "一级标题", "个"),
        (ParagraphType.HEADING_L2, "二级标题", "个"),
        (ParagraphType.HEADING_L3, "三级标题", "个"),
        (ParagraphType.FIGURE_CAPTION, "图标题", "个"),
        (ParagraphType.TABLE_CAPTION, "表标题", "个"),
        (ParagraphType.SECTION_HEADING, "非编号章节", "个"),
    ):
        count = stats.get(key, 0)
        if count:
            structure_bits.append(describe_structure_count(count, label, unit))

    structure_description = "、".join(structure_bits) or "正文段落已统一为小四、首行缩进和 1.5 倍行距"
    if not header_text:
        header_description = "正文未设置固定默认页眉，页脚仍会插入可更新的自动页码字段"
    elif title_text and page_setup.get("header_text") != title_text:
        header_description = f"页眉显示“{header_text}”，超长标题会自动缩成更适合打印的运行页眉"
    else:
        header_description = f"页眉显示“{header_text}”，页脚插入可更新的自动页码字段"

    reference_count = stats.get(ParagraphType.REFERENCE_ENTRY, 0)
    if reference_count:
        reference_description = f"识别到 {reference_count} 条参考文献，已统一为左对齐、五号字号和更紧凑的文献列表样式"
    else:
        reference_description = "暂未识别到“参考文献”标题，本次仍已完成正文和标题层级排版"

    return {
        "highlights": [
            {
                "eyebrow": "页面设置",
                "title": "A4 页面与规范页边距已应用",
                "description": f"纸张为 A4，上下 {top_margin:.2f} cm，左右 {left_margin:.2f} cm",
            },
            {
                "eyebrow": "页眉页码",
                "title": "页眉与居中页码已自动生成",
                "description": header_description,
            },
            {
                "eyebrow": "结构识别",
                "title": "标题层级已完成识别和套用",
                "description": structure_description,
            },
            {
                "eyebrow": "参考文献",
                "title": "尾部参考文献已单独整理",
                "description": reference_description,
            },
        ],
        "outline": outline,
    }


# ============================================================
# 路由
# ============================================================
@app.after_request
def add_api_response_headers(response):
    if is_api_request():
        response.headers["Cache-Control"] = "no-store"
        response.headers["X-Content-Type-Options"] = "nosniff"

    return response


@app.route("/")
def index():
    """提供主页"""
    return send_from_directory(app.static_folder, "index.html")


@app.route("/api/health", methods=["GET"])
def api_health():
    """健康检查接口，方便部署后探活与诊断。"""
    return jsonify(get_health_payload())


@app.errorhandler(RequestEntityTooLarge)
def handle_file_too_large(_error):
    """统一返回 JSON，避免前端把 413 误判成网络错误。"""
    if is_api_request():
        return json_error("文件大小超过 50MB 限制，请压缩后重试。", 413)

    return _error


@app.errorhandler(404)
def handle_not_found(_error):
    if is_api_request():
        return json_error("接口不存在，请确认请求地址是否正确。", 404)

    return _error


@app.errorhandler(405)
def handle_method_not_allowed(_error):
    if is_api_request():
        return json_error("请求方法不受支持，请检查接口调用方式。", 405)

    return _error


@app.errorhandler(Exception)
def handle_unexpected_error(error):
    if isinstance(error, HTTPException):
        return error

    logger.error(f"未处理异常: {error}", exc_info=True)

    if is_api_request():
        return json_error("服务器内部错误，请稍后重试。", 500)

    raise error


@app.route("/api/format", methods=["POST"])
def api_format():
    """
    接收上传的 .docx 文件，进行排版处理，返回处理结果。
    """
    cleanup_expired_files(UPLOAD_FOLDER, OUTPUT_FOLDER)

    if "file" not in request.files:
        return json_error("未检测到上传文件", 400)

    file = request.files["file"]

    if file.filename == "":
        return json_error("未选择文件", 400)

    if not allowed_file(file.filename):
        return json_error("仅支持 .docx 格式的文件", 400)

    input_path = None
    output_path = None

    try:
        # 生成唯一文件名，避免冲突
        job_id = str(uuid.uuid4())[:8]
        original_name = get_display_name(file.filename)
        
        input_filename = f"{job_id}_input.docx"
        output_filename = f"{job_id}_output.docx"

        input_path = UPLOAD_FOLDER / input_filename
        output_path = OUTPUT_FOLDER / output_filename

        # 保存上传文件
        file.save(str(input_path))
        file_size = input_path.stat().st_size
        logger.info(f"收到文件: {file.filename} ({file_size / 1024:.1f} KB)")

        # 执行排版
        start_time = time.time()
        format_result = format_academic_paper(str(input_path), str(output_path))
        elapsed = time.time() - start_time

        if not format_result:
            with suppress(OSError):
                output_path.unlink()
            return json_error("排版处理失败，请检查文档格式是否正确", 500)

        # 读取排版后的文件大小
        output_size = output_path.stat().st_size
        format_summary = normalize_format_summary(format_result)

        return jsonify({
            "success": True,
            "job_id": job_id,
            "original_name": original_name,
            "download_url": f"/api/download/{output_filename}",
            "download_name": f"{original_name}_排版后.docx",
            "format_summary": format_summary,
            "preview": build_preview(format_summary),
            "stats": {
                "input_size": f"{file_size / 1024:.1f} KB",
                "output_size": f"{output_size / 1024:.1f} KB",
                "elapsed": f"{elapsed:.2f}s",
            }
        })

    except Exception as e:
        if output_path is not None:
            with suppress(OSError):
                output_path.unlink()
        logger.error(f"处理过程中出现异常: {e}", exc_info=True)
        return json_error("服务器内部错误，请稍后重试。", 500)
    finally:
        if input_path is not None:
            with suppress(OSError):
                input_path.unlink()


@app.route("/api/format_text", methods=["POST"])
def api_format_text():
    """
    接收纯文本内容，进行排版处理并生成 .docx，返回下载链接。
    """
    cleanup_expired_files(UPLOAD_FOLDER, OUTPUT_FOLDER)

    output_path = None

    try:
        data = request.get_json(silent=True) or request.form or {}
        text = data.get("text", "").strip()

        if not text:
            return json_error("请输入有效的文字内容", 400)

        job_id = str(uuid.uuid4())[:8]
        original_name = "黏贴文本排版"
        output_filename = f"{job_id}_output.docx"
        output_path = OUTPUT_FOLDER / output_filename

        start_time = time.time()
        format_result = format_academic_paper_from_text(text, str(output_path))
        elapsed = time.time() - start_time

        if not format_result:
            with suppress(OSError):
                output_path.unlink()
            return json_error("排版处理失败，请检查文本内容", 500)

        output_size = output_path.stat().st_size
        format_summary = normalize_format_summary(format_result)

        return jsonify({
            "success": True,
            "job_id": job_id,
            "original_name": original_name,
            "download_url": f"/api/download/{output_filename}",
            "download_name": f"{original_name}_结果.docx",
            "format_summary": format_summary,
            "preview": build_preview(format_summary),
            "stats": {
                "input_size": f"{len(text.encode('utf-8')) / 1024:.1f} KB",
                "output_size": f"{output_size / 1024:.1f} KB",
                "elapsed": f"{elapsed:.2f}s",
            }
        })
    except Exception as e:
        if output_path is not None:
            with suppress(OSError):
                output_path.unlink()
        logger.error(f"文本处理异常: {e}", exc_info=True)
        return json_error("服务器内部错误，请稍后重试。", 500)


@app.route("/api/format_merge", methods=["POST"])
def api_format_merge():
    """接收封面文档和正文文档，排版正文后合并为一个文档。"""
    cleanup_expired_files(UPLOAD_FOLDER, OUTPUT_FOLDER)

    if "cover" not in request.files:
        return json_error("未检测到封面文档", 400)
    if "body" not in request.files:
        return json_error("未检测到正文文档", 400)

    cover_file = request.files["cover"]
    body_file = request.files["body"]

    if cover_file.filename == "":
        return json_error("未选择封面文档", 400)
    if body_file.filename == "":
        return json_error("未选择正文文档", 400)

    if not allowed_file(cover_file.filename):
        return json_error("封面文档仅支持 .docx 格式", 400)
    if not allowed_file(body_file.filename):
        return json_error("正文文档仅支持 .docx 格式", 400)

    cover_path = None
    body_path = None
    output_path = None

    try:
        job_id = str(uuid.uuid4())[:8]
        body_name = get_display_name(body_file.filename)

        cover_path = UPLOAD_FOLDER / f"{job_id}_cover.docx"
        body_path = UPLOAD_FOLDER / f"{job_id}_body.docx"
        output_filename = f"{job_id}_output.docx"
        output_path = OUTPUT_FOLDER / output_filename

        cover_file.save(str(cover_path))
        body_file.save(str(body_path))

        cover_size = cover_path.stat().st_size
        body_size = body_path.stat().st_size
        logger.info(f"收到合并请求: 封面={cover_file.filename} ({cover_size / 1024:.1f} KB), 正文={body_file.filename} ({body_size / 1024:.1f} KB)")

        start_time = time.time()
        result = merge_cover_and_body(str(cover_path), str(body_path), str(output_path))
        elapsed = time.time() - start_time

        if not result:
            with suppress(OSError):
                output_path.unlink()
            return json_error("合并排版失败，请检查文档格式是否正确", 500)

        output_size = output_path.stat().st_size

        response_data = {
            "success": True,
            "job_id": job_id,
            "original_name": body_name,
            "download_url": f"/api/download/{output_filename}",
            "download_name": f"{body_name}_合并排版后.docx",
            "stats": {
                "input_size": f"{(cover_size + body_size) / 1024:.1f} KB",
                "output_size": f"{output_size / 1024:.1f} KB",
                "elapsed": f"{elapsed:.2f}s",
            }
        }

        if isinstance(result, dict):
            response_data["format_result"] = {
                "outline": result.get("outline", []),
                "page_setup": result.get("page_setup", {}),
                "stats": {k: v for k, v in result.get("stats", {}).items()},
                "table_paragraphs": result.get("table_paragraphs", 0),
            }

        return jsonify(response_data)

    except Exception as e:
        if output_path is not None:
            with suppress(OSError):
                output_path.unlink()
        logger.error(f"合并处理过程中出现异常: {e}", exc_info=True)
        return json_error("服务器内部错误，请稍后重试。", 500)
    finally:
        for path in (cover_path, body_path):
            if path is not None:
                with suppress(OSError):
                    path.unlink()


@app.route("/api/download/<filename>")
def api_download(filename):
    """提供排版后文件的下载"""
    cleanup_expired_files(OUTPUT_FOLDER)

    if not is_generated_output(filename):
        return json_error("文件不存在或已过期", 404)

    file_path = OUTPUT_FOLDER / filename

    if not file_path.exists():
        return json_error("文件不存在或已过期", 404)

    download_name = get_download_name(request.args.get("name", ""), filename)

    return send_file(
        str(file_path),
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


# ============================================================
# 启动
# ============================================================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=True)
