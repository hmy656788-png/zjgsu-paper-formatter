#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
学术论文自动排版工具 - Web 服务端
基于 Flask 提供文件上传、排版处理、结果下载功能。
"""

import re
import json
import uuid
import time
import threading
import logging
from contextlib import suppress
from pathlib import Path

from flask import Flask, request, jsonify, send_file, send_from_directory, Response
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

try:
    from docxcompose.composer import Composer as _DocxComposer  # noqa: F401
    HAS_DOCXCOMPOSE = True
except ImportError:  # pragma: no cover - optional dependency fallback
    HAS_DOCXCOMPOSE = False

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
JOB_TTL_SECONDS = TEMP_FILE_TTL_SECONDS
PROCESSING_STEP_LABELS = {
    1: "解析文档结构",
    2: "识别标题层级",
    3: "应用排版规则",
    4: "生成输出文档",
}
COVER_FORM_FIELDS = (
    "title",
    "cover_title",
    "course_title",
    "college",
    "teacher",
    "class_name",
    "student_name",
    "student_id",
    "school_name",
)

logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)
PROGRESS_JOBS = {}
PROGRESS_JOBS_LOCK = threading.Lock()

if CORS is None:
    logger.warning("未检测到 Flask-Cors，已跳过 CORS 配置；同源部署不受影响。")


class JobProcessingError(Exception):
    """包装可直接返回给用户的处理失败信息。"""

    def __init__(self, message: str, status_code: int = 500):
        super().__init__(message)
        self.message = message
        self.status_code = status_code


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
        "features": {
            "cover_merge": HAS_DOCXCOMPOSE,
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


def cleanup_expired_jobs(max_age_seconds: int = JOB_TTL_SECONDS) -> None:
    """清理已结束且过期的异步任务状态。"""
    cutoff = time.time() - max_age_seconds

    with PROGRESS_JOBS_LOCK:
        expired_job_ids = [
            job_id
            for job_id, job in PROGRESS_JOBS.items()
            if job.get("updated_at", 0) < cutoff and job.get("status") in {"done", "error"}
        ]
        for job_id in expired_job_ids:
            PROGRESS_JOBS.pop(job_id, None)


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


def extract_cover_info(payload) -> dict | None:
    """从表单或 JSON 载荷中提取自动封面信息。"""
    if not payload:
        return None

    enabled = str(payload.get("generate_cover", "")).strip().lower() in {"1", "true", "yes", "on"}
    if not enabled:
        return None

    cover_info = {}
    for field in COVER_FORM_FIELDS:
        raw_value = payload.get(field)
        if raw_value is None:
            continue
        value = str(raw_value).strip()
        if value:
            cover_info[field] = value

    return cover_info


def normalize_format_summary(result) -> dict:
    """兼容旧布尔返回值，统一整理为结构化摘要。"""
    if isinstance(result, dict):
        return {
            "stats": result.get("stats", {}),
            "page_setup": result.get("page_setup", {}),
            "outline": result.get("outline", []),
            "title_text": result.get("title_text", ""),
            "table_paragraphs": result.get("table_paragraphs", 0),
            "equation_paragraphs": result.get("equation_paragraphs", 0),
            "resized_images": result.get("resized_images", 0),
            "cover_generated": result.get("cover_generated", False),
        }

    return {
        "stats": {},
        "page_setup": {},
        "outline": [],
        "title_text": "",
        "table_paragraphs": 0,
        "equation_paragraphs": 0,
        "resized_images": 0,
        "cover_generated": False,
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
    cover_generated = bool(summary.get("cover_generated"))
    equation_paragraphs = int(summary.get("equation_paragraphs", 0) or 0)

    top_margin = page_setup.get("margins_cm", {}).get("top", 2.54)
    left_margin = page_setup.get("margins_cm", {}).get("left", 3.18)
    header_text = truncate_preview_text(page_setup.get("header_text", ""))

    structure_bits = []
    for key, label, unit in (
        (ParagraphType.TITLE, "论文标题", "个"),
        (ParagraphType.ENGLISH_ABSTRACT_HEADING, "英文摘要", "个"),
        (ParagraphType.HEADING_L1, "一级标题", "个"),
        (ParagraphType.HEADING_L2, "二级标题", "个"),
        (ParagraphType.HEADING_L3, "三级标题", "个"),
        (ParagraphType.FIGURE_CAPTION, "图标题", "个"),
        (ParagraphType.TABLE_CAPTION, "表标题", "个"),
        (ParagraphType.SECTION_HEADING, "非编号章节", "个"),
        (ParagraphType.REFERENCES_HEADING, "参考文献标题", "个"),
    ):
        count = stats.get(key, 0)
        if count:
            structure_bits.append(describe_structure_count(count, label, unit))
    if equation_paragraphs:
        structure_bits.append(describe_structure_count(equation_paragraphs, "公式段落", "个"))

    structure_description = "、".join(structure_bits) or "正文段落已统一为小四、首行缩进和 1.5 倍行距"
    page_description = f"纸张为 A4，上下 {top_margin:.2f} cm，左右 {left_margin:.2f} cm"
    if cover_generated:
        page_description += "，并在首页插入了自动生成的课程论文封面"
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
                "description": page_description,
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


def create_progress_job(kind: str) -> dict:
    """创建一个可供 SSE 订阅的异步处理任务。"""
    job_id = str(uuid.uuid4())[:8]
    job = {
        "id": job_id,
        "kind": kind,
        "status": "queued",
        "created_at": time.time(),
        "updated_at": time.time(),
        "events": [],
        "next_event_id": 0,
        "condition": threading.Condition(),
        "result": None,
        "error": None,
        "events_url": f"/api/jobs/{job_id}/events",
        "result_url": f"/api/jobs/{job_id}/result",
    }

    with PROGRESS_JOBS_LOCK:
        PROGRESS_JOBS[job_id] = job

    return job


def get_progress_job(job_id: str):
    with PROGRESS_JOBS_LOCK:
        return PROGRESS_JOBS.get(job_id)


def _append_job_event_locked(job: dict, event: str, data: dict) -> None:
    job["next_event_id"] += 1
    job["events"].append(
        {
            "id": job["next_event_id"],
            "event": event,
            "data": data,
        }
    )
    job["updated_at"] = time.time()
    job["condition"].notify_all()


def emit_job_progress(job: dict, step: int, message: str, detail: str | None = None) -> None:
    with job["condition"]:
        if job["status"] == "queued":
            job["status"] = "running"

        payload = {
            "job_id": job["id"],
            "step": step,
            "step_label": PROCESSING_STEP_LABELS.get(step, "处理中"),
            "message": message,
        }
        if detail:
            payload["detail"] = detail

        _append_job_event_locked(job, "progress", payload)


def complete_progress_job(job: dict, result_payload: dict) -> None:
    with job["condition"]:
        job["status"] = "done"
        job["result"] = result_payload
        job["error"] = None
        _append_job_event_locked(
            job,
            "complete",
            {
                "job_id": job["id"],
                "message": "排版完成",
                "result_url": job["result_url"],
            },
        )


def fail_progress_job(job: dict, message: str, status_code: int = 500) -> None:
    with job["condition"]:
        job["status"] = "error"
        job["error"] = {
            "message": message,
            "status_code": status_code,
        }
        _append_job_event_locked(
            job,
            "failed",
            {
                "job_id": job["id"],
                "message": message,
            },
        )


def build_job_progress_callback(job: dict):
    def callback(payload: dict):
        emit_job_progress(
            job,
            step=int(payload.get("step", 1)),
            message=payload.get("message", "正在处理文档"),
            detail=payload.get("detail"),
        )

    return callback


def launch_background_job(job: dict, task_name: str, work_fn, cleanup_paths=()) -> None:
    def runner():
        progress_callback = build_job_progress_callback(job)
        try:
            result_payload = work_fn(progress_callback)
            complete_progress_job(job, result_payload)
        except JobProcessingError as exc:
            logger.error(f"{task_name}失败: {exc.message}")
            fail_progress_job(job, exc.message, exc.status_code)
        except Exception as exc:  # pragma: no cover - 作为最终兜底
            logger.error(f"{task_name}出现异常: {exc}", exc_info=True)
            fail_progress_job(job, "服务器内部错误，请稍后重试。", 500)
        finally:
            for path in cleanup_paths:
                with suppress(OSError):
                    Path(path).unlink()

    threading.Thread(target=runner, daemon=True).start()


def build_async_job_response(job: dict) -> dict:
    return {
        "success": True,
        "job_id": job["id"],
        "status": job["status"],
        "events_url": job["events_url"],
        "result_url": job["result_url"],
    }


def build_success_response(
    *,
    job_id: str,
    original_name: str,
    output_filename: str,
    download_name: str,
    input_size_text: str,
    output_size_text: str,
    elapsed: float,
    format_result,
    include_format_result_alias: bool = False,
) -> dict:
    format_summary = normalize_format_summary(format_result)
    response_data = {
        "success": True,
        "job_id": job_id,
        "original_name": original_name,
        "download_url": f"/api/download/{output_filename}",
        "download_name": download_name,
        "format_summary": format_summary,
        "preview": build_preview(format_summary),
        "stats": {
            "input_size": input_size_text,
            "output_size": output_size_text,
            "elapsed": f"{elapsed:.2f}s",
        },
    }

    if include_format_result_alias:
        response_data["format_result"] = format_summary

    return response_data


def process_uploaded_document(input_path: Path, original_name: str, job_id: str, progress_callback=None, cover_info=None) -> dict:
    output_filename = f"{job_id}_output.docx"
    output_path = OUTPUT_FOLDER / output_filename

    try:
        input_size_bytes = input_path.stat().st_size
        start_time = time.time()
        format_result = format_academic_paper(
            str(input_path),
            str(output_path),
            progress_callback=progress_callback,
            cover_info=cover_info,
        )
        elapsed = time.time() - start_time

        if not format_result:
            with suppress(OSError):
                output_path.unlink()
            raise JobProcessingError("排版处理失败，请检查文档格式是否正确", 500)

        output_size_bytes = output_path.stat().st_size
        return build_success_response(
            job_id=job_id,
            original_name=original_name,
            output_filename=output_filename,
            download_name=f"{original_name}_排版后.docx",
            input_size_text=f"{input_size_bytes / 1024:.1f} KB",
            output_size_text=f"{output_size_bytes / 1024:.1f} KB",
            elapsed=elapsed,
            format_result=format_result,
        )
    except FileNotFoundError as exc:
        raise JobProcessingError(f"上传文件不存在：{exc}", 500) from exc


def process_text_document(text: str, job_id: str, progress_callback=None, cover_info=None) -> dict:
    output_filename = f"{job_id}_output.docx"
    output_path = OUTPUT_FOLDER / output_filename
    original_name = "黏贴文本排版"

    start_time = time.time()
    format_result = format_academic_paper_from_text(
        text,
        str(output_path),
        progress_callback=progress_callback,
        cover_info=cover_info,
    )
    elapsed = time.time() - start_time

    if not format_result:
        with suppress(OSError):
            output_path.unlink()
        raise JobProcessingError("排版处理失败，请检查文本内容", 500)

    output_size_bytes = output_path.stat().st_size
    return build_success_response(
        job_id=job_id,
        original_name=original_name,
        output_filename=output_filename,
        download_name=f"{original_name}_结果.docx",
        input_size_text=f"{len(text.encode('utf-8')) / 1024:.1f} KB",
        output_size_text=f"{output_size_bytes / 1024:.1f} KB",
        elapsed=elapsed,
        format_result=format_result,
    )


def process_merged_document(cover_path: Path, body_path: Path, body_name: str, job_id: str, progress_callback=None) -> dict:
    output_filename = f"{job_id}_output.docx"
    output_path = OUTPUT_FOLDER / output_filename

    cover_size_bytes = cover_path.stat().st_size
    body_size_bytes = body_path.stat().st_size
    start_time = time.time()
    format_result = merge_cover_and_body(
        str(cover_path),
        str(body_path),
        str(output_path),
        progress_callback=progress_callback,
    )
    elapsed = time.time() - start_time

    if not format_result:
        with suppress(OSError):
            output_path.unlink()
        raise JobProcessingError("合并排版失败，请检查文档格式是否正确", 500)

    output_size_bytes = output_path.stat().st_size
    return build_success_response(
        job_id=job_id,
        original_name=body_name,
        output_filename=output_filename,
        download_name=f"{body_name}_合并排版后.docx",
        input_size_text=f"{(cover_size_bytes + body_size_bytes) / 1024:.1f} KB",
        output_size_text=f"{output_size_bytes / 1024:.1f} KB",
        elapsed=elapsed,
        format_result=format_result,
        include_format_result_alias=True,
    )


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
    cleanup_expired_jobs()

    if "file" not in request.files:
        return json_error("未检测到上传文件", 400)

    file = request.files["file"]

    if file.filename == "":
        return json_error("未选择文件", 400)

    if not allowed_file(file.filename):
        return json_error("仅支持 .docx 格式的文件", 400)

    input_path = None
    try:
        job_id = str(uuid.uuid4())[:8]
        original_name = get_display_name(file.filename)
        input_path = UPLOAD_FOLDER / f"{job_id}_input.docx"
        cover_info = extract_cover_info(request.form)

        file.save(str(input_path))
        file_size = input_path.stat().st_size
        logger.info(f"收到文件: {file.filename} ({file_size / 1024:.1f} KB)")

        response_data = process_uploaded_document(input_path, original_name, job_id, cover_info=cover_info)
        return jsonify(response_data)

    except JobProcessingError as exc:
        return json_error(exc.message, exc.status_code)
    except Exception as e:
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
    cleanup_expired_jobs()

    try:
        data = request.get_json(silent=True) or request.form or {}
        text = data.get("text", "").strip()
        cover_info = extract_cover_info(data)

        if not text:
            return json_error("请输入有效的文字内容", 400)

        job_id = str(uuid.uuid4())[:8]
        original_name = "黏贴文本排版"
        return jsonify(process_text_document(text, job_id, cover_info=cover_info))
    except JobProcessingError as exc:
        return json_error(exc.message, exc.status_code)
    except Exception as e:
        logger.error(f"文本处理异常: {e}", exc_info=True)
        return json_error("服务器内部错误，请稍后重试。", 500)


@app.route("/api/format_merge", methods=["POST"])
def api_format_merge():
    """接收封面文档和正文文档，排版正文后合并为一个文档。"""
    cleanup_expired_files(UPLOAD_FOLDER, OUTPUT_FOLDER)
    cleanup_expired_jobs()

    if not HAS_DOCXCOMPOSE:
        return json_error("当前服务未安装封面合并组件，请改用“自动生成模板封面”或补装 docxcompose。", 503)

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
    try:
        job_id = str(uuid.uuid4())[:8]
        body_name = get_display_name(body_file.filename)

        cover_path = UPLOAD_FOLDER / f"{job_id}_cover.docx"
        body_path = UPLOAD_FOLDER / f"{job_id}_body.docx"

        cover_file.save(str(cover_path))
        body_file.save(str(body_path))

        cover_size = cover_path.stat().st_size
        body_size = body_path.stat().st_size
        logger.info(f"收到合并请求: 封面={cover_file.filename} ({cover_size / 1024:.1f} KB), 正文={body_file.filename} ({body_size / 1024:.1f} KB)")

        response_data = process_merged_document(cover_path, body_path, body_name, job_id)
        return jsonify(response_data)

    except JobProcessingError as exc:
        return json_error(exc.message, exc.status_code)
    except Exception as e:
        logger.error(f"合并处理过程中出现异常: {e}", exc_info=True)
        return json_error("服务器内部错误，请稍后重试。", 500)
    finally:
        for path in (cover_path, body_path):
            if path is not None:
                with suppress(OSError):
                    path.unlink()


@app.route("/api/format_async", methods=["POST"])
def api_format_async():
    """创建正文排版异步任务，并通过 SSE 推送真实进度。"""
    cleanup_expired_files(UPLOAD_FOLDER, OUTPUT_FOLDER)
    cleanup_expired_jobs()

    if "file" not in request.files:
        return json_error("未检测到上传文件", 400)

    file = request.files["file"]

    if file.filename == "":
        return json_error("未选择文件", 400)

    if not allowed_file(file.filename):
        return json_error("仅支持 .docx 格式的文件", 400)

    job = create_progress_job("format")
    input_path = UPLOAD_FOLDER / f"{job['id']}_input.docx"
    original_name = get_display_name(file.filename)
    cover_info = extract_cover_info(request.form)

    try:
        file.save(str(input_path))
        file_size = input_path.stat().st_size
        logger.info(f"收到异步排版请求: {file.filename} ({file_size / 1024:.1f} KB)")

        emit_job_progress(job, 1, "文件上传完成，正在准备排版", f"{file_size / 1024:.1f} KB")
        launch_background_job(
            job,
            "异步正文排版",
            lambda progress_callback: process_uploaded_document(
                input_path,
                original_name,
                job["id"],
                progress_callback=progress_callback,
                cover_info=cover_info,
            ),
            cleanup_paths=(input_path,),
        )
        return jsonify(build_async_job_response(job)), 202
    except Exception:
        with PROGRESS_JOBS_LOCK:
            PROGRESS_JOBS.pop(job["id"], None)
        with suppress(OSError):
            input_path.unlink()
        raise


@app.route("/api/format_text_async", methods=["POST"])
def api_format_text_async():
    """创建黏贴文本排版异步任务。"""
    cleanup_expired_files(UPLOAD_FOLDER, OUTPUT_FOLDER)
    cleanup_expired_jobs()

    data = request.get_json(silent=True) or request.form or {}
    text = data.get("text", "").strip()
    cover_info = extract_cover_info(data)

    if not text:
        return json_error("请输入有效的文字内容", 400)

    job = create_progress_job("format_text")
    emit_job_progress(job, 1, "文本内容已接收，正在创建排版任务")
    launch_background_job(
        job,
        "异步文本排版",
        lambda progress_callback: process_text_document(
            text,
            job["id"],
            progress_callback=progress_callback,
            cover_info=cover_info,
        ),
    )
    return jsonify(build_async_job_response(job)), 202


@app.route("/api/format_merge_async", methods=["POST"])
def api_format_merge_async():
    """创建封面 + 正文合并排版异步任务。"""
    cleanup_expired_files(UPLOAD_FOLDER, OUTPUT_FOLDER)
    cleanup_expired_jobs()

    if not HAS_DOCXCOMPOSE:
        return json_error("当前服务未安装封面合并组件，请改用“自动生成模板封面”或补装 docxcompose。", 503)

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

    job = create_progress_job("format_merge")
    cover_path = UPLOAD_FOLDER / f"{job['id']}_cover.docx"
    body_path = UPLOAD_FOLDER / f"{job['id']}_body.docx"
    body_name = get_display_name(body_file.filename)

    try:
        cover_file.save(str(cover_path))
        body_file.save(str(body_path))
        cover_size = cover_path.stat().st_size
        body_size = body_path.stat().st_size
        logger.info(
            f"收到异步合并请求: 封面={cover_file.filename} ({cover_size / 1024:.1f} KB), "
            f"正文={body_file.filename} ({body_size / 1024:.1f} KB)"
        )

        emit_job_progress(
            job,
            1,
            "封面与正文上传完成，正在准备合并排版",
            f"总大小 {(cover_size + body_size) / 1024:.1f} KB",
        )
        launch_background_job(
            job,
            "异步合并排版",
            lambda progress_callback: process_merged_document(
                cover_path,
                body_path,
                body_name,
                job["id"],
                progress_callback=progress_callback,
            ),
            cleanup_paths=(cover_path, body_path),
        )
        return jsonify(build_async_job_response(job)), 202
    except Exception:
        with PROGRESS_JOBS_LOCK:
            PROGRESS_JOBS.pop(job["id"], None)
        for path in (cover_path, body_path):
            with suppress(OSError):
                path.unlink()
        raise


@app.route("/api/jobs/<job_id>/events", methods=["GET"])
def api_job_events(job_id):
    """输出指定任务的 Server-Sent Events 进度流。"""
    cleanup_expired_jobs()
    job = get_progress_job(job_id)
    if job is None:
        return json_error("任务不存在或已过期", 404)

    last_event_id = request.headers.get("Last-Event-ID", "").strip()
    try:
        sent_count = max(0, int(last_event_id))
    except ValueError:
        sent_count = 0

    def generate():
        delivered_count = sent_count

        while True:
            keepalive = False
            with job["condition"]:
                notified = job["condition"].wait_for(
                    lambda: len(job["events"]) > delivered_count or job["status"] in {"done", "error"},
                    timeout=15,
                )
                pending_events = job["events"][delivered_count:]
                finished = job["status"] in {"done", "error"}
                if not notified and not pending_events and not finished:
                    keepalive = True

            if keepalive:
                yield ": keep-alive\n\n"
                continue

            for event in pending_events:
                delivered_count += 1
                payload = json.dumps(event["data"], ensure_ascii=False)
                yield f"id: {event['id']}\nevent: {event['event']}\ndata: {payload}\n\n"

            if finished and delivered_count >= len(job["events"]):
                break

    return Response(
        generate(),
        mimetype="text/event-stream",
        headers={
            "Cache-Control": "no-store",
            "X-Accel-Buffering": "no",
        },
    )


@app.route("/api/jobs/<job_id>/result", methods=["GET"])
def api_job_result(job_id):
    """读取异步任务的最终结果。"""
    cleanup_expired_jobs()
    job = get_progress_job(job_id)
    if job is None:
        return json_error("任务不存在或已过期", 404)

    if job["status"] in {"queued", "running"}:
        return jsonify({"success": False, "status": "processing"}), 202

    if job["status"] == "error":
        error_info = job.get("error") or {}
        return json_error(error_info.get("message", "处理失败"), error_info.get("status_code", 500))

    return jsonify(job["result"])


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
