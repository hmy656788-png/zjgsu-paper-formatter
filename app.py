#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
学术论文自动排版工具 - Web 服务端
基于 Flask 提供文件上传、排版处理、结果下载功能。
"""

import os
import uuid
import time
import logging
from pathlib import Path

from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename

from format_paper import format_academic_paper, format_academic_paper_from_text

# ============================================================
# 应用配置
# ============================================================
app = Flask(__name__, static_folder="static", static_url_path="/static")
CORS(app)  # 允许跨域请求（适应前后台分离的云部署架构）

# 在 Serverless 环境中，通常只有 /tmp 目录有写入权限
UPLOAD_FOLDER = Path("/tmp") / "uploads"
OUTPUT_FOLDER = Path("/tmp") / "outputs"
UPLOAD_FOLDER.mkdir(exist_ok=True, parents=True)
OUTPUT_FOLDER.mkdir(exist_ok=True, parents=True)

app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB 上传限制

ALLOWED_EXTENSIONS = {".docx"}

logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)


def allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


# ============================================================
# 路由
# ============================================================
@app.route("/")
def index():
    """提供主页"""
    return send_from_directory("static", "index.html")


@app.route("/api/format", methods=["POST"])
def api_format():
    """
    接收上传的 .docx 文件，进行排版处理，返回处理结果。
    """
    if "file" not in request.files:
        return jsonify({"success": False, "error": "未检测到上传文件"}), 400

    file = request.files["file"]

    if file.filename == "":
        return jsonify({"success": False, "error": "未选择文件"}), 400

    if not allowed_file(file.filename):
        return jsonify({"success": False, "error": "仅支持 .docx 格式的文件"}), 400

    try:
        # 生成唯一文件名，避免冲突
        job_id = str(uuid.uuid4())[:8]
        original_name = Path(secure_filename(file.filename)).stem or "document"
        
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
        success = format_academic_paper(str(input_path), str(output_path))
        elapsed = time.time() - start_time

        if not success:
            return jsonify({
                "success": False,
                "error": "排版处理失败，请检查文档格式是否正确"
            }), 500

        # 读取排版后的文件大小
        output_size = output_path.stat().st_size

        return jsonify({
            "success": True,
            "job_id": job_id,
            "original_name": original_name,
            "download_url": f"/api/download/{output_filename}",
            "download_name": f"{original_name}_排版后.docx",
            "stats": {
                "input_size": f"{file_size / 1024:.1f} KB",
                "output_size": f"{output_size / 1024:.1f} KB",
                "elapsed": f"{elapsed:.2f}s",
            }
        })

    except Exception as e:
        logger.error(f"处理过程中出现异常: {e}", exc_info=True)
        return jsonify({
            "success": False,
            "error": f"服务器内部错误: {str(e)}"
        }), 500


@app.route("/api/format_text", methods=["POST"])
def api_format_text():
    """
    接收纯文本内容，进行排版处理并生成 .docx，返回下载链接。
    """
    try:
        data = request.json or request.form
        text = data.get("text", "").strip()

        if not text:
            return jsonify({"success": False, "error": "请输入有效的文字内容"}), 400

        job_id = str(uuid.uuid4())[:8]
        original_name = "黏贴文本排版"
        output_filename = f"{job_id}_output.docx"
        output_path = OUTPUT_FOLDER / output_filename

        start_time = time.time()
        success = format_academic_paper_from_text(text, str(output_path))
        elapsed = time.time() - start_time

        if not success:
            return jsonify({"success": False, "error": "排版处理失败，请检查文本内容"}), 500

        output_size = output_path.stat().st_size

        return jsonify({
            "success": True,
            "job_id": job_id,
            "original_name": original_name,
            "download_url": f"/api/download/{output_filename}",
            "download_name": f"{original_name}_结果.docx",
            "stats": {
                "input_size": f"{len(text.encode('utf-8')) / 1024:.1f} KB",
                "output_size": f"{output_size / 1024:.1f} KB",
                "elapsed": f"{elapsed:.2f}s",
            }
        })
    except Exception as e:
        logger.error(f"文本处理异常: {e}", exc_info=True)
        return jsonify({"success": False, "error": f"服务器内部错误: {str(e)}"}), 500


@app.route("/api/download/<filename>")
def api_download(filename):
    """提供排版后文件的下载"""
    file_path = OUTPUT_FOLDER / filename

    if not file_path.exists():
        return jsonify({"success": False, "error": "文件不存在或已过期"}), 404

    download_name = request.args.get("name", filename)

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
