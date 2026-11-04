# -*- coding: utf-8 -*-
"""PPT 换皮工作台 - 本地网站
上传 .pptx/.ppt -> 自动套用北大红卡片模板 -> 下载成品
"""
import os
import re
import uuid
import shutil
import threading
import traceback

from flask import Flask, request, jsonify, send_file, render_template

import engine

BASE = os.path.dirname(os.path.abspath(__file__))
UPLOAD = os.path.join(BASE, "uploads")
OUT = os.path.join(BASE, "outputs")
def _find_soffice():
    """自动探测 LibreOffice 的 soffice 可执行文件（macOS / Linux / 自定义挂载）。"""
    c = shutil.which("soffice")
    if c:
        return c
    for p in (
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/Volumes/LibreOffice/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice",
        "/usr/local/bin/soffice",
    ):
        if os.path.exists(p):
            return p
    return None


SOFFICE = _find_soffice()
os.makedirs(UPLOAD, exist_ok=True)
os.makedirs(OUT, exist_ok=True)

JOBS = {}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024


def _ppt_to_pptx(src):
    """老 .ppt 先用 LibreOffice 转成 .pptx（同目录，同名）。"""
    if not SOFFICE:
        raise RuntimeError("未找到 LibreOffice（soffice），无法转换 .ppt，请改用 .pptx 上传")
    d = os.path.dirname(src)
    env = dict(os.environ, FONTCONFIG_FILE="/tmp/fonts.conf")
    import subprocess
    subprocess.run(
        [SOFFICE, "--headless", "--convert-to", "pptx", "--outdir", d, src],
        env=env, check=True, capture_output=True, timeout=300,
    )
    out = os.path.splitext(src)[0] + ".pptx"
    if not os.path.exists(out):
        raise RuntimeError(".ppt 转换失败，请改用 .pptx 上传")
    return out


def run_job(jid, src, out_name):
    def prog(m):
        JOBS[jid]["msg"] = m

    try:
        if src.lower().endswith(".ppt"):
            JOBS[jid]["msg"] = "转换旧版 .ppt 格式…"
            src = _ppt_to_pptx(src)
        safe = re.sub(r'[\\/:*?"<>|]', "_", out_name)
        out = os.path.join(OUT, f"{jid}_{safe}")
        engine.plan_and_build(src, out, progress=prog)
        JOBS[jid].update(status="done", out=out, name=safe)
        # 清理中间产物
        if src != out and os.path.dirname(src) == UPLOAD:
            try:
                os.remove(src)
            except OSError:
                pass
    except Exception as e:  # noqa: BLE001
        traceback.print_exc()
        JOBS[jid].update(status="error", msg=str(e))


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/convert", methods=["POST"])
def convert():
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify(error="请选择一个 PPT 文件"), 400
    if not f.filename.lower().endswith((".pptx", ".ppt")):
        return jsonify(error="只支持 .pptx / .ppt 文件"), 400
    jid = uuid.uuid4().hex[:8]
    src = os.path.join(UPLOAD, f"{jid}_{f.filename}")
    f.save(src)
    out_name = os.path.splitext(f.filename)[0] + "-北大模板版.pptx"
    JOBS[jid] = {"status": "running", "msg": "读取文件…", "name": out_name}
    threading.Thread(target=run_job, args=(jid, src, out_name), daemon=True).start()
    return jsonify(job=jid)


@app.route("/api/status/<jid>")
def status(jid):
    j = JOBS.get(jid)
    if not j:
        return jsonify(status="unknown"), 404
    return jsonify(j)


@app.route("/download/<jid>")
def download(jid):
    j = JOBS.get(jid)
    if not j or j.get("status") != "done":
        return "任务不存在或未完成", 404
    return send_file(j["out"], as_attachment=True, download_name=j["name"])


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5177, debug=False)
