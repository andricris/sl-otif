#!/usr/bin/env python3
"""
SL & OTIF — Flask backend (openpyxl, pure Python).

Endpoints:
  GET  /api/health              -> {"status":"ok"}
  POST /api/process             -> multipart form:
                                     files : one or more .xlsx
                                     tol   : tolerance days (optional, default 7)
                                   returns JSON:
                                     { "results": [ { name, out_name, stats, data_b64 } ] }

Run:
  pip install -r requirements.txt
  python app.py                 # http://localhost:5000
"""
import base64, os, sys, traceback

from flask import Flask, jsonify, request

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from sl_otif_core import process_bytes, TOL  # noqa: E402

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024  # 64 MB


def _cors(resp):
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
    resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    return resp


@app.after_request
def add_cors(resp):
    return _cors(resp)


@app.route("/api/health", methods=["GET", "OPTIONS"])
def health():
    if request.method == "OPTIONS":
        return ("", 204)
    return jsonify(status="ok", tol_default=TOL)


@app.route("/api/process", methods=["POST", "OPTIONS"])
def process():
    if request.method == "OPTIONS":
        return ("", 204)
    files = request.files.getlist("files")
    if not files:
        return jsonify(error="no files uploaded (field name must be 'files')"), 400
    try:
        tol = int(request.form.get("tol", TOL))
    except (TypeError, ValueError):
        tol = TOL
    tol = max(0, min(60, tol))

    results = []
    for f in files:
        fname = f.filename or "upload.xlsx"
        if not fname.lower().endswith((".xlsx", ".xlsm")):
            results.append({"name": fname, "error": "not an .xlsx file"})
            continue
        try:
            out_bytes, stats, out_filename = process_bytes(f.read(), tol, fname)
            results.append({
                "name": fname,
                "out_name": out_filename,
                "stats": {k: v for k, v in stats.items() if k not in ("agg",)},
                "agg": stats["agg"],
                "data_b64": base64.b64encode(out_bytes).decode("ascii"),
            })
        except Exception as e:  # noqa: BLE001
            traceback.print_exc()
            results.append({"name": fname, "error": str(e)})
    return jsonify(results=results, tol=tol)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    debug = os.environ.get("DEBUG", "0") == "1"
    app.run(host="0.0.0.0", port=port, debug=debug, threaded=True)
