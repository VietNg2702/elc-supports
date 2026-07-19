"""Local API server to process DOCX files into XLSX using docx2excel logic."""

from __future__ import annotations

import os
import uuid
from pathlib import Path

from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename

from docx2excel import write_to_excel


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
ALLOWED_EXTENSIONS = {".docx"}

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)

# For production, set CORS_ORIGINS to your GitHub Pages URL, e.g.:
# https://<username>.github.io
cors_origins = os.getenv("CORS_ORIGINS", "*")
origin_list = [origin.strip() for origin in cors_origins.split(",") if origin.strip()]
CORS(app, resources={r"/process": {"origins": origin_list or "*"}})


@app.get("/health")
def health() -> tuple[dict[str, str], int]:
    return {"status": "ok"}, 200


@app.post("/process")
def process_file():
    if "inputFile" not in request.files:
        return jsonify({"error": "No file uploaded in 'inputFile'"}), 400

    uploaded = request.files["inputFile"]
    if not uploaded.filename:
        return jsonify({"error": "Empty filename"}), 400

    filename = secure_filename(uploaded.filename)
    suffix = Path(filename).suffix.lower()
    if suffix not in ALLOWED_EXTENSIONS:
        return jsonify({"error": "Only .docx files are supported"}), 400

    token = uuid.uuid4().hex
    input_name = f"{token}_{filename}"
    output_name = f"{Path(filename).stem}_converted.xlsx"

    input_path = UPLOAD_DIR / input_name
    output_path = OUTPUT_DIR / f"{token}_{output_name}"

    uploaded.save(input_path)

    try:
        write_to_excel(input_path, output_path)
        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as exc:  # noqa: BLE001
        return jsonify({"error": str(exc)}), 500


if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=True)
