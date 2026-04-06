from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZIP_DEFLATED, ZipFile

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.utils import secure_filename

from reportreader import build_report_artifacts, default_poppler_path, default_tesseract_path


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.lower().endswith(".pdf")


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/api/process")
def process_pdf():
    upload = request.files.get("pdf_file")
    if upload is None or not upload.filename:
        return jsonify({"error": "Choose a PDF file first."}), 400
    if not allowed_file(upload.filename):
        return jsonify({"error": "Only PDF files are supported."}), 400

    poppler_path = default_poppler_path()
    tesseract_path = default_tesseract_path()
    if poppler_path is None:
        return jsonify({"error": "Poppler could not be found on the server."}), 500
    if tesseract_path is None:
        return jsonify({"error": "Tesseract could not be found on the server."}), 500

    safe_name = secure_filename(Path(upload.filename).stem) or "report"

    try:
        with TemporaryDirectory() as temp_dir:
            pdf_path = Path(temp_dir) / f"{safe_name}.pdf"
            upload.save(pdf_path)

            artifacts = build_report_artifacts(
                pdf_path=pdf_path,
                poppler_path=poppler_path,
                tesseract_path=tesseract_path,
                dpi=220,
            )

            zip_buffer = BytesIO()
            with ZipFile(zip_buffer, mode="w", compression=ZIP_DEFLATED) as archive:
                archive.writestr(f"{safe_name}_part_totals.csv", artifacts.csv_bytes)
                archive.writestr(f"{safe_name}_part_totals.xlsx", artifacts.excel_bytes)
            zip_buffer.seek(0)

            return send_file(
                zip_buffer,
                mimetype="application/zip",
                as_attachment=True,
                download_name=f"{safe_name}_outputs.zip",
            )
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": f"Processing failed: {exc}"}), 500


if __name__ == "__main__":
    app.run(debug=True)
