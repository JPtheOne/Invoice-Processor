import os, io, tempfile, datetime as dt
from werkzeug.utils import secure_filename
from processor import process_cfdi, unzip_folder  # importing our functions
from flask import Flask, render_template, request, send_file, jsonify, make_response

app = Flask(__name__, template_folder="../templates")

XML_EXT = {".xml"}
ZIP_EXT = {".zip"}

def ext(path): 
    return os.path.splitext(path)[1].lower()

@app.get("/")
def index():
    return render_template("index.html")

@app.post("/process-folder")
def process_folder():
    files = request.files.getlist("folder")  
    output_name = (request.form.get("output_name", "")).strip()

    if not files:
        return jsonify({"error": "No se subieron archivos"}), 400
    if not output_name:
        output_name = f"Excel_final_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}"
    output_filename = f"{secure_filename(output_name)}.xlsx"

    counters = {"Total": 0, "I/E": 0, "P": 0, "N": 0, "Desconocido": 0}

    try:
        with tempfile.TemporaryDirectory() as workdir:
            raw_dir = os.path.join(workdir, "raw")
            unzipped_dir = os.path.join(workdir, "unzipped")
            os.makedirs(raw_dir, exist_ok=True)
            os.makedirs(unzipped_dir, exist_ok=True)

            saved = []
            for f in files:
                if not f.filename: 
                    continue
                fname = secure_filename(f.filename)
                path = os.path.join(raw_dir, fname)
                f.save(path)
                saved.append(path)

            xmls = []
            for p in saved:
                if ext(p) in ZIP_EXT:
                    extracted = unzip_folder(p, unzipped_dir)
                    xmls.extend([q for q in extracted if ext(q) in XML_EXT])

            xmls.extend([p for p in saved if ext(p) in XML_EXT])

            if not xmls:
                return jsonify({"error": "No se encontraron XML"}), 400

            out_path = os.path.join(workdir, output_filename)
            for xml in xmls:
                counters["Total"] += 1
                process_cfdi(xml, out_path, counters)

            with open(out_path, "rb") as f:
                payload = io.BytesIO(f.read())

        # ✅ Construimos la respuesta con headers personalizados (sin "/" en nombres)
        response = make_response(send_file(
            payload,
            as_attachment=True,
            download_name=output_filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ))

        response.headers["X-Counter-Total"] = str(counters["Total"])
        response.headers["X-Counter-IE"] = str(counters["I/E"])
        response.headers["X-Counter-P"] = str(counters["P"])
        response.headers["X-Counter-N"] = str(counters["N"])
        response.headers["X-Counter-Desconocido"] = str(counters["Desconocido"])

        # Habilitar que el frontend pueda leer esos headers
        response.headers["Access-Control-Expose-Headers"] = (
            "X-Counter-Total, X-Counter-IE, X-Counter-P, X-Counter-N, X-Counter-Desconocido, Content-Disposition"
        )

        return response

    except Exception as e:
        return jsonify({"error": f"Error: {e}"}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
