
from flask import Flask, render_template, request, redirect, send_file
import pandas as pd
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
DATA_FILE = 'clientes.xlsx'

COLUMNAS = [
    "Cliente", "Estado", "Rubro", "Teléfono", "Mail", "Comentario", "Zona", "Domicilio",
    "Localidad", "Provincia", "Comercial", "Contacto", "Departamento Oficina", "Último contacto"
]

@app.route("/")
def index():
    return render_template("formulario.html")

@app.route("/guardar", methods=["POST"])
def guardar():
    datos = {campo: request.form.get(campo, "") for campo in COLUMNAS}
    nuevo_df = pd.DataFrame([datos])
    if os.path.exists(DATA_FILE):
        df = pd.read_excel(DATA_FILE)
        df = pd.concat([df, nuevo_df], ignore_index=True)
    else:
        df = nuevo_df
    df.to_excel(DATA_FILE, index=False)
    return redirect("/")

@app.route("/importar_excel", methods=["GET", "POST"])
def importar_excel():
    if request.method == "POST":
        archivo = request.files["archivo"]
        if archivo and archivo.filename.endswith((".xls", ".xlsx")):
            df = pd.read_excel(archivo)
            if os.path.exists(DATA_FILE):
                df_existente = pd.read_excel(DATA_FILE)
                df = pd.concat([df_existente, df], ignore_index=True)
            df.to_excel(DATA_FILE, index=False)
            return redirect("/")
    return render_template("importar_excel.html")

@app.route("/listado")
def listado():
    if os.path.exists(DATA_FILE):
        df = pd.read_excel(DATA_FILE)
        return df.to_html(index=False)
    return "No hay datos cargados."

@app.route("/descargar_excel")
def descargar_excel():
    if os.path.exists(DATA_FILE):
        return send_file(DATA_FILE, as_attachment=True)
    return "No hay archivo para descargar."

if __name__ == "__main__":
    app.run(debug=True)
