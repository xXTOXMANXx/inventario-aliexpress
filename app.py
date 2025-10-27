from flask import Flask, render_template, request, send_file
import os
import subprocess

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():

    # 1) ELIMINAR archivo antiguo
    if os.path.exists("stock_actual.xlsx"):
        os.remove("stock_actual.xlsx")

    file = request.files['file']
    file.save("stock_actual.xlsx")  # Renombra automáticamente al requerido

    # 2) Ejecuta actualizar_inventario.py (actualiza stock_real.xlsx desde GOOGLE SHEETS)
    subprocess.run(["python", "actualizar_inventario.py"], check=True)

    # 3) Ejecuta actualizar_stock.py (actualiza el archivo cargado)
    subprocess.run(["python", "actualizar_stock.py"], check=True)

    # 4) Cuando termine, descarga el archivo resultante
    return send_file(
        "stock_actual_actualizado.xlsx",
        as_attachment=True
    )

if __name__ == '__main__':
    app.run(debug=True)
