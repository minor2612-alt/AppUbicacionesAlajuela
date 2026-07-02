from flask import Flask
import pandas as pd

app = Flask(__name__)

EXCEL_FILE = "data/Ubicaciones Alajuela Glide.xlsx"

@app.route("/")
def inicio():
    try:
        df = pd.read_excel(EXCEL_FILE)
        return f"""
        <h1>App Ubicaciones Alajuela</h1>
        <p>✅ El archivo Excel fue leído correctamente.</p>
        <p>Total de productos: <strong>{len(df)}</strong></p>
        """
    except Exception as e:
        return f"<h2>Error al leer el Excel:</h2><pre>{e}</pre>"

if __name__ == "__main__":
    app.run(debug=True) 