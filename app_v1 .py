from flask import Flask, request
import pandas as pd

app = Flask(__name__)

EXCEL_FILE = "data/Ubicaciones Alajuela Glide.xlsx"


@app.route("/")
def inicio():
    try:
        df = pd.read_excel(EXCEL_FILE)

        busqueda = request.args.get("buscar", "").strip()

        resultado = ""

        if busqueda:
            filtro = df.astype(str).apply(
                lambda fila: fila.str.contains(busqueda, case=False, na=False)
            ).any(axis=1)

            encontrados = df[filtro]

            if len(encontrados) > 0:
                resultado = encontrados[["PRODUCTO", "CODIGO", "UBICACION"]].to_html(index=False)
            else:
                resultado = "<p><b>No se encontraron resultados.</b></p>"

        return f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>App Ubicaciones Alajuela</title>

<style>
body {{
    font-family: Arial, sans-serif;
    background-color: #f5f5f5;
    margin: 0;
    padding: 30px;
}}

.contenedor {{
    max-width: 900px;
    margin: auto;
    background: white;
    padding: 30px;
    border-radius: 12px;
    box-shadow: 0 4px 15px rgba(0,0,0,.15);
}}

h1 {{
    color: #0b5ed7;
}}

input {{
    font-size: 18px;
    padding: 8px;
    width: 300px;
}}

button {{
    font-size: 18px;
    padding: 8px 16px;
    background-color: #0b5ed7;
    color: white;
    border: none;
    border-radius: 5px;
    cursor: pointer;
}}

button:hover {{
    background-color: #084298;
}}

table {{
    border-collapse: collapse;
    width: 100%;
    margin-top: 20px;
}}

th, td {{
    border: 1px solid #ccc;
    padding: 10px;
    text-align: left;
}}

th {{
    background-color: #0b5ed7;
    color: white;
}}
</style>

</head>

<body>

<div class="contenedor">

<h1>App Ubicaciones Alajuela</h1>

<form method="GET">
    <input type="text" name="buscar" placeholder="Buscar producto o código" value="{busqueda}">
    <button type="submit">Buscar</button>
</form>

<br>

{resultado}

</div>

</body>
</html>
"""

    except Exception as e:
        return f"<h2>Error al leer el Excel:</h2><pre>{e}</pre>"


if __name__ == "__main__":
    app.run(debug=True) 
