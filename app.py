from flask import Flask, request, render_template, redirect, url_for, session
import pandas as pd
import shutil
import os
from datetime import datetime

app = Flask(__name__)

EXCEL_FILE = "data/Ubicaciones Alajuela Glide.xlsx"
app.secret_key = "construplaza_alajuela_2026"


def respaldar_excel():
    carpeta_respaldo = "respaldos"

    if not os.path.exists(carpeta_respaldo):
        os.makedirs(carpeta_respaldo)

    fecha = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    nombre_respaldo = f"ubicaciones_respaldo_{fecha}.xlsx"

    shutil.copy(EXCEL_FILE, os.path.join(carpeta_respaldo, nombre_respaldo))


@app.route("/")
def inicio():
    df = pd.read_excel(EXCEL_FILE)

    busqueda = request.args.get("buscar", "").strip()
    resultado = ""

    if busqueda:
        filtro = df.astype(str).apply(
            lambda fila: fila.str.contains(busqueda, case=False, na=False)
        ).any(axis=1)

        encontrados = df[filtro]

        if not encontrados.empty:
            resultado = encontrados[["PRODUCTO", "CODIGO", "UBICACION"]].to_html(index=False)
        else:
            resultado = "<p><b>No se encontraron resultados.</b></p>"

    return render_template("index.html", busqueda=busqueda, resultado=resultado)


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        usuario = request.form.get("usuario", "").strip()
        password = request.form.get("password", "").strip()

        if usuario == "admin" and password == "1234":
            session["admin"] = True
            return redirect(url_for("admin"))

        return render_template("login.html", error="Usuario o contraseña incorrectos.")

    return render_template("login.html")


@app.route("/admin")
def admin():
    if not session.get("admin"):
        return redirect(url_for("login"))

    return render_template("admin.html")


@app.route("/inventario")
def inventario():
    if not session.get("admin"):
        return redirect(url_for("login"))

    buscar = request.args.get("buscar", "").strip()
    df = pd.read_excel(EXCEL_FILE)

    if buscar:
        filtro = df.astype(str).apply(
            lambda fila: fila.str.contains(buscar, case=False, na=False)
        ).any(axis=1)
        df = df[filtro]

    tabla = df[["PRODUCTO", "CODIGO", "UBICACION"]].to_html(index=False)

    return render_template("inventario.html", tabla=tabla, buscar=buscar)


@app.route("/nuevo", methods=["GET", "POST"])
def nuevo():
    if not session.get("admin"):
        return redirect(url_for("login"))

    if request.method == "POST":
        producto = request.form["producto"].strip()
        codigo = request.form["codigo"].strip()
        ubicacion = request.form["ubicacion"].strip()

        df = pd.read_excel(EXCEL_FILE)

        nuevo_producto = pd.DataFrame({
            "PRODUCTO": [producto],
            "CODIGO": [codigo],
            "UBICACION": [ubicacion]
        })

        respaldar_excel()

        df = pd.concat([df, nuevo_producto], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)

        return redirect(url_for("admin"))

    return render_template("nuevo.html")


@app.route("/editar", methods=["GET", "POST"])
def editar():
    if not session.get("admin"):
        return redirect(url_for("login"))

    if request.method == "POST":
        codigo_buscar = request.form["codigo_buscar"].strip()
        producto = request.form["producto"].strip()
        codigo = request.form["codigo"].strip()
        ubicacion = request.form["ubicacion"].strip()

        df = pd.read_excel(EXCEL_FILE)
        df["CODIGO"] = df["CODIGO"].astype(str)

        filtro = df["CODIGO"].str.strip() == codigo_buscar

        if not filtro.any():
            return render_template("editar.html", error="No existe ese código.")

        indice = df[filtro].index[0]

        if producto:
            df.at[indice, "PRODUCTO"] = producto

        if codigo:
            df.at[indice, "CODIGO"] = codigo

        if ubicacion:
            df.at[indice, "UBICACION"] = ubicacion

        respaldar_excel()

        df.to_excel(EXCEL_FILE, index=False)

        return redirect(url_for("admin"))

    return render_template("editar.html")


@app.route("/eliminar", methods=["GET", "POST"])
def eliminar():
    if not session.get("admin"):
        return redirect(url_for("login"))

    if request.method == "POST":
        codigo = request.form["codigo"].strip()

        df = pd.read_excel(EXCEL_FILE)
        df["CODIGO"] = df["CODIGO"].astype(str)

        if codigo not in df["CODIGO"].str.strip().values:
            return render_template("eliminar.html", error="No existe un producto con ese código.")

        respaldar_excel()

        df = df[df["CODIGO"].str.strip() != codigo]
        df.to_excel(EXCEL_FILE, index=False)

        return redirect(url_for("admin"))

    return render_template("eliminar.html")


@app.route("/logout")
def logout():
    session.pop("admin", None)
    return redirect(url_for("inicio"))


if __name__ == "__main__":
    app.run(debug=True) 
