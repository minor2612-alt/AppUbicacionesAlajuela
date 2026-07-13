from __future__ import annotations

import os
from pathlib import Path

import pandas as pd
from flask import Flask, redirect, render_template, request, session, url_for
from sqlalchemy import (
    Column,
    Integer,
    MetaData,
    String,
    Table,
    Text,
    create_engine,
    delete,
    func,
    insert,
    or_,
    select,
    update,
)
from werkzeug.security import check_password_hash, generate_password_hash 

app = Flask(__name__)

# En Render se usa la dirección de Neon guardada en DATABASE_URL.
# En la computadora se usa una base SQLite local para poder hacer pruebas.
DATABASE_URL = os.getenv(
    "DATABASE_URL",
    "sqlite:///ubicaciones_neon_local.db",
)

app.secret_key = os.getenv(
    "SECRET_KEY",
    "construplaza_alajuela_2026",
)

ADMIN_USER = os.getenv("ADMIN_USER", "admin")
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "1234")

EXCEL_FILE = Path("data/Ubicaciones Alajuela Glide.xlsx")

engine = create_engine(
    DATABASE_URL,
    pool_pre_ping=True,
)

metadata = MetaData()

productos = Table(
    "productos",
    metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("producto", String(250), nullable=False),
    Column("codigo", String(100), nullable=False),
    Column("ubicacion", String(150), nullable=False),
)

configuracion = Table(
    "configuracion",
    metadata,
    Column("clave", String(100), primary_key=True),
    Column("valor", Text, nullable=False),
)


def limpiar_texto(valor) -> str:
    """Convierte valores vacíos o NaN en texto limpio."""
    if valor is None or pd.isna(valor):
        return ""

    texto = str(valor).strip()

    if texto.lower() == "nan":
        return ""

    # Evita que códigos numéricos de Excel terminen como 12345.0
    if texto.endswith(".0"):
        parte_numerica = texto[:-2]
        if parte_numerica.isdigit():
            texto = parte_numerica

    return texto


def crear_tablas_e_importar_excel() -> None:
    """
    Crea las tablas y copia el Excel a la base de datos una sola vez.

    Si la tabla ya contiene productos, no vuelve a importarlos.
    """
    metadata.create_all(engine)

    with engine.begin() as conexion:
        importacion = conexion.execute(
            select(configuracion.c.valor).where(
                configuracion.c.clave == "excel_importado"
            )
        ).scalar_one_or_none()

        if importacion == "si":
            return

        cantidad = conexion.execute(
            select(func.count()).select_from(productos)
        ).scalar_one()

        if cantidad > 0:
            conexion.execute(
                insert(configuracion).values(
                    clave="excel_importado",
                    valor="si",
                )
            )
            return

        if not EXCEL_FILE.exists():
            conexion.execute(
                insert(configuracion).values(
                    clave="excel_importado",
                    valor="si",
                )
            )
            return

        df = pd.read_excel(EXCEL_FILE, dtype=str)
        df.columns = [str(columna).strip().upper() for columna in df.columns]

        columnas_necesarias = {"PRODUCTO", "CODIGO", "UBICACION"}

        if not columnas_necesarias.issubset(df.columns):
            faltantes = columnas_necesarias.difference(df.columns)
            raise RuntimeError(
                "Al Excel le faltan estas columnas: "
                + ", ".join(sorted(faltantes))
            )

        registros = []

        for _, fila in df.iterrows():
            producto = limpiar_texto(fila.get("PRODUCTO"))
            codigo = limpiar_texto(fila.get("CODIGO"))
            ubicacion = limpiar_texto(fila.get("UBICACION"))

            # Ignora únicamente las filas completamente vacías.
            if not producto and not codigo and not ubicacion:
                continue

            registros.append(
                {
                    "producto": producto,
                    "codigo": codigo,
                    "ubicacion": ubicacion,
                }
            )

        if registros:
            conexion.execute(insert(productos), registros)

        conexion.execute(
            insert(configuracion).values(
                clave="excel_importado",
                valor="si",
            )
        )


def buscar_productos(texto: str = "") -> list[dict]:
    texto = texto.strip()

    consulta = select(
        productos.c.id,
        productos.c.producto,
        productos.c.codigo,
        productos.c.ubicacion,
    ).order_by(
        productos.c.producto,
        productos.c.codigo,
        productos.c.ubicacion,
    )

    if texto:
        patron = f"%{texto.lower()}%"

        consulta = consulta.where(
            or_(
                func.lower(productos.c.producto).like(patron),
                func.lower(productos.c.codigo).like(patron),
                func.lower(productos.c.ubicacion).like(patron),
            )
        )

    with engine.connect() as conexion:
        filas = conexion.execute(consulta).mappings().all()

    return [dict(fila) for fila in filas]


def crear_tabla_html(filas: list[dict]) -> str:
    if not filas:
        return "<p><b>No se encontraron resultados.</b></p>"

    df = pd.DataFrame(
        [
            {
                "PRODUCTO": fila["producto"],
                "CODIGO": fila["codigo"],
                "UBICACION": fila["ubicacion"],
            }
            for fila in filas
        ]
    )

    return df.to_html(
        index=False,
        escape=True,
        classes="tabla",
        border=0,
    )


crear_tablas_e_importar_excel()


@app.route("/")
def inicio():
    busqueda = request.args.get("buscar", "").strip()
    resultado = ""

    if busqueda:
        filas = buscar_productos(busqueda)
        resultado = crear_tabla_html(filas)

    return render_template(
        "index.html",
        busqueda=busqueda,
        resultado=resultado,
    )


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        usuario = request.form.get("usuario", "").strip()
        password = request.form.get("password", "").strip()

        if usuario == ADMIN_USER and password == ADMIN_PASSWORD:
            session["admin"] = True
            return redirect(url_for("admin"))

        return render_template(
            "login.html",
            error="Usuario o contraseña incorrectos.",
        )

    return render_template("login.html")


@app.route("/admin")
def admin():
    if not session.get("admin"):
        return redirect(url_for("login"))

    return render_template("admin.html")

@app.route("/cambiar_password", methods=["GET", "POST"])
def cambiar_password():
    global ADMIN_PASSWORD
    if not session.get("admin"):
        return redirect(url_for("login"))

    mensaje = ""

    if request.method == "POST":
        actual = request.form.get("actual", "")
        nueva = request.form.get("nueva", "")
        confirmar = request.form.get("confirmar", "")

        if actual != ADMIN_PASSWORD:
            mensaje = "La contraseña actual es incorrecta."
        elif nueva != confirmar:
            mensaje = "Las contraseñas nuevas no coinciden."
        else:
            os.environ["ADMIN_PASSWORD"] = nueva
            
            ADMIN_PASSWORD = nueva
            mensaje = "Contraseña cambiada correctamente."

    return render_template("cambiar_password.html", mensaje=mensaje) 

@app.route("/inventario")
def inventario():
    if not session.get("admin"):
        return redirect(url_for("login"))

    buscar = request.args.get("buscar", "").strip()
    filas = buscar_productos(buscar)
    tabla = crear_tabla_html(filas)

    return render_template(
        "inventario.html",
        tabla=tabla,
        buscar=buscar,
    )


@app.route("/nuevo", methods=["GET", "POST"])
def nuevo():
    if not session.get("admin"):
        return redirect(url_for("login"))

    if request.method == "POST":
        producto = request.form.get("producto", "").strip()
        codigo = request.form.get("codigo", "").strip()
        ubicacion = request.form.get("ubicacion", "").strip()

        if not producto or not codigo or not ubicacion:
            return render_template(
                "nuevo.html",
                error="Producto, código y ubicación son obligatorios.",
            )

        # Se permiten códigos repetidos porque un mismo producto
        # puede estar en varias ubicaciones.
        with engine.begin() as conexion:
            conexion.execute(
                insert(productos).values(
                    producto=producto,
                    codigo=codigo,
                    ubicacion=ubicacion,
                )
            )

        return redirect(url_for("admin"))

    return render_template("nuevo.html")


@app.route("/editar", methods=["GET", "POST"])
def editar():
    if not session.get("admin"):
        return redirect(url_for("login"))

    if request.method == "POST":
        codigo_buscar = request.form.get("codigo_buscar", "").strip()
        id_registro = request.form.get("id_registro", "").strip()

        producto_nuevo = request.form.get("producto", "").strip()
        codigo_nuevo = request.form.get("codigo", "").strip()
        ubicacion_nueva = request.form.get("ubicacion", "").strip()

        if not codigo_buscar:
            return render_template(
                "editar.html",
                error="Debe escribir el código que desea editar.",
            )

        with engine.begin() as conexion:
            coincidencias = conexion.execute(
                select(
                    productos.c.id,
                    productos.c.producto,
                    productos.c.codigo,
                    productos.c.ubicacion,
                ).where(
                    func.lower(productos.c.codigo)
                    == codigo_buscar.lower()
                )
            ).mappings().all()

            if not coincidencias:
                return render_template(
                    "editar.html",
                    error="No existe ese código.",
                    codigo_buscar=codigo_buscar,
                )

            # Si existen varias ubicaciones y todavía no se ha
            # seleccionado una fila, mostramos la lista.
            if len(coincidencias) > 1 and not id_registro:
                return render_template(
                    "editar.html",
                    coincidencias=coincidencias,
                    codigo_buscar=codigo_buscar,
                    producto=producto_nuevo,
                    codigo=codigo_nuevo,
                    ubicacion=ubicacion_nueva,
                )

            # Si se seleccionó una fila, buscamos ese ID exacto.
            if id_registro:
                try:
                    id_seleccionado = int(id_registro)
                except ValueError:
                    return render_template(
                        "editar.html",
                        error="La selección no es válida.",
                        coincidencias=coincidencias,
                        codigo_buscar=codigo_buscar,
                    )

                fila = next(
                    (
                        registro
                        for registro in coincidencias
                        if registro["id"] == id_seleccionado
                    ),
                    None,
                )

                if fila is None:
                    return render_template(
                        "editar.html",
                        error="El registro seleccionado no corresponde a ese código.",
                        coincidencias=coincidencias,
                        codigo_buscar=codigo_buscar,
                    )
            else:
                # Si solamente existe una coincidencia.
                fila = coincidencias[0]

            cambios = {}

            if producto_nuevo:
                cambios["producto"] = producto_nuevo

            if codigo_nuevo:
                cambios["codigo"] = codigo_nuevo

            if ubicacion_nueva:
                cambios["ubicacion"] = ubicacion_nueva

            if not cambios:
                return render_template(
                    "editar.html",
                    error="Debe escribir al menos un dato nuevo.",
                    coincidencias=(
                        coincidencias
                        if len(coincidencias) > 1
                        else None
                    ),
                    codigo_buscar=codigo_buscar,
                )

            conexion.execute(
                update(productos)
                .where(productos.c.id == fila["id"])
                .values(**cambios)
            )

        return redirect(url_for("admin"))

    return render_template("editar.html")


@app.route("/eliminar", methods=["GET", "POST"])
def eliminar():
    if not session.get("admin"):
        return redirect(url_for("login"))

    if request.method == "POST":
        codigo = request.form.get("codigo", "").strip()
        id_registro = request.form.get("id_registro", "").strip()

        if not codigo:
            return render_template(
                "eliminar.html",
                error="Debe escribir un código.",
            )

        with engine.begin() as conexion:
            coincidencias = conexion.execute(
                select(
                    productos.c.id,
                    productos.c.producto,
                    productos.c.codigo,
                    productos.c.ubicacion,
                ).where(
                    func.lower(productos.c.codigo) == codigo.lower()
                )
            ).mappings().all()

            if not coincidencias:
                return render_template(
                    "eliminar.html",
                    error="No existe un producto con ese código.",
                    codigo=codigo,
                )

            # Primera etapa: mostrar todos los registros encontrados.
            if not id_registro:
                return render_template(
                    "eliminar.html",
                    coincidencias=coincidencias,
                    codigo=codigo,
                )

            try:
                id_seleccionado = int(id_registro)
            except ValueError:
                return render_template(
                    "eliminar.html",
                    error="La selección no es válida.",
                    coincidencias=coincidencias,
                    codigo=codigo,
                )

            fila = next(
                (
                    registro
                    for registro in coincidencias
                    if registro["id"] == id_seleccionado
                ),
                None,
            )

            if fila is None:
                return render_template(
                    "eliminar.html",
                    error="El registro seleccionado no corresponde a ese código.",
                    coincidencias=coincidencias,
                    codigo=codigo,
                )

            conexion.execute(
                delete(productos).where(
                    productos.c.id == fila["id"]
                )
            )

        return redirect(url_for("admin"))

    return render_template("eliminar.html") 


@app.route("/logout")
def logout():
    session.pop("admin", None)
    return redirect(url_for("inicio"))


if __name__ == "__main__":
    app.run(debug=True) 
