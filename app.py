from __future__ import annotations

import os
import tempfile
from pathlib import Path
import unicodedata
import pandas as pd
from html import escape
from flask import Flask, redirect, render_template, request, session, jsonify, url_for, send_from_directory
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
    inspect,
    or_,
    select,
    text,
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
    Column("codigo_barras", String(100), nullable=True),
    Column("sucursal", String(100), nullable=False, server_default="Alajuela"),
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

def asegurar_columna_sucursal() -> None:
    inspector = inspect(engine)
    columnas = {columna["name"] for columna in inspector.get_columns("productos")}

    if "sucursal" not in columnas:
        with engine.begin() as conexion:
            conexion.execute(
                text(
                    "ALTER TABLE productos "
                    "ADD COLUMN sucursal VARCHAR(100) "
                    "NOT NULL DEFAULT 'Alajuela'"
                )
            ) 

def crear_tablas_e_importar_excel() -> None:
    """
    Crea las tablas y copia el Excel a la base de datos una sola vez.

    Si la tabla ya contiene productos, no vuelve a importarlos.
    """
    metadata.create_all(engine)
    asegurar_columna_sucursal()
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
def quitar_acentos(texto: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    ) 
def variantes_singular_plural(texto: str) -> list[str]:
    texto = texto.strip().lower()

    if not texto:
        return []

    palabras = texto.split()
    ultima = palabras[-1]
    variantes_ultima = {ultima}

    vocales = "aeiouáéíóú"

    # Convertir plural a singular
    if ultima.endswith("ces") and len(ultima) > 3:
        variantes_ultima.add(ultima[:-3] + "z")

    elif ultima.endswith("es") and len(ultima) > 2:
        letra_anterior = ultima[-3]

        if letra_anterior not in vocales:
            variantes_ultima.add(ultima[:-2])

    elif ultima.endswith("s") and len(ultima) > 1:
        letra_anterior = ultima[-2]

        if letra_anterior in vocales:
            variantes_ultima.add(ultima[:-1])

    # Convertir singular a plural
    if ultima.endswith("z"):
        variantes_ultima.add(ultima[:-1] + "ces")

    elif ultima[-1] in vocales:
        variantes_ultima.add(ultima + "s")

    else:
        variantes_ultima.add(ultima + "es")

    variantes = []

    for variante_ultima in variantes_ultima:
        palabras_variantes = palabras[:-1] + [variante_ultima]
        variantes.append(" ".join(palabras_variantes))

    return variantes 

def buscar_productos(texto: str = "") -> list[dict]:
    texto = texto.strip()

    consulta = select(
        productos.c.id,
        productos.c.producto,
        productos.c.codigo,
        productos.c.codigo_barras,
        productos.c.ubicacion,
    ).order_by(
        productos.c.producto,
        productos.c.codigo,
        productos.c.ubicacion,
    )

    if texto:
        variantes = variantes_singular_plural(texto)
        condiciones = []

        for variante in variantes:
            patron = f"%{variante}%"

            condiciones.extend(
                [
                    func.lower(productos.c.producto).like(patron),
                    func.lower(productos.c.codigo).like(patron),
                    func.lower(func.coalesce(productos.c.codigo_barras, "")).like(patron),
                    func.lower(productos.c.ubicacion).like(patron),
                ]
            )

        consulta = consulta.where(or_(*condiciones))

    with engine.connect() as conexion:
        filas = conexion.execute(consulta).mappings().all()

    return [dict(fila) for fila in filas] 


def crear_tabla_html(filas: list[dict]) -> str:
    if not filas:
        return """
        <div class="sin-resultados">
            <strong>No se encontraron resultados.</strong>
        </div>
        """

    tarjetas = []

    for fila in filas:
        producto = escape(str(fila.get("producto", "")))
        codigo = escape(str(fila.get("codigo", "")))
        ubicacion = escape(str(fila.get("ubicacion", "")))

        tarjeta = f"""
        <article class="tarjeta-producto">
            <div class="dato-producto">
                <span class="etiqueta">📦 Producto</span>
                <span class="valor producto">{producto}</span>
            </div>

            <div class="dato-producto">
                <span class="etiqueta">🔖 Código</span>
                <span class="valor">{codigo}</span>
            </div>

            <div class="dato-producto">
                <span class="etiqueta">📍 Ubicación</span>
                <span class="valor ubicacion">{ubicacion}</span>
            </div>
        </article>
        """

        tarjetas.append(tarjeta)

    return '<div class="lista-resultados">' + "".join(tarjetas) + "</div>" 



crear_tablas_e_importar_excel()

@app.route("/estado")
def estado():
    respuesta = jsonify({"estado": "listo"})
    respuesta.headers["Access-Control-Allow-Origin"] = "*"
    return respuesta
@app.route("/service-worker.js")
def service_worker():
    return send_from_directory("static", "service-worker.js")
    return app.send_static_file("service-worker.js")
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
@app.route("/actualizar_excel", methods=["GET", "POST"])
def actualizar_excel():
    if not session.get("admin"):
        return redirect(url_for("login"))

    mensaje = ""
    resumen = None

    if request.method == "POST":
        archivo = request.files.get("archivo")

        if not archivo or not archivo.filename:
            mensaje = "Debe seleccionar un archivo de Excel."

        else:
            try:
                sufijo = Path(archivo.filename).suffix.lower()

                temporal = tempfile.NamedTemporaryFile(delete=False, suffix=sufijo)
                ruta_temporal = temporal.name
                temporal.close()
                archivo.save(ruta_temporal)
                session["excel_temporal"] = ruta_temporal
                df = pd.read_excel(ruta_temporal, dtype=str)
                # Limpia y uniforma los nombres de las columnas.
                df.columns = [
                    str(columna).strip().upper()
                    for columna in df.columns
                ]

                # Acepta diferentes formas de escribir los encabezados.
                df = df.rename(
                    columns={
                        "CÓDIGO": "CODIGO",
                        "CÓDIGO DE BARRAS": "CODIGO_BARRAS",
                        "CODIGO DE BARRAS": "CODIGO_BARRAS",
                        "CÓDIGO_BARRAS": "CODIGO_BARRAS",
                        "UBICACIÓN": "UBICACION",
                    }
                )

                columnas_necesarias = {
                    "PRODUCTO",
                    "CODIGO",
                    "CODIGO_BARRAS",
                    "UBICACION",
                    "SUCURSAL",
                }

                faltantes = columnas_necesarias.difference(df.columns)

                if faltantes:
                    mensaje = (
                        "Al archivo le faltan estas columnas: "
                        + ", ".join(sorted(faltantes))
                    )

                else:
                    resumen = {
                        "filas": len(df),
                        "columnas": list(df.columns),
                        "existentes": 0,
                        "nuevas_ubicaciones": 0,
                        "productos_nuevos": 0,
                        "sin_cambios": 0,
                    }

                    with engine.begin() as conexion:
                        for _, fila in df.iterrows():
                            codigo = limpiar_texto(fila.get("CODIGO"))
                            ubicacion = limpiar_texto(fila.get("UBICACION"))
                            sucursal = (
                                limpiar_texto(fila.get("SUCURSAL"))
                                or "Alajuela"
                            )

                            if not codigo or not ubicacion:
                                continue

                            filas_existentes = conexion.execute(
                                select(
                                    productos.c.codigo,
                                    productos.c.ubicacion,
                                    productos.c.sucursal,
                                ).where(
                                    func.lower(productos.c.codigo)
                                    == codigo.lower()
                                )
                            ).fetchall()

                            if not filas_existentes:
                                resumen["productos_nuevos"] += 1
                                continue

                            resumen["existentes"] += 1

                            misma_ubicacion = any(
                                limpiar_texto(
                                    fila_db.ubicacion
                                ).lower()
                                == ubicacion.lower()
                                and limpiar_texto(
                                    fila_db.sucursal
                                ).lower()
                                == sucursal.lower()
                                for fila_db in filas_existentes
                            )

                            if misma_ubicacion:
                                resumen["sin_cambios"] += 1
                            else:
                                resumen["nuevas_ubicaciones"] += 1

                    mensaje = (
                        f"Archivo leído correctamente. "
                        f"Se encontraron {len(df)} filas."
                    )

            except Exception as error:
                mensaje = f"No se pudo leer el archivo: {error}"

    return render_template(
        "actualizar_excel.html",
        mensaje=mensaje,
        resumen=resumen,
    ) 
@app.route("/aplicar_cambios_excel", methods=["POST"])
def aplicar_cambios_excel():
    if not session.get("admin"):
        return redirect(url_for("login"))

    ruta_temporal = session.get("excel_temporal")

    if not ruta_temporal or not os.path.exists(ruta_temporal):
        return render_template(
            "actualizar_excel.html",
            mensaje="No hay un archivo revisado para aplicar.",
            resumen=None,
        )

    agregadas = 0
    nuevos = 0
    sin_cambios = 0

    try:
        df = pd.read_excel(ruta_temporal, dtype=str)

        df.columns = [
            str(columna).strip().upper()
            for columna in df.columns
        ]

        df = df.rename(
            columns={
                "CÓDIGO": "CODIGO",
                "CÓDIGO DE BARRAS": "CODIGO_BARRAS",
                "CODIGO DE BARRAS": "CODIGO_BARRAS",
                "CÓDIGO_BARRAS": "CODIGO_BARRAS",
                "UBICACIÓN": "UBICACION",
            }
        )

        with engine.begin() as conexion:
            for _, fila in df.iterrows():
                producto = limpiar_texto(fila.get("PRODUCTO"))
                codigo = limpiar_texto(fila.get("CODIGO"))
                codigo_barras = limpiar_texto(
                    fila.get("CODIGO_BARRAS")
                )
                ubicacion = limpiar_texto(fila.get("UBICACION"))
                sucursal = (
                    limpiar_texto(fila.get("SUCURSAL"))
                    or "Alajuela"
                )

                if not producto or not codigo or not ubicacion:
                    continue

                existentes = conexion.execute(
                    select(productos).where(
                        func.lower(productos.c.codigo)
                        == codigo.lower()
                    )
                ).fetchall()

                misma_ubicacion = any(
                    limpiar_texto(
                        registro.ubicacion
                    ).lower() == ubicacion.lower()
                    and limpiar_texto(
                        registro.sucursal
                    ).lower() == sucursal.lower()
                    for registro in existentes
                )

                if misma_ubicacion:
                    sin_cambios += 1
                    continue

                conexion.execute(
                    insert(productos).values(
                        producto=producto,
                        codigo=codigo,
                        codigo_barras=codigo_barras,
                        sucursal=sucursal,
                        ubicacion=ubicacion,
                    )
                )

                if existentes:
                    agregadas += 1
                else:
                    nuevos += 1

        mensaje = (
            f"Cambios aplicados correctamente. "
            f"Nuevas ubicaciones: {agregadas}. "
            f"Productos nuevos: {nuevos}. "
            f"Sin cambios: {sin_cambios}."
        )

        session.pop("excel_temporal", None)

        try:
            os.remove(ruta_temporal)
        except OSError:
            pass

        return render_template(
            "actualizar_excel.html",
            mensaje=mensaje,
            resumen=None,
        )

    except Exception as error:
        return render_template(
            "actualizar_excel.html",
            mensaje=f"No se pudieron aplicar los cambios: {error}",
            resumen=None,
        ) 


@app.route("/nuevo", methods=["GET", "POST"])
def nuevo():
    if not session.get("admin"):
        return redirect(url_for("login"))

    if request.method == "POST":
        producto = request.form.get("producto", "").strip()
        codigo = request.form.get("codigo", "").strip()
        codigo_barras = request.form.get("codigo_barras", "").strip() 
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
                    codigo_barras=codigo_barras or None,
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
        codigo_barras_nuevo = request.form.get("codigo_barras", "").strip()
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
            if codigo_barras_nuevo:
                cambios["codigo_barras"] = codigo_barras_nuevo
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
@app.route("/eliminar_ubicacion", methods=["GET", "POST"])
def eliminar_ubicacion():
    if not session.get("admin"):
        return redirect(url_for("login"))

    if request.method == "POST":
        ubicacion_buscar = request.form.get(
            "ubicacion_buscar",
            "",
        ).strip()

        id_registro = request.form.get(
            "id_registro",
            "",
        ).strip()

        accion = request.form.get(
            "accion",
            "",
        ).strip()

        if not ubicacion_buscar:
            return render_template(
                "eliminar_ubicacion.html",
                error="Debe escribir una ubicación.",
            )

        with engine.begin() as conexion:
            coincidencias = conexion.execute(
                select(
                    productos.c.id,
                    productos.c.producto,
                    productos.c.codigo,
                    productos.c.ubicacion,
                )
                .where(
                    func.lower(productos.c.ubicacion)
                    == ubicacion_buscar.lower()
                )
                .order_by(
                    productos.c.producto,
                    productos.c.codigo,
                )
            ).mappings().all()

            if not coincidencias:
                return render_template(
                    "eliminar_ubicacion.html",
                    error="No existen productos en esa ubicación.",
                    ubicacion_buscar=ubicacion_buscar,
                )

            if accion == "eliminar_toda":
                cantidad = len(coincidencias)

                conexion.execute(
                    delete(productos).where(
                        func.lower(productos.c.ubicacion)
                        == ubicacion_buscar.lower()
                    )
                )

                return render_template(
                    "eliminar_ubicacion.html",
                    mensaje=(
                        f"Se eliminó completamente la ubicación "
                        f"{ubicacion_buscar}, junto con "
                        f"{cantidad} registro(s)."
                    ),
                    ubicacion_buscar="",
                    coincidencias=[],
                )

            if not id_registro:
                return render_template(
                    "eliminar_ubicacion.html",
                    coincidencias=coincidencias,
                    ubicacion_buscar=ubicacion_buscar,
                )

            try:
                id_seleccionado = int(id_registro)
            except ValueError:
                return render_template(
                    "eliminar_ubicacion.html",
                    error="La selección no es válida.",
                    coincidencias=coincidencias,
                    ubicacion_buscar=ubicacion_buscar,
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
                    "eliminar_ubicacion.html",
                    error=(
                        "El registro seleccionado no corresponde "
                        "a esa ubicación."
                    ),
                    coincidencias=coincidencias,
                    ubicacion_buscar=ubicacion_buscar,
                )

            conexion.execute(
    delete(productos).where(
        productos.c.id == fila["id"]
    )
) 

            codigo_modificado = fila["codigo"]

        with engine.connect() as conexion:
            coincidencias_restantes = conexion.execute(
                select(
                    productos.c.id,
                    productos.c.producto,
                    productos.c.codigo,
                    productos.c.ubicacion,
                )
                .where(
                    func.lower(productos.c.ubicacion)
                    == ubicacion_buscar.lower()
                )
                .order_by(
                    productos.c.producto,
                    productos.c.codigo,
                )
            ).mappings().all()

        return render_template(
            "eliminar_ubicacion.html",
            mensaje=(
    f"Se eliminó el registro con código "
    f"{codigo_modificado} correctamente."
), 
            coincidencias=coincidencias_restantes,
            ubicacion_buscar=ubicacion_buscar,
        )

    return render_template("eliminar_ubicacion.html") 



@app.route("/logout")
def logout():
    session.pop("admin", None)
    return redirect(url_for("inicio"))


if __name__ == "__main__":
    app.run(debug=True) 
