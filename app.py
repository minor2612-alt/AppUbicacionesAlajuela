from __future__ import annotations

import os
import tempfile
from pathlib import Path
import unicodedata
import pandas as pd
from html import escape
from urllib.parse import quote
from flask import Flask, redirect, render_template, request, session, jsonify, url_for, send_from_directory
from sqlalchemy import (
 
    Boolean,
    Column,
    DateTime,
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

# Etapa 1 del sistema de usuarios y permisos.
# Estas tablas se crean sin modificar todavía el login administrativo actual.
usuarios = Table(
    "usuarios",
    metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("usuario", String(100), nullable=False, unique=True),
    Column("nombre", String(150), nullable=False),
    Column("password_hash", String(255), nullable=False),
    Column("activo", Boolean, nullable=False, server_default=text("true")),
    Column("es_superadmin", Boolean, nullable=False, server_default=text("false")),
    Column("puede_agregar", Boolean, nullable=False, server_default=text("false")),
    Column("puede_editar", Boolean, nullable=False, server_default=text("false")),
    Column("puede_eliminar", Boolean, nullable=False, server_default=text("false")),
    Column("puede_eliminar_ubicacion", Boolean, nullable=False, server_default=text("false")),
    Column("puede_actualizar_excel", Boolean, nullable=False, server_default=text("false")),
    Column("creado_en", DateTime, nullable=False, server_default=func.now()),
)

historial_cambios = Table(
    "historial_cambios",
    metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("usuario_id", Integer, nullable=True),
    # Se guarda también el nombre del usuario para conservar la trazabilidad
    # aunque posteriormente esa cuenta sea desactivada o eliminada.
    Column("usuario", String(100), nullable=False),
    Column("accion", String(100), nullable=False),
    Column("entidad", String(100), nullable=False),
    Column("registro_id", Integer, nullable=True),
    Column("codigo", String(100), nullable=True),
    Column("datos_antes", Text, nullable=True),
    Column("datos_despues", Text, nullable=True),
    Column("detalle", Text, nullable=True),
    Column("fecha_hora", DateTime, nullable=False, server_default=func.now()),
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
        mensaje_whatsapp = (
            f"Producto: {producto}\n"
            f"Código: {codigo}\n"
            f"Ubicación: {ubicacion}"
        )
        enlace_whatsapp = f"https://wa.me/50686928249?text={quote(mensaje_whatsapp)}"
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
            <a href="{enlace_whatsapp}"
               target="_blank"
               class="boton-whatsapp">
                📱 Enviar por WhatsApp
            </a> 
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

        # Primero comprobamos el administrador principal
        if usuario == ADMIN_USER and password == ADMIN_PASSWORD:
            session["admin"] = True
            session["usuario"] = ADMIN_USER
            session["es_superadmin"] = True
            return redirect(url_for("admin"))

        # Después comprobamos los usuarios creados
        with engine.connect() as conexion:
            usuario_db = conexion.execute(
                select(usuarios).where(
                    func.lower(usuarios.c.usuario) == usuario.lower()
                )
            ).mappings().first()

        if (
            usuario_db
            and usuario_db["activo"]
            and check_password_hash(usuario_db["password_hash"], password)
        ):
            session["admin"] = True
            session["usuario"] = usuario_db["usuario"]
            session["usuario_id"] = usuario_db["id"]
            session["es_superadmin"] = usuario_db["es_superadmin"]
            session["puede_agregar"] = usuario_db["puede_agregar"]
            session["puede_editar"] = usuario_db["puede_editar"]
            session["puede_eliminar"] = usuario_db["puede_eliminar"]
            session["puede_eliminar_ubicacion"] = usuario_db["puede_eliminar_ubicacion"]
            session["puede_actualizar_excel"] = usuario_db["puede_actualizar_excel"]

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

    buscar = request.args.get("buscar", "").strip()
    resultados = []

    if buscar:
        resultados = buscar_productos(buscar)

    return render_template(
        "admin.html",
        buscar=buscar,
        resultados=resultados,
    ) 

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
    if not session.get("es_superadmin") and not session.get("puede_actualizar_excel"):
        return redirect(url_for("admin"))
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

                            if sucursal.strip().lower() != "alajuela" or not codigo or not ubicacion:
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

                if sucursal.strip().lower() != "alajuela" or not producto or not codigo or not ubicacion:
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
    if not session.get("es_superadmin") and not session.get("puede_agregar"):
        return redirect(url_for("admin")) 
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
    if not session.get("es_superadmin") and not session.get("puede_editar"):
        return redirect(url_for("admin"))
    # Entrada directa desde el buscador de Administración.
    if request.method == "GET":
        codigo_buscar = request.args.get("codigo", "").strip()

        if not codigo_buscar:
            return render_template("editar.html")

        with engine.connect() as conexion:
            coincidencias = conexion.execute(
                select(
                    productos.c.id,
                    productos.c.producto,
                    productos.c.codigo,
                    productos.c.codigo_barras,
                    productos.c.ubicacion,
                    productos.c.sucursal,
                ).where(
                    func.lower(productos.c.codigo) == codigo_buscar.lower(),
                    func.lower(productos.c.sucursal) == "alajuela",
                )
            ).mappings().all()

        if not coincidencias:
            return render_template(
                "editar.html",
                error="No existe ese código en la sucursal Alajuela.",
                codigo_buscar=codigo_buscar,
            )

        return render_template(
            "editar.html",
            coincidencias=coincidencias,
            codigo_buscar=codigo_buscar,
        )

    # Desde aquí continúa únicamente el POST.
    codigo_buscar = request.form.get("codigo_buscar", "").strip()
    accion = request.form.get("accion", "buscar").strip()

    producto_masivo = request.form.get("producto_masivo", "").strip()
    codigo_masivo = request.form.get("codigo_masivo", "").strip()
    codigo_barras_masivo = request.form.get(
        "codigo_barras_masivo", ""
    ).strip()
    sucursal_masiva = request.form.get("sucursal_masiva", "").strip()

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
                productos.c.codigo_barras,
                productos.c.ubicacion,
                productos.c.sucursal,
            ).where(
                func.lower(productos.c.codigo) == codigo_buscar.lower(),
                func.lower(productos.c.sucursal) == "alajuela",
            )
        ).mappings().all()

        if not coincidencias:
            return render_template(
                "editar.html",
                error="No existe ese código en la sucursal Alajuela.",
                codigo_buscar=codigo_buscar,
            )

        # Primera etapa: mostramos todas las ubicaciones encontradas.
        if accion != "guardar":
            return render_template(
                "editar.html",
                coincidencias=coincidencias,
                codigo_buscar=codigo_buscar,
            )

        # Segunda etapa: recibimos las filas marcadas para modificar.
        ids_seleccionados = request.form.getlist("ids_seleccionados")

        if not ids_seleccionados:
            return render_template(
                "editar.html",
                error="Debe seleccionar al menos una ubicación para editar.",
                coincidencias=coincidencias,
                codigo_buscar=codigo_buscar,
            )

        try:
            ids_seleccionados = [
                int(id_registro)
                for id_registro in ids_seleccionados
            ]
        except ValueError:
            return render_template(
                "editar.html",
                error="La selección contiene un registro no válido.",
                coincidencias=coincidencias,
                codigo_buscar=codigo_buscar,
            )

        ids_validos = {fila["id"] for fila in coincidencias}

        if not set(ids_seleccionados).issubset(ids_validos):
            return render_template(
                "editar.html",
                error=(
                    "Una de las filas seleccionadas "
                    "no corresponde a ese código."
                ),
                coincidencias=coincidencias,
                codigo_buscar=codigo_buscar,
            )

        cambios_masivos = {}

        if producto_masivo:
            cambios_masivos["producto"] = producto_masivo

        if codigo_masivo:
            cambios_masivos["codigo"] = codigo_masivo

        if codigo_barras_masivo:
            cambios_masivos["codigo_barras"] = codigo_barras_masivo

        if sucursal_masiva:
            cambios_masivos["sucursal"] = sucursal_masiva

        for fila in coincidencias:
            id_registro = fila["id"]

            if id_registro not in ids_seleccionados:
                continue

            producto_nuevo = request.form.get(
                f"producto_{id_registro}", ""
            ).strip()

            codigo_nuevo = request.form.get(
                f"codigo_{id_registro}", ""
            ).strip()

            codigo_barras_nuevo = request.form.get(
                f"codigo_barras_{id_registro}", ""
            ).strip()

            ubicacion_nueva = request.form.get(
                f"ubicacion_{id_registro}", ""
            ).strip()

            sucursal_nueva = request.form.get(
                f"sucursal_{id_registro}", ""
            ).strip()

            producto_nuevo = cambios_masivos.get(
                "producto",
                producto_nuevo,
            )

            codigo_nuevo = cambios_masivos.get(
                "codigo",
                codigo_nuevo,
            )

            codigo_barras_nuevo = cambios_masivos.get(
                "codigo_barras",
                codigo_barras_nuevo,
            )

            sucursal_nueva = cambios_masivos.get(
                "sucursal",
                sucursal_nueva,
            )

            if (
                not producto_nuevo
                or not codigo_nuevo
                or not ubicacion_nueva
                or not sucursal_nueva
            ):
                return render_template(
                    "editar.html",
                    error=(
                        "Producto, código, ubicación y sucursal "
                        "no pueden quedar vacíos."
                    ),
                    coincidencias=coincidencias,
                    codigo_buscar=codigo_buscar,
                )

            conexion.execute(
                update(productos)
                .where(productos.c.id == id_registro)
                .values(
                    producto=producto_nuevo,
                    codigo=codigo_nuevo,
                    codigo_barras=codigo_barras_nuevo or None,
                    ubicacion=ubicacion_nueva,
                    sucursal=sucursal_nueva,
                )
            )

    return redirect(url_for("admin")) 



@app.route("/eliminar", methods=["GET", "POST"])
def eliminar():
    if not session.get("admin"):
        return redirect(url_for("login"))
    if not session.get("es_superadmin") and not session.get("puede_eliminar"):
        return redirect(url_for("admin"))
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

    codigo = request.args.get("codigo", "").strip()
    id_registro = request.args.get("id", "").strip()

    if not codigo:
        return render_template("eliminar.html")

    with engine.connect() as conexion:
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

    return render_template(
        "eliminar.html",
        coincidencias=coincidencias,
        codigo=codigo,
        id_registro=id_registro,
    ) 

@app.route("/eliminar_ubicacion", methods=["GET", "POST"])
def eliminar_ubicacion():
    if not session.get("admin"):
        return redirect(url_for("login"))
    if not session.get("es_superadmin") and not session.get("puede_eliminar_ubicacion"):
        return redirect(url_for("admin"))
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



@app.route("/usuarios", methods=["GET", "POST"])
def gestionar_usuarios():
    """Lista y crea usuarios con permisos. Solo disponible para el administrador principal."""
    if not session.get("admin"):
        return redirect(url_for("login"))
    if not session.get("es_superadmin"):
     return redirect(url_for("admin"))
    mensaje = ""
    error = ""

    if request.method == "POST":
        usuario = request.form.get("usuario", "").strip()
        nombre = request.form.get("nombre", "").strip()
        password = request.form.get("password", "").strip()

        if not usuario or not nombre or not password:
            error = "Usuario, nombre y contraseña son obligatorios."
        else:
            with engine.begin() as conexion:
                existente = conexion.execute(
                    select(usuarios.c.id).where(
                        func.lower(usuarios.c.usuario) == usuario.lower()
                    )
                ).scalar_one_or_none()

                if existente is not None:
                    error = "Ese nombre de usuario ya existe."
                else:
                    conexion.execute(
                        insert(usuarios).values(
                            usuario=usuario,
                            nombre=nombre,
                            password_hash=generate_password_hash(password),
                            activo=True,
                            es_superadmin=False,
                            puede_agregar=request.form.get("puede_agregar") == "on",
                            puede_editar=request.form.get("puede_editar") == "on",
                            puede_eliminar=request.form.get("puede_eliminar") == "on",
                            puede_eliminar_ubicacion=(
                                request.form.get("puede_eliminar_ubicacion") == "on"
                            ),
                            puede_actualizar_excel=(
                                request.form.get("puede_actualizar_excel") == "on"
                            ),
                        )
                    )
                    mensaje = "Usuario creado correctamente."

    with engine.connect() as conexion:
        lista_usuarios = conexion.execute(
            select(
                usuarios.c.id,
                usuarios.c.usuario,
                usuarios.c.nombre,
                usuarios.c.activo,
                usuarios.c.es_superadmin,
                usuarios.c.puede_agregar,
                usuarios.c.puede_editar,
                usuarios.c.puede_eliminar,
                usuarios.c.puede_eliminar_ubicacion,
                usuarios.c.puede_actualizar_excel,
            ).order_by(usuarios.c.usuario)
        ).mappings().all()

    return render_template(
        "usuarios.html",
        mensaje=mensaje,
        error=error,
        usuarios_lista=lista_usuarios,
    ) 
         


@app.route("/usuarios/<int:usuario_id>/editar", methods=["GET", "POST"])
def editar_usuario(usuario_id):
    """Consulta o actualiza permisos de un usuario. No permite desactivar al superadmin."""
    if not session.get("admin"):
        return redirect(url_for("login"))
    if not session.get("es_superadmin"):
        return redirect(url_for("admin"))
    with engine.begin() as conexion:
        fila = conexion.execute(
            select(usuarios).where(usuarios.c.id == usuario_id)
        ).mappings().first()

        if fila is None:
            return jsonify({"error": "Usuario no encontrado."}), 404

        if request.method == "POST":
            nombre = request.form.get("nombre", "").strip() or fila["nombre"]
            password = request.form.get("password", "").strip()
            es_superadmin = bool(fila["es_superadmin"])

            valores = {
                "nombre": nombre,
                "activo": True if es_superadmin else request.form.get("activo") == "on",
                "puede_agregar": True if es_superadmin else request.form.get("puede_agregar") == "on",
                "puede_editar": True if es_superadmin else request.form.get("puede_editar") == "on",
                "puede_eliminar": True if es_superadmin else request.form.get("puede_eliminar") == "on",
                "puede_eliminar_ubicacion": (
                    True if es_superadmin
                    else request.form.get("puede_eliminar_ubicacion") == "on"
                ),
                "puede_actualizar_excel": (
                    True if es_superadmin
                    else request.form.get("puede_actualizar_excel") == "on"
                ),
            }

            if password:
                valores["password_hash"] = generate_password_hash(password)

            conexion.execute(
                update(usuarios)
                .where(usuarios.c.id == usuario_id)
                .values(**valores)
            )

            fila = conexion.execute(
                select(usuarios).where(usuarios.c.id == usuario_id)
            ).mappings().first()

    usuario = dict(fila)
    usuario.pop("password_hash", None)

    return render_template(
    "editar_usuario.html",
    usuario=usuario,
) 


@app.route("/logout")
def logout():
    session.pop("admin", None)
    return redirect(url_for("inicio"))


if __name__ == "__main__":
    app.run(debug=True)
