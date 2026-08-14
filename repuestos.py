from pathlib import Path
import sqlite3
import unicodedata
from datetime import datetime

from flask import (
    Blueprint,
    render_template,
    request,
    jsonify,
    redirect,
    url_for,
    flash,
)


repuestos_bp = Blueprint(
    "repuestos",
    __name__,
    url_prefix="/repuestos",
)


BASE_DIR = Path(__file__).resolve().parent
CARPETA_DATOS = BASE_DIR / "datos"
RUTA_BD = CARPETA_DATOS / "repuestos.db"


def conectar_bd():
    CARPETA_DATOS.mkdir(
        parents=True,
        exist_ok=True,
    )

    conexion = sqlite3.connect(
        RUTA_BD,
        timeout=60,
    )

    conexion.row_factory = sqlite3.Row

    conexion.execute(
        "PRAGMA foreign_keys = ON"
    )

    conexion.execute(
        "PRAGMA busy_timeout = 60000"
    )

    return conexion


def cargar_openpyxl():
    try:
        from openpyxl import load_workbook

        return load_workbook

    except ModuleNotFoundError as error:
        raise RuntimeError(
            "Falta instalar openpyxl. "
            "Ejecute: python -m pip install openpyxl"
        ) from error


def normalizar_encabezado(
    valor,
) -> str:
    texto = str(
        valor or ""
    ).strip().upper()

    texto = "".join(
        caracter
        for caracter in unicodedata.normalize(
            "NFD",
            texto,
        )
        if unicodedata.category(
            caracter
        ) != "Mn"
    )

    texto = " ".join(
        texto.split()
    )

    return texto


def convertir_numero(
    valor,
) -> float:
    if valor in (
        None,
        "",
    ):
        return 0.0

    if isinstance(
        valor,
        (
            int,
            float,
        ),
    ):
        return float(
            valor
        )

    texto = str(
        valor
    ).strip()

    if not texto:
        return 0.0

    texto = (
        texto
        .replace("$", "")
        .replace("COP", "")
        .replace(" ", "")
    )

    try:
        return float(
            texto
        )

    except ValueError:
        pass

    try:
        if (
            "." in texto
            and "," in texto
        ):
            texto = (
                texto
                .replace(".", "")
                .replace(",", ".")
            )

        elif "," in texto:
            texto = texto.replace(
                ",",
                ".",
            )

        return float(
            texto
        )

    except ValueError:
        return 0.0


def normalizar_codigo_sap(
    valor,
) -> str:
    if valor in (
        None,
        "",
    ):
        return ""

    if isinstance(
        valor,
        float,
    ) and valor.is_integer():
        return str(
            int(
                valor
            )
        )

    return str(
        valor
    ).strip()


def inicializar_bd():
    conexion = conectar_bd()

    try:
        cursor = conexion.cursor()

        cursor.executescript(
            """
            CREATE TABLE IF NOT EXISTS materiales_sap (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo_sap TEXT UNIQUE,
                descripcion TEXT NOT NULL,
                unidad_medida TEXT,
                existencia REAL DEFAULT 0,
                valor_inventario REAL DEFAULT 0,
                valor_unitario REAL DEFAULT 0,
                origen_dato TEXT DEFAULT 'SAP',
                fecha_actualizacion TEXT,
                activo INTEGER DEFAULT 1
            );

            CREATE TABLE IF NOT EXISTS solicitudes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo_solicitud TEXT UNIQUE NOT NULL,
                semana INTEGER NOT NULL,
                anio INTEGER NOT NULL,
                fecha_creacion TEXT NOT NULL,
                area TEXT DEFAULT 'MP TAPA PLASTICA',
                estado TEXT DEFAULT 'BORRADOR',
                archivo_generado TEXT
            );

            CREATE TABLE IF NOT EXISTS solicitud_detalle (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                solicitud_id INTEGER NOT NULL,
                codigo_sap TEXT,
                descripcion TEXT NOT NULL,
                origen TEXT,
                unidad_medida TEXT,
                existencia_al_solicitar REAL DEFAULT 0,
                consumo_promedio REAL DEFAULT 0,
                tiempo_entrega REAL DEFAULT 0,
                cantidad_pedir REAL DEFAULT 0,
                valor_unitario REAL DEFAULT 0,
                observaciones TEXT,
                origen_dato TEXT DEFAULT 'SAP',
                estado TEXT DEFAULT 'PENDIENTE',
                cantidad_recibida REAL DEFAULT 0,
                cantidad_retirada REAL DEFAULT 0,
                fecha_recibido TEXT,
                fecha_retiro TEXT,
                FOREIGN KEY (
                    solicitud_id
                )
                REFERENCES solicitudes(id)
                ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS secciones (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT UNIQUE NOT NULL,
                descripcion TEXT,
                activo INTEGER DEFAULT 1,
                fecha_creacion TEXT
            );

            CREATE TABLE IF NOT EXISTS subsecciones (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                seccion_id INTEGER NOT NULL,
                nombre TEXT NOT NULL,
                descripcion TEXT,
                activo INTEGER DEFAULT 1,
                fecha_creacion TEXT,
                UNIQUE (
                    seccion_id,
                    nombre
                ),
                FOREIGN KEY (
                    seccion_id
                )
                REFERENCES secciones(id)
                ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS inventario_tecnico (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo_sap TEXT,
                descripcion TEXT NOT NULL,
                unidad_medida TEXT,
                seccion_id INTEGER NOT NULL,
                subseccion_id INTEGER,
                tipo_control TEXT DEFAULT 'manual',
                stock_minimo REAL DEFAULT 0,
                stock_objetivo REAL DEFAULT 0,
                factor_minimo REAL DEFAULT 1,
                factor_objetivo REAL DEFAULT 1,
                origen_dato TEXT DEFAULT 'SAP',
                observaciones TEXT,
                activo INTEGER DEFAULT 1,
                fecha_creacion TEXT,
                FOREIGN KEY (
                    seccion_id
                )
                REFERENCES secciones(id),
                FOREIGN KEY (
                    subseccion_id
                )
                REFERENCES subsecciones(id)
            );

            CREATE TABLE IF NOT EXISTS inventario_maquinas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                inventario_id INTEGER NOT NULL,
                maquina_id INTEGER NOT NULL,
                cantidad_necesaria REAL DEFAULT 0,
                FOREIGN KEY (
                    inventario_id
                )
                REFERENCES inventario_tecnico(id)
                ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS recepciones (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                detalle_solicitud_id INTEGER NOT NULL,
                cantidad REAL NOT NULL,
                fecha_recepcion TEXT NOT NULL,
                confirmado INTEGER DEFAULT 0,
                observaciones TEXT,
                FOREIGN KEY (
                    detalle_solicitud_id
                )
                REFERENCES solicitud_detalle(id)
                ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS retiros (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                detalle_solicitud_id INTEGER NOT NULL,
                cantidad REAL NOT NULL,
                fecha_retiro TEXT NOT NULL,
                observaciones TEXT,
                FOREIGN KEY (
                    detalle_solicitud_id
                )
                REFERENCES solicitud_detalle(id)
                ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS snapshots_sap (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha_carga TEXT NOT NULL,
                codigo_sap TEXT NOT NULL,
                descripcion TEXT,
                unidad_medida TEXT,
                existencia REAL DEFAULT 0,
                valor_inventario REAL DEFAULT 0,
                valor_unitario REAL DEFAULT 0
            );

            CREATE TABLE IF NOT EXISTS cargas_sap (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha_carga TEXT NOT NULL,
                nombre_archivo TEXT,
                cantidad_registros INTEGER DEFAULT 0
            );
            """
        )

        conexion.commit()

    finally:
        conexion.close()


def obtener_resumen():
    conexion = conectar_bd()

    try:
        solicitudes_activas = conexion.execute(
            """
            SELECT COUNT(*) AS total
            FROM solicitudes
            WHERE estado NOT IN (
                'CERRADA',
                'ANULADA'
            )
            """
        ).fetchone()["total"]

        pendientes_recibir = conexion.execute(
            """
            SELECT COUNT(*) AS total
            FROM solicitud_detalle
            WHERE estado IN (
                'PENDIENTE',
                'RECIBIDO_PARCIAL'
            )
            """
        ).fetchone()["total"]

        recibidos_retirar = conexion.execute(
            """
            SELECT COUNT(*) AS total
            FROM solicitud_detalle
            WHERE estado IN (
                'RECIBIDO',
                'PENDIENTE_RETIRO'
            )
            AND cantidad_recibida > cantidad_retirada
            """
        ).fetchone()["total"]

        bajo_stock = conexion.execute(
            """
            SELECT COUNT(*) AS total
            FROM inventario_tecnico inventario
            LEFT JOIN materiales_sap material
                ON material.codigo_sap =
                   inventario.codigo_sap
            WHERE inventario.activo = 1
            AND COALESCE(
                material.existencia,
                0
            ) < inventario.stock_minimo
            """
        ).fetchone()["total"]

        return {
            "solicitudes_activas":
                int(
                    solicitudes_activas
                    or 0
                ),

            "pendientes_recibir":
                int(
                    pendientes_recibir
                    or 0
                ),

            "recibidos_retirar":
                int(
                    recibidos_retirar
                    or 0
                ),

            "bajo_stock":
                int(
                    bajo_stock
                    or 0
                ),
        }

    finally:
        conexion.close()


def buscar_materiales(
    texto: str,
):
    texto = (
        texto
        .strip()
        .upper()
    )

    if not texto:
        return []

    conexion = conectar_bd()

    try:
        filas = conexion.execute(
            """
            SELECT
                id,
                codigo_sap,
                descripcion,
                unidad_medida,
                existencia,
                valor_inventario,
                valor_unitario,
                origen_dato

            FROM materiales_sap

            WHERE activo = 1

            AND (
                UPPER(
                    COALESCE(
                        codigo_sap,
                        ''
                    )
                ) LIKE ?

                OR

                UPPER(
                    descripcion
                ) LIKE ?
            )

            ORDER BY
                CASE
                    WHEN UPPER(
                        COALESCE(
                            codigo_sap,
                            ''
                        )
                    ) = ?
                    THEN 0
                    ELSE 1
                END,
                descripcion

            LIMIT 30
            """,
            (
                f"%{texto}%",
                f"%{texto}%",
                texto,
            ),
        ).fetchall()

        return [
            dict(
                fila
            )
            for fila in filas
        ]

    finally:
        conexion.close()


def buscar_pedido_pendiente(
    codigo_sap: str,
):
    if not codigo_sap:
        return None

    conexion = conectar_bd()

    try:
        fila = conexion.execute(
            """
            SELECT
                sd.id,
                sd.codigo_sap,
                sd.descripcion,
                sd.cantidad_pedir,
                sd.cantidad_recibida,
                sd.estado,
                s.codigo_solicitud,
                s.semana,
                s.anio

            FROM solicitud_detalle sd

            INNER JOIN solicitudes s
                ON s.id = sd.solicitud_id

            WHERE sd.codigo_sap = ?

            AND sd.estado IN (
                'PENDIENTE',
                'RECIBIDO_PARCIAL',
                'PENDIENTE_RETIRO'
            )

            AND s.estado NOT IN (
                'CERRADA',
                'ANULADA'
            )

            ORDER BY sd.id DESC

            LIMIT 1
            """,
            (
                codigo_sap,
            ),
        ).fetchone()

        if fila is None:
            return None

        resultado = dict(
            fila
        )

        resultado[
            "cantidad_pendiente"
        ] = max(
            0,
            float(
                resultado[
                    "cantidad_pedir"
                ]
                or 0
            )
            - float(
                resultado[
                    "cantidad_recibida"
                ]
                or 0
            ),
        )

        return resultado

    finally:
        conexion.close()


def encontrar_fila_encabezados(
    hoja,
) -> int:
    requeridos = {
        "MATERIAL",
        "TEXTO BREVE DE MATERIAL",
    }

    limite = min(
        hoja.max_row,
        30,
    )

    for fila in range(
        1,
        limite + 1,
    ):
        encontrados = set()

        for columna in range(
            1,
            hoja.max_column + 1,
        ):
            encabezado = normalizar_encabezado(
                hoja.cell(
                    fila,
                    columna,
                ).value
            )

            if encabezado:
                encontrados.add(
                    encabezado
                )

        if requeridos.issubset(
            encontrados
        ):
            return fila

    raise ValueError(
        "No se encontró una fila de encabezados válida. "
        "El archivo debe contener como mínimo las columnas "
        "'Material' y 'Texto breve de material'."
    )


def obtener_mapa_columnas(
    hoja,
    fila_encabezado: int,
) -> dict[str, int]:
    mapa = {}

    for columna in range(
        1,
        hoja.max_column + 1,
    ):
        encabezado = normalizar_encabezado(
            hoja.cell(
                fila_encabezado,
                columna,
            ).value
        )

        if encabezado:
            mapa[
                encabezado
            ] = columna

    return mapa


def buscar_columna(
    mapa: dict[str, int],
    opciones: list[str],
):
    for opcion in opciones:
        normalizada = normalizar_encabezado(
            opcion
        )

        if normalizada in mapa:
            return mapa[
                normalizada
            ]

    return None


def importar_base_sap(
    ruta_excel: Path,
) -> dict:
    load_workbook = cargar_openpyxl()

    workbook = load_workbook(
        ruta_excel,
        data_only=True,
        read_only=True,
    )

    try:
        hoja = workbook[
            workbook.sheetnames[0]
        ]

        fila_encabezado = encontrar_fila_encabezados(
            hoja
        )

        encabezados = obtener_mapa_columnas(
            hoja,
            fila_encabezado,
        )

        col_codigo = buscar_columna(
            encabezados,
            [
                "MATERIAL",
            ],
        )

        col_descripcion = buscar_columna(
            encabezados,
            [
                "TEXTO BREVE DE MATERIAL",
            ],
        )

        col_unidad = buscar_columna(
            encabezados,
            [
                "UNIDAD MEDIDA BASE",
                "UNIDAD DE MEDIDA BASE",
                "UM",
                "UNIDAD",
            ],
        )

        col_existencia = buscar_columna(
            encabezados,
            [
                "LIBRE UTILIZACION",
                "LIBRE UTIL.",
                "STOCK LIBRE UTILIZACION",
            ],
        )

        col_valor = buscar_columna(
            encabezados,
            [
                "VALOR LIBRE UTIL.",
                "VALOR LIBRE UTIL",
                "VALOR LIBRE UTILIZACION",
            ],
        )

        faltantes = []

        if not col_codigo:
            faltantes.append(
                "Material"
            )

        if not col_descripcion:
            faltantes.append(
                "Texto breve de material"
            )

        if not col_unidad:
            faltantes.append(
                "Unidad medida base"
            )

        if not col_existencia:
            faltantes.append(
                "Libre utilización"
            )

        if not col_valor:
            faltantes.append(
                "Valor libre util."
            )

        if faltantes:
            raise ValueError(
                "El archivo SAP no contiene las columnas requeridas: "
                + ", ".join(
                    faltantes
                )
            )

        fecha_carga = datetime.now().isoformat(
            timespec="seconds"
        )

        materiales = []
        snapshots = []

        cantidad = 0
        omitidos = 0

        indice_codigo = col_codigo - 1
        indice_descripcion = col_descripcion - 1
        indice_unidad = col_unidad - 1
        indice_existencia = col_existencia - 1
        indice_valor = col_valor - 1

        for fila in hoja.iter_rows(
            min_row=(
                fila_encabezado
                + 1
            ),
            values_only=True,
        ):
            codigo = normalizar_codigo_sap(
                fila[
                    indice_codigo
                ]
            )

            descripcion = str(
                fila[
                    indice_descripcion
                ]
                or ""
            ).strip()

            if not codigo:
                omitidos += 1
                continue

            if not descripcion:
                omitidos += 1
                continue

            unidad = str(
                fila[
                    indice_unidad
                ]
                or ""
            ).strip()

            existencia = convertir_numero(
                fila[
                    indice_existencia
                ]
            )

            valor_inventario = convertir_numero(
                fila[
                    indice_valor
                ]
            )

            valor_unitario = 0.0

            if existencia > 0:
                valor_unitario = (
                    valor_inventario
                    / existencia
                )

            materiales.append(
                (
                    codigo,
                    descripcion,
                    unidad,
                    existencia,
                    valor_inventario,
                    valor_unitario,
                    fecha_carga,
                )
            )

            snapshots.append(
                (
                    fecha_carga,
                    codigo,
                    descripcion,
                    unidad,
                    existencia,
                    valor_inventario,
                    valor_unitario,
                )
            )

            cantidad += 1

        conexion = conectar_bd()

        try:
            conexion.execute(
                """
                UPDATE materiales_sap
                SET activo = 0
                """
            )

            conexion.executemany(
                """
                INSERT INTO materiales_sap (
                    codigo_sap,
                    descripcion,
                    unidad_medida,
                    existencia,
                    valor_inventario,
                    valor_unitario,
                    origen_dato,
                    fecha_actualizacion,
                    activo
                )

                VALUES (
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    'SAP',
                    ?,
                    1
                )

                ON CONFLICT(
                    codigo_sap
                )

                DO UPDATE SET
                    descripcion =
                        excluded.descripcion,

                    unidad_medida =
                        excluded.unidad_medida,

                    existencia =
                        excluded.existencia,

                    valor_inventario =
                        excluded.valor_inventario,

                    valor_unitario =
                        excluded.valor_unitario,

                    fecha_actualizacion =
                        excluded.fecha_actualizacion,

                    origen_dato =
                        'SAP',

                    activo =
                        1
                """,
                materiales,
            )

            conexion.executemany(
                """
                INSERT INTO snapshots_sap (
                    fecha_carga,
                    codigo_sap,
                    descripcion,
                    unidad_medida,
                    existencia,
                    valor_inventario,
                    valor_unitario
                )

                VALUES (
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?
                )
                """,
                snapshots,
            )

            conexion.execute(
                """
                INSERT INTO cargas_sap (
                    fecha_carga,
                    nombre_archivo,
                    cantidad_registros
                )

                VALUES (
                    ?,
                    ?,
                    ?
                )
                """,
                (
                    fecha_carga,
                    ruta_excel.name,
                    cantidad,
                ),
            )

            conexion.commit()

            return {
                "cantidad":
                    cantidad,

                "omitidos":
                    omitidos,

                "fecha_carga":
                    fecha_carga,

                "archivo":
                    ruta_excel.name,

                "fila_encabezado":
                    fila_encabezado,
            }

        except Exception:
            conexion.rollback()
            raise

        finally:
            conexion.close()

    finally:
        workbook.close()


@repuestos_bp.route("/")
def inicio():
    inicializar_bd()

    return render_template(
        "repuestos/inicio.html",
        resumen=obtener_resumen(),
    )


@repuestos_bp.route("/nueva")
def nueva_solicitud():
    inicializar_bd()

    return render_template(
        "repuestos/nueva_solicitud.html"
    )


@repuestos_bp.route(
    "/actualizar-sap",
    methods=[
        "GET",
        "POST",
    ],
)
def actualizar_sap():
    inicializar_bd()

    if request.method == "POST":
        archivo = request.files.get(
            "archivo_sap"
        )

        if (
            archivo is None
            or not archivo.filename
        ):
            flash(
                "Seleccione un archivo Excel.",
                "error",
            )

            return redirect(
                url_for(
                    "repuestos.actualizar_sap"
                )
            )

        extension = Path(
            archivo.filename
        ).suffix.lower()

        if extension not in (
            ".xlsx",
            ".xlsm",
        ):
            flash(
                "Solo se permiten archivos .xlsx o .xlsm.",
                "error",
            )

            return redirect(
                url_for(
                    "repuestos.actualizar_sap"
                )
            )

        carpeta = (
            CARPETA_DATOS
            / "cargas_sap"
        )

        carpeta.mkdir(
            parents=True,
            exist_ok=True,
        )

        nombre_original = Path(
            archivo.filename
        ).name

        nombre = (
            datetime.now().strftime(
                "%Y%m%d_%H%M%S"
            )
            + "_"
            + nombre_original
        )

        ruta = (
            carpeta
            / nombre
        )

        archivo.save(
            ruta
        )

        try:
            resultado = importar_base_sap(
                ruta
            )

            mensaje = (
                f"Base SAP actualizada correctamente. "
                f"{resultado['cantidad']} materiales cargados."
            )

            if resultado[
                "omitidos"
            ]:
                mensaje += (
                    f" {resultado['omitidos']} filas fueron omitidas "
                    f"por no tener código o descripción."
                )

            flash(
                mensaje,
                "exito",
            )

        except Exception as error:
            flash(
                f"No fue posible actualizar la base SAP: {error}",
                "error",
            )

        return redirect(
            url_for(
                "repuestos.actualizar_sap"
            )
        )

    conexion = conectar_bd()

    try:
        ultima_carga = conexion.execute(
            """
            SELECT
                fecha_carga,
                nombre_archivo,
                cantidad_registros

            FROM cargas_sap

            ORDER BY id DESC

            LIMIT 1
            """
        ).fetchone()

        total_materiales = conexion.execute(
            """
            SELECT COUNT(*) AS total
            FROM materiales_sap
            WHERE activo = 1
            """
        ).fetchone()["total"]

    finally:
        conexion.close()

    return render_template(
        "repuestos/actualizar_sap.html",
        ultima_carga=ultima_carga,
        total_materiales=int(
            total_materiales
            or 0
        ),
    )


@repuestos_bp.route(
    "/api/materiales"
)
def api_materiales():
    texto = request.args.get(
        "q",
        "",
    )

    materiales = buscar_materiales(
        texto
    )

    return jsonify(
        materiales
    )


@repuestos_bp.route(
    "/api/material/<codigo_sap>/pendiente"
)
def api_material_pendiente(
    codigo_sap,
):
    pendiente = buscar_pedido_pendiente(
        codigo_sap
    )

    return jsonify(
        {
            "pendiente":
                pendiente
        }
    )
