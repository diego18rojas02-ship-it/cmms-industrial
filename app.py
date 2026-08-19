import json
import os
from repuestos import repuestos_bp
from calendar import monthrange
from datetime import date, datetime
from typing import Any

from flask import Flask, flash, redirect, render_template, request, url_for
from programacion import programacion_bp
from tiempos_perdidos import tiempos_perdidos_bp

try:
    import psycopg
    from psycopg.types.json import Jsonb
except ModuleNotFoundError:
    psycopg = None
    Jsonb = None


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_DATOS = os.path.join(BASE_DIR, "datos_cmms.json")

app = Flask(
    __name__,
    static_folder=os.path.join(BASE_DIR, "static"),
    template_folder=os.path.join(BASE_DIR, "templates"),
)

app.secret_key = os.environ.get("SECRET_KEY", "CMMS2026")
app.register_blueprint(programacion_bp)
app.register_blueprint(tiempos_perdidos_bp)
app.register_blueprint(repuestos_bp)

DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()


def datos_vacios() -> dict[str, Any]:
    return {
        "maquinas": [],
        "registros_horometros": [],
        "mantenimientos": {},
        "programaciones": [],
    }


def normalizar_datos(
    datos: dict[str, Any] | None,
) -> dict[str, Any]:
    if not isinstance(datos, dict):
        datos = datos_vacios()

    datos.setdefault("maquinas", [])
    datos.setdefault("registros_horometros", [])
    datos.setdefault("mantenimientos", {})
    datos.setdefault("programaciones", [])

    return datos


def cargar_datos_json_local() -> dict[str, Any]:
    if not os.path.exists(ARCHIVO_DATOS):
        return datos_vacios()

    try:
        with open(
            ARCHIVO_DATOS,
            "r",
            encoding="utf-8",
        ) as archivo:
            datos = json.load(archivo)
    except (
        OSError,
        json.JSONDecodeError,
    ):
        return datos_vacios()

    return normalizar_datos(datos)


def guardar_datos_json_local(
    datos: dict[str, Any],
) -> None:
    temporal = f"{ARCHIVO_DATOS}.tmp"

    with open(
        temporal,
        "w",
        encoding="utf-8",
    ) as archivo:
        json.dump(
            normalizar_datos(datos),
            archivo,
            ensure_ascii=False,
            indent=2,
        )

    os.replace(
        temporal,
        ARCHIVO_DATOS,
    )


def conectar_postgres():
    if not DATABASE_URL:
        raise RuntimeError(
            "DATABASE_URL no está configurada."
        )

    if psycopg is None:
        raise RuntimeError(
            "Falta instalar psycopg. "
            "Agregue psycopg[binary] a requirements.txt."
        )

    return psycopg.connect(
        DATABASE_URL,
        connect_timeout=15,
    )


def inicializar_postgres() -> None:
    if not DATABASE_URL:
        return

    with conectar_postgres() as conexion:
        with conexion.cursor() as cursor:
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS cmms_estado (
                    id SMALLINT PRIMARY KEY,
                    datos JSONB NOT NULL,
                    actualizado_en TIMESTAMPTZ NOT NULL DEFAULT NOW()
                )
                """
            )

            cursor.execute(
                """
                SELECT datos
                FROM cmms_estado
                WHERE id = 1
                """
            )

            fila = cursor.fetchone()

            if fila is None:
                datos_iniciales = cargar_datos_json_local()

                cursor.execute(
                    """
                    INSERT INTO cmms_estado (
                        id,
                        datos,
                        actualizado_en
                    )
                    VALUES (
                        1,
                        %s,
                        NOW()
                    )
                    """,
                    (
                        Jsonb(
                            normalizar_datos(
                                datos_iniciales
                            )
                        ),
                    ),
                )

        conexion.commit()


def cargar_datos_postgres() -> dict[str, Any]:
    inicializar_postgres()

    with conectar_postgres() as conexion:
        with conexion.cursor() as cursor:
            cursor.execute(
                """
                SELECT datos
                FROM cmms_estado
                WHERE id = 1
                """
            )

            fila = cursor.fetchone()

            if fila is None:
                return datos_vacios()

            datos = fila[0]

            if isinstance(datos, str):
                datos = json.loads(datos)

            return normalizar_datos(datos)


def guardar_datos_postgres(
    datos: dict[str, Any],
) -> None:
    inicializar_postgres()

    datos = normalizar_datos(datos)

    with conectar_postgres() as conexion:
        with conexion.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO cmms_estado (
                    id,
                    datos,
                    actualizado_en
                )
                VALUES (
                    1,
                    %s,
                    NOW()
                )
                ON CONFLICT (id)
                DO UPDATE SET
                    datos = EXCLUDED.datos,
                    actualizado_en = NOW()
                """,
                (
                    Jsonb(datos),
                ),
            )

        conexion.commit()


def cargar_datos() -> dict[str, Any]:
    if DATABASE_URL:
        try:
            return cargar_datos_postgres()
        except Exception as error:
            print(
                f"Error cargando PostgreSQL: {error}"
            )

    return cargar_datos_json_local()


def guardar_datos(
    datos: dict[str, Any],
) -> None:
    if DATABASE_URL:
        guardar_datos_postgres(
            datos
        )
        return

    guardar_datos_json_local(
        datos
    )

def buscar_maquina(
    datos: dict[str, Any],
    maquina_id: int,
) -> dict[str, Any] | None:
    for maquina in datos.get("maquinas", []):
        if int(maquina.get("id", 0)) == int(maquina_id):
            return maquina
    return None


def siguiente_maquina(
    datos: dict[str, Any],
    maquina_id: int,
) -> dict[str, Any] | None:
    maquinas = datos.get("maquinas", [])

    for indice, maquina in enumerate(maquinas):
        if int(maquina.get("id", 0)) == int(maquina_id):
            if indice + 1 < len(maquinas):
                return maquinas[indice + 1]
            return None

    return None


def obtener_horometro_actual(maquina: dict[str, Any]) -> int:
    valor = maquina.get(
        "horometro",
        maquina.get("horometro_actual", 0),
    )

    try:
        return int(valor)
    except (TypeError, ValueError):
        return 0


def obtener_controles_maquina(
    maquina: dict[str, Any],
) -> list[str]:
    controles: list[str] = []

    for control in maquina.get("controles", []):
        if isinstance(control, str):
            controles.append(control)
        elif isinstance(control, dict):
            codigo = control.get("codigo")
            if codigo:
                controles.append(str(codigo))

    return controles


def maquina_registrada_hoy(
    datos: dict[str, Any],
    maquina_id: int,
) -> bool:
    hoy = date.today().isoformat()

    for registro in datos.get("registros_horometros", []):
        if (
            int(registro.get("maquina_id", 0)) == int(maquina_id)
            and registro.get("fecha_registro") == hoy
            and registro.get("modo_registro", "horometro") == "horometro"
        ):
            return True

    return False


def obtener_ultimo_horometro_intervencion(
    datos: dict[str, Any],
    maquina_id: int,
) -> int:
    ultimo = 0

    for registro in datos.get("registros_horometros", []):
        if int(registro.get("maquina_id", 0)) != int(maquina_id):
            continue

        valor = registro.get("horometro_intervencion")

        if valor in (None, ""):
            continue

        try:
            ultimo = max(ultimo, int(valor))
        except (TypeError, ValueError):
            continue

    return ultimo


def prioridad_estado(estado: str) -> int:
    return {
        "sin_datos": 0,
        "en_tiempo": 1,
        "proximo": 2,
        "vencido": 3,
        "critico": 4,
    }.get(estado, 0)


def texto_estado(estado: str) -> str:
    return {
        "sin_datos": "Sin datos",
        "en_tiempo": "En tiempo",
        "proximo": "Próximo",
        "vencido": "Vencido",
        "critico": "Crítico",
    }.get(estado, "Sin datos")


def icono_estado(estado: str) -> str:
    return {
        "sin_datos": "⚪",
        "en_tiempo": "🟢",
        "proximo": "🟡",
        "vencido": "🟠",
        "critico": "🔴",
    }.get(estado, "⚪")


def estado_general_maquina(
    maquina: dict[str, Any],
) -> str:
    controles = maquina.get("controles", [])

    if not controles:
        return "sin_datos"

    estados = []

    for control in controles:
        if isinstance(control, dict):
            estados.append(control.get("estado", "sin_datos"))
        else:
            estados.append("sin_datos")

    return max(estados, key=prioridad_estado)


def convertir_fecha(valor: Any) -> date | None:
    if not valor:
        return None

    if isinstance(valor, date):
        return valor

    texto = str(valor).strip()

    for formato in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(texto, formato).date()
        except ValueError:
            continue

    return None


def fecha_visible(valor: Any) -> str:
    fecha = convertir_fecha(valor)

    if not fecha:
        return "Sin fecha"

    return fecha.strftime("%d/%m/%Y")


def dias_transcurridos(valor: Any) -> int | None:
    fecha = convertir_fecha(valor)

    if not fecha:
        return None

    return max(0, (date.today() - fecha).days)


def sumar_meses(fecha: date, meses: int) -> date:
    total = fecha.month - 1 + meses
    anio = fecha.year + total // 12
    mes = total % 12 + 1
    dia = min(fecha.day, monthrange(anio, mes)[1])

    return date(anio, mes, dia)


def meses_dias_restantes(
    valor_fecha: Any,
    meses_objetivo: int,
) -> tuple[int | None, int | None, bool]:
    fecha = convertir_fecha(valor_fecha)

    if not fecha:
        return None, None, False

    limite = sumar_meses(fecha, meses_objetivo)
    hoy = date.today()

    if hoy > limite:
        vencido_dias = (hoy - limite).days
        return 0, vencido_dias, True

    meses = 0
    cursor = hoy

    while True:
        siguiente = sumar_meses(cursor, 1)

        if siguiente <= limite:
            meses += 1
            cursor = siguiente
        else:
            break

    dias = (limite - cursor).days
    return meses, dias, False


def texto_meses_restantes(
    valor_fecha: Any,
    meses_objetivo: int,
) -> tuple[str, str]:
    meses, dias, vencido = meses_dias_restantes(
        valor_fecha,
        meses_objetivo,
    )

    if meses is None:
        return "sin_datos", "🗓 Sin información"

    if vencido:
        return "vencido", f"🗓 Vencido {dias} días"

    if meses == 0 and dias <= 30:
        return "proximo", f"🗓 {dias} días"

    if meses == 1:
        texto = "1 mes"
    else:
        texto = f"{meses} meses"

    if dias:
        texto += f" {dias} días"

    return "en_tiempo", f"🗓 {texto}"


def horas_y_dias_trabajo(
    horometro_actual: int,
    horometro_intervencion: int,
) -> tuple[int, float]:
    horas = max(
        0,
        int(horometro_actual) - int(horometro_intervencion),
    )

    dias = round(horas / 24, 1)
    return horas, dias


def estimar_meses_por_ritmo(
    horas_faltantes: int,
    horas_trabajadas: int,
    fecha_intervencion: Any,
) -> float | None:
    dias = dias_transcurridos(fecha_intervencion)

    if not dias or dias <= 0 or horas_trabajadas <= 0:
        return None

    horas_por_dia = horas_trabajadas / dias

    if horas_por_dia <= 0:
        return None

    dias_estimados = horas_faltantes / horas_por_dia
    return round(dias_estimados / 30.44, 1)


def formato_horas(valor: int) -> str:
    return f"{int(valor):,} hrs".replace(",", ".")


def preparar_maquinas(
    datos: dict[str, Any],
) -> list[dict[str, Any]]:
    maquinas = datos.get("maquinas", [])
    mantenimientos = datos.get("mantenimientos", {})

    configuracion = {
        "rutina": {
            "nombre": "Rutina",
            "icono": "🛠",
        },
        "punzones": {
            "nombre": "Punzones",
            "icono": "🔨",
        },
        "distribuidores": {
            "nombre": "Distribuidores",
            "icono": "💧",
        },
        "bridas": {
            "nombre": "Bridas y actuadores",
            "icono": "⚙",
        },
        "aceite": {
            "nombre": "Aceite",
            "icono": "🛢",
        },
    }

    resultado = []

    for maquina_original in maquinas:
        maquina = dict(maquina_original)
        maquina_id = str(maquina.get("id", 0))
        horometro_actual = obtener_horometro_actual(maquina_original)
        registros_maquina = mantenimientos.get(maquina_id, {})
        controles = []

        for codigo in obtener_controles_maquina(maquina_original):
            config = configuracion.get(
                codigo,
                {
                    "nombre": codigo.capitalize(),
                    "icono": "⚙",
                },
            )

            control = {
                "codigo": codigo,
                "nombre": config["nombre"],
                "icono": config["icono"],
                "estado": "sin_datos",
                "estado_texto": "Sin datos",
                "estado_icono": "⚪",
                "datos": [],
            }

            if codigo == "rutina":
                registro = registros_maquina.get("rutina", {})

                if registro:
                    horometro_intervencion = int(
                        registro.get("horometro_intervencion", 0)
                    )

                    fecha = registro.get("fecha_intervencion", "")
                    frecuencia = str(
                        registro.get("frecuencia", "")
                    ).lower()

                    dias_fecha = dias_transcurridos(fecha)
                    horas_trabajadas, dias_trabajo = horas_y_dias_trabajo(
                        horometro_actual,
                        horometro_intervencion,
                    )

                    limite = 30 if frecuencia == "mensual" else 15

                    if dias_fecha is None:
                        estado = "sin_datos"
                    elif dias_fecha > limite:
                        estado = "vencido"
                    elif limite - dias_fecha <= 3:
                        estado = "proximo"
                    else:
                        estado = "en_tiempo"

                    control["estado"] = estado
                    control["estado_texto"] = texto_estado(estado)
                    control["estado_icono"] = icono_estado(estado)

                    control["datos"] = [
                        f"⚙ {formato_horas(horometro_intervencion)}",
                        f"📈 {formato_horas(horometro_actual)}",
                        f"📅 {fecha_visible(fecha)}",
                        {
                            "tipo": "doble",
                            "izquierda": (
                                f"⏱ {dias_fecha} días"
                                if dias_fecha is not None
                                else "⏱ Sin fecha"
                            ),
                            "derecha": f"⚡ {dias_trabajo:.1f} días",
                        },
                        f"🔄 {frecuencia.capitalize() if frecuencia else 'Sin frecuencia'}",
                    ]

            elif codigo == "punzones":
                registro = registros_maquina.get("punzones", {})

                if registro:
                    horometro_intervencion = int(
                        registro.get("horometro_intervencion", 0)
                    )

                    fecha = registro.get("fecha_intervencion", "")

                    horas_trabajadas = max(
                        0,
                        horometro_actual - horometro_intervencion,
                    )

                    horas_faltantes = max(
                        0,
                        1800 - horas_trabajadas,
                    )

                    if horas_trabajadas >= 1800:
                        estado = "vencido"
                    elif horas_faltantes <= 200:
                        estado = "proximo"
                    else:
                        estado = "en_tiempo"

                    meses_estimados = estimar_meses_por_ritmo(
                        horas_faltantes,
                        horas_trabajadas,
                        fecha,
                    )

                    control["estado"] = estado
                    control["estado_texto"] = texto_estado(estado)
                    control["estado_icono"] = icono_estado(estado)

                    control["datos"] = [
                        f"⚙ {formato_horas(horometro_intervencion)}",
                        f"📈 {formato_horas(horometro_actual)}",
                        {
                            "tipo": "doble",
                            "izquierda": f"⚡ {formato_horas(horas_trabajadas)}",
                            "derecha": f"🎯 {formato_horas(horas_faltantes)}",
                        },
                        (
                            f"🗓 ≈ {meses_estimados:.1f} meses"
                            if meses_estimados is not None
                            else "🗓 Sin estimación"
                        ),
                    ]

            elif codigo == "distribuidores":
                registros = registros_maquina.get(
                    "distribuidores",
                    {},
                )

                grupos = []
                estados = []

                for nombre, clave in (
                    ("Agua", "distribuidor_agua"),
                    ("Aceite", "distribuidor_aceite"),
                ):
                    registro = registros.get(clave, {})

                    if registro:
                        fecha = registro.get(
                            "fecha_intervencion",
                            "",
                        )

                        estado, restante = texto_meses_restantes(
                            fecha,
                            6,
                        )

                        grupos.append(
                            {
                                "nombre": nombre,
                                "estado": estado,
                                "estado_texto": texto_estado(estado),
                                "estado_icono": icono_estado(estado),
                                "datos": [
                                    f"📅 {fecha_visible(fecha)}",
                                    restante,
                                ],
                            }
                        )

                        estados.append(estado)

                    else:
                        grupos.append(
                            {
                                "nombre": nombre,
                                "estado": "sin_datos",
                                "estado_texto": "Sin datos",
                                "estado_icono": "⚪",
                                "datos": [
                                    "📅 Sin información",
                                    "🗓 Sin información",
                                ],
                            }
                        )
                        estados.append("sin_datos")

                control["grupos"] = grupos
                control.pop("datos", None)

                estado = max(
                    estados,
                    key=prioridad_estado,
                )

                control["estado"] = estado
                control["estado_texto"] = texto_estado(estado)
                control["estado_icono"] = icono_estado(estado)

            elif codigo == "bridas":
                registro = registros_maquina.get("bridas", {})

                if registro:
                    fecha = registro.get("fecha_intervencion", "")
                    estado, restante = texto_meses_restantes(
                        fecha,
                        12,
                    )

                    control["estado"] = estado
                    control["estado_texto"] = texto_estado(estado)
                    control["estado_icono"] = icono_estado(estado)
                    control["datos"] = [
                        f"📅 {fecha_visible(fecha)}",
                        restante,
                    ]

            elif codigo == "aceite":
                registro = registros_maquina.get("aceite", {})

                if registro:
                    horometro_intervencion = int(
                        registro.get("horometro_intervencion", 0)
                    )

                    fecha = registro.get("fecha_intervencion", "")

                    horas_trabajadas = max(
                        0,
                        horometro_actual - horometro_intervencion,
                    )

                    horas_faltantes = max(
                        0,
                        40000 - horas_trabajadas,
                    )

                    if horas_trabajadas >= 40000:
                        estado = "vencido"
                    elif horas_faltantes <= 5000:
                        estado = "proximo"
                    else:
                        estado = "en_tiempo"

                    meses_estimados = estimar_meses_por_ritmo(
                        horas_faltantes,
                        horas_trabajadas,
                        fecha,
                    )

                    control["estado"] = estado
                    control["estado_texto"] = texto_estado(estado)
                    control["estado_icono"] = icono_estado(estado)

                    control["datos"] = [
                        f"⚙ {formato_horas(horometro_intervencion)}",
                        f"📈 {formato_horas(horometro_actual)}",
                        {
                            "tipo": "doble",
                            "izquierda": f"⚡ {formato_horas(horas_trabajadas)}",
                            "derecha": f"🎯 {formato_horas(horas_faltantes)}",
                        },
                        (
                            f"🗓 ≈ {meses_estimados:.1f} meses"
                            if meses_estimados is not None
                            else "🗓 Sin estimación"
                        ),
                    ]

            controles.append(control)

        maquina["controles"] = controles

        estado = estado_general_maquina(maquina)

        maquina["estado_general"] = estado
        maquina["estado_general_texto"] = texto_estado(estado)

        resultado.append(maquina)

    return resultado


def crear_resumen(
    maquinas: list[dict[str, Any]],
) -> dict[str, int]:
    resumen = {
        "total": len(maquinas),
        "en_tiempo": 0,
        "proximo": 0,
        "vencido": 0,
        "critico": 0,
        "sin_datos": 0,
    }

    for maquina in maquinas:
        estado = maquina.get("estado_general", "sin_datos")

        if estado in resumen:
            resumen[estado] += 1

    return resumen


def actualizar_control(
    datos: dict[str, Any],
    maquina_id: int,
    control: str,
    horometro_intervencion: int,
    fecha_intervencion: str,
    frecuencia: str | None = None,
    distribuidor: str | None = None,
) -> None:
    registro_maquina = datos.setdefault(
        "mantenimientos",
        {},
    ).setdefault(
        str(maquina_id),
        {},
    )

    if control == "rutina":
        registro = registro_maquina.setdefault(
            "rutina",
            {},
        )

        registro["horometro_intervencion"] = (
            horometro_intervencion
        )

        registro["fecha_intervencion"] = (
            fecha_intervencion
        )

        registro["frecuencia"] = (
            frecuencia
            or "quincenal"
        )

    elif control == "punzones":
        registro = registro_maquina.setdefault(
            "punzones",
            {},
        )

        registro["horometro_intervencion"] = (
            horometro_intervencion
        )

        registro["fecha_intervencion"] = (
            fecha_intervencion
        )

    elif control == "distribuidores":
        distribuidores = registro_maquina.setdefault(
            "distribuidores",
            {},
        )

        if distribuidor == "agua":
            codigos = [
                "distribuidor_agua"
            ]

        elif distribuidor == "aceite":
            codigos = [
                "distribuidor_aceite"
            ]

        elif distribuidor == "ambos":
            codigos = [
                "distribuidor_agua",
                "distribuidor_aceite",
            ]

        else:
            raise ValueError(
                "El distribuidor seleccionado no es válido."
            )

        for codigo in codigos:
            registro = distribuidores.setdefault(
                codigo,
                {},
            )

            registro["horometro_intervencion"] = (
                horometro_intervencion
            )

            registro["fecha_intervencion"] = (
                fecha_intervencion
            )

    elif control == "bridas":
        registro = registro_maquina.setdefault(
            "bridas",
            {},
        )

        registro["horometro_intervencion"] = (
            horometro_intervencion
        )

        registro["fecha_intervencion"] = (
            fecha_intervencion
        )

    elif control == "aceite":
        registro = registro_maquina.setdefault(
            "aceite",
            {},
        )

        registro["horometro_intervencion"] = (
            horometro_intervencion
        )

        registro["fecha_intervencion"] = (
            fecha_intervencion
        )


def registrar_mantenimiento(
    datos: dict[str, Any],
    maquina: dict[str, Any],
    maquina_id: int,
    horometro_intervencion: int,
    fecha_intervencion: str,
    frecuencia: str,
    cambio_aceite: bool,
    novedad_distribuidor: bool,
    distribuidor_novedad: str,
) -> None:
    controles = obtener_controles_maquina(maquina)

    if frecuencia in (
        "quincenal",
        "mensual",
    ):
        actualizar_control(
            datos=datos,
            maquina_id=maquina_id,
            control="rutina",
            horometro_intervencion=horometro_intervencion,
            fecha_intervencion=fecha_intervencion,
            frecuencia=frecuencia,
        )

    elif frecuencia == "trimestral":
        if "punzones" in controles:
            actualizar_control(
                datos=datos,
                maquina_id=maquina_id,
                control="punzones",
                horometro_intervencion=horometro_intervencion,
                fecha_intervencion=fecha_intervencion,
            )

        actualizar_control(
            datos=datos,
            maquina_id=maquina_id,
            control="rutina",
            horometro_intervencion=horometro_intervencion,
            fecha_intervencion=fecha_intervencion,
            frecuencia="quincenal",
        )

    elif frecuencia == "semestral":
        if "distribuidores" in controles:
            distribuidor_control = "ambos"

            if novedad_distribuidor:
                distribuidor_control = distribuidor_novedad

            actualizar_control(
                datos=datos,
                maquina_id=maquina_id,
                control="distribuidores",
                horometro_intervencion=horometro_intervencion,
                fecha_intervencion=fecha_intervencion,
                distribuidor=distribuidor_control,
            )

        actualizar_control(
            datos=datos,
            maquina_id=maquina_id,
            control="rutina",
            horometro_intervencion=horometro_intervencion,
            fecha_intervencion=fecha_intervencion,
            frecuencia="quincenal",
        )

    elif frecuencia == "anual":
        if "bridas" in controles:
            actualizar_control(
                datos=datos,
                maquina_id=maquina_id,
                control="bridas",
                horometro_intervencion=horometro_intervencion,
                fecha_intervencion=fecha_intervencion,
            )

        actualizar_control(
            datos=datos,
            maquina_id=maquina_id,
            control="rutina",
            horometro_intervencion=horometro_intervencion,
            fecha_intervencion=fecha_intervencion,
            frecuencia="quincenal",
        )

    if (
        cambio_aceite
        and "aceite" in controles
    ):
        actualizar_control(
            datos=datos,
            maquina_id=maquina_id,
            control="aceite",
            horometro_intervencion=horometro_intervencion,
            fecha_intervencion=fecha_intervencion,
        )


@app.route("/")
def inicio():
    datos = cargar_datos()
    maquinas = preparar_maquinas(datos)
    resumen = crear_resumen(maquinas)

    return render_template(
        "inicio.html",
        maquinas=maquinas,
        resumen=resumen,
    )


@app.route(
    "/horometros",
    methods=[
        "GET",
        "POST",
    ],
)
def horometros():
    datos = cargar_datos()
    maquinas = datos.get("maquinas", [])

    maquina_id = request.args.get(
        "maquina_id",
        type=int,
    )

    modo_get = request.args.get(
        "modo",
        "horometro",
    )

    if modo_get not in (
        "horometro",
        "mantenimiento",
    ):
        modo_get = "horometro"

    maquina_seleccionada = (
        buscar_maquina(
            datos,
            maquina_id,
        )
        if maquina_id is not None
        else None
    )

    mensaje = ""

    if request.method == "POST":
        modo = request.form.get(
            "modo_registro",
            "horometro",
        ).strip()

        if modo not in (
            "horometro",
            "mantenimiento",
        ):
            modo = "horometro"

        maquina_id = request.form.get(
            "maquina_id",
            type=int,
        )

        if maquina_id is None:
            flash(
                "No se recibió la máquina seleccionada.",
                "error",
            )
            return redirect(
                url_for(
                    "horometros",
                    modo=modo,
                )
            )

        maquina = buscar_maquina(
            datos,
            maquina_id,
        )

        if not maquina:
            flash(
                "La máquina seleccionada no existe.",
                "error",
            )
            return redirect(
                url_for(
                    "horometros",
                    modo=modo,
                )
            )

        horometro_anterior = obtener_horometro_actual(
            maquina
        )

        horometro_actual = horometro_anterior

        if modo == "horometro":
            texto_actual = request.form.get(
                "horometro_actual",
                "",
            ).strip()

            if not texto_actual.isdigit():
                flash(
                    "El horómetro actual debe contener solo números.",
                    "error",
                )
                return redirect(
                    url_for(
                        "horometros",
                        maquina_id=maquina_id,
                        modo=modo,
                    )
                )

            horometro_actual = int(
                texto_actual
            )

            if horometro_actual < horometro_anterior:
                flash(
                    "El horómetro actual debe ser igual o mayor que "
                    f"{horometro_anterior:,}.",
                    "error",
                )
                return redirect(
                    url_for(
                        "horometros",
                        maquina_id=maquina_id,
                        modo=modo,
                    )
                )

        if modo == "mantenimiento":
            ejecuto_mantenimiento = True
        else:
            ejecuto_mantenimiento = (
                request.form.get(
                    "ejecuto_mantenimiento"
                )
                == "si"
            )

        horometro_intervencion: int | None = None
        fecha_intervencion = ""
        frecuencia = ""
        cambio_aceite = False
        novedad_distribuidor = False
        distribuidor_novedad = ""

        if ejecuto_mantenimiento:
            texto_intervencion = (
                request.form.get(
                    "horometro_intervencion",
                    "",
                ).strip()
            )

            fecha_intervencion = (
                request.form.get(
                    "fecha_intervencion",
                    "",
                ).strip()
            )

            frecuencia = (
                request.form.get(
                    "frecuencia",
                    "",
                ).strip()
            )

            cambio_aceite = (
                request.form.get(
                    "cambio_aceite"
                )
                == "si"
            )

            novedad_distribuidor = (
                request.form.get(
                    "novedad_distribuidor"
                )
                == "si"
            )

            distribuidor_novedad = (
                request.form.get(
                    "distribuidor_novedad",
                    "",
                ).strip()
            )

            if not texto_intervencion.isdigit():
                flash(
                    "El horómetro de intervención debe contener solo números.",
                    "error",
                )
                return redirect(
                    url_for(
                        "horometros",
                        maquina_id=maquina_id,
                        modo=modo,
                    )
                )

            horometro_intervencion = int(
                texto_intervencion
            )

            ultimo_intervencion = (
                obtener_ultimo_horometro_intervencion(
                    datos,
                    maquina_id,
                )
            )

            if horometro_intervencion < ultimo_intervencion:
                flash(
                    "El horómetro de intervención debe ser igual o mayor "
                    f"que el último registrado: {ultimo_intervencion:,}.",
                    "error",
                )
                return redirect(
                    url_for(
                        "horometros",
                        maquina_id=maquina_id,
                        modo=modo,
                    )
                )

            limite_horometro = (
                horometro_actual
                if modo == "horometro"
                else horometro_anterior
            )

            if horometro_intervencion > limite_horometro:
                flash(
                    "El horómetro de intervención no puede ser mayor "
                    "que el último horómetro registrado de la máquina.",
                    "error",
                )
                return redirect(
                    url_for(
                        "horometros",
                        maquina_id=maquina_id,
                        modo=modo,
                    )
                )

            if not fecha_intervencion:
                flash(
                    "Debe seleccionar la fecha de intervención.",
                    "error",
                )
                return redirect(
                    url_for(
                        "horometros",
                        maquina_id=maquina_id,
                        modo=modo,
                    )
                )

            if not frecuencia and not cambio_aceite:
                flash(
                    "Seleccione una frecuencia o marque Cambio de aceite.",
                    "error",
                )
                return redirect(
                    url_for(
                        "horometros",
                        maquina_id=maquina_id,
                        modo=modo,
                    )
                )

            if novedad_distribuidor:
                if frecuencia != "semestral":
                    flash(
                        "La novedad de distribuidores solo puede "
                        "registrarse con frecuencia semestral.",
                        "error",
                    )
                    return redirect(
                        url_for(
                            "horometros",
                            maquina_id=maquina_id,
                            modo=modo,
                        )
                    )

                if distribuidor_novedad not in (
                    "agua",
                    "aceite",
                ):
                    flash(
                        "Debe seleccionar el distribuidor afectado.",
                        "error",
                    )
                    return redirect(
                        url_for(
                            "horometros",
                            maquina_id=maquina_id,
                            modo=modo,
                        )
                    )

        if modo == "horometro":
            maquina["horometro"] = horometro_actual
            maquina["horometro_actual"] = horometro_actual

        registro: dict[str, Any] = {
            "modo_registro": modo,
            "maquina_id": maquina_id,
            "maquina": maquina.get(
                "nombre",
                "",
            ),
            "fecha_registro": date.today().isoformat(),
            "horometro_anterior": horometro_anterior,
            "horometro_actual": (
                horometro_actual
                if modo == "horometro"
                else horometro_anterior
            ),
            "ejecuto_mantenimiento": ejecuto_mantenimiento,
        }

        if ejecuto_mantenimiento:
            registro.update(
                {
                    "horometro_intervencion":
                        horometro_intervencion,
                    "fecha_intervencion":
                        fecha_intervencion,
                    "frecuencia_ejecutada":
                        frecuencia,
                    "cambio_aceite":
                        cambio_aceite,
                    "novedad_distribuidor":
                        novedad_distribuidor,
                    "distribuidor_novedad":
                        distribuidor_novedad,
                }
            )

            registrar_mantenimiento(
                datos=datos,
                maquina=maquina,
                maquina_id=maquina_id,
                horometro_intervencion=horometro_intervencion,
                fecha_intervencion=fecha_intervencion,
                frecuencia=frecuencia,
                cambio_aceite=cambio_aceite,
                novedad_distribuidor=novedad_distribuidor,
                distribuidor_novedad=distribuidor_novedad,
            )

        datos.setdefault(
            "registros_horometros",
            [],
        ).append(
            registro
        )

        guardar_datos(
            datos
        )

        siguiente = siguiente_maquina(
            datos,
            maquina_id,
        )

        if modo == "mantenimiento":
            if siguiente:
                return redirect(
                    url_for(
                        "horometros",
                        maquina_id=siguiente["id"],
                        modo="mantenimiento",
                        mantenimiento_guardado=1,
                    )
                )

            return redirect(
                url_for(
                    "horometros",
                    modo="mantenimiento",
                    mantenimiento_terminado=1,
                )
            )

        if siguiente:
            return redirect(
                url_for(
                    "horometros",
                    maquina_id=siguiente["id"],
                    modo="horometro",
                    guardado=1,
                )
            )

        return redirect(
            url_for(
                "horometros",
                modo="horometro",
                terminado=1,
            )
        )

    tarjetas = []

    for maquina in maquinas:
        tarjeta = dict(
            maquina
        )

        tarjeta["registrada_hoy"] = (
            maquina_registrada_hoy(
                datos,
                maquina.get(
                    "id",
                    0,
                ),
            )
        )

        tarjetas.append(
            tarjeta
        )

    if request.args.get("guardado") == "1":
        mensaje = (
            "Horómetro guardado. Continúe con la siguiente máquina."
        )

    if request.args.get("terminado") == "1":
        mensaje = (
            "Registro diario de horómetros completado."
        )

    if request.args.get("mantenimiento_guardado") == "1":
        mensaje = (
            "Mantenimiento guardado. Continúe con la siguiente máquina."
        )

    if request.args.get("mantenimiento_terminado") == "1":
        mensaje = (
            "Registro de mantenimientos completado."
        )

    ultimo_intervencion = 0

    if maquina_seleccionada:
        ultimo_intervencion = (
            obtener_ultimo_horometro_intervencion(
                datos,
                maquina_seleccionada.get(
                    "id",
                    0,
                ),
            )
        )

    return render_template(
        "horometros.html",
        maquinas=tarjetas,
        maquina_seleccionada=maquina_seleccionada,
        mensaje=mensaje,
        fecha_hoy=date.today().isoformat(),
        ultimo_horometro_intervencion=ultimo_intervencion,
        modo_registro=modo_get,
    )


if __name__ == "__main__":
    app.run(
        host="127.0.0.1",
        port=5000,
        debug=True,
    )
