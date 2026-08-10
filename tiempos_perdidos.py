from __future__ import annotations

import calendar
import os
import re
import shutil
import unicodedata
import uuid
from collections import defaultdict
from copy import copy
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from flask import (
    Blueprint,
    current_app,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    url_for,
)
from werkzeug.utils import secure_filename


tiempos_perdidos_bp = Blueprint(
    "tiempos_perdidos",
    __name__,
    url_prefix="/tiempos-perdidos",
)

EXTENSIONES_PERMITIDAS = {".xlsx", ".xlsm"}

CLASES_ORDEN = {
    "PM01A": "correctivo",
    "PM01B": "preventivo",
    "PM02": "preventivo",
    "PM02A": "preventivo",
    "PM02B": "preventivo",
    "PMV02": "preventivo",
    "PM03": "cambio",
}

HOJAS_DESTINO_TIEMPOS = {
    "correctivo": "tiempos Perdidos TP IBERPLAST",
    "preventivo": "tiempos Mtto Prev TP",
    "cambio": "tiempos Cambios TP",
}

ORDEN_MAQUINAS = [
    "CCM1",
    "CCM2",
    "CCM3",
    "CCM4",
    "CCM5",
    "CCM6",
    "CCM7",
    "CCM8",
    "PMV2",
    "PMV4",
]

NOMBRES_MAQUINA_FORMATO = {
    "CCM1": "CCM 01",
    "CCM2": "CCM 02",
    "CCM3": "CCM 03",
    "CCM4": "CCM 04",
    "CCM5": "CCM 05",
    "CCM6": "CCM 06",
    "CCM7": "CCM 07",
    "CCM8": "CCM 08",
    "PMV2": "PMV 02",
    "PMV4": "PMV 04",
}

HOJAS_TRABAJOS = {
    "CCM1": "Trabajos Realizados CCM1",
    "CCM2": "Trabajos Realizados CCM2",
    "CCM3": "Trabajos Realizados CCM3",
    "CCM4": "Trabajos Realizados CCM64-1",
    "CCM5": "Trabajos Realizados CCM64-2",
    "CCM6": "Trabajos Realizados CCM 06",
    "CCM7": "Trabajos Realizados CCM 07",
    "CCM8": "Trabajos Realizados CCM 08",
    "PMV2": "Trabajos Realizados PMV 02",
    "PMV4": "Trabajos Realizados PMV 04",
}

MESES_ES = {
    1: "ENERO",
    2: "FEBRERO",
    3: "MARZO",
    4: "ABRIL",
    5: "MAYO",
    6: "JUNIO",
    7: "JULIO",
    8: "AGOSTO",
    9: "SEPTIEMBRE",
    10: "OCTUBRE",
    11: "NOVIEMBRE",
    12: "DICIEMBRE",
}

MESES_TEXTO = {
    "ENERO": 1,
    "FEBRERO": 2,
    "MARZO": 3,
    "ABRIL": 4,
    "MAYO": 5,
    "JUNIO": 6,
    "JULIO": 7,
    "AGOSTO": 8,
    "SEPTIEMBRE": 9,
    "SETIEMBRE": 9,
    "OCTUBRE": 10,
    "NOVIEMBRE": 11,
    "DICIEMBRE": 12,
}

ENCABEZADOS = {
    "dia": {
        "DIA",
        "FECHA",
    },
    "mes": {
        "MES",
    },
    "anio": {
        "ANO",
        "ANIO",
    },
    "equipo": {
        "EQUIPO",
        "MAQUINA",
    },
    "clase": {
        "CLASE DE ORDEN",
        "CLASE ORDEN",
    },
    "descripcion": {
        "DESCRIPCION PARADA",
        "DESCRIPCION DE PARADA",
        "DESCRIPCION",
    },
    "trabajo": {
        "TRABAJO REALIZADO",
    },
    "tiempo": {
        "TIEMPO PERDIDO MIN",
        "TIEMPO PERDIDO MINUTOS",
        "TIEMPO PERDIDO",
    },
    "estado": {
        "ESTADO",
        "ESTADO TRABAJO",
        "ESTADO DE TRABAJO",
        "ESTADO DEL TRABAJO",
    },
}


def _openpyxl():
    try:
        from openpyxl import load_workbook
        from openpyxl.comments import Comment
        from openpyxl.formula.translate import Translator
    except ModuleNotFoundError as error:
        raise RuntimeError(
            "Falta instalar openpyxl. Ejecute: "
            "python -m pip install -r requirements.txt"
        ) from error

    return load_workbook, Comment, Translator


def normalizar(valor: Any) -> str:
    if valor is None:
        return ""

    texto = str(valor).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(
        caracter
        for caracter in texto
        if not unicodedata.combining(caracter)
    )
    texto = texto.replace("º", " ").replace("°", " ")
    texto = re.sub(r"[\r\n\t]+", " ", texto)
    texto = re.sub(r"[^A-Z0-9]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def ruta_base() -> Path:
    return Path(current_app.root_path)


def carpeta_uploads() -> Path:
    ruta = ruta_base() / "uploads_tiempos"
    ruta.mkdir(parents=True, exist_ok=True)
    return ruta


def carpeta_salidas() -> Path:
    ruta = ruta_base() / "salidas"
    ruta.mkdir(parents=True, exist_ok=True)
    return ruta


def carpeta_formatos() -> Path:
    ruta = ruta_base() / "formatos"
    ruta.mkdir(parents=True, exist_ok=True)
    return ruta


def resolver_maquina(valor: Any) -> str | None:
    texto = normalizar(valor)

    if not texto:
        return None

    coincidencia = re.search(
        r"\bTERMOCOMPRESORA\s+SACMI(?:\s+N|\s+NO)?\s*0?([1-8])\b",
        texto,
    )
    if coincidencia:
        return f"CCM{int(coincidencia.group(1))}"

    coincidencia = re.search(
        r"\bSACMI\s+PMV(?:\s+N|\s+NO)?\s*0?(2|4)\b",
        texto,
    )
    if coincidencia:
        return f"PMV{int(coincidencia.group(1))}"

    if re.search(r"\bCCM\s*64\s*1\b", texto):
        return "CCM4"

    if re.search(r"\bCCM\s*64\s*2\b", texto):
        return "CCM5"

    coincidencia = re.search(r"\bCCM\s*0?([1-8])\b", texto)
    if coincidencia:
        return f"CCM{int(coincidencia.group(1))}"

    coincidencia = re.search(r"\bPMV\s*0?(2|4)\b", texto)
    if coincidencia:
        return f"PMV{int(coincidencia.group(1))}"

    return None


def clase_normalizada(valor: Any) -> str:
    return normalizar(valor).replace(" ", "")


def minutos_desde_valor(valor: Any) -> int:
    if valor in (None, ""):
        return 0

    if isinstance(valor, bool):
        return 0

    if isinstance(valor, (int, float)):
        return max(0, int(round(float(valor))))

    texto = str(valor).strip().replace(",", ".")

    try:
        return max(0, int(round(float(texto))))
    except ValueError:
        return 0


def mes_desde_valor(valor: Any) -> int | None:
    if valor in (None, ""):
        return None

    if isinstance(valor, (int, float)):
        numero = int(valor)
        return numero if 1 <= numero <= 12 else None

    return MESES_TEXTO.get(normalizar(valor))


def anio_desde_valor(valor: Any) -> int | None:
    try:
        numero = int(valor)
    except (TypeError, ValueError):
        return None

    return numero if 2000 <= numero <= 2100 else None


def fecha_desde_valores(
    valor_dia: Any,
    valor_mes: Any = None,
    valor_anio: Any = None,
) -> date | None:
    if isinstance(valor_dia, datetime):
        return valor_dia.date()

    if isinstance(valor_dia, date):
        return valor_dia

    texto = str(valor_dia or "").strip()

    for formato in (
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%d/%m/%y",
    ):
        try:
            return datetime.strptime(texto, formato).date()
        except ValueError:
            pass

    try:
        dia = int(float(texto))
    except (TypeError, ValueError):
        return None

    mes = mes_desde_valor(valor_mes)
    anio = anio_desde_valor(valor_anio)

    if not mes or not anio:
        return None

    try:
        return date(anio, mes, dia)
    except ValueError:
        return None


def _valor_fila(
    fila: tuple[Any, ...],
    columna_excel: int | None,
) -> Any:
    if columna_excel is None:
        return None

    indice = columna_excel - 1

    if indice < 0 or indice >= len(fila):
        return None

    return fila[indice]


def buscar_hoja_plastica(workbook):
    hojas = [
        hoja
        for hoja in workbook.worksheets
        if "PLAST" in normalizar(hoja.title)
        and "METAL" not in normalizar(hoja.title)
    ]

    if not hojas:
        raise ValueError(
            "No se encontró una hoja correspondiente a Plástica."
        )

    hojas.sort(key=lambda hoja: hoja.max_row, reverse=True)
    return hojas[0]


def buscar_fila_encabezados(hoja) -> tuple[int, dict[str, int]]:
    obligatorios = {
        "dia",
        "equipo",
        "clase",
        "descripcion",
        "trabajo",
        "tiempo",
    }

    for numero_fila, fila in enumerate(
        hoja.iter_rows(
            min_row=1,
            max_row=min(80, hoja.max_row),
            values_only=True,
        ),
        start=1,
    ):
        mapa: dict[str, int] = {}

        for columna, valor in enumerate(fila, start=1):
            encabezado = normalizar(valor)

            if not encabezado:
                continue

            for campo, opciones in ENCABEZADOS.items():
                if encabezado in opciones:
                    mapa[campo] = columna

            if encabezado.startswith("ESTADO DE TRABAJO"):
                mapa["estado"] = columna

        if obligatorios.issubset(mapa):
            return numero_fila, mapa

    raise ValueError(
        "No se encontraron los encabezados requeridos en la hoja de Plástica."
    )


def leer_entrega_turno(ruta: Path) -> dict[str, Any]:
    load_workbook, _, _ = _openpyxl()

    workbook = load_workbook(
        ruta,
        read_only=True,
        data_only=True,
    )

    try:
        hoja = buscar_hoja_plastica(workbook)
        fila_encabezados, columnas = buscar_fila_encabezados(hoja)

        columna_maxima = max(columnas.values())
        registros: list[dict[str, Any]] = []
        clases_ignoradas: set[str] = set()
        maquinas_ignoradas: set[str] = set()
        vacias_consecutivas = 0

        for numero_fila, fila in enumerate(
            hoja.iter_rows(
                min_row=fila_encabezados + 1,
                max_col=columna_maxima,
                values_only=True,
            ),
            start=fila_encabezados + 1,
        ):
            equipo_original = _valor_fila(
                fila,
                columnas.get("equipo"),
            )

            if (
                equipo_original is None
                or str(equipo_original).strip() == ""
            ):
                vacias_consecutivas += 1

                if vacias_consecutivas >= 10:
                    break

                continue

            vacias_consecutivas = 0

            maquina = resolver_maquina(equipo_original)

            if maquina not in ORDEN_MAQUINAS:
                maquinas_ignoradas.add(str(equipo_original).strip())
                continue

            clase_original = _valor_fila(
                fila,
                columnas.get("clase"),
            )
            clase = clase_normalizada(clase_original)

            if clase not in CLASES_ORDEN:
                if clase:
                    clases_ignoradas.add(str(clase_original).strip())
                continue

            fecha_evento = fecha_desde_valores(
                _valor_fila(fila, columnas.get("dia")),
                _valor_fila(fila, columnas.get("mes")),
                _valor_fila(fila, columnas.get("anio")),
            )

            if fecha_evento is None:
                continue

            minutos = minutos_desde_valor(
                _valor_fila(
                    fila,
                    columnas.get("tiempo"),
                )
            )

            if minutos <= 0:
                continue

            registros.append(
                {
                    "fila_origen": numero_fila,
                    "fecha": fecha_evento,
                    "dia": fecha_evento.day,
                    "maquina": maquina,
                    "equipo_origen": str(equipo_original).strip(),
                    "clase_orden": clase,
                    "categoria": CLASES_ORDEN[clase],
                    "descripcion": str(
                        _valor_fila(
                            fila,
                            columnas.get("descripcion"),
                        )
                        or ""
                    ).strip(),
                    "trabajo_realizado": str(
                        _valor_fila(
                            fila,
                            columnas.get("trabajo"),
                        )
                        or ""
                    ).strip(),
                    "estado": str(
                        _valor_fila(
                            fila,
                            columnas.get("estado"),
                        )
                        or ""
                    ).strip(),
                    "minutos": minutos,
                }
            )

        if not registros:
            raise ValueError(
                "No se encontraron registros válidos para las máquinas "
                "y clases de orden configuradas."
            )

        return {
            "hoja": hoja.title,
            "registros": registros,
            "clases_ignoradas": sorted(clases_ignoradas),
            "maquinas_no_reconocidas": sorted(maquinas_ignoradas),
        }

    finally:
        workbook.close()


def periodo_29_28(mes: int, anio: int) -> list[date]:
    if mes == 1:
        mes_anterior = 12
        anio_anterior = anio - 1
    else:
        mes_anterior = mes - 1
        anio_anterior = anio

    ultimo_anterior = calendar.monthrange(
        anio_anterior,
        mes_anterior,
    )[1]

    dia_inicial = 29 if ultimo_anterior >= 29 else ultimo_anterior

    inicio = date(
        anio_anterior,
        mes_anterior,
        dia_inicial,
    )
    fin = date(anio, mes, 28)

    resultado = []
    actual = inicio

    while actual <= fin:
        resultado.append(actual)
        actual += timedelta(days=1)

    return resultado


def determinar_periodo(
    registros: list[dict[str, Any]],
) -> tuple[int, int, list[date]]:
    fechas = [
        registro["fecha"]
        for registro in registros
        if registro.get("fecha")
    ]

    if not fechas:
        hoy = date.today()
        return hoy.month, hoy.year, periodo_29_28(hoy.month, hoy.year)

    fecha_maxima = max(fechas)

    if fecha_maxima.day <= 28:
        mes_reporte = fecha_maxima.month
        anio_reporte = fecha_maxima.year
    else:
        siguiente = fecha_maxima + timedelta(days=7)
        mes_reporte = siguiente.month
        anio_reporte = siguiente.year

    periodo = periodo_29_28(
        mes_reporte,
        anio_reporte,
    )

    return mes_reporte, anio_reporte, periodo


def agregar_unico(lista: list[str], valor: Any) -> None:
    texto = str(valor or "").strip()

    if not texto:
        return

    clave = normalizar(texto)

    if not any(
        normalizar(existente) == clave
        for existente in lista
    ):
        lista.append(texto)


def construir_agregados_diarios(
    registros: list[dict[str, Any]],
) -> dict[
    str,
    dict[tuple[date, str], dict[str, Any]],
]:
    resultado = {
        "correctivo": {},
        "preventivo": {},
        "cambio": {},
    }

    for registro in registros:
        categoria = registro["categoria"]
        clave = (
            registro["fecha"],
            registro["maquina"],
        )

        grupo = resultado[categoria].setdefault(
            clave,
            {
                "minutos": 0,
                "descripciones": [],
            },
        )

        grupo["minutos"] += registro["minutos"]

        agregar_unico(
            grupo["descripciones"],
            registro["descripcion"],
        )

    return resultado


def construir_fallas_correctivas(
    registros: list[dict[str, Any]],
) -> dict[str, list[dict[str, Any]]]:
    grupos: dict[
        tuple[str, str],
        dict[str, Any],
    ] = {}

    for registro in registros:
        if registro["categoria"] != "correctivo":
            continue

        descripcion_original = str(
            registro["descripcion"] or ""
        ).strip()

        descripcion_clave = normalizar(descripcion_original)

        if not descripcion_clave:
            descripcion_clave = "SIN DESCRIPCION"

        clave = (
            registro["maquina"],
            descripcion_clave,
        )

        grupo = grupos.setdefault(
            clave,
            {
                "maquina": registro["maquina"],
                "descripcion": (
                    descripcion_original
                    or "Sin descripción de parada"
                ),
                "minutos": 0,
                "trabajos_terminados": [],
                "terminaciones": 0,
            },
        )

        grupo["minutos"] += registro["minutos"]

        estado = normalizar(registro.get("estado"))

        if "TERMINADO" in estado:
            grupo["terminaciones"] += 1
            agregar_unico(
                grupo["trabajos_terminados"],
                registro.get("trabajo_realizado"),
            )

    por_maquina = {
        maquina: []
        for maquina in ORDEN_MAQUINAS
    }

    for grupo in grupos.values():
        grupo["eventos"] = max(
            1,
            grupo["terminaciones"],
        )
        grupo["trabajo_realizado"] = "\n".join(
            grupo["trabajos_terminados"]
        )
        por_maquina[groupo_maquina := grupo["maquina"]].append(grupo)

    for maquina in ORDEN_MAQUINAS:
        por_maquina[maquina].sort(
            key=lambda item: (
                -item["minutos"],
                normalizar(item["descripcion"]),
            )
        )

        for numero, falla in enumerate(
            por_maquina[maquina],
            start=1,
        ):
            falla["falla_no"] = numero

    return por_maquina


def todas_fallas_correctivas(
    fallas_por_maquina: dict[str, list[dict[str, Any]]],
) -> list[dict[str, Any]]:
    resultado = []

    for maquina in ORDEN_MAQUINAS:
        resultado.extend(
            dict(falla)
            for falla in fallas_por_maquina.get(maquina, [])
        )

    return resultado


def buscar_hoja(workbook, nombre: str):
    if nombre in workbook.sheetnames:
        return workbook[nombre]

    objetivo = normalizar(nombre)

    for hoja in workbook.worksheets:
        if normalizar(hoja.title) == objetivo:
            return hoja

    raise ValueError(
        f"No se encontró la hoja '{nombre}'."
    )


def buscar_fila_con_texto(
    hoja,
    texto: str,
    columna: int,
    inicio: int = 1,
) -> int | None:
    objetivo = normalizar(texto)

    for fila in range(inicio, hoja.max_row + 1):
        valor = hoja.cell(fila, columna).value

        if objetivo in normalizar(valor):
            return fila

    return None


def buscar_fila_total_formula(
    hoja,
    columna: int,
    inicio: int,
) -> int | None:
    for fila in range(inicio, hoja.max_row + 1):
        valor = hoja.cell(fila, columna).value

        if (
            isinstance(valor, str)
            and valor.upper().startswith("=SUM(")
        ):
            return fila

    return None


def copiar_estilo_celda(origen, destino) -> None:
    if origen.has_style:
        destino._style = copy(origen._style)

    destino.number_format = origen.number_format
    destino.font = copy(origen.font)
    destino.fill = copy(origen.fill)
    destino.border = copy(origen.border)
    destino.alignment = copy(origen.alignment)
    destino.protection = copy(origen.protection)


def copiar_estilo_fila(
    hoja,
    fila_origen: int,
    fila_destino: int,
    columnas: range,
) -> None:
    for columna in columnas:
        copiar_estilo_celda(
            hoja.cell(fila_origen, columna),
            hoja.cell(fila_destino, columna),
        )

    altura = hoja.row_dimensions[fila_origen].height

    if altura is not None:
        hoja.row_dimensions[fila_destino].height = altura


def es_merged_cell(celda) -> bool:
    return celda.__class__.__name__ == "MergedCell"


def limpiar_rango(
    hoja,
    fila_inicio: int,
    fila_fin: int,
    columna_inicio: int,
    columna_fin: int,
) -> None:
    for fila in range(fila_inicio, fila_fin + 1):
        for columna in range(columna_inicio, columna_fin + 1):
            celda = hoja.cell(fila, columna)

            if es_merged_cell(celda):
                continue

            celda.value = None
            celda.comment = None


def fila_inicio_matriz(hoja) -> int:
    fila_encabezado = None

    for fila in range(1, min(20, hoja.max_row) + 1):
        if "DIA PLANTA" in normalizar(
            hoja.cell(fila, 1).value
        ):
            fila_encabezado = fila
            break

    if fila_encabezado is None:
        raise ValueError(
            f"No se encontró 'DÍA / planta' en la hoja '{hoja.title}'."
        )

    fin_encabezado = fila_encabezado

    for rango in hoja.merged_cells.ranges:
        if (
            rango.min_col == 1
            and rango.max_col == 1
            and rango.min_row <= fila_encabezado <= rango.max_row
        ):
            fin_encabezado = max(
                fin_encabezado,
                rango.max_row,
            )

    return fin_encabezado + 1


def columnas_maquinas(hoja) -> dict[str, int]:
    resultado: dict[str, int] = {}

    for fila in range(1, min(20, hoja.max_row) + 1):
        if "DIA PLANTA" not in normalizar(
            hoja.cell(fila, 1).value
        ):
            continue

        for columna in range(2, hoja.max_column + 1):
            maquina = resolver_maquina(
                hoja.cell(fila, columna).value
            )

            if maquina in ORDEN_MAQUINAS:
                resultado[maquina] = columna

        break

    faltantes = [
        maquina
        for maquina in ORDEN_MAQUINAS
        if maquina not in resultado
    ]

    if faltantes:
        raise ValueError(
            "El formato no contiene las columnas: "
            + ", ".join(faltantes)
        )

    return resultado


def comentario_descripciones(descripciones: list[str]) -> str:
    unicas: list[str] = []

    for descripcion in descripciones:
        agregar_unico(unicas, descripcion)

    if not unicas:
        return "Sin descripción de parada."

    if len(unicas) == 1:
        return unicas[0]

    return "\n".join(
        f"{numero}. {descripcion}"
        for numero, descripcion in enumerate(
            unicas,
            start=1,
        )
    )


def preparar_matriz(
    hoja,
    periodo: list[date],
    incluir_total_dia: bool = False,
) -> tuple[
    dict[date, int],
    dict[str, int],
    int,
]:
    inicio = fila_inicio_matriz(hoja)
    columnas = columnas_maquinas(hoja)

    fila_total = buscar_fila_total_formula(
        hoja,
        min(columnas.values()),
        inicio,
    )

    if fila_total is None:
        fila_total = hoja.max_row

    capacidad = fila_total - inicio

    if len(periodo) > capacidad:
        extra = len(periodo) - capacidad
        hoja.insert_rows(fila_total, amount=extra)

        for fila in range(
            fila_total,
            fila_total + extra,
        ):
            copiar_estilo_fila(
                hoja,
                inicio,
                fila,
                range(1, hoja.max_column + 1),
            )

        fila_total += extra

    limpiar_rango(
        hoja,
        inicio,
        fila_total - 1,
        1,
        hoja.max_column,
    )

    filas_fecha: dict[date, int] = {}

    for offset, fecha_actual in enumerate(periodo):
        fila = inicio + offset
        filas_fecha[fecha_actual] = fila

        celda_fecha = hoja.cell(fila, 1)
        celda_fecha.value = fecha_actual
        celda_fecha.number_format = "dd/mm/yyyy"

        for columna in columnas.values():
            hoja.cell(fila, columna).value = 0

        if incluir_total_dia:
            columna_total = max(columnas.values()) + 1
            hoja.cell(
                fila,
                columna_total,
            ).value = (
                f"=SUM("
                f"{hoja.cell(fila, min(columnas.values())).coordinate}:"
                f"{hoja.cell(fila, max(columnas.values())).coordinate}"
                f")"
            )

    ultima_fila_periodo = inicio + len(periodo) - 1

    for fila in range(
        ultima_fila_periodo + 1,
        fila_total,
    ):
        for columna in range(1, hoja.max_column + 1):
            celda = hoja.cell(fila, columna)

            if not es_merged_cell(celda):
                celda.value = None
                celda.comment = None

    for maquina, columna in columnas.items():
        hoja.cell(
            fila_total,
            columna,
        ).value = (
            f"=SUM("
            f"{hoja.cell(inicio, columna).coordinate}:"
            f"{hoja.cell(ultima_fila_periodo, columna).coordinate}"
            f")"
        )

    if incluir_total_dia:
        columna_total = max(columnas.values()) + 1
        hoja.cell(
            fila_total,
            columna_total,
        ).value = (
            f"=SUM("
            f"{hoja.cell(fila_total, min(columnas.values())).coordinate}:"
            f"{hoja.cell(fila_total, max(columnas.values())).coordinate}"
            f")"
        )

    return filas_fecha, columnas, fila_total


def llenar_matriz(
    hoja,
    datos: dict[tuple[date, str], dict[str, Any]],
    periodo: list[date],
    incluir_total_dia: bool = False,
) -> None:
    _, Comment, _ = _openpyxl()

    filas_fecha, columnas, _ = preparar_matriz(
        hoja,
        periodo,
        incluir_total_dia=incluir_total_dia,
    )

    fechas_validas = set(periodo)

    for (fecha_evento, maquina), grupo in datos.items():
        if (
            fecha_evento not in fechas_validas
            or maquina not in columnas
        ):
            continue

        celda = hoja.cell(
            filas_fecha[fecha_evento],
            columnas[maquina],
        )

        celda.value = grupo["minutos"]
        celda.comment = Comment(
            comentario_descripciones(
                grupo["descripciones"]
            ),
            "CMMS Industrial",
        )


def asegurar_capacidad_tabla(
    hoja,
    fila_inicio: int,
    fila_total: int,
    cantidad: int,
    columnas_estilo: range,
) -> int:
    capacidad = fila_total - fila_inicio

    if cantidad <= capacidad:
        return fila_total

    extra = cantidad - capacidad
    hoja.insert_rows(fila_total, amount=extra)

    for fila in range(
        fila_total,
        fila_total + extra,
    ):
        copiar_estilo_fila(
            hoja,
            fila_inicio,
            fila,
            columnas_estilo,
        )

    return fila_total + extra


def llenar_trabajos_realizados(
    workbook,
    fallas_por_maquina: dict[str, list[dict[str, Any]]],
) -> None:
    for maquina in ORDEN_MAQUINAS:
        hoja = buscar_hoja(
            workbook,
            HOJAS_TRABAJOS[maquina],
        )

        fila_inicio = 8
        fila_total = buscar_fila_total_formula(
            hoja,
            4,
            fila_inicio,
        )

        if fila_total is None:
            fila_total = max(
                hoja.max_row + 1,
                fila_inicio + 13,
            )

        fallas = fallas_por_maquina.get(maquina, [])

        fila_total = asegurar_capacidad_tabla(
            hoja,
            fila_inicio,
            fila_total,
            len(fallas),
            range(1, 6),
        )

        limpiar_rango(
            hoja,
            fila_inicio,
            fila_total - 1,
            1,
            5,
        )

        for offset, falla in enumerate(fallas):
            fila = fila_inicio + offset
            copiar_estilo_fila(
                hoja,
                fila_inicio,
                fila,
                range(1, 6),
            )

            hoja.cell(fila, 1).value = falla["falla_no"]
            hoja.cell(fila, 2).value = falla["descripcion"]
            hoja.cell(fila, 3).value = falla["trabajo_realizado"]
            hoja.cell(fila, 4).value = falla["minutos"]
            hoja.cell(fila, 5).value = falla["eventos"]

        ultima = (
            fila_inicio + len(fallas) - 1
            if fallas
            else fila_inicio
        )

        hoja.cell(
            fila_total,
            4,
        ).value = f"=SUM(D{fila_inicio}:D{ultima})"


def descombinar_columna_g_fallas(
    hoja,
    fila_inicio: int,
    fila_total: int,
) -> None:
    for rango in list(hoja.merged_cells.ranges):
        if (
            rango.min_col == 7
            and rango.max_col == 7
            and rango.max_row >= fila_inicio
            and rango.min_row < fila_total
        ):
            hoja.unmerge_cells(str(rango))


def llenar_fallas_tapas(
    hoja,
    fallas: list[dict[str, Any]],
) -> None:
    fila_inicio = 4
    fila_total = buscar_fila_con_texto(
        hoja,
        "TOTAL TIEMPOS PERDIDOS AREA",
        4,
        inicio=fila_inicio,
    )

    if fila_total is None:
        fila_total = 41

    descombinar_columna_g_fallas(
        hoja,
        fila_inicio,
        fila_total,
    )

    fila_total = asegurar_capacidad_tabla(
        hoja,
        fila_inicio,
        fila_total,
        len(fallas),
        range(2, 8),
    )

    limpiar_rango(
        hoja,
        fila_inicio,
        fila_total - 1,
        2,
        7,
    )

    posiciones: dict[str, list[int]] = defaultdict(list)

    for offset, falla in enumerate(fallas):
        fila = fila_inicio + offset

        copiar_estilo_fila(
            hoja,
            fila_inicio,
            fila,
            range(2, 8),
        )

        hoja.cell(fila, 2).value = falla["falla_no"]
        hoja.cell(fila, 3).value = falla["maquina"]
        hoja.cell(fila, 4).value = falla["descripcion"]
        hoja.cell(fila, 5).value = falla["minutos"]
        hoja.cell(fila, 6).value = falla["eventos"]

        posiciones[falla["maquina"]].append(fila)

    for maquina, filas in posiciones.items():
        primera = filas[0]
        ultima = filas[-1]

        hoja.cell(
            primera,
            7,
        ).value = f"=SUM(E{primera}:E{ultima})"

        if ultima > primera:
            hoja.merge_cells(
                start_row=primera,
                start_column=7,
                end_row=ultima,
                end_column=7,
            )

    ultima_datos = (
        fila_inicio + len(fallas) - 1
        if fallas
        else fila_inicio
    )

    hoja.cell(
        fila_total,
        4,
    ).value = "TOTAL TIEMPOS PERDIDOS AREA"
    hoja.cell(
        fila_total,
        5,
    ).value = f"=SUM(E{fila_inicio}:E{ultima_datos})"


def llenar_datos_pareto(
    hoja,
    fallas: list[dict[str, Any]],
) -> None:
    fila_inicio = 4
    fila_total = buscar_fila_con_texto(
        hoja,
        "TOTAL TIEMPOS PERDIDOS AREA",
        6,
        inicio=fila_inicio,
    )

    if fila_total is None:
        fila_total = 40

    ordenadas = sorted(
        fallas,
        key=lambda item: (
            -item["minutos"],
            item["maquina"],
            normalizar(item["descripcion"]),
        ),
    )

    fila_total = asegurar_capacidad_tabla(
        hoja,
        fila_inicio,
        fila_total,
        len(ordenadas),
        range(4, 9),
    )

    limpiar_rango(
        hoja,
        fila_inicio,
        fila_total - 1,
        4,
        8,
    )

    for numero, falla in enumerate(
        ordenadas,
        start=1,
    ):
        fila = fila_inicio + numero - 1

        copiar_estilo_fila(
            hoja,
            fila_inicio,
            fila,
            range(4, 9),
        )

        hoja.cell(fila, 4).value = numero
        hoja.cell(fila, 5).value = falla["maquina"]
        hoja.cell(fila, 6).value = falla["descripcion"]
        hoja.cell(fila, 7).value = falla["minutos"]
        hoja.cell(fila, 8).value = falla["eventos"]

    ultima_datos = (
        fila_inicio + len(ordenadas) - 1
        if ordenadas
        else fila_inicio
    )

    hoja.cell(
        fila_total,
        6,
    ).value = "TOTAL TIEMPOS PERDIDOS AREA"
    hoja.cell(
        fila_total,
        7,
    ).value = f"=SUM(G{fila_inicio}:G{ultima_datos})"

    hoja["J4"] = "PARETO 80/20"
    hoja["J5"] = "EVENTOS 20%"
    hoja["K5"] = (
        f"=COUNTA(D{fila_inicio}:D{ultima_datos})*20%"
    )
    hoja["J6"] = "TIEMPO 80 %"
    hoja["K6"] = f"=G{fila_total}*80%"


def actualizar_mes_anio_reporte(
    workbook,
    mes: int,
    anio: int,
) -> None:
    for hoja in workbook.worksheets:
        for fila in range(1, min(15, hoja.max_row) + 1):
            for columna in range(1, min(12, hoja.max_column) + 1):
                valor = normalizar(
                    hoja.cell(fila, columna).value
                )

                if valor == "MES":
                    destino = hoja.cell(fila, columna + 1)

                    if not es_merged_cell(destino):
                        destino.value = MESES_ES[mes].capitalize()

                if valor == "ANO" or valor == "ANIO":
                    destino = hoja.cell(fila, columna + 1)

                    if not es_merged_cell(destino):
                        destino.value = anio


def diligenciar_formato_tiempos(
    plantilla: Path,
    salida: Path,
    agregados: dict[
        str,
        dict[tuple[date, str], dict[str, Any]],
    ],
    fallas_por_maquina: dict[str, list[dict[str, Any]]],
    periodo: list[date],
    mes: int,
    anio: int,
) -> None:
    load_workbook, _, _ = _openpyxl()

    shutil.copy2(plantilla, salida)

    workbook = load_workbook(
        salida,
        keep_vba=salida.suffix.lower() == ".xlsm",
    )

    try:
        for categoria, nombre_hoja in HOJAS_DESTINO_TIEMPOS.items():
            llenar_matriz(
                buscar_hoja(workbook, nombre_hoja),
                agregados[categoria],
                periodo,
                incluir_total_dia=False,
            )

        llenar_trabajos_realizados(
            workbook,
            fallas_por_maquina,
        )

        actualizar_mes_anio_reporte(
            workbook,
            mes,
            anio,
        )

        workbook.save(salida)

    finally:
        workbook.close()


def diligenciar_pareto(
    plantilla: Path,
    salida: Path,
    agregados: dict[
        str,
        dict[tuple[date, str], dict[str, Any]],
    ],
    fallas_por_maquina: dict[str, list[dict[str, Any]]],
    periodo: list[date],
    mes: int,
    anio: int,
) -> None:
    load_workbook, _, _ = _openpyxl()

    shutil.copy2(plantilla, salida)

    workbook = load_workbook(
        salida,
        keep_vba=salida.suffix.lower() == ".xlsm",
    )

    try:
        llenar_matriz(
            buscar_hoja(
                workbook,
                "TIEMPOS PERDIDOS",
            ),
            agregados["correctivo"],
            periodo,
            incluir_total_dia=True,
        )

        fallas = todas_fallas_correctivas(
            fallas_por_maquina
        )

        llenar_fallas_tapas(
            buscar_hoja(
                workbook,
                "FALLAS TAPAS",
            ),
            fallas,
        )

        llenar_datos_pareto(
            buscar_hoja(
                workbook,
                "DATOS PARETO",
            ),
            fallas,
        )

        actualizar_mes_anio_reporte(
            workbook,
            mes,
            anio,
        )

        workbook.save(salida)

    finally:
        workbook.close()


def localizar_plantilla(
    candidatos: list[str],
    palabras: list[str],
) -> Path:
    carpeta = carpeta_formatos()

    for nombre in candidatos:
        ruta = carpeta / nombre

        if ruta.exists() and not ruta.name.startswith("~$"):
            return ruta

    archivos = [
        archivo
        for archivo in list(carpeta.glob("*.xlsx"))
        + list(carpeta.glob("*.xlsm"))
        if not archivo.name.startswith("~$")
    ]

    for archivo in archivos:
        nombre = normalizar(archivo.name)

        if all(
            normalizar(palabra) in nombre
            for palabra in palabras
        ):
            return archivo

    raise FileNotFoundError(
        "No se encontró el formato requerido en la carpeta 'formatos'."
    )


def resumen_proceso(
    registros: list[dict[str, Any]],
) -> dict[str, int]:
    resultado = {
        "total_registros": len(registros),
        "correctivos": 0,
        "preventivos": 0,
        "cambios": 0,
        "min_correctivo": 0,
        "min_preventivo": 0,
        "min_cambio": 0,
    }

    for registro in registros:
        categoria = registro["categoria"]

        if categoria == "correctivo":
            resultado["correctivos"] += 1
            resultado["min_correctivo"] += registro["minutos"]

        elif categoria == "preventivo":
            resultado["preventivos"] += 1
            resultado["min_preventivo"] += registro["minutos"]

        elif categoria == "cambio":
            resultado["cambios"] += 1
            resultado["min_cambio"] += registro["minutos"]

    return resultado


def nombre_archivo_salida(
    tipo: str,
    mes: int,
    anio: int,
) -> str:
    return (
        f"{tipo} {MESES_ES[mes]} {anio}.xlsx"
    )


@tiempos_perdidos_bp.route(
    "/",
    methods=["GET", "POST"],
)
def index():
    resultado = None
    ruta_entrada: Path | None = None

    if request.method == "POST":
        archivo = request.files.get("archivo_entrega")

        if (
            archivo is None
            or not archivo.filename
        ):
            flash(
                "Seleccione el archivo de Entrega de Turno.",
                "error",
            )
            return redirect(
                url_for("tiempos_perdidos.index")
            )

        extension = Path(
            archivo.filename
        ).suffix.lower()

        if extension not in EXTENSIONES_PERMITIDAS:
            flash(
                "El archivo debe ser .xlsx o .xlsm.",
                "error",
            )
            return redirect(
                url_for("tiempos_perdidos.index")
            )

        identificador = uuid.uuid4().hex[:10]
        nombre_seguro = secure_filename(
            archivo.filename
        )
        ruta_entrada = (
            carpeta_uploads()
            / f"{identificador}_{nombre_seguro}"
        )

        try:
            archivo.save(ruta_entrada)

            lectura = leer_entrega_turno(
                ruta_entrada
            )
            registros = lectura["registros"]

            mes, anio, periodo = determinar_periodo(
                registros
            )

            agregados = construir_agregados_diarios(
                registros
            )
            fallas_por_maquina = construir_fallas_correctivas(
                registros
            )

            plantilla_tiempos = localizar_plantilla(
                [
                    "TIEMPOS TAPA PLASTICA.xlsx",
                    "TIEMPOS JUNIO 2026 TAPA PLASTICA IBERPLAST.xlsx",
                ],
                ["TIEMPOS", "TAPA", "PLASTICA"],
            )

            plantilla_pareto = localizar_plantilla(
                [
                    "PARETO 80-20 TAPAS.xlsx",
                    "PARETO 80-20 TAPAS JUNIO 2026.xlsx",
                ],
                ["PARETO", "TAPAS"],
            )

            salida_tiempos = (
                carpeta_salidas()
                / nombre_archivo_salida(
                    "TIEMPOS TAPA PLASTICA",
                    mes,
                    anio,
                )
            )

            salida_pareto = (
                carpeta_salidas()
                / nombre_archivo_salida(
                    "PARETO 80-20 TAPAS",
                    mes,
                    anio,
                )
            )

            diligenciar_formato_tiempos(
                plantilla_tiempos,
                salida_tiempos,
                agregados,
                fallas_por_maquina,
                periodo,
                mes,
                anio,
            )

            diligenciar_pareto(
                plantilla_pareto,
                salida_pareto,
                agregados,
                fallas_por_maquina,
                periodo,
                mes,
                anio,
            )

            resumen = resumen_proceso(registros)

            resultado = {
                **resumen,
                "hoja": lectura["hoja"],
                "no_reconocidas": lectura[
                    "maquinas_no_reconocidas"
                ],
                "clases_ignoradas": lectura[
                    "clases_ignoradas"
                ],
                "archivo_tiempos": salida_tiempos.name,
                "archivo_pareto": salida_pareto.name,
                "periodo": (
                    periodo[0].strftime("%d/%m/%Y")
                    + " - "
                    + periodo[-1].strftime("%d/%m/%Y")
                ),
            }

        except Exception as error:
            import traceback

            traceback.print_exc()

            flash(
                str(error),
                "error",
            )

            return redirect(
                url_for("tiempos_perdidos.index")
            )

        finally:
            if (
                ruta_entrada is not None
                and ruta_entrada.exists()
            ):
                try:
                    ruta_entrada.unlink()
                except OSError:
                    pass

    return render_template(
        "tiempos_perdidos.html",
        resultado=resultado,
    )


@tiempos_perdidos_bp.route(
    "/descargar/<path:nombre>",
    methods=["GET"],
)
def descargar(nombre: str):

    nombre_archivo = Path(
        nombre
    ).name

    ruta = (
        carpeta_salidas()
        / nombre_archivo
    )

    if not ruta.exists():

        flash(
            "El archivo solicitado no existe.",
            "error",
        )

        return redirect(
            url_for(
                "tiempos_perdidos.index"
            )
        )

    return send_file(
        ruta,
        as_attachment=True,
        download_name=nombre_archivo,
    )