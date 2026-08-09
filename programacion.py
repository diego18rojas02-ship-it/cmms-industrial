import json
import os
from calendar import monthrange
from datetime import date, datetime, timedelta
from typing import Any

from flask import Blueprint, jsonify, render_template, request

programacion_bp = Blueprint("programacion", __name__)


def _archivo_datos() -> str:
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "datos_cmms.json")


def cargar_datos() -> dict[str, Any]:
    ruta = _archivo_datos()
    if not os.path.exists(ruta):
        return {"maquinas": [], "registros_horometros": [], "mantenimientos": {}, "programaciones": []}

    with open(ruta, "r", encoding="utf-8") as archivo:
        datos = json.load(archivo)

    datos.setdefault("maquinas", [])
    datos.setdefault("registros_horometros", [])
    datos.setdefault("mantenimientos", {})
    datos.setdefault("programaciones", [])
    return datos


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
            pass
    return None


def sumar_meses(fecha: date, meses: int) -> date:
    total = fecha.month - 1 + meses
    anio = fecha.year + total // 12
    mes = total % 12 + 1
    dia = min(fecha.day, monthrange(anio, mes)[1])
    return date(anio, mes, dia)


def pascua(anio: int) -> date:
    a = anio % 19
    b = anio // 100
    c = anio % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mes = (h + l - 7 * m + 114) // 31
    dia = ((h + l - 7 * m + 114) % 31) + 1
    return date(anio, mes, dia)


def mover_a_lunes(fecha: date) -> date:
    dias = (7 - fecha.weekday()) % 7
    return fecha + timedelta(days=dias)


def festivos_colombia(anio: int) -> dict[str, str]:
    domingo_pascua = pascua(anio)
    festivos = {
        date(anio, 1, 1): "Año Nuevo",
        mover_a_lunes(date(anio, 1, 6)): "Reyes Magos",
        mover_a_lunes(date(anio, 3, 19)): "San José",
        domingo_pascua - timedelta(days=3): "Jueves Santo",
        domingo_pascua - timedelta(days=2): "Viernes Santo",
        date(anio, 5, 1): "Día del Trabajo",
        mover_a_lunes(domingo_pascua + timedelta(days=39)): "Ascensión",
        mover_a_lunes(domingo_pascua + timedelta(days=60)): "Corpus Christi",
        mover_a_lunes(domingo_pascua + timedelta(days=68)): "Sagrado Corazón",
        mover_a_lunes(date(anio, 6, 29)): "San Pedro y San Pablo",
        date(anio, 7, 20): "Independencia",
        date(anio, 8, 7): "Batalla de Boyacá",
        mover_a_lunes(date(anio, 8, 15)): "Asunción",
        mover_a_lunes(date(anio, 10, 12)): "Día de la Raza",
        mover_a_lunes(date(anio, 11, 1)): "Todos los Santos",
        mover_a_lunes(date(anio, 11, 11)): "Independencia de Cartagena",
        date(anio, 12, 8): "Inmaculada Concepción",
        date(anio, 12, 25): "Navidad",
    }
    return {fecha.isoformat(): nombre for fecha, nombre in festivos.items()}


def obtener_horometro_actual(maquina: dict[str, Any]) -> int:
    valor = maquina.get("horometro", maquina.get("horometro_actual", 0))
    try:
        return int(valor)
    except (TypeError, ValueError):
        return 0


def obtener_controles(maquina: dict[str, Any]) -> list[str]:
    controles = []
    for control in maquina.get("controles", []):
        if isinstance(control, str):
            controles.append(control)
        elif isinstance(control, dict) and control.get("codigo"):
            controles.append(str(control["codigo"]))
    return controles


def horas_por_dia_promedio(datos: dict[str, Any], maquina_id: int) -> float:
    registros = []

    for registro in datos.get("registros_horometros", []):
        if int(registro.get("maquina_id", 0)) != int(maquina_id):
            continue
        if registro.get("modo_registro", "horometro") != "horometro":
            continue

        fecha_registro = convertir_fecha(registro.get("fecha_registro"))
        try:
            horometro = int(registro.get("horometro_actual", 0))
        except (TypeError, ValueError):
            continue

        if fecha_registro:
            registros.append((fecha_registro, horometro))

    registros.sort(key=lambda x: x[0])

    if len(registros) >= 2:
        fecha_inicial, horometro_inicial = registros[0]
        fecha_final, horometro_final = registros[-1]
        dias = (fecha_final - fecha_inicial).days
        horas = horometro_final - horometro_inicial
        if dias > 0 and horas > 0:
            return max(1.0, horas / dias)

    return 24.0


def fecha_objetivo_por_horas(
    datos: dict[str, Any],
    maquina: dict[str, Any],
    horometro_base: int,
    horas_objetivo: int,
) -> tuple[date, int, float]:
    actual = obtener_horometro_actual(maquina)
    horas_trabajadas = max(0, actual - horometro_base)
    horas_faltantes = horas_objetivo - horas_trabajadas
    ritmo = horas_por_dia_promedio(datos, int(maquina.get("id", 0)))

    if horas_faltantes <= 0:
        return date.today(), horas_faltantes, ritmo

    dias_faltantes = horas_faltantes / ritmo
    fecha_objetivo = date.today() + timedelta(days=max(1, round(dias_faltantes)))
    return fecha_objetivo, horas_faltantes, ritmo


def siguiente_frecuencia_rutina(ultima_frecuencia: str) -> str:
    return "mensual" if (ultima_frecuencia or "").lower() == "quincenal" else "quincenal"


def analizar_rutina(maquina: dict[str, Any], registro: dict[str, Any]) -> dict[str, Any] | None:
    if not registro:
        return None

    fecha = convertir_fecha(registro.get("fecha_intervencion"))
    if not fecha:
        return None

    ultima_frecuencia = str(registro.get("frecuencia", "quincenal")).lower()
    siguiente = siguiente_frecuencia_rutina(ultima_frecuencia)
    dias_objetivo = 30 if siguiente == "mensual" else 15
    vencimiento = fecha + timedelta(days=dias_objetivo)
    dias_restantes = (vencimiento - date.today()).days

    return {
        "tipo": "rutina",
        "frecuencia": siguiente,
        "fecha_objetivo": vencimiento.isoformat(),
        "dias_restantes": dias_restantes,
        "tiempo": "8 horas" if siguiente == "mensual" else "4 horas",
        "prioridad": 0 if dias_restantes < 0 else 1 if dias_restantes <= 4 else 2,
    }


def analizar_mayor_horas(
    datos: dict[str, Any],
    maquina: dict[str, Any],
    registro: dict[str, Any],
    frecuencia: str,
    tipo: str,
    horas_objetivo: int,
) -> dict[str, Any] | None:
    if not registro:
        return None

    try:
        base = int(registro.get("horometro_intervencion", 0))
    except (TypeError, ValueError):
        return None

    fecha_objetivo, faltantes, ritmo = fecha_objetivo_por_horas(
        datos,
        maquina,
        base,
        horas_objetivo,
    )

    dias_restantes = (fecha_objetivo - date.today()).days

    if faltantes <= 0:
        estado, prioridad = "vencido", 0
    elif dias_restantes <= 4:
        estado, prioridad = "proximo", 1
    else:
        estado, prioridad = "en_tiempo", 2

    return {
        "tipo": tipo,
        "frecuencia": frecuencia,
        "fecha_objetivo": fecha_objetivo.isoformat(),
        "dias_restantes": dias_restantes,
        "horas_faltantes": faltantes,
        "ritmo": round(ritmo, 1),
        "estado": estado,
        "prioridad": prioridad,
        "tiempo": "Jornada",
    }


def analizar_mayor_meses(
    registro: dict[str, Any],
    frecuencia: str,
    tipo: str,
    meses: int,
) -> dict[str, Any] | None:
    if not registro:
        return None

    fecha = convertir_fecha(registro.get("fecha_intervencion"))
    if not fecha:
        return None

    objetivo = sumar_meses(fecha, meses)
    dias_restantes = (objetivo - date.today()).days

    if dias_restantes < 0:
        estado, prioridad = "vencido", 0
    elif dias_restantes <= 4:
        estado, prioridad = "proximo", 1
    else:
        estado, prioridad = "en_tiempo", 2

    return {
        "tipo": tipo,
        "frecuencia": frecuencia,
        "fecha_objetivo": objetivo.isoformat(),
        "dias_restantes": dias_restantes,
        "estado": estado,
        "prioridad": prioridad,
        "tiempo": "Jornada",
    }


def analizar_maquinas(datos: dict[str, Any]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    rutinas = []
    mayores = []
    mantenimientos = datos.get("mantenimientos", {})

    for maquina in datos.get("maquinas", []):
        maquina_id = str(maquina.get("id", 0))
        nombre = maquina.get("nombre", f"Máquina {maquina_id}")
        registros = mantenimientos.get(maquina_id, {})

        rutina = analizar_rutina(maquina, registros.get("rutina", {}))
        if rutina:
            rutina.update({"maquina_id": int(maquina.get("id", 0)), "maquina": nombre})
            rutinas.append(rutina)

        controles = obtener_controles(maquina)

        if "punzones" in controles:
            dato = analizar_mayor_horas(
                datos, maquina, registros.get("punzones", {}),
                "Trimestral", "Punzones", 1800
            )
            if dato:
                dato.update({"maquina_id": int(maquina.get("id", 0)), "maquina": nombre})
                mayores.append(dato)

        if "distribuidores" in controles:
            distribuidores = registros.get("distribuidores", {})
            fechas = []
            for clave in ("distribuidor_agua", "distribuidor_aceite"):
                fecha = convertir_fecha(distribuidores.get(clave, {}).get("fecha_intervencion"))
                if fecha:
                    fechas.append(fecha)

            if fechas:
                dato = analizar_mayor_meses(
                    {"fecha_intervencion": min(fechas).isoformat()},
                    "Semestral",
                    "Distribuidores",
                    6,
                )
                if dato:
                    dato.update({"maquina_id": int(maquina.get("id", 0)), "maquina": nombre})
                    mayores.append(dato)

        if "bridas" in controles:
            dato = analizar_mayor_meses(
                registros.get("bridas", {}),
                "Anual",
                "Bridas y actuadores",
                12,
            )
            if dato:
                dato.update({"maquina_id": int(maquina.get("id", 0)), "maquina": nombre})
                mayores.append(dato)

        if "aceite" in controles:
            dato = analizar_mayor_horas(
                datos, maquina, registros.get("aceite", {}),
                "Aceite", "Cambio de aceite", 40000
            )
            if dato:
                dato.update({"maquina_id": int(maquina.get("id", 0)), "maquina": nombre})
                mayores.append(dato)

    rutinas.sort(key=lambda item: (item["prioridad"], item["fecha_objetivo"]))
    mayores.sort(key=lambda item: (item["prioridad"], item["fecha_objetivo"]))
    return rutinas, mayores


def parsear_fechas(valores: list[str]) -> list[date]:
    resultado = []
    for valor in valores:
        fecha = convertir_fecha(valor)
        if fecha:
            resultado.append(fecha)
    return sorted(set(resultado))


def inicio_semana(fecha: date) -> date:
    return fecha - timedelta(days=fecha.weekday())


def misma_semana(fechas: list[date]) -> bool:
    if not fechas:
        return False
    semana = inicio_semana(fechas[0])
    return all(inicio_semana(fecha) == semana for fecha in fechas)


def es_dia_habil(fecha: date, festivos: dict[str, str]) -> bool:
    return fecha.weekday() < 5 and fecha.isoformat() not in festivos


def mejor_dia_para_rutina(item: dict[str, Any], seleccionados: list[date]) -> date | None:
    if not seleccionados:
        return None

    objetivo = convertir_fecha(item.get("fecha_objetivo"))
    if not objetivo:
        return seleccionados[0]

    if objetivo < seleccionados[0]:
        return seleccionados[0]

    candidatos = [
        fecha
        for fecha in seleccionados
        if abs((fecha - objetivo).days) <= 4
    ]

    if candidatos:
        return min(candidatos, key=lambda fecha: (abs((fecha - objetivo).days), fecha))

    if objetivo in seleccionados:
        return objetivo

    if seleccionados[0] <= objetivo <= seleccionados[-1]:
        return min(seleccionados, key=lambda fecha: abs((fecha - objetivo).days))

    return None


@programacion_bp.route("/programacion", methods=["GET"])
def programacion():
    datos = cargar_datos()
    hoy = date.today()

    festivos = {}
    for anio in {hoy.year, hoy.year + 1}:
        festivos.update(festivos_colombia(anio))

    _, mayores = analizar_maquinas(datos)

    mayores_relevantes = [
        item
        for item in mayores
        if item.get("estado") in ("vencido", "proximo")
    ]

    resumen_mayores = {
        "total": len(mayores_relevantes),
        "vencidos": sum(1 for item in mayores_relevantes if item.get("estado") == "vencido"),
        "proximos": sum(1 for item in mayores_relevantes if item.get("estado") == "proximo"),
    }

    return render_template(
        "programacion.html",
        fecha_hoy=hoy.isoformat(),
        festivos=festivos,
        mayores=mayores_relevantes,
        resumen_mayores=resumen_mayores,
    )


@programacion_bp.route("/programacion/generar", methods=["POST"])
def generar_programacion():
    datos = cargar_datos()
    payload = request.get_json(silent=True) or {}
    seleccionados = parsear_fechas(payload.get("fechas", []))

    if not seleccionados:
        return jsonify({"ok": False, "mensaje": "Seleccione al menos un día."}), 400

    if not misma_semana(seleccionados):
        return jsonify({"ok": False, "mensaje": "Solo puede seleccionar días de una misma semana."}), 400

    festivos = {}
    for anio in {fecha.year for fecha in seleccionados}:
        festivos.update(festivos_colombia(anio))

    seleccionados = [
        fecha
        for fecha in seleccionados
        if es_dia_habil(fecha, festivos)
    ]

    if not seleccionados:
        return jsonify({"ok": False, "mensaje": "No hay días hábiles seleccionados."}), 400

    rutinas, _ = analizar_maquinas(datos)
    filas = []

    for item in rutinas:
        dia = mejor_dia_para_rutina(item, seleccionados)
        if not dia:
            continue

        objetivo = convertir_fecha(item.get("fecha_objetivo"))
        if not objetivo:
            continue

        vencido = objetivo < seleccionados[0]
        proximo = abs((dia - objetivo).days) <= 4
        dentro_semana = seleccionados[0] <= objetivo <= seleccionados[-1]

        if not (vencido or proximo or dentro_semana):
            continue

        filas.append({
            "id": f"rutina-{item['maquina_id']}",
            "maquina_id": item["maquina_id"],
            "maquina": item["maquina"],
            "dia": dia.isoformat(),
            "tiempo": item["tiempo"],
            "frecuencia": item["frecuencia"].capitalize(),
            "tipo": "rutina",
            "prioridad": item["prioridad"],
        })

    filas.sort(key=lambda fila: (fila["dia"], fila["prioridad"], fila["maquina"]))

    return jsonify({
        "ok": True,
        "programacion": filas,
        "dias_disponibles": [fecha.isoformat() for fecha in seleccionados],
    })


@programacion_bp.route("/programacion/agregar-mayores", methods=["POST"])
def agregar_mayores():
    datos = cargar_datos()
    payload = request.get_json(silent=True) or {}
    seleccionados = parsear_fechas(payload.get("fechas", []))
    mayores_ids = {str(valor) for valor in payload.get("mayores", [])}

    if not seleccionados:
        return jsonify({"ok": False, "mensaje": "No hay días seleccionados."}), 400

    if not misma_semana(seleccionados):
        return jsonify({"ok": False, "mensaje": "Solo puede trabajar con una semana."}), 400

    _, mayores = analizar_maquinas(datos)
    filas = []

    for item in mayores:
        identificador = f"{item['maquina_id']}|{item['tipo']}|{item['frecuencia']}"
        if identificador not in mayores_ids:
            continue

        for dia in seleccionados:
            filas.append({
                "id": f"mayor-{item['maquina_id']}-{item['frecuencia']}-{dia.isoformat()}",
                "maquina_id": item["maquina_id"],
                "maquina": item["maquina"],
                "dia": dia.isoformat(),
                "tiempo": "Jornada",
                "frecuencia": item["frecuencia"],
                "tipo": "mayor",
                "prioridad": item["prioridad"],
            })

    return jsonify({"ok": True, "programacion": filas})
