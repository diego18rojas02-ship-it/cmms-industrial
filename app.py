from flask import Flask, render_template_string, request
from datetime import date
import json
import os
import calendar

app = Flask(__name__)

ARCHIVO = "datos_cmms.json"

MAQUINAS_BASE = [
    {"codigo": "CCM6", "nombre": "CCM6", "ultimo": 25000},
    {"codigo": "DICER200", "nombre": "Dicer 200", "ultimo": 18000},
    {"codigo": "DICER250", "nombre": "Dicer 250", "ultimo": 22000},
    {"codigo": "AUTOCLAVE", "nombre": "Autoclave", "ultimo": 12000},
    {"codigo": "COMPACTADORA", "nombre": "Compactadora", "ultimo": 15000},
    {"codigo": "PULVERIZADORA", "nombre": "Pulverizadora", "ultimo": 9000},
    {"codigo": "MINICARGADOR", "nombre": "Minicargador", "ultimo": 6000},
    {"codigo": "MONTACARGA", "nombre": "Montacarga", "ultimo": 8000}
]

FRECUENCIAS_CONTROL = {
    "rutina": 360,
    "punzones": 1800,
    "distribuidores": 40000,
    "bridas": 40000,
    "aceite": 40000
}

ALERTAS = {
    "rutina": 120,
    "punzones": 200,
    "distribuidores": 3000,
    "bridas": 3000,
    "aceite": 3000
}

FRECUENCIAS_MANTENIMIENTO = {
    "quincenal": {
        "nombre": "Quincenal",
        "categoria": "Rutina",
        "control": "rutina_horometro"
    },
    "mensual": {
        "nombre": "Mensual",
        "categoria": "Rutina",
        "control": "rutina_horometro"
    },
    "trimestral": {
        "nombre": "Trimestral",
        "categoria": "Mayor",
        "control": "punzones_horometro"
    },
    "semestral": {
        "nombre": "Semestral",
        "categoria": "Mayor",
        "control": "distribuidores_horometro"
    },
    "anual": {
        "nombre": "Anual",
        "categoria": "Mayor",
        "control": "bridas_horometro"
    }
}


def estructura_controles():
    return {
        m["codigo"]: {
            "rutina_horometro": 0,
            "punzones_horometro": 0,
            "distribuidores_horometro": 0,
            "bridas_horometro": 0,
            "aceite_horometro": 0
        }
        for m in MAQUINAS_BASE
    }


def cargar():
    if os.path.exists(ARCHIVO):
        try:
            with open(ARCHIVO, "r", encoding="utf-8") as f:
                datos = json.load(f)

            maquinas = datos.get("maquinas", MAQUINAS_BASE)
            controles = datos.get("controles", estructura_controles())
            registros = datos.get("registros", [])

            base = estructura_controles()

            for codigo in base:
                if codigo not in controles:
                    controles[codigo] = base[codigo]

                for campo in base[codigo]:
                    if campo not in controles[codigo]:
                        controles[codigo][campo] = 0

            return maquinas, controles, registros

        except Exception:
            return MAQUINAS_BASE, estructura_controles(), []

    return MAQUINAS_BASE, estructura_controles(), []


def guardar():
    with open(ARCHIVO, "w", encoding="utf-8") as f:
        json.dump(
            {
                "maquinas": maquinas,
                "controles": controles,
                "registros": registros
            },
            f,
            ensure_ascii=False,
            indent=4
        )


maquinas, controles, registros = cargar()


def maquina_por_codigo(codigo):
    for m in maquinas:
        if m["codigo"] == codigo:
            return m
    return None


def nombre_maquina(codigo):
    m = maquina_por_codigo(codigo)
    return m["nombre"] if m else codigo


def calcular_faltante(codigo, campo, frecuencia):
    maquina = maquina_por_codigo(codigo)

    if not maquina:
        return 0

    actual = int(maquina.get("ultimo", 0))
    ultimo = int(controles[codigo].get(campo, 0))
    consumido = actual - ultimo

    return frecuencia - consumido


def estado_visual(faltante, alerta, frecuencia):
    if faltante < -(frecuencia * 0.20):
        return "critical", "Crítico"

    if faltante < 0:
        return "expired", "Vencido"

    if faltante <= alerta:
        return "warning", "Por vencer"

    return "ok", "Normal"


def formato_dias(horas):
    return f"{abs(horas) / 24:.1f}d"


def formato_meses(horas):
    return f"{abs(horas) / 24 / 30:.1f}m"


def texto_tiempo(faltante, tipo):
    valor = formato_dias(faltante) if tipo == "dias" else formato_meses(faltante)

    if faltante < 0:
        return f"{valor} venc."

    return valor


def fila_alarma(icono, titulo, texto, clase, estado):
    return f"""
    <div class="alarm-row {clase}" title="{titulo} - {estado}">
        <div class="alarm-icon {clase}">
            {icono}
        </div>
        <div class="alarm-time">{texto}</div>
    </div>
    """


BASE_HTML = """
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>CMMS Industrial</title>

<style>
body{
    margin:0;
    font-family:Segoe UI, Arial, sans-serif;
    background:#0b0f19;
    color:white;
}

.sidebar{
    width:72px;
    height:100vh;
    position:fixed;
    left:0;
    top:0;
    background:#020617;
    border-right:1px solid #1f2937;
    display:flex;
    flex-direction:column;
    align-items:center;
    padding-top:20px;
}

.logo{
    font-size:24px;
    margin-bottom:25px;
}

.menu-item{
    width:48px;
    height:48px;
    margin-bottom:10px;
    border-radius:14px;
    display:flex;
    align-items:center;
    justify-content:center;
    color:white;
    text-decoration:none;
    position:relative;
    font-size:20px;
}

.menu-item:hover{
    background:#111827;
}

.menu-item span{
    display:none;
    position:absolute;
    left:65px;
    background:#111827;
    padding:10px 14px;
    border-radius:10px;
    white-space:nowrap;
    font-size:12px;
    z-index:99;
}

.menu-item:hover span{
    display:block;
}

.main{
    margin-left:72px;
    padding:18px;
}

.header,
.card,
.machine-card{
    background:#111827;
    border:1px solid #1f2937;
    border-radius:16px;
    padding:14px;
    margin-bottom:14px;
    box-shadow:0 4px 16px rgba(0,0,0,.35);
}

.header h1{
    margin:0;
    font-size:24px;
}

.header p{
    color:#94a3b8;
    font-size:13px;
}

.dashboard{
    display:grid;
    grid-template-columns:repeat(4,1fr);
    gap:14px;
}

.machine-title{
    margin-bottom:8px;
}

.machine-title h2{
    margin:0;
    font-size:15px;
}

.machine-title span{
    font-size:11px;
    color:#94a3b8;
}

.alarm-panel{
    display:grid;
    grid-template-columns:repeat(5,1fr);
    gap:6px;
}

.alarm-row{
    display:flex;
    flex-direction:column;
    align-items:center;
    justify-content:center;
    gap:4px;
    min-height:74px;
    border-radius:12px;
    border:1px solid #374151;
    background:#1f2937;
    transition:.2s;
    cursor:pointer;
}

.alarm-row:hover{
    transform:translateY(-2px);
}

.alarm-row.ok{
    border-bottom:5px solid #22c55e;
}

.alarm-row.warning{
    border-bottom:5px solid #facc15;
    background:#292411;
}

.alarm-row.expired{
    border-bottom:5px solid #ef4444;
    background:#2a1215;
}

.alarm-row.critical{
    border-bottom:5px solid #991b1b;
    background:#3b0d0d;
}

.alarm-icon{
    width:30px;
    height:30px;
    border-radius:10px;
    display:flex;
    align-items:center;
    justify-content:center;
    font-size:18px;
}

.alarm-icon.ok{
    background:#052e16;
}

.alarm-icon.warning{
    background:#422006;
    animation:pulseSoft 2s infinite;
}

.alarm-icon.expired{
    background:#450a0a;
    animation:blinkExpired 1.5s infinite;
}

.alarm-icon.critical{
    background:#7f1d1d;
    animation:blinkCritical .7s infinite;
}

.custom-icon{
    width:24px;
    height:24px;
    object-fit:contain;
}

.alarm-time{
    font-size:10px;
    font-weight:700;
    color:#e2e8f0;
}

button{
    background:#0284c7;
    border:none;
    color:white;
    padding:10px 14px;
    border-radius:10px;
    cursor:pointer;
    margin:3px;
}

button:hover{
    background:#0369a1;
}

input,
select{
    width:100%;
    padding:10px;
    border-radius:10px;
    border:1px solid #334155;
    background:#0f172a;
    color:white;
    margin-top:5px;
    margin-bottom:12px;
    box-sizing:border-box;
}

.grid{
    display:grid;
    grid-template-columns:repeat(4,1fr);
    gap:12px;
}

.indicador{
    font-size:24px;
    font-weight:bold;
}

table{
    width:100%;
    border-collapse:collapse;
}

th{
    background:#1e293b;
    color:#93c5fd;
}

td,
th{
    padding:9px;
    border-bottom:1px solid #334155;
    font-size:13px;
}

.form-grid{
    display:grid;
    grid-template-columns:repeat(2,1fr);
    gap:14px;
}

.alerta{
    background:#052e16;
    color:#86efac;
    padding:10px;
    border-radius:10px;
    margin-bottom:12px;
}

.error{
    background:#450a0a;
    color:#fecaca;
    padding:10px;
    border-radius:10px;
    margin-bottom:12px;
}

.machine-select-grid{
    display:grid;
    grid-template-columns:repeat(4,1fr);
    gap:12px;
}

.select-card{
    background:#1f2937;
    border:1px solid #334155;
    border-radius:14px;
    padding:14px;
    cursor:pointer;
    transition:.2s;
}

.select-card:hover{
    transform:translateY(-2px);
    border-color:#38bdf8;
}

.select-card.locked{
    opacity:.45;
    cursor:not-allowed;
    border-color:#475569;
}

.select-card h3{
    margin:0 0 8px 0;
    font-size:14px;
}

.select-card p{
    margin:0;
    color:#94a3b8;
    font-size:12px;
}

.checkbox-line{
    display:flex;
    align-items:center;
    gap:8px;
    margin:8px 0 12px 0;
}

.checkbox-line input{
    width:auto;
    margin:0;
}

.mantenimiento-box{
    display:none;
    background:#0f172a;
    border:1px solid #334155;
    border-radius:14px;
    padding:12px;
    margin-bottom:12px;
}

.mantenimiento-grid{
    display:grid;
    grid-template-columns:repeat(2,1fr);
    gap:10px;
    align-items:start;
}

.mantenimiento-grid .campo{
    display:flex;
    flex-direction:column;
}

.mantenimiento-grid label{
    font-size:12px;
    color:#cbd5e1;
    margin-bottom:4px;
}

.mantenimiento-grid input,
.mantenimiento-grid select{
    padding:7px 9px;
    font-size:12px;
    border-radius:8px;
    border:1px solid #334155;
    background:#020617;
    color:white;
    width:100%;
    box-sizing:border-box;
}

.check-mini{
    display:flex;
    align-items:center;
    gap:8px;
    height:100%;
    margin-top:18px;
    font-size:12px;
    color:#e2e8f0;
}

.check-mini input{
    width:auto;
    margin:0;
}

.calendar-box{
    max-width:430px;
    margin:auto;
    background:#020617;
    padding:18px;
    border-radius:18px;
    border:1px solid #334155;
}

.calendar-title{
    text-align:center;
    margin-bottom:14px;
    font-weight:bold;
    text-transform:capitalize;
}

.calendar-grid,
.calendar-head{
    display:grid;
    grid-template-columns:repeat(7,1fr);
    gap:7px;
    text-align:center;
}

.calendar-head div{
    color:#94a3b8;
    font-size:12px;
}

.day{
    height:38px;
    border-radius:50%;
    background:#111827;
    display:flex;
    align-items:center;
    justify-content:center;
    cursor:pointer;
}

.day:hover{
    background:#1d4ed8;
}

.day.selected{
    background:#fb923c;
}

.day.disabled{
    opacity:.3;
    cursor:not-allowed;
}

@keyframes pulseSoft{
    0%{box-shadow:0 0 0 0 rgba(250,204,21,.4);}
    70%{box-shadow:0 0 0 8px rgba(250,204,21,0);}
    100%{box-shadow:0 0 0 0 rgba(250,204,21,0);}
}

@keyframes blinkExpired{
    0%,100%{opacity:1;}
    50%{opacity:.4;}
}

@keyframes blinkCritical{
    0%,100%{opacity:1;transform:scale(1.08);}
    50%{opacity:.3;transform:scale(.95);}
}

@media(max-width:1200px){
    .dashboard,
    .machine-select-grid{
        grid-template-columns:repeat(2,1fr);
    }
}

@media(max-width:700px){
    .dashboard,
    .grid,
    .form-grid,
    .machine-select-grid,
    .mantenimiento-grid{
        grid-template-columns:1fr;
    }
}
</style>
</head>

<body>
<div class="sidebar">
    <div class="logo">⚙️</div>
    <a class="menu-item" href="/">🏠<span>Inicio</span></a>
    <a class="menu-item" href="/horometros">⏱<span>Horómetros</span></a>
    <a class="menu-item" href="/intervenciones">🔩<span>Intervenciones</span></a>
    <a class="menu-item" href="/programacion">📅<span>Programación</span></a>
</div>

<div class="main">
    {{contenido|safe}}
</div>
</body>
</html>
"""


@app.route("/")
def inicio():
    dashboard = ""

    for m in maquinas:
        codigo = m["codigo"]

        rutina = calcular_faltante(
            codigo,
            "rutina_horometro",
            FRECUENCIAS_CONTROL["rutina"]
        )
        clase_rutina, estado_rutina = estado_visual(
            rutina,
            ALERTAS["rutina"],
            FRECUENCIAS_CONTROL["rutina"]
        )

        punzones = calcular_faltante(
            codigo,
            "punzones_horometro",
            FRECUENCIAS_CONTROL["punzones"]
        )
        clase_punzones, estado_punzones = estado_visual(
            punzones,
            ALERTAS["punzones"],
            FRECUENCIAS_CONTROL["punzones"]
        )

        distribuidores = calcular_faltante(
            codigo,
            "distribuidores_horometro",
            FRECUENCIAS_CONTROL["distribuidores"]
        )
        clase_distribuidores, estado_distribuidores = estado_visual(
            distribuidores,
            ALERTAS["distribuidores"],
            FRECUENCIAS_CONTROL["distribuidores"]
        )

        bridas = calcular_faltante(
            codigo,
            "bridas_horometro",
            FRECUENCIAS_CONTROL["bridas"]
        )
        clase_bridas, estado_bridas = estado_visual(
            bridas,
            ALERTAS["bridas"],
            FRECUENCIAS_CONTROL["bridas"]
        )

        aceite = calcular_faltante(
            codigo,
            "aceite_horometro",
            FRECUENCIAS_CONTROL["aceite"]
        )
        clase_aceite, estado_aceite = estado_visual(
            aceite,
            ALERTAS["aceite"],
            FRECUENCIAS_CONTROL["aceite"]
        )

        icono_punzon = '<img src="/static/punzon.png" class="custom-icon">'

        dashboard += f"""
        <div class="machine-card">
            <div class="machine-title">
                <h2>{m["nombre"]}</h2>
                <span>Horómetro: {m["ultimo"]}</span>
            </div>

            <div class="alarm-panel">
                {fila_alarma("🔩", "Rutina / lubricación", texto_tiempo(rutina, "dias"), clase_rutina, estado_rutina)}
                {fila_alarma(icono_punzon, "Punzón formación tapas", texto_tiempo(punzones, "dias"), clase_punzones, estado_punzones)}
                {fila_alarma("💧", "Distribuidores", texto_tiempo(distribuidores, "meses"), clase_distribuidores, estado_distribuidores)}
                {fila_alarma("⚙️", "Bridas", texto_tiempo(bridas, "meses"), clase_bridas, estado_bridas)}
                {fila_alarma("🛢️", "Aceite", texto_tiempo(aceite, "meses"), clase_aceite, estado_aceite)}
            </div>
        </div>
        """

    contenido = f"""
    <div class="header">
        <h1>CMMS Industrial</h1>
        <p>Panel de alarmas de mantenimiento.</p>
    </div>

    <div class="grid">
        <div class="card">
            <h3>Máquinas</h3>
            <div class="indicador">{len(maquinas)}</div>
        </div>

        <div class="card">
            <h3>Rutina</h3>
            <div class="indicador">Días</div>
        </div>

        <div class="card">
            <h3>Mayores</h3>
            <div class="indicador">Meses</div>
        </div>

        <div class="card">
            <h3>Aceite</h3>
            <div class="indicador">Meses</div>
        </div>
    </div>

    <div class="card">
        <a href="/horometros"><button>Registrar horómetro</button></a>
        <a href="/intervenciones"><button>Editar intervención</button></a>
        <a href="/programacion"><button>Programación automática</button></a>
    </div>

    <div class="dashboard">
        {dashboard}
    </div>
    """

    return render_template_string(BASE_HTML, contenido=contenido)


@app.route("/horometros", methods=["GET", "POST"])
def horometros():
    mensaje = ""
    clase = "alerta"
    hoy = str(date.today())
    mostrar_formulario = False

    if request.method == "POST":
        codigo = request.form.get("maquina")
        horometro = int(request.form.get("horometro"))
        ejecuto = request.form.get("ejecuto_mantenimiento")
        frecuencia = request.form.get("frecuencia_mantenimiento")
        horometro_intervencion_txt = request.form.get("horometro_intervencion")
        fecha_ejecucion = request.form.get("fecha_ejecucion")
        cambio_aceite = request.form.get("cambio_aceite")

        maquina = maquina_por_codigo(codigo)

        ya_registrado = any(
            r.get("maquina") == codigo and r.get("fecha") == hoy
            for r in registros
        )

        if not maquina:
            mensaje = "Debe seleccionar una máquina."
            clase = "error"
            mostrar_formulario = True

        elif ya_registrado:
            mensaje = "Esta máquina ya tiene horómetro registrado hoy. Se habilita nuevamente mañana."
            clase = "error"

        elif horometro < int(maquina["ultimo"]):
            mensaje = "El horómetro actual no puede ser menor al último registrado."
            clase = "error"
            mostrar_formulario = True

        else:
            maquina["ultimo"] = horometro

            registro = {
                "id": str(len(registros) + 1),
                "fecha": hoy,
                "maquina": codigo,
                "horometro": horometro,
                "ejecuto_mantenimiento": "Sí" if ejecuto == "Si" else "No",
                "fecha_ejecucion_mantenimiento": "",
                "frecuencia": "",
                "categoria": "",
                "control_actualizado": "",
                "horometro_intervencion": "",
                "cambio_aceite": "No"
            }

            if ejecuto == "Si":
                if not frecuencia and cambio_aceite != "Si":
                    mensaje = "Debe seleccionar una frecuencia o marcar cambio de aceite."
                    clase = "error"
                    mostrar_formulario = True

                elif not horometro_intervencion_txt:
                    mensaje = "Debe ingresar el horómetro de intervención."
                    clase = "error"
                    mostrar_formulario = True

                else:
                    horometro_intervencion = int(horometro_intervencion_txt)

                    if horometro_intervencion > horometro:
                        mensaje = "El horómetro de intervención no puede ser mayor al horómetro actual."
                        clase = "error"
                        mostrar_formulario = True

                    else:
                        if frecuencia:
                            info = FRECUENCIAS_MANTENIMIENTO[frecuencia]
                            control = info["control"]

                            controles[codigo][control] = horometro_intervencion

                            registro["frecuencia"] = info["nombre"]
                            registro["categoria"] = info["categoria"]
                            registro["control_actualizado"] = control

                        if cambio_aceite == "Si":
                            controles[codigo]["aceite_horometro"] = horometro_intervencion
                            registro["cambio_aceite"] = "Sí"

                        registro["fecha_ejecucion_mantenimiento"] = fecha_ejecucion if fecha_ejecucion else hoy
                        registro["horometro_intervencion"] = horometro_intervencion

                        registros.append(registro)
                        guardar()

                        mensaje = "Horómetro y mantenimiento guardados correctamente."

            elif clase != "error":
                registros.append(registro)
                guardar()

                mensaje = "Horómetro guardado correctamente."

    tarjetas = ""

    for m in maquinas:
        codigo = m["codigo"]

        ya_registrado = any(
            r.get("maquina") == codigo and r.get("fecha") == hoy
            for r in registros
        )

        if ya_registrado:
            tarjetas += f"""
            <div class="select-card locked">
                <h3>🔒 {m["nombre"]}</h3>
                <p>Registrado hoy</p>
                <p>Último: {m["ultimo"]}</p>
            </div>
            """

        else:
            tarjetas += f"""
            <div class="select-card" onclick="seleccionarMaquina('{codigo}', '{m["nombre"]}', '{m["ultimo"]}')">
                <h3>{m["nombre"]}</h3>
                <p>Último horómetro: {m["ultimo"]}</p>
            </div>
            """

    filas = ""

    for r in registros:
        filas += f"""
        <tr>
            <td>{r.get("fecha", "")}</td>
            <td>{nombre_maquina(r.get("maquina", ""))}</td>
            <td>{r.get("horometro", "")}</td>
            <td>{r.get("ejecuto_mantenimiento", "No")}</td>
            <td>{r.get("fecha_ejecucion_mantenimiento", "")}</td>
            <td>{r.get("frecuencia", "")}</td>
            <td>{r.get("categoria", "")}</td>
            <td>{r.get("horometro_intervencion", "")}</td>
            <td>{r.get("cambio_aceite", "No")}</td>
        </tr>
        """

    display_form = "block" if mostrar_formulario or mensaje else "none"

    contenido = f"""
    <div class="header">
        <h1>Registro de horómetros</h1>
        <p>Seleccione una máquina para registrar horómetro actual y frecuencia ejecutada.</p>
    </div>

    <div class="card">
        <h2>Máquinas</h2>
        <div class="machine-select-grid">
            {tarjetas}
        </div>
    </div>

    <div class="card" id="formularioRegistro" style="display:{display_form};">
        {f'<div class="{clase}">{mensaje}</div>' if mensaje else ''}

        <form method="POST">
            <input type="hidden" name="maquina" id="maquinaSeleccionada">

            <h3 id="tituloMaquina">Seleccione una máquina</h3>

            <div class="form-grid">
                <div>
                    <label>Horómetro actual</label>
                    <input type="number" name="horometro" id="horometro" required>
                </div>
            </div>

            <label class="checkbox-line">
                <input type="checkbox" name="ejecuto_mantenimiento" value="Si" id="checkMant" onchange="mostrarMantenimiento()">
                ¿Se ejecutó mantenimiento?
            </label>

            <div class="mantenimiento-box" id="boxMantenimiento">
                <div class="mantenimiento-grid">

                    <div class="campo">
                        <label>Fecha ejecución</label>
                        <input
                            type="date"
                            name="fecha_ejecucion"
                            value="{hoy}"
                            onclick="this.showPicker && this.showPicker()"
                            onfocus="this.showPicker && this.showPicker()"
                            required
                        >
                    </div>

                    <div class="campo">
                        <label>Frecuencia ejecutada</label>
                        <select name="frecuencia_mantenimiento">
                            <option value="">Seleccione</option>
                            <option value="quincenal">Quincenal</option>
                            <option value="mensual">Mensual</option>
                            <option value="trimestral">Trimestral</option>
                            <option value="semestral">Semestral</option>
                            <option value="anual">Anual</option>
                        </select>
                    </div>

                    <div class="campo">
                        <label>Horómetro intervención</label>
                        <input type="number" name="horometro_intervencion">
                    </div>

                    <div class="campo">
                        <label class="check-mini">
                            <input type="checkbox" name="cambio_aceite" value="Si">
                            Cambio de aceite
                        </label>
                    </div>

                </div>
            </div>

            <button type="submit">Guardar horómetro</button>
        </form>
    </div>

    <div class="card">
        <h2>Histórico de horómetros</h2>

        <table>
            <tr>
                <th>Fecha registro</th>
                <th>Máquina</th>
                <th>Horómetro actual</th>
                <th>¿Mantenimiento?</th>
                <th>Fecha ejecución</th>
                <th>Frecuencia</th>
                <th>Categoría</th>
                <th>Horómetro intervención</th>
                <th>Aceite</th>
            </tr>

            {filas}
        </table>
    </div>

    <script>
        function seleccionarMaquina(codigo, nombre, ultimo){{
            document.getElementById("formularioRegistro").style.display = "block";
            document.getElementById("maquinaSeleccionada").value = codigo;

            document.getElementById("tituloMaquina").innerText =
                "Registrando: " + nombre + " | Último: " + ultimo;

            document.getElementById("horometro").focus();

            window.scrollTo({{
                top: 0,
                behavior: "smooth"
            }});
        }}

        function mostrarMantenimiento(){{
            let check = document.getElementById("checkMant");
            let box = document.getElementById("boxMantenimiento");

            box.style.display = check.checked ? "block" : "none";
        }}
    </script>
    """

    return render_template_string(BASE_HTML, contenido=contenido)


@app.route("/intervenciones", methods=["GET", "POST"])
def intervenciones():
    mensaje = ""
    clase = "alerta"

    if request.method == "POST":
        codigo = request.form.get("maquina")
        tipo = request.form.get("tipo")
        horometro = int(request.form.get("horometro"))
        maquina = maquina_por_codigo(codigo)

        if not maquina:
            mensaje = "Debe seleccionar una máquina."
            clase = "error"

        elif horometro > int(maquina["ultimo"]):
            mensaje = "El horómetro de intervención no puede ser mayor al horómetro actual."
            clase = "error"

        else:
            controles[codigo][tipo] = horometro
            guardar()

            mensaje = "Intervención editada correctamente."

    opciones = ""

    for m in maquinas:
        opciones += f"""
        <option value="{m["codigo"]}">
            {m["nombre"]} - Actual {m["ultimo"]}
        </option>
        """

    filas = ""

    for m in maquinas:
        codigo = m["codigo"]

        filas += f"""
        <tr>
            <td>{m["nombre"]}</td>
            <td>{controles[codigo].get("rutina_horometro", 0)}</td>
            <td>{controles[codigo].get("punzones_horometro", 0)}</td>
            <td>{controles[codigo].get("distribuidores_horometro", 0)}</td>
            <td>{controles[codigo].get("bridas_horometro", 0)}</td>
            <td>{controles[codigo].get("aceite_horometro", 0)}</td>
        </tr>
        """

    contenido = f"""
    <div class="header">
        <h1>Editar última intervención</h1>
        <p>Use esta opción solo para correcciones manuales.</p>
    </div>

    <div class="card">
        {f'<div class="{clase}">{mensaje}</div>' if mensaje else ''}

        <form method="POST">
            <div class="form-grid">
                <div>
                    <label>Máquina</label>
                    <select name="maquina" required>
                        <option value="">Seleccione</option>
                        {opciones}
                    </select>
                </div>

                <div>
                    <label>Control a editar</label>
                    <select name="tipo" required>
                        <option value="">Seleccione</option>
                        <option value="rutina_horometro">Rutina / Lubricación</option>
                        <option value="punzones_horometro">Trimestral - Punzón formación tapas</option>
                        <option value="distribuidores_horometro">Semestral - Distribuidores</option>
                        <option value="bridas_horometro">Anual - Bridas y actuadores</option>
                        <option value="aceite_horometro">Cambio de aceite</option>
                    </select>
                </div>
            </div>

            <label>Horómetro última intervención</label>
            <input type="number" name="horometro" required>

            <button type="submit">Editar intervención</button>
        </form>
    </div>

    <div class="card">
        <h2>Últimas intervenciones</h2>

        <table>
            <tr>
                <th>Máquina</th>
                <th>Rutina</th>
                <th>Punzón</th>
                <th>Distribuidores</th>
                <th>Bridas</th>
                <th>Aceite</th>
            </tr>

            {filas}
        </table>
    </div>
    """

    return render_template_string(BASE_HTML, contenido=contenido)


@app.route("/programacion")
def programacion():
    hoy = date.today()
    anio = hoy.year
    mes = hoy.month

    nombres_meses = [
        "",
        "enero",
        "febrero",
        "marzo",
        "abril",
        "mayo",
        "junio",
        "julio",
        "agosto",
        "septiembre",
        "octubre",
        "noviembre",
        "diciembre"
    ]

    cal = calendar.Calendar(firstweekday=6)
    semanas = cal.monthdatescalendar(anio, mes)

    dias_html = ""

    for semana in semanas:
        for dia in semana:
            clase = "day"

            if dia.month != mes or dia.weekday() >= 5:
                clase += " disabled"

            dias_html += f"""
            <div class="{clase}" onclick="seleccionarDia(this)">
                {dia.day}
            </div>
            """

    contenido = f"""
    <div class="header">
        <h1>Programación automática</h1>
        <p>Seleccione los días hábiles para programar mantenimientos.</p>
    </div>

    <div class="card">
        <div class="calendar-box">
            <div class="calendar-title">
                {nombres_meses[mes]} {anio}
            </div>

            <div class="calendar-head">
                <div>DO</div>
                <div>LU</div>
                <div>MA</div>
                <div>MI</div>
                <div>JU</div>
                <div>VI</div>
                <div>SA</div>
            </div>

            <div class="calendar-grid">
                {dias_html}
            </div>
        </div>
    </div>

    <div class="card">
        <button onclick="generarTabla()">
            Generar programación automática
        </button>
    </div>

    <div class="card">
        <h2>Resultado para Outlook</h2>
        <div id="resultado">
            Seleccione días y genere la programación.
        </div>
    </div>

    <script>
        function seleccionarDia(elemento){{
            if(elemento.classList.contains("disabled")) return;
            elemento.classList.toggle("selected");
        }}

        function generarTabla(){{
            let dias = document.querySelectorAll(".day.selected");
            let html = "<table><tr><th>Día</th><th>Actividad</th><th>Estado</th></tr>";

            dias.forEach(function(d){{
                html += "<tr><td>" + d.innerText + "</td><td>Mantenimiento programado</td><td>Por confirmar</td></tr>";
            }});

            html += "</table>";

            document.getElementById("resultado").innerHTML = html;
        }}
    </script>
    """

    return render_template_string(BASE_HTML, contenido=contenido)


if __name__ == "__main__":
    app.run(debug=False)