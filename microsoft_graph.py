import html
import os
from typing import Any

import msal
import requests
from flask import Blueprint, jsonify, redirect, request, session, url_for


microsoft_graph_bp = Blueprint(
    "microsoft_graph",
    __name__,
    url_prefix="/microsoft",
)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPES = ["User.Read", "Mail.ReadWrite"]


def _configuracion() -> tuple[str, str, str, str]:
    client_id = os.environ.get("MICROSOFT_CLIENT_ID", "").strip()
    client_secret = os.environ.get("MICROSOFT_CLIENT_SECRET", "").strip()
    tenant_id = os.environ.get("MICROSOFT_TENANT_ID", "").strip()
    redirect_uri = os.environ.get("MICROSOFT_REDIRECT_URI", "").strip()

    faltantes = []
    if not client_id:
        faltantes.append("MICROSOFT_CLIENT_ID")
    if not client_secret:
        faltantes.append("MICROSOFT_CLIENT_SECRET")
    if not tenant_id:
        faltantes.append("MICROSOFT_TENANT_ID")
    if not redirect_uri:
        faltantes.append("MICROSOFT_REDIRECT_URI")

    if faltantes:
        raise RuntimeError("Falta configurar: " + ", ".join(faltantes))

    authority = "https://login.microsoftonline.com/" + tenant_id
    return client_id, client_secret, authority, redirect_uri


def _cargar_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    serializado = session.get("microsoft_token_cache")

    if serializado:
        cache.deserialize(serializado)

    return cache


def _guardar_cache(cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        session["microsoft_token_cache"] = cache.serialize()


def _crear_app_msal(
    cache: msal.SerializableTokenCache,
) -> msal.ConfidentialClientApplication:
    client_id, client_secret, authority, _ = _configuracion()

    return msal.ConfidentialClientApplication(
        client_id=client_id,
        authority=authority,
        client_credential=client_secret,
        token_cache=cache,
    )


def _token_silencioso() -> str | None:
    cache = _cargar_cache()
    app = _crear_app_msal(cache)
    cuentas = app.get_accounts()

    if not cuentas:
        _guardar_cache(cache)
        return None

    resultado = app.acquire_token_silent(
        SCOPES,
        account=cuentas[0],
    )

    _guardar_cache(cache)

    if not resultado:
        return None

    return resultado.get("access_token")


def _crear_flujo_login() -> str:
    cache = _cargar_cache()
    app = _crear_app_msal(cache)
    _, _, _, redirect_uri = _configuracion()

    flujo = app.initiate_auth_code_flow(
        scopes=SCOPES,
        redirect_uri=redirect_uri,
    )

    if "auth_uri" not in flujo:
        raise RuntimeError(
            "Microsoft no devolvió una URL de autenticación."
        )

    session["microsoft_auth_flow"] = flujo
    _guardar_cache(cache)

    return flujo["auth_uri"]


def _escapar(valor: Any) -> str:
    return html.escape(str(valor if valor is not None else ""))


def _tabla_html(payload: dict[str, Any]) -> str:
    filas = payload.get("programacion", [])
    semana = _escapar(
        payload.get("semana", "Semana seleccionada")
    )

    filas_html = []

    for fila in filas:
        es_mayor = str(fila.get("tipo", "")).lower() == "mayor"
        fondo = "#eef6ff" if es_mayor else "#ffffff"

        etiqueta = (
            '<span style="display:inline-block;margin-left:6px;'
            'padding:2px 7px;border-radius:999px;'
            'background:#dbeafe;color:#1d4ed8;'
            'font-size:10px;font-weight:700;">Mayor</span>'
            if es_mayor
            else ""
        )

        filas_html.append(
            f"""
            <tr style="background:{fondo};">
                <td style="padding:11px 13px;border:1px solid #d8e2ec;font-family:Arial,sans-serif;font-size:13px;color:#172033;">
                    <strong>{_escapar(fila.get("maquina", ""))}</strong>
                    {etiqueta}
                </td>
                <td style="padding:11px 13px;border:1px solid #d8e2ec;font-family:Arial,sans-serif;font-size:13px;color:#172033;">
                    {_escapar(fila.get("dia_texto", fila.get("dia", "")))}
                </td>
                <td style="padding:11px 13px;border:1px solid #d8e2ec;font-family:Arial,sans-serif;font-size:13px;color:#172033;text-align:center;">
                    {_escapar(fila.get("tiempo", ""))}
                </td>
                <td style="padding:11px 13px;border:1px solid #d8e2ec;font-family:Arial,sans-serif;font-size:13px;color:#172033;">
                    {_escapar(fila.get("frecuencia", ""))}
                </td>
            </tr>
            """
        )

    cuerpo_filas = "".join(filas_html)

    return f"""
    <div style="margin:0;padding:0;font-family:Arial,sans-serif;color:#172033;">

        <table role="presentation" cellspacing="0" cellpadding="0" border="0"
               style="width:100%;max-width:780px;margin:0;border-collapse:collapse;">
            <tr>
                <td style="padding:0 0 16px 0;">
                    <div style="margin-bottom:5px;font-size:11px;font-weight:700;letter-spacing:1.2px;color:#2563eb;">
                        CMMS INDUSTRIAL
                    </div>

                    <div style="margin-bottom:4px;font-size:22px;line-height:1.25;font-weight:700;color:#0f172a;">
                        Programación semanal de mantenimiento
                    </div>

                    <div style="font-size:12px;color:#64748b;">
                        {semana}
                    </div>
                </td>
            </tr>
        </table>

        <table cellspacing="0" cellpadding="0" border="0"
               style="width:100%;max-width:780px;border-collapse:collapse;border:1px solid #d8e2ec;">
            <thead>
                <tr style="background:#0f2942;">
                    <th style="padding:11px 13px;border:1px solid #28445e;font-family:Arial,sans-serif;font-size:11px;color:#ffffff;text-align:left;text-transform:uppercase;letter-spacing:.5px;">Máquina</th>
                    <th style="padding:11px 13px;border:1px solid #28445e;font-family:Arial,sans-serif;font-size:11px;color:#ffffff;text-align:left;text-transform:uppercase;letter-spacing:.5px;">Día</th>
                    <th style="padding:11px 13px;border:1px solid #28445e;font-family:Arial,sans-serif;font-size:11px;color:#ffffff;text-align:center;text-transform:uppercase;letter-spacing:.5px;">Tiempo</th>
                    <th style="padding:11px 13px;border:1px solid #28445e;font-family:Arial,sans-serif;font-size:11px;color:#ffffff;text-align:left;text-transform:uppercase;letter-spacing:.5px;">Frecuencia</th>
                </tr>
            </thead>
            <tbody>
                {cuerpo_filas}
            </tbody>
        </table>

        <table role="presentation" cellspacing="0" cellpadding="0" border="0"
               style="width:100%;max-width:780px;margin-top:14px;border-collapse:collapse;">
            <tr>
                <td style="font-family:Arial,sans-serif;font-size:11px;line-height:1.45;color:#64748b;">
                    Programación generada desde CMMS Industrial.
                </td>
            </tr>
        </table>

    </div>
    """


def _crear_borrador_graph(
    access_token: str,
    payload: dict[str, Any],
) -> dict[str, Any]:
    mensaje = {
        "subject": "Programación semanal de mantenimiento",
        "body": {
            "contentType": "HTML",
            "content": _tabla_html(payload),
        },
        "toRecipients": [],
    }

    respuesta = requests.post(
        GRAPH_BASE + "/me/messages",
        headers={
            "Authorization": "Bearer " + access_token,
            "Content-Type": "application/json",
        },
        json=mensaje,
        timeout=30,
    )

    if respuesta.status_code != 201:
        try:
            detalle = respuesta.json()
        except ValueError:
            detalle = respuesta.text

        raise RuntimeError(
            "Microsoft Graph no pudo crear el borrador. "
            f"HTTP {respuesta.status_code}: {detalle}"
        )

    borrador = respuesta.json()
    web_link = borrador.get("webLink")

    if not web_link:
        message_id = borrador.get("id")

        if message_id:
            consulta = requests.get(
                GRAPH_BASE
                + "/me/messages/"
                + message_id
                + "?$select=id,webLink",
                headers={
                    "Authorization": "Bearer " + access_token,
                },
                timeout=30,
            )

            if consulta.ok:
                web_link = consulta.json().get("webLink")

    return {
        "id": borrador.get("id"),
        "web_link": web_link,
    }


@microsoft_graph_bp.route("/crear-borrador", methods=["POST"])
def crear_borrador():
    payload = request.get_json(silent=True) or {}
    programacion = payload.get("programacion")

    if not isinstance(programacion, list) or not programacion:
        return jsonify({
            "ok": False,
            "mensaje": "No hay programación para enviar al correo.",
        }), 400

    session["correo_pendiente"] = payload

    try:
        token = _token_silencioso()

        if not token:
            login_url = _crear_flujo_login()

            return jsonify({
                "ok": True,
                "login_required": True,
                "login_url": login_url,
            })

        borrador = _crear_borrador_graph(
            token,
            payload,
        )

        session.pop("correo_pendiente", None)

        return jsonify({
            "ok": True,
            "login_required": False,
            "web_url": borrador.get("web_link"),
            "message_id": borrador.get("id"),
        })

    except Exception as error:
        return jsonify({
            "ok": False,
            "mensaje": str(error),
        }), 500


@microsoft_graph_bp.route("/login", methods=["GET"])
def login():
    try:
        return redirect(_crear_flujo_login())
    except Exception as error:
        return (
            "No fue posible iniciar sesión con Microsoft: "
            + html.escape(str(error)),
            500,
        )


@microsoft_graph_bp.route("/callback", methods=["GET"])
def callback():
    cache = _cargar_cache()
    app = _crear_app_msal(cache)
    flujo = session.get("microsoft_auth_flow")

    if not flujo:
        return (
            "La sesión de autenticación expiró. "
            "Vuelva a Programación y pulse Correo nuevamente.",
            400,
        )

    try:
        resultado = app.acquire_token_by_auth_code_flow(
            flujo,
            request.args,
        )
    except ValueError:
        return (
            "La respuesta de Microsoft no coincide con "
            "la sesión de autenticación iniciada.",
            400,
        )

    _guardar_cache(cache)
    session.pop("microsoft_auth_flow", None)

    if "error" in resultado:
        descripcion = resultado.get(
            "error_description",
            resultado.get("error", "Error de autenticación"),
        )

        return (
            "Microsoft no autorizó el acceso: "
            + html.escape(str(descripcion)),
            400,
        )

    access_token = resultado.get("access_token")

    if not access_token:
        return (
            "Microsoft no devolvió un token de acceso.",
            400,
        )

    payload = session.pop("correo_pendiente", None)

    if not payload:
        return redirect(
            url_for("programacion.programacion")
        )

    try:
        borrador = _crear_borrador_graph(
            access_token,
            payload,
        )
    except Exception as error:
        return (
            "Se inició sesión, pero no fue posible crear "
            "el borrador: "
            + html.escape(str(error)),
            500,
        )

    web_link = borrador.get("web_link")

    if web_link:
        return redirect(web_link)

    return redirect(
        url_for("programacion.programacion")
    )


@microsoft_graph_bp.route("/cerrar-sesion", methods=["GET"])
def cerrar_sesion():
    session.pop("microsoft_token_cache", None)
    session.pop("microsoft_auth_flow", None)
    session.pop("correo_pendiente", None)

    return redirect(
        url_for("programacion.programacion")
    )
