#! .\venv_ofi\Scripts\python.exe
from __future__ import annotations

import io
import os
import re
from datetime import date, datetime
from pathlib import Path
from urllib.parse import quote

import requests
from dotenv import load_dotenv
from flask import Flask, jsonify, render_template, request, send_file, send_from_directory, url_for
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

app = Flask(__name__)
app.config.update(
    JSON_AS_ASCII=False,
    MAX_CONTENT_LENGTH=16 * 1024,
    SEND_FILE_MAX_AGE_DEFAULT=0,
    TEMPLATES_AUTO_RELOAD=True,
)
app.jinja_env.auto_reload = True

XAFIRO_BASE_URL = os.getenv("XAFIRO_BASE_URL", "https://hotel.xafiro.net").rstrip("/")
NEO_AUTH_BASE_URL = os.getenv("NEO_AUTH_BASE_URL", "https://neo.castillodechancay.com/ms_auth").rstrip("/")
NEO_BOLETERIA_BASE_URL = os.getenv(
    "NEO_BOLETERIA_BASE_URL",
    "https://neo.castillodechancay.com/ms_boleteria",
).rstrip("/")
REQUEST_TIMEOUT = (20, 180)


class XafiroError(RuntimeError):
    """Error controlado al comunicarse con Xafiro."""


class NeoError(RuntimeError):
    """Error controlado al comunicarse con Neo."""


def static_asset(filename: str) -> str:
    asset_path = BASE_DIR / "static" / filename
    version = int(asset_path.stat().st_mtime) if asset_path.exists() else 0
    return url_for("static", filename=filename, v=version)


def animation_asset(filename: str) -> str:
    asset_path = BASE_DIR / "animaciones" / filename
    version = int(asset_path.stat().st_mtime) if asset_path.exists() else 0
    return url_for("animation_file", filename=filename, v=version)


app.jinja_env.globals["static_asset"] = static_asset
app.jinja_env.globals["animation_asset"] = animation_asset


def xafiro_request_error(exc: requests.RequestException, action: str) -> XafiroError:
    detail = str(exc)

    if "WinError 10013" in detail:
        return XafiroError(
            f"La PC bloqueó la conexión de Python hacia Xafiro al {action}. "
            "Revisa firewall, antivirus o permisos de red para Python."
        )
    if isinstance(exc, requests.Timeout):
        return XafiroError(
            f"Xafiro tardó demasiado en responder al {action}. "
            "Intenta nuevamente o usa un rango de fechas más corto."
        )
    if isinstance(exc, requests.ConnectionError):
        return XafiroError(
            f"No se pudo conectar con Xafiro al {action}. "
            "Verifica internet, VPN/proxy o que https://hotel.xafiro.net esté disponible."
        )
    if isinstance(exc, requests.HTTPError) and exc.response is not None:
        return XafiroError(
            f"Xafiro respondió con error HTTP {exc.response.status_code} al {action}."
        )

    return XafiroError(f"No se pudo completar la comunicación con Xafiro al {action}.")


def build_xafiro_session() -> requests.Session:
    retry_strategy = Retry(
        total=2,
        connect=2,
        read=2,
        backoff_factor=0.7,
        status_forcelist=(502, 503, 504),
        allowed_methods=frozenset({"GET", "POST"}),
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry_strategy)

    session = requests.Session()
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update(
        {
            "User-Agent": "Contabilidad-Chancay/0.1",
            "Accept": "application/vnd.ms-excel, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/octet-stream",
        }
    )
    return session


def neo_request_error(exc: requests.RequestException, action: str) -> NeoError:
    detail = str(exc)

    if "WinError 10013" in detail:
        return NeoError(
            f"La PC bloqueó la conexión de Python hacia Neo al {action}. "
            "Revisa firewall, antivirus o permisos de red para Python."
        )
    if isinstance(exc, requests.Timeout):
        return NeoError(
            f"Neo tardó demasiado en responder al {action}. "
            "Intenta nuevamente o usa un rango de fechas más corto."
        )
    if isinstance(exc, requests.ConnectionError):
        return NeoError(
            f"No se pudo conectar con Neo al {action}. "
            "Verifica internet, VPN/proxy o que https://neo.castillodechancay.com esté disponible."
        )
    if isinstance(exc, requests.HTTPError) and exc.response is not None:
        return NeoError(f"Neo respondió con error HTTP {exc.response.status_code} al {action}.")

    return NeoError(f"No se pudo completar la comunicación con Neo al {action}.")


def build_neo_session() -> requests.Session:
    retry_strategy = Retry(
        total=2,
        connect=2,
        read=2,
        backoff_factor=0.7,
        status_forcelist=(502, 503, 504),
        allowed_methods=frozenset({"GET", "POST"}),
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry_strategy)

    session = requests.Session()
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update(
        {
            "User-Agent": "Contabilidad-Chancay/0.1",
            "Accept": "application/json, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/octet-stream",
        }
    )
    return session


def parse_iso_date(raw_value: object, field_name: str) -> date:
    if not isinstance(raw_value, str):
        raise ValueError(f"{field_name} es obligatorio.")

    try:
        return datetime.strptime(raw_value, "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError(f"{field_name} debe tener el formato AAAA-MM-DD.") from exc


def xafiro_credentials() -> dict[str, str]:
    mapping = {
        "empresa": os.getenv("XAFIRO_EMPRESA", "").strip(),
        "usuario": os.getenv("XAFIRO_USUARIO", "").strip(),
        "pass": os.getenv("XAFIRO_PASSWORD", "").strip(),
    }
    if not all(mapping.values()):
        raise XafiroError("Faltan credenciales de Xafiro en el archivo .env.")
    return mapping


def neo_credentials() -> dict[str, str]:
    mapping = {
        "usuario": (
            os.getenv("NEO_USUARIO", "").strip()
            or os.getenv("Neo_usuario", "").strip()
        ),
        "password": (
            os.getenv("NEO_PASSWORD", "").strip()
            or os.getenv("Neo_password", "").strip()
        ),
    }
    if not all(mapping.values()):
        raise NeoError("Faltan credenciales de Neo en el archivo .env.")
    return mapping


def login_to_xafiro(session: requests.Session) -> None:
    credentials = xafiro_credentials()

    try:
        response = session.post(
            f"{XAFIRO_BASE_URL}/",
            data={**credentials, "enviar": "Ingresar"},
            timeout=REQUEST_TIMEOUT,
        )
        response.raise_for_status()
    except requests.RequestException as exc:
        raise xafiro_request_error(exc, "iniciar sesión") from exc

    # El formulario vuelve a mostrarse cuando Xafiro rechaza las credenciales.
    if re.search(r'name=[\"\']empresa[\"\']', response.text, re.IGNORECASE):
        raise XafiroError("Xafiro rechazó las credenciales configuradas.")


def download_xafiro_report(start_date: date, end_date: date) -> tuple[bytes, str, str]:
    with build_xafiro_session() as session:
        login_to_xafiro(session)

        export_url = (
            f"{XAFIRO_BASE_URL}/ticket/export_ticket_report/"
            f"{quote(start_date.isoformat())}/{quote(end_date.isoformat())}"
        )

        try:
            response = session.get(export_url, timeout=REQUEST_TIMEOUT)
            response.raise_for_status()
        except requests.RequestException as exc:
            raise xafiro_request_error(exc, "descargar el reporte") from exc

        content_type = response.headers.get("Content-Type", "").split(";", 1)[0].strip()
        if "text/html" in content_type.lower():
            raise XafiroError("La sesión de Xafiro terminó antes de descargar el reporte.")
        if not response.content:
            raise XafiroError("Xafiro devolvió un reporte vacío.")

        # Xafiro declara actualmente `application/vnd.ms-excel`, aunque el
        # archivo que entrega es un libro XLSX (contenedor ZIP).
        is_xlsx = response.content.startswith(b"PK\x03\x04")
        extension = ".xlsx" if is_xlsx else ".xls"
        if is_xlsx:
            content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        filename = f"facturacion_xafiro_{start_date.isoformat()}_{end_date.isoformat()}{extension}"
        return response.content, filename, content_type or "application/vnd.ms-excel"


def login_to_neo(session: requests.Session) -> None:
    credentials = neo_credentials()

    try:
        response = session.post(
            f"{NEO_AUTH_BASE_URL}/usuario/login",
            json=credentials,
            timeout=REQUEST_TIMEOUT,
        )
        response.raise_for_status()
    except requests.RequestException as exc:
        raise neo_request_error(exc, "iniciar sesión") from exc

    try:
        payload = response.json()
    except ValueError as exc:
        raise NeoError("Neo devolvió una respuesta de login inválida.") from exc

    if not payload.get("estado"):
        raise NeoError(payload.get("message") or "Neo rechazó las credenciales configuradas.")

    token = ((payload.get("data") or {}).get("access_token") or {}).get("token")
    if not token:
        raise NeoError("Neo no devolvió token de acceso para descargar el reporte.")

    session.headers.update({"Authorization": f"Bearer {token}"})


def fetch_neo_report_rows(session: requests.Session, start_date: date, end_date: date) -> list[dict[str, object]]:
    report_url = (
        f"{NEO_BOLETERIA_BASE_URL}/reporte/reporte-detallado-ticket"
        f"?fecha_inicio={quote(start_date.isoformat())}&fecha_fin={quote(end_date.isoformat())}"
    )

    try:
        response = session.get(report_url, timeout=REQUEST_TIMEOUT)
        response.raise_for_status()
    except requests.RequestException as exc:
        raise neo_request_error(exc, "descargar el reporte") from exc

    try:
        payload = response.json()
    except ValueError as exc:
        raise NeoError("Neo devolvió un reporte con formato inválido.") from exc

    if not payload.get("estado"):
        raise NeoError(payload.get("message") or "Neo no pudo generar el reporte solicitado.")

    rows = payload.get("data")
    if not isinstance(rows, list):
        raise NeoError("Neo devolvió el reporte sin filas válidas.")

    return [row for row in rows if isinstance(row, dict)]


def excel_value(value: object) -> object:
    if isinstance(value, (dict, list)):
        import json

        return json.dumps(value, ensure_ascii=False)
    return value


def build_neo_workbook(rows: list[dict[str, object]]) -> bytes:
    columns = [
        "fecha_operacion",
        "hora_operacion",
        "codigo_venta",
        "numero_operacion",
        "id_venta",
        "id_externo",
        "producto",
        "categoria",
        "cantidad",
        "precio",
        "tipo_pago",
        "tipo_caja",
        "tipo_colita",
        "estado_comprobante",
        "comprobantes",
        "detalles_ticket",
        "gratis",
        "validado",
    ]
    extra_columns = sorted({key for row in rows for key in row.keys()} - set(columns))
    columns.extend(extra_columns)

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Facturación Neo"
    worksheet.freeze_panes = "A2"

    header_fill = PatternFill("solid", fgColor="173F33")
    header_font = Font(color="FFFFFF", bold=True)
    worksheet.append([column.replace("_", " ").title() for column in columns])
    worksheet.auto_filter.ref = f"A1:{worksheet.cell(row=1, column=len(columns)).coordinate}"
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font

    for row in rows:
        worksheet.append([excel_value(row.get(column)) for column in columns])

    for column_cells in worksheet.columns:
        header = str(column_cells[0].value or "")
        max_length = max(len(str(cell.value or "")) for cell in column_cells[:80])
        worksheet.column_dimensions[column_cells[0].column_letter].width = min(
            max(max_length, len(header)) + 2,
            42,
        )

    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


def register_neo_export(session: requests.Session, filename: str) -> None:
    try:
        response = session.post(
            f"{NEO_BOLETERIA_BASE_URL}/export-log/trabajar-export-log",
            json={"filename": filename, "origen": "reporte_detallado_tickets"},
            timeout=REQUEST_TIMEOUT,
        )
        response.raise_for_status()
    except requests.RequestException as exc:
        raise neo_request_error(exc, "registrar la exportación") from exc


def download_neo_report(start_date: date, end_date: date) -> tuple[bytes, str, str]:
    filename = f"facturacion_neo_{start_date.isoformat()}_{end_date.isoformat()}.xlsx"

    with build_neo_session() as session:
        login_to_neo(session)
        rows = fetch_neo_report_rows(session, start_date, end_date)
        content = build_neo_workbook(rows)
        register_neo_export(session, filename)

    return (
        content,
        filename,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/animaciones/<path:filename>")
def animation_file(filename: str):
    return send_from_directory(BASE_DIR / "animaciones", filename, max_age=0)


@app.get("/api/health")
def health():
    credentials_ready = all(
        os.getenv(key, "").strip()
        for key in ("XAFIRO_EMPRESA", "XAFIRO_USUARIO", "XAFIRO_PASSWORD")
    )
    neo_ready = bool(
        (os.getenv("NEO_USUARIO", "").strip() or os.getenv("Neo_usuario", "").strip())
        and (os.getenv("NEO_PASSWORD", "").strip() or os.getenv("Neo_password", "").strip())
    )
    return jsonify(
        {
            "status": "ok",
            "xafiro_configurado": credentials_ready,
            "neo_configurado": neo_ready,
        }
    )


@app.post("/api/xafiro/export")
def export_xafiro():
    payload = request.get_json(silent=True) or {}

    try:
        start_date = parse_iso_date(payload.get("fecha_inicio"), "La fecha inicial")
        end_date = parse_iso_date(payload.get("fecha_fin"), "La fecha final")
        if start_date > end_date:
            raise ValueError("La fecha inicial no puede ser posterior a la fecha final.")
        if (end_date - start_date).days > 366:
            raise ValueError("El rango máximo permitido es de 367 días.")

        content, filename, mimetype = download_xafiro_report(start_date, end_date)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except XafiroError as exc:
        return jsonify({"error": str(exc)}), 502

    return send_file(
        io.BytesIO(content),
        mimetype=mimetype,
        as_attachment=True,
        download_name=filename,
        max_age=0,
    )


@app.post("/api/neo/export")
def export_neo():
    payload = request.get_json(silent=True) or {}

    try:
        start_date = parse_iso_date(payload.get("fecha_inicio"), "La fecha inicial")
        end_date = parse_iso_date(payload.get("fecha_fin"), "La fecha final")
        if start_date > end_date:
            raise ValueError("La fecha inicial no puede ser posterior a la fecha final.")
        if (end_date - start_date).days > 366:
            raise ValueError("El rango máximo permitido es de 367 días.")

        content, filename, mimetype = download_neo_report(start_date, end_date)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except NeoError as exc:
        return jsonify({"error": str(exc)}), 502

    return send_file(
        io.BytesIO(content),
        mimetype=mimetype,
        as_attachment=True,
        download_name=filename,
        max_age=0,
    )


@app.errorhandler(413)
def payload_too_large(_error):
    return jsonify({"error": "La solicitud es demasiado grande."}), 413


if __name__ == "__main__":
    host = os.getenv("APP_HOST", "127.0.0.1")
    port = int(os.getenv("APP_PORT", "5050"))
    debug = os.getenv("APP_DEBUG", "false").lower() == "true"
    app.run(host=host, port=port, debug=debug)
