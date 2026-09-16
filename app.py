import os
import csv
import re
import ssl
import html
import time
import smtplib
import tempfile
import threading
import mimetypes
import subprocess
import webbrowser
import unicodedata
from datetime import date
from pathlib import Path
from email.message import EmailMessage
from string import Template

import pandas as pd
from PIL import Image, ImageChops

import customtkinter as ctk
from tkinter import filedialog, messagebox

import smtp_keys


# =========================
# CONFIGURACIÓN POR DEFECTO
# =========================

DEFAULT_COURSE_NAME = "Ingrese curso"
# course1 = NOMBRE CORTO del curso en Moodle (la carga de usuarios no busca por número ID)
DEFAULT_MOODLE_COURSE_FIELD = ""
DEFAULT_MOODLE_TYPE1 = 1
DEFAULT_PROFILE_FIELD_NAME = "profile_field_rut"

DEFAULT_PASSWORD_YEAR = date.today().year
# placeholders: {username}, {year}, {rut}, {email}
DEFAULT_PASSWORD_PATTERN = "{username}{year}"

DEFAULT_AULA_URL = "https://aulavirtual.perfeccionatec.cl/"

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465
SENDER = "perfeccionatec@gmail.com"

THROTTLE_SECONDS = 1.0
MAX_RETRIES = 3

USERNAME_NORMALIZE_ACCENTS = True

ASSETS_DIR = Path(__file__).resolve().parent / "assets"
LOGO_PATH = ASSETS_DIR / "perfeccionatec.png"
LOGO_CID = "logo_perfeccionatec"
MANUAL_PATH = ASSETS_DIR / "Manual de ingreso Perfeccionatec.pdf"
MANUAL_ATTACHMENT_NAME = "Manual_Ingreso_Aula.pdf"

SUBJECT_TEMPLATE = Template("Tus credenciales — Aula $nombre_curso")

PREHEADER_TEMPLATE = Template(
    "Tus credenciales del Aula Virtual ya están listas. Ingresa y parte con el curso de $nombre_curso."
)

# Sin <svg>: Gmail y Outlook los eliminan. El logo va como imagen embebida (cid).
HTML_TEMPLATE = Template(r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Bienvenido al Aula Virtual PerfeccionaTEC</title>
</head>
<body style="margin:0; padding:0; background-color:#eaeff4; font-family:'Plus Jakarta Sans', Arial, Calibri, sans-serif;">

<!-- Preheader (oculto, mejora la vista previa en el inbox) -->
<div style="display:none; max-height:0; overflow:hidden; opacity:0;">
$preheader
</div>

<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#eaeff4; padding:32px 16px;">
  <tr>
    <td align="center">
      <table role="presentation" width="600" cellpadding="0" cellspacing="0" style="max-width:600px; width:100%; background-color:#ffffff; border:1px solid #dcdfe0;">

        <!-- Header -->
        <tr>
          <td style="background-color:#ffffff; padding:24px 40px;">
            <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
              <tr>
                <td valign="middle">
                  <img src="$logo_src" width="140" alt="PerfeccionaTEC Capacitación" style="display:block; width:140px; height:auto; border:0; outline:none; text-decoration:none;">
                </td>
                <td valign="middle" align="right" style="font-family:'IBM Plex Mono', monospace; font-size:11px; letter-spacing:0.22em; text-transform:uppercase; color:#093995; font-weight:700;">
                  Aula Virtual
                </td>
              </tr>
            </table>
          </td>
        </tr>

        <!-- Franja de curso -->
        <tr>
          <td style="background-color:#093995; padding:12px 40px; font-family:'IBM Plex Mono', monospace; font-size:12px; color:#c6d8f9;">
            Curso asignado &nbsp;·&nbsp; $nombre_curso
          </td>
        </tr>

        <!-- Cuerpo -->
        <tr>
          <td style="padding:40px 40px 8px 40px;">
            <div style="font-weight:800; font-size:26px; letter-spacing:-0.02em; color:#161e2e; line-height:1.25;">
              ¡Bienvenido, $primer_nombre!
            </div>
            <p style="font-size:15px; line-height:1.6; color:#3d4451; margin:16px 0 0 0;">
              Ya quedaste inscrito en el curso <strong style="color:#161e2e;">$nombre_curso</strong>. Aquí tienes tus credenciales para entrar al Aula Virtual y partir con el módulo 1.
            </p>
            <p style="font-size:15px; line-height:1.6; color:#3d4451; margin:12px 0 0 0;">
              Te adjuntamos el <strong style="color:#161e2e;">Manual de Ingreso al Aula</strong>, con el paso a paso para entrar y moverte por la plataforma sin problemas.
            </p>
          </td>
        </tr>

        <!-- Credenciales -->
        <tr>
          <td style="padding:24px 40px 8px 40px;">
            <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#eaeff4; border:1px solid #dcdfe0;">
              <tr>
                <td style="padding:22px 24px;">
                  <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
                    <tr>
                      <td style="font-family:'IBM Plex Mono', monospace; font-size:11px; color:#7c8186; text-transform:uppercase; letter-spacing:0.08em; padding-bottom:4px;">Usuario</td>
                    </tr>
                    <tr>
                      <td style="font-family:'IBM Plex Mono', monospace; font-size:16px; color:#093995; font-weight:700; padding-bottom:14px;">$usuario</td>
                    </tr>
                    <tr>
                      <td style="font-family:'IBM Plex Mono', monospace; font-size:11px; color:#7c8186; text-transform:uppercase; letter-spacing:0.08em; padding-bottom:4px;">Contraseña</td>
                    </tr>
                    <tr>
                      <td style="font-family:'IBM Plex Mono', monospace; font-size:16px; color:#093995; font-weight:700;">$contrasena</td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>

        <!-- Adjunto -->
        <tr>
          <td style="padding:14px 40px 0 40px;">
            <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-left:3px solid #093995;">
              <tr>
                <td style="padding:4px 0 4px 14px; font-size:13px; color:#3d4451; line-height:1.5;">
                  <strong style="color:#161e2e;">Manual de ingreso:</strong> revisa el archivo adjunto en este correo.
                </td>
              </tr>
            </table>
          </td>
        </tr>

        <!-- Recomendación de seguridad -->
        <tr>
          <td style="padding:14px 40px 0 40px;">
            <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-left:3px solid #c26300;">
              <tr>
                <td style="padding:4px 0 4px 14px; font-size:13px; color:#3d4451; line-height:1.5;">
                  <strong style="color:#161e2e;">Recomendación:</strong> cambia tu contraseña apenas inicies sesión, por una que solo tú conozcas.
                </td>
              </tr>
            </table>
          </td>
        </tr>

        <!-- Botón CTA -->
        <tr>
          <td align="center" style="padding:32px 40px 8px 40px;">
            <table role="presentation" cellpadding="0" cellspacing="0">
              <tr>
                <td style="background-color:#093995;">
                  <a href="$aula_url" style="display:inline-block; padding:14px 36px; font-size:15px; font-weight:700; color:#ffffff; text-decoration:none; font-family:'Plus Jakarta Sans', Arial, sans-serif;">
                    Acceder al Aula
                  </a>
                </td>
              </tr>
            </table>
            <p style="font-size:12px; color:#7c8186; margin:14px 0 0 0;">
              O copia este enlace en tu navegador:<br>
              <a href="$aula_url" style="color:#093995;">$aula_url</a>
            </p>
          </td>
        </tr>

        <!-- Primeros pasos -->
        <tr>
          <td style="padding:32px 40px 32px 40px;">
            <div style="font-weight:700; font-size:15px; color:#161e2e; padding-bottom:14px; border-bottom:1px solid #dcdfe0;">
              Tus primeros pasos
            </div>
            <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:14px;">
              <tr>
                <td style="padding:8px 0; vertical-align:top; width:28px; font-family:'IBM Plex Mono', monospace; font-size:13px; color:#c26300; font-weight:700;">01</td>
                <td style="padding:8px 0; font-size:14px; color:#3d4451; line-height:1.5;">Inicia sesión y cambia tu contraseña.</td>
              </tr>
              <tr>
                <td style="padding:8px 0; vertical-align:top; width:28px; font-family:'IBM Plex Mono', monospace; font-size:13px; color:#c26300; font-weight:700;">02</td>
                <td style="padding:8px 0; font-size:14px; color:#3d4451; line-height:1.5;">Revisa el programa del curso y las fechas de las clases.</td>
              </tr>
              <tr>
                <td style="padding:8px 0; vertical-align:top; width:28px; font-family:'IBM Plex Mono', monospace; font-size:13px; color:#c26300; font-weight:700;">03</td>
                <td style="padding:8px 0; font-size:14px; color:#3d4451; line-height:1.5;">Entra al módulo 1 y completa la primera actividad.</td>
              </tr>
            </table>
          </td>
        </tr>

        <!-- Footer -->
        <tr>
          <td style="padding:24px 40px 32px 40px; border-top:1px solid #dcdfe0;">
            <p style="font-size:12px; color:#7c8186; margin:0 0 6px 0;">
              ¿Dudas o problemas de acceso? Responde este correo y te ayudamos.
            </p>
            <p style="font-size:12px; color:#7c8186; margin:0;">
              © PerfeccionaTEC — Este mensaje fue enviado automáticamente.
            </p>
            <p style="font-family:'IBM Plex Mono', monospace; font-size:11px; font-weight:700; color:#c26300; margin:14px 0 0 0;">
              PerfeccionaTEC · Perfecciónate donde estés.
            </p>
          </td>
        </tr>

      </table>
    </td>
  </tr>
</table>

</body>
</html>
""")

PLAIN_TEMPLATE = Template("""
¡Bienvenido, $primer_nombre!

Ya quedaste inscrito en el curso $nombre_curso. Aquí tienes tus credenciales para entrar al Aula Virtual PerfeccionaTEC:

Usuario: $usuario
Contraseña: $contrasena

Acceder al Aula: $aula_url

Te adjuntamos el Manual de Ingreso al Aula, con el paso a paso para entrar y moverte por la plataforma. Revisa el archivo adjunto en este correo.

Recomendación: cambia tu contraseña apenas inicies sesión, por una que solo tú conozcas.

Tus primeros pasos:
01. Inicia sesión y cambia tu contraseña.
02. Revisa el programa del curso y las fechas de las clases.
03. Entra al módulo 1 y completa la primera actividad.

¿Dudas o problemas de acceso? Responde este correo y te ayudamos.

PerfeccionaTEC · Perfecciónate donde estés.
""".strip())


def render_email(user, course_name, aula_url, logo_src):
    """Devuelve (asunto, texto plano, html) para un usuario."""
    nombre = user["nombre"]
    values = {
        "primer_nombre": nombre.split()[0] if nombre else "",
        "usuario": user["usuario"],
        "contrasena": user["contrasena"],
        "aula_url": aula_url,
        "nombre_curso": course_name,
    }
    subject = SUBJECT_TEMPLATE.substitute(nombre_curso=course_name)
    plain = PLAIN_TEMPLATE.substitute(values)
    html_body = HTML_TEMPLATE.substitute(
        {k: html.escape(str(v)) for k, v in values.items()},
        preheader=html.escape(PREHEADER_TEMPLATE.substitute(nombre_curso=course_name)),
        logo_src=html.escape(logo_src),
    )
    return subject, plain, html_body


PREVIEW_BROWSERS = [
    Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "Microsoft/Edge/Application/msedge.exe",
    Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Microsoft/Edge/Application/msedge.exe",
    Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Google/Chrome/Application/chrome.exe",
    Path(os.environ.get("LOCALAPPDATA", "")) / "Google/Chrome/Application/chrome.exe",
]
PREVIEW_BG_COLOR = (0xEA, 0xEF, 0xF4)  # fondo del <body> del correo
PREVIEW_DEMO_USER = {
    "email": "demo@correo.cl",
    "nombre": "Nombre Apellido",
    "usuario": "usuario.demo",
    "contrasena": "Cambiar123*",
}


def render_html_screenshot(html_body, width=680, height=2400):
    """Renderiza el HTML con Edge/Chrome headless y devuelve la imagen recortada (PIL)."""
    browser = next((b for b in PREVIEW_BROWSERS if b.is_file()), None)
    if browser is None:
        raise FileNotFoundError("No se encontró Microsoft Edge ni Google Chrome para renderizar la vista previa.")

    work_dir = Path(tempfile.gettempdir()) / "perfeccionatec_preview"
    work_dir.mkdir(exist_ok=True)
    html_path = work_dir / "correo.html"
    png_path = work_dir / "correo.png"
    html_path.write_text(html_body, encoding="utf-8")
    png_path.unlink(missing_ok=True)

    subprocess.run(
        [
            str(browser), "--headless=new", "--disable-gpu", "--hide-scrollbars",
            "--no-first-run", "--allow-file-access-from-files",
            f"--user-data-dir={work_dir / 'profile'}",
            f"--window-size={width},{height}",
            f"--screenshot={png_path}",
            html_path.as_uri(),
        ],
        capture_output=True,
        timeout=60,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    if not png_path.is_file():
        raise RuntimeError("El navegador no generó la captura de la vista previa.")

    with Image.open(png_path) as shot:
        img = shot.convert("RGB")
    # Recorta el fondo sobrante bajo el correo
    bbox = ImageChops.difference(img, Image.new("RGB", img.size, PREVIEW_BG_COLOR)).getbbox()
    if bbox:
        img = img.crop((0, 0, img.width, min(img.height, bbox[3] + 32)))
    return img


def load_email_assets():
    """Lee el logo y el manual PDF. Lanza FileNotFoundError si falta alguno."""
    missing = [str(p) for p in (LOGO_PATH, MANUAL_PATH) if not p.is_file()]
    if missing:
        raise FileNotFoundError("Faltan archivos en assets/:\n" + "\n".join(missing))
    return {
        "logo_bytes": LOGO_PATH.read_bytes(),
        "logo_type": mimetypes.guess_type(LOGO_PATH.name)[0] or "image/png",
        "manual_bytes": MANUAL_PATH.read_bytes(),
    }

def normalize_simple(s: str) -> str:
    if not isinstance(s, str):
        s = str(s)
    nfkd = unicodedata.normalize("NFKD", s)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def normalize_username(u: str) -> str:
    u = (
        u.lower()
        .replace(" ", "")
        .replace(".", "")
        .replace(",", "")
        .replace("'", "")
        .replace('"', "")
    )
    if USERNAME_NORMALIZE_ACCENTS:
        u = normalize_simple(u)
    return u


def select_single_email(email_raw: str) -> str:
    if not isinstance(email_raw, str):
        email_raw = str(email_raw)
    txt = email_raw.replace("\n", " ").strip()
    tokens = re.split(r"[,\s;]+", txt)
    for t in tokens:
        if "@" in t:
            return t
    return txt


def build_username_from_row(row) -> str:
    nombres = str(row.get("nombres", "")).strip()
    apellidos = str(row.get("apellidos", "")).strip()

    first_name = nombres.split()[0] if nombres else ""
    ap_tokens = apellidos.split()
    first_surname = ap_tokens[0] if len(ap_tokens) >= 1 else ""
    second_surname_initials = ap_tokens[1][:2] if len(ap_tokens) >= 2 else ""

    raw = f"{first_name}{first_surname}{second_surname_initials}"
    return normalize_username(raw)


def build_password(pattern: str, year: int, username: str, rut: str, email: str) -> str:
    pwd = pattern.format(
        username=username,
        year=year,
        rut=rut,
        email=email,
    )
    pwd = normalize_simple(pwd)
    return pwd


HEADER_SEARCH_ROWS = 30

# Columna interna -> reconoce el encabezado ya normalizado (minúsculas, sin tildes ni espacios extra)
EXCEL_COLUMN_MATCHERS = {
    "rut": lambda h: h.startswith("rut"),  # "RUT", "Rut (con punto y con guión)"
    "nombres": lambda h: h in ("nombres", "nombre"),
    "apellidos": lambda h: h in ("apellidos", "apellido"),
    "email": lambda h: h.startswith("correo") or h in ("email", "e-mail", "mail"),
}


def normalize_header(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return " ".join(normalize_simple(str(value)).lower().split())


def find_header_row(raw: pd.DataFrame):
    """Devuelve (índice de fila, {columna interna: índice de columna}) o (None, {})."""
    for row_idx in range(min(HEADER_SEARCH_ROWS, len(raw))):
        headers = [normalize_header(v) for v in raw.iloc[row_idx]]
        found = {}
        for key, matches in EXCEL_COLUMN_MATCHERS.items():
            col = next((i for i, h in enumerate(headers) if h and matches(h)), None)
            if col is not None:
                found[key] = col
        if len(found) == len(EXCEL_COLUMN_MATCHERS):
            return row_idx, found
    return None, {}


def read_participants_excel(excel_path: str) -> pd.DataFrame:
    """
    Lee la primera hoja y busca la fila de encabezados en las primeras filas.
    Devuelve las filas de datos con columnas: rut, nombres, apellidos, email.
    """
    raw = pd.read_excel(excel_path, sheet_name=0, header=None)

    row_idx, found = find_header_row(raw)
    if row_idx is not None:
        data = raw.iloc[row_idx + 1:, list(found.values())].copy()
        data.columns = list(found.keys())
        return data

    raise ValueError(
        "No se encontró la fila de encabezados en el Excel.\n"
        f"Se buscó en las primeras {HEADER_SEARCH_ROWS} filas de la primera hoja una fila con las columnas:\n"
        "RUT, Nombres, Apellidos y Correo (sin importar mayúsculas, tildes ni espacios)."
    )


def cell_text(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return " ".join(str(value).split())  # sin saltos de línea ni espacios dobles


def format_table(df: pd.DataFrame, max_col_width: int = 40) -> str:
    """Tabla de texto con columnas alineadas para las vistas previas."""
    headers = [cell_text(c) for c in df.columns]
    rows = [[cell_text(v) for v in row] for row in df.itertuples(index=False)]
    widths = [
        min(max_col_width, max([len(h)] + [len(r[i]) for r in rows]))
        for i, h in enumerate(headers)
    ]
    fit = lambda text, w: (text if len(text) <= w else text[:w - 1] + "…").ljust(w)
    lines = [" │ ".join(fit(h, w) for h, w in zip(headers, widths))]
    lines.append("─┼─".join("─" * w for w in widths))
    lines += [" │ ".join(fit(v, w) for v, w in zip(r, widths)) for r in rows]
    return "\n".join(line.rstrip() for line in lines) + "\n"


def excel_preview_table(excel_path: str):
    """
    Tabla para la pestaña Excel: desde la fila de encabezados detectada, sin filas ni columnas vacías.
    Devuelve (DataFrame, texto informativo).
    """
    raw = pd.read_excel(excel_path, sheet_name=0, header=None)
    raw = raw.dropna(how="all").dropna(axis=1, how="all")
    row_idx, found = find_header_row(raw.reset_index(drop=True))
    if row_idx is None:
        raw.columns = [f"Col {i + 1}" for i in range(raw.shape[1])]
        return raw, "No se detectaron las columnas RUT, Nombres, Apellidos y Correo"

    raw = raw.reset_index(drop=True)
    table = raw.iloc[row_idx + 1:].copy()
    table.columns = [cell_text(h) or f"Col {i + 1}" for i, h in enumerate(raw.iloc[row_idx])]
    rut_col, nombres_col = table.columns[found["rut"]], table.columns[found["nombres"]]
    valid = table.iloc[:, found["rut"]].map(cell_text).ne("") & table.iloc[:, found["nombres"]].map(cell_text).ne("")
    info = f"Encabezados detectados · participantes con {rut_col} y {nombres_col}: {int(valid.sum())}"
    return table, info


def normalize_excel_to_moodle_csv(
    excel_path: str,
    csv_output_path: str,
    course_field: str,
    type1_value: int,
    profile_field_name: str,
    password_pattern: str,
    password_year: int,
):
    clean_df = read_participants_excel(excel_path)

    has_value = lambda col: clean_df[col].notna() & (clean_df[col].astype(str).str.strip() != "")
    participants = clean_df[has_value("rut") & has_value("nombres")].copy()

    moodle = pd.DataFrame()
    moodle["firstname"] = (
        participants["nombres"].astype(str).str.strip().str.title().str.split().str[0]
    )
    moodle["lastname"] = participants["apellidos"].astype(str).str.strip().str.title()
    moodle["email"] = participants["email"].apply(select_single_email)
    moodle[profile_field_name] = participants["rut"].astype(str).str.strip()

    moodle["username"] = participants.apply(build_username_from_row, axis=1)
    moodle["password"] = [
        build_password(password_pattern, password_year, u, r, e)
        for u, r, e in zip(
            moodle["username"],
            moodle[profile_field_name],
            moodle["email"],
        )
    ]
    moodle["type1"] = type1_value
    moodle["course1"] = course_field

    moodle = moodle[
        ["username", "password", "firstname", "lastname", "email", profile_field_name, "type1", "course1"]
    ]

    # utf-8-sig: con BOM para que Excel muestre bien las tildes (Moodle ignora el BOM)
    moodle.to_csv(csv_output_path, index=False, encoding="utf-8-sig")
    return moodle


def load_users_from_csv(path: str):
    users = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames:
            return users

        fieldnames = [fn.lower() for fn in reader.fieldnames]

        format_old = "email" in fieldnames and "usuario" in fieldnames
        format_moodle = "email" in fieldnames and "username" in fieldnames and "password" in fieldnames

        for row in reader:
            if format_old:
                email = (row.get("email") or "").strip()
                nombre = (row.get("nombre") or "").strip()
                usuario = (row.get("usuario") or "").strip()
                contrasena = (row.get("contrasena") or "").strip()
            elif format_moodle:
                email = (row.get("email") or "").strip()
                firstname = (row.get("firstname") or "").strip()
                lastname = (row.get("lastname") or "").strip()
                nombre = (firstname + " " + lastname).strip() or email.split("@")[0].title()
                usuario = (row.get("username") or "").strip()
                contrasena = (row.get("password") or "").strip()
            else:
                email = (row.get("email") or "").strip()
                nombre = (row.get("nombre") or "").strip() or email.split("@")[0].title()
                usuario = (row.get("usuario") or row.get("username") or email.split("@")[0]).strip()
                contrasena = (row.get("contrasena") or row.get("password") or "").strip()

            if email:
                users.append({
                    "email": email,
                    "nombre": nombre,
                    "usuario": usuario,
                    "contrasena": contrasena,
                })
    return users


def build_message(sender, recipient, subject, plain, html_body, assets):
    """multipart/mixed[ alternative[ texto, related[ html, logo ] ], manual.pdf ]"""
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.set_content(plain)
    msg.add_alternative(html_body, subtype="html")

    logo_main, logo_sub = assets["logo_type"].split("/")
    msg.get_payload()[-1].add_related(
        assets["logo_bytes"],
        maintype=logo_main,
        subtype=logo_sub,
        cid=f"<{LOGO_CID}>",
        disposition="inline",
        filename=LOGO_PATH.name,
    )
    msg.add_attachment(
        assets["manual_bytes"],
        maintype="application",
        subtype="pdf",
        filename=MANUAL_ATTACHMENT_NAME,
    )
    return msg


def send_all(sender, smtp_password, users, course_name, aula_url, assets, log_func,
             on_login=None, subject_prefix=""):
    """
    Envío con contador y 'cuenta regresiva':
    - [1/50] Enviando a...
    - [ENVIADO 1/50] ... (quedan 49)
    """
    total = len(users)
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context) as smtp:
        smtp.login(sender, smtp_password)
        if on_login:
            on_login()
        for idx, u in enumerate(users, start=1):
            restantes = total - idx
            log_func(f"[{idx}/{total}] Enviando a {u['email']}...")

            subject, plain, html_body = render_email(u, course_name, aula_url, f"cid:{LOGO_CID}")
            subject = subject_prefix + subject
            msg = build_message(sender, u["email"], subject, plain, html_body, assets)

            sent = False
            for attempt in range(1, MAX_RETRIES + 1):
                try:
                    smtp.send_message(msg)
                    log_func(f"[ENVIADO {idx}/{total}] {u['email']} (quedan {restantes})")
                    sent = True
                    break
                except Exception as e:
                    log_func(f"[ERROR] intento {attempt} con {u['email']}: {e}")
                    time.sleep(2 * attempt)

            if not sent:
                log_func(f"[FALLO {idx}/{total}] No se pudo enviar a {u['email']} después de {MAX_RETRIES} intentos.")

            time.sleep(THROTTLE_SECONDS)


# APP CON CUSTOMTKINTER


class MoodleApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.title("Moodle CSV + Envío de Credenciales")
        self.geometry("1400x900")
        self.minsize(1150, 600)
        # Maximizada para que se vea todo; con after() porque CTk pisa el estado inicial
        self.after(50, lambda: self.state("zoomed") if os.name == "nt" else None)

        self.var_course_name = ctk.StringVar(value=DEFAULT_COURSE_NAME)
        self.var_course1 = ctk.StringVar(value=DEFAULT_MOODLE_COURSE_FIELD)
        self.var_type1 = ctk.StringVar(value=str(DEFAULT_MOODLE_TYPE1))
        self.var_profile_field = ctk.StringVar(value=DEFAULT_PROFILE_FIELD_NAME)

        self.var_password_year = ctk.StringVar(value=str(DEFAULT_PASSWORD_YEAR))
        self.var_password_pattern = ctk.StringVar(value=DEFAULT_PASSWORD_PATTERN)

        self.var_aula_url = ctk.StringVar(value=DEFAULT_AULA_URL)

        self.var_excel_path = ctk.StringVar(value="")
        self.var_csv_output_path = ctk.StringVar(value="")
        self.var_csv_mail_path = ctk.StringVar(value="")

        self.df_excel_raw = None
        self.df_moodle = None
        self.df_csv_mail = None
        self.users_mail = []

        self.sending_thread = None
        self.sending = False

        self.info_excel = None
        self.info_moodle = None
        self.info_csv = None

        self.build_ui()
        self.after(300, self.update_email_preview_first_user)

    # ---------- UI ----------
    def build_ui(self):
        main_frame = ctk.CTkFrame(self, corner_radius=0)
        main_frame.pack(fill="both", expand=True)

        # uniform: el ancho de cada panel no cambia según el contenido de la pestaña abierta
        main_frame.grid_columnconfigure(0, weight=2, minsize=420, uniform="panels")
        main_frame.grid_columnconfigure(1, weight=3, uniform="panels")
        main_frame.grid_rowconfigure(2, weight=1)

        header = ctk.CTkFrame(main_frame)
        header.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(5, 0), padx=5)
        header.grid_columnconfigure(0, weight=1)

        title_label = ctk.CTkLabel(
            header,
            text="Panel de carga Moodle + Envío de credenciales",
            font=("Segoe UI Semibold", 18),
        )
        title_label.grid(row=0, column=0, sticky="w", padx=10, pady=5)

        subtitle = ctk.CTkLabel(
            header,
            text="Normaliza planillas, genera CSV para Moodle y envía credenciales, todo desde un mismo lugar.",
            font=("Segoe UI", 12),
            text_color=("gray80", "gray80")
        )
        subtitle.grid(row=1, column=0, sticky="w", padx=10, pady=(0, 8))

        smtp_bar = ctk.CTkFrame(main_frame)
        smtp_bar.grid(row=1, column=0, columnspan=2, sticky="ew", pady=5, padx=5)
        self.build_smtp_bar(smtp_bar)

        left = ctk.CTkFrame(main_frame)
        left.grid(row=2, column=0, sticky="nsew", padx=(5, 2), pady=(0, 5))
        self.build_left_panel(left)

        right = ctk.CTkFrame(main_frame)
        right.grid(row=2, column=1, sticky="nsew", padx=(2, 5), pady=(0, 5))
        right.grid_rowconfigure(1, weight=1)
        right.grid_columnconfigure(0, weight=1)
        self.build_right_panel(right)

    def build_smtp_bar(self, parent: ctk.CTkFrame):
        parent.grid_columnconfigure(2, weight=1)

        ctk.CTkLabel(parent, text="Gmail", font=("Segoe UI Semibold", 14)).grid(
            row=0, column=0, sticky="w", padx=(10, 12), pady=(6, 0)
        )
        ctk.CTkLabel(
            parent,
            text=f"Remitente: {SENDER}  ·  {SMTP_SERVER}:{SMTP_PORT}",
            font=("Segoe UI", 11),
            text_color=("gray70", "gray70"),
        ).grid(row=0, column=1, columnspan=4, sticky="w", pady=(6, 0))

        ctk.CTkLabel(parent, text="Clave de aplicación:").grid(row=1, column=0, sticky="w", padx=(10, 4), pady=8)

        key_row = ctk.CTkFrame(parent, fg_color="transparent")
        key_row.grid(row=1, column=1, sticky="w", pady=8)
        self.smtp_key_entry = ctk.CTkEntry(key_row, width=200, show="•", placeholder_text="")
        self.smtp_key_entry.pack(side="left")
        self.smtp_key_toggle = ctk.CTkButton(
            key_row, text="Mostrar", width=70, command=self.toggle_smtp_key_visibility
        )
        self.smtp_key_toggle.pack(side="left", padx=(4, 0))
        self.smtp_key_menu = ctk.CTkOptionMenu(
            key_row, width=260, values=["Sin claves encontradas"], command=self.select_saved_smtp_key
        )
        self.smtp_key_menu.pack(side="left", padx=(8, 0))
        ctk.CTkButton(
            key_row, text="Olvidar", width=70, command=self.forget_selected_smtp_key,
            fg_color="gray35", hover_color="gray25",
        ).pack(side="left", padx=(4, 0))

        test_row = ctk.CTkFrame(parent, fg_color="transparent")
        test_row.grid(row=1, column=3, sticky="e", padx=10, pady=8)
        ctk.CTkLabel(test_row, text="Correo de prueba:").pack(side="left", padx=(0, 4))
        self.test_email_entry = ctk.CTkEntry(test_row, width=240, placeholder_text="correo@ejemplo.cl")
        self.test_email_entry.pack(side="left")
        ctk.CTkButton(
            test_row, text="Enviar prueba", width=110, command=self.action_send_test_email,
            fg_color="#c26300", hover_color="#9a4f00",
        ).pack(side="left", padx=(4, 0))

        self.smtp_key_options = {}
        self.load_saved_smtp_keys(select_first=True)

    def build_left_panel(self, parent: ctk.CTkFrame):
        parent.grid_rowconfigure(0, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        # Formulario con scroll; las acciones quedan fijas abajo y siempre visibles
        form = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        form.grid(row=0, column=0, sticky="nsew")
        form.grid_columnconfigure(0, weight=1)

        course_frame = ctk.CTkFrame(form)
        course_frame.grid(row=0, column=0, sticky="ew", padx=4, pady=6)
        course_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(course_frame, text="Curso y Moodle", font=("Segoe UI Semibold", 14)).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 6)
        )

        ctk.CTkLabel(course_frame, text="Nombre del curso (para el correo):").grid(
            row=1, column=0, sticky="w"
        )
        ctk.CTkEntry(course_frame, textvariable=self.var_course_name).grid(
            row=1, column=1, sticky="ew", padx=4, pady=2
        )

        ctk.CTkLabel(course_frame, text="Nombre corto del curso (course1):").grid(row=2, column=0, sticky="w")
        course1_row = ctk.CTkFrame(course_frame, fg_color="transparent")
        course1_row.grid(row=2, column=1, sticky="ew", padx=4, pady=2)
        ctk.CTkEntry(course1_row, textvariable=self.var_course1, width=160).pack(side="left")
        ctk.CTkLabel(
            course1_row,
            text="Moodle › Configuración del curso › «Nombre corto» (no el número ID)",
            font=("Segoe UI", 9),
            text_color=("gray78", "gray78"),
        ).pack(side="left", padx=4)

        ctk.CTkLabel(course_frame, text="type1:").grid(row=3, column=0, sticky="w")
        type_row = ctk.CTkFrame(course_frame, fg_color="transparent")
        type_row.grid(row=3, column=1, sticky="ew", padx=4, pady=2)
        ctk.CTkEntry(type_row, textvariable=self.var_type1, width=60).pack(side="left")
        ctk.CTkLabel(
            type_row,
            text="1 = crear/actualizar usuario y matricular",
            font=("Segoe UI", 9),
            text_color=("gray78", "gray78"),
        ).pack(side="left", padx=4)

        info_type = ctk.CTkLabel(
            course_frame,
            text="En Moodle puedes usar type2, type3, etc. como pares con course2, course3 para otros cursos.",
            font=("Segoe UI", 9),
            text_color=("gray70", "gray70"),
            wraplength=360,
            justify="left",
        )
        info_type.grid(row=4, column=0, columnspan=2, sticky="w", pady=(4, 0))

        ctk.CTkLabel(course_frame, text="Campo RUT en Moodle:").grid(row=5, column=0, sticky="w", pady=(6, 0))
        ctk.CTkEntry(course_frame, textvariable=self.var_profile_field).grid(
            row=5, column=1, sticky="ew", padx=4, pady=(6, 4)
        )

        pwd_frame = ctk.CTkFrame(form)
        pwd_frame.grid(row=1, column=0, sticky="ew", padx=4, pady=6)
        pwd_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(pwd_frame, text="Contraseñas", font=("Segoe UI Semibold", 14)).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 6)
        )

        ctk.CTkLabel(pwd_frame, text="Año:").grid(row=1, column=0, sticky="w")
        ctk.CTkEntry(pwd_frame, textvariable=self.var_password_year, width=80).grid(
            row=1, column=1, sticky="w", padx=4, pady=2
        )

        ctk.CTkLabel(pwd_frame, text="Patrón de contraseña:").grid(row=2, column=0, sticky="w")
        ctk.CTkEntry(pwd_frame, textvariable=self.var_password_pattern).grid(
            row=2, column=1, sticky="ew", padx=4, pady=2
        )

        ctk.CTkLabel(
            pwd_frame,
            text="Placeholders disponibles: {username}, {year}, {rut}, {email}",
            font=("Segoe UI", 9),
            text_color=("gray70", "gray70"),
            wraplength=360,
            justify="left",
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(4, 0))

        aula_frame = ctk.CTkFrame(form)
        aula_frame.grid(row=2, column=0, sticky="ew", padx=4, pady=6)
        aula_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(aula_frame, text="Aula Virtual", font=("Segoe UI Semibold", 14)).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 6)
        )

        ctk.CTkLabel(aula_frame, text="URL Aula:").grid(row=1, column=0, sticky="w")
        ctk.CTkEntry(aula_frame, textvariable=self.var_aula_url).grid(
            row=1, column=1, sticky="ew", padx=4, pady=2
        )

        files_frame = ctk.CTkFrame(form)
        files_frame.grid(row=3, column=0, sticky="ew", padx=4, pady=6)
        files_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(files_frame, text="Archivos", font=("Segoe UI Semibold", 14)).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 6)
        )

        ctk.CTkLabel(files_frame, text="Excel de participantes:").grid(row=1, column=0, sticky="w")
        ctk.CTkEntry(files_frame, textvariable=self.var_excel_path).grid(
            row=1, column=1, sticky="ew", padx=4, pady=2
        )
        ctk.CTkButton(files_frame, text="Buscar", command=self.browse_excel, width=80).grid(
            row=1, column=2, padx=2, pady=2
        )

        ctk.CTkLabel(files_frame, text="CSV Moodle (salida):").grid(row=2, column=0, sticky="w")
        ctk.CTkEntry(files_frame, textvariable=self.var_csv_output_path).grid(
            row=2, column=1, sticky="ew", padx=4, pady=2
        )
        ctk.CTkButton(files_frame, text="Guardar como", command=self.browse_csv_output, width=80).grid(
            row=2, column=2, padx=2, pady=2
        )

        ctk.CTkLabel(files_frame, text="CSV para envío de correos:").grid(row=3, column=0, sticky="w")
        ctk.CTkEntry(files_frame, textvariable=self.var_csv_mail_path).grid(
            row=3, column=1, sticky="ew", padx=4, pady=2
        )
        ctk.CTkButton(files_frame, text="Buscar", command=self.browse_csv_mail, width=80).grid(
            row=3, column=2, padx=2, pady=2
        )

        csv_btns = ctk.CTkFrame(files_frame, fg_color="transparent")
        csv_btns.grid(row=4, column=1, columnspan=2, sticky="w", padx=4, pady=(4, 8))
        ctk.CTkButton(
            csv_btns,
            text="Usar CSV Moodle como fuente de correos",
            command=self.use_moodle_csv_for_mail,
        ).pack(side="left")
        ctk.CTkButton(
            csv_btns,
            text="Ver en carpeta",
            width=110,
            command=self.show_csv_in_folder,
            fg_color="gray35",
            hover_color="gray25",
        ).pack(side="left", padx=(6, 0))

        btns = ctk.CTkFrame(parent)
        btns.grid(row=1, column=0, sticky="ew", padx=8, pady=8)
        btns.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkButton(
            btns,
            text="1) Generar CSV Moodle",
            command=self.action_generate_csv,
            fg_color="#22c55e",
            hover_color="#16a34a",
            height=36,
        ).grid(row=0, column=0, sticky="ew", padx=(0, 4))

        ctk.CTkButton(
            btns,
            text="2) Enviar correos",
            command=self.action_send_emails,
            fg_color="#3b82f6",
            hover_color="#1d4ed8",
            height=36,
        ).grid(row=0, column=1, sticky="ew", padx=(4, 0))

    # ---------- Clave SMTP ----------
    def toggle_smtp_key_visibility(self):
        hidden = self.smtp_key_entry.cget("show") == "•"
        self.smtp_key_entry.configure(show="" if hidden else "•")
        self.smtp_key_toggle.configure(text="Ocultar" if hidden else "Mostrar")

    def load_saved_smtp_keys(self, select_first=False):
        keys, errors = smtp_keys.find_keys()
        for err in errors:
            self.after(0, lambda e=err: self.log(f"[Claves] {e}"))

        self.smtp_key_options = {}
        for k in keys:
            label = f"{smtp_keys.mask(k['password'])}  ·  {k['source']}"
            self.smtp_key_options[label] = k

        if self.smtp_key_options:
            labels = list(self.smtp_key_options)
            self.smtp_key_menu.configure(values=labels)
            self.smtp_key_menu.set(f"Claves encontradas ({len(labels)})")
            if select_first and not self.smtp_key_entry.get():
                self.select_saved_smtp_key(labels[0])
        else:
            self.smtp_key_menu.configure(values=["Sin claves encontradas"])
            self.smtp_key_menu.set("Sin claves encontradas")

    def select_saved_smtp_key(self, label):
        key = self.smtp_key_options.get(label)
        if key is None:
            return
        self.smtp_key_entry.delete(0, "end")
        self.smtp_key_entry.insert(0, key["password"])
        self.smtp_key_menu.set(label)

    def forget_selected_smtp_key(self):
        key = self.smtp_key_options.get(self.smtp_key_menu.get())
        if key is None:
            messagebox.showinfo("Claves", "Selecciona en el desplegable la clave guardada que quieres olvidar.")
            return
        if key["source"] != smtp_keys.SOURCE_CACHE:
            messagebox.showinfo(
                "Claves",
                f"Esta clave viene de una {key['source']}.\nBórrala desde las variables de entorno de Windows.",
            )
            return
        if not messagebox.askyesno("Olvidar clave", f"¿Olvidar la clave {smtp_keys.mask(key['password'])}?"):
            return
        try:
            smtp_keys.delete_key(key["password"])
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo borrar la clave:\n{e}")
            return
        if self.smtp_key_entry.get() == key["password"]:
            self.smtp_key_entry.delete(0, "end")
        self.log(f"[Claves] Clave {smtp_keys.mask(key['password'])} olvidada.")
        self.load_saved_smtp_keys()

    def get_smtp_key_or_warn(self):
        key = self.smtp_key_entry.get().strip()
        if not key:
            messagebox.showerror("Falta la clave", "Ingresa la clave de aplicación de Gmail en la barra superior.")
            self.smtp_key_entry.focus_set()
        return key

    def remember_smtp_key(self, key):
        """Se llama desde el hilo de envío tras un login SMTP correcto."""
        try:
            smtp_keys.save_key(key)
            self.after(0, self.load_saved_smtp_keys)
        except Exception as e:
            self.log_threadsafe(f"[Claves] No se pudo guardar la clave: {e}")

    def build_right_panel(self, parent: ctk.CTkFrame):

        tabs = ctk.CTkTabview(parent)
        tabs.grid(row=0, column=0, sticky="nsew", padx=8, pady=(8, 4))
        self.tabs = tabs
        parent.grid_rowconfigure(0, weight=2)
        parent.grid_rowconfigure(1, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        tab_excel = tabs.add("Excel")
        tab_moodle = tabs.add("Moodle CSV")
        tab_csv = tabs.add("CSV envío")
        tab_email = tabs.add("Correo (preview)")

        self.tree_excel, self.info_excel = self.make_table(tab_excel, "Vista previa Excel importado")
        self.tree_moodle, self.info_moodle = self.make_table(tab_moodle, "Vista previa datos formateados (Moodle)")
        self.tree_csv, self.info_csv = self.make_table(tab_csv, "Vista previa CSV para envío")

        self.build_email_preview(tab_email)

        queue_frame = ctk.CTkFrame(parent)
        queue_frame.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
        queue_frame.grid_rowconfigure(1, weight=1)
        queue_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(queue_frame, text="Cola de envío / Log", font=("Segoe UI Semibold", 13)).grid(
            row=0, column=0, sticky="w", padx=4, pady=(4, 0)
        )

        self.text_log = ctk.CTkTextbox(queue_frame, wrap="word")
        self.text_log.grid(row=1, column=0, sticky="nsew", padx=4, pady=4)

    def make_table(self, parent: ctk.CTkFrame, title: str):
        parent.grid_rowconfigure(2, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(parent, text=title, font=("Segoe UI Semibold", 13)).grid(
            row=0, column=0, sticky="w", padx=4, pady=(4, 2)
        )

        info_label = ctk.CTkLabel(
            parent,
            text="Sin datos",
            font=("Segoe UI", 10),
            text_color=("gray70", "gray70"),
        )
        info_label.grid(row=1, column=0, sticky="w", padx=4, pady=(0, 2))

        tree = ctk.CTkTextbox(parent, wrap="none", font=("Consolas", 11))
        tree.grid(row=2, column=0, sticky="nsew", padx=4, pady=4)
        tree.insert("1.0", "(sin datos)")
        tree.configure(state="disabled")

        return tree, info_label

    def build_email_preview(self, parent: ctk.CTkFrame):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_columnconfigure(1, weight=2)

        ctk.CTkLabel(parent, text="Previsualización de correo", font=("Segoe UI Semibold", 13)).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=4, pady=(4, 2)
        )

        header_frame = ctk.CTkFrame(parent)
        header_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=4, pady=4)
        header_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(header_frame, text="Para:").grid(row=0, column=0, sticky="w")
        self.email_to_entry = ctk.CTkEntry(header_frame)
        self.email_to_entry.grid(row=0, column=1, sticky="ew", padx=4, pady=2)

        ctk.CTkLabel(header_frame, text="Asunto:").grid(row=1, column=0, sticky="w")
        self.email_subject_entry = ctk.CTkEntry(header_frame)
        self.email_subject_entry.grid(row=1, column=1, sticky="ew", padx=4, pady=2)


        plain_frame = ctk.CTkFrame(parent)
        html_frame = ctk.CTkFrame(parent)
        plain_frame.grid(row=2, column=0, sticky="nsew", padx=(4, 2), pady=4)
        html_frame.grid(row=2, column=1, sticky="nsew", padx=(2, 4), pady=4)
        parent.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(plain_frame, text="Texto plano (fallback):").pack(anchor="w", padx=4, pady=(4, 2))
        self.email_plain_text = ctk.CTkTextbox(plain_frame, wrap="word")
        self.email_plain_text.pack(fill="both", expand=True, padx=4, pady=(0, 4))

        ctk.CTkLabel(
            html_frame,
            text="HTML (renderizado):",
        ).pack(anchor="w", padx=4, pady=(4, 2))

        self.html_preview_scroll = ctk.CTkScrollableFrame(html_frame, fg_color="#eaeff4")
        self.html_preview_scroll.pack(fill="both", expand=True, padx=4, pady=(0, 4))
        self.html_preview = None
        self.set_html_preview_content(text="Generando vista previa...")
        self.html_preview_image = None
        self.html_preview_pil = None
        self.html_preview_seq = 0
        self.html_preview_scroll.bind("<Configure>", lambda e: self.fit_html_preview_image())

        preview_btns = ctk.CTkFrame(parent, fg_color="transparent")
        preview_btns.grid(row=3, column=0, columnspan=2, sticky="e", padx=8, pady=(0, 4))

        ctk.CTkButton(
            preview_btns,
            text="Abrir en navegador",
            command=self.open_email_preview_in_browser,
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            preview_btns,
            text="Actualizar preview (primer usuario)",
            command=self.update_email_preview_first_user,
        ).pack(side="left")

    def log(self, msg: str):
        self.text_log.insert("end", msg + "\n")
        self.text_log.see("end")

    def log_threadsafe(self, msg: str):
        self.after(0, lambda: self.log(msg))

    def show_df_in_textbox(self, textbox: ctk.CTkTextbox, df: pd.DataFrame,
                           info_label: ctk.CTkLabel, source_label: str, extra_info: str = ""):
        textbox.configure(state="normal")
        textbox.delete("1.0", "end")

        if df is None or df.empty:
            textbox.insert("1.0", "(sin datos)")
            textbox.configure(state="disabled")
            info_label.configure(text=f"{source_label}: 0 filas")
            return

        max_rows = 500
        textbox.insert("end", format_table(df.iloc[:max_rows]))
        if len(df) > max_rows:
            textbox.insert("end", f"\n... ({len(df) - max_rows} filas más)\n")

        textbox.configure(state="disabled")

        base_info = f"{source_label}: {len(df)} filas, {len(df.columns)} columnas"
        if extra_info:
            base_info += f" · {extra_info}"
        info_label.configure(text=base_info)

    def browse_excel(self):
        path = filedialog.askopenfilename(
            title="Seleccionar Excel de participantes",
            filetypes=[("Excel", "*.xlsx *.xls")]
        )
        if path:
            self.var_excel_path.set(path)
            # Siempre junto al Excel elegido, para no sobrescribir el CSV de otro curso
            self.var_csv_output_path.set(os.path.splitext(path)[0] + "_moodle.csv")
            self.refresh_excel_preview()

    def browse_csv_output(self):
        path = filedialog.asksaveasfilename(
            title="Guardar CSV Moodle",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")]
        )
        if path:
            self.var_csv_output_path.set(path)

    def browse_csv_mail(self):
        path = filedialog.askopenfilename(
            title="Seleccionar CSV para envío de correos",
            filetypes=[("CSV", "*.csv")]
        )
        if path:
            self.var_csv_mail_path.set(path)
            self.refresh_csv_mail_preview()

    def use_moodle_csv_for_mail(self):
        csv_out = self.var_csv_output_path.get().strip()
        if not csv_out or not os.path.isfile(csv_out):
            messagebox.showwarning("Atención", "Primero genera el CSV Moodle (botón «1) Generar CSV Moodle»).")
            return
        self.var_csv_mail_path.set(csv_out)
        self.refresh_csv_mail_preview()

    def show_csv_in_folder(self):
        csv_out = self.var_csv_output_path.get().strip()
        if not csv_out or not os.path.isfile(csv_out):
            messagebox.showwarning("Atención", "Todavía no existe el CSV Moodle. Genera el CSV primero.")
            return
        subprocess.Popen(f'explorer /select,"{os.path.normpath(csv_out)}"')

    def refresh_excel_preview(self):
        path = self.var_excel_path.get().strip()
        if not path or not os.path.isfile(path):
            self.df_excel_raw = None
            empty = pd.DataFrame()
            self.show_df_in_textbox(self.tree_excel, empty, self.info_excel, "Excel")
            return
        try:
            df, info = excel_preview_table(path)
            self.df_excel_raw = df
            self.show_df_in_textbox(
                self.tree_excel,
                df,
                self.info_excel,
                "Excel",
                extra_info=f"{os.path.basename(path)} · {info}",
            )
            self.tabs.set("Excel")
            self.log(f"[Preview] Excel cargado: {path}")
        except Exception as e:
            self.log(f"[ERROR Excel preview] {e}")
            messagebox.showerror("Error", f"No se pudo leer el Excel:\n{e}")

    def refresh_moodle_preview(self):
        df = self.df_moodle if self.df_moodle is not None else pd.DataFrame()
        extra = ""
        if self.df_moodle is not None:
            extra = f"Usuarios generados: {len(self.df_moodle)}"
        self.show_df_in_textbox(self.tree_moodle, df, self.info_moodle, "Moodle CSV", extra_info=extra)
        if self.df_moodle is not None:
            self.log("[Preview] Vista Moodle actualizada.")

    def refresh_csv_mail_preview(self, show_tab=True):
        path = self.var_csv_mail_path.get().strip()
        if not path or not os.path.isfile(path):
            self.df_csv_mail = None
            self.users_mail = []
            empty = pd.DataFrame()
            self.show_df_in_textbox(self.tree_csv, empty, self.info_csv, "CSV envío")
            return
        try:
            df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
            self.df_csv_mail = df
            self.users_mail = load_users_from_csv(path)
            extra = f"Usuarios válidos para envío: {len(self.users_mail)}"
            self.show_df_in_textbox(self.tree_csv, df, self.info_csv, "CSV envío", extra_info=extra)
            if show_tab:
                # Un solo set() por acción: CTkTabview oculta las otras pestañas 100 ms después,
                # y dos set() seguidos dejan en blanco la última pestaña elegida
                self.tabs.set("CSV envío")
            self.log(f"[Preview] CSV envío cargado: {path}")
            self.update_email_preview_first_user()
        except Exception as e:
            self.log(f"[ERROR CSV preview] {e}")
            messagebox.showerror("Error", f"No se pudo leer el CSV:\n{e}")

    def update_email_preview_first_user(self):
        u = self.users_mail[0] if self.users_mail else PREVIEW_DEMO_USER
        subject, plain, html_body = self.render_preview(u)

        self.email_to_entry.delete(0, "end")
        self.email_to_entry.insert(0, u["email"])

        self.email_subject_entry.delete(0, "end")
        self.email_subject_entry.insert(0, subject)

        self.email_plain_text.configure(state="normal")
        self.email_plain_text.delete("1.0", "end")
        self.email_plain_text.insert("1.0", plain)
        self.email_plain_text.configure(state="disabled")

        self.render_html_preview(html_body)

        origen = "primer usuario del CSV" if self.users_mail else "datos de ejemplo"
        self.log(f"[Preview] Correo ({origen}) para: {u['email']}")

    def render_html_preview(self, html_body):
        """Renderiza el HTML en segundo plano y muestra la imagen en la pestaña."""
        self.html_preview_seq += 1
        seq = self.html_preview_seq
        self.html_preview_pil = None
        self.set_html_preview_content(text="Generando vista previa...")

        def worker():
            try:
                img = render_html_screenshot(html_body)
                self.after(0, lambda: self.show_html_preview_image(seq, img))
            except Exception as e:
                msg = f"No se pudo generar la vista previa: {e}\nUsa «Abrir en navegador»."
                self.after(0, lambda: self.show_html_preview_image(seq, None, msg))

        threading.Thread(target=worker, daemon=True).start()

    def show_html_preview_image(self, seq, img, error_msg=""):
        if seq != self.html_preview_seq:
            return  # llegó una vista previa más nueva
        if img is None:
            self.set_html_preview_content(text=error_msg)
            self.log(f"[Preview] {error_msg}")
            return
        self.html_preview_pil = img
        self.fit_html_preview_image()

    def fit_html_preview_image(self):
        """Ajusta la imagen al ancho disponible (sin agrandarla sobre su tamaño real)."""
        img = self.html_preview_pil
        if img is None:
            return
        target_w = min(img.width, max(360, self.html_preview_scroll.winfo_width() - 30))
        size = (target_w, round(img.height * target_w / img.width))
        if self.html_preview_image is not None and self.html_preview_image.cget("size") == size:
            return
        self.set_html_preview_content(image=ctk.CTkImage(light_image=img, dark_image=img, size=size))

    def set_html_preview_content(self, text="", image=None):
        # Se recrea el label: reconfigurar la imagen de un CTkLabel falla ("pyimage does not exist")
        if self.html_preview is not None:
            self.html_preview.destroy()
        self.html_preview_image = image
        self.html_preview = ctk.CTkLabel(
            self.html_preview_scroll,
            text=text,
            image=image,
            text_color="#3d4451",
            wraplength=320,
        )
        self.html_preview.pack(anchor="n", pady=8)

    def render_preview(self, user):
        """Igual que el correo real, pero con el logo desde disco en vez de cid:."""
        course_name = self.var_course_name.get().strip() or DEFAULT_COURSE_NAME
        aula_url = self.var_aula_url.get().strip() or DEFAULT_AULA_URL
        return render_email(user, course_name, aula_url, LOGO_PATH.as_uri())

    def open_email_preview_in_browser(self):
        user = self.users_mail[0] if self.users_mail else PREVIEW_DEMO_USER
        _, _, html_body = self.render_preview(user)
        preview_path = Path(tempfile.gettempdir()) / "perfeccionatec_preview_correo.html"
        preview_path.write_text(html_body, encoding="utf-8")
        webbrowser.open(preview_path.as_uri())
        self.log(f"[Preview] Abierto en navegador: {preview_path}")

    def action_generate_csv(self):
        excel_path = self.var_excel_path.get().strip()
        csv_out = self.var_csv_output_path.get().strip()

        if not excel_path or not os.path.isfile(excel_path):
            messagebox.showerror("Error", "Selecciona un Excel válido.")
            return

        if not csv_out:
            messagebox.showerror("Error", "Define una ruta de salida para el CSV.")
            return

        course1 = self.var_course1.get().strip()
        if not course1:
            messagebox.showerror(
                "Falta el nombre corto",
                "Ingresa el NOMBRE CORTO del curso tal como aparece en Moodle "
                "(Configuración del curso › Nombre corto).\n\n"
                "Moodle matricula por nombre corto: si usas el número ID no encontrará el curso.",
            )
            return

        try:
            type1 = int(self.var_type1.get().strip())
            profile_field = self.var_profile_field.get().strip() or DEFAULT_PROFILE_FIELD_NAME

            password_pattern = self.var_password_pattern.get().strip() or DEFAULT_PASSWORD_PATTERN
            password_year = int(self.var_password_year.get().strip() or DEFAULT_PASSWORD_YEAR)

            self.log("Normalizando Excel -> CSV Moodle...")
            df_moodle = normalize_excel_to_moodle_csv(
                excel_path=excel_path,
                csv_output_path=csv_out,
                course_field=course1,
                type1_value=type1,
                profile_field_name=profile_field,
                password_pattern=password_pattern,
                password_year=password_year,
            )
            self.df_moodle = df_moodle
            self.log(f"[OK] CSV generado ({len(df_moodle)} usuarios): {csv_out}")

            # Deja listo el CSV recién generado como fuente de correos
            mail_path = self.var_csv_mail_path.get().strip()
            if not mail_path or os.path.normcase(os.path.abspath(mail_path)) == os.path.normcase(os.path.abspath(csv_out)):
                self.var_csv_mail_path.set(csv_out)
                self.refresh_csv_mail_preview(show_tab=False)

            self.refresh_moodle_preview()
            self.tabs.set("Moodle CSV")
            messagebox.showinfo(
                "Éxito",
                f"CSV Moodle generado con {len(df_moodle)} usuarios en:\n{csv_out}\n\n"
                "Revisa la pestaña «Moodle CSV». Usa «Ver en carpeta» para ubicar el archivo.",
            )
        except Exception as e:
            self.log(f"[ERROR] {e}")
            messagebox.showerror("Error", f"Ocurrió un error al generar el CSV:\n{e}")

    def action_send_emails(self):
        if self.sending:
            messagebox.showinfo("Aviso", "Ya hay un envío en curso.")
            return

        csv_path = self.var_csv_mail_path.get().strip()
        if not csv_path or not os.path.isfile(csv_path):
            messagebox.showerror("Error", "Selecciona un CSV válido para envío de correos.")
            return

        users = load_users_from_csv(csv_path)
        if not users:
            messagebox.showerror("Error", "No se encontraron usuarios válidos en el CSV.")
            return

        self.users_mail = users
        self.refresh_csv_mail_preview()

        try:
            assets = load_email_assets()
        except FileNotFoundError as e:
            self.log(f"[ERROR] {e}")
            messagebox.showerror("Error", str(e))
            return

        smtp_password = self.get_smtp_key_or_warn()
        if not smtp_password:
            return

        if not messagebox.askyesno(
            "Confirmar envío",
            f"Se enviarán correos a {len(users)} usuarios.\n\n¿Continuar?"
        ):
            self.log("Envío cancelado por el usuario.")
            return

        self.log(f"== Iniciando envío a {len(users)} usuarios ==")
        self.start_sending(users, smtp_password, assets)

    def action_send_test_email(self):
        if self.sending:
            messagebox.showinfo("Aviso", "Ya hay un envío en curso.")
            return

        test_email = self.test_email_entry.get().strip()
        if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", test_email):
            messagebox.showerror("Correo de prueba", "Ingresa un correo de prueba válido.")
            self.test_email_entry.focus_set()
            return

        smtp_password = self.get_smtp_key_or_warn()
        if not smtp_password:
            return

        try:
            assets = load_email_assets()
        except FileNotFoundError as e:
            self.log(f"[ERROR] {e}")
            messagebox.showerror("Error", str(e))
            return

        # Mismo contenido que recibiría el primer usuario del CSV (o datos de ejemplo)
        base_user = self.users_mail[0] if self.users_mail else PREVIEW_DEMO_USER
        test_user = {**base_user, "email": test_email}
        origen = "primer usuario del CSV" if self.users_mail else "datos de ejemplo"
        self.log(f"== Envío de prueba a {test_email} ({origen}) ==")
        self.start_sending([test_user], smtp_password, assets, subject_prefix="[PRUEBA] ")

    def start_sending(self, users, smtp_password, assets, subject_prefix=""):
        course_name = self.var_course_name.get().strip() or DEFAULT_COURSE_NAME
        aula_url = self.var_aula_url.get().strip() or DEFAULT_AULA_URL
        self.sending = True

        def worker():
            try:
                send_all(
                    sender=SENDER,
                    smtp_password=smtp_password,
                    users=users,
                    course_name=course_name,
                    aula_url=aula_url,
                    assets=assets,
                    log_func=self.log_threadsafe,
                    on_login=lambda: self.remember_smtp_key(smtp_password),
                    subject_prefix=subject_prefix,
                )
                self.log_threadsafe("== Proceso de envío finalizado ==")
            except smtplib.SMTPAuthenticationError:
                self.log_threadsafe(
                    f"[ERROR] Gmail rechazó la clave de aplicación para {SENDER}. "
                    "Revisa que sea una clave de aplicación vigente de esa cuenta."
                )
            except Exception as e:
                self.log_threadsafe(f"[ERROR general envío] {e}")
            finally:
                self.sending = False

        t = threading.Thread(target=worker, daemon=True)
        t.start()
        self.sending_thread = t


if __name__ == "__main__":
    app = MoodleApp()
    app.mainloop()
