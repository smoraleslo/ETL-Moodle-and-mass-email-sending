import os
import csv
import re
import ssl
import time
import smtplib
import threading
import unicodedata
from pathlib import Path
from email.message import EmailMessage
from string import Template

import pandas as pd

import customtkinter as ctk
from tkinter import filedialog, messagebox

try:
    from tkhtmlview import HTMLLabel
    HAS_HTML_PREVIEW = True
except ImportError:
    HAS_HTML_PREVIEW = False


# =========================
# CONFIGURACIÓN POR DEFECTO
# =========================

DEFAULT_COURSE_NAME = "Ingrese curso"
DEFAULT_MOODLE_COURSE_FIELD = "Ingrese ID"
DEFAULT_MOODLE_TYPE1 = 1
DEFAULT_PROFILE_FIELD_NAME = "profile_field_rut"

DEFAULT_PASSWORD_YEAR = 2025
# placeholders: {username}, {year}, {rut}, {email}
DEFAULT_PASSWORD_PATTERN = "{username}{year}"

DEFAULT_AULA_URL = "https://aulavirtual.perfeccionatec.cl/"

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465
SENDER = "perfeccionatec@gmail.com"

THROTTLE_SECONDS = 1.0
MAX_RETRIES = 3

USERNAME_NORMALIZE_ACCENTS = True

SUBJECT_TEMPLATE = Template("Tus credenciales — Aula $nombre_curso")

PREHEADER_TEMPLATE = Template(
    "Tu acceso al Aula Virtual de Perfeccionatec. Usuario: $usuario."
)

HTML_TEMPLATE = Template(r"""
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width">
  <title>Credenciales de acceso</title>
</head>
<body style="margin:0;padding:0;background:#f4f7fb;font-family:Arial,Helvetica,sans-serif;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f4f7fb;">
    <tr>
      <td align="center" style="padding:24px;">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0"
               style="max-width:640px;background:#ffffff;border-radius:16px;overflow:hidden;
                      border:1px solid #e6ecf5;box-shadow:0 12px 30px rgba(15,23,42,0.12);">
          <tr>
            <td style="background:#0b63f6; padding:24px;">
              <h1 style="margin:0;color:#ffffff;font-size:22px;line-height:1.3;">
                Aula Virtual Perfeccionatec
              </h1>
              <p style="margin:8px 0 0 0;color:#dfe9ff;font-size:14px;">
                Añadido al curso: $nombre_curso
              </p>
            </td>
          </tr>
          <!-- Sin 'transparent' para evitar problemas con Tk / tkhtmlview -->
          <tr>
            <td style="display:none; height:0; width:0; overflow:hidden;">
              $preheader
            </td>
          </tr>
          <tr>
            <td style="padding:28px 24px 12px 24px;">
              <p style="margin:0 0 12px 0;font-size:16px;color:#1a1f36;">
                Hola <strong>$nombre</strong>,
              </p>
              <p style="margin:0 0 16px 0;font-size:16px;color:#1a1f36;">
                Te compartimos tus <strong>credenciales de acceso</strong> al Aula Virtual:
              </p>
              <table role="presentation" cellpadding="0" cellspacing="0"
                     style="width:100%;margin:8px 0 16px 0;background:#f8fafc;
                            border:1px solid #e6ecf5;border-radius:12px;">
                <tr>
                  <td style="padding:14px 16px;font-size:14px;color:#1a1f36;">
                    <div style="margin-bottom:6px;">
                      <strong>Usuario:</strong>
                      <span style="font-family:Consolas,Menlo,monospace;">$usuario</span>
                    </div>
                    <div>
                      <strong>Contraseña:</strong>
                      <span style="font-family:Consolas,Menlo,monospace;">$contrasena</span>
                    </div>
                  </td>
                </tr>
              </table>
              <p style="margin:0 0 20px 0;font-size:14px;color:#425466;">
                Recomendación: cambia tu contraseña al iniciar sesión por una que solo tú conozcas.
              </p>
              <table role="presentation" cellpadding="0" cellspacing="0" style="margin:0 0 8px 0;">
                <tr>
                  <td align="left">
                    <a href="$aula_url"
                       style="display:inline-block;background:#0b63f6;color:#ffffff;text-decoration:none;
                              font-size:15px;line-height:1;padding:14px 18px;border-radius:10px;
                              border:1px solid #0b5ae0;">
                      Acceder al Aula
                    </a>
                  </td>
                </tr>
              </table>
              <p style="margin:8px 0 0 0;font-size:13px;color:#6b7280;">
                Enlace directo:
                <a href="$aula_url" style="color:#0b63f6;text-decoration:none;">$aula_url</a>
              </p>
            </td>
          </tr>
          <tr>
            <td style="padding:18px 24px 24px 24px;">
              <hr style="border:none;border-top:1px solid #e6ecf5;margin:0 0 12px 0;">
              <p style="margin:0;font-size:12px;color:#6b7280;">
                ¿Dudas o problemas de acceso? Responde este correo y te ayudamos.
              </p>
              <p style="margin:6px 0 0 0;font-size:12px;color:#94a3b8;">
                © Perfeccionatec — Este mensaje fue enviado automáticamente.
              </p>
            </td>
          </tr>
        </table>
        <p style="margin:12px 0 0 0;font-size:11px;color:#94a3b8;">
          Si el botón no funciona, copia y pega en tu navegador: $aula_url
        </p>
      </td>
    </tr>
  </table>
</body>
</html>
""")

PLAIN_TEMPLATE = Template("""
Hola $nombre,

Te compartimos tus credenciales de acceso al Aula Virtual Perfeccionatec ($nombre_curso).

Usuario: $usuario
Contraseña: $contrasena

Acceso: $aula_url

Recomendación: cambia tu contraseña al iniciar sesión.

Saludos,
Equipo Perfeccionatec
""".strip())

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


def normalize_excel_to_moodle_csv(
    excel_path: str,
    csv_output_path: str,
    course_field: str,
    type1_value: int,
    profile_field_name: str,
    password_pattern: str,
    password_year: int,
):
    df = pd.read_excel(excel_path, sheet_name=0)

    header = df.iloc[3]
    clean_df = df.iloc[4:].copy()
    clean_df.columns = header.values

    clean_df = clean_df.rename(columns={
        "Rut (con punto y con guión)": "rut",
        "Nombres ": "nombres",
        "Apellidos": "apellidos",
        "Correo electrónico": "email",
    })

    participants = clean_df[clean_df["rut"].notna() & clean_df["nombres"].notna()].copy()

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

    moodle.to_csv(csv_output_path, index=False, encoding="utf-8")
    return moodle


def load_users_from_csv(path: str):
    users = []
    with open(path, newline="", encoding="utf-8") as f:
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


def build_message(sender, recipient, subject, plain, html):
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.set_content(plain)
    msg.add_alternative(html, subtype="html")
    return msg


def send_all(sender, smtp_password, users, course_name, aula_url, log_func):
    """
    Envío con contador y 'cuenta regresiva':
    - [1/50] Enviando a...
    - [ENVIADO 1/50] ... (quedan 49)
    """
    total = len(users)
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context) as smtp:
        smtp.login(sender, smtp_password)
        for idx, u in enumerate(users, start=1):
            restantes = total - idx
            log_func(f"[{idx}/{total}] Enviando a {u['email']}...")

            preheader = PREHEADER_TEMPLATE.substitute(usuario=u["usuario"])
            subject = SUBJECT_TEMPLATE.substitute(nombre_curso=course_name)

            plain = PLAIN_TEMPLATE.substitute(
                nombre=u["nombre"],
                usuario=u["usuario"],
                contrasena=u["contrasena"],
                aula_url=aula_url,
                nombre_curso=course_name,
            )
            html = HTML_TEMPLATE.substitute(
                nombre=u["nombre"],
                usuario=u["usuario"],
                contrasena=u["contrasena"],
                aula_url=aula_url,
                preheader=preheader,
                nombre_curso=course_name,
            )

            msg = build_message(sender, u["email"], subject, plain, html)

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
        self.geometry("1200x720")
        self.minsize(1100, 650)

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

    # ---------- UI ----------
    def build_ui(self):
        main_frame = ctk.CTkFrame(self, corner_radius=0)
        main_frame.pack(fill="both", expand=True)

        main_frame.grid_columnconfigure(0, weight=1, minsize=380)
        main_frame.grid_columnconfigure(1, weight=2)
        main_frame.grid_rowconfigure(1, weight=1)

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

        left = ctk.CTkFrame(main_frame)
        left.grid(row=1, column=0, sticky="nsew", padx=(5, 2), pady=(0, 5))
        left.grid_rowconfigure(4, weight=1)
        left.grid_columnconfigure(0, weight=1)
        self.build_left_panel(left)

        right = ctk.CTkFrame(main_frame)
        right.grid(row=1, column=1, sticky="nsew", padx=(2, 5), pady=(0, 5))
        right.grid_rowconfigure(1, weight=1)
        right.grid_columnconfigure(0, weight=1)
        self.build_right_panel(right)

    def build_left_panel(self, parent: ctk.CTkFrame):
        course_frame = ctk.CTkFrame(parent)
        course_frame.grid(row=0, column=0, sticky="ew", padx=8, pady=8)
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

        ctk.CTkLabel(course_frame, text="course1 (Moodle):").grid(row=2, column=0, sticky="w")
        ctk.CTkEntry(course_frame, textvariable=self.var_course1, width=80).grid(
            row=2, column=1, sticky="w", padx=4, pady=2
        )

        ctk.CTkLabel(course_frame, text="type1:").grid(row=3, column=0, sticky="w")
        type_row = ctk.CTkFrame(course_frame)
        type_row.grid(row=3, column=1, sticky="ew", padx=4, pady=2)
        ctk.CTkEntry(type_row, textvariable=self.var_type1, width=60).pack(side="left")
        ctk.CTkLabel(
            type_row,
            text="1 = crear/actualizar usuario y matricular (recomendado)",
            font=("Segoe UI", 9),
            text_color=("gray78", "gray78"),
        ).pack(side="left", padx=4)

        info_type = ctk.CTkLabel(
            course_frame,
            text="En Moodle puedes usar type2, type3, etc. como pares con course2, course3 para otros cursos.",
            font=("Segoe UI", 9),
            text_color=("gray70", "gray70"),
            wraplength=320,
            justify="left",
        )
        info_type.grid(row=4, column=0, columnspan=2, sticky="w", pady=(4, 0))

        ctk.CTkLabel(course_frame, text="Campo RUT en Moodle:").grid(row=5, column=0, sticky="w", pady=(6, 0))
        ctk.CTkEntry(course_frame, textvariable=self.var_profile_field).grid(
            row=5, column=1, sticky="ew", padx=4, pady=(6, 4)
        )

        pwd_frame = ctk.CTkFrame(parent)
        pwd_frame.grid(row=1, column=0, sticky="ew", padx=8, pady=8)
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
            wraplength=320,
            justify="left",
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(4, 0))

        aula_frame = ctk.CTkFrame(parent)
        aula_frame.grid(row=2, column=0, sticky="ew", padx=8, pady=8)
        aula_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(aula_frame, text="Aula Virtual", font=("Segoe UI Semibold", 14)).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 6)
        )

        ctk.CTkLabel(aula_frame, text="URL Aula:").grid(row=1, column=0, sticky="w")
        ctk.CTkEntry(aula_frame, textvariable=self.var_aula_url).grid(
            row=1, column=1, sticky="ew", padx=4, pady=2
        )


        smtp_frame = ctk.CTkFrame(parent)
        smtp_frame.grid(row=3, column=0, sticky="ew", padx=8, pady=8)
        smtp_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(smtp_frame, text="SMTP (fijo en el código)", font=("Segoe UI Semibold", 14)).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 6)
        )

        ctk.CTkLabel(smtp_frame, text="Servidor:").grid(row=1, column=0, sticky="w")
        server_entry = ctk.CTkEntry(smtp_frame)
        server_entry.insert(0, SMTP_SERVER)
        server_entry.configure(state="disabled")
        server_entry.grid(row=1, column=1, sticky="ew", padx=4, pady=2)

        ctk.CTkLabel(smtp_frame, text="Puerto:").grid(row=2, column=0, sticky="w")
        port_entry = ctk.CTkEntry(smtp_frame, width=80)
        port_entry.insert(0, str(SMTP_PORT))
        port_entry.configure(state="disabled")
        port_entry.grid(row=2, column=1, sticky="w", padx=4, pady=2)

        ctk.CTkLabel(smtp_frame, text="Remitente:").grid(row=3, column=0, sticky="w")
        sender_entry = ctk.CTkEntry(smtp_frame)
        sender_entry.insert(0, SENDER)
        sender_entry.configure(state="disabled")
        sender_entry.grid(row=3, column=1, sticky="ew", padx=4, pady=2)

        files_frame = ctk.CTkFrame(parent)
        files_frame.grid(row=4, column=0, sticky="nsew", padx=8, pady=8)
        files_frame.grid_columnconfigure(1, weight=1)
        files_frame.grid_rowconfigure(3, weight=1)

        ctk.CTkLabel(files_frame, text="Archivos y acciones", font=("Segoe UI Semibold", 14)).grid(
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
        ctk.CTkButton(files_frame, text="Guardar como", command=self.browse_csv_output, width=120).grid(
            row=2, column=2, padx=2, pady=2
        )


        ctk.CTkLabel(files_frame, text="CSV para envío de correos:").grid(row=3, column=0, sticky="w")
        ctk.CTkEntry(files_frame, textvariable=self.var_csv_mail_path).grid(
            row=3, column=1, sticky="ew", padx=4, pady=2
        )
        ctk.CTkButton(files_frame, text="Buscar", command=self.browse_csv_mail, width=80).grid(
            row=3, column=2, padx=2, pady=2
        )

        ctk.CTkButton(
            files_frame,
            text="Usar CSV Moodle como fuente de correos",
            command=self.use_moodle_csv_for_mail,
        ).grid(row=4, column=1, sticky="w", padx=4, pady=(4, 8))

        btns = ctk.CTkFrame(files_frame)
        btns.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        btns.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkButton(
            btns,
            text="1) Generar CSV Moodle",
            command=self.action_generate_csv,
            fg_color="#22c55e",
            hover_color="#16a34a",
        ).grid(row=0, column=0, sticky="ew", padx=(0, 4))

        ctk.CTkButton(
            btns,
            text="2) Enviar correos",
            command=self.action_send_emails,
            fg_color="#3b82f6",
            hover_color="#1d4ed8",
        ).grid(row=0, column=1, sticky="ew", padx=(4, 0))

    def build_right_panel(self, parent: ctk.CTkFrame):

        tabs = ctk.CTkTabview(parent)
        tabs.grid(row=0, column=0, sticky="nsew", padx=8, pady=(8, 4))
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

        container = ctk.CTkScrollableFrame(parent)
        container.grid(row=2, column=0, sticky="nsew", padx=4, pady=4)

        tree = ctk.CTkTextbox(container, height=200)
        tree.configure(font=("Consolas", 10))
        tree.pack(fill="both", expand=True)

        return tree, info_label

    def build_email_preview(self, parent: ctk.CTkFrame):
        parent.grid_rowconfigure(3, weight=1)
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_columnconfigure(1, weight=1)

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
            text="HTML (vista aproximada):",
        ).pack(anchor="w", padx=4, pady=(4, 2))

        if HAS_HTML_PREVIEW:
            self.html_preview = HTMLLabel(html_frame, html="", background="white")
            self.html_preview.pack(fill="both", expand=True, padx=4, pady=(0, 4))
        else:
            self.html_preview = ctk.CTkTextbox(html_frame, wrap="word")
            self.html_preview.insert(
                "1.0",
                "Para ver el HTML renderizado instala tkhtmlview:\n\npip install tkhtmlview",
            )
            self.html_preview.configure(state="disabled")
            self.html_preview.pack(fill="both", expand=True, padx=4, pady=(0, 4))

        ctk.CTkButton(
            parent,
            text="Actualizar preview (primer usuario)",
            command=self.update_email_preview_first_user,
        ).grid(row=3, column=0, columnspan=2, sticky="e", padx=8, pady=(0, 4))

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

        cols = list(df.columns)
        max_cols = min(len(cols), 8)
        use_cols = cols[:max_cols]

        header_line = " | ".join([str(c) for c in use_cols])
        textbox.insert("end", header_line + "\n")
        textbox.insert("end", "-" * len(header_line) + "\n")

        max_rows = 50
        for _, row in df.iloc[:max_rows].iterrows():
            vals = [str(row[c]) for c in use_cols]
            line = " | ".join(vals)
            textbox.insert("end", line + "\n")

        if len(df) > max_rows:
            textbox.insert("end", f"... ({len(df) - max_rows} filas más)\n")

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
            out_path = os.path.splitext(path)[0] + "_moodle.csv"
            if not self.var_csv_output_path.get():
                self.var_csv_output_path.set(out_path)
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
        if self.var_csv_output_path.get():
            self.var_csv_mail_path.set(self.var_csv_output_path.get())
            self.refresh_csv_mail_preview()
        else:
            messagebox.showwarning("Atención", "Primero genera el CSV Moodle.")

    def refresh_excel_preview(self):
        path = self.var_excel_path.get().strip()
        if not path or not os.path.isfile(path):
            self.df_excel_raw = None
            empty = pd.DataFrame()
            self.show_df_in_textbox(self.tree_excel, empty, self.info_excel, "Excel")
            return
        try:
            df = pd.read_excel(path, sheet_name=0)
            self.df_excel_raw = df
            self.show_df_in_textbox(
                self.tree_excel,
                df,
                self.info_excel,
                "Excel",
                extra_info=os.path.basename(path),
            )
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

    def refresh_csv_mail_preview(self):
        path = self.var_csv_mail_path.get().strip()
        if not path or not os.path.isfile(path):
            self.df_csv_mail = None
            self.users_mail = []
            empty = pd.DataFrame()
            self.show_df_in_textbox(self.tree_csv, empty, self.info_csv, "CSV envío")
            return
        try:
            df = pd.read_csv(path)
            self.df_csv_mail = df
            self.users_mail = load_users_from_csv(path)
            extra = f"Usuarios válidos para envío: {len(self.users_mail)}"
            self.show_df_in_textbox(self.tree_csv, df, self.info_csv, "CSV envío", extra_info=extra)
            self.log(f"[Preview] CSV envío cargado: {path}")
            self.update_email_preview_first_user()
        except Exception as e:
            self.log(f"[ERROR CSV preview] {e}")
            messagebox.showerror("Error", f"No se pudo leer el CSV:\n{e}")

    def update_email_preview_first_user(self):
        if not self.users_mail:
            self.email_to_entry.delete(0, "end")
            self.email_subject_entry.delete(0, "end")
            self.email_plain_text.configure(state="normal")
            self.email_plain_text.delete("1.0", "end")
            self.email_plain_text.insert("1.0", "No hay usuarios cargados desde el CSV.")
            self.email_plain_text.configure(state="disabled")
            if HAS_HTML_PREVIEW:
                self.html_preview.set_html("<p>No hay usuarios cargados.</p>")
            self.log("[Preview] Sin usuarios para correo.")
            return

        u = self.users_mail[0]
        course_name = self.var_course_name.get().strip() or DEFAULT_COURSE_NAME
        aula_url = self.var_aula_url.get().strip() or DEFAULT_AULA_URL

        preheader = PREHEADER_TEMPLATE.substitute(usuario=u["usuario"])
        subject = SUBJECT_TEMPLATE.substitute(nombre_curso=course_name)
        plain = PLAIN_TEMPLATE.substitute(
            nombre=u["nombre"],
            usuario=u["usuario"],
            contrasena=u["contrasena"],
            aula_url=aula_url,
            nombre_curso=course_name,
        )
        html = HTML_TEMPLATE.substitute(
            nombre=u["nombre"],
            usuario=u["usuario"],
            contrasena=u["contrasena"],
            aula_url=aula_url,
            preheader=preheader,
            nombre_curso=course_name,
        )

        self.email_to_entry.delete(0, "end")
        self.email_to_entry.insert(0, u["email"])

        self.email_subject_entry.delete(0, "end")
        self.email_subject_entry.insert(0, subject)

        self.email_plain_text.configure(state="normal")
        self.email_plain_text.delete("1.0", "end")
        self.email_plain_text.insert("1.0", plain)
        self.email_plain_text.configure(state="disabled")

        if HAS_HTML_PREVIEW:
            self.html_preview.set_html(html)
        else:
            self.html_preview.configure(state="normal")
            self.html_preview.delete("1.0", "end")
            self.html_preview.insert(
                "1.0",
                "Instala tkhtmlview para ver el HTML renderizado:\n\npip install tkhtmlview\n\n---\n\n" + html,
            )
            self.html_preview.configure(state="disabled")

        self.log(f"[Preview] Correo ejemplo para: {u['email']}")

    def action_generate_csv(self):
        excel_path = self.var_excel_path.get().strip()
        csv_out = self.var_csv_output_path.get().strip()

        if not excel_path or not os.path.isfile(excel_path):
            messagebox.showerror("Error", "Selecciona un Excel válido.")
            return

        if not csv_out:
            messagebox.showerror("Error", "Define una ruta de salida para el CSV.")
            return

        try:
            course1 = self.var_course1.get().strip()
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
            self.refresh_moodle_preview()
            self.log(f"[OK] CSV generado: {csv_out}")
            messagebox.showinfo("Éxito", f"CSV Moodle generado en:\n{csv_out}")
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

        pwd_win = ctk.CTkInputDialog(
            title="App Password Gmail",
            text="Introduce tu App Password de Gmail (no se guardará):",
        )
        smtp_password = pwd_win.get_input()
        if not smtp_password:
            self.log("Envío cancelado: no se ingresó App Password.")
            return

        course_name = self.var_course_name.get().strip() or DEFAULT_COURSE_NAME
        aula_url = self.var_aula_url.get().strip() or DEFAULT_AULA_URL

        if not messagebox.askyesno(
            "Confirmar envío",
            f"Se enviarán correos a {len(users)} usuarios.\n\n¿Continuar?"
        ):
            self.log("Envío cancelado por el usuario.")
            return

        self.sending = True
        total = len(users)
        self.log(f"== Iniciando envío a {total} usuarios ==")

        def worker():
            try:
                send_all(
                    sender=SENDER,
                    smtp_password=smtp_password,
                    users=users,
                    course_name=course_name,
                    aula_url=aula_url,
                    log_func=self.log_threadsafe,
                )
                self.log_threadsafe("== Proceso de envío finalizado ==")
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
