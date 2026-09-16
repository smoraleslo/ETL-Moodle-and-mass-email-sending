# Moodle CSV + Envío de Credenciales (PerfeccionaTEC)

Aplicación de escritorio en **Python + CustomTkinter** para optimizar el flujo operativo de **matriculación y onboarding** de estudiantes en Moodle.

Este proyecto resuelve un problema muy real de backoffice educativo:  
**tomar una planilla Excel de participantes, normalizarla, generar el CSV compatible con Moodle y enviar credenciales por correo** con una interfaz única, evitando trabajo manual repetitivo, errores humanos y retrasos de coordinación.

---

## Instalación y ejecución (Windows)

```powershell
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python app.py
```

La ventana abre maximizada. Los botones **1) Generar CSV Moodle** y **2) Enviar correos** quedan fijos abajo a la izquierda; el formulario tiene scroll si la pantalla es pequeña.

---

## ¿Qué hace?

### 1) Normaliza Excel → CSV Moodle

La app lee la **primera hoja** del Excel con este formato:

| Filas | Contenido |
|---|---|
| 1 – 4 | Encabezado del formulario (título, curso, fecha…). No se leen. |
| 5 | Nombres de columna. |
| 6 en adelante | Participantes. Se omiten las filas sin RUT o sin nombres. |

Columnas obligatorias (el nombre debe coincidir exactamente; las demás columnas se ignoran):

- `Rut (con punto y con guión)`
- `Nombres ` ← **con un espacio al final**
- `Apellidos`
- `Correo electrónico` (si la celda trae varios correos, se usa el primero)

Con eso genera un CSV listo para importar en Moodle con:

- `username` autogenerado desde **primer nombre + primer apellido + 2 letras del segundo apellido**, en minúsculas y sin tildes.
- `password` generada por patrón configurable (por defecto `{username}{year}`, con el **año actual**).
- `firstname`, `lastname`, `email`.
- Campo de perfil para RUT (`profile_field_rut` por defecto).
- Campos de matrícula: `type1` (por defecto `1`) y `course1` (ID del curso en Moodle).

En [`ejemplos/planilla_inscripcion_prueba.xlsx`](ejemplos/planilla_inscripcion_prueba.xlsx) hay una planilla de prueba con datos ficticios (correos `@example.com`) que cubre casos borde: tildes, mayúsculas, un solo apellido, espacios sobrantes, dos correos en una celda, filas vacías y filas sin RUT.

### 2) Envía las credenciales por correo

Puedes cargar un CSV externo de correos o reutilizar el **CSV Moodle generado** como fuente de envío.

Cada correo incluye:

- saludo con el primer nombre, curso asignado, usuario y contraseña,
- botón y enlace directo al Aula Virtual,
- **logo de PerfeccionaTEC** embebido en el correo (no depende de imágenes externas),
- **Manual de Ingreso al Aula en PDF adjunto** (`Manual_Ingreso_Aula.pdf`),
- preheader y versión de texto plano (fallback).

La plantilla HTML usa tablas y estilos en línea, sin SVG, para verse bien en Gmail y Outlook. El envío es secuencial, con reintentos y pausa entre correos.

### 3) Envío de prueba

En la barra superior, **Correo de prueba** + **Enviar prueba** manda un único correo a la dirección indicada, con el asunto prefijado `[PRUEBA]`. Usa los datos del primer usuario del CSV cargado o, si no hay CSV, datos de ejemplo.

### 4) Vista previa y control operativo

Pestañas de vista previa:

- **Excel**
- **Moodle CSV**
- **CSV envío**
- **Correo (preview)**: el HTML se renderiza con Microsoft Edge o Google Chrome en modo headless y se muestra como imagen, igual a como se verá en el navegador. El botón **Abrir en navegador** abre la misma vista en una pestaña.

Y un panel de log con seguimiento secuencial: intentos, errores, reintentos y control de “quedan X”.

---

## Clave de aplicación de Gmail

Los correos salen desde `perfeccionatec@gmail.com` por `smtp.gmail.com:465`. Gmail exige una **clave de aplicación** (requiere verificación en dos pasos); la contraseña normal de la cuenta no funciona.

- Se ingresa en la barra superior, oculta, con botón **Mostrar / Ocultar**.
- Tras un login SMTP correcto se **guarda cifrada con DPAPI de Windows** en `%LOCALAPPDATA%\PerfeccionaTEC\smtp_keys.bin`. Solo la puede leer el mismo usuario de Windows en el mismo equipo, y nunca se guarda dentro del repositorio.
- Al iniciar, la app busca claves en esa caché y en las variables de entorno `PERFECCIONATEC_SMTP_PASSWORD`, `GMAIL_APP_PASSWORD` y `SMTP_APP_PASSWORD`. Las encontradas aparecen en el desplegable **Claves encontradas** (solo se muestran los últimos 4 caracteres) y la más reciente se precarga.
- **Olvidar** borra de la caché la clave seleccionada (las de variables de entorno se eliminan desde Windows).

---

## Estructura

```
app.py              Interfaz, ETL Excel → CSV, plantillas y envío de correos
smtp_keys.py        Guardado cifrado y búsqueda de claves de aplicación
assets/
  perfeccionatec.png                      Logo embebido en el correo
  Manual de ingreso Perfeccionatec.pdf    Manual adjunto a cada correo
ejemplos/
  planilla_inscripcion_prueba.xlsx        Planilla de prueba con datos ficticios
```

---

## Limitaciones conocidas

- **Apellidos compuestos**: “de la Fuente Araya” genera el usuario `tomasdela`, porque toma “de” como primer apellido y “la” como segundo.
- Los espacios dobles dentro de un apellido se mantienen en `lastname`.
- Cada correo pesa ~3,3 MB por el PDF adjunto. Una cuenta Gmail estándar permite alrededor de 500 destinatarios por día.
- El guardado cifrado de claves solo funciona en Windows.

---

## Estado actual y roadmap

Hoy estamos aquí:

✅ Normalización de Excel → generación de CSV compatible con Moodle  
✅ Envío automatizado de credenciales por correo, con logo y manual adjunto  
✅ Envío de prueba y vista previa renderizada del correo  
✅ Gestión segura de la clave de aplicación de Gmail  
✅ Log de trazabilidad del envío

En términos prácticos:  
**nos encontramos aquí, pero a futuro la idea es subir usuarios automáticamente a Moodle.**

### Visión a futuro

La evolución natural del producto es pasar de un flujo basado en importación manual a un modelo **end-to-end**:

🚀 Subir usuarios automáticamente a Moodle, idealmente mediante:

- integración directa con la **API de Moodle**,
- validaciones previas de duplicidad y campos obligatorios,
- ejecución de matrícula por curso desde la interfaz,
- auditoría, métricas y reporte de resultados.

En términos de madurez operativa, el objetivo es transformar esta herramienta en un **módulo de automatización de onboarding**, donde el CSV sea un **respaldo opcional**, no el corazón del flujo.

En resumen:  
**hoy estandarizamos y aceleramos el proceso manual; mañana lo eliminamos.**

---

## Stack

- Python 3.10+ (probado con 3.14)
- `pandas` y `openpyxl` (lectura de Excel)
- `customtkinter` (interfaz)
- `pillow` (vista previa del correo)
- Microsoft Edge o Google Chrome (render de la vista previa)
