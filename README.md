# Moodle CSV + Envío de Credenciales (Perfeccionatec)

Aplicación de escritorio en **Python + CustomTkinter** para optimizar el flujo operativo de **matriculación y onboarding** de estudiantes en Moodle.

Este proyecto resuelve un problema muy real de backoffice educativo:  
**tomar una planilla Excel de participantes, normalizarla, generar el CSV compatible con Moodle y enviar credenciales por correo** con una interfaz única, evitando trabajo manual repetitivo, errores humanos y retrasos de coordinación.

---

## ¿Qué hace?

### 1) Normaliza Excel → CSV Moodle

A partir de una planilla Excel con columnas típicas de inscripción, la app genera un CSV listo para importar en Moodle con:

- `username` autogenerado desde **nombre + primer apellido + 2 letras del segundo apellido**.
- `password` generada por patrón configurable.
- `firstname`, `lastname`, `email`.
- Campo de perfil para RUT (`profile_field_rut` por defecto).
- Campos de matrícula por curso:
  - `type1` (por defecto `1`)
  - `course1` (por defecto `PSP`)

### 2) Usa el CSV para enviar correos

Puedes cargar:
- un CSV externo de correos, o
- reutilizar el **CSV Moodle generado** como fuente de envío.

El email incluye:

- asunto dinámico,
- preheader,
- versión de texto plano (fallback),
- versión HTML simple pero bien presentada,
- enlace directo al Aula.

### 3) Preview y control operativo

Incluye pestañas de vista previa:

- **Excel**
- **Moodle CSV**
- **CSV envío**
- **Correo (preview)**

Y un panel de log con seguimiento secuencial:

- intentos,
- errores,
- reintentos,
- control de “quedan X”.

---

## Estado actual y roadmap

Hoy estamos aquí:

✅ Normalización de Excel → generación de CSV compatible con Moodle  
✅ Envío automatizado de credenciales por correo desde la misma app  
✅ Preview operacional + log de trazabilidad

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

- Python 3.10+ (recomendado 3.11)
- `pandas`
- `customtkinter`
- `tkhtmlview` (opcional, para ver el HTML renderizado dentro de la app)

---
