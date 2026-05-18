# -*- coding: utf-8 -*-

"""
=============================================================
  LUBRICENTRO O'HIGGINS — Envío automático de WhatsApp
=============================================================

  Requisitos:
    pip install openpyxl pyautogui

  Uso:
    python lubricentro_whatsapp.py

=============================================================
"""

import openpyxl
import webbrowser
import urllib.parse
import time
import random
import pyautogui
import sys
import ctypes
import winsound

from datetime import date

# ─────────────────────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────────────────────

RUTA_EXCEL = r"\\DESKTOP-PCFS020\compartir\SOPORTE 2\LUBRICENTRO\Macros\Nuevas planillas de mensajes\Base de datos.xlsx"

NOMBRE_HOJA = "BASE 2026"

FILA_ENCABEZADOS = 1

# ── Columnas Excel ──────────────────────────────────────────
COL_CLIENTE       = "Cliente"
COL_TELEFONO      = "Teléfono"           # columna B — se limpia automáticamente
COL_AUTO          = "Auto"
COL_PATENTE       = "Patente"
COL_KMTS_PROX     = "Kmts Prox. Cambio"
COL_NUM_CONTACTOS = "Num de contacto"    # vacío = pendiente | 1 = enviado | 0 = no contactado
COL_FECHA_ENVIO   = "Repuesta 1"

# ── Configuración envío ─────────────────────────────────────
CANTIDAD_POR_TANDA = 10

DELAY_MIN = 4
DELAY_MAX = 6

ESPERA_CARGA_CHAT = 3

PREFIJO = "549"

# ─────────────────────────────────────────────────────────────
# MENSAJE
# ─────────────────────────────────────────────────────────────

MENSAJE = """🚘 *Recordatorio de Service*

Tu *{auto}* dominio *{patente}* está próximo a los *{km} kms* ⏱️

🔧 En *Lubricentro O'Higgins* tenemos un beneficio exclusivo para vos:
📅 *10% OFF* en cambio de aceite reservando tu turno por WhatsApp.

💥 ¡Consultá disponibilidad!

_Promoción válida abonando en efectivo o transferencia._
"""

# ─────────────────────────────────────────────────────────────
# FUNCIONES
# ─────────────────────────────────────────────────────────────

def limpiar_telefono(raw):
    """
    Limpia el campo teléfono de la columna B.
    Soporta formatos como:
      -Cel:3512078250
      ' -Cel: 3512078250'  (con espacios)
      -Cel:3512078250 -Cel:3516789012  (dos números → toma el primero)
      3512078250
    Devuelve el número normalizado para wa.me (con prefijo 549).
    """
    if not raw:
        return None

    texto = str(raw).strip()

    # Quitar todos los prefijos posibles incluyendo variantes con espacios
    for prefijo in ["-Cel:", "Cel:", "-cel:", "cel:", "-CEL:", "CEL:"]:
        texto = texto.replace(prefijo, " ")

    # Quitar guiones sueltos que no son parte de un número
    texto = texto.replace("-", " ")

    # Dividir por espacios y tomar el primer bloque con suficientes dígitos
    partes = texto.split()
    numero_raw = None
    for parte in partes:
        digitos = "".join(c for c in parte if c.isdigit())
        if len(digitos) >= 8:
            numero_raw = digitos
            break

    if not numero_raw:
        return None

    # Normalizar prefijo
    if numero_raw.startswith("0"):
        numero_raw = numero_raw[1:]

    if numero_raw.startswith("54"):
        return numero_raw

    return PREFIJO + numero_raw


def formatear_km(km):
    """Formatea kilómetros con puntos de miles"""
    try:
        return f"{int(float(km)):,}".replace(",", ".")
    except:
        return str(km)


def formatear_auto(texto):
    """Capitaliza nombre del vehículo respetando siglas"""
    if not texto:
        return ""

    especiales = {
        "BMW", "GTI", "XLS", "GLI", "VTI", "TDI",
        "HDI", "SRX", "AMG", "X5", "X6", "GNC"
    }

    palabras = str(texto).strip().split()
    resultado = []

    for p in palabras:
        if p.upper() in especiales:
            resultado.append(p.upper())
        else:
            resultado.append(p.capitalize())

    return " ".join(resultado)


def cargar_excel():
    """Carga el workbook y devuelve hoja y mapa de columnas"""
    try:
        wb = openpyxl.load_workbook(RUTA_EXCEL)
    except FileNotFoundError:
        print(f"\n❌ No encontré el archivo:\n   {RUTA_EXCEL}")
        mostrar_alerta_error("No encontré el archivo Excel.")
        sys.exit(1)

    if NOMBRE_HOJA not in wb.sheetnames:
        print(f"\n❌ La hoja '{NOMBRE_HOJA}' no existe.")
        print(f"   Hojas disponibles: {wb.sheetnames}")
        mostrar_alerta_error("La hoja configurada no existe.")
        sys.exit(1)

    ws = wb[NOMBRE_HOJA]

    encabezados = {}
    for col in ws.iter_cols(min_row=FILA_ENCABEZADOS, max_row=FILA_ENCABEZADOS):
        for cell in col:
            if cell.value:
                encabezados[str(cell.value).strip()] = cell.column

    return wb, ws, encabezados


def obtener_col(encabezados, nombre_col):
    """Busca columna por nombre exacto, luego parcial"""
    if nombre_col in encabezados:
        return encabezados[nombre_col]
    for k, v in encabezados.items():
        if nombre_col.lower() in k.lower():
            return v
    return None


def seleccionar_contactos(ws, encabezados):
    """Devuelve contactos con Num de contacto vacío (nunca notificados)"""
    col_tel    = obtener_col(encabezados, COL_TELEFONO)
    col_auto   = obtener_col(encabezados, COL_AUTO)
    col_pat    = obtener_col(encabezados, COL_PATENTE)
    col_km     = obtener_col(encabezados, COL_KMTS_PROX)
    col_nombre = obtener_col(encabezados, COL_CLIENTE)
    col_num    = obtener_col(encabezados, COL_NUM_CONTACTOS)
    col_fecha  = obtener_col(encabezados, COL_FECHA_ENVIO)

    if not col_tel or not col_pat or not col_km:
        print("\n❌ Faltan columnas clave.")
        mostrar_alerta_error("Faltan columnas clave en el Excel.")
        sys.exit(1)

    contactos = []

    for row in ws.iter_rows(min_row=FILA_ENCABEZADOS + 1, values_only=False):
        fila_num      = row[0].row
        telefono_raw  = row[col_tel - 1].value    if col_tel    else None
        auto_raw      = row[col_auto - 1].value   if col_auto   else ""
        patente       = row[col_pat - 1].value    if col_pat    else None
        km            = row[col_km - 1].value     if col_km     else None
        nombre        = row[col_nombre - 1].value if col_nombre else ""
        num_contactos = row[col_num - 1].value    if col_num    else None

        # Saltar filas sin datos esenciales
        if not telefono_raw or not patente or not km:
            continue

        # Solo incluir los que tienen la columna vacía (nunca notificados)
        if num_contactos is None or num_contactos == "":
            pass
        else:
            continue

        telefono = limpiar_telefono(telefono_raw)
        if not telefono:
            continue

        contactos.append({
            "fila":      fila_num,
            "telefono":  telefono,
            "auto":      formatear_auto(auto_raw),
            "patente":   str(patente).strip().upper(),
            "km":        formatear_km(km),
            "nombre":    str(nombre).strip() if nombre else "Cliente",
            "col_num":   col_num,
            "col_fecha": col_fecha,
        })

    return contactos


def enviar_mensaje(contacto):
    """Abre WhatsApp Desktop y envía el mensaje"""
    mensaje = MENSAJE.format(
        auto=contacto["auto"],
        patente=contacto["patente"],
        km=contacto["km"]
    )

    mensaje_url = urllib.parse.quote(mensaje, safe="")
    url = (
        f"whatsapp://send?"
        f"phone={contacto['telefono']}"
        f"&text={mensaje_url}"
    )

    print(f"   📱 Abriendo WhatsApp con {contacto['nombre']} ({contacto['patente']})...")
    webbrowser.open(url)
    time.sleep(ESPERA_CARGA_CHAT)

    pyautogui.press("enter")
    print(f"   ✅ Mensaje enviado a {contacto['telefono']}")


def marcar_enviado(ws, contacto, valor=1):
    """Actualiza Num de contacto y fecha en Excel"""
    hoy = date.today().strftime("%d/%m/%Y")

    if contacto["col_num"]:
        ws.cell(row=contacto["fila"], column=contacto["col_num"]).value = valor

    if contacto["col_fecha"]:
        ws.cell(row=contacto["fila"], column=contacto["col_fecha"]).value = hoy


def mostrar_alerta_error(mensaje):
    """Popup de error con sonido"""
    winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
    ctypes.windll.user32.MessageBoxW(
        0, mensaje,
        "Error — Lubricentro O'Higgins",
        0x00001000 | 0x00000010
    )


def mostrar_alerta_final(enviados, errores):
    """Popup de resumen al finalizar"""
    pyautogui.hotkey("win", "d")
    time.sleep(0.5)

    mensaje = f"Proceso finalizado.\n\n✅ Enviados: {enviados}"
    if errores:
        mensaje += f"\n⚠️ Con error: {errores} (marcados con 0 en la base)"

    winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
    ctypes.windll.user32.MessageBoxW(
        0, mensaje,
        "Lubricentro O'Higgins",
        0x00001000 | 0x00000040
    )


# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────

def main():
    print("=" * 55)
    print("  LUBRICENTRO O'HIGGINS — WhatsApp automático")
    print("=" * 55)

    print("\n📂 Cargando base de datos...")
    wb, ws, encabezados = cargar_excel()
    print(f"✅ Base cargada: {NOMBRE_HOJA}")

    print("\n🔍 Buscando pendientes...")
    todos = seleccionar_contactos(ws, encabezados)

    if not todos:
        print("\n✅ No hay pendientes.")
        mostrar_alerta_final(enviados=0, errores=0)
        sys.exit(0)

    tanda = todos[:CANTIDAD_POR_TANDA]

    print(f"\n📋 Encontré {len(todos)} pendientes — enviando {len(tanda)} hoy")
    print("-" * 55)
    for i, c in enumerate(tanda, 1):
        print(f"  {i:2}. {c['nombre']:<25} | {c['patente']} | {c['km']} km")
    print("-" * 55)

    confirmacion = input(f"\n¿Confirmar envío de {len(tanda)} mensajes? (s/n): ").strip().lower()
    if confirmacion != "s":
        print("❌ Envío cancelado.")
        sys.exit(0)

    print("\n⚠️  Asegurate de tener WhatsApp Desktop abierto.")
    print("⏳ Comenzando en 5 segundos...")
    time.sleep(5)

    enviados = 0
    errores  = 0

    for i, contacto in enumerate(tanda, 1):
        print(f"\n[{i}/{len(tanda)}] {contacto['nombre']} — {contacto['patente']}")
        try:
            enviar_mensaje(contacto)
            marcar_enviado(ws, contacto, valor=1)
            enviados += 1
        except Exception as e:
            print(f"   ⚠️  Error al enviar: {e} — marcando como no contactado (0)")
            marcar_enviado(ws, contacto, valor=0)
            errores += 1

        if i < len(tanda):
            delay = random.uniform(DELAY_MIN, DELAY_MAX)
            print(f"   ⏳ Esperando {delay:.1f}s...")
            time.sleep(delay)

    print("\n💾 Guardando cambios...")
    try:
        wb.save(RUTA_EXCEL)
        print("✅ Base actualizada.")
    except PermissionError:
        print("⚠️  No pude guardar el Excel (¿está abierto?).")
        mostrar_alerta_error("No pude guardar el Excel.\n\n¿Está abierto?")

    print(f"\n{'='*55}")
    print(f"  ✅ Enviados:  {enviados}")
    if errores:
        print(f"  ⚠️  Con error: {errores}")
    print(f"{'='*55}\n")

    mostrar_alerta_final(enviados, errores)


if __name__ == "__main__":
    main()