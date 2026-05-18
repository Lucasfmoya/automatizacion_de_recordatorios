# 📱 Lubricentro O'Higgins — Guía de instalación y uso

---

## ¿Qué hace este script?

Lee tu base de datos Excel, busca todos los clientes que **aún no fueron notificados**
(columna "Num de contactos" vacía o en 0), toma los primeros 10, y les manda el mensaje
personalizado con la patente y los kilómetros via WhatsApp Desktop.

Al terminar **actualiza automáticamente** tu Excel: pone `1` en "Num de contactos" y la
fecha de hoy en la columna de fecha.

---

## PASO 1 — Instalar Python

1. Entrá a https://www.python.org/downloads/
2. Bajá la última versión (botón amarillo grande "Download Python 3.x.x")
3. Ejecutá el instalador
4. **MUY IMPORTANTE**: tildá la opción **"Add Python to PATH"** antes de instalar
5. Finalizá la instalación normalmente

---

## PASO 2 — Instalar las librerías necesarias

1. Abrí el **Símbolo del sistema** (buscá "cmd" en el menú Inicio)
2. Pegá este comando y presioná Enter:

```
pip install openpyxl pyautogui
```

Vas a ver texto descargándose. Cuando vuelva el cursor, listo.

---

## PASO 3 — Configurar el script

Abrí el archivo `lubricentro_whatsapp.py` con el **Bloc de notas** o con
**Visual Studio Code** (que ya tenés instalado, lo vi en tu escritorio).

Buscá la sección que dice `# CONFIGURACIÓN` y editá estas líneas:

### 3.1 — Ruta de tu Excel

```python
RUTA_EXCEL = r"C:\Ruta\A\Tu\BASE_2024.xlsx"
```

Cambiá por la ruta real de tu archivo. Por ejemplo:

```python
RUTA_EXCEL = r"C:\Users\TuPC\Desktop\Proximo_cambio_BASE.xlsx"
```

> ⚠️ **Importante**: tu archivo actual es `.xls` (formato viejo). El script necesita
> que lo guardes como `.xlsx`:
> - Abrilo en Excel
> - Archivo → Guardar como → Tipo: **Libro de Excel (.xlsx)**
> - Guardalo con ese nuevo nombre

### 3.2 — Nombre de la hoja

```python
NOMBRE_HOJA = "BASE 2024"
```

Verificá que coincida exactamente con el nombre de la pestaña en tu Excel.

### 3.3 — Nombres de columnas

Si alguna columna tiene un nombre distinto al que aparece en el script, actualizalo.
Abrí tu Excel y fijate los encabezados exactos de fila 1.

Los que más probablemente necesites verificar:

| Variable en el script | Columna en tu Excel |
|---|---|
| `COL_TELEFONO_LIMPIO` | La columna con el número limpio (sin "Cel:") |
| `COL_NUM_CONTACTOS` | La columna donde ponés 1 cuando notificás |
| `COL_FECHA_ENVIO` | La columna donde guardás la fecha de envío |

### 3.4 — Cantidad de mensajes por tanda

```python
CANTIDAD_POR_TANDA = 10
```

Cambiá el número si querés mandar más o menos por día.

---

## PASO 4 — Convertir el Excel a .xlsx

Como mencioné arriba, el script trabaja con `.xlsx`. Hacelo una sola vez:

1. Abrí tu `Proximo_cambio_BASE_al_31-03-2024.xls` en Excel
2. **Archivo → Guardar como**
3. Elegí la carpeta donde lo querés tener
4. En "Tipo", seleccioná **Libro de Excel (.xlsx)**
5. Guardalo — de ahora en más usás ese archivo

---

## PASO 5 — Usar el script cada vez que quieras mandar mensajes

### Antes de correrlo:
- ✅ Abrí **WhatsApp Desktop** y asegurate de tener sesión iniciada
- ✅ **Cerrá** el archivo Excel (si está abierto en Excel, el script no puede actualizarlo)
- ✅ Tenés conexión a internet

### Cómo correrlo:

**Opción A — Doble click** (más fácil):
1. Hacé click derecho sobre `lubricentro_whatsapp.py`
2. Elegí "Abrir con" → Python

**Opción B — Desde cmd** (si la opción A no funciona):
1. Abrí el Símbolo del sistema
2. Navegá a la carpeta donde está el script:
   ```
   cd C:\Users\TuPC\Desktop
   ```
3. Ejecutá:
   ```
   python lubricentro_whatsapp.py
   ```

### Lo que vas a ver:

```
=======================================================
  LUBRICENTRO O'HIGGINS — Envío automático WhatsApp
=======================================================

📂 Cargando base de datos...
✅ Base cargada: BASE 2024

🔍 Buscando contactos pendientes...
📋 Encontré 47 pendientes — enviando tanda de 10 hoy

   1. JAVIER PEREZ            | OEY771 | 232.167 km
   2. AUGUSTO ALONSO          | AA415JF | 111.079 km
   ...

¿Confirmar envío de 10 mensajes? (s/n):
```

Escribís `s` y Enter. El script hace el resto solo.

---

## ¿Qué hace exactamente con cada mensaje?

1. Toma el número de teléfono de la fila
2. Abre en el navegador: `https://wa.me/549XXXXXXXXXX?text=mensaje_personalizado`
3. WhatsApp Desktop se abre en ese chat con el mensaje ya escrito
4. El script presiona Enter automáticamente para enviarlo
5. Espera entre 4 y 6 segundos (aleatorio) antes del siguiente
6. Marca la fila en el Excel: `1` en "Num de contactos" + fecha de hoy

---

## Preguntas frecuentes

**¿El script puede ejecutarse con el Excel abierto?**
Puede leerlo sí, pero no puede guardarlo. Para que actualice el Excel,
cerralo antes de correr el script.

**¿Qué pasa si un número es incorrecto?**
WhatsApp Desktop abre pero no encuentra el contacto. El script sigue con el siguiente.
Ese contacto igual queda marcado como enviado — si no querés eso, podés volver a poner
0 en su columna "Num de contactos".

**¿Cuántos mensajes por día es seguro enviar?**
Con 10 por día tenés bajo riesgo. Si querés subir a 20–30, aumentá también los delays
a 8–12 segundos. Nunca mandar más de 50 por día para evitar bloqueos.

**¿El script funciona si no tengo WhatsApp Desktop?**
No, necesita WhatsApp Desktop instalado en la PC. Descargalo de:
https://www.whatsapp.com/download

---

## Estructura de columnas esperada en el Excel

| Columna | Nombre | Uso |
|---|---|---|
| A | Cliente | Nombre del cliente |
| B | Teléfono | Con formato "Cel:XXXXXXXXXX" |
| C | Contacto | Número limpio (solo dígitos) |
| D | Patente | Dominio del vehículo |
| F | Auto | Marca/modelo |
| J | Kmts Prox. Cambio | Kilómetros del próximo service |
| K | Prox. Cambio | Fecha estimada |
| L | Num de (contactos) | 0 = pendiente, 1 = notificado |
| N | Repuesta 1 / Fecha envío | Fecha en que se envió el mensaje |
