"""
# Extracción Automatizada de Datos de Formatos de Legalización de Transporte mediante Visión por Computador y OCR

**Participantes**

- María Liliana Alzate Hincapié 
- Adrián Arango Salas

---

## 1. Introducción

Los formatos de legalización de gastos de transporte se diligencian y se archivan, en muchos
casos, como PDF escaneados. Esto obliga a transcribir a mano, fila por fila y celda por celda, la
información que luego se necesita para la contabilidad o el análisis: fecha, descripción del
viaje, centro responsable, proyecto, beneficiario, número de tiquete o factura y valor. Es un
proceso lento y propenso a errores de digitación.

El objetivo de este trabajo es **automatizar la extracción de esos datos** y convertirlos en una
tabla estructurada (CSV y Excel) lista para ser analizada. Para lograrlo, el programa combina:

- **Visión por computador (OpenCV):** endereza la imagen escaneada, detecta las líneas de la
  tabla y recorta cada celda de forma individual.
- **Reconocimiento óptico de caracteres (Tesseract OCR):** lee el texto de cada celda con una
  configuración distinta según el tipo de dato (números, fechas, valores monetarios o texto libre).
- **Lectura directa de texto (pdfplumber):** si el PDF es digital y no escaneado, extrae las
  tablas sin necesidad de OCR.
- **Procesamiento de datos (pandas):** organiza la información, selecciona las columnas de
  interés y exporta los resultados.

El programa detecta por sí mismo si el PDF es escaneado o digital y aplica el método adecuado.

### Tecnologías utilizadas

| Librería | Uso |
|---|---|
| `pypdf` | Leer el PDF y extraer las imágenes de las páginas escaneadas |
| `pdfplumber` | Extraer tablas de PDFs digitales |
| `opencv-python` (`cv2`) | Corrección de inclinación, detección de la cuadrícula y limpieza de celdas |
| `numpy` | Manejo de imágenes como matrices |
| `pytesseract` | Interfaz de Python con el motor OCR Tesseract |
| `pandas` / `openpyxl` | Construcción de la tabla y exportación a CSV y Excel |

## 2. Resumen del flujo

    PDF de entrada
         │
         ▼
    ¿Tiene texto seleccionable?
         │
         ├── Sí ──► extract_from_digital_pdf() ──────────────┐
         │                                                    │
         └── No (escaneado)                                   │
                │                                             │
                ▼                                             │
       Enderezar y deskew (OpenCV)                            │
                │                                             │
                ▼                                             │
       Detectar filas y columnas                              │
                │                                             │
                ▼                                             │
       Aislar texto de cada celda                             │
                │                                             │
                ▼                                             │
       OCR con Tesseract + limpieza                           │
                │                                             │
                └─────────────────────────────────────────────┤
                                                              ▼
                                           DataFrame con 7 columnas objetivo
                                                              │
                                                              ▼
                                          legalizacion_datos.csv / .xlsx

## 3. Resultado esperado

Una tabla con una fila por cada gasto de transporte legalizado y las columnas **Fecha,
Descripción del viaje, Centro responsable, Proyecto, Nombre del beneficiario, No. de
tiquete/factura y Valor**, disponible en CSV y Excel para su revisión y análisis posterior.

## 4. Requisitos de ejecución

- Python 3 con las librerías: `opencv-python`, `numpy`, `pandas`, `openpyxl`, `pypdf`,
  `pdfplumber` y `pytesseract`.
- Tesseract OCR instalado en Windows, con los paquetes de idioma español e inglés.
- El archivo PDF en la misma carpeta del script.

## 5. Descripción del proceso paso a paso

Cada paso se describe en un bloque de comentarios (`# ...`) justo antes del código que lo
implementa. El orden de ejecución es el de los pasos 1 a 11.
"""

import os
import re
import cv2
import numpy as np
import pandas as pd
from pypdf import PdfReader
import pdfplumber
import pytesseract

# ===========================================================================
# ## Paso 1. Configuración del entorno (Tesseract)
#
# Al ejecutar el programa, lo primero que hace es **buscar el ejecutable de
# Tesseract** en las rutas habituales de instalación en Windows
# (`Program Files`, `Program Files (x86)` y la carpeta local del usuario).
# Cuando encuentra `tesseract.exe`, se lo indica a `pytesseract`. Así el
# usuario no tiene que configurar la ruta manualmente.
# ===========================================================================
possible_tesseract_paths = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    os.path.expanduser(r"~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"),
]

for p in possible_tesseract_paths:
    if os.path.isfile(p):
        pytesseract.pytesseract.tesseract_cmd = p
        break

# ===========================================================================
# ## Paso 2. Definición de la estructura del formulario
#
# El programa carga tres constantes que describen el formato:
#
# - **`COLUMN_NAMES`:** las 13 columnas del formulario, de izquierda a derecha
#   (No. de recibo, cuenta contable, nombre de la cuenta, tercero, nombre del
#   tercero, área, centro responsable, proyecto, descripción del viaje, nombre
#   del beneficiario, fecha, No. de tiquete/factura y valor).
# - **`TARGET_COLUMNS`:** las 7 columnas que realmente se exportan: `Fecha`,
#   `Descripcion_Viaje`, `Centro_Responsable`, `Proyecto`,
#   `Nombre_Beneficiario`, `No_Tiquete_Factura` y `Valor`.
# - **`DESKEWED_COL_RATIOS`:** posiciones relativas (proporciones del ancho de
#   la página) de las líneas verticales de la tabla. Sirven como **plan de
#   respaldo** si la detección automática de columnas falla.
# ===========================================================================

# Full 13-column layout of the form (Left to Right)
COLUMN_NAMES = [
    "No_Recibo",             # Col 0
    "Cuenta_Contable",       # Col 1
    "Nombre_Cuenta",         # Col 2
    "Tercero",               # Col 3
    "Nombre_Tercero",        # Col 4
    "Area",                  # Col 5
    "Centro_Responsable",    # Col 6 (E)
    "Proyecto",              # Col 7 (F)
    "Descripcion_Viaje",     # Col 8 (C)
    "Nombre_Beneficiario",   # Col 9
    "Fecha",                 # Col 10 (A)
    "No_Tiquete_Factura",    # Col 11
    "Valor"                  # Col 12
]

# Fields requested for export
TARGET_COLUMNS = [
    "Fecha",
    "Descripcion_Viaje",
    "Centro_Responsable",
    "Proyecto",
    "Nombre_Beneficiario",
    "No_Tiquete_Factura",
    "Valor"
]

# Fallback column ratios on a deskewed page
DESKEWED_COL_RATIOS = [
    0.0367, 0.0815, 0.1348, 0.2067, 0.2515, 0.3224, 0.3476,
    0.3879, 0.4715, 0.5994, 0.7633, 0.8258, 0.8703, 0.9270
]


# ===========================================================================
# ## Paso 3. Identificación del tipo de PDF — `is_scanned_pdf()`
#
# Se revisa cada página del PDF buscando texto seleccionable. Si alguna página
# contiene más de 50 caracteres de texto, el archivo se considera **digital**.
# Si no, se considera **escaneado** (solo imágenes). Esta decisión define cuál
# de los dos caminos de extracción se usa más adelante.
# ===========================================================================
def is_scanned_pdf(pdf_path: str) -> bool:
    """Checks if the PDF is an image scan or has native selectable text."""
    reader = PdfReader(pdf_path)
    for page in reader.pages:
        text = page.extract_text()
        if text and len(text.strip()) > 50:
            return False
    return True


# ===========================================================================
# ## Paso 4. Enderezado de la imagen — `ensure_landscape_and_deskew()`
#
# Solo para PDFs escaneados. Esta función prepara la imagen para que la tabla
# quede perfectamente horizontal:
#
# 1. **Orientación:** si la imagen es más alta que ancha, la gira 90° para
#    dejarla en formato horizontal.
# 2. **Binarización:** convierte la imagen a escala de grises y luego a blanco
#    y negro invertido para resaltar las líneas.
# 3. **Detección de líneas horizontales:** con una operación morfológica aísla
#    las líneas largas de la tabla y con la transformada de Hough mide su ángulo.
# 4. **Corrección:** calcula la mediana de los ángulos (ignorando los mayores
#    a 5°) y rota la imagen en sentido contrario para eliminar la inclinación
#    del escáner.
# ===========================================================================
def ensure_landscape_and_deskew(img: np.ndarray) -> np.ndarray:
    """Rotates sideways scans upright and corrects sub-degree scanner tilt."""
    h, w = img.shape[:2]
    if h > w:
        img = cv2.rotate(img, cv2.ROTATE_90_COUNTERCLOCKWISE)
        h, w = img.shape[:2]

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)

    kernel_h = cv2.getStructuringElement(cv2.MORPH_RECT, (60, 1))
    h_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel_h)

    lines = cv2.HoughLinesP(h_lines, 1, np.pi / 1800, threshold=100, minLineLength=400, maxLineGap=50)
    angles = []
    if lines is not None:
        for x1, y1, x2, y2 in np.array(lines).reshape(-1, 4):
            if abs(x2 - x1) > 400:
                angle = np.degrees(np.arctan2(y2 - y1, x2 - x1))
                if abs(angle) < 5.0:
                    angles.append(angle)

    if angles:
        median_angle = float(np.median(angles))
        M = cv2.getRotationMatrix2D((w // 2, h // 2), median_angle, 1.0)
        img = cv2.warpAffine(img, M, (w, h), flags=cv2.INTER_CUBIC, borderValue=(255, 255, 255))

    return img


# ===========================================================================
# ## Paso 5. Detección de la cuadrícula de la tabla — `detect_table_grid()`
#
# Sobre la imagen ya enderezada, el programa localiza dónde están las filas y
# las columnas:
#
# 1. **Filas:** detecta las líneas horizontales que cruzan casi todo el ancho
#    de la tabla (>= 80 % de la proyección máxima) y fusiona las que están a
#    menos de 80 px, para obtener las coordenadas Y de cada fila. La primera
#    fila corresponde al encabezado.
# 2. **Columnas:** dentro del alto de la tabla detecta las líneas verticales y
#    fusiona las que están a menos de 40 px, para obtener las coordenadas X de
#    cada columna.
# 3. **Validación:** si no se detectan exactamente 14 líneas verticales
#    (13 columnas), usa las proporciones de respaldo `DESKEWED_COL_RATIOS`.
# ===========================================================================
def detect_table_grid(img: np.ndarray):
    """Finds the exact horizontal row boundaries and vertical column boundaries on the deskewed image."""
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)

    # 1. Detect full-width horizontal table lines
    kernel_h = cv2.getStructuringElement(cv2.MORPH_RECT, (100, 1))
    h_lines = cv2.morphologyEx(thresh[:, int(img.shape[1] * 0.04):int(img.shape[1] * 0.93)], cv2.MORPH_OPEN, kernel_h)
    h_prof = np.sum(h_lines, axis=1)

    # Table lines span the full width of the table (>= 80% of max projection)
    thresh_h = np.max(h_prof) * 0.80
    raw_y = []
    cluster = []
    for y, val in enumerate(h_prof):
        if val >= thresh_h:
            cluster.append(y)
        elif cluster:
            raw_y.append(max(cluster, key=lambda idx: h_prof[idx]))
            cluster = []

    # Merge line edges closer than 80px (header is ~104px tall, data rows are ~170-195px tall)
    row_ys = []
    for y in raw_y:
        if not row_ys or (y - row_ys[-1]) >= 80:
            row_ys.append(y)

    # 2. Detect vertical column lines within the table vertical span
    y_top = row_ys[0] if len(row_ys) >= 2 else int(img.shape[0] * 0.21)
    y_bot = row_ys[-1] if len(row_ys) >= 2 else int(img.shape[0] * 0.77)

    kernel_v = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 80))
    v_lines = cv2.morphologyEx(thresh[y_top:y_bot, :], cv2.MORPH_OPEN, kernel_v)
    v_prof = np.sum(v_lines, axis=0)

    thresh_v = np.max(v_prof) * 0.30
    raw_x = []
    cluster_x = []
    for x, val in enumerate(v_prof):
        if val >= thresh_v:
            cluster_x.append(x)
        elif cluster_x:
            raw_x.append(max(cluster_x, key=lambda idx: v_prof[idx]))
            cluster_x = []

    col_xs = []
    for x in raw_x:
        if not col_xs or (x - col_xs[-1]) >= 40:
            col_xs.append(x)

    if len(col_xs) != 14:
        col_xs = [int(r * img.shape[1]) for r in DESKEWED_COL_RATIOS]

    return gray, row_ys, col_xs


# ===========================================================================
# ## Paso 6. Aislamiento del texto de cada celda — `extract_clean_cell_image()`
#
# Para cada celda, la función recorta solo el texto y descarta bordes y ruido:
#
# 1. Recorta la celda dejando un margen de 10 px para evitar los bordes de la
#    tabla.
# 2. Identifica los **componentes conectados** (cada carácter o trazo) y
#    descarta los que son demasiado pequeños, demasiado anchos o con poca área.
# 3. Elimina las **motas de polvo** aisladas que no tienen otros caracteres
#    cercanos.
# 4. Calcula el rectángulo mínimo que contiene el texto real y recorta la
#    celda a ese tamaño.
# 5. Normaliza el contraste, blanquea el fondo, agrega un borde blanco y
#    **amplía la imagen al doble** para mejorar la precisión del OCR.
#
# Si la celda está vacía, devuelve `None`.
# ===========================================================================
def extract_clean_cell_image(gray_img: np.ndarray, y1: int, y2: int, x1: int, x2: int):
    """Isolates actual text characters inside the cell, removing borders and scanner speckles."""
    cell = gray_img[y1 + 10:y2 - 10, x1 + 10:x2 - 10]
    if cell.size == 0:
        return None

    _, inv = cv2.threshold(cell, 185, 255, cv2.THRESH_BINARY_INV)
    num_labels, _, stats, centroids = cv2.connectedComponentsWithStats(inv, connectivity=8)

    candidates = []
    for i in range(1, num_labels):
        area = stats[i, cv2.CC_STAT_AREA]
        w_cc = stats[i, cv2.CC_STAT_WIDTH]
        h_cc = stats[i, cv2.CC_STAT_HEIGHT]
        if h_cc >= 10 and area >= 15 and w_cc < cell.shape[1] * 0.92:
            candidates.append(i)

    if not candidates:
        return None

    # Filter out isolated dust speckles far from characters
    core_indices = []
    for i in candidates:
        cx_i, cy_i = centroids[i]
        w_i = stats[i, cv2.CC_STAT_WIDTH]
        neighbors = sum(
            1 for j in candidates if j != i
            and abs(centroids[j][1] - cy_i) < 28
            and abs(centroids[j][0] - cx_i) < 80
        )
        if neighbors >= 1 or w_i >= 25:
            core_indices.append(i)

    if not core_indices:
        core_indices = candidates

    cy_min = min(stats[i, cv2.CC_STAT_TOP] for i in core_indices)
    cy_max = max(stats[i, cv2.CC_STAT_TOP] + stats[i, cv2.CC_STAT_HEIGHT] for i in core_indices)
    cx_min = min(stats[i, cv2.CC_STAT_LEFT] for i in core_indices)
    cx_max = max(stats[i, cv2.CC_STAT_LEFT] + stats[i, cv2.CC_STAT_WIDTH] for i in core_indices)

    by1 = max(0, cy_min - 5)
    by2 = min(cell.shape[0], cy_max + 5)
    bx1 = max(0, cx_min - 8)
    bx2 = min(cell.shape[1], cx_max + 8)

    tight_gray = cell[by1:by2, bx1:bx2]
    tight_norm = cv2.normalize(tight_gray, None, 0, 255, cv2.NORM_MINMAX)
    tight_norm[tight_norm > 210] = 255

    # Pad with clean white space and upscale 2x for high OCR accuracy
    padded = cv2.copyMakeBorder(tight_norm, 16, 16, 16, 16, cv2.BORDER_CONSTANT, value=255)
    upscaled = cv2.resize(padded, None, fx=2.0, fy=2.0, interpolation=cv2.INTER_CUBIC)
    return upscaled


# ===========================================================================
# ## Paso 7. Limpieza del texto reconocido — `clean_ocr_text()`
#
# Después del OCR, el texto se normaliza según el tipo de columna:
#
# - Se eliminan símbolos espurios y se unifican los espacios.
# - **Columnas numéricas** (centro responsable, proyecto, área, cuenta
#   contable, tercero): se conservan solo los dígitos.
# - **Fecha:** se busca el patrón `dd/mm/aaaa`; si no aparece, se conservan
#   solo dígitos y barras.
# - **Valor:** se conservan solo dígitos, puntos, comas y el signo `$`.
# - **No. de tiquete/factura:** se recortan signos de puntuación en los
#   extremos.
# ===========================================================================
def clean_ocr_text(txt: str, col_name: str) -> str:
    """Cleans OCR text according to the column's expected data format."""
    txt = re.sub(r"[\|\_~<>\*°`«»\"“”‘’]", "", txt)
    txt = re.sub(r"\s+", " ", txt).strip()

    if col_name in ["Centro_Responsable", "Proyecto", "Area", "Cuenta_Contable", "Tercero"]:
        txt = re.sub(r"[^\d]", "", txt)
    elif col_name == "Fecha":
        m = re.search(r"\d{1,2}/\d{1,2}/\d{4}", txt)
        txt = m.group(0) if m else re.sub(r"[^\d/]", "", txt)
    elif col_name == "Valor":
        txt = re.sub(r"[^\d\.\$,]", "", txt).strip(".,")
    elif col_name == "No_Tiquete_Factura":
        txt = txt.strip(".,:- ")

    return txt


# ===========================================================================
# ## Paso 8. Extracción desde PDF escaneado — `extract_from_scanned_pdf()`
#
# Es el flujo principal para documentos escaneados. Por cada página e imagen
# del PDF:
#
# 1. Decodifica la imagen y aplica el **enderezado** (Paso 4).
# 2. Detecta la **cuadrícula** (Paso 5) y arma los intervalos de las 13
#    columnas.
# 3. Recorre las filas de datos, omitiendo el encabezado y las filas delgadas
#    (< 80 px) como la línea de totales.
# 4. Para cada celda, la aísla (Paso 6) y ejecuta Tesseract con una
#    **configuración específica por campo**:
#    - Campos numéricos: solo dígitos `0-9`.
#    - Fecha: dígitos y `/`.
#    - Valor: dígitos, `.`, `$` y `,`.
#    - Tiquete/factura: línea de texto sin restricción (`--psm 7`).
#    - Campos de texto (nombres, descripción): bloque de texto (`--psm 6`).
#
#    El idioma es español + inglés (`spa+eng`).
# 5. Limpia el texto (Paso 7) y arma un diccionario por fila.
# 6. Descarta filas vacías o encabezados repetidos.
# ===========================================================================
def extract_from_scanned_pdf(pdf_path: str) -> list[dict]:
    reader = PdfReader(pdf_path)
    records = []

    for page in reader.pages:
        if not page.images:
            continue

        for img_obj in page.images:
            nparr = np.frombuffer(img_obj.data, np.uint8)
            img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)

            # 1. Rotate and deskew
            img = ensure_landscape_and_deskew(img)

            # 2. Find exact deskewed grid lines
            gray, row_ys, col_xs = detect_table_grid(img)
            col_intervals = [(col_xs[i], col_xs[i + 1]) for i in range(min(13, len(col_xs) - 1))]

            # 3. Iterate over data rows (skip row_ys[0]..row_ys[1] which is the header)
            for r in range(1, len(row_ys) - 1):
                y1, y2 = row_ys[r], row_ys[r + 1]
                if (y2 - y1) < 80:  # Skip thin summary/total line at the bottom
                    continue

                row_dict = {}
                for c_idx, (x1, x2) in enumerate(col_intervals):
                    col_name = COLUMN_NAMES[c_idx]
                    cell_img = extract_clean_cell_image(gray, y1, y2, x1, x2)

                    if cell_img is None:
                        row_dict[col_name] = ""
                        continue

                    # Configure OCR per field type
                    if col_name in ["Centro_Responsable", "Proyecto", "Area", "Cuenta_Contable"]:
                        cfg = "--psm 7 -c tessedit_char_whitelist=0123456789"
                    elif col_name == "Fecha":
                        cfg = "--psm 7 -c tessedit_char_whitelist=0123456789/"
                    elif col_name == "Valor":
                        cfg = "--psm 7 -c tessedit_char_whitelist=0123456789.$,"
                    elif col_name == "No_Tiquete_Factura":
                        cfg = "--psm 7"
                    else:
                        cfg = "--psm 6"

                    raw_txt = pytesseract.image_to_string(cell_img, lang="spa+eng", config=cfg)
                    row_dict[col_name] = clean_ocr_text(raw_txt, col_name)

                # Skip empty rows or header repetitions
                if not any(row_dict.values()) or "Cuenta Contable" in row_dict.get("Descripcion_Viaje", ""):
                    continue

                records.append(row_dict)

    return records


# ===========================================================================
# ## Paso 9. Extracción desde PDF digital — `extract_from_digital_pdf()`
#
# Si el PDF tiene texto nativo, no se necesita OCR. Con `pdfplumber` se
# extraen las tablas de cada página, se limpian los espacios, se omiten filas
# vacías o de encabezado ("Cuenta Contable") y se asocia cada fila con los
# nombres de columna definidos en el Paso 2.
# ===========================================================================
def extract_from_digital_pdf(pdf_path: str) -> list[dict]:
    records = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                for row in table:
                    clean_row = [re.sub(r"\s+", " ", c).strip() if c else "" for c in row]
                    if any(clean_row) and not any("Cuenta Contable" in c for c in clean_row):
                        if len(clean_row) >= len(COLUMN_NAMES):
                            records.append(dict(zip(COLUMN_NAMES, clean_row[:len(COLUMN_NAMES)])))
    return records


# ===========================================================================
# ## Paso 10. Consolidación y exportación — `run_pipeline()`
#
# Es la función que orquesta todo el proceso:
#
# 1. Verifica que el archivo PDF exista; si no, lanza un error.
# 2. Decide el método de extracción según el tipo de PDF (Paso 3) y lo
#    ejecuta (Paso 8 o Paso 9).
# 3. Si no se obtuvieron filas, informa que no se extrajo ningún registro
#    válido.
# 4. Construye un `DataFrame` de pandas y **conserva únicamente las 7
#    columnas objetivo**.
# 5. Exporta los resultados a:
#    - `legalizacion_datos.csv` (UTF-8 con BOM, para que Excel lea bien las
#      tildes).
#    - `legalizacion_datos.xlsx` (Excel).
# 6. Muestra en consola la cantidad de filas exportadas, las rutas de los
#    archivos y la tabla resultante.
# ===========================================================================
def run_pipeline(pdf_filename: str, output_csv: str = "legalizacion_datos.csv", output_xlsx: str = "legalizacion_datos.xlsx"):
    if not os.path.isfile(pdf_filename):
        raise FileNotFoundError(f"File not found: {pdf_filename}")

    print(f"Processing '{pdf_filename}'...")
    if is_scanned_pdf(pdf_filename):
        print("Detected: Scanned PDF. Running deskew + connected-component cell isolation + OCR...")
        records = extract_from_scanned_pdf(pdf_filename)
    else:
        print("Detected: Digital PDF. Extracting text directly...")
        records = extract_from_digital_pdf(pdf_filename)

    if not records:
        print("No valid rows extracted.")
        return

    df = pd.DataFrame(records)
    selected_cols = [c for c in TARGET_COLUMNS if c in df.columns]
    df_final = df[selected_cols]

    df_final.to_csv(output_csv, index=False, encoding="utf-8-sig")
    df_final.to_excel(output_xlsx, index=False, engine="openpyxl")

    print(f"\nSuccessfully exported {len(df_final)} rows.")
    print(f"CSV:   {os.path.abspath(output_csv)}")
    print(f"Excel: {os.path.abspath(output_xlsx)}\n")
    print(df_final.to_string(index=False))


# ===========================================================================
# ## Paso 11. Punto de entrada
#
# Al ejecutar el script directamente, se procesa el archivo
# `Archivo Legalización transporte.pdf` mediante `run_pipeline()`.
# ===========================================================================
if __name__ == "__main__":
    PDF_FILE = "Archivo Legalización transporte.pdf"
    run_pipeline(PDF_FILE)
