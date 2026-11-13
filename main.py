# -*- coding: utf-8 -*-
"""
GESTOR DE HOJAS EXCEL  –  BASE v0.1
===================================
Primer código base con todas las funcionalidades del diseño (DISENO_SOFTWARE.md).
Estilo y temas tomados de RAIZ.py (PySide6).

Dependencias:  pip install PySide6 openpyxl xlwings requests cryptography

Puntos marcados con  TODO-EXCEL  son los que hay que ajustar cuando
se referencie la plantilla real de Excel (nombres de hoja, filas, columnas).
"""
import sys
import os
import re
import copy
import shutil
import unicodedata
from decimal import Decimal, ROUND_HALF_UP, ROUND_UP, ROUND_DOWN, InvalidOperation
from datetime import datetime
from functools import partial

from PySide6.QtWidgets import (
    QApplication, QWidget, QDialog, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, QFrame, QScrollArea,
    QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
    QMessageBox, QFileDialog, QColorDialog, QSizePolicy, QMenu,
    QStyledItemDelegate, QStyleOptionViewItem, QStyle,
)
from PySide6.QtGui import (
    QColor, QPixmap, QPainter, QTransform, QIcon, QFont, QShortcut,
    QKeySequence, QRegularExpressionValidator, QPalette,
)
from PySide6.QtCore import (
    Qt, QTimer, QThread, Signal, QRegularExpression, QPoint, QRect, QEvent,
    QItemSelectionModel,
)

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side

from license_manager import check_license

# ═════════════════════════════════════════════════════════════════════
#  CONFIGURACIÓN GENERAL  (TODO-EXCEL: ajustar con la plantilla real)
# ═════════════════════════════════════════════════════════════════════
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _carpeta_recursos(meipass, base):
    """
    Carpeta de RECURSOS de solo lectura (logo y plantilla "machote").

    En un .exe onefile de PyInstaller los datos se descomprimen en `_MEIPASS`,
    una carpeta temporal que se BORRA al cerrar el programa; en código fuente
    los recursos viven junto a main.py. Todo lo que solo se lee viene de aquí.
    """
    return meipass or base


def _carpeta_datos(es_frozen, ejecutable, base):
    """
    Carpeta de DATOS que se ESCRIBEN (plantilla editable y backups).

    Nunca puede ser `_MEIPASS`: lo que se escriba ahí se pierde al cerrar, que
    era la causa de que el Guardado Global no llegara a la plantilla real.
    Con PyInstaller se usa la carpeta del .exe; en código fuente, la de main.py.
    """
    if es_frozen:
        return os.path.dirname(os.path.abspath(ejecutable))
    return base


NOMBRE_PLANTILLA = "ANEXO SICOA 2025.xlsm"
DIR_RECURSOS = _carpeta_recursos(getattr(sys, "_MEIPASS", None), BASE_DIR)
DIR_DATOS = _carpeta_datos(getattr(sys, "frozen", False), sys.executable, BASE_DIR)
RUTA_LOGO = os.path.join(DIR_RECURSOS, "Imagenes", "Logo.ico")
# Plantilla empaquetada: SOLO se lee (semilla para crear la editable).
PLANTILLA_ORIGEN = os.path.join(DIR_RECURSOS, "Plantilla", NOMBRE_PLANTILLA)
# Plantilla EDITABLE: es la que se lee al arrancar y la que Guardar→Global modifica.
PLANTILLA_XLSX = os.path.join(DIR_DATOS, "Plantilla", NOMBRE_PLANTILLA)
HOJA_BD = "CODIGOS"       # hoja de la plantilla donde vive la tabla CODIGOS_SICOA
TABLA_BD = "CODIGOS_SICOA"
# Nombres EXACTOS de las columnas de la tabla CODIGOS_SICOA (para BUSCARX).
COL_BD_CODIGO = "Codigo Componente"
COL_BD_DESC = "Descripción del componente"
COL_BD_UNIDAD = "Unidad de medida"
COL_BD_PRECIO = "Precio unitario"
# Alto fijo para descripciones que no caben en el ancho de la plantilla.
ALTO_DESC_LARGA = 45
UMBRAL_DESC_LARGA = 30      # más de 30 caracteres = alto 45
UMBRAL_DESC_AUTO = 110      # más de 110 caracteres = AutoFit real (2+ líneas)
# Alto del "Total en letras": 15 pt por línea (alto nativo de Excel a 11 pt).
# No se usa AutoFit ni se descombina A:G; se cuentan las líneas reales del texto
# con el ancho medido de la combinación. Ver _xl_alto_letras.
LETRAS_ALTO_LINEA = 15      # alto de una línea de texto normal (11 pt)
LETRAS_MAX_CAR = 110        # tope duro de caracteres por línea (criterio de WuMDrex)
LETRAS_MIN_CAR = 40         # piso: un ancho muy angosto no debe inventar líneas
LETRAS_FACTOR_CAR = 0.58    # ancho medio de un carácter / tamaño de fuente
LETRAS_MARGEN_PT = 6        # margen interno de la celda combinada (izq. + der.)
LETRAS_SEGURIDAD = 0.92     # 8 % menos capacidad para no cortar la última palabra
LETRAS_ALTO_MAX = 409       # límite duro del alto de fila en Excel
LETRAS_ANCHO_MIN_COL = 5   # columnas mínimas del combinado del "Total en letras"
HOJA_CTRL = "_CTRL"       # hoja oculta de control (totales anteriores)
MAX_ABREV = 17            # 31 (límite Excel) - 14 (prefijo "ANE. - VAR. - ")
# Tipo de Documento: un código es ATE solo si es numérico y MAYOR a este umbral.
# 2000 exacto NO es ATE. Lo comparten la fórmula de Excel y Estado._tipo_documento.
UMBRAL_CODIGO_ATE = 2000
FMT_MONEDA_XL = '"Q"#,##0.00'

HOJAS = ["ANE", "ANE_VAR", "VARIACION", "OC", "OTS", "ATE"]
ETIQUETA = {"ANE": "ANE", "ANE_VAR": "ANE - VAR", "VARIACION": "VARIACIÓN",
            "OC": "OC", "OTS": "OTS", "ATE": "ATE"}

# Prefijos REALES de hoja en el machote (con punto y espacio).
PREFIJOS_HOJA = {
    "ANE": "ANE. - ",
    "ANE_VAR": "ANE. - VAR. - ",
    "VARIACION": "VARIACIÓN - ",
    "OC": "OC - ",
    "OTS": "OTS - ",
    "ATE": "ATE - ",
}

# Hojas base del machote que se copian para cada entidad.
BASE_COMUNIDAD = {
    "ANE": "ANE. - COMUNIDAD",
    "ANE_VAR": "ANE. - VAR. - COMUNIDAD",
    "VARIACION": "VARIACIÓN - COMUNIDAD",
    "OC": "OC - COMUNIDAD",
    "OTS": "OTS - COMUNIDAD",
    "ATE": "ATE - COMUNIDAD",
}
BASE_RAMAL = {
    "ANE": "ANE. - RAMAL",
    "ANE_VAR": "ANE. - VAR. - RAMAL",
    "VARIACION": "VARIACIÓN - RAMAL",
    "OC": "OC - RAMAL",
    "OTS": "OTS - RAMAL",
    "ATE": "ATE - RAMAL",
}

# Prefijos de tabla dentro de las hojas base.
TABLA_PREFIJO = {
    "ANE": "A",
    "ANE_VAR": "AV",
    "VARIACION": "V",
    "OC": "OC",
    "OTS": "OTS",
    "ATE": "ATE",
}

# Modos de redondeo de cada Precio total (los mismos nombres que usa Excel en español).
# "0.00 SR" = multiplicación SIN redondeo: la celda guarda el producto exacto y
# el formato de moneda de la columna lo muestra con 2 decimales.
MODOS_REDONDEO = [
    ("Redondear", "ROUND"),
    ("0.00 SR", "SR"),
    ("Redondear.menos", "ROUNDDOWN"),
]
# Modos aceptados al LEER un archivo. ROUNDUP ya no se ofrece en el menú, pero se
# conserva para no romper archivos exportados con la versión anterior.
MODOS_VALIDOS = ("ROUND", "SR", "ROUNDDOWN", "ROUNDUP")
# Textos del menú (incluido el viejo "Redondear.más") para leer archivos viejos.
MODO_POR_TEXTO = {"Redondear": "ROUND", "0.00 SR": "SR",
                  "Redondear.más": "ROUNDUP", "Redondear.menos": "ROUNDDOWN"}
MODO_DEFECTO = "ROUND"
DEPENDEN_DE_VAR = ["VARIACION", "OC", "OTS", "ATE"]
EDITABLES = ("ANE", "ANE_VAR")          # hojas donde se capturan componentes
FMT_CANTIDAD_KM_XL = "0.000"            # kilometraje: tres decimales
# Cálculo en MANUAL (xlCalculationManual). Al exportar/importar se escriben
# cientos de fórmulas BUSCARX/SUMAR.SI; con cálculo automático Excel recalcula
# el libro después de CADA escritura y eso domina el tiempo. Todo lo que LEE un
# valor calculado llama antes a `app.calculate()`.
XL_CALCULOMANUAL = -4135
# Tope para leer columnas completas de una vez: UsedRange puede venir inflado
# hasta la última fila de la hoja y leer un millón de filas costaría más que el
# recorrido fila por fila anterior.
XL_MAX_FILAS_LECTURA = 5000
HISTORIAL_MAX = 30          # acciones que se pueden deshacer seguidas

# clave, título de la tabla, texto del combo "Línea"
LINEAS = [("MT", "MEDIA TENSION", "Media tensión"),
          ("MTBT", "MEDIA Y BAJA TENSION", "Media más baja tensión"),
          ("BT", "BAJA TENSION", "Baja tensión"),
          ("G", "GENERAL", "General")]
VOLTAJES = [("Línea 7.6 kV", "v76"), ("Línea 19.9 kV", "v199")]
PALETA = ["#3498db", "#27ae60", "#e67e22", "#9b59b6", "#16a085",
          "#e74c3c", "#f39c12", "#2980b9", "#8e44ad", "#d35400"]




# ═════════════════════════════════════════════════════════════════════
#  UTILIDADES
# ═════════════════════════════════════════════════════════════════════
def norm(s):
    """minúsculas y sin acentos (para buscar)."""
    s = unicodedata.normalize("NFD", str(s))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower().strip()


def _nivel_coincidencia(valor, texto, campo):
    """
    Prioridad de una coincidencia de búsqueda.

    0 = exacta, 1 = empieza con lo escrito, 2 = alguna palabra empieza con
    lo escrito (solo descripción), 3 = contiene lo escrito, 99 = no coincide.
    """
    if not texto:
        return 0
    v = norm(valor)
    if v == texto:
        return 0
    if v.startswith(texto):
        return 1
    if campo == "descripcion":
        palabras = re.findall(r"[a-z0-9]+", v)
        if any(p.startswith(texto) for p in palabras):
            return 2
    if texto in v:
        return 3
    return 99


def D(x):
    try:
        return Decimal(str(x).replace(",", "").replace("Q", "").strip() or "0")
    except InvalidOperation:
        return Decimal("0")


def es_km(unidad):
    return norm(unidad) in ("kilometro", "kilometros")


def fmt_cant(c, unidad):
    c = D(c)
    return f"{c:.3f}" if es_km(unidad) else f"{int(c)}"


def fmt_q(x):
    return "Q {:,.2f}".format(D(x).quantize(Decimal("0.01"), ROUND_HALF_UP))


def texto_total(valor, modo=MODO_DEFECTO):
    """Texto de una celda Precio total: sin sufijo, con ' SR', '+' o '-'."""
    txt = fmt_q(valor)
    if modo == "SR":
        return txt + " SR"
    if modo == "ROUNDUP":
        return txt + " +"
    if modo == "ROUNDDOWN":
        return txt + " -"
    return txt


def redondear_monto(x, modo="ROUND"):
    """Aplica el criterio de ROUND / ROUNDDOWN de Excel; SR devuelve el producto sin redondear."""
    q = Decimal("0.01")
    v = D(x)
    if modo == "SR":
        return v                      # producto exacto, sin redondear
    if modo == "ROUNDUP":
        return v.quantize(q, ROUND_UP)
    if modo == "ROUNDDOWN":
        return v.quantize(q, ROUND_DOWN)
    return v.quantize(q, ROUND_HALF_UP)


def contraste(color):
    c = QColor(color)
    return "#000000" if (c.red() * 0.299 + c.green() * 0.587 + c.blue() * 0.114) > 160 else "#ffffff"


def normalizar_abrev(texto, otros):
    """
    Devuelve (nombre, None) o (None, mensaje_de_error).

    La abreviatura se usa TAMBIÉN en nombres de tabla de Excel, así que se
    limita a letras/números/guion_bajo (sin acentos, espacios ni guiones).
    """
    t = unicodedata.normalize("NFKD", texto.strip()).encode("ascii", "ignore").decode()
    t = re.sub(r"[^A-Za-z0-9_]", "_", t)
    t = re.sub(r"_+", "_", t).strip("_")
    if not t:
        return None, "El nombre no puede estar vacío."
    if len(t) > MAX_ABREV:
        return None, f"Máximo {MAX_ABREV} caracteres."
    if t.upper() in {o.upper() for o in otros}:
        return None, "Ya existe un nombre igual. Escriba uno diferente."
    return t, None


def nombre_archivo_anexo(estado):
    """
    Nombre sugerido para el Excel final:
    ANEXO - <abreviatura> - dd.mm.aaaa.xlsx

    Siempre se usa la abreviatura: la del proyecto cuando hay uno solo y la
    del Ramal cuando hay dos o más. El nombre largo no se usa para el archivo.
    """
    if len(estado.proyectos) == 1:
        etiqueta = (estado.proyectos[0].get("abrev") or "COMUNIDAD").strip()
    else:
        etiqueta = (estado.ramal.get("abrev") or "RAMAL").strip()
    etiqueta = re.sub(r'[\\/:*?"<>|]+', "-", etiqueta)
    etiqueta = re.sub(r"-{2,}", "-", etiqueta)
    etiqueta = re.sub(r"\s+", " ", etiqueta).strip(" .-")
    if not etiqueta:
        etiqueta = "ANEXO"
    return f"ANEXO - {etiqueta} - {datetime.now():%d.%m.%Y}.xlsx"


# ── Total en letras ──────────────────────────────────────────────────
_UNI = ["CERO", "UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE",
        "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISÉIS", "DIECISIETE",
        "DIECIOCHO", "DIECINUEVE", "VEINTE", "VEINTIÚN", "VEINTIDÓS", "VEINTITRÉS",
        "VEINTICUATRO", "VEINTICINCO", "VEINTISÉIS", "VEINTISIETE", "VEINTIOCHO", "VEINTINUEVE"]
_DEC = {3: "TREINTA", 4: "CUARENTA", 5: "CINCUENTA", 6: "SESENTA", 7: "SETENTA",
        8: "OCHENTA", 9: "NOVENTA"}
_CEN = {1: "CIENTO", 2: "DOSCIENTOS", 3: "TRESCIENTOS", 4: "CUATROCIENTOS", 5: "QUINIENTOS",
        6: "SEISCIENTOS", 7: "SETECIENTOS", 8: "OCHOCIENTOS", 9: "NOVECIENTOS"}


def _menor_mil(n):
    if n == 100:
        return "CIEN"
    c, r = divmod(n, 100)
    partes = []
    if c:
        partes.append(_CEN[c])
    if r:
        if r < 30:
            partes.append(_UNI[r])
        else:
            dd, u = divmod(r, 10)
            partes.append(_DEC[dd] + (" Y " + _UNI[u] if u else ""))
    return " ".join(partes)


def entero_a_letras(n):
    if n == 0:
        return "CERO"
    partes = []
    mill, resto = divmod(n, 1_000_000)
    if mill:
        partes.append("UN MILLÓN" if mill == 1 else entero_a_letras(mill) + " MILLONES")
    miles, resto = divmod(resto, 1000)
    if miles:
        partes.append("MIL" if miles == 1 else _menor_mil(miles) + " MIL")
    if resto:
        partes.append(_menor_mil(resto))
    return " ".join(partes)


def total_en_letras(total):
    total = D(total).quantize(Decimal("0.01"), ROUND_HALF_UP)
    if total < 0:                       # p. ej. hoja VARIACION con diferencia negativa
        return "MENOS " + total_en_letras(-total)
    ent = int(total)
    cent = int((total - ent) * 100)
    if ent == 1:
        unidad = "QUETZAL"
    elif ent > 0 and ent % 1_000_000 == 0:
        unidad = "DE QUETZALES"
    else:
        unidad = "QUETZALES"
    c_un = "CENTAVO" if cent == 1 else "CENTAVOS"
    return f"{entero_a_letras(ent)} {unidad} CON {entero_a_letras(cent)} {c_un} DE QUETZAL"


# ═════════════════════════════════════════════════════════════════════
#  BASE DE DATOS DE COMPONENTES (vive en la hoja BD de la plantilla)
#  Columnas: A Código | B Descripción | C Unidad | D Precio |
#            E MT | F MT+BT | G BT | H G | I 7.6 kV | J 19.9 kV     (1 / 0)
#  TODO-EXCEL: ajustar columnas/hoja a la plantilla real
# ═════════════════════════════════════════════════════════════════════
BD_COLS = ["codigo", "descripcion", "unidad", "precio", "MT", "MTBT", "BT", "G", "v76", "v199"]
BD_ENC = ["CODIGO COMPONENTE", "Descripción del componente", "Unidad", "Precio Unitario",
          "MT", "MT + BT", "BT", "G", "7.6 kV", "19.9 kV"]


def _cod(v):
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def clasificar_componente(codigo, descripcion):
    """
    Deduce en qué tablas del ANE puede ir un código cuando la tabla
    CODIGOS_SICOA solo trae las 4 columnas oficiales.

    Es una clasificación INICIAL; el usuario puede ajustarla en la ventana
    Componentes (esas marcas se guardan en el JSON del proyecto, no en Excel).
    """
    cod = _cod(codigo).upper()
    m = re.match(r"([A-Z]+)", cod)
    pref = m.group(1) if m else ""
    d = norm(descripcion)
    f = {"MT": 0, "MTBT": 0, "BT": 0, "G": 0, "v76": 0, "v199": 0}

    if pref in ("P", "B", "S", "R"):
        # Postes, fundaciones, retenciones y excavación sirven en cualquier línea.
        f.update(MT=1, MTBT=1, BT=1)
    elif pref == "E":
        # Estructuras de distribución primaria.
        f.update(MT=1, MTBT=1)
        if "baja tension" in d:
            f["BT"] = 1
    elif pref == "C":
        # Conductores: la descripción dice si son MT, MT+BT o BT.
        if "media y baja" in d:
            f["MTBT"] = 1
        elif "media tension" in d:
            f["MT"] = 1
        elif "baja tension" in d:
            f["BT"] = 1
        else:
            f.update(MT=1, MTBT=1, BT=1)
    elif pref == "T":
        # Transformadores: van en media y media+baja; el voltaje se deduce del texto.
        f.update(MT=1, MTBT=1)
    elif pref == "M":
        # Planos y rótulos son GENERAL; el resto es material de línea.
        if "plano" in d or "rotulo" in d or "rotulacion" in d:
            f["G"] = 1
        else:
            f.update(MT=1, MTBT=1, BT=1)
    else:
        f.update(MT=1, MTBT=1, BT=1)

    # Alias de voltaje usados en las descripciones:
    # 7.6 kV y 13.2 kV son la misma línea; 19.9 kV y 34.5 kV también.
    # Solo se aplica a materiales de media tensión; si el texto trae ambos
    # voltajes se deja sin restricción (v76 = v199 = 0).
    if f["MT"] or f["MTBT"]:
        d76 = any(x in d for x in ("7.6", "7,6", "13.2", "13,2"))
        d199 = any(x in d for x in ("19.9", "19,9", "34.5", "34,5"))
        if d76 and not d199:
            f["v76"] = 1
        elif d199 and not d76:
            f["v199"] = 1
    return f


def componente_nuevo(codigo, descripcion, unidad, precio):
    """
    Crea un componente NUEVO con las seis marcas de línea/voltaje en 0.

    A diferencia de `clasificar_componente` (que deduce las líneas desde la
    descripción para las bases de 4 columnas), un código nuevo debe nacer SIN
    marca: el usuario decide en qué líneas va y guarda esa decisión. Así un
    nombre que menciona un voltaje no deja el código marcado en una línea
    equivocada antes de que el usuario lo revise.
    """
    return {"codigo": _cod(codigo),
            "descripcion": str(descripcion or "").strip(),
            "unidad": str(unidad or "").strip(),
            "precio": str(D(precio).quantize(Decimal("0.01"))),
            "MT": 0, "MTBT": 0, "BT": 0, "G": 0, "v76": 0, "v199": 0}


def bd_leer_ws(ws):
    """
    Lee la tabla CODIGOS_SICOA (4 columnas oficiales) o una hoja BD antigua
    de 10 columnas (código, descripción, unidad, precio + marcas).
    """
    bd = []
    encabezado = [norm(c.value) for c in ws[1]]
    formato_viejo = any(x in encabezado for x in ("mt", "mt + bt", "7.6 kv", "19.9 kv"))

    for fila in ws.iter_rows(min_row=2, values_only=True):
        if not fila or fila[0] in (None, ""):
            continue
        fila = list(fila) + [None] * (10 - len(fila))
        c = dict(codigo=_cod(fila[0]), descripcion=str(fila[1] or ""), unidad=str(fila[2] or ""),
                 precio=str(D(fila[3])))
        if formato_viejo:
            for k, v in zip(BD_COLS[4:], fila[4:10]):
                c[k] = 1 if str(v).strip() in ("1", "True", "true") else 0
            if c["v76"] and c["v199"]:       # nunca ambos
                c["v76"] = c["v199"] = 0
        else:
            # CODIGOS_SICOA: la clasificación se deduce y luego el usuario
            # puede afinarla desde la ventana Componentes.
            c.update(clasificar_componente(c["codigo"], c["descripcion"]))
        bd.append(c)
    return bd


def _bd_dict_de_fila(fila, con_marcas):
    """
    Convierte una fila COM de CODIGOS_SICOA en el dict interno de la BD.

    · con_marcas=True  → la tabla trae las 6 columnas de clasificación y se
      respetan tal cual (es la decisión que el usuario guardó).
    · con_marcas=False → plantilla vieja de 4 columnas: se deduce con
      `clasificar_componente`, igual que al leer el .xlsm con openpyxl.

    COM devuelve los números como float (1.0) y los booleanos como True, así
    que la marca acepta 1, 1.0, True y "true"; cualquier otra cosa es 0.
    """
    fila = list(fila) + [None] * (10 - len(fila))
    c = dict(codigo=_cod(fila[0]), descripcion=str(fila[1] or ""),
             unidad=str(fila[2] or ""), precio=str(D(fila[3])))
    if con_marcas:
        for k, v in zip(BD_COLS[4:], fila[4:10]):
            try:
                c[k] = 1 if float(v) else 0
            except (TypeError, ValueError):
                c[k] = 1 if str(v).strip().lower() in ("true", "verdadero") else 0
        if c["v76"] and c["v199"]:       # nunca ambos
            c["v76"] = c["v199"] = 0
    else:
        c.update(clasificar_componente(c["codigo"], c["descripcion"]))
    return c


def bd_escribir_ws(ws, bd):
    ws.delete_rows(1, ws.max_row + 1)
    ws.append(BD_ENC)
    for c in ws[1]:
        c.font = Font(bold=True)
    for c in bd:
        ws.append([c["codigo"], c["descripcion"], c["unidad"], float(D(c["precio"]))] +
                  [int(c[k]) for k in BD_COLS[4:]])
    for fila in ws.iter_rows(min_row=2, min_col=4, max_col=4):
        fila[0].number_format = FMT_MONEDA_XL
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 50
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 16


# Caché de la BD de la plantilla. `cargar_bd_inicial` se llama en Estado(),
# desde_dict y nuevo_proyecto; abrir el .xlsm con openpyxl cada vez cuesta lo
# suyo. La clave lleva ruta + mtime + tamaño, así que si Guardar→Global
# reescribe la plantilla se relee sola.
_BD_CACHE = {"clave": None, "bd": None}


def cargar_bd_inicial():
    try:
        # Estado() puede construirse sin pasar por main(); se asegura aquí que
        # la plantilla editable exista junto al .exe antes de leerla.
        preparar_plantilla_editable()
        if not os.path.exists(PLANTILLA_XLSX):
            print("No existe la plantilla editable:", PLANTILLA_XLSX)
            return []
        try:
            st = os.stat(PLANTILLA_XLSX)
            clave = (os.path.abspath(PLANTILLA_XLSX), st.st_mtime_ns, st.st_size)
        except OSError:
            clave = None
        if clave is not None and _BD_CACHE["clave"] == clave and _BD_CACHE["bd"]:
            # Copia superficial de cada dict: nadie debe mutilar el caché con
            # las marcas/descripciones que edita la ventana Componentes.
            return [dict(c) for c in _BD_CACHE["bd"]]
        wb = load_workbook(PLANTILLA_XLSX, read_only=True, data_only=True)
        try:
            bd = []
            if HOJA_BD in wb.sheetnames:
                bd = bd_leer_ws(wb[HOJA_BD])
        finally:
            wb.close()
        if bd:
            _BD_CACHE["clave"] = clave
            _BD_CACHE["bd"] = bd
            return [dict(c) for c in bd]
    except Exception as e:                       # plantilla dañada o abierta
        print("No se pudo leer la BD de la plantilla:", e)
    return []


def _xl_leer_bd(wb):
    """
    Lee la tabla CODIGOS_SICOA del libro ABIERTO, no de la plantilla.

    Un anexo guardado en Local lleva sus propios códigos, que la plantilla
    puede no tener todavía; por eso al abrir un Excel se lee primero esta
    tabla. Si el libro no la trae (o no tiene las seis columnas), se avisa con
    una lista vacía para que el llamador caiga a la plantilla.
    """
    try:
        ws = wb.sheets[HOJA_BD]
    except Exception:
        return []
    tabla = _xl_tabla(ws, TABLA_BD)
    if tabla is None or tabla.api.DataBodyRange is None:
        return []
    etiquetas = set()
    try:
        for i in range(1, int(tabla.api.ListColumns.Count) + 1):
            etiquetas.add(_xl_normalizar_etiqueta(tabla.api.ListColumns(i).Name))
    except Exception:
        etiquetas = set()
    # Comparación exacta normalizada: "MT" no debe confundirse con "MT + BT".
    con_marcas = all(_xl_normalizar_etiqueta(x) in etiquetas for x in BD_ENC[4:])

    bd = []
    for fila in _xl_valores(tabla.api.DataBodyRange.Value):
        if not fila or fila[0] in (None, ""):
            continue
        c = _bd_dict_de_fila(fila, con_marcas)
        if c["codigo"]:
            bd.append(c)
    return bd


def preparar_plantilla_editable():
    """
    Copia la plantilla empaquetada junto al .exe la PRIMERA vez que arranca.

    PyInstaller onefile descomprime los recursos en `_MEIPASS`, que se borra al
    cerrar: si se escribiera ahí, el Guardado Global se perdería. Por eso la
    plantilla editable vive en DIR_DATOS (junto al .exe) y, si YA existe, NO se
    sobrescribe al recompilar: WuMDrex pudo haber agregado códigos con
    Guardar→Global y recompilar no debe borrarlos.
    """
    try:
        if os.path.exists(PLANTILLA_XLSX):
            return PLANTILLA_XLSX
        if os.path.abspath(PLANTILLA_ORIGEN) == os.path.abspath(PLANTILLA_XLSX):
            return PLANTILLA_XLSX          # código fuente: el mismo archivo
        os.makedirs(os.path.dirname(PLANTILLA_XLSX), exist_ok=True)
        shutil.copyfile(PLANTILLA_ORIGEN, PLANTILLA_XLSX)
    except OSError as e:
        # No se aborta: la app sigue en memoria y el error se reporta al guardar.
        print(f"[ANEXOS] No se pudo crear la plantilla editable en "
              f"{PLANTILLA_XLSX}: {e}")
    return PLANTILLA_XLSX


def _plantilla_bloqueada(ruta=None):
    """
    True si la plantilla está abierta en Excel o protegida contra escritura.

    No se decide por el archivo temporal `~$ANEXO SICOA 2025.xlsm`: queda
    huérfano cuando Excel se cierra de forma anormal. Se usa la prueba real de
    abrir el archivo en modo lectura/escritura.
    """
    ruta = ruta or PLANTILLA_XLSX
    if not os.path.exists(ruta):
        return False
    archivo = None
    try:
        archivo = open(ruta, "r+b")
        return False
    except OSError:
        return True
    finally:
        if archivo is not None:
            archivo.close()


def _respaldo_plantilla():
    """Copia de seguridad de la plantilla en backups/ (junto al .exe)."""
    os.makedirs(os.path.join(DIR_DATOS, "backups"), exist_ok=True)
    ahora = datetime.now()
    destino = os.path.join(DIR_DATOS, "backups",
                           f"Backup_{ahora:%d.%m.%Y}_Plantilla_{ahora:%H%M%S}.xlsm")
    shutil.copyfile(PLANTILLA_XLSX, destino)
    return destino


def _abrir_excel():
    """Abre Excel oculto con xlwings (para conservar dibujos, tablas y macros)."""
    import pythoncom
    import xlwings as xw
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    # Sin esto Excel recalcula el libro tras CADA fórmula escrita y los eventos
    # de tabla disparan trabajo extra; con cálculo manual la exportación deja
    # de irse en recalcular. AutoFillFormulasInLists se apaga para que Excel no
    # propague la fórmula de la columna calculada a todas las filas (lo hace
    # `_xl_escribir_formulas` por tramos).
    try:
        app.api.Calculation = XL_CALCULOMANUAL          # xlCalculationManual
        app.api.EnableEvents = False
        app.api.AutoCorrect.AutoFillFormulasInLists = False
    except Exception:
        pass
    return app, pythoncom


def _cerrar_excel(app, pythoncom):
    try:
        app.quit()
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _xl_valores(valores):
    """
    Normaliza Range.Value de COM a una lista de tuplas (una por fila).

    COM devuelve una tupla de escalares cuando el rango tiene UNA sola fila;
    en ese caso hay que envolverla para no perder las demás columnas.
    """
    if valores is None:
        return []
    if not isinstance(valores, tuple):
        return [(valores,)]
    if not valores:
        return []
    if isinstance(valores[0], tuple):
        return [f if isinstance(f, tuple) else (f,) for f in valores]
    return [valores]


def _xl_asegurar_columnas_clasificacion(tabla):
    """Agrega a CODIGOS_SICOA las 6 columnas de clasificación si faltan."""
    nombres = {str(tabla.api.ListColumns(i).Name).strip()
               for i in range(1, int(tabla.api.ListColumns.Count) + 1)}
    for nombre in BD_ENC[4:]:
        if nombre not in nombres:
            col = tabla.api.ListColumns.Add()
            col.Name = nombre
            nombres.add(nombre)


def bd_guardar_plantilla(bd):
    """
    Guardado GLOBAL: actualiza la tabla CODIGOS_SICOA de la plantilla.

    · Solo agrega códigos nuevos y actualiza descripción/unidad/precio.
    · Asegura las columnas de clasificación (MT, MT + BT, BT, G, 7.6, 19.9)
      y guarda sus marcas como 1/0.
    · NO borra ni reordena filas: las fórmulas del machote dependen del orden.
    · Conserva la propiedad de tabla y los dibujos/macros del .xlsm.
    · Escribe en la plantilla EDITABLE (junto al .exe), nunca en `_MEIPASS`.
    · Si la plantilla está abierta o en solo lectura, lanza PermissionError con
      la ruta para que la interfaz lo muestre en vez de fallar en silencio.
    """
    preparar_plantilla_editable()
    if not os.path.exists(PLANTILLA_XLSX):
        raise FileNotFoundError(f"No existe la plantilla: {PLANTILLA_XLSX}")
    bloqueada = (f"La plantilla está abierta en Excel o protegida:\n"
                 f"{PLANTILLA_XLSX}\n\nCiérrela y vuelva a intentar.")
    if _plantilla_bloqueada():
        raise PermissionError(bloqueada)
    _respaldo_plantilla()

    app, pythoncom = _abrir_excel()
    wb = None
    try:
        wb = app.books.open(PLANTILLA_XLSX)
        if bool(wb.api.ReadOnly):
            raise PermissionError(bloqueada)
        ws = wb.sheets[HOJA_BD]
        tabla = None
        for t in ws.tables:
            if t.name.upper() == TABLA_BD.upper():
                tabla = t
                break
        if tabla is None:
            raise RuntimeError(f"No se encontró la tabla {TABLA_BD} en la hoja {HOJA_BD}.")

        _xl_asegurar_columnas_clasificacion(tabla)

        existentes = {}                 # código -> (índice en el cuerpo, fila leída)
        if tabla.api.DataBodyRange is not None:
            for i, fila in enumerate(_xl_valores(tabla.api.DataBodyRange.Value), start=1):
                cod = _cod(fila[0]) if fila and fila[0] not in (None, "") else ""
                if cod:
                    existentes[cod] = (i, fila)

        def _fila_igual(actual, valores):
            """True si la fila de Excel ya tiene exactamente los mismos 10 valores."""
            if not actual or len(actual) < 10:
                return False
            try:
                if _cod(actual[0]) != _cod(valores[0]):
                    return False
                if str(actual[1] or "") != str(valores[1] or ""):
                    return False
                if str(actual[2] or "") != str(valores[2] or ""):
                    return False
                if float(actual[3] or 0) != float(valores[3] or 0):
                    return False
                for k in range(4, 10):
                    if int(actual[k] or 0) != int(valores[k] or 0):
                        return False
            except (TypeError, ValueError, InvalidOperation):
                return False
            return True

        fila_inicio = int(tabla.api.DataBodyRange.Row) if tabla.api.DataBodyRange is not None else 2
        nuevas = []
        for c in bd:
            valores = [c["codigo"], c["descripcion"], c["unidad"], float(D(c["precio"]))] + \
                      [int(c.get(k, 0)) for k in BD_COLS[4:]]
            actual = existentes.get(c["codigo"])
            if actual is None:
                nuevas.append(valores)
                continue
            # Escribir SOLO lo que cambió: reescribir las diez columnas de los
            # cientos de códigos iguales costaba cientos de llamadas COM en
            # cada Guardado Global.
            if _fila_igual(actual[1], valores):
                continue
            fila_excel = fila_inicio + actual[0] - 1
            ws.range(fila_excel, 1).resize(1, 10).value = valores

        if nuevas:
            # Los códigos nuevos van al final: se agregan las filas con
            # ListRows.Add y se escriben en UN solo bloque contiguo. Si el
            # cuerpo no quedó del tamaño esperado, se escribe fila por fila.
            inicio = int(tabla.api.ListRows.Count)
            for _ in nuevas:
                tabla.api.ListRows.Add()
            if int(tabla.api.ListRows.Count) == inicio + len(nuevas):
                ws.range(fila_inicio + inicio, 1).resize(len(nuevas), 10).value = nuevas
            else:
                total = int(tabla.api.ListRows.Count)
                for k, valores in enumerate(nuevas):
                    if inicio + k >= total:
                        tabla.api.ListRows.Add()
                        total = int(tabla.api.ListRows.Count)
                    ws.range(fila_inicio + inicio + k, 1).resize(1, 10).value = valores

        # Comprobación de disco: si Excel no escribió (ruta equivocada, solo
        # lectura, OneDrive bloqueando), se avisa en vez de dar éxito falso.
        tam_antes = os.path.getsize(PLANTILLA_XLSX)
        mtime_antes = os.path.getmtime(PLANTILLA_XLSX)
        wb.save()
        if (os.path.getsize(PLANTILLA_XLSX) == tam_antes and
                os.path.getmtime(PLANTILLA_XLSX) == mtime_antes):
            raise RuntimeError("La plantilla no se modificó en disco: " + PLANTILLA_XLSX)
    finally:
        try:
            if wb is not None:
                wb.close()
        finally:
            _cerrar_excel(app, pythoncom)


# ═════════════════════════════════════════════════════════════════════
#  ESTADO DEL PROYECTO (proyectos, ramal, hojas, componentes agregados)
# ═════════════════════════════════════════════════════════════════════
def _vacio():
    return {k: [] for k, _, _ in LINEAS}


class Estado:
    # ── bandera de cambios sin guardar ──────────────────────────────
    # Es una propiedad para que la ventana actualice el título en el momento
    # exacto en que cambia, sin tener que llamar a `actualizar_titulo()` en las
    # ~20 asignaciones repartidas por el archivo.
    _cambios = False
    _al_avisar_cambios = None       # callback opcional que pone `Principal`

    @property
    def cambios(self):
        return self._cambios

    @cambios.setter
    def cambios(self, valor):
        nuevo = bool(valor)
        if nuevo == getattr(self, "_cambios", False):
            return                  # sin cambio real: no se avisa dos veces
        self._cambios = nuevo
        aviso = self._al_avisar_cambios
        if aviso is not None:
            aviso()

    def __init__(self):
        self.bd = cargar_bd_inicial()
        self.reset()

    # ── creación / reinicio ─────────────────────────────────────────
    def _base(self, abrev, color, ramal=False):
        return {"abrev": abrev, "nombre": "", "color": color,
                "hojas": {h: (h == "ANE") for h in HOJAS},
                "visibles": ["ANE"], "items": {"ANE": _vacio()},
                "redondeo_doc": {}, "orden_manual": []}

    def ordenar_por_codigo(self, lista):
        """Ordena por la posición del código en la base; los externos van al final."""
        self._indice_bd()
        indice = self._posiciones
        lista.sort(key=lambda it: (indice.get(it.get("codigo", ""), len(indice)),
                                   str(it.get("codigo", ""))))

    def posicion_por_codigo(self, lista, codigo):
        """Posición donde debe entrar un código nuevo según CODIGOS_SICOA.
        No reordena la lista: solo calcula el índice de inserción para no
        mandar al final un material que pertenece a una fila intermedia.
        """
        self._indice_bd()
        indice = self._posiciones
        clave = (indice.get(codigo, len(indice)), str(codigo))
        for i, it in enumerate(lista):
            actual = (indice.get(it.get("codigo", ""), len(indice)),
                      str(it.get("codigo", "")))
            if actual > clave:
                return i
        return len(lista)

    def _indice_bd(self):
        """
        Devuelve {codigo: componente} en caché.

        `comp` se llama por cada fila de cada tabla; buscar en la lista cada
        vez era lineal. El índice se reconstruye si cambió la LISTA (otra
        referencia) o su tamaño; en la misma pasada se arma {codigo: posición}
        para no recorrer la BD al ordenar ni al insertar. Guarda REFERENCIAS a
        los mismos dicts, así que las marcas/descripciones editadas al vuelo se
        ven sin reindexar.
        """
        if (self._bd_ref is not self.bd
                or self._indice is None
                or self._bd_len != len(self.bd)):
            indice, posiciones = {}, {}
            for i, c in enumerate(self.bd):
                # setdefault: con códigos repetidos manda la PRIMERA fila,
                # igual que la búsqueda lineal anterior.
                indice.setdefault(c["codigo"], c)
                posiciones.setdefault(c["codigo"], i)
            self._indice = indice
            self._posiciones = posiciones
            self._bd_ref = self.bd
            self._bd_len = len(self.bd)
        return self._indice

    def reindexar_bd(self):
        """Fuerza a reconstruir el índice (p. ej. tras cambiar un código)."""
        self._indice = None
        self._bd_ref = None
        self._bd_len = 0
        self._indice_bd()

    def reset(self):
        # Caché del índice de la BD ({codigo: componente} y {codigo: posición}).
        self._bd_ref = None
        self._indice = None
        self._posiciones = None
        self._bd_len = 0
        self.bd_local = False
        self.proyectos = [self._base("COMUNIDAD_1", PALETA[0])]
        self.ramal = self._base("RAMAL", "#7f8c8d", True)
        self.ramal_presente = False       # un Excel abierto puede traer el Ramal solo
        self.archivo = None
        self.totales_previos = {}
        self.titulos_borrar = set()      # hojas que hay que quitar del Excel al guardar
        self.cambios = False

    # ── consulta ────────────────────────────────────────────────────
    def ramal_visible(self):
        return self.ramal_presente or len(self.proyectos) >= 2

    def bases(self):
        return self.proyectos + ([self.ramal] if self.ramal_visible() else [])

    def es_ramal(self, base):
        return base is self.ramal

    def titulo(self, base, hoja):
        # La misma abreviatura segura se usa en la hoja y en los nombres de tabla.
        return f"{PREFIJOS_HOJA[hoja]}{_xl_token(base['abrev'])}"

    def tabs(self):
        return [(b, h) for b in self.bases() for h in HOJAS
                if b["hojas"][h] and h in b["visibles"]]

    def otros_nombres(self, base):
        n = [b["abrev"] for b in self.proyectos if b is not base]
        if base is not self.ramal:
            n.append(self.ramal["abrev"])
        return n

    def siguiente_nombre(self):
        usados = {x.upper() for x in self.otros_nombres(None)}
        i = 1
        while f"COMUNIDAD_{i}" in usados:
            i += 1
        return f"COMUNIDAD_{i}"

    def comp(self, codigo):
        return self._indice_bd().get(str(codigo))

    def codigo_en_uso(self, codigo):
        usos = []
        # Se revisan también las filas propias del Ramal: si un componente se
        # usa solo ahí, borrarlo de la BD dejaría la fila apuntando a la nada.
        for b in self.bases():
            for h, lineas in b["items"].items():
                if any(it["codigo"] == codigo for lst in lineas.values() for it in lst):
                    usos.append(self.titulo(b, h))
        return usos

    # ── modificaciones ──────────────────────────────────────────────
    def _borrar_titulos(self, base):
        for h in HOJAS:
            if base["hojas"][h]:
                self.titulos_borrar.add(self.titulo(base, h))

    def agregar_proyecto(self):
        nombre = self.siguiente_nombre()
        b = self._base(nombre, PALETA[len(self.proyectos) % len(PALETA)])
        self.proyectos.append(b)
        self.cambios = True
        if len(self.proyectos) == 2:           # el Ramal aparece justo ahora
            self._sembrar_ramal()
        self._sincronizar_ramal()
        return b

    def eliminar_proyecto(self, base):
        self._borrar_titulos(base)
        if len(self.proyectos) == 1:                       # único: solo se reinicia
            if self.ramal_visible():
                self._borrar_titulos(self.ramal)
            self.proyectos = [self._base("COMUNIDAD_1", PALETA[0])]
            self.ramal = self._base("RAMAL", "#7f8c8d", True)
            self.ramal_presente = False
        else:
            self.proyectos.remove(base)
            if len(self.proyectos) < 2:                    # el Ramal deja de existir
                self._borrar_titulos(self.ramal)
                self.ramal = self._base("RAMAL", "#7f8c8d", True)
                self.ramal_presente = False
        self.cambios = True
        self._sincronizar_ramal()

    def renombrar(self, base, nuevo):
        self._borrar_titulos(base)           # los nombres viejos se quitan del Excel
        base["abrev"] = nuevo
        self.cambios = True

    def activar_hoja(self, base, hoja):
        base["hojas"][hoja] = True
        if hoja not in base["visibles"]:
            base["visibles"].append(hoja)
        if hoja in EDITABLES and hoja not in base["items"]:
            base["items"][hoja] = _vacio()
        self.cambios = True
        if base is not self.ramal:
            self._sincronizar_ramal()

    def quitar_hoja(self, base, hoja):
        self.titulos_borrar.add(self.titulo(base, hoja))
        base["hojas"][hoja] = False
        if hoja in base["visibles"]:
            base["visibles"].remove(hoja)
        base["items"].pop(hoja, None)
        pref = f"{hoja}:"
        base["orden_manual"] = [c for c in base.get("orden_manual", [])
                                if not c.startswith(pref)]
        self.cambios = True
        self.asegurar_visibles()
        if base is not self.ramal:
            self._sincronizar_ramal()

    def asegurar_visibles(self):
        if not self.tabs():
            # Mostrar la primera hoja que exista; si no hay ninguna, crear ANE.
            for b in self.bases():
                for h in HOJAS:
                    if b["hojas"].get(h):
                        if h not in b["visibles"]:
                            b["visibles"].append(h)
                        if h in EDITABLES:
                            b["items"].setdefault(h, _vacio())
                        return
            b = self.bases()[0]
            b["visibles"].append("ANE")
            b["hojas"]["ANE"] = True
            b["items"].setdefault("ANE", _vacio())

    def _sincronizar_ramal(self):
        """
        El Ramal hereda ANE y ANE-VAR de los proyectos y es dueño de sus hojas
        derivadas (VARIACIÓN, OC, OTS y ATE).

        ANE existe siempre; ANE-VAR existe solo si algún proyecto la tiene.
        Las derivadas necesitan ANE-VAR para calcularse, así que no se copian
        de los proyectos: las activa el usuario en el Ramal y, si el Ramal se
        queda sin ANE-VAR, se apagan solas para no dejar pestañas inválidas.
        """
        if not self.ramal_visible():
            return
        tiene_var = any(p["hojas"].get("ANE_VAR") for p in self.proyectos)
        self.ramal["hojas"]["ANE"] = True
        self.ramal["hojas"]["ANE_VAR"] = tiene_var
        if not tiene_var:
            for h in DEPENDEN_DE_VAR:
                self.ramal["hojas"][h] = False
        # Las hojas nuevas aparecen visibles; las que el usuario ocultó siguen
        # ocultas. Así el Ramal no pierde la configuración de pestañas.
        self.ramal["visibles"] = [h for h in self.ramal.get("visibles", [])
                                  if self.ramal["hojas"].get(h)]
        for h in HOJAS:
            if self.ramal["hojas"].get(h) and h not in self.ramal["visibles"]:
                self.ramal["visibles"].append(h)
        # El Ramal conserva lo capturado en sus hojas editables: las filas
        # propias y los ajustes de redondeo no se pierden al sincronizar.
        for h in EDITABLES:
            if self.ramal["hojas"].get(h):
                self.ramal.setdefault("items", {}).setdefault(h, _vacio())
            else:
                self.ramal.get("items", {}).pop(h, None)
        self.asegurar_visibles()
        self.cambios = True

    def _sembrar_ramal(self):
        """
        Copia UNA sola vez la unión de hojas de los proyectos al Ramal.

        Se llama en la transición de 1 a 2 proyectos: cuando el Ramal aparece
        por primera vez arranca con las hojas que ya traen los proyectos. De
        ahí en adelante el Ramal es dueño de sus derivadas, así que no se
        vuelve a sembrar (no se reactiva lo que el usuario apagó).
        """
        if not self.ramal_visible():
            return
        for h in HOJAS:
            if any(p["hojas"].get(h) for p in self.proyectos):
                self.ramal["hojas"][h] = True
                if h not in self.ramal.get("visibles", []):
                    self.ramal["visibles"].append(h)
        self.ramal["visibles"] = [h for h in self.ramal.get("visibles", [])
                                  if self.ramal["hojas"].get(h)]
        self.cambios = True

    def renombrar_codigo(self, viejo, nuevo):
        # El Ramal puede tener filas propias; si no se renombran, quedan
        # apuntando a un código que ya no existe.
        for b in self.proyectos + [self.ramal]:
            for lineas in b["items"].values():
                for lst in lineas.values():
                    for it in lst:
                        if it["codigo"] == viejo:
                            it["codigo"] = nuevo
        # Las claves del índice son los códigos: hay que rehacerlo.
        self.reindexar_bd()

    # ── cálculo de filas por hoja ───────────────────────────────────
    def _cant_proyectos(self, hoja):
        """Suma de las cantidades de los proyectos para una hoja (sin Ramal)."""
        res = {k: {} for k, _, _ in LINEAS}
        for src in self.proyectos:
            lineas = src["items"].get(hoja)
            if not lineas:
                continue
            for k in res:
                for it in lineas[k]:
                    res[k][it["codigo"]] = res[k].get(it["codigo"], Decimal(0)) + D(it["cantidad"])
        return res

    def _cant_propia(self, hoja):
        """
        Suma de las filas capturadas EN el Ramal, por línea y código.

        Ojo: en `items` del Ramal viven dos cosas distintas. Las filas propias
        llevan su cantidad capturada; los ajustes de redondeo de una fila
        heredada viven con cantidad 0. Un ajuste con cantidad 0 NO es una fila
        propia: si el código deja de existir en todos los proyectos no debe
        reaparecer como fila editable ni exportarse.
        """
        res = {k: {} for k, _, _ in LINEAS}
        lineas = (self.ramal.get("items") or {}).get(hoja)
        if not lineas:
            return res
        for k in res:
            for it in lineas.get(k, []):
                res[k][it["codigo"]] = res[k].get(it["codigo"], Decimal(0)) + D(it["cantidad"])
        for k in res:
            for cod in [c for c, q in res[k].items() if q == 0]:
                res[k].pop(cod)
        return res

    def codigo_heredado(self, hoja, linea, codigo):
        """True si el código ya viene de algún proyecto en esa hoja y línea."""
        return self._cant_proyectos(hoja).get(linea, {}).get(codigo, Decimal(0)) > 0

    def _cant_editable(self, base, hoja):
        """
        Cantidades de una hoja editable. En el Ramal es la suma de los
        proyectos MÁS lo capturado en el propio Ramal; en un proyecto es solo
        lo capturado en ese proyecto.
        """
        if base is not self.ramal:
            res = {k: {} for k, _, _ in LINEAS}
            lineas = base["items"].get(hoja)
            if not lineas:
                return res
            for k in res:
                for it in lineas[k]:
                    res[k][it["codigo"]] = res[k].get(it["codigo"], Decimal(0)) + D(it["cantidad"])
            return res
        propias = self._cant_propia(hoja)
        res = self._cant_proyectos(hoja)
        for k in res:
            for cod, q in propias[k].items():
                res[k][cod] = res[k].get(cod, Decimal(0)) + q
        return res

    def _cantidades(self, base, hoja):
        if hoja in EDITABLES:
            return self._cant_editable(base, hoja)
        a = self._cant_editable(base, "ANE")
        v = self._cant_editable(base, "ANE_VAR")
        if hoja == "VARIACION":                     # TODO-EXCEL: lógica real de variación
            res = {}
            for k in a:
                res[k] = {}
                for cod in dict.fromkeys(list(v[k]) + list(a[k])):
                    dif = v[k].get(cod, Decimal(0)) - a[k].get(cod, Decimal(0))
                    if dif != 0:
                        res[k][cod] = dif
            return res
        return v                                    # OC / OTS / ATE: vista provisional

    def _modo_en(self, base, hoja, linea, codigo):
        """Modo individual guardado en la fila, o None si el código no está."""
        for it in (base.get("items", {}).get(hoja) or {}).get(linea, []):
            if it.get("codigo") == codigo:
                return it.get("redondeo", MODO_DEFECTO)
        return None

    def _modo_proyectos(self, hoja, linea, codigo):
        """Modo heredado de los proyectos, sin mirar el override del Ramal."""
        for src in self.proyectos:
            if hoja in EDITABLES:
                modo = self._modo_en(src, hoja, linea, codigo)
            elif hoja == "VARIACION":
                modo = self._modo_en(src, "ANE_VAR", linea, codigo)
                if modo is None:
                    modo = self._modo_en(src, "ANE", linea, codigo)
            else:                                   # OC / OTS / ATE provisional
                modo = self._modo_en(src, "ANE_VAR", linea, codigo)
            if modo is not None:
                return modo
        return MODO_DEFECTO

    def _modo_agregado(self, base, hoja, linea, codigo):
        """Modo para filas derivadas (Ramal, VARIACIÓN, OC/OTS/ATE)."""
        # El Ramal puede tener un redondeo propio por fila; si existe, manda.
        if base is self.ramal:
            modo = self._modo_en(base, hoja, linea, codigo)
            if modo is not None:
                return modo
            return self._modo_proyectos(hoja, linea, codigo)
        if hoja in EDITABLES:
            modo = self._modo_en(base, hoja, linea, codigo)
        elif hoja == "VARIACION":
            modo = self._modo_en(base, "ANE_VAR", linea, codigo)
            if modo is None:
                modo = self._modo_en(base, "ANE", linea, codigo)
        else:                                   # OC / OTS / ATE provisional
            modo = self._modo_en(base, "ANE_VAR", linea, codigo)
        return modo if modo is not None else MODO_DEFECTO

    def _tipo_documento(self, codigo, diferencia):
        """
        Misma regla que la fórmula de Excel: el código es ATE solo si es
        numérico Y mayor a UMBRAL_CODIGO_ATE. El valor exacto del umbral no es
        ATE y un código alfanumérico tampoco; en esos casos manda la cantidad
        variada. El orden de las ramas replica el IF anidado de la hoja.
        """
        texto = str(codigo).strip()
        # isascii() descarta dígitos Unicode ("²", "٣") que Excel no leería
        # como número; isdigit() vacío devuelve False, así que "" no revienta.
        if (texto.isascii() and texto.isdigit()
                and int(texto) > UMBRAL_CODIGO_ATE):
            return "ATE"
        if diferencia > 0:
            return "OTS"
        if diferencia < 0:
            return "OC"
        return "---"

    def _modo_doc_ramal_heredado(self, linea, codigo):
        """Modo de documento que el Ramal hereda si no tiene uno propio."""
        clave = f"{linea}:{codigo}"
        for src in self.proyectos:
            modo = (src.get("redondeo_doc") or {}).get(clave)
            if modo in MODOS_VALIDOS:
                return modo
        # Sin modo de documento en los proyectos, manda el ANE-VAR efectivo
        # del Ramal (que puede tener su propio override de redondeo).
        return self._modo_agregado(self.ramal, "ANE_VAR", linea, codigo)

    def modo_doc(self, base, linea, codigo):
        """Modo de redondeo propio de OC/OTS/ATE para un codigo."""
        clave = f"{linea}:{codigo}"
        if base is self.ramal:
            modo = (base.get("redondeo_doc") or {}).get(clave)
            if modo in MODOS_VALIDOS:
                return modo
            return self._modo_doc_ramal_heredado(linea, codigo)
        modo = (base.get("redondeo_doc") or {}).get(clave)
        if modo in MODOS_VALIDOS:
            return modo
        return self._modo_agregado(base, "ANE_VAR", linea, codigo)

    def set_modo_doc(self, base, linea, codigo, modo):
        if modo not in MODOS_VALIDOS:
            modo = MODO_DEFECTO
        base.setdefault("redondeo_doc", {})[f"{linea}:{codigo}"] = modo

    def set_modo_linea(self, base, hoja, linea, codigo, modo):
        """Cambia el redondeo de una fila de ANE/ANE-VAR (proyecto o Ramal)."""
        if modo not in MODOS_VALIDOS:
            modo = MODO_DEFECTO
        if base is self.ramal:
            # El Ramal no captura materiales: la fila solo guarda el redondeo.
            lineas = base.setdefault("items", {}).setdefault(hoja, _vacio())
            lista = lineas.setdefault(linea, [])
            for it in lista:
                if it.get("codigo") == codigo:
                    it["redondeo"] = modo
                    self.cambios = True
                    return True
            lista.append({"codigo": codigo, "cantidad": "0", "redondeo": modo})
            self.cambios = True
            return True
        for it in (base.get("items", {}).get(hoja) or {}).get(linea, []):
            if it.get("codigo") == codigo:
                it["redondeo"] = modo
                self.cambios = True
                return True
        return False

    def _limpiar_redondeos_ramal(self):
        """
        Deja en el Ramal solo los redondeos que difieren de los proyectos.

        Al abrir un Excel, todas las filas del Ramal traen su redondeo efectivo.
        Si se conservaran todas como override, el Ramal dejaría de seguir los
        cambios de redondeo de los proyectos. Se descartan las que coinciden
        con la herencia y solo quedan las decisiones propias del Ramal.
        """
        if not self.ramal_visible():
            return
        for hoja in EDITABLES:
            # «presentes» son las filas que existen de verdad en el Ramal: las
            # heredadas y las propias, ambas con cantidad. Un ajuste de redondeo
            # (cantidad 0) de un código que ya no está en ningún proyecto no
            # cuenta como presente, así que se descarta abajo.
            heredadas = self._cant_proyectos(hoja)
            propias = self._cant_propia(hoja)
            presentes = {k: {cod for cod, q in heredadas[k].items() if q > 0}
                         | {cod for cod, q in propias[k].items() if q > 0}
                         for k in heredadas}
            lineas = (self.ramal.get("items") or {}).get(hoja) or {}
            for linea, lista in list(lineas.items()):
                for it in list(lista):
                    codigo = it.get("codigo")
                    if codigo is None or codigo not in presentes.get(linea, {}):
                        lista.remove(it)
                        continue
                    # Una fila capturada en el Ramal es un DATO, no un override
                    # de redondeo: se conserva aunque su redondeo coincida con
                    # el heredado de los proyectos.
                    if D(it.get("cantidad", "0")) > 0:
                        continue
                    heredado = self._modo_proyectos(hoja, linea, codigo)
                    if it.get("redondeo", MODO_DEFECTO) == heredado:
                        lista.remove(it)
                if not lista:
                    lineas.pop(linea, None)
        for clave in list((self.ramal.get("redondeo_doc") or {}).keys()):
            linea, sep, codigo = clave.partition(":")
            if not sep or not codigo:
                continue
            heredado = self._modo_doc_ramal_heredado(linea, codigo)
            if self.ramal["redondeo_doc"].get(clave) == heredado:
                self.ramal["redondeo_doc"].pop(clave, None)
        self.cambios = True

    def _fila_variacion(self, base, linea, codigo, cantidad_ane, cantidad_av):
        c = self.comp(codigo)
        precio = D(c["precio"]) if c else Decimal(0)
        diferencia = cantidad_av - cantidad_ane
        tipo = self._tipo_documento(codigo, diferencia)
        modo_ane = self._modo_agregado(base, "ANE", linea, codigo)
        modo_av = self._modo_agregado(base, "ANE_VAR", linea, codigo)
        modo_doc = self.modo_doc(base, linea, codigo)
        total_ane = redondear_monto(precio * cantidad_ane, modo_ane)
        total_av = redondear_monto(precio * cantidad_av, modo_av)
        bruto_doc = precio * diferencia
        monto_doc = (redondear_monto(bruto_doc, modo_doc)
                     if tipo in ("OC", "OTS", "ATE") else Decimal(0))
        return dict(
            codigo=codigo,
            desc=c["descripcion"] if c else "(no está en la BD)",
            unidad=c["unidad"] if c else "",
            precio=precio,
            cantidad=cantidad_ane,
            cantidad_ane=cantidad_ane,
            cantidad_av=cantidad_av,
            cantidad_var=diferencia,
            bruto=bruto_doc,
            redondeo=modo_doc,
            modo_doc=modo_doc,
            modo_ane=modo_ane,
            modo_av=modo_av,
            total_ane=total_ane,
            total_av=total_av,
            tipo=tipo,
            monto_doc=monto_doc,
            total=monto_doc,
        )

    def _filas_derivadas(self, base, hoja):
        a = self._cant_editable(base, "ANE")
        v = self._cant_editable(base, "ANE_VAR")
        out = {}
        for k, _, _ in LINEAS:
            rows = []
            for cod in dict.fromkeys(list(a[k].keys()) + list(v[k].keys())):
                fila = self._fila_variacion(
                    base, k, cod,
                    a[k].get(cod, Decimal(0)),
                    v[k].get(cod, Decimal(0)))
                if hoja in ("OC", "OTS", "ATE") and fila["tipo"] != hoja:
                    continue
                rows.append(fila)
            out[k] = rows
        return out

    def filas(self, base, hoja):
        out = {}
        # En ANE / ANE-VAR de un proyecto, cada fila conserva SU modo individual.
        if hoja in EDITABLES and base is not self.ramal:
            lineas = base.get("items", {}).get(hoja) or {}
            for k, _, _ in LINEAS:
                rows = []
                for it in lineas.get(k, []):
                    cod = it["codigo"]
                    c = self.comp(cod)
                    precio = D(c["precio"]) if c else Decimal(0)
                    q = D(it.get("cantidad", "0"))
                    modo = it.get("redondeo", MODO_DEFECTO)
                    bruto = precio * q
                    total = redondear_monto(bruto, modo)
                    rows.append(dict(codigo=cod, desc=c["descripcion"] if c else "(no está en la BD)",
                                     unidad=c["unidad"] if c else "", precio=precio,
                                     cantidad=q, bruto=bruto, redondeo=modo, total=total))
                out[k] = rows
            return out

        if hoja in EDITABLES:                 # Ramal: proyectos + filas propias
            # Las filas del Ramal se muestran en el mismo orden que
            # CODIGOS_SICOA, sin reordenar las cantidades sumadas.
            indice = {c["codigo"]: i for i, c in enumerate(self.bd)}
            propias = self._cant_propia(hoja)
            heredadas = self._cant_proyectos(hoja)
            for k, cants in self._cant_editable(base, hoja).items():
                rows = []
                for cod in sorted(cants,
                                  key=lambda x: (indice.get(x, len(indice)), str(x))):
                    q = cants[cod]
                    c = self.comp(cod)
                    precio = D(c["precio"]) if c else Decimal(0)
                    modo = self._modo_agregado(base, hoja, k, cod)
                    bruto = precio * q
                    total = redondear_monto(bruto, modo)
                    rows.append(dict(codigo=cod,
                                     desc=c["descripcion"] if c else "(no está en la BD)",
                                     unidad=c["unidad"] if c else "", precio=precio,
                                     cantidad=q, bruto=bruto, redondeo=modo,
                                     total=total,
                                     # «propio» = fila capturada en el Ramal y en
                                     # ningún proyecto: solo esas se editan y se
                                     # borran; las heredadas solo cambian su redondeo.
                                     propio=not (heredadas.get(k, {}).get(cod, Decimal(0)) > 0),
                                     cant_propia=propias.get(k, {}).get(cod, Decimal(0))))
                out[k] = rows
            return out

        return self._filas_derivadas(base, hoja)

    @staticmethod
    def subtotales(filas):
        """
        Total de cada tabla = suma de sus celdas de Precio total.

        Cada celda ya viene redondeada desde `filas`; por eso aquí NO se
        vuelve a redondear. El general será la suma de estos subtotales.
        """
        return {k: sum((r["total"] for r in lista), Decimal(0))
                for k, lista in filas.items()}

    @staticmethod
    def total(filas):
        """Total general = suma de los totales de cada tabla, sin redondear otra vez."""
        return sum(Estado.subtotales(filas).values(), Decimal(0))

    # ── serialización ───────────────────────────────────────────────
    def a_dict(self):
        # La clasificación (MT/MTBT/BT/G/voltaje) es metadato del proyecto:
        # se guarda siempre, aunque la base de códigos sea la global.
        clasif = {c["codigo"]: {k: int(c.get(k, 0))
                               for k in ("MT", "MTBT", "BT", "G", "v76", "v199")}
                  for c in self.bd}
        return {"version": 1, "bd": self.bd if self.bd_local else None,
                "clasificacion": clasif,
                "proyectos": self.proyectos, "ramal": self.ramal,
                "ramal_presente": bool(getattr(self, "ramal_presente", False)),
                "totales_previos": self.totales_previos}

    def desde_dict(self, d):
        self.reset()
        if d.get("bd"):
            self.bd, self.bd_local = d["bd"], True
        else:
            self.bd, self.bd_local = cargar_bd_inicial(), False
        for c in self.bd:
            f = (d.get("clasificacion") or {}).get(c["codigo"])
            if f:
                c.update({k: int(f.get(k, 0))
                          for k in ("MT", "MTBT", "BT", "G", "v76", "v199")})
        self.proyectos = d["proyectos"]
        self.ramal = d["ramal"]
        self.ramal_presente = bool(d.get("ramal_presente", False))
        # Compatibilidad: si el proyecto viejo tenía un modo global, se copia
        # a cada fila que aún no tenga modo individual. Si no, se usa Redondear.
        modo_previo = d.get("redondeo", MODO_DEFECTO)
        if modo_previo not in MODOS_VALIDOS:
            modo_previo = MODO_DEFECTO
        for b in self.proyectos + [self.ramal]:
            b["abrev"] = _xl_token(b.get("abrev", "X"))
            b.setdefault("redondeo_doc", {})
            b.setdefault("orden_manual", [])
            manual = set(b.get("orden_manual") or [])
            for hoja, lineas in (b.get("items") or {}).items():
                for linea, lista in lineas.items():
                    for it in lista:
                        it.setdefault("redondeo", modo_previo)
                    if f"{hoja}:{linea}" not in manual:
                        self.ordenar_por_codigo(lista)
        self.totales_previos = d.get("totales_previos", {})


# ═════════════════════════════════════════════════════════════════════
#  EXPORTACIÓN A EXCEL   (TODO-EXCEL: adaptar al formato de la plantilla)
# ═════════════════════════════════════════════════════════════════════
def _xl_token(abrev):
    """Nombre seguro para tablas de Excel (sin acentos, espacios ni guiones)."""
    s = unicodedata.normalize("NFKD", str(abrev)).encode("ascii", "ignore").decode()
    s = re.sub(r"[^A-Za-z0-9_]", "_", s)
    return (s or "X")[:MAX_ABREV]


def _xl_color_bgr(color):
    """Convierte #RRGGBB al entero BGR que usa Excel para el color de pestaña."""
    c = QColor(str(color or "#3498db"))
    if not c.isValid():
        c = QColor("#3498db")
    return c.red() | (c.green() << 8) | (c.blue() << 16)


def _xl_color_hex(ws, defecto=None):
    """Lee el color visible de la pestaña y lo devuelve como #RRGGBB."""
    try:
        color = ws.api.Tab.Color
        if isinstance(color, bool) or not color:
            return defecto
        color = int(color)
        return f"#{color & 0xFF:02x}{(color >> 8) & 0xFF:02x}{(color >> 16) & 0xFF:02x}"
    except Exception:
        return defecto


def _xl_tabla(ws, nombre):
    objetivo = nombre.upper()
    for t in ws.tables:
        if t.name.upper() == objetivo:
            return t
    return None


def _xl_sufijo(linea):
    return {"MT": "MT", "MTBT": "MT_BT", "BT": "BT", "G": "G"}[linea]


def _xl_nombre_tabla(hoja, token, linea):
    return f"{TABLA_PREFIJO[hoja]}_{_xl_sufijo(linea)}_{token}"


def _xl_col_cantidad(tabla):
    """
    Devuelve el índice de la columna 'Cantidad a ejecutar' de una tabla.

    Se busca por nombre y no por posición fija para no depender del índice 6
    (columna F): al importar un anexo hay que leer esa columna para distinguir
    las cantidades de un proyecto (números) de las del Ramal (SUMIF).
    Si el encabezado no aparece se asume 6, que es la posición actual en la
    plantilla.
    """
    try:
        for i in range(1, int(tabla.api.ListColumns.Count) + 1):
            nombre = str(tabla.api.ListColumns(i).Name or "")
            if _xl_normalizar_etiqueta(nombre) == "CANTIDAD A EJECUTAR":
                return i
    except Exception:
        pass
    return 6


def _xl_normalizar_etiqueta(valor):
    """Colapsa espacios y saltos de línea para comparar encabezados."""
    return " ".join(str(valor or "").replace("\n", " ").split()).upper()


def _xl_normalizar_col_variada(ws):
    """
    Deja 'CANTIDAD VARIADA' con un solo espacio, sin un Replace global.

    El Replace sobre toda la hoja puede tardar minutos en OC/OTS/ATE, así que
    solo se renombra la columna de cada tabla y los rótulos manuales de la
    columna K. Las fórmulas de la app usan después el nombre real de la tabla.
    """
    for t in ws.tables:
        try:
            for i in range(1, int(t.api.ListColumns.Count) + 1):
                col = t.api.ListColumns(i)
                nombre = str(col.Name or "")
                if (_xl_normalizar_etiqueta(nombre) == "CANTIDAD VARIADA"
                        and nombre != "CANTIDAD VARIADA"):
                    col.Name = "CANTIDAD VARIADA"
        except Exception:
            continue
    try:
        ultima = int(ws.api.UsedRange.Row + ws.api.UsedRange.Rows.Count)
        n = min(max(ultima, 1), XL_MAX_FILAS_LECTURA)
        # Una sola lectura de K1:K{n}; solo se escribe la celda que cambia
        # (normalmente ninguna), en vez de una llamada COM por celda.
        valores = _xl_valores(ws.range((1, 11), (n, 11)).options(ndim=2).value)
        for i, fila in enumerate(valores, start=1):
            valor = fila[0] if fila else None
            if (isinstance(valor, str)
                    and _xl_normalizar_etiqueta(valor) == "CANTIDAD VARIADA"
                    and valor != "CANTIDAD VARIADA"):
                ws.range(f"K{i}").value = "CANTIDAD VARIADA"
    except Exception:
        pass


def _xl_col_variada(tabla):
    """Devuelve el nombre real de la columna CANTIDAD VARIADA de la tabla."""
    try:
        for i in range(1, int(tabla.api.ListColumns.Count) + 1):
            nombre = str(tabla.api.ListColumns(i).Name or "")
            if _xl_normalizar_etiqueta(nombre) == "CANTIDAD VARIADA":
                return nombre
    except Exception:
        pass
    return "CANTIDAD VARIADA"


def _xl_ajustar_filas(tabla, n):
    actual = int(tabla.api.ListRows.Count)
    # Agregar sigue siendo ListRows.Add() en bucle: hacerlo en bloque con
    # Resize/Insert puede pisar el título de la sección siguiente del machote.
    while actual < n:
        tabla.api.ListRows.Add()
        actual += 1
    if actual > n:
        # Borrar en bloque: una llamada COM en vez de una por fila. Si no se
        # logra el tamaño esperado se cae al borrado fila por fila de antes.
        try:
            tabla.api.DataBodyRange.Rows(f"{n + 1}:{actual}").Delete()
            if int(tabla.api.ListRows.Count) == n:
                return
        except Exception:
            pass
        actual = int(tabla.api.ListRows.Count)
        while actual > n:
            tabla.api.ListRows(actual).Delete()
            actual -= 1


def _xl_codigo_celda(codigo):
    """
    Normaliza el código que se escribe en la columna CODIGO COMPONENTE.

    Si el código es numérico se escribe como número para que BUSCARX
    (XLOOKUP) y SUMAR.SI (SUMIF) lo encuentren sin depender del tipo de dato;
    un alfanumérico (P001, E051) se escribe tal cual, como texto. Antes ANE
    lo escribía siempre como texto y VARIACIÓN/OC/OTS/ATE como número, así
    que la BUSCARX de "Cantidad a ejecutar" podía devolver 0 y descuadrar
    VARIACIÓN. isascii() descarta dígitos Unicode ("²") que Excel no leería
    como número.
    """
    texto = str(codigo).strip()
    if texto.isascii() and texto.isdigit():
        return int(texto)
    return codigo


def _xl_columna(tabla, idx, valores, formulas=False):
    if not valores:
        return
    rango = tabla.api.ListColumns(idx).DataBodyRange
    if rango is None:
        return
    datos = [[v] for v in valores]
    if formulas:
        rango.Formula = datos
    else:
        rango.Value = datos


def _xl_escribir_formulas(rango, formulas):
    """Escribe fórmulas agrupando filas consecutivas con la MISMA fórmula.

    Una llamada COM por bloque en vez de por celda. Antes de darlo por bueno se
    relee el inicio de cada bloque: en una columna calculada Excel puede
    reimponer la fórmula de la primera fila a todas, así que si no coincide se
    cae al recorrido celda por celda (comportamiento anterior).
    """

    def _normal(f):
        return "".join(str(f or "").upper().split())

    if not formulas:
        return
    # Tramos contiguos con la misma fórmula. Asignar un string a un rango de
    # varias celdas escribe esa fórmula en todas; las referencias [@[...]] y
    # $A no dependen de la fila, así que el resultado es el mismo.
    tramos = []
    inicio = 0
    for i in range(1, len(formulas) + 1):
        if i == len(formulas) or formulas[i] != formulas[inicio]:
            tramos.append((inicio + 1, i, formulas[inicio]))
            inicio = i
    try:
        for a, b, formula in tramos:
            bloque = rango.Cells(a, 1) if a == b else rango.Rows(f"{a}:{b}")
            bloque.Formula = formula
        for a, b, formula in tramos:
            if _normal(rango.Cells(a, 1).Formula) != _normal(formula):
                raise ValueError("Excel reimpuso la fórmula de la columna")
            if b > a and _normal(rango.Cells(b, 1).Formula) != _normal(formula):
                raise ValueError("Excel no escribió todo el bloque")
        return
    except Exception:
        pass
    # Degradación: celda por celda, como antes.
    for i, formula in enumerate(formulas, start=1):
        try:
            rango.Cells(i, 1).Formula = formula
        except Exception:
            pass


def _xl_columna_formulas(tabla, idx, formulas, valores=None):
    """
    Escribe las fórmulas de una columna de tabla por tramos contiguos.

    En una columna calculada de tabla, Excel puede copiar la fórmula de la
    primera fila a todas; para modos de redondeo distintos por fila hay que
    romper primero la columna calculada escribiendo valores y luego asignar
    las fórmulas. `_xl_escribir_formulas` agrupa las filas con la misma
    fórmula y, si Excel reimpone la columna, se degrada al recorrido celda por
    celda de antes.
    """
    if not formulas:
        return
    rango = tabla.api.ListColumns(idx).DataBodyRange
    if rango is None:
        return
    if valores is not None:
        rango.Value = [[v] for v in valores]
    _xl_escribir_formulas(rango, formulas)


def _xl_encabezado(tabla):
    return int(tabla.api.HeaderRowRange.Row)


def _xl_alto_descripcion(ws, fila, descripcion):
    """Deja la fila en 45 si la descripción no cabe en el ancho de la plantilla."""
    texto = str(descripcion or "")
    if len(texto) > UMBRAL_DESC_LARGA or "\n" in texto:
        try:
            ws.api.Rows(f"{fila}:{fila}").RowHeight = ALTO_DESC_LARGA
        except Exception:
            pass


def _rangos_contiguos(numeros):
    """Agrupa enteros en tramos consecutivos: [1,2,3,7,9,10] → [(1,3),(7,7),(9,10)].

    Sirve para escribir un formato/alto por TRAMO en vez de una llamada COM por
    celda. Se ordena y depura por si el llamador no lo hizo.
    """
    tramos = []
    for n in sorted(set(int(x) for x in numeros)):
        if tramos and n == tramos[-1][1] + 1:
            tramos[-1] = (tramos[-1][0], n)
        else:
            tramos.append((n, n))
    return tramos


def _xl_autoalto_descripciones(wb, estado):
    """
    Ajusta el alto real de las filas de descripción DESPUÉS de calcular Excel.

    La plantilla usa BUSCARX, así que el texto definitivo no existe hasta que
    Excel recalcula. Por eso el alto se ajusta en una pasada final, con las
    MISMAS reglas para todas las hojas (ANE/ANE-VAR/VARIACIÓN/OC/OTS/ATE):
      - conserva 45 para textos de más de 30 caracteres;
      - usa AutoFit real para los que pasan de 110 (dos líneas) o de 220
        (tres líneas) y para los que traen salto de línea, sin cortar la
        descripción.

    Las descripciones se leen de una vez y los altos se escriben por tramos
    contiguos: antes eran tres llamadas COM por fila.
    """
    for base in estado.bases():
        for hoja in HOJAS:
            if not base["hojas"][hoja]:
                continue
            try:
                ws = wb.sheets[estado.titulo(base, hoja)]
            except Exception:
                continue
            token = _xl_token(base["abrev"])
            for linea, _, _ in LINEAS:
                tabla = _xl_tabla(ws, _xl_nombre_tabla(hoja, token, linea))
                if tabla is None:
                    continue
                try:
                    filas = int(tabla.api.ListRows.Count)
                except Exception:
                    continue
                if filas <= 0:
                    continue
                encabezado = _xl_encabezado(tabla)
                try:
                    rango = tabla.api.DataBodyRange
                except Exception:
                    continue
                if rango is None:
                    continue
                col_desc = 3 if hoja in EDITABLES else 2
                try:
                    rango.Columns(col_desc).WrapText = True
                    # AutoFit sobre todo el rango es una sola llamada a Excel;
                    # luego solo se corrige el alto de las filas que lo piden.
                    rango.Rows.AutoFit()
                except Exception:
                    pass
                try:
                    textos = _xl_valores(rango.Columns(col_desc).Value)
                except Exception:
                    continue
                # Filas cuyo alto depende del AutoFit recién aplicado.
                altos = {}
                for i, fila in enumerate(textos[:filas], start=1):
                    texto = str((fila[0] if fila else "") or "")
                    fila_excel = encabezado + i
                    try:
                        if len(texto) > UMBRAL_DESC_AUTO or "\n" in texto:
                            altos[fila_excel] = max(
                                15.0, float(ws.api.Rows(fila_excel).RowHeight))
                        elif len(texto) > UMBRAL_DESC_LARGA:
                            altos[fila_excel] = max(
                                ALTO_DESC_LARGA, float(ws.api.Rows(fila_excel).RowHeight))
                        # 30 caracteres o menos, sin salto: se deja como está.
                    except Exception:
                        continue
                # Un RowHeight por tramo contiguo que comparte el mismo alto.
                por_alto = {}
                for fila_excel, alto in altos.items():
                    por_alto.setdefault(alto, []).append(fila_excel)
                for alto, filas_alto in por_alto.items():
                    for a, b in _rangos_contiguos(filas_alto):
                        try:
                            ws.api.Rows(f"{a}:{b}").RowHeight = alto
                        except Exception:
                            pass


def _xl_ocultar_tabla(ws, tabla, visible, titulo=True, titulo_despues=False):
    """
    Oculta la sección completa de una tabla vacía.

    No basta con esconder el encabezado: al borrar las filas de datos queda
    una fila en blanco dentro del rango de la tabla y se ve un hueco. Por eso
    se oculta todo el rango de la tabla (encabezado + fila en blanco) y,
    cuando corresponde, también el título de categoría.

    `titulo=True` busca el título en la fila anterior a la tabla (ANE/ANE-VAR).
    `titulo_despues=True` lo busca en la fila siguiente (VARIACIÓN y OC/OTS/ATE),
    donde el rótulo de la sección va después de las filas de la tabla.
    """
    fila = _xl_encabezado(tabla)
    fin = fila + int(tabla.api.Range.Rows.Count) - 1
    ws.api.Rows(f"{fila}:{fin}").Hidden = not visible
    if not titulo and not titulo_despues:
        return
    fila_titulo = fin + 1 if titulo_despues else fila - 1
    try:
        # El título puede estar en A (ANE/ANE-VAR/VARIACIÓN) o en B (OC/OTS/ATE).
        encontrado = False
        for col in ("A", "B"):
            valor = ws.range(f"{col}{fila_titulo}").value
            if isinstance(valor, str):
                t = norm(valor)
                if t.startswith("media") or t.startswith("baja") or t.startswith("general"):
                    encontrado = True
                    break
        if encontrado:
            ws.api.Rows(fila_titulo).Hidden = not visible
    except Exception:
        pass


def _xl_ocultar_encabezados_tablas(ws):
    """Oculta la fila de encabezado de TODAS las tablas de una hoja."""
    for t in ws.tables:
        try:
            t.api.HeaderRowRange.EntireRow.Hidden = True
        except Exception:
            pass


def _xl_renombrar_tablas_base(wb):
    """
    Libera los nombres de tabla que usarán las hojas de proyecto.

    Excel exige nombres de tabla únicos en todo el libro. Si el proyecto se
    llama igual que la base (p. ej. RAMAL), la tabla copiada no podría tomar
    el nombre final mientras la tabla base siga existiendo. Renombramos las
    tablas base antes de copiar; las hojas base quedan ocultas al final.
    """
    for nombre in list(BASE_COMUNIDAD.values()) + list(BASE_RAMAL.values()):
        try:
            ws = wb.sheets[nombre]
        except Exception:
            continue
        for t in list(ws.tables):
            try:
                t.api.Name = t.name + "__BASE"
            except Exception:
                pass


def _xl_renombrar_hojas_base(wb):
    """
    Libera los nombres de hoja base (ANE. - COMUNIDAD, ANE. - RAMAL, ...).

    Sin esto, un proyecto cuya abreviatura sea COMUNIDAD o el Ramal (RAMAL)
    intentaría crear una hoja con el mismo nombre que la base y Excel lo
    rechazaría. Las hojas base quedan ocultas y con nombres internos.
    """
    mapa = {}
    for nombre in list(BASE_COMUNIDAD.values()) + list(BASE_RAMAL.values()):
        try:
            ws = wb.sheets[nombre]
        except Exception:
            continue
        nuevo = f"_BASE_{len(mapa) + 1}"
        ws.name = nuevo
        mapa[nombre] = nuevo
    return mapa


def _xl_crear_hoja(wb, estado, base_nombre, nuevo_nombre, token, base, hoja, mapa_base):
    base_ws = wb.sheets[base_nombre]
    base_ws.api.Copy(After=wb.sheets[-1].api)
    ws = wb.sheets[-1]
    ws.name = nuevo_nombre
    # Las tablas copiadas pueden traer sufijos de Excel. Se renombran por
    # índice en el orden fijo MT, MT+BT, BT, G con el nombre final del token.
    for i, (linea, _, _) in enumerate(LINEAS, start=1):
        nuevo = _xl_nombre_tabla(hoja, token, linea)
        try:
            ws.api.ListObjects(i).Name = nuevo
        except Exception:
            pass
    # Las referencias directas a otra hoja del machote están en ANE_VAR A1
    # y OC/OTS/ATE B12 (apuntan al ANE). Se reemplazan por la hoja del
    # proyecto para poder quitar después las hojas base del archivo final.
    if hoja == "ANE_VAR":
        try:
            f = ws.range("A1").formula
            if isinstance(f, str) and mapa_base.get("ANE") and base["hojas"].get("ANE"):
                titulo_ane = estado.titulo(base, "ANE")
                # El nombre de hoja lleva espacios/puntos: la referencia debe
                # ir entre comillas simples o Excel rechaza la fórmula.
                ref = "'" + titulo_ane.replace("'", "''") + "'"
                f = f.replace("'" + mapa_base["ANE"] + "'", ref)
                f = f.replace(mapa_base["ANE"], ref)
                ws.range("A1").formula = f
            elif not base["hojas"].get("ANE"):
                ws.range("A1").value = str(base.get("nombre") or "")
        except Exception:
            pass
    if hoja in ("OC", "OTS", "ATE"):
        try:
            if base["hojas"].get("ANE") and mapa_base.get("ANE"):
                f = ws.range("B12").formula
                if isinstance(f, str):
                    titulo_ane = estado.titulo(base, "ANE")
                    # El nombre de hoja lleva espacios/puntos: la referencia debe
                    # ir entre comillas simples o Excel rechaza la fórmula.
                    ref = "'" + titulo_ane.replace("'", "''") + "'"
                    f = f.replace("'" + mapa_base["ANE"] + "'", ref)
                    f = f.replace(mapa_base["ANE"], ref)
                    ws.range("B12").formula = f
            else:
                # Sin ANE exportada no debe quedar una referencia a una hoja
                # que ya no existe; se escribe el nombre como valor.
                ws.range("B12").value = str(base.get("nombre") or "")
        except Exception:
            pass
    if hoja in ("ANE", "ANE_VAR"):
        nombre = str(base.get("nombre") or "").strip()
        if nombre:
            ws.range("A1").value = nombre
    try:
        # Fijar el BGR exacto y quitar el tinte de tema que Excel aplica al
        # asignar ws.color; así el color vuelve igual al abrir el archivo.
        ws.api.Tab.Color = _xl_color_bgr(base["color"])
        try:
            ws.api.Tab.TintAndShade = 0
        except Exception:
            pass
    except Exception:
        try:
            ws.color = base["color"]
        except Exception:
            pass
    return ws


def _xl_limpiar_hojas(wb):
    # INICIO ya no forma parte de la plantilla ni del entregable; si un machote
    # viejo todavía la trae, se descarta junto con las demás hojas sobrantes.
    keep = {"CODIGOS"} | set(BASE_COMUNIDAD.values()) | set(BASE_RAMAL.values())
    for ws in list(wb.sheets):
        if ws.name not in keep:
            ws.api.Delete()


def _xl_escribir_bd(wb, bd):
    """
    Escribe la base de componentes en la copia temporal del libro.

    Se llama SIEMPRE en la exportación: así el .xlsx final lleva los códigos
    nuevos (guardado Local) y, si el guardado fue Global, los que ya trae la
    plantilla. La hoja CODIGOS no se puede quitar nunca porque las BUSCARX de
    ANE/ANE-VAR dependen de esa tabla.
    """
    if not bd:
        return
    try:
        ws = wb.sheets[HOJA_BD]
    except Exception:
        raise RuntimeError(f"El archivo no tiene la hoja {HOJA_BD}; sin ella las "
                           f"BUSCARX de ANE/ANE-VAR quedarían rotas.")
    tabla = _xl_tabla(ws, TABLA_BD)
    if tabla is None:
        raise RuntimeError(f"El archivo no tiene la tabla {TABLA_BD} en la hoja {HOJA_BD}.")

    _xl_asegurar_columnas_clasificacion(tabla)
    _xl_ajustar_filas(tabla, len(bd))
    valores = [[c["codigo"], c["descripcion"], c["unidad"], float(D(c["precio"]))] +
               [int(c.get(k, 0) or 0) for k in BD_COLS[4:]] for c in bd]
    fila_inicio = _xl_encabezado(tabla) + 1
    ws.range(fila_inicio, 1).resize(len(bd), len(BD_ENC)).value = valores


def _xl_config_inicio(wb, estado):
    # La plantilla puede venir sin la hoja INICIO; en ese caso no hay
    # configuración que escribir y la exportación debe continuar.
    try:
        ws = wb.sheets["INICIO"]
    except Exception:
        return
    tabla = _xl_tabla(ws, "Configuration")
    if tabla is not None:
        _xl_ajustar_filas(tabla, len(estado.proyectos))
        for i, b in enumerate(estado.proyectos, start=1):
            ab = _xl_token(b["abrev"])
            vals = [i, b["nombre"] or ab, ab, ab]
            vals += [bool(b["hojas"][h]) for h in HOJAS]
            vals += [True]
            tabla.api.DataBodyRange.Rows(i).Value = [vals]

    tabla_r = _xl_tabla(ws, "Configuration_Ramal")
    if tabla_r is not None:
        _xl_ajustar_filas(tabla_r, 1)
        if estado.ramal_visible():
            b = estado.ramal
            ab = _xl_token(b["abrev"])
            vals = [1, b["nombre"] or ab, ab, ab]
            vals += [bool(b["hojas"][h]) for h in HOJAS]
            vals += [True]
        else:
            vals = [None] * 11
        tabla_r.api.DataBodyRange.Rows(1).Value = [vals]

    ws.range("C7").value = len(estado.proyectos)


def _xl_llenar_ane(wb, estado, base, hoja):
    ws = wb.sheets[estado.titulo(base, hoja)]
    token = _xl_token(base["abrev"])
    filas = estado.filas(base, hoja)
    for linea, _, texto_linea in LINEAS:
        tabla = _xl_tabla(ws, _xl_nombre_tabla(hoja, token, linea))
        if tabla is None:
            continue
        lista = filas[linea]
        _xl_ajustar_filas(tabla, len(lista))
        if lista:
            enc = _xl_encabezado(tabla)
            # El código se escribe con el mismo tipo que en VARIACIÓN/OC/OTS/ATE
            # (número si es numérico) para que el BUSCARX y el SUMAR.SI del
            # Ramal lo encuentren sin choque de tipos.
            codigos_celda = [_xl_codigo_celda(r["codigo"]) for r in lista]
            _xl_columna(tabla, 1, codigos_celda)
            _xl_formato_condicional_codigo_numerico(tabla)
            # El resto de tablas escribe su texto de línea (Media tensión,
            # Media más baja tensión, ...). La tabla General deja Tipo vacío
            # porque WuMDrex no quiere la palabra "General" ahí: la importación
            # se reconstruye por código y cantidad (columnas 1, 6 y 7), nunca
            # por este texto.
            texto_tipo = "" if linea == "G" else texto_linea
            _xl_columna(tabla, 2, [texto_tipo] * len(lista))
            # Referencias estructuradas: la tabla de códigos y la propia fila
            # de la tabla, sin depender de la dirección absoluta de la celda.
            cod = f"{TABLA_BD}[{COL_BD_CODIGO}]"
            f_desc = (f'=_xlfn.XLOOKUP([@[CODIGO COMPONENTE]],{cod},'
                      f'{TABLA_BD}[{COL_BD_DESC}],"")')
            f_uni = (f'=_xlfn.XLOOKUP([@[CODIGO COMPONENTE]],{cod},'
                     f'{TABLA_BD}[{COL_BD_UNIDAD}],"")')
            f_precio = (f'=_xlfn.XLOOKUP([@[CODIGO COMPONENTE]],{cod},'
                        f'{TABLA_BD}[{COL_BD_PRECIO}],0)')
            for i, r in enumerate(lista, start=1):
                _xl_alto_descripcion(ws, enc + i, r["desc"])
            _xl_columna(tabla, 3, [f_desc] * len(lista), formulas=True)
            _xl_columna(tabla, 4, [f_uni] * len(lista), formulas=True)
            _xl_columna(tabla, 5, [f_precio] * len(lista), formulas=True)
            _xl_columna(tabla, 6, [float(r["cantidad"]) for r in lista])
            formulas = [_xl_formula_celda_modo(
                "[@[Precio unitario]]*[@[Cantidad a ejecutar]]", r["redondeo"])
                for r in lista]
            _xl_columna_formulas(tabla, 7, formulas,
                                 valores=[float(r["total"]) for r in lista])
            _xl_formato_moneda(tabla, (5, 7))
            _xl_formato_cantidades_km(tabla, [r["unidad"] for r in lista], (6,))
        _xl_ocultar_tabla(ws, tabla, bool(lista))
    _xl_ocultar_encabezados_tablas(ws)


def _xl_borrar_cant_total(tabla):
    """Quita la columna CANT TOTAL de una tabla; ya no se usa en VARIACIÓN."""
    try:
        if int(tabla.api.ListColumns.Count) < 16:
            return
        nombre = str(tabla.api.ListColumns(16).Name or "").upper()
        if "CANT" in nombre and "TOTAL" in nombre:
            tabla.api.ListColumns(16).Delete()
    except Exception:
        pass


def _xl_formato_condicional_codigo_numerico(tabla):
    """
    Oculta con formato condicional el código de la columna 1 cuando es un
    número MAYOR a UMBRAL_CODIGO_ATE, es decir, solo los códigos ATE.

    La regla se aplica solo al cuerpo de la tabla, no al encabezado. La
    expresión usa únicamente operadores aritméticos y de comparación
    ((A+0=A)*(A>umbral)) en lugar de funciones como ISNUMBER o AND: así el
    motor de formato condicional la evalúa igual en Excel en español o en
    inglés, sin depender del idioma con que estén escritas las funciones.
    Al reexportar se borra además la regla vieja (=A+0=A sin umbral), que
    blanqueaba toda la columna numérica, incluidos los códigos ≤ umbral; el
    borrado exige que la dirección sea la MISMA celda para no eliminar reglas
    ajenas que casualmente tengan esa forma.
    """
    try:
        columna = tabla.api.ListColumns(1)
        nombre = str(columna.Name or "").replace("\n", " ").strip().upper()
        if nombre not in ("CODIGO COMPONENTE", "NO."):
            return
        rango = columna.DataBodyRange
        if rango is None:
            return
        celda = rango.Cells(1, 1)
        direccion = str(celda.Address).replace("$", "")
        formula = (f"=({direccion}+0={direccion})"
                   f"*({direccion}>{UMBRAL_CODIGO_ATE})")
        formula_vieja = f"={direccion}+0={direccion}"
        reglas = rango.FormatConditions
        # Excel puede devolver Formula1 con o sin el "=" inicial; se normaliza
        # sin espacios, en mayúsculas y sin ese signo para comparar parejo.
        objetivo = {formula.replace(" ", "").upper().lstrip("="),
                    formula_vieja.replace(" ", "").upper().lstrip("=")}
        # Evitar reglas duplicadas si la exportación se repite sobre la hoja.
        # La retroreferencia \1 exige que el +0= compare la MISMA celda.
        for i in range(int(reglas.Count), 0, -1):
            try:
                existente = (str(reglas.Item(i).Formula1)
                             .replace(" ", "").upper().lstrip("="))
                if existente in objetivo or re.fullmatch(
                        r"=?([A-Z]{1,3}\d+)\+0=\1$", existente):
                    reglas.Item(i).Delete()
            except Exception:
                pass
        regla = reglas.Add(Type=2, Formula1=formula)
        regla.Font.Color = 16777215  # blanco: el código ATE (> umbral) no se ve
    except Exception:
        pass


def _xl_formato_moneda(tabla, columnas):
    """Aplica formato de moneda a las columnas de precio de una tabla."""
    for col in columnas:
        try:
            rango = tabla.api.ListColumns(col).DataBodyRange
            if rango is not None:
                rango.NumberFormat = FMT_MONEDA_XL
        except Exception:
            pass


def _xl_formato_cantidades_km(tabla, unidades, columnas):
    """Deja las cantidades de kilometraje con tres decimales (0.000).

    Se agrupan las filas de kilometraje en tramos contiguos y se aplica el
    formato por tramo: antes era una llamada COM por celda.
    """
    filas_km = [i for i, unidad in enumerate(unidades, start=1) if es_km(unidad)]
    if not filas_km:
        return
    tramos = _rangos_contiguos(filas_km)
    for col in columnas:
        try:
            rango = tabla.api.ListColumns(col).DataBodyRange
        except Exception:
            continue
        if rango is None:
            continue
        for a, b in tramos:
            try:
                rango.Rows(f"{a}:{b}").NumberFormat = FMT_CANTIDAD_KM_XL
            except Exception:
                pass


def _xl_formulas_base_tabla(ane_tab, av_tab, modo_ane, modo_av, col_var):
    """Fórmulas B-L comunes a VARIACIÓN y a OC/OTS/ATE, con referencias de tabla."""
    cod = f"{TABLA_BD}[{COL_BD_CODIGO}]"
    f_b = f'=_xlfn.XLOOKUP([@[No.]],{cod},{TABLA_BD}[{COL_BD_DESC}],"")'
    f_c = f'=_xlfn.XLOOKUP([@[No.]],{cod},{TABLA_BD}[{COL_BD_UNIDAD}],"")'
    f_d = (f'=_xlfn.XLOOKUP([@[No.]],{ane_tab}[CODIGO COMPONENTE],'
           f'{ane_tab}[Cantidad a ejecutar],0)')
    f_e = f'=_xlfn.XLOOKUP([@[No.]],{cod},{TABLA_BD}[{COL_BD_PRECIO}],0)'
    f_f = _xl_formula_celda_modo(
        "[@[CANTIDAD]]*[@[PRECIO UNITARIO (Q)]]", modo_ane)
    f_g = (f'=_xlfn.XLOOKUP([@[No.]],{av_tab}[CODIGO COMPONENTE],'
           f'{av_tab}[Cantidad a ejecutar],0)')
    f_h = f'=_xlfn.XLOOKUP([@[No.]],{cod},{TABLA_BD}[{COL_BD_PRECIO}],0)'
    f_i = _xl_formula_celda_modo(
        "[@[CANTIDAD - V]]*[@[PRECIO UNITARIO ( Q )]]", modo_av)
    f_j = (f'=IF(AND(ISNUMBER([@[No.]]),[@[No.]]>{UMBRAL_CODIGO_ATE}),"ATE",'
           f'IF([@[{col_var}]]>0,"OTS",'
           f'IF([@[{col_var}]]<0,"OC","---")))')
    f_k = "=[@[CANTIDAD - V]]-[@[CANTIDAD]]"
    f_l = f'=_xlfn.XLOOKUP([@[No.]],{cod},{TABLA_BD}[{COL_BD_PRECIO}],0)'
    return [f_b, f_c, f_d, f_e, f_f, f_g, f_h, f_i, f_j, f_k, f_l]


def _xl_formula_documento(tipo, modo, col_var):
    """Monto de la columna OC/OTS/ATE, con el redondeo elegido para esa fila."""
    calculo = _xl_formula_celda_modo(
        f"[@[{col_var}]]*[@[PRECIO UNITARIO ( Q)]]", modo)[1:]
    return f'=IF([@[Tipo de Documento]]="{tipo}",{calculo},"")'


def _xl_llenar_variacion(wb, estado, base):
    ws = wb.sheets[estado.titulo(base, "VARIACION")]
    _xl_normalizar_col_variada(ws)
    token = _xl_token(base["abrev"])
    ane = estado._cant_editable(base, "ANE")
    av = estado._cant_editable(base, "ANE_VAR")
    quitar_cant_total = False
    for t in ws.tables:
        try:
            if (int(t.api.ListColumns.Count) >= 16
                    and "CANT TOTAL" in str(t.api.ListColumns(16).Name).upper()):
                quitar_cant_total = True
                break
        except Exception:
            continue
    for linea, _, _ in LINEAS:
        codigos = list(dict.fromkeys(list(ane[linea].keys()) + list(av[linea].keys())))
        tabla = _xl_tabla(ws, _xl_nombre_tabla("VARIACION", token, linea))
        if tabla is None:
            continue
        _xl_ajustar_filas(tabla, len(codigos))
        if not codigos:
            _xl_ocultar_tabla(ws, tabla, False, titulo=False, titulo_despues=True)
            _xl_borrar_cant_total(tabla)
            continue

        filas = []
        for cod in codigos:
            c = estado.comp(cod)
            filas.append(dict(
                codigo=cod,
                desc=c["descripcion"] if c else "(no está en la BD)",
                unidad=c["unidad"] if c else "",
                precio=float(D(c["precio"])) if c else 0.0,
                modo_ane=estado._modo_agregado(base, "ANE", linea, cod),
                modo_av=estado._modo_agregado(base, "ANE_VAR", linea, cod),
                modo_doc=estado.modo_doc(base, linea, cod),
            ))

        # En esta hoja solo se escriben los códigos como valor. Todo lo demás
        # se resuelve con BUSCARX o con fórmulas de la propia tabla.
        valores_cod = [_xl_codigo_celda(f["codigo"]) for f in filas]
        _xl_columna(tabla, 1, valores_cod)
        _xl_formato_condicional_codigo_numerico(tabla)

        n = len(filas)
        col_var = _xl_col_variada(tabla)
        ane_tab = _xl_nombre_tabla("ANE", token, linea)
        av_tab = _xl_nombre_tabla("ANE_VAR", token, linea)
        formulas = [_xl_formulas_base_tabla(ane_tab, av_tab,
                                            f["modo_ane"], f["modo_av"],
                                            col_var)
                    for f in filas]
        f_m = [_xl_formula_documento("OC", f["modo_doc"], col_var) for f in filas]
        f_n = [_xl_formula_documento("OTS", f["modo_doc"], col_var) for f in filas]
        f_o = [_xl_formula_documento("ATE", f["modo_doc"], col_var) for f in filas]

        _xl_columna(tabla, 2, [x[0] for x in formulas], formulas=True)
        _xl_columna(tabla, 3, [x[1] for x in formulas], formulas=True)
        _xl_columna(tabla, 4, [x[2] for x in formulas], formulas=True)
        _xl_columna(tabla, 5, [x[3] for x in formulas], formulas=True)
        _xl_columna_formulas(tabla, 6, [x[4] for x in formulas], valores=[0] * n)
        _xl_columna(tabla, 7, [x[5] for x in formulas], formulas=True)
        _xl_columna(tabla, 8, [x[6] for x in formulas], formulas=True)
        _xl_columna_formulas(tabla, 9, [x[7] for x in formulas], valores=[0] * n)
        _xl_columna(tabla, 10, [x[8] for x in formulas], formulas=True)
        _xl_columna(tabla, 11, [x[9] for x in formulas], formulas=True)
        _xl_columna(tabla, 12, [x[10] for x in formulas], formulas=True)
        _xl_columna_formulas(tabla, 13, f_m, valores=[0] * n)
        _xl_columna_formulas(tabla, 14, f_n, valores=[0] * n)
        _xl_columna_formulas(tabla, 15, f_o, valores=[0] * n)
        _xl_formato_moneda(tabla, (5, 6, 8, 9, 12, 13, 14, 15))
        _xl_formato_cantidades_km(tabla, [f["unidad"] for f in filas], (4, 7, 11))
        _xl_ocultar_tabla(ws, tabla, True, titulo=False, titulo_despues=True)
        # La columna se quita al final: si se elimina antes, Excel no permite
        # volver a agregar filas a la tabla durante la misma exportación.
        _xl_borrar_cant_total(tabla)
    if quitar_cant_total:
        # Limpia el encabezado manual que quedaba fuera de las tablas.
        try:
            ws.api.Columns(16).Delete()
        except Exception:
            pass
    # WuMDrex pidió ocultar los encabezados de tabla también en VARIACIÓN:
    # la hoja ya tiene su fila de encabezado manual.
    _xl_ocultar_encabezados_tablas(ws)


def _xl_llenar_documento(wb, estado, base, hoja):
    """Llena OC/OTS/ATE con las filas de su tipo y fórmulas de tabla."""
    ws = wb.sheets[estado.titulo(base, hoja)]
    _xl_normalizar_col_variada(ws)
    token = _xl_token(base["abrev"])
    filas = estado.filas(base, hoja)
    for linea, _, _ in LINEAS:
        lista = filas[linea]
        tabla = _xl_tabla(ws, _xl_nombre_tabla(hoja, token, linea))
        if tabla is None:
            continue
        _xl_ajustar_filas(tabla, len(lista))
        if not lista:
            _xl_ocultar_tabla(ws, tabla, False, titulo=True, titulo_despues=True)
            continue

        valores_cod = [_xl_codigo_celda(f["codigo"]) for f in lista]
        _xl_columna(tabla, 1, valores_cod)
        _xl_formato_condicional_codigo_numerico(tabla)

        n = len(lista)
        col_var = _xl_col_variada(tabla)
        ane_tab = _xl_nombre_tabla("ANE", token, linea)
        av_tab = _xl_nombre_tabla("ANE_VAR", token, linea)
        formulas = [_xl_formulas_base_tabla(ane_tab, av_tab,
                                            f["modo_ane"], f["modo_av"],
                                            col_var)
                    for f in lista]
        f_doc = [_xl_formula_documento(hoja, f["modo_doc"], col_var)
                 for f in lista]

        _xl_columna(tabla, 2, [x[0] for x in formulas], formulas=True)
        _xl_columna(tabla, 3, [x[1] for x in formulas], formulas=True)
        _xl_columna(tabla, 4, [x[2] for x in formulas], formulas=True)
        _xl_columna(tabla, 5, [x[3] for x in formulas], formulas=True)
        _xl_columna_formulas(tabla, 6, [x[4] for x in formulas], valores=[0] * n)
        _xl_columna(tabla, 7, [x[5] for x in formulas], formulas=True)
        _xl_columna(tabla, 8, [x[6] for x in formulas], formulas=True)
        _xl_columna_formulas(tabla, 9, [x[7] for x in formulas], valores=[0] * n)
        _xl_columna(tabla, 10, [x[8] for x in formulas], formulas=True)
        _xl_columna(tabla, 11, [x[9] for x in formulas], formulas=True)
        _xl_columna(tabla, 12, [x[10] for x in formulas], formulas=True)
        _xl_columna_formulas(tabla, 13, f_doc, valores=[0] * n)
        _xl_formato_moneda(tabla, (5, 6, 8, 9, 12, 13))
        _xl_formato_cantidades_km(tabla, [f["unidad"] for f in lista], (4, 7, 11))

        enc = _xl_encabezado(tabla)
        for i, fila in enumerate(lista, start=1):
            _xl_alto_descripcion(ws, enc + i, fila["desc"])
        _xl_ocultar_tabla(ws, tabla, True, titulo=True, titulo_despues=True)
    _xl_ocultar_encabezados_tablas(ws)


def _xl_ocultar_tablas_hoja(wb, estado, base, hoja):
    """Oculta las secciones de una hoja cuyas tablas aún no se llenan (OC/OTS/ATE)."""
    ws = wb.sheets[estado.titulo(base, hoja)]
    token = _xl_token(base["abrev"])
    for linea, _, _ in LINEAS:
        tabla = _xl_tabla(ws, _xl_nombre_tabla(hoja, token, linea))
        if tabla is not None:
            _xl_ocultar_tabla(ws, tabla, False, titulo=False)


def _xl_literal_cantidad(valor):
    """
    Texto de una cantidad para escribirla dentro de una fórmula de Excel.

    Decimal → texto con punto decimal y sin ceros de cola (2, 2.5, 0.75),
    nunca en notación científica y sin depender del separador local: así el
    literal se puede pegar en Range.Formula (convención en-US).
    """
    try:
        d = Decimal(str(valor))
    except (InvalidOperation, ValueError):
        return "0"
    if d == 0:
        return "0"
    return format(d.normalize(), "f")


def _xl_formulas_ramal(wb, estado):
    """
    Cantidad del Ramal = suma de los proyectos + las filas propias del Ramal.

    Las filas heredadas siguen siendo un SUMIF por proyecto; las capturadas en
    el Ramal no viven en ninguna tabla de proyecto, así que se suman como
    literal al final. Sin filas propias la fórmula queda idéntica a antes.
    """
    for hoja in EDITABLES:
        if not estado.ramal["hojas"][hoja]:
            continue
        ws = wb.sheets[estado.titulo(estado.ramal, hoja)]
        token = _xl_token(estado.ramal["abrev"])
        propias = estado._cant_propia(hoja)
        for linea, _, _ in LINEAS:
            tabla = _xl_tabla(ws, _xl_nombre_tabla(hoja, token, linea))
            if tabla is None or tabla.api.DataBodyRange is None:
                continue
            enc = _xl_encabezado(tabla)
            n = int(tabla.api.ListRows.Count)
            if n <= 0:
                continue
            # El código de cada fila se lee en UNA llamada; antes era una
            # lectura y una escritura de fórmula por fila.
            try:
                codigos = _xl_valores(tabla.api.DataBodyRange.Columns(1).Value)
            except Exception:
                continue
            formulas = []
            for i in range(n):
                partes = []
                for p in estado.proyectos:
                    if not p["hojas"][hoja]:
                        continue
                    ptab = _xl_nombre_tabla(hoja, _xl_token(p["abrev"]), linea)
                    partes.append(f"SUMIF({ptab}[CODIGO COMPONENTE],"
                                  f"[@[CODIGO COMPONENTE]],"
                                  f"{ptab}[Cantidad a ejecutar])")
                fila = codigos[i] if i < len(codigos) else ()
                cod = _cod(fila[0]) if fila else ""
                propia = propias.get(linea, {}).get(cod, Decimal(0))
                if propia > 0:
                    partes.append(_xl_literal_cantidad(propia))
                formulas.append("=" + ("+".join(partes) if partes else "0"))
            _xl_escribir_formulas(
                ws.range((enc + 1, 6), (enc + n, 6)).api, formulas)


def _xl_formula_celda_modo(expr, modo):
    """Fórmula de una celda Precio total con el modo individual de esa fila."""
    if modo == "SR":
        return f"=({expr})"          # multiplicación sin redondeo
    funcion = modo if modo in MODOS_VALIDOS else MODO_DEFECTO
    return f"={funcion}({expr},2)"


def _xl_modo_desde_formula(formula):
    """Recupera ROUND / ROUNDUP / ROUNDDOWN / SR de la fórmula de una celda."""
    s = str(formula or "").upper()
    if "ROUNDUP" in s:
        return "ROUNDUP"
    if "ROUNDDOWN" in s:
        return "ROUNDDOWN"
    if "ROUND" in s:
        return "ROUND"
    if "*" in s:
        return "SR"                  # multiplicación sin redondeo
    return MODO_DEFECTO


def _xl_rango_total(ws, total_cell):
    """
    Devuelve el objeto COM donde SÍ se puede escribir el total.

    El TOTAL del machote puede estar en una celda combinada; Excel guarda el
    valor en el ancla del combinado, así que se resuelve al ancla ANTES de
    escribir. La única condición para escribir es que en esa posición no haya
    un rótulo: se acepta una fórmula o una celda vacía, y se rechaza cualquier
    texto ("TOTAL", "Total en letras: …") para no borrar la etiqueta.

    Antes se vetaba cualquier combinado que empezara en la columna A y el total
    del machote (`=SUMA(t1[Precio total]+t2[Precio total]+…)`) se quedaba
    intacto con su #N/D, sin ningún aviso.
    """
    try:
        rango = ws.range(total_cell).api
    except Exception:
        return None
    try:
        if bool(rango.MergeCells):
            rango = rango.MergeArea
    except Exception:
        pass
    try:
        actual = rango.Formula
    except Exception:
        return None
    if actual is None:
        return rango
    if isinstance(actual, str) and (not actual.strip() or actual.startswith("=")):
        return rango
    return None


def _total_con_coma(formula):
    """True si el total quedó como UNA suma con argumentos separados por coma."""
    s = str(formula or "").upper().replace("SUMA", "SUM")
    i = s.find("SUM(")
    if i < 0:
        return False
    dentro = s[i + 4:]
    fin = dentro.find(")")
    if fin >= 0:
        dentro = dentro[:fin]
    return "," in dentro


def _xl_suma_sin_plus(formula):
    """
    Convierte la suma del machote `=SUMA(a+b+c)` en `=SUMA(a,b,c)`.

    Con `+` entre columnas de tablas distintas Excel aplica intersección
    implícita en una celda que no pertenece a esas tablas y el total sale
    #N/D; con la coma cada argumento es un vector completo y la suma es la
    correcta. Devuelve None si la fórmula no tiene esa forma, para no tocar
    nada que no sea una suma de columnas de tabla.
    """
    s = str(formula or "").strip()
    if not s.startswith("="):
        return None
    i = s.upper().find("SUM(")
    if i < 0:
        return None
    ini = i + 4
    nivel = 1
    j = ini
    while j < len(s) and nivel:
        if s[j] == "(":
            nivel += 1
        elif s[j] == ")":
            nivel -= 1
        j += 1
    if nivel:
        return None
    dentro = s[ini:j - 1]
    if "," in dentro:
        return None                  # ya está con comas: no se toca
    partes, actual, nivel = [], [], 0
    for ch in dentro:
        if ch == "(":
            nivel += 1
        elif ch == ")":
            nivel -= 1
        if ch == "+" and nivel == 0:
            partes.append("".join(actual))
            actual = []
            continue
        actual.append(ch)
    partes.append("".join(actual))
    if len(partes) < 2:
        return None
    return s[:ini] + ",".join(p.strip() for p in partes) + s[j - 1:]


def _xl_arreglar_suma_plus(ws):
    """
    Red de seguridad del total: reescribe EN EL LUGAR cualquier suma del
    machote que traiga `+` dentro de una sola SUMA.

    Se aplica DESPUÉS de `_xl_asignar_total` porque esa escritura puede fallar
    en silencio (celda combinada, propiedad rechazada por Excel, rango
    inexistente) y el archivo saldría con #N/D sin que nadie se entere. Solo
    toca fórmulas que YA existen y llevan referencias de tabla (`[`), así que
    no crea, borra ni desplaza ninguna celda ni ningún rótulo.
    """
    try:
        filas = int(ws.api.UsedRange.Rows.Count)
    except Exception:
        return
    n = min(max(filas, 1), XL_MAX_FILAS_LECTURA)
    for col in ("F", "G", "I", "M", "N", "O"):
        try:
            formulas = _xl_valores(
                ws.range(f"{col}1:{col}{n}").options(ndim=2).formula)
        except Exception:
            continue
        for i, fila in enumerate(formulas, start=1):
            f = fila[0] if fila else None
            if not isinstance(f, str) or "[" not in f:
                continue
            nueva = _xl_suma_sin_plus(f)
            if not nueva or nueva == f:
                continue
            try:
                ws.range(f"{col}{i}").formula = nueva
            except Exception:
                continue


def _xl_asignar_total(ws, total_cell, expresiones):
    """
    Deja el total como UNA sola suma de las tablas de la hoja.

    Las expresiones que recibe son referencias a la columna de cada tabla
    (por ejemplo `ANE_MT_TOKEN[Precio total]`), SIN envolver en SUM: la suma se
    arma aquí una única vez, con los argumentos separados por coma.

    Con `+` entre columnas de tablas distintas Excel aplica intersección
    implícita en una celda que no pertenece a esas tablas y el total sale
    #N/D, así que la forma con coma es la primera opción. Si Excel no la
    acepta se cae a la forma de la versión anterior del programa
    (`=SUM(a)+SUM(b)+…`), que tampoco tiene intersección implícita y da el
    mismo número. Cada intento se VERIFICA releyendo la fórmula.

    Cada celda ya viene redondeada; el total NO se redondea otra vez.
    """
    if not expresiones:
        return
    rango = _xl_rango_total(ws, total_cell)
    if rango is None:
        return
    con_coma = "=SUM(" + ", ".join(expresiones) + ")"
    suelta = "=" + " + ".join(f"SUM({e})" for e in expresiones)
    intentos = (
        ("Formula2", con_coma, True),
        ("Formula", con_coma, True),
        ("FormulaLocal", con_coma.replace("SUM(", "SUMA("), True),
        ("Formula2", suelta, False),
        ("Formula", suelta, False),
        ("FormulaLocal", suelta.replace("SUM(", "SUMA("), False),
    )
    for propiedad, texto, quiere_coma in intentos:
        try:
            setattr(rango, propiedad, texto)
        except Exception:
            continue
        try:
            leida = str(rango.Formula or "")
        except Exception:
            continue
        if quiere_coma:
            if _total_con_coma(leida):
                return
        elif "SUM(" in leida.upper().replace("SUMA(", "SUM("):
            return


def _xl_aplicar_redondeos(wb, estado):
    """Cierra los totales de cada hoja: solo suman celdas ya redondeadas."""
    for base in estado.bases():
        for hoja in HOJAS:
            if not base["hojas"][hoja]:
                continue
            ws = wb.sheets[estado.titulo(base, hoja)]
            fila = _xl_fila_total(ws)
            if fila is not None:
                token = _xl_token(base["abrev"])
                tablas = [_xl_nombre_tabla(hoja, token, linea)
                          for linea, _, _ in LINEAS]

                if hoja in ("ANE", "ANE_VAR"):
                    _xl_asignar_total(
                        ws, f"G{fila}",
                        [f"{t}[Precio total]" for t in tablas])

                elif hoja == "VARIACION":
                    _xl_asignar_total(
                        ws, f"F{fila}",
                        [f"{t}[PRECIO TOTAL (Q)]" for t in tablas])
                    _xl_asignar_total(
                        ws, f"I{fila}",
                        [f"{t}[PRECIO TOTAL ( Q )]" for t in tablas])
                    _xl_asignar_total(
                        ws, f"M{fila}",
                        [f"{t}[OC]" for t in tablas])
                    _xl_asignar_total(
                        ws, f"N{fila}",
                        [f"{t}[OTS]" for t in tablas])
                    _xl_asignar_total(
                        ws, f"O{fila}",
                        [f"{t}[ATE]" for t in tablas])

                else:
                    col_doc = {"OC": "OC", "OTS": "OTS", "ATE": "ATE"}[hoja]
                    _xl_asignar_total(
                        ws, f"M{fila}",
                        [f"{t}[{col_doc}]" for t in tablas])

            # Red de seguridad: si la escritura de arriba no prosperó, la
            # fórmula del machote se queda con `+` y el total sale #N/D.
            _xl_arreglar_suma_plus(ws)


def _lineas_texto(texto, por_linea):
    """
    Cuenta las líneas reales que ocupa `texto` en un ajuste de párrafo.

    Se recorre igual que lo hace Excel: cada salto de línea manual (`\\n`)
    empieza un párrafo nuevo y, dentro del párrafo, se acumulan palabras
    mientras quepan en `por_linea`. Una palabra más ancha que la línea se
    parte en `ceil(largo / por_linea)` líneas. Nunca devuelve 0: un texto
    vacío ocupa una línea.

    Función pura (sin Excel ni `unicodedata`): los acentos cuentan 1 porque
    Python ya trabaja con caracteres, no con bytes.
    """
    try:
        por_linea = int(por_linea)
    except Exception:
        por_linea = LETRAS_MAX_CAR
    if por_linea < 1:
        por_linea = 1

    total = 0
    for parrafo in str(texto or "").split("\n"):
        palabras = parrafo.split()
        if not palabras:
            # Un párrafo vacío (o una línea en blanco) también ocupa alto.
            total += 1
            continue
        lineas = 1
        usado = 0   # caracteres ya ocupados en la línea actual
        for palabra in palabras:
            largo = len(palabra)
            if largo > por_linea:
                # Excel parte la palabra: primero termina la línea en curso.
                if usado:
                    lineas += 1
                lineas += -(-largo // por_linea) - 1
                resto = largo % por_linea
                usado = resto if resto else por_linea
                continue
            if usado == 0:
                usado = largo
            elif usado + 1 + largo <= por_linea:
                usado += 1 + largo
            else:
                lineas += 1
                usado = largo
        total += lineas

    return max(1, total)


def _xl_alto_letras(ws, fila):
    """
    Ajusta el alto de la fila del "Total en letras".

    La celda está combinada de A a G y Excel no calcula el alto de una fila
    con celdas combinadas. La versión anterior descombinaba, ensanchaba la
    columna A, ejecutaba AutoFit y restauraba; el resultado era impredecible
    (un texto de 170 caracteres quedaba en 2 líneas y WuMDrex lo ve cortado).

    Ahora no se toca ni el ancho ni la combinación: se mide el ancho real de
    la combinación en puntos, se estima cuántos caracteres caben por línea y
    se cuentan las líneas del texto. El alto es 15 pt por línea, nunca menor
    al alto original del machote y nunca mayor a 409 pt (límite de Excel).
    """
    try:
        celda = ws.range(f"A{fila}")
        # Si el machote tenía "ajustar al tamaño", el texto se ENCOGE en vez de
        # envolverse y el alto calculado no serviría: se fuerza el ajuste.
        try:
            celda.api.WrapText = True
            celda.api.ShrinkToFit = False
        except Exception:
            pass

        # Una combinación en vertical (o de un rango raro) no se puede medir
        # como una sola línea de texto; se deja el alto como está.
        try:
            merge = celda.api.MergeArea
            if int(merge.Rows.Count) != 1:
                return
            ancho_pt = float(merge.Width)
        except Exception:
            return

        texto = str(celda.value or "")
        try:
            tam = float(celda.api.Font.Size or 11.0)
        except Exception:
            tam = 11.0
        if tam <= 0:
            tam = 11.0
        try:
            negrita = bool(celda.api.Font.Bold)
        except Exception:
            negrita = False

        # Caracteres por línea: ancho útil / ancho medio de carácter. La
        # negrita es un 5 % más ancha. El tope es 110 (criterio de WuMDrex).
        try:
            if ancho_pt <= 0:
                raise ValueError("ancho no medible")
            factor = LETRAS_FACTOR_CAR + (0.03 if negrita else 0.0)
            car_por_linea = int((ancho_pt - LETRAS_MARGEN_PT) / (tam * factor))
            car_por_linea = int(car_por_linea * LETRAS_SEGURIDAD)
        except Exception:
            car_por_linea = LETRAS_MAX_CAR
        car_por_linea = max(LETRAS_MIN_CAR, min(LETRAS_MAX_CAR, car_por_linea))

        lineas = _lineas_texto(texto, car_por_linea)
        alto_original = float(ws.api.Rows(fila).RowHeight)
        # Una sola línea conserva el alto del machote; 2+ líneas crecen a 15
        # pt por línea. Nunca se reduce el alto original.
        alto = max(alto_original, LETRAS_ALTO_LINEA * lineas)
        ws.api.Rows(fila).RowHeight = min(LETRAS_ALTO_MAX, alto)
    except Exception:
        pass


def _xl_fila_valores(valores):
    """
    Aplana lo que devuelve COM para UNA fila a una tupla de escalares.

    Sirve para no depender de la forma exacta que entrega xlwings: una fila
    puede llegar como tupla plana de 8 valores o como tupla anidada
    `((v1, …, v8),)`, y en los dos casos se necesita recorrer los 8.
    """
    if not isinstance(valores, (tuple, list)):
        return (valores,)
    salida = []
    for v in valores:
        if isinstance(v, (tuple, list)):
            salida.extend(v)
        else:
            salida.append(v)
    return tuple(salida)


def _xl_fila_letras(ws, fila_total, filas):
    """
    Devuelve la fila donde va el "Total en letras", o None si no la halla.

    Se busca con tres señales, en este orden, porque la celda cambió con el
    machote y solo la primera estaba contemplada (y exigía el nombre exacto de
    la macro y en la columna A):

    1. Una fórmula que mencione la macro (`CONVERTIRMONTO`), el monto
       (`QUETZALES`) o el texto de la fila (`EN LETRAS`), sin distinguir
       mayúsculas y en cualquier columna de la fila.
    2. La fila cuya columna A está combinada de ancho (`LETRAS_ANCHO_MIN_COL`
       columnas o más) y lleva una FÓRMULA: en el machote ese renglón combinado
       de A a G es el del "Total en letras". Se saltan la fila del TOTAL y TODAS
       las anteriores, porque ahí está el título A1 del proyecto: en las hojas
       ANE - VAR ese A1 es una referencia con fórmula al ANE (`='ANE. - X'!$A$1`)
       y también está combinado, así que sin el salto esta señal devolvería la
       fila 1 y se escribiría el monto encima del título del proyecto.
    3. Cualquier celda con error de nombre (`#NOMBRE?` / `#NAME?`), que es lo
       que deja la macro al guardar el libro sin macros. Va al final porque
       depende de que Excel ya haya recalculado.

    Se lee UNA fila por llamada (8 columnas) y se aplana con `_xl_fila_valores`:
    la lectura en bloque de una sola columna no es fiable.
    """
    combinada = None
    for r in range(1, filas + 1):
        if r <= fila_total:
            continue
        try:
            formulas = _xl_fila_valores(
                ws.range((r, 1), (r, 8)).options(ndim=2).formula)
        except Exception:
            formulas = ()
        for f in formulas:
            if isinstance(f, str):
                u = f.upper()
                if ("CONVERTIRMONTO" in u or "QUETZALES" in u
                        or "EN LETRAS" in u):
                    return r
        if combinada is None:
            try:
                api = ws.range(f"A{r}").api
                if (bool(api.MergeCells)
                        and int(api.MergeArea.Columns.Count) >= LETRAS_ANCHO_MIN_COL
                        and str(api.MergeArea.Formula or "").startswith("=")):
                    combinada = r
            except Exception:
                pass
    if combinada is not None:
        return combinada
    for r in range(1, filas + 1):
        if r <= fila_total:
            continue
        try:
            valores = _xl_fila_valores(
                ws.range((r, 1), (r, 8)).options(ndim=2).value)
        except Exception:
            continue
        for v in valores:
            if isinstance(v, str):
                u = v.strip().upper()
                if u.startswith("#") and ("NOMBRE" in u or "NAME" in u):
                    return r
    return None


def _xl_reemplazar_letras(wb, estado):
    """El machote usa una macro para el monto en letras; el .xlsx final no lleva macros."""
    for base in estado.bases():
        for hoja in EDITABLES:
            if not base["hojas"][hoja]:
                continue
            ws = wb.sheets[estado.titulo(base, hoja)]
            fila_total = _xl_fila_total(ws)
            if fila_total is None:
                continue
            try:
                total = D(ws.range(f"G{fila_total}").value)
            except Exception:
                continue
            # Última fila REAL de la hoja (UsedRange.Row no siempre es 1).
            usada = ws.api.UsedRange
            filas = int(usada.Row) + int(usada.Rows.Count) - 1
            fila = _xl_fila_letras(ws, fila_total, filas)
            if fila is None:
                print(f"[letras] no se encontró la fila del Total en letras "
                      f"en la hoja {ws.name} (TOTAL en la fila {fila_total})")
                continue
            ws.range(f"A{fila}").value = "Total en letras: " + total_en_letras(total)
            _xl_alto_letras(ws, fila)


def _xl_ordenar_hojas(wb, estado):
    # Si la plantilla no trae INICIO, se usa CODIGOS como primer ancla.
    try:
        anterior = wb.sheets["INICIO"]
    except Exception:
        try:
            anterior = wb.sheets["CODIGOS"]
        except Exception:
            anterior = None
    for base in estado.bases():
        for hoja in HOJAS:
            if not base["hojas"][hoja]:
                continue
            try:
                ws = wb.sheets[estado.titulo(base, hoja)]
                if anterior is not None and ws.name != anterior.name:
                    ws.api.Move(After=anterior.api)
                anterior = ws
            except Exception:
                continue


def _xl_eliminar_bases(wb, nombres):
    """
    Quita del archivo final las hojas base copiadas del machote.

    Solo deben quedar las hojas del proyecto/Ramal que el usuario activó.
    El machote original sigue disponible para la próxima exportación, así que
    no hace falta conservar copias `_BASE_n` en el `.xlsx` final.
    """
    try:
        # Algunas operaciones del libro pueden devolver DisplayAlerts a True;
        # sin esto Excel cancelaría el borrado de hojas con datos.
        wb.app.api.DisplayAlerts = False
    except Exception:
        pass
    for nombre in nombres:
        try:
            wb.sheets[nombre].api.Delete()
        except Exception:
            pass


def _xl_aplicar_visibilidad(wb, estado):
    """El Excel conserva exactamente las hojas visibles/ocultas de la app."""
    for base in estado.bases():
        for hoja in HOJAS:
            if not base["hojas"].get(hoja):
                continue
            try:
                ws = wb.sheets[estado.titulo(base, hoja)]
                ws.api.Visible = -1 if hoja in base["visibles"] else 0
            except Exception:
                pass


def _xl_activar_visible(wb, estado):
    """Activa una hoja visible antes de quitar las hojas base."""
    for base in estado.bases():
        for hoja in HOJAS:
            if base["hojas"].get(hoja) and hoja in base["visibles"]:
                try:
                    wb.sheets[estado.titulo(base, hoja)].activate()
                    return
                except Exception:
                    pass
    for nombre in ("INICIO", "CODIGOS"):
        try:
            wb.sheets[nombre].activate()
            return
        except Exception:
            pass
    try:
        wb.sheets[0].activate()
    except Exception:
        pass


def exportar_xlsx(estado, ruta, solo_hojas=None):
    """
    Genera el producto final .xlsx SIN macros a partir del machote.

    · Copia las hojas base COMUNIDAD/RAMAL y las renombra por proyecto.
    · Inserta SOLO los códigos que tienen cantidad en las tablas A_*/AV_*.
    · La VARIACIÓN lleva la unión de códigos de ANE y ANE-VAR.
    · Las tablas vacías NO se eliminan: se les borran las filas y se oculta
      su sección. Así las fórmulas TOTAL y el diseño del machote siguen sanos.
    """
    import tempfile
    if not os.path.exists(PLANTILLA_XLSX):
        raise FileNotFoundError(f"No existe la plantilla: {PLANTILLA_XLSX}")

    salida_tmp = os.path.abspath(ruta) + ".nuevo.xlsx"
    try:
        if os.path.exists(salida_tmp):
            os.remove(salida_tmp)
    except PermissionError:
        raise PermissionError("El archivo de salida está abierto en Excel.")

    # El libro de TRABAJO se abre en la carpeta temporal del sistema: si se
    # creara junto al destino (OneDrive, unidad de red), la sincronización
    # competiría con Excel. `salida_tmp` sí queda junto a `ruta` porque
    # os.replace falla entre volúmenes distintos.
    fd, temp = tempfile.mkstemp(suffix=".xlsm", dir=tempfile.gettempdir())
    os.close(fd)
    shutil.copyfile(PLANTILLA_XLSX, temp)

    app, pythoncom = _abrir_excel()
    wb = None
    try:
        wb = app.books.open(temp)
        _xl_limpiar_hojas(wb)
        # La hoja CODIGOS viaja en el archivo final: lleva los códigos nuevos
        # (Local) o los de la plantilla (Global) para que las BUSCARX los hallen.
        _xl_escribir_bd(wb, estado.bd)
        _xl_renombrar_tablas_base(wb)
        mapa_bases = _xl_renombrar_hojas_base(wb)
        mapa_com = {h: mapa_bases[BASE_COMUNIDAD[h]] for h in HOJAS
                    if BASE_COMUNIDAD[h] in mapa_bases}
        mapa_ram = {h: mapa_bases[BASE_RAMAL[h]] for h in HOJAS
                    if BASE_RAMAL[h] in mapa_bases}
        _xl_config_inicio(wb, estado)

        for base in estado.proyectos:
            for hoja in HOJAS:
                if base["hojas"][hoja]:
                    _xl_crear_hoja(wb, estado, mapa_com[hoja], estado.titulo(base, hoja),
                                   _xl_token(base["abrev"]), base, hoja, mapa_com)

        if estado.ramal_visible():
            for hoja in HOJAS:
                if estado.ramal["hojas"][hoja]:
                    _xl_crear_hoja(wb, estado, mapa_ram[hoja], estado.titulo(estado.ramal, hoja),
                                   _xl_token(estado.ramal["abrev"]), estado.ramal, hoja, mapa_ram)

        for base in estado.bases():
            for hoja in EDITABLES:
                if base["hojas"][hoja]:
                    _xl_llenar_ane(wb, estado, base, hoja)

        for base in estado.bases():
            if base["hojas"]["VARIACION"]:
                _xl_llenar_variacion(wb, estado, base)

        # OC / OTS / ATE: cada hoja recibe solo las filas de su tipo, con las
        # mismas fórmulas de tabla que VARIACIÓN.
        for base in estado.bases():
            for hoja in ("OC", "OTS", "ATE"):
                if base["hojas"][hoja]:
                    _xl_llenar_documento(wb, estado, base, hoja)

        if estado.ramal_visible():
            _xl_formulas_ramal(wb, estado)

        _xl_aplicar_redondeos(wb, estado)
        _xl_ordenar_hojas(wb, estado)
        # Activar una hoja que quedará visible ANTES de ocultar el resto
        # (Excel no permite ocultar la hoja activa).
        _xl_activar_visible(wb, estado)
        # Conservar el estado visible/oculto de cada hoja del proyecto.
        _xl_aplicar_visibilidad(wb, estado)
        _xl_eliminar_bases(wb, list(mapa_bases.values()))
        app.calculate()
        _xl_reemplazar_letras(wb, estado)
        app.calculate()
        _xl_autoalto_descripciones(wb, estado)
        wb.api.SaveAs(salida_tmp, FileFormat=51)   # xlOpenXMLWorkbook (.xlsx)
        wb.close()
        wb = None
        os.replace(salida_tmp, ruta)
    finally:
        try:
            if wb is not None:
                wb.close()
        finally:
            _cerrar_excel(app, pythoncom)
            for tmp in (temp, salida_tmp):
                try:
                    if os.path.exists(tmp):
                        os.remove(tmp)
                except OSError:
                    pass

    estado.titulos_borrar.clear()
    return {estado.titulo(b, h): estado.total(estado.filas(b, h))
            for b in estado.bases() for h in ("ANE", "ANE_VAR", "VARIACION")
            if b["hojas"][h]}


# ═════════════════════════════════════════════════════════════════════
#  ACTUALIZAR (🔄)
# ═════════════════════════════════════════════════════════════════════
def _xl_fila_total(ws):
    """
    Devuelve la fila donde está "TOTAL", o None si no aparece.

    Se lee celda por celda (columna A y luego B, fila a fila), que es la forma
    que estaba verificada con Excel real. La lectura en bloque de A1:B{n} se
    retiró porque con el machote actual la exportación dejó de encontrar el
    total y dejó de reemplazar el "Total en letras".

    La fila del "Total en letras" también empieza por TOTAL; se descarta a
    propósito.

    ANE/VARIACIÓN ponen TOTAL en A; OC/OTS/ATE lo ponen en B.
    """
    try:
        filas = int(ws.api.UsedRange.Rows.Count)
    except Exception:
        return None
    for r in range(1, filas + 1):
        for col in ("A", "B"):
            try:
                v = ws.range(f"{col}{r}").value
            except Exception:
                continue
            if isinstance(v, str):
                t = v.strip().upper()
                if t.startswith("TOTAL") and "EN LETRAS" not in t:
                    return r
    return None


def _xl_grupos_desde_wb(wb):
    """
    Agrupa las hojas del proyecto por abreviatura.

    Devuelve una lista en el orden en que aparecen las hojas en el libro:
    [{"token": "COMUNIDAD_1", "hojas": {"ANE": ws, ...}}, ...]
    """
    prefijos = sorted(((p, h) for h, p in PREFIJOS_HOJA.items()),
                      key=lambda x: len(x[0]), reverse=True)
    grupos = {}
    for ws in wb.sheets:
        nombre = ws.name
        for prefijo, hoja in prefijos:
            if not nombre.startswith(prefijo):
                continue
            token = nombre[len(prefijo):].strip()
            if not token:
                break
            clave = token.upper()
            grupo = grupos.setdefault(clave, {"token": token, "hojas": {}})
            grupo["hojas"][hoja] = ws
            break
    return list(grupos.values())


def _xl_firma_ramal(grupo):
    """
    Calcula qué tan probable es que un grupo de hojas sea el Ramal.

    El Ramal ya no se puede reconocer por el nombre de la hoja: el usuario
    puede renombrarlo (por ejemplo RAMAL2). Su firma está en las tablas:
    en ANE / ANE-VAR la columna 'Cantidad a ejecutar' lleva fórmulas
    SUMIF / SUMAR.SI que suman las tablas de los proyectos, mientras que un
    proyecto escribe cantidades numéricas en esa misma columna.

    Se revisan TODAS las tablas de ANE/ANE-VAR, sin filtrar por nombre, para
    que una tabla renombrada a mano no provoque un falso negativo. Devuelve
    1000 solo si MÁS DE LA MITAD de las celdas revisadas contienen
    SUMIF/SUMAR.SI; una sola fórmula suelta dentro de un proyecto no basta.
    En cualquier otro caso devuelve 0.
    """
    try:
        revisadas = 0
        con_sumif = 0
        for hoja in EDITABLES:
            ws = grupo["hojas"].get(hoja)
            if ws is None:
                continue
            for tabla in ws.tables:
                # Cada tabla se procesa por separado: una tabla dañada no debe
                # abortar la lectura de las demás.
                try:
                    col = _xl_col_cantidad(tabla)
                    try:
                        cuerpo = tabla.api.ListColumns(col).DataBodyRange
                    except Exception:
                        continue
                    if cuerpo is None:
                        continue
                    try:
                        formulas = _xl_valores(cuerpo.Formula)
                    except Exception:
                        try:
                            formulas = _xl_valores(cuerpo.FormulaLocal)
                        except Exception:
                            continue
                    for fila in formulas:
                        texto = str((fila[0] if fila else "") or "").upper()
                        revisadas += 1
                        if "SUMIF(" in texto or "SUMAR.SI(" in texto:
                            con_sumif += 1
                except Exception:
                    continue
        if revisadas and con_sumif * 2 > revisadas:
            return 1000
        return 0
    except Exception:
        return 0


def _xl_indice_ramal(tokens, firmas):
    """
    Elige cuál de los grupos del libro es el Ramal.

    Es una función pura (no toca Excel) para poder probarla por separado:
    recibe la abreviatura de cada grupo y su firma, y devuelve el índice del
    Ramal o None. Con un solo grupo nunca hay Ramal; un proyecto vacío no
    debe perderse, así que ante firma igual se prefiere el nombre RAMAL y,
    si aun así hay empate, no se marca ninguno.
    """
    if len(tokens) != len(firmas):
        return None
    if len(tokens) <= 1:
        return None
    candidatos = [(firma, i) for i, firma in enumerate(firmas) if firma > 0]
    if candidatos:
        mayor = max(firma for firma, _ in candidatos)
        mejores = [i for firma, i in candidatos if firma == mayor]
        if len(mejores) == 1:
            return mejores[0]
        llamados = [i for i in mejores
                    if _xl_token(tokens[i]).upper() == "RAMAL"]
        if len(llamados) == 1:
            return llamados[0]
        return None
    for i, token in enumerate(tokens):
        if _xl_token(token).upper() == "RAMAL":
            return i
    return None


def _xl_nombre_desde_hoja(ws):
    """Lee el nombre del proyecto en A1; el texto del machote significa vacío."""
    if ws is None:
        return ""
    try:
        valor = ws.range("A1").value
    except Exception:
        return ""
    texto = str(valor or "").strip()
    if not texto or norm(texto).startswith("nombre del proyecto"):
        return ""
    return texto


def _xl_nombre_desde_config(wb, token):
    """Respaldo: nombre guardado en la tabla Configuration de INICIO."""
    try:
        ws = wb.sheets["INICIO"]
        tabla = _xl_tabla(ws, "Configuration")
        if tabla is None or tabla.api.DataBodyRange is None:
            return ""
        for fila in _xl_valores(tabla.api.DataBodyRange.Value):
            if len(fila) < 3:
                continue
            if _xl_token(str(fila[2] or "")) == token:
                return str(fila[1] or "").strip()
    except Exception:
        pass
    return ""


def _xl_hoja_visible(ws):
    try:
        return int(ws.api.Visible) == -1
    except Exception:
        return True


def _xl_leer_items(estado, base, hoja, token, externos, ws):
    """Carga las filas de ANE / ANE-VAR conservando el orden de la tabla."""
    base["items"][hoja] = _vacio()
    es_ramal = base is estado.ramal
    titulo = estado.titulo(base, hoja)
    codigos_bd = {c["codigo"] for c in estado.bd}
    # Se calcula UNA sola vez: al leer el Ramal se recorría todos los proyectos
    # por CADA fila, lo que hacía la importación cuadrática.
    cant_proy = estado._cant_proyectos(hoja) if es_ramal else None
    fila_modo = _xl_fila_total(ws)
    modo_old = None
    if fila_modo is not None:
        texto_modo = ws.range(f"H{fila_modo}").value
        modo_old = MODO_POR_TEXTO.get(texto_modo)
    for linea, _, _ in LINEAS:
        tabla = _xl_tabla(ws, _xl_nombre_tabla(hoja, token, linea))
        if tabla is None or tabla.api.DataBodyRange is None:
            continue
        try:
            formulas = _xl_valores(tabla.api.ListColumns(7).DataBodyRange.Formula)
        except Exception:
            formulas = []
        for i, fila in enumerate(_xl_valores(tabla.api.DataBodyRange.Value)):
            cod = _cod(fila[0]) if fila and fila[0] not in (None, "") else ""
            if not cod:
                continue
            cantidad = D(fila[5]) if len(fila) > 5 else Decimal(0)
            if cantidad <= 0:
                continue
            if es_ramal:
                # La columna 6 del Ramal trae «proyectos + propia»; lo propio se
                # obtiene restando la suma de los proyectos, que se leyeron
                # antes. Si da 0 queda el override de redondeo, como hasta hoy.
                cantidad = (cantidad
                            - cant_proy[linea].get(cod, Decimal(0))
                            ).quantize(Decimal("0.001"))
                if cantidad < 0:
                    cantidad = Decimal(0)
            formula = formulas[i][0] if i < len(formulas) else ""
            s_formula = str(formula or "").upper()
            if "CHOOSE(" in s_formula and "MATCH(" in s_formula:
                modo = modo_old or MODO_DEFECTO
            else:
                modo = _xl_modo_desde_formula(formula)
            base["items"][hoja][linea].append(
                {"codigo": cod, "cantidad": str(cantidad), "redondeo": modo})
            if cod not in codigos_bd:
                externos.append((
                    titulo, "-", cod,
                    str(fila[2] or "") if len(fila) > 2 else "",
                    str(fila[3] or "") if len(fila) > 3 else "",
                    str(D(fila[4])) if len(fila) > 4 else "0",
                ))
        if base["items"][hoja][linea]:
            # Se marca orden manual solo si el orden leído no es el de
            # CODIGOS_SICOA (hubo arrastres); si coincide, el próximo
            # material nuevo se inserta en su posición de la base.
            copia = list(base["items"][hoja][linea])
            estado.ordenar_por_codigo(copia)
            if ([it["codigo"] for it in copia]
                    != [it["codigo"] for it in base["items"][hoja][linea]]):
                clave = f"{hoja}:{linea}"
                if clave not in base["orden_manual"]:
                    base["orden_manual"].append(clave)


def _xl_leer_modos_doc(estado, base, hojas_ws):
    """Lee el redondeo elegido en las columnas OC/OTS/ATE de VARIACIÓN."""
    base["redondeo_doc"] = {}
    ws = hojas_ws.get("VARIACION")
    if ws is None:
        return
    token = _xl_token(base["abrev"])
    for linea, _, _ in LINEAS:
        tabla = _xl_tabla(ws, _xl_nombre_tabla("VARIACION", token, linea))
        if tabla is None or tabla.api.DataBodyRange is None:
            continue
        try:
            valores = _xl_valores(tabla.api.DataBodyRange.Value)
        except Exception:
            continue
        formulas = {}
        for nombre, col in (("OC", 13), ("OTS", 14), ("ATE", 15)):
            try:
                formulas[nombre] = _xl_valores(
                    tabla.api.ListColumns(col).DataBodyRange.Formula)
            except Exception:
                formulas[nombre] = []
        for i, fila in enumerate(valores):
            cod = _cod(fila[0]) if fila and fila[0] not in (None, "") else ""
            if not cod:
                continue
            tipo = str(fila[9] or "").strip().upper() if len(fila) > 9 else ""
            if tipo not in formulas:
                continue
            lista = formulas.get(tipo) or []
            formula = lista[i][0] if i < len(lista) else ""
            base["redondeo_doc"][f"{linea}:{cod}"] = _xl_modo_desde_formula(formula)


def cargar_proyecto_desde_excel(estado, ruta):
    """
    Reconstruye el estado completo desde un Excel exportado por el programa.

    No usa JSON ni hoja auxiliar: el nombre y el color salen de las hojas, la
    abreviatura del nombre de las hojas/tablas, el orden de las filas de las
    tablas y la visibilidad de cada pestaña.
    """
    app, pythoncom = _abrir_excel()
    wb = None
    try:
        wb = app.books.open(os.path.abspath(ruta))
        app.calculate()
        grupos = _xl_grupos_desde_wb(wb)
        if not grupos:
            raise ValueError("El archivo no contiene hojas de proyecto (ANE, OC, etc.).")

        # El Ramal puede estar renombrado (RAMAL2, etc.), así que se reconoce
        # por las fórmulas SUMIF de la columna 'Cantidad a ejecutar'.
        firmas = [_xl_firma_ramal(g) for g in grupos]
        i_ramal = _xl_indice_ramal([g["token"] for g in grupos], firmas)

        estado.reset()
        # La BD del propio archivo manda: un anexo guardado en Local puede traer
        # códigos que la plantilla todavía no tiene. Solo si el archivo no trae
        # CODIGOS_SICOA se cae a la tabla de la plantilla.
        bd_archivo = _xl_leer_bd(wb)
        estado.bd = bd_archivo or cargar_bd_inicial()
        estado.bd_local = False
        # La BD del archivo puede traer códigos nuevos: el índice en caché
        # tiene que reflejarlos antes de leer las tablas.
        estado.reindexar_bd()
        estado.archivo = os.path.abspath(ruta)
        estado.titulos_borrar.clear()
        estado.totales_previos = {}

        proyectos, ramal, pares = [], None, []
        for i, grupo in enumerate(grupos):
            token = _xl_token(grupo["token"])
            es_ramal = (i == i_ramal)
            base = estado._base(token, PALETA[i % len(PALETA)], es_ramal)
            base["abrev"] = token
            base["hojas"] = {h: h in grupo["hojas"] for h in HOJAS}
            base["visibles"] = [h for h in HOJAS
                                if h in grupo["hojas"]
                                and _xl_hoja_visible(grupo["hojas"][h])]
            base["items"] = {h: _vacio() for h in EDITABLES if h in grupo["hojas"]}
            base["orden_manual"] = []

            ws_nombre = grupo["hojas"].get("ANE") or next(iter(grupo["hojas"].values()))
            base["nombre"] = _xl_nombre_desde_hoja(ws_nombre)
            if not base["nombre"]:
                base["nombre"] = _xl_nombre_desde_config(wb, token)

            color = None
            for hoja in HOJAS:
                if hoja in grupo["hojas"]:
                    color = _xl_color_hex(grupo["hojas"][hoja], None)
                    if color:
                        break
            base["color"] = color or PALETA[i % len(PALETA)]

            if es_ramal:
                ramal = base
            else:
                proyectos.append(base)
            pares.append((base, grupo["hojas"]))

        if not proyectos:
            raise ValueError("El archivo no contiene un proyecto (hoja ANE).")

        estado.proyectos = proyectos
        estado.ramal = ramal or estado._base("RAMAL", "#7f8c8d", True)
        estado.ramal_presente = ramal is not None

        externos = []
        # Dos pasadas: los proyectos primero, porque las filas propias del
        # Ramal se calculan restando la suma de proyectos a la cantidad leída.
        for base, hojas_ws in [par for par in pares if par[0] is not estado.ramal]:
            for hoja in EDITABLES:
                if base["hojas"].get(hoja):
                    _xl_leer_items(estado, base, hoja, base["abrev"],
                                   externos, hojas_ws[hoja])
            _xl_leer_modos_doc(estado, base, hojas_ws)
        for base, hojas_ws in [par for par in pares if par[0] is estado.ramal]:
            for hoja in EDITABLES:
                if base["hojas"].get(hoja):
                    _xl_leer_items(estado, base, hoja, base["abrev"],
                                   externos, hojas_ws[hoja])
            _xl_leer_modos_doc(estado, base, hojas_ws)

        estado._sincronizar_ramal()
        estado._limpiar_redondeos_ramal()
        estado.asegurar_visibles()
        estado.cambios = False
    finally:
        try:
            if wb is not None:
                wb.close()
        finally:
            _cerrar_excel(app, pythoncom)


def _xl_importar_desde_excel(estado, ruta):
    """
    Solo lectura: verifica el archivo y no modifica el estado ni el libro.

    Abre el Excel, recalcula y lee las tablas ANE/ANE-VAR únicamente para
    detectar diferencias (códigos externos, filas duplicadas y hojas activas
    ausentes) y para leer los totales de cada hoja. No reimporta filas, no
    toca el orden manual, no cambia la BD de la sesión, no reescribe modos y
    nunca guarda el libro.
    """
    res = dict(totales={}, externos=[], duplicados=[], faltantes=[], bd=0,
               verificacion=True)
    app, pythoncom = _abrir_excel()
    wb = None
    try:
        wb = app.books.open(os.path.abspath(ruta))
        app.calculate()
        # La tabla CODIGOS_SICOA del ARCHIVO solo se lee: sirve para saber
        # cuántos códigos trae y para detectar componentes externos. La BD de
        # la sesión NO se reemplaza (Actualizar ya no reimporta al estado).
        bd_archivo = _xl_leer_bd(wb)
        res["bd"] = len(bd_archivo)
        hojas = {ws.name for ws in wb.sheets}
        codigos = {c["codigo"] for c in (bd_archivo or estado.bd)}

        for base in estado.proyectos:
            for hoja in EDITABLES:
                if not base["hojas"][hoja]:
                    continue
                titulo = estado.titulo(base, hoja)
                if titulo not in hojas:
                    # Se reporta la ausencia; la memoria del proyecto se deja
                    # intacta para no perder las filas que el usuario capturó.
                    res["faltantes"].append(titulo)
                    continue
                ws = wb.sheets[titulo]
                token = _xl_token(base["abrev"])
                for linea, _, _ in LINEAS:
                    tabla = _xl_tabla(ws, _xl_nombre_tabla(hoja, token, linea))
                    if tabla is None or tabla.api.DataBodyRange is None:
                        continue
                    vistos = set()            # un código una sola vez por tabla
                    for fila in _xl_valores(tabla.api.DataBodyRange.Value):
                        cod = _cod(fila[0]) if fila and fila[0] not in (None, "") else ""
                        if not cod:
                            continue
                        qty = D(fila[5]) if len(fila) > 5 else Decimal(0)
                        if qty <= 0:
                            continue
                        if cod in vistos:
                            # Fila repetida en el mismo archivo: solo se
                            # reporta, no se descarta nada en la app.
                            res["duplicados"].append(
                                (titulo, linea, cod, str(qty)))
                            continue
                        vistos.add(cod)
                        if cod not in codigos:
                            res["externos"].append((
                                titulo, "-", cod,
                                str(fila[2] or "") if len(fila) > 2 else "",
                                str(fila[3] or "") if len(fila) > 3 else "",
                                str(D(fila[4])) if len(fila) > 4 else "0",
                            ))

        for base in estado.bases():
            for hoja in HOJAS:
                if not base["hojas"][hoja]:
                    continue
                titulo = estado.titulo(base, hoja)
                if titulo not in hojas:
                    continue
                ws = wb.sheets[titulo]
                fila = _xl_fila_total(ws)
                if fila is None:
                    continue
                if hoja in ("ANE", "ANE_VAR"):
                    col = "G"
                elif hoja == "VARIACION":
                    # El total de VARIACIÓN es la suma de las variaciones
                    # (columnas OC/OTS/ATE), no el total nuevo de la columna I.
                    try:
                        res["totales"][titulo] = redondear_monto(
                            D(ws.range(f"M{fila}").value) +
                            D(ws.range(f"N{fila}").value) +
                            D(ws.range(f"O{fila}").value))
                    except Exception:
                        res["totales"][titulo] = Decimal(0)
                    continue
                else:
                    col = "M"
                try:
                    res["totales"][titulo] = redondear_monto(
                        D(ws.range(f"{col}{fila}").value))
                except Exception:
                    res["totales"][titulo] = Decimal(0)
    finally:
        try:
            if wb is not None:
                wb.close()
        finally:
            _cerrar_excel(app, pythoncom)
    return res


def ejecutar_actualizacion(estado, ruta):
    """Verificación de solo lectura: abre el Excel, calcula y compara.

    No exporta, no reimporta filas al estado, no borra ni agrega hojas y no
    marca cambios. Solo devuelve las diferencias encontradas para el reporte.
    """
    res = dict(totales=[], externos=[], duplicados=[], faltantes=[],
               actualizadas=[], bd=0, verificacion=True, error=None)
    try:
        datos = _xl_importar_desde_excel(estado, ruta)
        for titulo, nuevo in datos["totales"].items():
            prev = estado.totales_previos.get(titulo)
            if prev is not None and redondear_monto(D(prev)) != redondear_monto(nuevo):
                res["totales"].append((titulo, str(prev), str(nuevo)))
        res["externos"] = datos["externos"]
        res["duplicados"] = datos["duplicados"]
        res["faltantes"] = datos["faltantes"]
        res["bd"] = datos["bd"]
    except PermissionError:
        res["error"] = ("No se pudo leer el archivo: ciérrelo en Excel "
                        "o verifique que no esté bloqueado.")
    except Exception as e:
        res["error"] = str(e)
    return res


class ActualizarWorker(QThread):
    listo = Signal(dict)

    def __init__(self, estado, ruta):
        super().__init__()
        self.estado, self.ruta = estado, ruta

    def run(self):
        self.listo.emit(ejecutar_actualizacion(self.estado, self.ruta))


# ═════════════════════════════════════════════════════════════════════
#  TEMA DÍA / NOCHE  (colores tomados de RAIZ.py)
# ═════════════════════════════════════════════════════════════════════
class Tema:
    oscuro = True


def css_tema(oscuro):
    if oscuro:
        bg, fg, inp, brd, rb, hov, grp, gbrd, sel = ("#2c3e50", "#ecf0f1", "#34495e", "#7f8c8d",
                                                     "#34495e", "#7f8c8d", "#34495e", "#7f8c8d", "#2980b9")
        alt = "#3b5068"
        dis_bg, dis_fg = "#566573", "#bdc3c7"
    else:
        bg, fg, inp, brd, rb, hov, grp, gbrd, sel = ("#f5f6fa", "#2c3e50", "#ffffff", "#bdc3c7",
                                                     "#d6eaf8", "#aed6f1", "#ffffff", "#e0e0e0", "#3498db")
        alt = "#eaf3fb"
        dis_bg, dis_fg = "#d5d8dc", "#566573"
    return f"""
    * {{ font-family: 'Segoe UI'; font-size: 8pt; }}
    QWidget {{ background-color: {bg}; color: {fg}; }}
    QScrollArea, QScrollArea > QWidget > QWidget {{ background: transparent; }}
    QLabel {{ background: transparent; }}
    QLabel[rol="titulo"] {{ font-size: 14pt; font-weight: bold; }}
    QLabel[rol="seccion"] {{ font-size: 10pt; font-weight: bold; }}
    QLabel[rol="banda"] {{ background-color: {grp}; border: 1px solid {gbrd}; font-weight: bold;
                           padding: 4px; }}
    QFrame[rol="grupo"] {{ background-color: {grp}; border: 1px solid {gbrd}; border-radius: 8px; }}
    QFrame[rol="grupo"] QLabel {{ background: transparent; border: none; }}
    QLineEdit, QComboBox {{ background-color: {inp}; color: {fg}; border: 2px solid {brd};
                            border-radius: 4px; padding: 4px; min-height: 18px; }}
    QLineEdit:focus, QComboBox:focus {{ border-color: #3498db; }}
    QLineEdit:disabled, QComboBox:disabled {{ background-color: {dis_bg}; color: {dis_fg}; }}
    QComboBox QAbstractItemView {{ background-color: {inp}; color: {fg}; border: 1px solid {brd};
                                  selection-background-color: {sel}; selection-color: white; }}
    QComboBox QAbstractItemView::item {{ background-color: {inp}; color: {fg}; padding: 4px;
                                        border: none; outline: 0; }}
    QComboBox QAbstractItemView::item:hover {{ background-color: {sel}; color: white; }}
    QComboBox QAbstractItemView::item:selected,
    QComboBox QAbstractItemView::item:selected:active,
    QComboBox QAbstractItemView::item:selected:!active {{ background-color: {sel}; color: white;
                                                          border: none; outline: 0; }}
    QListView {{ background-color: {inp}; color: {fg}; selection-background-color: {sel};
                 selection-color: white; }}
    QPushButton {{ background-color: {inp}; color: {fg}; border: 1px solid {brd};
                   border-radius: 4px; padding: 5px 10px; min-height: 18px; }}
    QPushButton:hover {{ background-color: {hov}; }}
    QPushButton:disabled {{ background-color: {dis_bg}; color: {dis_fg}; }}
    QPushButton[rol="primario"] {{ background-color: #3498db; color: white; border: none; font-weight: bold; }}
    QPushButton[rol="primario"]:hover {{ background-color: #2980b9; }}
    QPushButton[rol="exito"] {{ background-color: #27ae60; color: white; border: none; font-weight: bold; }}
    QPushButton[rol="exito"]:hover {{ background-color: #229954; }}
    QPushButton[rol="peligro"] {{ background-color: #e74c3c; color: white; border: none; }}
    QPushButton[rol="peligro"]:hover {{ background-color: #c0392b; }}
    QPushButton[rol="confirmar"] {{ background-color: #16a085; color: white; border: none; font-weight: bold; }}
    QPushButton[rol="confirmar"]:hover {{ background-color: #138d75; }}
    QPushButton[rol="radio"] {{ background-color: {rb}; border: 1px solid {brd}; }}
    QPushButton[rol="radio"]:hover {{ background-color: {hov}; }}
    QPushButton[rol="radio"]:checked {{ background-color: #3498db; color: white; font-weight: bold;
                                        border: 2px solid #21618c; }}
    QPushButton[rol="primario"]:disabled, QPushButton[rol="exito"]:disabled,
    QPushButton[rol="peligro"]:disabled, QPushButton[rol="confirmar"]:disabled,
    QPushButton[rol="radio"]:disabled {{ background-color: {dis_bg}; color: {dis_fg};
                                         border: 1px solid {brd}; }}
    QMenu {{ background-color: {inp}; color: {fg}; border: 1px solid {brd}; }}
    QMenu::item {{ padding: 4px 18px; }}
    QMenu::item:selected {{ background-color: {sel}; color: white; }}
    QTableWidget {{ background-color: {inp}; color: {fg}; gridline-color: {brd}; border: 1px solid {brd};
                    alternate-background-color: {alt}; selection-background-color: {sel};
                    selection-color: white; }}
    QHeaderView::section {{ background-color: {grp}; color: {fg}; border: 1px solid {brd};
                            padding: 3px; font-weight: bold; }}
    QTableWidget::item {{ padding: 2px; }}
    QScrollBar:vertical {{ background: {inp}; width: 12px; }}
    QScrollBar::handle:vertical {{ background: {brd}; border-radius: 5px; min-height: 24px; }}
    QScrollBar:horizontal {{ background: {inp}; height: 12px; }}
    QScrollBar::handle:horizontal {{ background: {brd}; border-radius: 5px; min-width: 24px; }}
    QScrollBar::add-line, QScrollBar::sub-line {{ width: 0; height: 0; }}
    QMessageBox {{ background-color: {bg}; }}
    """


class DelegadoCombo(QStyledItemDelegate):
    """Pinta en azul la fila activa del desplegable, sin depender del estilo nativo."""

    def paint(self, painter, option, index):
        activo = bool(option.state & (QStyle.State_Selected | QStyle.State_MouseOver))
        if not activo:
            super().paint(painter, option, index)
            return
        if Tema.oscuro:
            azul = "#2980b9" if option.state & QStyle.State_Selected else "#3498db"
        else:
            azul = "#3498db" if option.state & QStyle.State_Selected else "#5dade2"
        painter.save()
        painter.fillRect(option.rect, QColor(azul))
        painter.setPen(QColor("white"))
        painter.setFont(option.font)
        texto = str(index.data() or "")
        painter.drawText(option.rect.adjusted(6, 0, -6, 0),
                         Qt.AlignVCenter | Qt.AlignLeft, texto)
        painter.restore()


def css_popup_combo():
    """Estilo de la lista desplegable de un QComboBox, igual al popup de búsqueda."""
    if Tema.oscuro:
        bg, fg, brd, sel = "#34495e", "#ecf0f1", "#7f8c8d", "#2980b9"
    else:
        bg, fg, brd, sel = "#ffffff", "#2c3e50", "#bdc3c7", "#3498db"
    return f"""
    QListView {{ background-color: {bg}; color: {fg}; border: 1px solid {brd};
                 selection-background-color: {sel}; selection-color: white; outline: 0; }}
    QListView::item {{ background-color: {bg}; color: {fg}; padding: 4px;
                       border: none; outline: 0; }}
    QListView::item:hover {{ background-color: {sel}; color: white; }}
    QListView::item:selected, QListView::item:selected:active,
    QListView::item:selected:!active {{ background-color: {sel}; color: white;
                                        border: none; outline: 0; }}
    """


def aplicar_estilo_combo(combo):
    """Fuerza el estilo de la lista desplegable incluso dentro de paneles con estilo propio."""
    vista = combo.view()
    if vista is not None:
        vista.setStyleSheet(css_popup_combo())
        # El color de selección también va en la paleta: algunos estilos nativos
        # dibujan el foco con una línea gris e ignoran el fondo del stylesheet.
        sel = QColor("#2980b9" if Tema.oscuro else "#3498db")
        pal = vista.palette()
        pal.setColor(QPalette.Highlight, sel)
        pal.setColor(QPalette.HighlightedText, QColor("white"))
        vista.setPalette(pal)
        vista.setSelectionMode(QAbstractItemView.SingleSelection)
        vista.setSelectionBehavior(QAbstractItemView.SelectRows)
        vista.setItemDelegate(DelegadoCombo(vista))


def aplicar_tema():
    app = QApplication.instance()
    app.setStyleSheet(css_tema(Tema.oscuro))
    for w in app.allWidgets():
        if isinstance(w, QComboBox):
            aplicar_estilo_combo(w)


def alternar_tema():
    Tema.oscuro = not Tema.oscuro
    aplicar_tema()


def icono_tema():
    return "☀️" if Tema.oscuro else "🌙"


_ICONO_APP = None


def icono_app():
    """Icono de la aplicación (barra de tareas y barra de título)."""
    global _ICONO_APP
    if _ICONO_APP is None:
        _ICONO_APP = QIcon(RUTA_LOGO) if os.path.exists(RUTA_LOGO) else QIcon()
    return _ICONO_APP


def _configurar_app_id_windows():
    """Evita que Windows use el icono de Python en la barra de tareas."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "WuMDrex.Anexos.1")
    except Exception:
        pass


# ═════════════════════════════════════════════════════════════════════
#  HELPERS DE UI
# ═════════════════════════════════════════════════════════════════════
def boton(texto, rol=None, tip=None, ancho=None):
    b = QPushButton(texto)
    if rol:
        b.setProperty("rol", rol)
    if tip:
        b.setToolTip(tip)
    if ancho:
        b.setFixedWidth(ancho)
    b.setCursor(Qt.PointingHandCursor)
    return b


def etiqueta(texto, rol=None):
    l = QLabel(texto)
    if rol:
        l.setProperty("rol", rol)
    return l


def grupo():
    f = QFrame()
    f.setProperty("rol", "grupo")
    return f


def estilo_color(color):
    return (f"QPushButton {{ background-color: {color}; color: {contraste(color)}; "
            f"border: 1px solid #21618c; font-weight: bold; }}")


def resaltar_botones_fila(t):
    """Pinta en blanco el texto de los botones de las filas seleccionadas.

    Los cellWidget no reciben selection-color de la tabla; hay que reflejar
    la selección con una propiedad dinámica para que el QSS la aplique.
    """
    filas = {i.row() for i in t.selectionModel().selectedIndexes()}
    for b in t.findChildren(QPushButton):
        if b.property("rol") != "redondeo":
            continue
        activo = b.property("fila") in filas
        if bool(b.property("sel")) == activo:
            continue
        b.setProperty("sel", activo)
        b.style().unpolish(b)
        b.style().polish(b)
        b.update()


def limpiar(layout):
    while layout.count():
        it = layout.takeAt(0)
        w = it.widget()
        if w:
            w.deleteLater()
        elif it.layout():
            limpiar(it.layout())


def msg(parent, titulo, texto, icono=QMessageBox.Information):
    m = QMessageBox(parent)
    m.setIcon(icono)
    m.setWindowTitle(titulo)
    m.setText(texto)
    m.addButton("Aceptar", QMessageBox.AcceptRole)
    m.exec()


def confirmar(parent, titulo, texto):
    m = QMessageBox(parent)
    m.setIcon(QMessageBox.Question)
    m.setWindowTitle(titulo)
    m.setText(texto)
    si = m.addButton("Sí", QMessageBox.YesRole)
    no = m.addButton("No", QMessageBox.NoRole)
    m.setEscapeButton(no)
    m.setDefaultButton(no)
    m.exec()
    return m.clickedButton() is si


def tres_opciones(parent, titulo, texto, a, b, c):
    """Devuelve 0, 1 o 2 según el botón (c = cancelar / Esc)."""
    m = QMessageBox(parent)
    m.setIcon(QMessageBox.Question)
    m.setWindowTitle(titulo)
    m.setText(texto)
    ba = m.addButton(a, QMessageBox.AcceptRole)
    bb = m.addButton(b, QMessageBox.DestructiveRole)
    bc = m.addButton(c, QMessageBox.RejectRole)
    m.setEscapeButton(bc)
    m.exec()
    return {ba: 0, bb: 1, bc: 2}.get(m.clickedButton(), 2)


class Toast(QFrame):
    """Aviso breve, flotante y sin bloquear la ventana principal."""

    def __init__(self, parent=None):
        super().__init__(parent, Qt.ToolTip | Qt.FramelessWindowHint
                         | Qt.WindowDoesNotAcceptFocus)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.setFocusPolicy(Qt.NoFocus)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(16, 9, 16, 9)
        self.lbl = QLabel("")
        self.lbl.setStyleSheet("color: white; font-weight: bold; background: transparent;")
        lay.addWidget(self.lbl)
        self.timer = QTimer(self)
        self.timer.setSingleShot(True)
        self.timer.timeout.connect(self.hide)
        self.hide()

    def mostrar(self, texto, parent=None, color="#27ae60"):
        self.lbl.setText("✔  " + texto)
        self.setStyleSheet(
            f"QFrame {{ background-color: {color}; border-radius: 8px; }}")
        self.adjustSize()
        if self.width() < 230:
            self.setFixedWidth(230)
        if parent is not None:
            x = parent.width() - self.width() - 24
            y = parent.height() - self.height() - 24
            pos = parent.mapToGlobal(QPoint(max(0, x), max(0, y)))
        else:
            pantalla = QApplication.primaryScreen()
            geo = pantalla.availableGeometry() if pantalla else None
            pos = QPoint(40, 40)
            if geo is not None:
                pos = QPoint(geo.right() - self.width() - 40,
                             geo.bottom() - self.height() - 40)
        self.move(pos)
        self.show()
        self.raise_()
        self.timer.start(2200)


class VentanaBase(QDialog):
    """Ventana con scroll, Esc = retroceder, Alt+B = día/noche, redimensionable."""

    def __init__(self, parent, titulo, tam, minimo):
        super().__init__(parent)
        self.setWindowTitle(titulo)
        self.setWindowIcon(icono_app())
        self.resize(*tam)
        self.setMinimumSize(*minimo)
        self.setWindowFlag(Qt.WindowMaximizeButtonHint, True)
        externo = QVBoxLayout(self)
        externo.setContentsMargins(0, 0, 0, 0)
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)
        cont = QWidget()
        self.cuerpo = QVBoxLayout(cont)
        self.cuerpo.setContentsMargins(10, 10, 10, 10)
        self.cuerpo.setSpacing(8)
        self.scroll.setWidget(cont)
        externo.addWidget(self.scroll)
        QShortcut(QKeySequence("Alt+B"), self).activated.connect(alternar_tema)


# ═════════════════════════════════════════════════════════════════════
#  BUSCADOR CON POPUP (tipo VSCode)   ·   Ctrl+Espacio = todos
# ═════════════════════════════════════════════════════════════════════
class PopupBusqueda(QFrame):
    elegido = Signal(str)
    ALTO_FILA = 22

    def __init__(self):
        super().__init__(None, Qt.ToolTip | Qt.FramelessWindowHint
                         | Qt.WindowDoesNotAcceptFocus)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self.setFocusPolicy(Qt.NoFocus)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        self.lbl_conteo = QLabel("")
        self.lbl_conteo.setFixedHeight(18)
        self.lbl_conteo.setContentsMargins(6, 0, 6, 0)
        f = self.lbl_conteo.font()
        f.setPointSize(max(7, f.pointSize() - 1))
        self.lbl_conteo.setFont(f)
        lay.addWidget(self.lbl_conteo)
        self.tabla = QTableWidget(0, 2)
        self.tabla.setFocusPolicy(Qt.NoFocus)
        self.tabla.horizontalHeader().setVisible(False)
        self.tabla.verticalHeader().setVisible(False)
        self.tabla.verticalHeader().setDefaultSectionSize(self.ALTO_FILA)
        self.tabla.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tabla.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tabla.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tabla.setShowGrid(False)
        self.tabla.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tabla.setColumnWidth(0, 90)
        self.tabla.cellClicked.connect(lambda r, c: self.elegido.emit(self.codigo_en(r)))
        lay.addWidget(self.tabla)

    def codigo_en(self, fila):
        it = self.tabla.item(fila, 0)
        return it.data(Qt.UserRole) if it else ""

    def actual(self):
        r = self.tabla.currentRow()
        return self.codigo_en(r) if r >= 0 else ""

    def mover(self, delta):
        n = self.tabla.rowCount()
        if n:
            self.tabla.selectRow(max(0, min(n - 1, self.tabla.currentRow() + delta)))

    def mostrar(self, campo, filas, max_filas):
        """filas: [(codigo, descripcion, compatible)]"""
        if not filas:
            self.hide()
            return
        # Vaciar primero fuerza a Windows a repintar la lista al reducirse.
        self.tabla.setRowCount(0)
        self.tabla.setRowCount(len(filas))
        self.lbl_conteo.setText(
            f"{len(filas)} coincidencia{'s' if len(filas) != 1 else ''}")
        gris = QColor("#95a5a6")
        for i, (cod, desc, ok) in enumerate(filas):
            a = QTableWidgetItem(cod)
            a.setData(Qt.UserRole, cod)
            b = QTableWidgetItem(desc if ok else desc + "   (no habilitado para este voltaje)")
            if not ok:
                a.setForeground(gris)
                b.setForeground(gris)
            self.tabla.setItem(i, 0, a)
            self.tabla.setItem(i, 1, b)
        self.tabla.selectRow(0)
        self.tabla.scrollToTop()
        w = max(campo.width(), 480)
        h = self.lbl_conteo.height() + min(len(filas), max_filas) * self.ALTO_FILA + 4
        self.setFixedSize(w, h)
        self.move(campo.mapToGlobal(QPoint(0, campo.height() + 2)))
        self.show()
        # En Windows el primer popup puede quedar detrás de la ventana
        # principal hasta que se vuelve a mostrar; subirlo evita que el
        # usuario tenga que escribir un segundo carácter para verlo.
        self.raise_()
        self.tabla.viewport().update()


class CampoBusqueda(QLineEdit):
    """
    proveedor(texto, todos) -> [(codigo, descripcion, compatible)]
    emite `elegido(codigo)` al seleccionar una fila del popup.
    """
    elegido = Signal(str)
    limpiar = Signal()

    def __init__(self, placeholder, proveedor, max_filas=lambda: 8):
        super().__init__()
        self.setPlaceholderText(placeholder)
        self.proveedor = proveedor
        self.max_filas = max_filas
        self.popup = PopupBusqueda()
        self.popup.elegido.connect(self._elegir)
        self.textEdited.connect(lambda t: self._mostrar(t, False))

    def _mostrar(self, texto, todos):
        if not texto.strip() and not todos:
            self.popup.hide()
            return
        self.popup.mostrar(self, self.proveedor(texto, todos), self.max_filas())

    def _elegir(self, codigo):
        self.popup.hide()
        if codigo:
            self.elegido.emit(codigo)

    def keyPressEvent(self, e):
        k, mod = e.key(), e.modifiers()
        visible = self.popup.isVisible()
        if k == Qt.Key_Space and mod & Qt.ControlModifier:
            self.clear()
            self.limpiar.emit()                        # limpia el resto de campos
            self._mostrar("", True)                    # TODOS desde este campo
            e.accept()
        elif visible and k in (Qt.Key_Down, Qt.Key_Up):
            self.popup.mover(1 if k == Qt.Key_Down else -1)
            e.accept()
        elif visible and k in (Qt.Key_Return, Qt.Key_Enter, Qt.Key_Tab):
            self._elegir(self.popup.actual())
            e.accept()
        elif visible and k == Qt.Key_Escape:
            self.popup.hide()                          # Esc cierra solo el popup
            e.accept()
        else:
            super().keyPressEvent(e)
            # Refuerzo para el primer carácter: si el popup no llegó a
            # mostrarse, se intenta de nuevo con el texto ya actualizado.
            if e.text().strip() and self.text().strip() and not self.popup.isVisible():
                self._mostrar(self.text(), False)

    def focusOutEvent(self, e):
        self.popup.hide()
        super().focusOutEvent(e)

    def closeEvent(self, e):
        self.popup.hide()
        super().closeEvent(e)


class ComboLinea(QComboBox):
    """
    Combo de Línea con navegación controlada:
    - Al abrir la lista, resalta la opción actual (o la primera).
    - Flechas arriba/abajo mueven SOLO el resaltado.
    - Enter confirma la opción resaltada y cierra la lista.
    - Escape cierra sin cambiar.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.view().installEventFilter(self)
        self._filtro_app = False
        self._filtro_popup = None

    def _instalar_filtro_popup(self):
        """Refuerza el teclado también en la ventana emergente y en la app."""
        popup = self.view().window()
        if (popup is not None and popup is not self.window() and
                popup is not self._filtro_popup):
            popup.installEventFilter(self)
            self._filtro_popup = popup
        app = QApplication.instance()
        if app is not None and not self._filtro_app:
            app.installEventFilter(self)
            self._filtro_app = True

    def hidePopup(self):
        if self._filtro_popup is not None:
            self._filtro_popup.removeEventFilter(self)
            self._filtro_popup = None
        app = QApplication.instance()
        if self._filtro_app and app is not None:
            app.removeEventFilter(self)
            self._filtro_app = False
        super().hidePopup()

    def showPopup(self):
        super().showPopup()
        self._instalar_filtro_popup()
        if not self.count():
            return
        fila = self.currentIndex()
        if fila < 0:
            fila = 0
        self._resaltar(fila)

    def _resaltar(self, fila):
        fila = max(0, min(fila, self.count() - 1))
        idx = self.model().index(fila, 0)
        vista = self.view()
        vista.setCurrentIndex(idx)
        sel = vista.selectionModel()
        if sel is not None:
            sel.select(idx, QItemSelectionModel.ClearAndSelect | QItemSelectionModel.Rows)
        vista.scrollTo(idx)
        vista.viewport().update()

    def _mover(self, paso):
        if not self.count():
            return
        vista = self.view()
        idx = vista.currentIndex()
        fila = idx.row() if idx.isValid() else self.currentIndex()
        if fila < 0:
            fila = 0 if paso > 0 else self.count() - 1
        else:
            fila = max(0, min(fila + paso, self.count() - 1))
        self._resaltar(fila)

    def _confirmar(self):
        vista = self.view()
        idx = vista.currentIndex()
        fila = idx.row() if idx.isValid() else self.currentIndex()
        if fila < 0 and self.count():
            fila = 0
        if fila >= 0:
            self.setCurrentIndex(fila)
        self.hidePopup()

    def _tecla_lista(self, e):
        if not self.view().isVisible():
            return False
        k = e.key()
        if k == Qt.Key_Up:
            self._mover(-1)
        elif k == Qt.Key_Down:
            self._mover(1)
        elif k in (Qt.Key_Return, Qt.Key_Enter):
            self._confirmar()
        elif k == Qt.Key_Escape:
            self.hidePopup()
        else:
            return False
        e.accept()
        return True

    def keyPressEvent(self, e):
        if self._tecla_lista(e):
            return
        super().keyPressEvent(e)

    def eventFilter(self, obj, e):
        if e.type() == QEvent.KeyPress:
            app = QApplication.instance()
            if (obj is self.view() or obj is self.view().window() or obj is app):
                if self._tecla_lista(e):
                    return True
        return super().eventFilter(obj, e)


# ═════════════════════════════════════════════════════════════════════
#  TABLA con tecla Supr (y arrastre opcional de filas)
# ═════════════════════════════════════════════════════════════════════
class TablaBase(QTableWidget):
    suprimir = Signal()
    filaMovida = Signal(int, int)

    def __init__(self, filas, columnas, arrastrable=False, columnas_no_arrastre=()):
        super().__init__(filas, columnas)
        self.arrastrable = arrastrable
        self.columnas_no_arrastre = set(columnas_no_arrastre)
        self._fila_press = -1
        self._col_press = -1
        self._y_press = 0
        self._fila_destino = -1
        self._arrastrando = False
        self._linea = None
        if arrastrable:
            self._crear_linea()

    def _crear_linea(self):
        if self._linea is not None:
            return
        self._linea = QFrame(self.viewport())
        self._linea.setFixedHeight(3)
        self._linea.setStyleSheet("background-color: #3498db;")
        self._linea.setAttribute(Qt.WA_TransparentForMouseEvents)
        self._linea.hide()

    def permitir_arrastre(self, activo):
        self.arrastrable = activo
        if activo:
            self._crear_linea()
        else:
            self._cancelar_arrastre()

    def _cancelar_arrastre(self):
        self._arrastrando = False
        self._fila_press = -1
        self._col_press = -1
        self._fila_destino = -1
        self.viewport().unsetCursor()
        if self._linea is not None:
            self._linea.hide()

    def _mover_linea(self, dst):
        if self._linea is None or dst < 0 or self.rowCount() == 0:
            return
        y = self.rowViewportPosition(dst)
        if dst > self._fila_press:
            y += self.rowHeight(dst)
        y = max(0, min(y, self.viewport().height() - self._linea.height()))
        self._linea.setGeometry(0, y, self.viewport().width(),
                                self._linea.height())
        self._linea.show()
        self._linea.raise_()

    def mousePressEvent(self, e):
        self._cancelar_arrastre()
        if self.arrastrable and e.button() == Qt.LeftButton:
            idx = self.indexAt(e.position().toPoint())
            if idx.isValid() and idx.column() not in self.columnas_no_arrastre:
                self._fila_press = idx.row()
                self._col_press = idx.column()
                self._y_press = e.position().toPoint().y()
                self._arrastrando = False
        super().mousePressEvent(e)

    def mouseMoveEvent(self, e):
        if (self.arrastrable and self._fila_press >= 0 and
                e.buttons() & Qt.LeftButton and self.rowCount() > 0):
            dy = abs(e.position().toPoint().y() - self._y_press)
            if (not self._arrastrando and
                    dy >= max(6, QApplication.startDragDistance())):
                self._arrastrando = True
                self.viewport().setCursor(Qt.ClosedHandCursor)
            if self._arrastrando:
                dst = self.rowAt(e.position().toPoint().y())
                if dst < 0:
                    dst = self.rowCount() - 1
                self._fila_destino = max(0, min(dst, self.rowCount() - 1))
                self._mover_linea(self._fila_destino)
                e.accept()
                return
        super().mouseMoveEvent(e)

    def mouseReleaseEvent(self, e):
        if self._arrastrando:
            src, dst = self._fila_press, self._fila_destino
            self._cancelar_arrastre()
            if src >= 0 and dst >= 0 and src != dst:
                self.filaMovida.emit(src, dst)
            e.accept()
            return
        self._fila_press = -1
        self._col_press = -1
        super().mouseReleaseEvent(e)

    def keyPressEvent(self, e):
        if e.key() == Qt.Key_Delete:
            self.suprimir.emit()
        elif (e.modifiers() & Qt.AltModifier and
              e.key() in (Qt.Key_Up, Qt.Key_Down)):
            r = self.currentRow()
            destino = r - 1 if e.key() == Qt.Key_Up else r + 1
            if 0 <= r < self.rowCount() and 0 <= destino < self.rowCount():
                self.filaMovida.emit(r, destino)
            e.accept()
        else:
            super().keyPressEvent(e)


# ═════════════════════════════════════════════════════════════════════
#  CONFIGURACIÓN
# ═════════════════════════════════════════════════════════════════════
class DlgConfig(VentanaBase):
    def __init__(self, parent, estado):
        super().__init__(parent, "Configuración", (480, 420), (420, 320))
        self.estado = estado
        self.cuerpo.addWidget(etiqueta("Configuración", "titulo"))
        self.proy_lay = QVBoxLayout()
        self.ramal_lay = QVBoxLayout()
        self.cuerpo.addWidget(etiqueta("Proyectos", "seccion"))
        self.cuerpo.addLayout(self.proy_lay)
        self.sep = QFrame()
        self.sep.setFrameShape(QFrame.HLine)
        self.cuerpo.addWidget(self.sep)
        self.lbl_ramal = etiqueta("Ramal", "seccion")
        self.cuerpo.addWidget(self.lbl_ramal)
        self.cuerpo.addLayout(self.ramal_lay)
        self.cuerpo.addStretch()
        self.construir()

    def construir(self):
        limpiar(self.proy_lay)
        limpiar(self.ramal_lay)
        for base in self.estado.proyectos:
            f = QHBoxLayout()
            b = boton(base["abrev"], tip="Editar proyecto")
            b.setStyleSheet(estilo_color(base["color"]))
            b.setMinimumWidth(180)
            b.clicked.connect(partial(self.editar, base))
            pin = boton("🖌", tip="Color de la pestaña", ancho=36)
            pin.clicked.connect(partial(self.color, base))
            pap = boton("🗑", "peligro", "Eliminar proyecto", 36)
            pap.clicked.connect(partial(self.eliminar, base))
            for w in (b, pin, pap):
                f.addWidget(w)
            f.addStretch()
            self.proy_lay.addLayout(f)
        # fila para crear uno nuevo
        f = QHBoxLayout()
        nuevo = QLabel(self.estado.siguiente_nombre())
        nuevo.setMinimumWidth(180)
        nuevo.setStyleSheet("border: 1px dashed #7f8c8d; padding: 5px; color: #95a5a6;")
        mas = boton("＋", "exito", "Crear otro proyecto", 36)
        mas.clicked.connect(self.nuevo)
        f.addWidget(nuevo)
        f.addWidget(mas)
        f.addStretch()
        self.proy_lay.addLayout(f)
        # ramal (solo con 2 o más proyectos)
        vis = self.estado.ramal_visible()
        self.sep.setVisible(vis)
        self.lbl_ramal.setVisible(vis)
        if vis:
            r = self.estado.ramal
            f = QHBoxLayout()
            b = boton(r["abrev"], tip="Editar ramal")
            b.setStyleSheet(estilo_color(r["color"]))
            b.setMinimumWidth(180)
            b.clicked.connect(partial(self.editar, r))
            pin = boton("🖌", tip="Color de la pestaña", ancho=36)
            pin.clicked.connect(partial(self.color, r))
            for w in (b, pin):
                f.addWidget(w)
            f.addStretch()
            self.ramal_lay.addLayout(f)

    def editar(self, base):
        DlgEdicion(self, self.estado, "config", base).exec()
        self.construir()

    def color(self, base):
        c = QColorDialog.getColor(QColor(base["color"]), self, "Color de la pestaña")
        if c.isValid():
            base["color"] = c.name()
            self.estado.cambios = True
            self.construir()

    def nuevo(self):
        self.estado.agregar_proyecto()
        self.construir()
        self.scroll.verticalScrollBar().setValue(self.scroll.verticalScrollBar().maximum())

    def eliminar(self, base):
        unico = len(self.estado.proyectos) == 1
        texto = "Se eliminará todo lo relacionado al proyecto."
        if unico:
            texto += "\n\nEs el único proyecto: solo se reiniciará (se quitan las hojas extra)."
        elif len(self.estado.proyectos) == 2:
            texto += "\n\nQuedará un solo proyecto, por lo que también se eliminará el Ramal."
            # Las filas capturadas en el Ramal no viven en ningún proyecto: hay
            # que avisar cuántas se pierden antes de borrarlo.
            propias = sum(1 for hoja in EDITABLES
                          for lst in ((self.estado.ramal.get("items") or {})
                                      .get(hoja) or {}).values()
                          for it in lst if D(it.get("cantidad", "0")) > 0)
            if propias:
                texto += (f"\nSe perderán {propias} fila(s) agregadas solo en el Ramal.")
        else:
            # El Ramal hereda ANE - VAR de los proyectos: si este es el único
            # que la tiene, al eliminarlo la sincronización del Ramal descarta
            # los items de esa hoja y con ellos las filas capturadas solo ahí.
            # Se comprueba ANTES de eliminar, con el proyecto aún en la lista.
            pierde_var = (base["hojas"].get("ANE_VAR")
                          and not any(p is not base and p["hojas"].get("ANE_VAR")
                                      for p in self.estado.proyectos)
                          and self.estado.ramal["hojas"].get("ANE_VAR")
                          and any(self.estado._cant_propia("ANE_VAR").values()))
            if pierde_var:
                texto += ("\n\nEs el único proyecto con ANE - VAR: el Ramal la perderá y con "
                          "ella las filas capturadas solo ahí.")
        if confirmar(self, "Eliminar proyecto", texto + "\n\n¿Desea continuar?"):
            self.estado.eliminar_proyecto(base)
            self.construir()


# ═════════════════════════════════════════════════════════════════════
#  EDICIÓN
# ═════════════════════════════════════════════════════════════════════
class DlgEdicion(VentanaBase):
    def __init__(self, parent, estado, origen, foco=None):
        super().__init__(parent, "Edición", (760, 460), (700, 340))
        self.estado, self.origen, self.foco = estado, origen, foco
        cab = QHBoxLayout()
        cab.addWidget(etiqueta("Edición", "titulo"))
        ojo = boton("👁", tip="Ir a Visualización", ancho=40)
        ojo.clicked.connect(self.ir_visualizacion)
        cab.addWidget(ojo)
        cab.addStretch()
        self.cuerpo.addLayout(cab)
        self.lista = QVBoxLayout()
        self.cuerpo.addLayout(self.lista)
        self.cuerpo.addStretch()
        self.construir()

    def ir_visualizacion(self):
        if self.origen == "vis":
            self.accept()
        else:
            DlgVisualizacion(self, self.estado, "edicion").exec()
            self.construir()

    def construir(self):
        limpiar(self.lista)
        self.widgets = {}
        self.lista.addWidget(etiqueta("Proyectos", "seccion"))
        for b in self.estado.proyectos:
            self.lista.addWidget(self.fila(b, False))
        if self.estado.ramal_visible():
            self.lista.addWidget(etiqueta("Ramal", "seccion"))
            self.lista.addWidget(self.fila(self.estado.ramal, True))
        if self.foco is not None and id(self.foco) in self.widgets:
            w = self.widgets[id(self.foco)]
            QTimer.singleShot(0, lambda: self.scroll.ensureWidgetVisible(w))

    def fila(self, base, es_ramal):
        g = grupo()
        self.widgets[id(base)] = g
        self.widgets[base["abrev"]] = g
        lay = QVBoxLayout(g)
        top = QHBoxLayout()
        le = QLineEdit(base["abrev"])
        le.setMaximumWidth(160)
        le.setStyleSheet(f"QLineEdit {{ border: 2px solid {base['color']}; font-weight: bold; }}")
        le.editingFinished.connect(partial(self.cambiar_abrev, base, le))
        top.addWidget(le)
        for h in HOJAS:
            b = boton(ETIQUETA[h], "radio")
            b.setCheckable(True)
            b.setChecked(base["hojas"][h])
            b.clicked.connect(partial(self.toggle, base, h, b, es_ramal))
            if es_ramal:
                if h in EDITABLES:
                    # ANE y ANE-VAR son herencia de los proyectos: el Ramal no
                    # decide sobre ellas.
                    b.setEnabled(False)
                    b.setToolTip("Se activa automáticamente desde los proyectos.")
                else:
                    b.setToolTip("Hoja independiente del Ramal; requiere ANE - VAR.")
            top.addWidget(b)
        top.addStretch()
        lay.addLayout(top)
        nom = QLineEdit(base["nombre"])
        nom.setPlaceholderText("Nombre del Ramal" if es_ramal else "Nombre del Proyecto")
        nom.textChanged.connect(partial(self.cambiar_nombre, base))
        lay.addWidget(nom)
        return g

    def cambiar_nombre(self, base, texto):
        base["nombre"] = texto
        self.estado.cambios = True

    def cambiar_abrev(self, base, le):
        if le.text() == base["abrev"]:
            return
        nuevo, err = normalizar_abrev(le.text(), self.estado.otros_nombres(base))
        if err:
            msg(self, "Abreviación no válida", err, QMessageBox.Warning)
            le.setText(base["abrev"])
            return
        self.estado.renombrar(base, nuevo)
        le.setText(nuevo)

    def toggle(self, base, hoja, btn, es_ramal):
        e = self.estado
        marcado = btn.isChecked()
        if marcado:                                     # ── activar
            if hoja in DEPENDEN_DE_VAR and not base["hojas"]["ANE_VAR"]:
                if es_ramal:
                    # En el Ramal ANE-VAR no se fuerza: depende de los proyectos.
                    btn.setChecked(False)
                    msg(self, "Hoja requerida",
                        f"{ETIQUETA[hoja]} requiere ANE - VAR en el Ramal, y se activa "
                        "solo cuando algún proyecto la tenga.")
                    return
                e.activar_hoja(base, "ANE_VAR")
                msg(self, "Hoja requerida",
                    f"{ETIQUETA[hoja]} requiere ANE - VAR; se activó automáticamente.")
            e.activar_hoja(base, hoja)
        else:                                           # ── desactivar
            if hoja == "ANE":
                btn.setChecked(True)
                msg(self, "No permitido", "La hoja ANE no se puede desmarcar.")
                return
            if hoja == "ANE_VAR":
                deps = [ETIQUETA[x] for x in DEPENDEN_DE_VAR if base["hojas"][x]]
                # El Ramal hereda ANE - VAR de los proyectos: si con esto se
                # queda sin ella, su sincronización descarta los items de esa
                # hoja y con ellos las filas capturadas solo en el Ramal. Hay
                # que avisarlo antes de aceptar.
                pierde_ramal = (e.ramal_visible()
                                and e.ramal["hojas"].get("ANE_VAR")
                                and (es_ramal or not any(
                                    p is not base and p["hojas"].get("ANE_VAR")
                                    for p in e.proyectos))
                                and any(e._cant_propia("ANE_VAR").values()))
                aviso = ("\nSe perderán las filas capturadas solo en ANE - VAR del Ramal."
                         if pierde_ramal else "")
                if deps:
                    texto = ("Si elimina esta hoja eliminará también las siguientes:\n\n  • "
                             + "\n  • ".join(deps) + aviso + "\n\n¿Desea continuar?")
                else:
                    texto = "¿Desea eliminar la hoja ANE - VAR?" + aviso
                ok = confirmar(self, "Eliminar hoja", texto)
                if not ok:
                    btn.setChecked(True)
                    return
                for x in DEPENDEN_DE_VAR:
                    if base["hojas"][x]:
                        e.quitar_hoja(base, x)
            else:
                if not confirmar(self, "Eliminar hoja",
                                 f"¿Desea eliminar la hoja {ETIQUETA[hoja]}?"
                                 + ("\nSe perderán los componentes capturados en ella."
                                    if hoja in EDITABLES else "")):
                    btn.setChecked(True)
                    return
            e.quitar_hoja(base, hoja)
        self.construir()


# ═════════════════════════════════════════════════════════════════════
#  VISUALIZACIÓN
# ═════════════════════════════════════════════════════════════════════
class DlgVisualizacion(VentanaBase):
    def __init__(self, parent, estado, origen):
        super().__init__(parent, "Visualización", (700, 340), (640, 260))
        self.estado, self.origen = estado, origen
        cab = QHBoxLayout()
        cab.addWidget(etiqueta("Visualización", "titulo"))
        self.btn_tema = boton(icono_tema(), tip="Día / Noche (Alt+B)", ancho=40)
        self.btn_tema.clicked.connect(self.cambiar_tema)
        cab.addWidget(self.btn_tema)
        cab.addStretch()
        self.cuerpo.addLayout(cab)
        self.lista = QVBoxLayout()
        self.cuerpo.addLayout(self.lista)
        self.cuerpo.addStretch()
        self.construir()

    def cambiar_tema(self):
        alternar_tema()
        self.btn_tema.setText(icono_tema())

    def construir(self):
        limpiar(self.lista)
        self.lista.addWidget(etiqueta("Proyectos", "seccion"))
        for b in self.estado.proyectos:
            self.lista.addWidget(self.fila(b))
        if self.estado.ramal_visible():
            self.lista.addWidget(etiqueta("Ramal", "seccion"))
            self.lista.addWidget(self.fila(self.estado.ramal))

    def fila(self, base):
        g = grupo()
        lay = QGridLayout(g)
        ab = boton(base["abrev"], tip="Ir a Edición")
        ab.setStyleSheet(estilo_color(base["color"]))
        ab.setMinimumWidth(150)
        ab.clicked.connect(partial(self.ir_edicion, base))
        lay.addWidget(ab, 0, 0)
        for col, h in enumerate(HOJAS, 1):          # columnas fijas: lo que no existe queda hueco
            lay.setColumnMinimumWidth(col, 80)
            if not base["hojas"][h]:
                continue
            b = boton(ETIQUETA[h], "radio")
            b.setCheckable(True)
            b.setChecked(h in base["visibles"])
            b.clicked.connect(partial(self.toggle, base, h, b))
            lay.addWidget(b, 0, col)
        lay.setColumnStretch(len(HOJAS) + 1, 1)
        return g

    def toggle(self, base, hoja, btn):
        if btn.isChecked():
            if hoja not in base["visibles"]:
                base["visibles"].append(hoja)
        else:
            if len(self.estado.tabs()) <= 1:
                btn.setChecked(True)
                msg(self, "No permitido", "Debe quedar al menos una hoja visible.")
                return
            base["visibles"].remove(hoja)
        self.estado.cambios = True

    def ir_edicion(self, base):
        if self.origen == "edicion":
            self.accept()
        else:
            DlgEdicion(self, self.estado, "vis", base).exec()
            self.construir()


# ═════════════════════════════════════════════════════════════════════
#  COMPONENTES  (base de datos principal)
# ═════════════════════════════════════════════════════════════════════
class DlgComponentes(VentanaBase):
    COLS = ["CODIGO\nCOMPONENTE", "Descripción del componente", "Unidad", "Precio\nUnitario",
            "MT", "MT\n+ BT", "BT", "G", "7.6\nkV", "19.9\nkV"]

    def __init__(self, parent, estado):
        super().__init__(parent, "Componentes", (980, 640), (860, 500))
        self.estado = estado
        self.datos = copy.deepcopy(estado.bd)
        self.renombres = []
        self.cambios = False
        self._interno = False

        cab = QHBoxLayout()
        cab.addWidget(etiqueta("Componentes", "titulo"))
        self.btn_guardar = boton("💾", tip="Guardar (Ctrl+S)", ancho=40)
        self.btn_guardar.clicked.connect(self.guardar)
        self.btn_guardar.setEnabled(False)
        cab.addWidget(self.btn_guardar)
        cab.addStretch()
        self.cuerpo.addLayout(cab)

        g = grupo()
        lay = QVBoxLayout(g)
        f1 = QHBoxLayout()
        self.cod = CampoBusqueda("Código", self.proveedor_codigo, self.max_filas)
        self.cod.setMaximumWidth(140)
        self.cod.setToolTip("Para componentes nuevos use solo números: solo los "
                            f"códigos numéricos mayores a {UMBRAL_CODIGO_ATE} se "
                            "clasifican como ATE.")
        self.desc = CampoBusqueda("Componente", self.proveedor_descripcion, self.max_filas)
        self.cod.elegido.connect(self.elegir)
        self.desc.elegido.connect(self.elegir)
        self.cod.textEdited.connect(self.cod_editado)
        self.desc.textEdited.connect(self.desc_editado)
        self.cod.limpiar.connect(self.limpiar_campos)
        self.desc.limpiar.connect(self.limpiar_campos)
        f1.addWidget(self.cod)
        f1.addWidget(self.desc, 1)
        f2 = QHBoxLayout()
        self.uni = QLineEdit()
        self.uni.setPlaceholderText("Unidad")
        self.uni.setMaximumWidth(140)
        self.precio = QLineEdit()
        self.precio.setPlaceholderText("Precio Unitario (Q)")
        self.precio.setMaximumWidth(180)
        self.precio.setValidator(QRegularExpressionValidator(QRegularExpression(r"^\d{0,9}(\.\d{0,2})?$")))
        self.btn_add = boton("Agregar", "primario")
        self.btn_add.clicked.connect(self.agregar)
        self.lbl_modo = QLabel("")
        for w in (self.uni, self.precio, self.btn_add, self.lbl_modo):
            f2.addWidget(w)
        f2.addStretch()
        for w in (self.uni, self.precio):
            w.textChanged.connect(self.validar_agregar)
        lay.addLayout(f1)
        lay.addLayout(f2)
        self.cuerpo.addWidget(g)

        self.tabla = TablaBase(0, len(self.COLS), arrastrable=True)
        self.tabla.setHorizontalHeaderLabels(self.COLS)
        self.tabla.setMinimumHeight(300)
        hh = self.tabla.horizontalHeader()
        hh.setFixedHeight(40)
        hh.setSectionResizeMode(QHeaderView.Fixed)
        hh.setSectionResizeMode(1, QHeaderView.Stretch)
        for i, w in enumerate((100, 0, 90, 100, 48, 48, 48, 48, 52, 52)):
            if w:
                self.tabla.setColumnWidth(i, w)
        self.tabla.verticalHeader().setDefaultSectionSize(26)
        self.tabla.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tabla.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tabla.itemChanged.connect(self.celda_cambiada)
        self.tabla.itemSelectionChanged.connect(self.fila_seleccionada)
        self.tabla.suprimir.connect(self.eliminar)
        self.tabla.filaMovida.connect(self.mover)
        self.cuerpo.addWidget(self.tabla, 1)
        QShortcut(QKeySequence("Ctrl+S"), self).activated.connect(self.guardar)
        self.refrescar()
        self.evaluar_modo()

    # ── búsqueda ────────────────────────────────────────────────────
    def max_filas(self):
        return max(8, (self.height() - 400) // 22 + 8)

    def proveedor_codigo(self, texto, todos):
        return self._proveedor(texto, todos, "codigo")

    def proveedor_descripcion(self, texto, todos):
        return self._proveedor(texto, todos, "descripcion")

    def _proveedor(self, texto, todos, campo):
        """Busca en el campo activo y pone primero las coincidencias iniciales."""
        t = norm(texto)
        filas = []
        for i, c in enumerate(self.datos):
            nivel = 0 if not t else _nivel_coincidencia(c[campo], t, campo)
            if nivel > 3:
                continue
            filas.append((nivel, i, c["codigo"], c["descripcion"]))
        filas.sort(key=lambda x: (x[0], x[1]))
        return [(cod, desc, True) for _, _, cod, desc in filas[:300]]

    def elegir(self, codigo):
        i = next((n for n, c in enumerate(self.datos) if c["codigo"] == codigo), None)
        if i is None:
            return
        self.tabla.selectRow(i)
        self.tabla.scrollToItem(self.tabla.item(i, 0))

    def limpiar_campos(self):
        """Ctrl+Espacio: borra los campos; el campo activo abre su propia lista."""
        self.cod.popup.hide()
        self.desc.popup.hide()
        self.cod.clear()
        self.desc.clear()
        self.uni.clear()
        self.precio.clear()
        self.evaluar_modo()

    # ── modos: existente (bloquea) / nuevo (desbloquea) ─────────────
    def _indice(self):
        c, t = self.cod.text().strip(), norm(self.desc.text())
        return next((i for i, x in enumerate(self.datos)
                     if (c and x["codigo"] == c) or (t and norm(x["descripcion"]) == t)), None)

    def cod_editado(self, texto):
        i = next((n for n, x in enumerate(self.datos) if x["codigo"] == texto.strip()), None)
        if i is not None:
            self.desc.setText(self.datos[i]["descripcion"])
        self.evaluar_modo()

    def desc_editado(self, texto):
        i = next((n for n, x in enumerate(self.datos) if norm(x["descripcion"]) == norm(texto)), None)
        if i is not None:
            self.cod.setText(self.datos[i]["codigo"])
        self.evaluar_modo()

    def evaluar_modo(self):
        i = self._indice()
        existe = i is not None
        for w in (self.uni, self.precio):
            w.setEnabled(not existe)
        if existe:
            c = self.datos[i]
            self.uni.setText(c["unidad"])
            self.precio.setText(str(D(c["precio"])))
            self.lbl_modo.setText("Existente: edítelo directamente en la tabla.")
            self._interno = True
            self.tabla.selectRow(i)
            self._interno = False
        else:
            self.lbl_modo.setText(
                f"Nuevo componente (numérico > {UMBRAL_CODIGO_ATE} es ATE)"
                if (self.cod.text() or self.desc.text()) else "")
        self.validar_agregar()

    def validar_agregar(self):
        if self._indice() is not None:
            self.btn_add.setEnabled(False)
            return
        c, t = self.cod.text().strip(), self.desc.text().strip()
        # Los componentes nuevos deben ser numéricos: solo así el Tipo de
        # Documento puede marcarlos como ATE (si además superan el umbral).
        ok = (bool(re.fullmatch(r"\d+", c)) and t and self.uni.text().strip()
              and self.precio.text().strip() and self.precio.text() != ".")
        self.btn_add.setEnabled(bool(ok))

    def agregar(self):
        codigo = self.cod.text().strip()
        if not re.fullmatch(r"\d+", codigo):
            msg(self, "Código no válido",
                "Los componentes nuevos deben tener un código numérico.",
                QMessageBox.Warning)
            return
        descripcion = self.desc.text().strip()
        # Las seis marcas nacen en 0: el usuario decide las líneas y guarda.
        nuevo = componente_nuevo(codigo, descripcion, self.uni.text().strip(),
                                 self.precio.text())
        self.datos.append(nuevo)
        self.marcar_cambios()
        self.cod.clear()
        self.desc.clear()
        self.uni.clear()
        self.precio.clear()
        self.refrescar()
        self.evaluar_modo()
        self.tabla.selectRow(len(self.datos) - 1)
        self.tabla.scrollToBottom()
        # El componente nuevo queda sin líneas: hay que marcarlas y guardar,
        # o el código no aparecerá en ninguna tabla del anexo.
        self.lbl_modo.setText("Marque las líneas del componente nuevo y guarde (💾).")

    def fila_seleccionada(self):
        if self._interno:
            return
        r = self.tabla.currentRow()
        if 0 <= r < len(self.datos):
            c = self.datos[r]
            self.cod.setText(c["codigo"])
            self.desc.setText(c["descripcion"])
            self.evaluar_modo()

    # ── tabla ───────────────────────────────────────────────────────
    def refrescar(self):
        self._interno = True
        self.tabla.blockSignals(True)
        self.tabla.setRowCount(len(self.datos))
        for r, c in enumerate(self.datos):
            for col, v in enumerate((c["codigo"], c["descripcion"], c["unidad"], fmt_q(c["precio"]))):
                it = QTableWidgetItem(v)
                if col == 3:
                    it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.tabla.setItem(r, col, it)
            for col, k in enumerate(("MT", "MTBT", "BT", "G", "v76", "v199"), 4):
                it = QTableWidgetItem()
                it.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                it.setCheckState(Qt.Checked if c[k] else Qt.Unchecked)
                it.setTextAlignment(Qt.AlignCenter)
                self.tabla.setItem(r, col, it)
            self._bloqueo_voltaje(r)
        self.tabla.blockSignals(False)
        self._interno = False

    def _bloqueo_voltaje(self, r):
        """Si un voltaje está marcado, el otro se deshabilita."""
        c = self.datos[r]
        prev = self.tabla.blockSignals(True)      # setFlags emite itemChanged → evitar recursión
        for col, k, otro in ((8, "v76", "v199"), (9, "v199", "v76")):
            it = self.tabla.item(r, col)
            fl = Qt.ItemIsUserCheckable | Qt.ItemIsSelectable
            if not c[otro]:
                fl |= Qt.ItemIsEnabled
            it.setFlags(fl)
        self.tabla.blockSignals(prev)

    def celda_cambiada(self, it):
        if self._interno:
            return
        r, col = it.row(), it.column()
        c = self.datos[r]
        texto = it.text().strip()
        revertir = None
        if col == 0:
            if (not texto or not re.fullmatch(r"[A-Za-z0-9_-]+", texto)
                    or any(x is not c and x["codigo"] == texto for x in self.datos)):
                revertir = "El código no puede estar vacío, repetido ni contener espacios."
            else:
                self.renombres.append((c["codigo"], texto))
                c["codigo"] = texto
        elif col == 1:
            if not texto or any(x is not c and norm(x["descripcion"]) == norm(texto) for x in self.datos):
                revertir = "La descripción no puede estar vacía ni repetida."
            else:
                c["descripcion"] = texto
        elif col == 2:
            if not texto:
                revertir = "La unidad no puede estar vacía."
            else:
                c["unidad"] = texto
        elif col == 3:
            v = D(texto)
            if v < 0 or not re.fullmatch(r"[Q\s]*[\d,]*\.?\d{0,2}", texto) or not re.search(r"\d", texto):
                revertir = "Precio no válido (máximo 2 decimales)."
            else:
                c["precio"] = str(v.quantize(Decimal("0.01")))
        else:
            k = ("MT", "MTBT", "BT", "G", "v76", "v199")[col - 4]
            marcado = it.checkState() == Qt.Checked
            otro = {"v76": "v199", "v199": "v76"}.get(k)
            if marcado and otro and c[otro]:
                pass                       # exclusión: no se permiten ambos voltajes (se revierte)
            else:
                c[k] = 1 if marcado else 0
                if otro:
                    self._bloqueo_voltaje(r)
        if revertir:
            msg(self, "Valor no válido", revertir, QMessageBox.Warning)
        else:
            self.marcar_cambios()
        # Un código pudo cambiar de nombre: el índice de la BD hay que rehacerlo.
        self.estado.reindexar_bd()
        self.refrescar()

    def eliminar(self):
        r = self.tabla.currentRow()
        if r < 0:
            return
        c = self.datos[r]
        usos = self.estado.codigo_en_uso(c["codigo"])
        if usos:
            msg(self, "Componente en uso",
                "No se puede eliminar: está en las hojas\n\n  • " + "\n  • ".join(usos),
                QMessageBox.Warning)
            return
        if confirmar(self, "Eliminar componente", f"¿Eliminar «{c['descripcion']}» de la base de datos?"):
            self.datos.pop(r)
            self.marcar_cambios()
            self.refrescar()

    def mover(self, src, dst):
        self.datos.insert(dst, self.datos.pop(src))
        self.marcar_cambios()
        self.refrescar()
        self.tabla.selectRow(dst)

    # ── guardado ────────────────────────────────────────────────────
    def marcar_cambios(self):
        self.cambios = True
        self.btn_guardar.setEnabled(True)

    def guardar(self):
        """Devuelve False si el usuario cancela."""
        if not self.cambios:
            return True
        m = QMessageBox(self)
        m.setIcon(QMessageBox.Question)
        m.setWindowTitle("Guardar componentes")
        m.setText("¿Dónde desea guardar la configuración?\n\n"
                  "Local: solo para este proyecto.\n"
                  "Global: modifica la plantilla de Excel (aplica a proyectos nuevos).")
        loc = m.addButton("Local", QMessageBox.AcceptRole)
        glo = m.addButton("Global", QMessageBox.AcceptRole)
        can = m.addButton("Cancelar", QMessageBox.RejectRole)
        m.setEscapeButton(can)
        m.exec()
        b = m.clickedButton()
        if b is can:
            return False
        try:
            if b is glo:
                bd_guardar_plantilla(self.datos)
        except PermissionError as ex:
            msg(self, "Error",
                str(ex) or "No se pudo escribir la plantilla (¿abierta en Excel?).",
                QMessageBox.Critical)
            return False
        except Exception as ex:
            msg(self, "Error al guardar",
                f"No se pudo actualizar la tabla {TABLA_BD}:\n\n{ex}", QMessageBox.Critical)
            return False
        self.estado.bd = copy.deepcopy(self.datos)
        for viejo, nuevo in self.renombres:
            self.estado.renombrar_codigo(viejo, nuevo)
        self.renombres.clear()
        self.estado.bd_local = (b is loc)
        self.estado.cambios = True
        # Cambiaron códigos/descripciones/marcas: el índice de la BD se rehace.
        self.estado.reindexar_bd()
        self.cambios = False
        self.btn_guardar.setEnabled(False)
        return True

    def reject(self):
        if self.cambios:
            op = tres_opciones(self, "Cambios sin guardar",
                               "Hay cambios en los componentes. ¿Desea guardarlos?",
                               "Guardar", "Descartar", "Cancelar")
            if op == 2 or (op == 0 and not self.guardar()):
                return
        super().reject()


# ═════════════════════════════════════════════════════════════════════
#  MENÚ PROYECTO y REPORTE DE ACTUALIZACIÓN
# ═════════════════════════════════════════════════════════════════════
COMANDOS_CONTEXTO = [
    ("Tabla", "Del", "Eliminar las filas seleccionadas"),
    ("Tabla", "Alt+↑ / Alt+↓", "Mover la fila seleccionada arriba o abajo"),
    ("Materiales", "Enter", "Agregar desde el campo Cantidad"),
    ("Buscador", "Ctrl+Espacio", "Limpiar los campos y mostrar todos"),
    ("Buscador", "↑ / ↓", "Navegar por los resultados"),
    ("Buscador", "Enter", "Elegir el resultado resaltado"),
    ("Buscador", "Esc", "Cerrar la lista sin elegir"),
    ("Ventanas", "Esc", "Cerrar la ventana o el popup activo"),
]


class DlgComandos(QDialog):
    """Tabla de comandos de la ventana principal, con filtro de texto."""

    def __init__(self, parent, filas):
        super().__init__(parent)
        self.setWindowTitle("Comandos de Anexos")
        self.setWindowIcon(icono_app())
        self.resize(760, 620)
        self.setMinimumSize(600, 400)
        self.filas = list(filas)

        lay = QVBoxLayout(self)
        cab = QHBoxLayout()
        cab.addWidget(QLabel("Comandos"))
        self.filtro = QLineEdit()
        self.filtro.setPlaceholderText("Filtrar comando o descripción…")
        self.filtro.setClearButtonEnabled(True)
        cab.addWidget(self.filtro, 1)
        lay.addLayout(cab)

        self.tabla = QTableWidget(0, 3)
        self.tabla.setHorizontalHeaderLabels(["Grupo", "Comando", "Descripción"])
        self.tabla.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tabla.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tabla.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tabla.setAlternatingRowColors(True)
        self.tabla.setShowGrid(False)
        self.tabla.verticalHeader().setVisible(False)
        self.tabla.verticalHeader().setDefaultSectionSize(24)
        hh = self.tabla.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(2, QHeaderView.Stretch)
        lay.addWidget(self.tabla, 1)

        pie = QHBoxLayout()
        self.lbl_conteo = QLabel("")
        self.filtro.textChanged.connect(self.refrescar)
        pie.addWidget(self.lbl_conteo)
        pie.addStretch()
        b = boton("Cerrar", "primario")
        b.clicked.connect(self.accept)
        pie.addWidget(b)
        lay.addLayout(pie)

        self.refrescar()
        self.filtro.setFocus()

    def refrescar(self):
        texto = norm(self.filtro.text())
        visibles = [f for f in self.filas
                    if not texto or texto in norm(" ".join(f))]
        self.tabla.setRowCount(len(visibles))
        for r, (grupo, comando, descripcion) in enumerate(visibles):
            for c, valor in enumerate((grupo, comando, descripcion)):
                it = QTableWidgetItem(valor)
                if c == 1:
                    fuente = it.font()
                    fuente.setFamily("Consolas")
                    fuente.setBold(True)
                    it.setFont(fuente)
                self.tabla.setItem(r, c, it)
        self.lbl_conteo.setText(f"{len(visibles)} comando(s)")


class DlgProyecto(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("Proyecto")
        self.setWindowIcon(icono_app())
        self.setFixedSize(360, 200)
        self.accion = None
        lay = QVBoxLayout(self)
        lay.addWidget(etiqueta("Proyecto", "seccion"))
        for txt, acc, rol in (("📄  Nuevo proyecto", "nuevo", "exito"),
                              ("📂  Abrir proyecto…", "abrir", "primario")):
            b = boton(txt, rol)
            b.clicked.connect(partial(self.elegir, acc))
            lay.addWidget(b)
        lay.addStretch()

    def elegir(self, acc):
        self.accion = acc
        self.accept()


class DlgReporte(VentanaBase):
    def __init__(self, parent, res, estado):
        # Actualizar es una verificación de solo lectura: el título y los
        # rótulos lo dicen para que nadie crea que el archivo se reescribió.
        verif = bool(res.get("verificacion"))
        super().__init__(parent,
                         "Resultado de la verificación" if verif
                         else "Resultado de la actualización",
                         (620, 420), (520, 300))
        self.estado = estado
        self.cuerpo.addWidget(etiqueta("Verificación" if verif else "Actualización",
                                       "titulo"))
        if verif:
            self.cuerpo.addWidget(QLabel(
                "✔ Verificación de solo lectura: el archivo no se modificó."))
        duplicados = res.get("duplicados") or []
        faltantes = res.get("faltantes") or []
        if (not res["totales"] and not res["externos"]
                and not duplicados and not faltantes):
            self.cuerpo.addWidget(QLabel("✔ Sin cambios detectados en los totales ni componentes externos."))
        if res["totales"]:
            self.cuerpo.addWidget(etiqueta("Totales que cambiaron", "seccion"))
            t = QTableWidget(len(res["totales"]), 4)
            t.setHorizontalHeaderLabels(
                ["Hoja", "Total guardado", "Total en el archivo", "Diferencia"]
                if verif else
                ["Hoja", "Total anterior", "Total actual", "Diferencia"])
            for i, (h, a, n) in enumerate(res["totales"]):
                for j, v in enumerate((h, fmt_q(a), fmt_q(n), fmt_q(D(n) - D(a)))):
                    t.setItem(i, j, QTableWidgetItem(v))
            t.horizontalHeader().setStretchLastSection(True)
            t.setEditTriggers(QAbstractItemView.NoEditTriggers)
            t.setMinimumHeight(40 + 26 * len(res["totales"]))
            self.cuerpo.addWidget(t)
        if res["externos"]:
            self.cuerpo.addWidget(etiqueta("Componentes que no están en la base de datos", "seccion"))
            for (hoja, fila, cod, desc, uni, precio) in res["externos"]:
                f = QHBoxLayout()
                f.addWidget(QLabel(f"{hoja} · fila {fila} · {cod} · {desc} · {uni} · {fmt_q(precio)}"), 1)
                b = boton("Agregar a la BD", "primario")
                b.clicked.connect(partial(self.agregar_bd, b, cod, desc, uni, precio))
                f.addWidget(b)
                self.cuerpo.addLayout(f)
        if duplicados:
            self.cuerpo.addWidget(etiqueta(
                "Filas duplicadas en el Excel" if verif
                else "Filas duplicadas (se conservó la primera)", "seccion"))
            nombres = {k: txt for k, _, txt in LINEAS}
            for (hoja, linea, cod, cant) in duplicados:
                self.cuerpo.addWidget(QLabel(
                    f"  •  {hoja} · {nombres.get(linea, linea)} · {cod}"
                    f" · cantidad {'en el archivo' if verif else 'descartada'} {cant}"))
        if faltantes:
            self.cuerpo.addWidget(etiqueta(
                "Hojas activas que no están en el Excel" if verif
                else "Hojas activas que no están en el Excel (se vaciaron)", "seccion"))
            self.cuerpo.addWidget(QLabel("  •  " + "\n  •  ".join(faltantes)))
        if res["actualizadas"]:
            self.cuerpo.addWidget(etiqueta("Hojas actualizadas", "seccion"))
            self.cuerpo.addWidget(QLabel("  •  " + "\n  •  ".join(res["actualizadas"])))
        if res.get("bd"):
            self.cuerpo.addWidget(QLabel(f"Base de códigos leída del archivo: {res['bd']} códigos."))
        self.cuerpo.addStretch()

    def agregar_bd(self, btn, cod, desc, uni, precio):
        if not self.estado.comp(cod):
            # Los componentes externos también nacen sin línea marcada.
            self.estado.bd.append(componente_nuevo(cod, desc, uni, precio))
            self.estado.cambios = True
        btn.setText("✔ Agregado (marque líneas)")
        btn.setEnabled(False)


# ═════════════════════════════════════════════════════════════════════
#  VENTANA PRINCIPAL
# ═════════════════════════════════════════════════════════════════════
class Principal(QWidget):
    ANCHOS = (95, 0, 80, 105, 95, 140)       # 0 = stretch (Descripción)
    ENC = ["CODIGO\nCOMPONENTE", "Descripción del componente", "Unidad", "Precio\nUnitario",
           "Cantidad\na ejecutar", "Precio\nTotal"]

    def __init__(self):
        super().__init__()
        self.estado = Estado()
        self.tab = None                       # (base, hoja) actual
        self.comp_sel = None                  # componente seleccionado en los campos
        self.ultima_linea = None              # línea del último material agregado
        self._panel_visible = True            # preferencia del usuario en hojas editables
        self._panel_ramal = False             # el panel del Ramal arranca plegado
        self.setWindowTitle("Anexos - WuMDrex")
        self.setWindowIcon(icono_app())
        self.toast = Toast(self)
        self._atajos = []
        # Historial de Ctrl+Z / Ctrl+Y (ver sección DESHACER / REHACER).
        self._pila_deshacer = []
        self._pila_rehacer = []
        self._en_historial = False
        # El título se refresca solo cuando cambia la bandera de cambios.
        self.estado._al_avisar_cambios = self.actualizar_titulo
        self.resize(820, 780)
        self.setMinimumSize(820, 780)
        self.construir()
        self.refrescar_todo()
        self._instalar_atajos()

    # ── autoajuste de las tablas ────────────────────────────────────
    def resizeEvent(self, e):
        super().resizeEvent(e)
        # El ancho de la columna Descripción cambia con la ventana; recalcular
        # el alto de las filas permite leer el componente completo.
        if hasattr(self, "secciones"):
            QTimer.singleShot(0, self._ajustar_alturas_tablas)

    def _ajustar_alto_tabla(self, t):
        """Alto de fila según el ancho real de cada columna y el texto."""
        if t.rowCount() == 0:
            return
        fm = t.fontMetrics()
        for r in range(t.rowCount()):
            alto = 24
            for col in range(t.columnCount()):
                it = t.item(r, col)
                if it is None or not it.text():
                    continue
                ancho = max(40, t.columnWidth(col) - 8)
                rect = fm.boundingRect(QRect(0, 0, ancho, 10000),
                                       Qt.TextWordWrap, it.text())
                alto = max(alto, rect.height() + 8)
            t.setRowHeight(r, alto)
        total = (t.horizontalHeader().height() +
                 sum(t.rowHeight(r) for r in range(t.rowCount())) + 4)
        if t.columnCount() > 6:
            total += t.horizontalScrollBar().sizeHint().height()
        t.setFixedHeight(max(total, t.horizontalHeader().height() + 4))

    def _ajustar_alturas_tablas(self):
        if self._cargando or not hasattr(self, "secciones"):
            return
        for _, t in self.secciones.values():
            if t.isVisible():
                self._ajustar_alto_tabla(t)

    # ── construcción ────────────────────────────────────────────────
    COLUMNAS_BASE = (
        ("No.\nCódigo", 90),
        ("Descripción del componente", 0),
        ("Unidad de\nmedida", 80),
        ("Cantidad", 80),
        ("Precio\nunitario (Q)", 100),
        ("Precio\ntotal (Q)", 105),
        ("Cantidad - V", 85),
        ("Precio\nunitario ( Q )", 100),
        ("Precio\ntotal ( Q )", 105),
        ("Tipo de\ndocumento", 95),
        ("Cantidad\nvariada", 85),
        ("Precio\nunitario ( Q)", 100),
    )
    COLUMNAS_VARIACION = COLUMNAS_BASE + (
        ("OC", 95), ("OTS", 95), ("ATE", 95),
    )

    def _columnas_hoja(self, hoja):
        if hoja in EDITABLES:
            return list(zip(self.ENC, self.ANCHOS))
        if hoja == "VARIACION":
            return list(self.COLUMNAS_VARIACION)
        if hoja in ("OC", "OTS", "ATE"):
            return list(self.COLUMNAS_BASE) + [(hoja, 100)]
        return list(self.COLUMNAS_VARIACION)

    def _configurar_tabla(self, t, hoja):
        """Adapta las columnas de la tabla a la hoja que se está viendo."""
        cols = self._columnas_hoja(hoja)
        t.blockSignals(True)
        t.setColumnCount(len(cols))
        t.setHorizontalHeaderLabels([c[0] for c in cols])
        hh = t.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.Fixed)
        if len(cols) > 1:
            hh.setSectionResizeMode(1, QHeaderView.Stretch)
        for i, (_, ancho) in enumerate(cols):
            if ancho:
                t.setColumnWidth(i, ancho)
        t.blockSignals(False)

    def construir(self):
        raiz = QVBoxLayout(self)
        raiz.setContentsMargins(8, 8, 8, 8)
        raiz.setSpacing(6)

        barra = QHBoxLayout()
        self.btn_cfg = boton("⚙", tip="Configuración", ancho=40)
        self.btn_proy = boton("Proyecto", tip="Nuevo / Abrir")
        self.btn_vis = boton("👁", tip="Visualización", ancho=40)
        self.btn_comp = boton("📚", tip="Componentes", ancho=40)
        self.btn_act = boton("🔄", tip="Actualizar", ancho=40)
        self.btn_gua = boton("💾", tip="Guardar (Ctrl+S)", ancho=40)
        self.btn_cfg.clicked.connect(self.abrir_config)
        self.btn_proy.clicked.connect(self.proyecto)
        self.btn_vis.clicked.connect(self.abrir_vis)
        self.btn_comp.clicked.connect(self.abrir_comp)
        self.btn_act.clicked.connect(self.actualizar)
        self.btn_gua.clicked.connect(self.guardar)
        for b in (self.btn_cfg, self.btn_proy, self.btn_vis, self.btn_comp, self.btn_act, self.btn_gua):
            barra.addWidget(b)
        barra.addStretch()
        barra.addWidget(QLabel("Voltaje:"))
        self.cmb_volt = QComboBox()
        for txt, _ in VOLTAJES:
            self.cmb_volt.addItem(txt)
        self.cmb_volt.currentIndexChanged.connect(self.voltaje_cambiado)
        aplicar_estilo_combo(self.cmb_volt)
        barra.addWidget(self.cmb_volt)
        raiz.addLayout(barra)

        # panel Materiales
        self.panel_g = grupo()
        pl = QVBoxLayout(self.panel_g)
        cab = QHBoxLayout()
        cab.addWidget(etiqueta("Materiales:", "seccion"))
        cab.addStretch()
        self.btn_colapsar = boton("▲", tip="Mostrar / ocultar el panel", ancho=30)
        self.btn_colapsar.clicked.connect(self.colapsar)
        cab.addWidget(self.btn_colapsar)
        pl.addLayout(cab)
        self.panel = QWidget()
        self.panel.setObjectName("panel_materiales")
        self.panel.setStyleSheet("QWidget#panel_materiales { background: transparent; }")
        gl = QVBoxLayout(self.panel)
        gl.setContentsMargins(0, 0, 0, 0)
        f1 = QHBoxLayout()
        self.cod = CampoBusqueda("Código", self.proveedor_codigo, self.max_filas)
        self.cod.setMaximumWidth(150)
        self.desc = CampoBusqueda("Componente", self.proveedor_descripcion, self.max_filas)
        self.cod.elegido.connect(self.elegir)
        self.desc.elegido.connect(self.elegir)
        self.cod.textEdited.connect(self.cod_editado)
        self.desc.textEdited.connect(self.desc_editado)
        # Ctrl+Espacio: limpia todos los campos; el campo activo abre su lista.
        self.cod.limpiar.connect(self.limpiar_campos_busqueda)
        self.desc.limpiar.connect(self.limpiar_campos_busqueda)
        f1.addWidget(self.cod)
        f1.addWidget(self.desc, 1)
        f2 = QHBoxLayout()
        self.cant = QLineEdit()
        self.cant.setPlaceholderText("Cantidad")
        self.cant.setMaximumWidth(150)
        self.cant.textChanged.connect(self.validar_agregar)
        self.lbl_unidad = QLabel("")
        self.lbl_unidad.setMinimumWidth(110)
        self.cmb_linea = ComboLinea()
        try:
            self.cmb_linea.setPlaceholderText("Línea (elija material)")
        except AttributeError:
            pass
        self.cmb_linea.setEnabled(False)
        self.cmb_linea.currentIndexChanged.connect(self.linea_cambiada)
        aplicar_estilo_combo(self.cmb_linea)
        self.btn_add = boton("Agregar", "primario")
        self.btn_add.clicked.connect(self.agregar)
        self.cant.returnPressed.connect(self.agregar)
        f2.addWidget(self.cant)
        f2.addWidget(self.lbl_unidad)
        f2.addWidget(QLabel("Línea:"))
        f2.addWidget(self.cmb_linea)
        f2.addStretch()
        f2.addWidget(self.btn_add)
        gl.addLayout(f1)
        gl.addLayout(f2)
        pl.addWidget(self.panel)
        raiz.addWidget(self.panel_g)
        self.set_validador_cantidad("")

        self.lbl_info = QLabel("")
        self.lbl_info.setWordWrap(True)
        self.lbl_info.setStyleSheet("color: #f39c12; font-weight: bold;")
        raiz.addWidget(self.lbl_info)

        # Aviso permanente de VARIACIÓN: se oculta cuando la variación cuadra.
        self.alerta_var = QLabel("")
        self.alerta_var.setWordWrap(True)
        self.alerta_var.setAlignment(Qt.AlignCenter)
        self.alerta_var.setStyleSheet(
            "QLabel { background-color: #e74c3c; color: white; font-weight: bold;"
            " border-radius: 8px; padding: 7px 14px; }")
        self.alerta_var.hide()
        raiz.addWidget(self.alerta_var)

        # tablas
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)
        cont = QWidget()
        self.tablas_lay = QVBoxLayout(cont)
        self.tablas_lay.setContentsMargins(0, 0, 0, 0)
        self.secciones = {}
        for k, titulo, _ in LINEAS:
            lbl = etiqueta(titulo, "banda")
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setWordWrap(True)
            t = TablaBase(0, 6, arrastrable=True, columnas_no_arrastre=(4, 5))
            self._configurar_tabla(t, "ANE")
            hh = t.horizontalHeader()
            hh.setFixedHeight(38)
            t.verticalHeader().setVisible(False)
            t.verticalHeader().setDefaultSectionSize(24)
            t.setWordWrap(True)
            t.setTextElideMode(Qt.ElideNone)
            t.setSelectionBehavior(QAbstractItemView.SelectRows)
            t.setSelectionMode(QAbstractItemView.ExtendedSelection)
            t.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            t.itemChanged.connect(partial(self.cantidad_editada, k))
            t.suprimir.connect(partial(self.eliminar_filas, k))
            t.filaMovida.connect(partial(self.mover_fila, k))
            t.itemSelectionChanged.connect(partial(resaltar_botones_fila, t))
            self.tablas_lay.addWidget(lbl)
            self.tablas_lay.addWidget(t)
            self.secciones[k] = (lbl, t)
        self.tablas_lay.addStretch()
        self.scroll.setWidget(cont)
        raiz.addWidget(self.scroll, 1)

        # pie
        pie = grupo()
        pl2 = QVBoxLayout(pie)
        ft = QHBoxLayout()
        ft.addWidget(etiqueta("TOTAL:", "seccion"))
        ft.addStretch()
        self.lbl_total = etiqueta("Q 0.00", "seccion")
        ft.addWidget(self.lbl_total)
        pl2.addLayout(ft)
        self.lbl_letras = QLabel("")
        self.lbl_letras.setWordWrap(True)
        pl2.addWidget(self.lbl_letras)
        raiz.addWidget(pie)

        # pestañas de hojas
        self.pest_scroll = QScrollArea()
        self.pest_scroll.setFixedHeight(46)
        self.pest_scroll.setWidgetResizable(True)
        self.pest_scroll.setFrameShape(QFrame.NoFrame)
        pw = QWidget()
        self.pest_lay = QHBoxLayout(pw)
        self.pest_lay.setContentsMargins(0, 0, 0, 0)
        self.pest_scroll.setWidget(pw)
        raiz.addWidget(self.pest_scroll)

        # spinner de Actualizar
        self._ang = 0
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._girar)
        self.worker = None

    # ── línea / voltaje / búsqueda ──────────────────────────────────
    def linea_actual(self):
        return self.cmb_linea.currentData()

    def voltaje_actual(self):
        return VOLTAJES[self.cmb_volt.currentIndex()][1]

    def voltaje_compatible(self, c):
        v = self.voltaje_actual()
        return (not c["v76"] and not c["v199"]) or bool(c[v])     # ninguno marcado = ambos

    def lineas_disponibles(self, c):
        """Líneas donde el material se puede agregar (ya filtradas por voltaje)."""
        if c is None or not self.voltaje_compatible(c):
            return []
        return [(k, txt) for k, _, txt in LINEAS if c.get(k)]

    def max_filas(self):
        return max(8, (self.height() - 780) // 22 + 8)

    def proveedor_codigo(self, texto, todos):
        return self._proveedor(texto, todos, "codigo")

    def proveedor_descripcion(self, texto, todos):
        return self._proveedor(texto, todos, "descripcion")

    def _proveedor(self, texto, todos, campo):
        """
        Busca en el campo activo y ordena: exacto, empieza con, palabra
        empieza con (solo descripción) y contiene. Los materiales sin
        ninguna línea marcada no se ofrecen para captura.
        """
        t = norm(texto)
        filas = []
        for i, c in enumerate(self.estado.bd):
            if not any(c.get(k) for k, _, _ in LINEAS):
                continue
            nivel = 0 if not t else _nivel_coincidencia(c[campo], t, campo)
            if nivel > 3:
                continue
            filas.append((nivel, i, c["codigo"], c["descripcion"],
                          self.voltaje_compatible(c)))
        filas.sort(key=lambda x: (x[0], x[1]))
        return [(cod, desc, ok) for _, _, cod, desc, ok in filas[:300]]

    def elegir(self, codigo):
        c = self.estado.comp(codigo)
        if not c:
            return
        self.comp_sel = c
        self.cod.setText(c["codigo"])
        self.desc.setText(c["descripcion"])
        self.lbl_unidad.setText(f"Unidad: {c['unidad']}")
        self.set_validador_cantidad(c["unidad"])
        self.actualizar_lineas(c)
        self.cant.setFocus()
        self.validar_agregar()

    def actualizar_lineas(self, c=None):
        """Después de elegir el material, muestra SOLO las líneas donde cabe."""
        self.cmb_linea.blockSignals(True)
        self.cmb_linea.clear()
        if c is None:
            self.cmb_linea.setEnabled(False)
        else:
            opciones = self.lineas_disponibles(c)
            for k, txt in opciones:
                self.cmb_linea.addItem(txt, k)
            self.cmb_linea.setEnabled(bool(opciones))
            if len(opciones) == 1:
                self.cmb_linea.setCurrentIndex(0)
            elif opciones:
                # Conservar la línea del último material agregado si el nuevo
                # material también la permite; si no, el usuario debe elegir.
                i = self.cmb_linea.findData(self.ultima_linea)
                self.cmb_linea.setCurrentIndex(i if i >= 0 else -1)
            sin_linea = not any(c.get(k) for k, _, _ in LINEAS)
            if not self.voltaje_compatible(c):
                self.lbl_unidad.setText(
                    f"Unidad: {c['unidad']}  — no habilitado para el voltaje actual")
            elif sin_linea:
                self.lbl_unidad.setText(
                    f"Unidad: {c['unidad']}  — sin línea marcada")
            elif not opciones:
                self.lbl_unidad.setText(
                    f"Unidad: {c['unidad']}  — sin línea compatible")
        self.cmb_linea.blockSignals(False)
        self.validar_agregar()

    def limpiar_campos_busqueda(self):
        """Ctrl+Espacio: borra los campos; el campo activo abre su propia lista."""
        self.cod.popup.hide()
        self.desc.popup.hide()
        self.cod.clear()
        self.desc.clear()
        self.cant.clear()
        self.limpiar_sel()
        self.set_validador_cantidad("")

    def redondeo_celda_cambiada(self, k, r, col, accion):
        if self._cargando:
            return
        base, hoja = self.tab
        modo = accion.data() if accion is not None else MODO_DEFECTO
        if modo not in MODOS_VALIDOS:
            modo = MODO_DEFECTO
        lista = self.estado.filas(base, hoja).get(k, [])
        if r >= len(lista):
            return
        codigo = lista[r]["codigo"]
        if hoja in EDITABLES:
            if not (self.editable() or self.estado.es_ramal(base)):
                return
            self._apuntar(base, [hoja], "Cambiar redondeo")
            self.estado.set_modo_linea(base, hoja, k, codigo, modo)
        elif hoja == "VARIACION" and col in (5, 8):
            # El redondeo de Contratado/Modificado en VARIACIÓN pertenece a la
            # fila de ANE/ANE-VAR de la que sale el total.
            origen = "ANE" if col == 5 else "ANE_VAR"
            self._apuntar(base, ["ANE", "ANE_VAR"], "Cambiar redondeo")
            self.estado.set_modo_linea(base, origen, k, codigo, modo)
        elif hoja in DEPENDEN_DE_VAR:
            self._apuntar(base, [], "Cambiar redondeo")
            self.estado.set_modo_doc(base, k, codigo, modo)
        else:
            return
        self.estado.cambios = True
        # Diferir el refresco evita destruir el botón que abrió el menú.
        QTimer.singleShot(0, self.refrescar_vista)

    def cod_editado(self, texto):
        c = self.estado.comp(texto.strip())
        if c:
            self.elegir(c["codigo"])
        else:
            self.limpiar_sel()

    def desc_editado(self, texto):
        c = next((x for x in self.estado.bd if norm(x["descripcion"]) == norm(texto)), None)
        if c:
            self.elegir(c["codigo"])
        else:
            self.limpiar_sel()

    def limpiar_sel(self):
        self.comp_sel = None
        self.lbl_unidad.setText("")
        self.actualizar_lineas(None)
        self.validar_agregar()

    def set_validador_cantidad(self, unidad):
        patron = r"^\d{0,9}(\.\d{0,3})?$" if es_km(unidad) else r"^\d{0,9}$"
        self.cant.setValidator(QRegularExpressionValidator(QRegularExpression(patron)))
        if self.cant.text() and not self.cant.hasAcceptableInput():
            self.cant.clear()

    @staticmethod
    def parse_cantidad(texto, unidad):
        t = texto.strip().rstrip(".")
        if not t:
            return None
        try:
            v = Decimal(t)
        except InvalidOperation:
            return None
        if es_km(unidad):
            v = v.quantize(Decimal("0.001"))
        elif v != v.to_integral_value():
            return None
        return v if v > 0 else None

    def editable(self):
        """Las hojas editables admiten captura, también en el Ramal.

        En el Ramal solo son editables las filas propias: las que vienen de los
        proyectos se validan fila por fila con fila_propia / codigo_heredado.
        """
        return bool(self.tab) and self.tab[1] in EDITABLES

    def fila_propia(self, k, r):
        """True si la fila mostrada es propia del Ramal (no viene de proyectos)."""
        if not self.tab:
            return True
        lista = self.estado.filas(self.tab[0], self.tab[1]).get(k, [])
        if r < 0 or r >= len(lista):
            return True
        return lista[r].get("propio", True)

    def material_en_linea_actual(self):
        """Indica si el material seleccionado ya existe en la línea actual."""
        if self.comp_sel is None or not self.editable():
            return False
        k = self.linea_actual()
        if k is None:
            return False
        base, hoja = self.tab
        # En el Ramal un código que viene de los proyectos no se modifica:
        # el botón debe decir «Agregar» y quedar deshabilitado.
        if (self.estado.es_ramal(base)
                and self.estado.codigo_heredado(hoja, k, self.comp_sel["codigo"])):
            return False
        lista = base["items"].get(hoja, {}).get(k, [])
        return any(it["codigo"] == self.comp_sel["codigo"] for it in lista)

    def validar_agregar(self):
        ok = (self.editable() and self.comp_sel is not None and
              self.linea_actual() is not None and
              self.parse_cantidad(self.cant.text(), self.comp_sel["unidad"]) is not None)
        tip = ""
        if ok and self.tab and self.estado.es_ramal(self.tab[0]):
            if self.estado.codigo_heredado(self.tab[1], self.linea_actual(),
                                           self.comp_sel["codigo"]):
                ok = False
                tip = ("Este código ya viene de los proyectos; en el Ramal solo "
                       "se puede cambiar su redondeo.")
        self.btn_add.setEnabled(ok)
        self.btn_add.setToolTip(tip)
        self.btn_add.setText("Modificar" if self.material_en_linea_actual() else "Agregar")

    def linea_cambiada(self):
        self.validar_agregar()

    def elegir_linea(self, clave):
        """Selecciona una línea del combo si el material la permite."""
        if self.comp_sel is None or not self.cmb_linea.isEnabled():
            return
        i = self.cmb_linea.findData(clave)
        if i >= 0:
            self.cmb_linea.setCurrentIndex(i)

    def cambiar_voltaje(self):
        """Alterna entre los voltajes disponibles (7.6 / 19.9 kV)."""
        if self.cmb_volt.count() > 1:
            self.cmb_volt.setCurrentIndex((self.cmb_volt.currentIndex() + 1)
                                          % self.cmb_volt.count())

    def voltaje_cambiado(self):
        """Avisa (sin borrar) de componentes ya agregados que no corresponden al nuevo voltaje."""
        if self.comp_sel is not None:
            self.actualizar_lineas(self.comp_sel)
        if not self.editable():
            return
        base, hoja = self.tab
        v = self.voltaje_actual()
        fuera = []
        for lista in (base["items"].get(hoja) or {}).values():
            for it in lista:
                c = self.estado.comp(it["codigo"])
                if c and (c["v76"] or c["v199"]) and not c[v]:
                    fuera.append(f"{c['codigo']} – {c['descripcion']}")
        if fuera:
            msg(self, "Aviso de voltaje",
                "Estos componentes ya agregados no corresponden al voltaje seleccionado "
                "(no se eliminaron):\n\n  • " + "\n  • ".join(fuera), QMessageBox.Warning)

    def colapsar(self):
        vis = not self.panel.isVisible()
        # El Ramal guarda su propia preferencia: arranca plegado y no comparte
        # el estado del panel de los proyectos.
        if self.tab and self.estado.es_ramal(self.tab[0]):
            self._panel_ramal = vis
        else:
            self._panel_visible = vis
        self.panel.setVisible(vis)
        self.btn_colapsar.setText("▲" if vis else "▼")

    # ── agregar / editar / eliminar componentes de la hoja ──────────
    def agregar(self):
        if not self.btn_add.isEnabled():
            return
        c, k = self.comp_sel, self.linea_actual()
        if c is None or k is None:
            msg(self, "Falta la línea",
                "Primero seleccione el material; luego el programa mostrará las líneas "
                "donde se puede agregar.", QMessageBox.Warning)
            return
        cant = self.parse_cantidad(self.cant.text(), c["unidad"])
        if cant is None:
            return
        base, hoja = self.tab
        if (self.estado.es_ramal(base)
                and self.estado.codigo_heredado(hoja, k, c["codigo"])):
            # Guarda defensiva: las filas heredadas solo cambian su redondeo.
            return
        self._apuntar(base, [hoja], "Agregar material")
        lineas = base.setdefault("items", {}).setdefault(hoja, _vacio())
        lista = lineas.setdefault(k, [])
        previo = next((it for it in lista if it["codigo"] == c["codigo"]), None)
        if previo:                                  # misma tabla → se modifica
            previo["cantidad"] = str(cant)
            previo.setdefault("redondeo", MODO_DEFECTO)
        else:                                       # otra tabla / nuevo → se agrega
            nuevo = {"codigo": c["codigo"], "cantidad": str(cant),
                     "redondeo": MODO_DEFECTO}
            lista.insert(self.estado.posicion_por_codigo(lista, c["codigo"]), nuevo)
        self.estado.cambios = True
        self.ultima_linea = k
        self.cod.clear()
        self.desc.clear()
        self.cant.clear()
        self.limpiar_sel()
        self.set_validador_cantidad("")
        self.refrescar_vista()
        self.cod.setFocus()

    def cantidad_editada(self, k, item):
        if item.column() != 4 or not self.editable() or self._cargando:
            return
        base, hoja = self.tab
        if self.estado.es_ramal(base):
            # En el Ramal la lista guardada no sigue el orden de la vista, así
            # que la fila se localiza por código y solo si es propia.
            vista = self.estado.filas(base, hoja).get(k, [])
            fila = vista[item.row()] if item.row() < len(vista) else None
            if fila is None or not fila.get("propio"):
                return
            it_data = next((it for it in ((base.get("items", {}).get(hoja) or {})
                                          .get(k, []))
                            if it.get("codigo") == fila["codigo"]), None)
            if it_data is None:
                return
            c = self.estado.comp(it_data["codigo"])
        else:
            it_data = base["items"][hoja][k][item.row()]
            c = self.estado.comp(it_data["codigo"])
        unidad = c["unidad"] if c else ""
        v = self.parse_cantidad(item.text(), unidad)
        if v is None:
            msg(self, "Cantidad no válida",
                "Kilómetro: hasta 3 decimales.\nOtras unidades: solo números enteros mayores a 0.",
                QMessageBox.Warning)
        else:
            self._apuntar(base, [hoja], "Modificar cantidad")
            it_data["cantidad"] = str(v)
            self.estado.cambios = True
        self.refrescar_vista()

    def mover_fila(self, k, src, dst):
        """Reordena una fila arrastrada y recuerda que el orden ya es manual."""
        if not self.editable():
            return
        base, hoja = self.tab
        if self.estado.es_ramal(base):
            # El orden visible del Ramal es siempre el de CODIGOS_SICOA.
            return
        lista = base["items"][hoja][k]
        if not lista:
            return
        src = max(0, min(src, len(lista) - 1))
        dst = max(0, min(dst, len(lista) - 1))
        if src == dst:
            return
        self._apuntar(base, [hoja], "Mover fila")
        lista.insert(dst, lista.pop(src))
        clave = f"{hoja}:{k}"
        manual = base.setdefault("orden_manual", [])
        if clave not in manual:
            manual.append(clave)
        self.estado.cambios = True
        self.refrescar_vista()
        self.secciones[k][1].selectRow(dst)

    def eliminar_filas(self, k):
        if not self.editable():
            return
        base, hoja = self.tab
        tabla = self.secciones[k][1]
        filas = sorted({i.row() for i in tabla.selectedIndexes()}, reverse=True)
        if not filas:
            return
        if self.estado.es_ramal(base):
            # Solo se borran las filas capturadas en el Ramal; las que vienen
            # de los proyectos se quedan con un único aviso.
            vista = self.estado.filas(base, hoja).get(k, [])
            codigos, heredada = [], False
            for r in filas:
                if r >= len(vista):
                    continue
                if self.fila_propia(k, r):
                    codigos.append(vista[r]["codigo"])
                else:
                    heredada = True
            if heredada:
                msg(self, "Filas del proyecto",
                    "Las filas que vienen de los proyectos no se eliminan desde el Ramal.",
                    QMessageBox.Warning)
            if not codigos:
                return
            self._apuntar(base, [hoja], "Eliminar filas")
            lineas = base.setdefault("items", {}).setdefault(hoja, _vacio())
            lineas[k] = [it for it in lineas.get(k, [])
                         if it.get("codigo") not in codigos]
            if not lineas[k]:
                clave = f"{hoja}:{k}"
                if clave in (base.get("orden_manual") or []):
                    base["orden_manual"].remove(clave)
            self.estado.cambios = True
            self.refrescar_vista()
            return
        self._apuntar(base, [hoja], "Eliminar filas")
        for r in filas:
            base["items"][hoja][k].pop(r)
        if not base["items"][hoja][k]:
            clave = f"{hoja}:{k}"
            if clave in (base.get("orden_manual") or []):
                base["orden_manual"].remove(clave)
        self.estado.cambios = True
        self.refrescar_vista()

    # ── pestañas y vista ────────────────────────────────────────────
    def refrescar_todo(self):
        self.estado.asegurar_visibles()
        tabs = self.estado.tabs()
        if not self.tab or not any(b is self.tab[0] and h == self.tab[1] for b, h in tabs):
            self.tab = tabs[0]
        limpiar(self.pest_lay)
        for base, hoja in tabs:
            b = QPushButton(self.estado.titulo(base, hoja))
            b.setCheckable(True)
            b.setChecked(base is self.tab[0] and hoja == self.tab[1])
            col = base["color"]
            borde = "4px solid white" if b.isChecked() else "1px solid #21618c"
            b.setStyleSheet(f"QPushButton {{ background-color: {col}; color: {contraste(col)}; "
                            f"border: {borde}; border-radius: 3px; padding: 6px 12px; "
                            f"font-weight: {'bold' if b.isChecked() else 'normal'}; }}")
            b.clicked.connect(partial(self.seleccionar_tab, base, hoja))
            self.pest_lay.addWidget(b)
        self.pest_lay.addStretch()
        self.refrescar_vista()

    def seleccionar_tab(self, base, hoja):
        self.tab = (base, hoja)
        self.refrescar_todo()

    def _indice_tab(self, tabs):
        if self.tab is None:
            return 0
        return next((i for i, (b, h) in enumerate(tabs)
                     if b is self.tab[0] and h == self.tab[1]), 0)

    def siguiente_hoja(self):
        tabs = self.estado.tabs()
        if len(tabs) > 1:
            i = self._indice_tab(tabs)
            self.seleccionar_tab(*tabs[(i + 1) % len(tabs)])

    def anterior_hoja(self):
        tabs = self.estado.tabs()
        if len(tabs) > 1:
            i = self._indice_tab(tabs)
            self.seleccionar_tab(*tabs[(i - 1) % len(tabs)])

    def alternar_sin_redondeo(self):
        """
        Alterna en bloque los precios de la hoja ANE / ANE-VAR activa.

        Primer uso: las filas en Redondear pasan a 0.00 SR (multiplicación sin
        redondeo). Al ejecutarlo otra vez, las que están en 0.00 SR vuelven a
        Redondear. Las filas con Redondear.menos (o con el Redondear.más
        heredado de archivos viejos) se dejan como están para no cambiar una
        decisión manual del usuario.
        """
        if not self.tab:
            return
        base, hoja = self.tab
        if hoja not in EDITABLES:
            self.toast.mostrar("Solo en ANE / ANE-VAR", self, "#e67e22")
            return
        if not (self.editable() or self.estado.es_ramal(base)):
            return
        filas = self.estado.filas(base, hoja)
        # El destino se decide con el estado actual de la hoja: mientras haya
        # filas en Redondear, se pasan a SR; cuando ya no quede ninguna, se
        # devuelven las SR a Redondear.
        hay_redondear = any(f["redondeo"] == "ROUND"
                            for k, _, _ in LINEAS for f in filas[k])
        hay_sr = any(f["redondeo"] == "SR"
                     for k, _, _ in LINEAS for f in filas[k])
        if hay_redondear:
            origen, destino, aviso = "ROUND", "SR", "Precios en 0.00 SR (sin redondeo)"
        elif hay_sr:
            origen, destino, aviso = "SR", "ROUND", "Precios con Redondear"
        else:
            return
        self._apuntar(base, [hoja], "Alternar 0.00 SR")
        cambios = 0
        for k, _, _ in LINEAS:
            for f in filas[k]:
                if f["redondeo"] != origen:
                    continue
                if self.estado.set_modo_linea(base, hoja, k, f["codigo"], destino):
                    cambios += 1
        if not cambios:
            return
        self.refrescar_vista()
        self.toast.mostrar(aviso, self)

    # ── deshacer / rehacer ──────────────────────────────────────────
    def _snap(self, base, hojas, etiqueta):
        """
        Copia mínima del estado que puede tocar UNA acción.

        Solo se guardan las hojas implicadas (una, o dos cuando VARIACIÓN
        cambia el redondeo de ANE y ANE-VAR a la vez) más el orden manual y
        los redondeos de documento. Copiar el proyecto completo en cada clic
        llenaría la memoria sin ninguna ventaja.
        """
        items = base.get("items") or {}
        return {
            "base": base,
            "etiqueta": etiqueta,
            "hojas": {h: (copy.deepcopy(items[h]) if h in items else None)
                      for h in hojas},
            "orden_manual": list(base.get("orden_manual") or []),
            "redondeo_doc": dict(base.get("redondeo_doc") or {}),
            "cambios": self.estado.cambios,
        }

    def _apuntar(self, base, hojas, etiqueta):
        """Registra el estado ANTERIOR de una acción, para Ctrl+Z."""
        if self._en_historial or base is None:
            return
        self._pila_deshacer.append(self._snap(base, hojas, etiqueta))
        if len(self._pila_deshacer) > HISTORIAL_MAX:
            self._pila_deshacer.pop(0)
        self._pila_rehacer.clear()

    def _aplicar(self, snap):
        """Devuelve una base al snapshot y repinta toda la ventana."""
        base = snap["base"]
        items = base.setdefault("items", {})
        for hoja, valor in snap["hojas"].items():
            if valor is None:
                items.pop(hoja, None)
            else:
                items[hoja] = copy.deepcopy(valor)
        base["orden_manual"] = list(snap["orden_manual"])
        base["redondeo_doc"] = dict(snap["redondeo_doc"])
        self.estado.cambios = snap["cambios"]
        self.refrescar_todo()

    def deshacer(self):
        """Ctrl+Z: vuelve al estado anterior a la última acción registrada."""
        if self._en_historial:
            return
        # En un campo de texto, Ctrl+Z deshace lo escrito, no la última acción.
        foco = QApplication.focusWidget()
        if isinstance(foco, QLineEdit) and foco.isUndoAvailable():
            foco.undo()
            return
        if not self._pila_deshacer:
            self.toast.mostrar("Nada que deshacer", self, "#e67e22")
            return
        snap = self._pila_deshacer.pop()
        self._pila_rehacer.append(
            self._snap(snap["base"], list(snap["hojas"]), snap["etiqueta"]))
        self._en_historial = True
        try:
            self._aplicar(snap)
        finally:
            self._en_historial = False
        self.toast.mostrar(f"Deshecho: {snap['etiqueta']}", self)

    def rehacer(self):
        """Ctrl+Y: reaplica lo que se quitó con Ctrl+Z."""
        if self._en_historial:
            return
        foco = QApplication.focusWidget()
        if isinstance(foco, QLineEdit) and foco.isRedoAvailable():
            foco.redo()
            return
        if not self._pila_rehacer:
            self.toast.mostrar("Nada que rehacer", self, "#e67e22")
            return
        snap = self._pila_rehacer.pop()
        self._pila_deshacer.append(
            self._snap(snap["base"], list(snap["hojas"]), snap["etiqueta"]))
        self._en_historial = True
        try:
            self._aplicar(snap)
        finally:
            self._en_historial = False
        self.toast.mostrar(f"Rehecho: {snap['etiqueta']}", self)

    def _limpiar_historial(self):
        """Guardar/abrir/crear fijan una base nueva: el historial deja de valer."""
        self._pila_deshacer.clear()
        self._pila_rehacer.clear()

    _cargando = False

    def _valores_fila(self, hoja, f):
        if hoja in EDITABLES:
            return [f["codigo"], f["desc"], f["unidad"], fmt_q(f["precio"]),
                    fmt_cant(f["cantidad"], f["unidad"]),
                    texto_total(f["total"], f["redondeo"])]
        if hoja == "VARIACION":
            tipo = f["tipo"]
            return [
                f["codigo"], f["desc"], f["unidad"],
                fmt_cant(f["cantidad_ane"], f["unidad"]),
                fmt_q(f["precio"]),
                texto_total(f["total_ane"], f["modo_ane"]),
                fmt_cant(f["cantidad_av"], f["unidad"]),
                fmt_q(f["precio"]),
                texto_total(f["total_av"], f["modo_av"]),
                tipo,
                fmt_cant(f["cantidad_var"], f["unidad"]),
                fmt_q(f["precio"]),
                texto_total(f["monto_doc"], f["redondeo"]) if tipo == "OC" else "",
                texto_total(f["monto_doc"], f["redondeo"]) if tipo == "OTS" else "",
                texto_total(f["monto_doc"], f["redondeo"]) if tipo == "ATE" else "",
            ]
        return [
            f["codigo"], f["desc"], f["unidad"],
            fmt_cant(f["cantidad_ane"], f["unidad"]),
            fmt_q(f["precio"]),
            texto_total(f["total_ane"], f["modo_ane"]),
            fmt_cant(f["cantidad_av"], f["unidad"]),
            fmt_q(f["precio"]),
            texto_total(f["total_av"], f["modo_av"]),
            f["tipo"],
            fmt_cant(f["cantidad_var"], f["unidad"]),
            fmt_q(f["precio"]),
            texto_total(f["monto_doc"], f["redondeo"]),
        ]

    def _celda_con_redondeo(self, hoja, col, f):
        if hoja in EDITABLES:
            return col == 5
        if hoja == "VARIACION":
            return ((col == 5 and f["cantidad_ane"] != 0) or
                    (col == 8 and f["cantidad_av"] != 0) or
                    (col == 12 and f["tipo"] == "OC") or
                    (col == 13 and f["tipo"] == "OTS") or
                    (col == 14 and f["tipo"] == "ATE"))
        return col == 12

    def _columna_doc(self, hoja, f):
        if hoja == "VARIACION":
            return {"OC": 12, "OTS": 13, "ATE": 14}.get(f["tipo"])
        if hoja in ("OC", "OTS", "ATE"):
            return 12
        return None

    def _boton_redondeo(self, t, k, r, col, f, bruto, modo, diferencia=False):
        btn = QPushButton(texto_total(redondear_monto(bruto, modo), modo))
        btn.setFlat(True)
        btn.setFixedHeight(22)
        btn.setCursor(Qt.PointingHandCursor)
        btn.setToolTip("Haga clic para elegir cómo redondear "
                       + ("esta diferencia." if diferencia else "este total."))
        # La selección de la tabla no llega a los cellWidget: el rol y la fila
        # permiten que resaltar_botones_fila marque el texto en blanco.
        btn.setProperty("rol", "redondeo")
        btn.setProperty("fila", r)
        btn.setStyleSheet(
            "QPushButton { border: none; background: transparent;"
            " text-align: right; padding: 0px 2px; }"
            "QPushButton::menu-indicator { image: none; width: 0px; }"
            "QPushButton[sel=\"true\"] { color: white; }")
        menu = QMenu(btn)
        for txt_modo, m in MODOS_REDONDEO:
            valor = redondear_monto(bruto, m)
            accion = menu.addAction(texto_total(valor, m))
            accion.setData(m)
            accion.setCheckable(True)
            accion.setChecked(m == modo)
            accion.setToolTip(txt_modo)
        btn.setMenu(menu)
        menu.triggered.connect(partial(self.redondeo_celda_cambiada, k, r, col))
        t.setCellWidget(r, col, btn)

    def refrescar_vista(self):
        self._cargando = True
        base, hoja = self.tab
        filas = self.estado.filas(base, hoja)
        ed = self.editable()
        ed_redondeo = ed or self.estado.es_ramal(base)
        config_doc = hoja in DEPENDEN_DE_VAR
        derivada = hoja in DEPENDEN_DE_VAR
        es_ramal = self.estado.es_ramal(base)
        # En VARIACIÓN/OC/OTS/ATE los materiales no se capturan: el panel
        # queda plegado, pero la flecha sigue disponible para consultarlo.
        self.panel.setEnabled(ed)
        # El Ramal tiene su propia preferencia y arranca plegado; se abre con la
        # flecha para capturar sus filas propias.
        visible_panel = self._panel_ramal if es_ramal else self._panel_visible
        if derivada:
            self.panel.setVisible(False)
            self.btn_colapsar.setText("▼")
        else:
            self.panel.setVisible(visible_panel)
            self.btn_colapsar.setText("▲" if visible_panel else "▼")
        # El grupo se muestra en todas las hojas; en el Ramal arranca plegado.
        self.panel_g.setVisible(True)
        info = ""
        if not derivada and not es_ramal:
            if not ed:
                info = f"{ETIQUETA[hoja]}: vista provisional de solo lectura."
        self.lbl_info.setText(info)
        self.lbl_info.setVisible(bool(info))
        subtotales = self.estado.subtotales(filas)
        for k, titulo, _ in LINEAS:
            lbl, t = self.secciones[k]
            self._configurar_tabla(t, hoja)
            # El Ramal conserva el orden de CODIGOS_SICOA: no se arrastra.
            t.permitir_arrastre(ed and not es_ramal)
            lista = filas[k]
            visible = bool(lista)
            lbl.setVisible(visible)
            t.setVisible(visible)
            if hoja == "VARIACION":
                sub_f = sum((f["total_ane"] for f in lista), Decimal(0))
                sub_i = sum((f["total_av"] for f in lista), Decimal(0))
                lbl.setText(f"{titulo}   —   Contratado: {fmt_q(sub_f)}"
                            f"   |   Modificado: {fmt_q(sub_i)}"
                            f"   |   Variación: {fmt_q(subtotales[k])}")
            else:
                lbl.setText(f"{titulo}   —   Total tabla: {fmt_q(subtotales[k])}")
            t.blockSignals(True)
            t.setRowCount(0)
            t.setRowCount(len(lista))
            for r, f in enumerate(lista):
                vals = self._valores_fila(hoja, f)
                for col, v in enumerate(vals):
                    if self._celda_con_redondeo(hoja, col, f):
                        v = ""
                    it = QTableWidgetItem(str(v))
                    flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable
                    # En el Ramal solo se edita la cantidad de las filas propias.
                    if col == 4 and ed and (not es_ramal or f.get("propio", True)):
                        flags |= Qt.ItemIsEditable
                    it.setFlags(flags)
                    if col >= 3:
                        it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    t.setItem(r, col, it)

                if hoja in EDITABLES and ed_redondeo:
                    self._boton_redondeo(t, k, r, 5, f, f["bruto"],
                                         f["redondeo"])
                elif hoja == "VARIACION":
                    # Contratado y Modificado también permiten elegir su
                    # redondeo, igual que las hojas ANE / ANE-VAR.
                    if f["cantidad_ane"] != 0:
                        self._boton_redondeo(
                            t, k, r, 5, f,
                            f["precio"] * f["cantidad_ane"], f["modo_ane"])
                    if f["cantidad_av"] != 0:
                        self._boton_redondeo(
                            t, k, r, 8, f,
                            f["precio"] * f["cantidad_av"], f["modo_av"])
                if config_doc:
                    col_doc = self._columna_doc(hoja, f)
                    if col_doc is not None:
                        self._boton_redondeo(t, k, r, col_doc, f,
                                             f.get("bruto", D(0)),
                                             f["redondeo"], diferencia=True)
            t.blockSignals(False)
            resaltar_botones_fila(t)
            self._ajustar_alto_tabla(t)

        # Aviso permanente de VARIACIÓN: la variación real (modificado -
        # contratado) debe coincidir con la suma de OC + OTS + ATE.
        if hoja == "VARIACION":
            total_f = sum((f["total_ane"] for lista in filas.values() for f in lista),
                          Decimal(0))
            total_i = sum((f["total_av"] for lista in filas.values() for f in lista),
                          Decimal(0))
            total_doc = sum((f["monto_doc"] for lista in filas.values() for f in lista),
                            Decimal(0))
            diferencia = (total_i - total_f - total_doc).quantize(Decimal("0.01"))
            if diferencia != 0:
                self.alerta_var.setText(f"No Cuadra por ({fmt_q(diferencia)})")
                self.alerta_var.show()
            else:
                self.alerta_var.hide()
        else:
            self.alerta_var.hide()

        total = sum(subtotales.values(), Decimal(0))
        self.lbl_total.setText(fmt_q(total))
        f = self.lbl_total.font()
        f.setBold(True)
        self.lbl_total.setFont(f)
        self.lbl_letras.setText("Total en letras: " + total_en_letras(total))
        self._cargando = False
        QTimer.singleShot(0, self._ajustar_alturas_tablas)
        self.validar_agregar()
        self.actualizar_titulo()

    def actualizar_titulo(self):
        n = os.path.basename(self.estado.archivo) if self.estado.archivo else "sin guardar"
        # El aviso sustituye al antiguo asterisco: el usuario debe ver de un
        # golpe que hay trabajo sin guardar, también en la barra de tareas.
        if self.estado.cambios:
            n += " - Cambios pendientes"
        self.setWindowTitle(f"Anexos - WuMDrex – {n}")

    # ── ventanas ────────────────────────────────────────────────────
    def abrir_config(self):
        DlgConfig(self, self.estado).exec()
        self.refrescar_todo()

    def abrir_vis(self):
        DlgVisualizacion(self, self.estado, "principal").exec()
        self.refrescar_todo()

    def abrir_comp(self):
        DlgComponentes(self, self.estado).exec()
        self.refrescar_todo()

    # ── proyecto: nuevo / abrir / guardar ───────────────────────────
    def preguntar_guardar(self):
        """True si se puede continuar (guardó o descartó)."""
        if not self.estado.cambios:
            return True
        op = tres_opciones(self, "Cambios sin guardar", "¿Desea guardar los cambios del proyecto?",
                           "Guardar", "No guardar", "Cancelar")
        if op == 0:
            return self.guardar()
        return op == 1

    def nuevo_proyecto(self):
        """Reinicia el estado sin pasar por el diálogo de Proyecto."""
        if not self.preguntar_guardar():
            return False
        self.estado.reset()
        self.estado.bd = cargar_bd_inicial()
        self.tab = None
        self.estado.cambios = False
        self.refrescar_todo()
        self._limpiar_historial()
        return True

    def abrir_proyecto(self):
        """Abre un Excel exportado sin pasar por el diálogo de Proyecto."""
        if not self.preguntar_guardar():
            return False
        ruta, _ = QFileDialog.getOpenFileName(self, "Abrir proyecto", "",
                                              "Excel (*.xlsx *.xlsm)")
        if not ruta:
            return False
        try:
            cargar_proyecto_desde_excel(self.estado, ruta)
            self.tab = None
        except Exception as e:
            msg(self, "Error", f"No se pudo abrir el Excel: {e}", QMessageBox.Critical)
            return False
        self.estado.cambios = False
        self.refrescar_todo()
        self._limpiar_historial()
        self.toast.mostrar("Archivo importado", self, "#2980b9")
        return True

    def proyecto(self):
        dlg = DlgProyecto(self)
        if not dlg.exec():
            return
        if dlg.accion == "nuevo":
            self.nuevo_proyecto()
        else:
            self.abrir_proyecto()

    def guardar(self):
        e = self.estado
        if not e.archivo:
            ruta, _ = QFileDialog.getSaveFileName(self, "Guardar copia de la plantilla",
                                                  nombre_archivo_anexo(e), "Excel (*.xlsx)")
            if not ruta:
                return False
            if not ruta.lower().endswith(".xlsx"):
                ruta += ".xlsx"
            e.archivo = ruta
        try:
            totales = exportar_xlsx(e, e.archivo)
            e.totales_previos = {t: str(v) for t, v in totales.items()}
        except PermissionError:
            msg(self, "Error", "No se pudo escribir el archivo (¿está abierto en Excel?).",
                QMessageBox.Critical)
            return False
        except Exception as ex:
            msg(self, "Error al exportar", f"No se pudo generar el Excel:\n\n{ex}",
                QMessageBox.Critical)
            return False
        e.cambios = False
        self.actualizar_titulo()
        self._limpiar_historial()
        self.toast.mostrar("Archivo guardado", self)
        return True

    def guardar_como(self):
        """Guarda una copia en otra ruta sin cambiar la ruta actual si falla."""
        e = self.estado
        ruta, _ = QFileDialog.getSaveFileName(self, "Guardar como",
                                              nombre_archivo_anexo(e), "Excel (*.xlsx)")
        if not ruta:
            return False
        if not ruta.lower().endswith(".xlsx"):
            ruta += ".xlsx"
        anterior = e.archivo
        e.archivo = ruta
        if self.guardar():
            return True
        e.archivo = anterior
        return False

    # ── actualizar (🔄) ──────────────────────────────────────────────
    def _icono_giro(self):
        pm = QPixmap(30, 30)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing)
        p.setPen(QColor("#ecf0f1" if Tema.oscuro else "#2c3e50"))
        f = QFont("Segoe UI", 16)
        p.setFont(f)
        p.drawText(pm.rect(), Qt.AlignCenter, "🔄")
        p.end()
        return pm.transformed(QTransform().rotate(self._ang), Qt.SmoothTransformation)

    def _girar(self):
        self._ang = (self._ang + 30) % 360
        self.btn_act.setIcon(QIcon(self._icono_giro()))

    def actualizar(self):
        e = self.estado
        if not e.archivo or not os.path.exists(e.archivo):
            msg(self, "Actualizar", "Primero guarde el proyecto en Excel (💾).")
            return
        self.btn_act.setText("")
        self.timer.start(60)
        self.setEnabled(False)
        self.worker = ActualizarWorker(e, e.archivo)
        self.worker.listo.connect(self.fin_actualizar)
        self.worker.start()

    def fin_actualizar(self, res):
        self.timer.stop()
        self.btn_act.setIcon(QIcon())
        self.btn_act.setText("🔄")
        self.setEnabled(True)
        if res["error"]:
            msg(self, "Actualizar", res["error"], QMessageBox.Critical)
            return
        if res.get("verificacion"):
            # Verificación de solo lectura: el estado y la vista no cambian,
            # así que no se tocan los totales previos ni se refresca la tabla.
            DlgReporte(self, res, self.estado).exec()
            return
        for titulo, _, nuevo in res["totales"]:
            self.estado.totales_previos[titulo] = nuevo
        DlgReporte(self, res, self.estado).exec()
        self.refrescar_todo()

    # ── comandos / cierre ───────────────────────────────────────────
    def _foco_codigo(self):
        self.cod.setFocus()
        self.cod.selectAll()

    def _foco_desc(self):
        self.desc.setFocus()
        self.desc.selectAll()

    def _foco_linea(self):
        if self.cmb_linea.isEnabled():
            self.cmb_linea.setFocus()

    def _comandos_globales(self):
        """Comandos que se instalan como atajos de toda la ventana."""
        return [
            ("Proyecto", "Ctrl+N", "Nuevo proyecto", self.nuevo_proyecto),
            ("Proyecto", "Ctrl+O", "Abrir proyecto desde Excel", self.abrir_proyecto),
            ("Proyecto", "Ctrl+P", "Abrir ventana Proyecto (Nuevo / Abrir)", self.proyecto),
            ("Proyecto", "Ctrl+S", "Guardar / exportar Excel", self.guardar),
            ("Proyecto", "Ctrl+Shift+S", "Guardar como…", self.guardar_como),
            ("Edición", "Ctrl+Z", "Deshacer el último cambio", self.deshacer),
            ("Edición", "Ctrl+Y", "Rehacer el cambio deshecho", self.rehacer),
            ("Ventanas", "F2", "Abrir Componentes (materiales)", self.abrir_comp),
            ("Ventanas", "F3", "Abrir Configuración", self.abrir_config),
            ("Ventanas", "F4", "Abrir Visualización", self.abrir_vis),
            ("Ventanas", "F5", "Actualizar totales", self.actualizar),
            ("Materiales", "Ctrl+M", "Enfocar el campo Código", self._foco_codigo),
            ("Materiales", "Ctrl+Shift+M", "Enfocar el campo Componente", self._foco_desc),
            ("Materiales", "Alt+L", "Enfocar el selector Línea", self._foco_linea),
            ("Materiales", "Ctrl+1", "Elegir línea Media tensión",
             lambda: self.elegir_linea("MT")),
            ("Materiales", "Ctrl+2", "Elegir línea Media más baja tensión",
             lambda: self.elegir_linea("MTBT")),
            ("Materiales", "Ctrl+3", "Elegir línea Baja tensión",
             lambda: self.elegir_linea("BT")),
            ("Materiales", "Ctrl+4", "Elegir línea General",
             lambda: self.elegir_linea("G")),
            ("Materiales", "Alt+V", "Alternar voltaje 7.6 / 19.9 kV", self.cambiar_voltaje),
            ("Materiales", "Ctrl+Q", "Agregar material", self.agregar),
            ("Hojas", "Ctrl+Tab", "Ir a la siguiente hoja", self.siguiente_hoja),
            ("Hojas", "Ctrl+Shift+Tab", "Ir a la hoja anterior", self.anterior_hoja),
            ("Hojas", "Ctrl+R", "Alternar Redondear / 0.00 SR en los precios de la hoja activa",
             self.alternar_sin_redondeo),
            ("Interfaz", "Alt+B", "Alternar Día / Noche", alternar_tema),
            ("Interfaz", "Ctrl+H", "Mostrar / ocultar panel de materiales", self.colapsar),
            ("Comandos", "F1", "Mostrar esta tabla", self.mostrar_comandos),
        ]

    def _comandos_tabla(self):
        filas = [(g, c, d) for g, c, d, _ in self._comandos_globales()]
        filas.extend(COMANDOS_CONTEXTO)
        orden = ["Proyecto", "Edición", "Ventanas", "Materiales", "Tabla", "Hojas",
                 "Buscador", "Interfaz", "Comandos"]
        filas.sort(key=lambda f: orden.index(f[0]) if f[0] in orden else len(orden))
        return filas

    def _instalar_atajos(self):
        """Registra los comandos globales y conserva sus QShortcut vivos."""
        for _, secuencia, _, accion in self._comandos_globales():
            atajo = QShortcut(QKeySequence(secuencia), self)
            atajo.activated.connect(accion)
            self._atajos.append(atajo)

    def mostrar_comandos(self):
        DlgComandos(self, self._comandos_tabla()).exec()

    def keyPressEvent(self, e):
        # En la ventana principal Esc no cierra el programa por sí solo; los
        # diálogos secundarios conservan su comportamiento normal de QDialog.
        # Con cambios sin guardar, Esc se comporta como la X de la ventana y
        # deja que closeEvent pregunte UNA sola vez (Guardar/No/Cancelar).
        if e.key() == Qt.Key_Escape:
            e.accept()
            if self.estado.cambios:
                self.close()
        else:
            super().keyPressEvent(e)

    def closeEvent(self, ev):
        if self.estado.cambios:
            op = tres_opciones(self, "Salir", "Hay cambios sin guardar. ¿Desea guardarlos?",
                               "Guardar", "Salir sin guardar", "Cancelar")
            if op == 2 or (op == 0 and not self.guardar()):
                ev.ignore()
                return
        ev.accept()
        QApplication.quit()


# ═════════════════════════════════════════════════════════════════════
def main():
    # Validación silenciosa de licencia: sin ventanas ni avisos al usuario.
    print("[ANEXOS] Iniciando validación de licencia...")
    if not check_license():
        print("[ANEXOS] Licencia inválida o expirada. El programa no se iniciará.")
        sys.exit(1)
    print("[ANEXOS] Licencia válida. Iniciando interfaz.")

    # La plantilla editable vive junto al .exe (no en _MEIPASS): se crea la
    # primera vez copiándola de los recursos. Si falla, la app igual arranca.
    try:
        preparar_plantilla_editable()
    except Exception as e:
        print("[ANEXOS] Aviso: no se pudo preparar la plantilla editable:", e)

    _configurar_app_id_windows()
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    if not icono_app().isNull():
        app.setWindowIcon(icono_app())
    aplicar_tema()
    w = Principal()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
