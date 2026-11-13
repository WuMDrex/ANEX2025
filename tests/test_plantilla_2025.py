# -*- coding: utf-8 -*-
"""Validación de la plantilla nueva `Plantilla/ANEXO SICOA 2025.xlsm`.

Cubre:
  1) que la configuración apunte al archivo nuevo (main.py, ANEX.spec, LEEME);
  2) la ESTRUCTURA de la plantilla (openpyxl, sin abrir Excel): hoja/tabla de
     la BD, las 6 columnas de clasificación, las 12 hojas base y sus 4 tablas;
  3) la lectura de la BD por la vía real (`cargar_bd_inicial`);
  4) una EXPORTACIÓN real con xlwings/COM (3 códigos distintos + proyecto);
  5) el roundtrip `cargar_proyecto_desde_excel`;
  6) que la plantilla real quede intacta (hash, tamaño y fecha).

La exportación se hace UNA vez y se reutiliza (cada exportación tarda ~30 s).
Los archivos quedan en %TEMP%\\opencode\\plantilla2025 para inspección.
"""

import hashlib
import os
import shutil
import sys

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RAIZ not in sys.path:
    sys.path.insert(0, RAIZ)

import pytest  # noqa: E402
from openpyxl import load_workbook  # noqa: E402
from openpyxl.utils import range_boundaries  # noqa: E402

import main  # noqa: E402

SALIDA = os.path.join(os.environ.get("LOCALAPPDATA") or RAIZ, "Temp",
                      "opencode", "plantilla2025")
COD_NUEVO = "2500"          # numérico > UMBRAL_CODIGO_ATE → debe ser ATE
DESC_NUEVA = "EQUIPO NUEVO DE PRUEBA ATE"
PRECIO_NUEVO = "1000.00"
ABREV = "PRUEBA2025"
HOJAS_PROYECTO = ["ANE", "ANE_VAR", "VARIACION", "OC", "OTS", "ATE"]

_cache = {}


def _sha256(ruta):
    h = hashlib.sha256()
    with open(ruta, "rb") as f:
        for bloque in iter(lambda: f.read(1 << 20), b""):
            h.update(bloque)
    return h.hexdigest()


# Huella de la plantilla ANTES de cualquier prueba de este archivo.
if "hash_previo" not in _cache:
    _cache["hash_previo"] = _sha256(main.PLANTILLA_XLSX)
    _cache["stat_previo"] = os.stat(main.PLANTILLA_XLSX)


def _norm(x):
    return main._xl_normalizar_etiqueta(x)


def _tabla(ws, nombre):
    """Tabla por nombre (openpyxl: `ws.tables[n]` devuelve la REF, no el objeto)."""
    for t in ws.tables.values():
        if t.name.upper() == str(nombre).upper():
            return t
    raise KeyError(f"{nombre} no existe en {ws.title}: "
                   f"{[t.name for t in ws.tables.values()]}")


def _filas_tabla(ws, nombre):
    """Filas de datos de una tabla (sin el encabezado), como tuplas de valores."""
    tabla = _tabla(ws, nombre)
    c1, f1, c2, f2 = range_boundaries(tabla.ref)
    salida = []
    for f in range(f1 + 1, f2 + 1):
        salida.append(tuple(ws.cell(row=f, column=c).value for c in range(c1, c2 + 1)))
    return salida


def _encabezado_tabla(ws, nombre):
    """Encabezado (fila 1) de una tabla, como lista de valores."""
    c1, f1, c2, _ = range_boundaries(_tabla(ws, nombre).ref)
    return [ws.cell(row=f1, column=c).value for c in range(c1, c2 + 1)]


def _errores_libro(wb):
    males = ("#N/A", "#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#N/A", "Err:")
    out = []
    for ws in wb.worksheets:
        for fila in ws.iter_rows():
            for cel in fila:
                v = cel.value
                if isinstance(v, str) and v.strip() in males:
                    out.append((ws.title, cel.coordinate, v))
    return out


def _fila_total(ws):
    """Fila donde está la palabra TOTAL y valor de la última columna numérica."""
    for f in range(1, ws.max_row + 1):
        for c in (1, 2, 3):
            if _norm(ws.cell(row=f, column=c).value) == "TOTAL":
                return f
    return None


# ═════════════════════════════════════════════════════════════════════
#  1) la configuración apunta a la plantilla 2025
# ═════════════════════════════════════════════════════════════════════
def test_config_apunta_a_plantilla_2025():
    ruta = main.PLANTILLA_XLSX
    assert os.path.basename(ruta) == "ANEXO SICOA 2025.xlsm", ruta
    assert os.path.exists(ruta), ruta
    assert "2026" not in os.path.basename(ruta)

    spec = os.path.join(RAIZ, "ANEX.spec")
    if os.path.exists(spec):
        with open(spec, encoding="utf-8", errors="replace") as f:
            txt = f.read()
        assert "ANEXO SICOA 2025.xlsm" in txt, "ANEX.spec no menciona la plantilla nueva"
        assert "ANEXO SICOA 2026" not in txt, "ANEX.spec sigue con la plantilla vieja"


# ═════════════════════════════════════════════════════════════════════
#  2) estructura de la plantilla (solo lectura, sin Excel)
# ═════════════════════════════════════════════════════════════════════
def test_estructura_bd_de_la_plantilla():
    wb = load_workbook(main.PLANTILLA_XLSX, data_only=True)
    try:
        assert main.HOJA_BD in wb.sheetnames, wb.sheetnames
        ws = wb[main.HOJA_BD]
        nombres = [str(t).upper() for t in ws.tables.keys()]
        assert main.TABLA_BD.upper() in nombres, nombres

        encabezado = [c.value for c in ws[1]]
        presentes = {_norm(v) for v in encabezado if v not in (None, "")}
        for col in (main.COL_BD_CODIGO, main.COL_BD_DESC,
                    main.COL_BD_UNIDAD, main.COL_BD_PRECIO):
            assert _norm(col) in presentes, f"falta la columna {col!r}: {encabezado}"
        for marca in ("MT", "MT + BT", "BT", "G", "7.6 kV", "19.9 kV"):
            assert _norm(marca) in presentes, f"falta la marca {marca!r}: {encabezado}"
        assert len([v for v in encabezado if v not in (None, "")]) == 10, encabezado

        filas = [r for r in ws.iter_rows(min_row=2, values_only=True)
                 if r and r[0] not in (None, "")]
        _cache["filas_bd"] = len(filas)
        assert len(filas) > 100, f"CODIGOS_SICOA solo tiene {len(filas)} filas"

        # El precio de la plantilla viene como TEXTO ('Q44,391.30'): D() debe
        # poder leerlo y, sobre todo, la exportación debe escribirlo numérico.
        no_parseables = [r[0] for r in filas if main.D(r[3]) == 0]
        assert not no_parseables, f"precios ilegibles: {no_parseables[:5]}"
    finally:
        wb.close()


def test_estructura_hojas_base():
    # data_only=False para poder ver las FÓRMULAS de referencia (B12 del
    # OC/OTS/ATE apunta al ANE por fórmula, no por texto).
    wb = load_workbook(main.PLANTILLA_XLSX, data_only=False)
    try:
        faltan = []
        for grupo in (main.BASE_COMUNIDAD, main.BASE_RAMAL):
            for hoja, nombre in grupo.items():
                if nombre not in wb.sheetnames:
                    faltan.append(nombre)
                    continue
                ws = wb[nombre]
                prefijo = main.TABLA_PREFIJO[hoja]
                prefijos = [f"{prefijo}_{main._xl_sufijo(k)}" for k, _, _ in main.LINEAS]
                nombres = [str(t).upper() for t in ws.tables.keys()]
                # Las tablas base se renombran por ÍNDICE (1..4), así que el
                # orden MT, MT_BT, BT, G tiene que ser el físico del machote.
                assert len(nombres) >= 4, f"{nombre}: solo {len(nombres)} tablas {nombres}"
                # El sufijo real del machote (_COMUNIDAD / _RAMAL) se deduce del
                # primer nombre para no fijarlo a mano.
                sufijo = nombres[0][len(prefijos[0]):]
                esperadas = [p + sufijo for p in prefijos]
                assert nombres[:4] == esperadas, \
                    f"{nombre}: orden de tablas {nombres} != {esperadas}"
        assert not faltan, f"hojas base ausentes: {faltan}"

        # La ANE base debe traer el encabezado que busca _xl_col_cantidad.
        ane = wb[main.BASE_COMUNIDAD["ANE"]]
        for tabla in ane.tables.values():
            enc = _encabezado_tabla(ane, tabla.name)
            assert any(_norm(v) == "CANTIDAD A EJECUTAR" for v in enc), (tabla.name, enc)

        # El "Total en letras" debe seguir en la combinación A:G que mide
        # _xl_alto_letras (si cambia la combinación, el alto se calcula mal).
        combis = {str(r) for r in ane.merged_cells.ranges}
        assert "A18:G18" in combis, sorted(combis)

        # INICIO ya no existe: la exportación debe tolerarlo (tarea 29).
        assert "INICIO" not in wb.sheetnames, "la plantilla nueva trae INICIO"

        # VARIACIÓN: los 4 títulos de sección y la columna K renombrable.
        var = wb[main.BASE_COMUNIDAD["VARIACION"]]
        for tabla in var.tables.values():
            ncols = range_boundaries(tabla.ref)[2] - range_boundaries(tabla.ref)[0] + 1
            enc = _encabezado_tabla(var, tabla.name)
            assert any(_norm(v) == "CANTIDAD VARIADA" for v in enc), (tabla.name, enc)
            _cache.setdefault("ncols_var", ncols)
            assert ncols in (15, 16), f"{tabla.name}: {ncols} columnas"

        # OC/OTS/ATE: la referencia al nombre del proyecto está en B12.
        for hoja, nombre in main.BASE_COMUNIDAD.items():
            if hoja not in ("OC", "OTS", "ATE"):
                continue
            ws = wb[nombre]
            ref = ws["B12"].value
            assert isinstance(ref, str) and ref.startswith("=") and "ANE" in ref, \
                f"{nombre}!B12 debería apuntar al ANE: {ref!r}"
    finally:
        wb.close()


# ═════════════════════════════════════════════════════════════════════
#  3) lectura de la BD por la vía real
# ═════════════════════════════════════════════════════════════════════
def test_bd_por_cargar_bd_inicial():
    bd = main.cargar_bd_inicial()
    assert bd, "cargar_bd_inicial() devolvió una lista vacía"
    assert len(bd) == _cache["filas_bd"], (len(bd), _cache["filas_bd"])
    assert len(bd) > 100, len(bd)

    for c in bd:
        assert c["codigo"], c
        assert c["descripcion"], c
        assert main.D(c["precio"]) > 0, c
        for k in main.BD_COLS[4:]:
            assert c[k] in (0, 1), (k, c[k])
        assert not (c["v76"] and c["v199"]), c      # nunca ambos voltajes

    con_marca = [c for c in bd if any(c[k] for k in main.BD_COLS[4:])]
    assert con_marca, "ningún código quedó clasificado: la app no ofrecería materiales"

    # Los ejemplos concretos con los que se exporta.
    for c in con_marca[:3]:
        print("   ejemplo -> %s / %s / %s / %s" %
              (c["codigo"], c["descripcion"][:40], c["precio"],
               {m: c[m] for m in main.BD_COLS[4:]}))
    _bd_y_codigos()


# ═════════════════════════════════════════════════════════════════════
#  preparación del estado y exportación real
# ═════════════════════════════════════════════════════════════════════
def _bd_y_codigos():
    """BD real + los dos códigos de ejemplo (se lee una sola vez)."""
    if "bd" in _cache:
        return _cache["bd"]
    bd = main.cargar_bd_inicial()
    assert bd, "cargar_bd_inicial() devolvió una lista vacía"
    _cache["bd"] = bd
    _cache["filas_bd"] = len(bd)
    _cache["cod_mt"] = next(c["codigo"] for c in bd if c["MT"] == 1 and c["BT"] == 0)
    _cache["cod_bt"] = next(c["codigo"] for c in bd if c["BT"] == 1)
    for k in ("cod_mt", "cod_bt"):
        c = next(x for x in bd if x["codigo"] == _cache[k])
        print("   %s -> %s / %s / precio %s / marcas %s" %
              (k, c["codigo"], c["descripcion"][:40], c["precio"],
               {m: c[m] for m in main.BD_COLS[4:]}))
    return bd


def _estado():
    if "estado" in _cache:
        return _cache["estado"]
    bd = [dict(c) for c in _bd_y_codigos()]
    nuevo = main.componente_nuevo(COD_NUEVO, DESC_NUEVA, "Unidad", PRECIO_NUEVO)
    nuevo["MT"] = 1
    nuevo["v76"] = 1
    bd.append(nuevo)

    e = main.Estado.__new__(main.Estado)
    e.bd = bd
    e.reset()
    b = e.proyectos[0]
    b["abrev"] = ABREV
    b["nombre"] = "COMUNIDAD PRUEBA 2025"
    b["hojas"] = {h: (h in HOJAS_PROYECTO) for h in main.HOJAS}
    b["visibles"] = ["ANE"]
    b["items"] = {h: main._vacio() for h in main.EDITABLES}
    b["items"]["ANE"]["MT"] = [
        {"codigo": _cache["cod_mt"], "cantidad": "2", "redondeo": "ROUND"},
        {"codigo": COD_NUEVO, "cantidad": "1", "redondeo": "ROUNDUP"},
    ]
    b["items"]["ANE_VAR"]["MT"] = [
        {"codigo": _cache["cod_mt"], "cantidad": "3", "redondeo": "ROUNDDOWN"},
    ]
    b["items"]["ANE_VAR"]["BT"] = [
        {"codigo": _cache["cod_bt"], "cantidad": "1", "redondeo": "ROUND"},
    ]
    e.bd_local = True
    e.cambios = True
    _cache["estado"] = e
    return e


def _exportado():
    if "ruta" not in _cache:
        os.makedirs(SALIDA, exist_ok=True)
        ruta = os.path.join(SALIDA, f"ANEXO - {ABREV} - plantilla2025.xlsx")
        if os.path.exists(ruta):
            os.remove(ruta)
        e = _estado()
        _cache["totales"] = main.exportar_xlsx(e, ruta)
        _cache["ruta"] = ruta
    return _cache["ruta"]


@pytest.mark.excel
def test_exportacion_real_con_plantilla_nueva():
    ruta = _exportado()
    e = _cache["estado"]
    assert os.path.exists(ruta), ruta
    assert os.path.getsize(ruta) > 20000

    wbv = load_workbook(ruta, data_only=True)
    wbf = load_workbook(ruta, data_only=False)
    try:
        # · hojas del proyecto creadas, bases eliminadas
        assert main.HOJA_BD in wbv.sheetnames, wbv.sheetnames
        assert not [n for n in wbv.sheetnames if n.startswith("_BASE")], wbv.sheetnames
        for hoja in HOJAS_PROYECTO:
            assert e.titulo(e.proyectos[0], hoja) in wbv.sheetnames, \
                (hoja, wbv.sheetnames)

        # · sin errores de fórmula en todo el libro
        malos = _errores_libro(wbv)
        assert not malos, f"celdas con error: {malos[:10]}"

        ws_ane = wbv[e.titulo(e.proyectos[0], "ANE")]
        # · la BD exportada quedó NUMÉRICA en el precio (la plantilla la trae texto)
        ws_bd = wbv[main.HOJA_BD]
        precios = [ws_bd.cell(row=f, column=4).value
                   for f in range(2, min(ws_bd.max_row, 20) + 1)]
        assert all(isinstance(p, (int, float)) for p in precios), precios[:5]
        codigos_bd = [str(ws_bd.cell(row=f, column=1).value).strip()
                      for f in range(2, ws_bd.max_row + 1)
                      if ws_bd.cell(row=f, column=1).value not in (None, "")]
        # La BD exportada = catálogo de la plantilla + el código nuevo (Local).
        assert len(codigos_bd) == len(e.bd), (len(codigos_bd), len(e.bd))
        assert COD_NUEVO in codigos_bd and codigos_bd.count(COD_NUEVO) == 1

        # · BUSCARX resueltas y totales iguales a los de la app
        esperado = e.filas(e.proyectos[0], "ANE")
        filas_excel = {}
        for linea, _, _ in main.LINEAS:
            nombre = main._xl_nombre_tabla("ANE", main._xl_token(ABREV), linea)
            for fila in _filas_tabla(ws_ane, nombre):
                filas_excel[str(fila[0]).strip()] = (linea, fila)
        for linea, _, _ in main.LINEAS:
            for r in esperado[linea]:
                clave = str(r["codigo"]).strip()
                assert clave in filas_excel, f"{r['codigo']} no está en ANE/{linea}"
                _, fila = filas_excel[clave]
                assert fila[2] == r["desc"], (clave, fila[2], r["desc"])
                assert fila[3] == r["unidad"], (clave, fila[3], r["unidad"])
                assert float(fila[4]) == float(r["precio"]), (clave, fila[4], r["precio"])
                assert float(fila[5]) == float(r["cantidad"]), (clave, fila[5])
                assert float(fila[6]) == float(r["total"]), \
                    (clave, r["redondeo"], fila[6], r["total"])
        # · fórmula del precio total (conserva el redondeo por celda)
        form = wbf[e.titulo(e.proyectos[0], "ANE")]
        f7 = [c.value for c in form["G"] if isinstance(c.value, str) and c.value.startswith("=")]
        assert any("ROUNDUP" in x.upper() for x in f7), f7[:3]

        # · formato condicional del código numérico (>2000) aplicado
        reglas = []
        for cf in ws_ane.conditional_formatting:
            for r in cf.rules:
                reglas.append((str(cf.sqref), getattr(r, "formula", None), r.type))
        assert reglas, "ANE no tiene formato condicional"
        assert any(x[1] and any(str(main.UMBRAL_CODIGO_ATE) in f for f in x[1])
                   for x in reglas), reglas[:5]

        # · total de la hoja ANE = suma de las celdas redondeadas
        f_total = _fila_total(ws_ane)
        assert f_total, "no se encontró la fila TOTAL en ANE"
        total_excel = float(ws_ane.cell(row=f_total, column=7).value)
        total_app = float(main.Estado.total(esperado))
        assert abs(total_excel - total_app) < 0.005, (total_excel, total_app)
        _cache["total_ane"] = total_excel

        # · VARIACIÓN: se quitó CANT TOTAL y el Tipo de Documento es correcto
        ws_var = wbv[e.titulo(e.proyectos[0], "VARIACION")]
        for tabla in ws_var.tables.values():
            c1, _, c2, _ = range_boundaries(tabla.ref)
            ncols = c2 - c1 + 1
            assert ncols == 15, \
                f"{tabla.name} tiene {ncols} columnas (CANT TOTAL no se quitó)"
        tipos = {}
        for tabla in ws_var.tables.values():
            for fila in _filas_tabla(ws_var, tabla.name):
                tipos[str(fila[0]).strip()] = fila[9]
        assert tipos.get(COD_NUEVO) == "ATE", tipos
        assert tipos.get(_cache["cod_mt"]) == "OTS", tipos   # variada +1
        _cache["tipos"] = tipos

        # · OTS recibe la fila de cod_mt; ATE la del código >2000
        ws_ots = wbv[e.titulo(e.proyectos[0], "OTS")]
        cods_ots = [str(f[0]).strip() for t in ws_ots.tables.values()
                    for f in _filas_tabla(ws_ots, t.name)]
        assert _cache["cod_mt"] in cods_ots, cods_ots
        ws_ate = wbv[e.titulo(e.proyectos[0], "ATE")]
        cods_ate = [str(f[0]).strip() for t in ws_ate.tables.values()
                    for f in _filas_tabla(ws_ate, t.name)]
        assert COD_NUEVO in cods_ate, cods_ate
    finally:
        wbv.close()
        wbf.close()


@pytest.mark.excel
def test_roundtrip_recupera_proyecto_codigos_y_redondeos():
    ruta = _exportado()
    e2 = main.Estado.__new__(main.Estado)
    e2.bd = []
    e2.reset()
    main.cargar_proyecto_desde_excel(e2, ruta)

    assert len(e2.proyectos) == 1, [p["abrev"] for p in e2.proyectos]
    b = e2.proyectos[0]
    assert b["abrev"] == ABREV, b["abrev"]
    assert b["nombre"] == "COMUNIDAD PRUEBA 2025", b["nombre"]
    for hoja in HOJAS_PROYECTO:
        assert b["hojas"][hoja], (hoja, b["hojas"])

    bd = {c["codigo"]: c for c in e2.bd}
    assert COD_NUEVO in bd, "el código nuevo no volvió de la BD del archivo"
    assert bd[COD_NUEVO]["descripcion"] == DESC_NUEVA
    assert float(main.D(bd[COD_NUEVO]["precio"])) == 1000.0
    assert len(e2.bd) == len(_cache["estado"].bd), (len(e2.bd), len(_cache["estado"].bd))

    ane = b["items"]["ANE"]["MT"]
    por_cod = {it["codigo"]: it for it in ane}
    assert _cache["cod_mt"] in por_cod, ane
    # COM devuelve las cantidades como número (2.0): se comparan como Decimal.
    assert main.D(por_cod[_cache["cod_mt"]]["cantidad"]) == main.D("2"), \
        por_cod[_cache["cod_mt"]]
    assert por_cod[_cache["cod_mt"]]["redondeo"] == "ROUND", por_cod[_cache["cod_mt"]]
    assert COD_NUEVO in por_cod and por_cod[COD_NUEVO]["redondeo"] == "ROUNDUP", ane

    av = {(it["codigo"], linea[0]): it for linea in main.LINEAS
          for it in b["items"]["ANE_VAR"][linea[0]]}
    assert av[(_cache["cod_mt"], "MT")]["redondeo"] == "ROUNDDOWN", av
    assert main.D(av[(_cache["cod_bt"], "BT")]["cantidad"]) == main.D("1"), av


# ═════════════════════════════════════════════════════════════════════
#  la plantilla REAL debe quedar intacta
# ═════════════════════════════════════════════════════════════════════
def test_plantilla_real_intacta():
    assert _sha256(main.PLANTILLA_XLSX) == _cache["hash_previo"], \
        "la plantilla real cambió de contenido"
    st = os.stat(main.PLANTILLA_XLSX)
    assert st.st_size == _cache["stat_previo"].st_size
    assert int(st.st_mtime) == int(_cache["stat_previo"].st_mtime)
