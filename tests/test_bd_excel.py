# -*- coding: utf-8 -*-
"""
Pruebas del cambio "la BD viaja dentro del .xlsx" (guardado Local/Global).

Requieren Excel instalado (xlwings/COM) → marcadas con @pytest.mark.excel.

Nunca escriben la plantilla real:
  · La exportación copia la plantilla a un temporal (así lo hace main.py).
  · El guardado Global se prueba sobre una COPIA del .xlsm, parcheando
    main.PLANTILLA_XLSX y anulando `_respaldo_plantilla` para no ensuciar
    backups/.
"""

import os
import shutil
import sys
import tempfile
import types

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RAIZ not in sys.path:
    sys.path.insert(0, RAIZ)

_falso = types.ModuleType("license_manager")
_falso.check_license = lambda: True
sys.modules.setdefault("license_manager", _falso)

import atexit                                  # noqa: E402

import pytest                                   # noqa: E402
import main                                     # noqa: E402


def _limpiar_temporales():
    """Borra las carpetas anexos_* que estas pruebas dejan en %TEMP%."""
    base = tempfile.gettempdir()
    try:
        nombres = os.listdir(base)
    except OSError:
        return
    for nombre in nombres:
        if nombre.startswith("anexos_"):
            shutil.rmtree(os.path.join(base, nombre), ignore_errors=True)


atexit.register(_limpiar_temporales)

_EstadoReal = main.Estado
COD_NUEVO = "999999"
DESC_NUEVA = "EQUIPO NUEVO DE PRUEBA 13.2 kV"
PRECIO_NUEVO = "100.00"

_cache = {}


# ── utilidades ───────────────────────────────────────────────────────
def _bd_base():
    """Catálogo real de la plantilla (solo lectura) + el código nuevo.

    Es el estado que tiene la app de verdad: `Estado.bd` llega con TODO el
    catálogo de la plantilla y encima se agregan los códigos nuevos.
    """
    if "bd" not in _cache:
        bd = main.cargar_bd_inicial()
        assert len(bd) > 100, f"catálogo leído: {len(bd)}"
        nuevo = main.componente_nuevo(COD_NUEVO, DESC_NUEVA, "U", PRECIO_NUEVO)
        nuevo["MT"] = 1
        nuevo["v76"] = 1
        bd.append(nuevo)
        _cache["bd"] = bd
    return _cache["bd"]


def _estado_local():
    """Estado mínimo: 1 proyecto COMUNIDAD_1 + código nuevo 999999 en MT."""
    e = _EstadoReal.__new__(_EstadoReal)
    e.bd = [dict(c) for c in _bd_base()]
    e.reset()
    b = e.proyectos[0]
    b["abrev"] = "COMUNIDAD_1"
    b["nombre"] = "COMUNIDAD DE PRUEBA"
    b["hojas"] = {h: (h == "ANE") for h in main.HOJAS}
    b["visibles"] = ["ANE"]

    b["items"] = {"ANE": main._vacio()}
    b["items"]["ANE"]["MT"] = [{"codigo": COD_NUEVO, "cantidad": "2",
                                "redondeo": "ROUND"}]
    e.bd_local = True
    e.cambios = True
    return e


def _exportar(destino):
    e = _estado_local()
    totales = main.exportar_xlsx(e, destino)
    return e, totales


def _exportado():
    """Exporta una sola vez y reutiliza el archivo en las demás pruebas."""
    if "ruta" not in _cache:
        carpeta = tempfile.mkdtemp(prefix="anexos_export_")
        ruta = os.path.join(carpeta, "ANEXO - COMUNIDAD_1 - 21.10.2026.xlsx")
        _cache["estado"], _cache["totales"] = _exportar(ruta)
        _cache["ruta"] = ruta
        _cache["carpeta"] = carpeta
    return _cache["ruta"]


def _filas_bd(ws):
    return [r for r in ws.iter_rows(min_row=2, values_only=True)
            if r and r[0] not in (None, "")]


def _buscar_codigo(ws, codigo):
    for fila in ws.iter_rows(min_row=1):
        for celda in fila:
            v = celda.value
            if v is not None and str(v).strip() == str(codigo):
                return celda.row, celda.column
    return None, None


def _errores_en_libro(wb):
    """Celdas con error de fórmula visibles en el libro."""
    males = ("#N/A", "#REF!", "#VALUE!", "#NAME?", "#DIV/0!")
    hallados = []
    for ws in wb.worksheets:
        for fila in ws.iter_rows():
            for celda in fila:
                v = celda.value
                if isinstance(v, str) and v in males:
                    hallados.append((ws.title, celda.coordinate, v))
    return hallados


# ═════════════════════════════════════════════════════════════════════
#  3) exportar_xlsx escribe CODIGOS_SICOA con el código nuevo
# ═════════════════════════════════════════════════════════════════════
@pytest.mark.excel
def test_exportar_incluye_bd_local_y_buscarx_resuelve():
    from openpyxl import load_workbook

    ruta = _exportado()
    assert os.path.exists(ruta)

    wbv = load_workbook(ruta, data_only=True)
    wbf = load_workbook(ruta, data_only=False)

    # · el .xlsx conserva CODIGOS y no arrastra las bases internas
    assert main.HOJA_BD in wbv.sheetnames, wbv.sheetnames
    bases = [n for n in wbv.sheetnames if n.startswith("_BASE")]
    assert not bases, f"quedaron hojas base: {bases}"

    # · la tabla CODIGOS_SICOA tiene 10 columnas y contiene el código nuevo
    ws = wbv[main.HOJA_BD]
    assert main.TABLA_BD in [str(t).upper() for t in ws.tables.keys()], list(ws.tables.keys())
    encabezado = [c.value for c in ws[1] if c.value not in (None, "")]
    assert len(encabezado) == 10, encabezado

    filas = _filas_bd(ws)
    codigos = [str(r[0]).strip() for r in filas]
    assert COD_NUEVO in codigos, "el código nuevo no llegó a CODIGOS_SICOA"
    assert codigos.count(COD_NUEVO) == 1
    # · no se perdió el catálogo de la plantilla
    assert len(filas) == len(_bd_base()), f"filas: {len(filas)} vs {len(_bd_base())}"
    primero = str(_bd_base()[0]["codigo"]).strip()
    assert primero in codigos, f"desapareció el código {primero}"

    fila = filas[codigos.index(COD_NUEVO)]
    # · el código se guarda como NÚMERO (lo exige el formato condicional/ATE)
    assert isinstance(fila[0], (int, float)), f"tipo de celda: {type(fila[0])} ({fila[0]!r})"
    assert int(fila[0]) == int(COD_NUEVO)
    assert fila[1] == DESC_NUEVA, fila[1]
    assert fila[2] == "U", fila[2]
    assert float(fila[3]) == 100.0, fila[3]
    # · las seis marcas: MT=1, v76=1, el resto 0
    marcas = dict(zip(("MT", "MTBT", "BT", "G", "v76", "v199"), fila[4:10]))
    assert marcas["MT"] == 1 and marcas["v76"] == 1, marcas
    for k in ("MTBT", "BT", "G", "v199"):
        assert marcas[k] == 0, marcas

    # · las BUSCARX de ANE resuelven descripción/unidad/precio del código nuevo
    titulo_ane = main.PREFIJOS_HOJA["ANE"] + "COMUNIDAD_1"
    assert titulo_ane in wbv.sheetnames, wbv.sheetnames
    ane = wbv[titulo_ane]
    fila_ane, _ = _buscar_codigo(ane, COD_NUEVO)
    assert fila_ane, "el código no está en la hoja ANE"
    desc = ane.cell(row=fila_ane, column=3).value
    uni = ane.cell(row=fila_ane, column=4).value
    precio = ane.cell(row=fila_ane, column=5).value
    total = ane.cell(row=fila_ane, column=7).value
    assert isinstance(desc, str) and desc == DESC_NUEVA, f"desc={desc!r}"
    assert uni == "U", f"uni={uni!r}"
    assert float(precio) == 100.0, f"precio={precio!r}"
    assert float(total) == 200.0, f"total={total!r}"

    # · la celda de ANE sigue siendo una fórmula XLOOKUP (no un valor pegado)
    formula = wbf[titulo_ane].cell(row=fila_ane, column=3).value
    assert isinstance(formula, str) and "XLOOKUP" in formula.upper(), formula

    # · sin errores de fórmula en ningún lado
    malos = _errores_en_libro(wbv)
    assert not malos, f"celdas con error: {malos[:10]}"


# ═════════════════════════════════════════════════════════════════════
#  3b) idempotencia: exportar dos veces no duplica el código
# ═════════════════════════════════════════════════════════════════════
@pytest.mark.excel
def test_exportar_dos_veces_no_duplica():
    from openpyxl import load_workbook

    primera = _exportado()
    carpeta = tempfile.mkdtemp(prefix="anexos_export2_")
    segunda = os.path.join(carpeta, "ANEXO - COMUNIDAD_1 - copia.xlsx")
    try:
        _exportar(segunda)
        a = _filas_bd(load_workbook(primera, data_only=True)[main.HOJA_BD])
        b = _filas_bd(load_workbook(segunda, data_only=True)[main.HOJA_BD])
        assert len(b) == len(a), f"filas: {len(a)} → {len(b)}"
        assert [str(r[0]).strip() for r in b].count(COD_NUEVO) == 1
    finally:
        shutil.rmtree(carpeta, ignore_errors=True)


# ═════════════════════════════════════════════════════════════════════
#  4) cargar_proyecto_desde_excel lee la BD del propio .xlsx
# ═════════════════════════════════════════════════════════════════════
@pytest.mark.excel
def test_reabrir_lee_la_bd_del_archivo():
    ruta = _exportado()
    e = _EstadoReal.__new__(_EstadoReal)
    e.bd = []
    e.reset()

    main.cargar_proyecto_desde_excel(e, ruta)

    c = e.comp(COD_NUEVO)
    assert c is not None, "al reabrir se perdió el código nuevo"
    assert c["descripcion"] == DESC_NUEVA, c
    assert float(c["precio"]) == 100.0, c["precio"]
    assert c["MT"] == 1, c
    assert c["v76"] == 1, c
    for k in ("MTBT", "BT", "G", "v199"):
        assert c[k] == 0, c
    # la BD del archivo manda: trae TODOS los códigos, no solo el nuevo
    assert len(e.bd) > 100, len(e.bd)


@pytest.mark.excel
def test_sin_la_llamada_a_escribir_bd_el_codigo_no_llega():
    """Mutación: si `exportar_xlsx` dejara de llamar a `_xl_escribir_bd`,
    el código nuevo NO aparecería en CODIGOS. Confirma que la aserción
    principal discrimina (no pasa por casualidad)."""
    from openpyxl import load_workbook

    original = main._xl_escribir_bd
    carpeta = tempfile.mkdtemp(prefix="anexos_mutacion_")
    ruta = os.path.join(carpeta, "sin_escribir_bd.xlsx")
    main._xl_escribir_bd = lambda wb, bd: None
    try:
        _exportar(ruta)
        ws = load_workbook(ruta, data_only=True)[main.HOJA_BD]
        codigos = [str(r[0]).strip() for r in _filas_bd(ws)]
        assert COD_NUEVO not in codigos, \
            "se escribió la BD aunque _xl_escribir_bd era un no-op"
    finally:
        main._xl_escribir_bd = original
        shutil.rmtree(carpeta, ignore_errors=True)


# ═════════════════════════════════════════════════════════════════════
#  4b) mutación: si no se lee la BD del archivo, se pierde el código
# ═════════════════════════════════════════════════════════════════════
@pytest.mark.excel
def test_reabrir_sin_leer_bd_del_archivo_pierde_el_codigo():
    ruta = _exportado()
    e = _EstadoReal.__new__(_EstadoReal)
    e.bd = []
    e.reset()

    original = main._xl_leer_bd
    main._xl_leer_bd = lambda wb: []          # fuerza caer a la plantilla
    try:
        main.cargar_proyecto_desde_excel(e, ruta)
    finally:
        main._xl_leer_bd = original

    assert e.comp(COD_NUEVO) is None, \
        "el código nuevo seguía apareciendo sin leer la BD del archivo"


# ═════════════════════════════════════════════════════════════════════
#  3c) BD vacía: no debe vaciar CODIGOS (debe salir temprano)
# ═════════════════════════════════════════════════════════════════════
@pytest.mark.excel
def test_escribir_bd_vacia_no_vacia_codigos():
    carpeta = tempfile.mkdtemp(prefix="anexos_bdvacia_")
    copia = os.path.join(carpeta, "copia.xlsm")
    shutil.copyfile(main.PLANTILLA_XLSX, copia)

    app, pythoncom = main._abrir_excel()
    wb = None
    try:
        wb = app.books.open(copia)
        ws = wb.sheets[main.HOJA_BD]
        tabla = main._xl_tabla(ws, main.TABLA_BD)
        antes = int(tabla.api.DataBodyRange.Rows.Count)
        assert antes > 100, antes

        main._xl_escribir_bd(wb, [])

        despues = int(tabla.api.DataBodyRange.Rows.Count)
        assert despues == antes, f"filas CODIGOS: {antes} → {despues}"

        # Control de mutación: con datos SÍ reescribe (y recorta) la tabla, así
        # que el caso vacío realmente dependió del `return` temprano.
        unico = main.componente_nuevo("999999", "X", "U", "1")
        unico["MT"] = 1
        main._xl_escribir_bd(wb, [unico])
        recortada = int(tabla.api.DataBodyRange.Rows.Count)
        assert recortada == 1, f"con 1 código quedaron {recortada} filas"
    finally:
        try:
            if wb is not None:
                wb.close()
        finally:
            main._cerrar_excel(app, pythoncom)
            shutil.rmtree(carpeta, ignore_errors=True)


# ═════════════════════════════════════════════════════════════════════
#  Global: bd_guardar_plantilla sobre una COPIA temporal
# ═════════════════════════════════════════════════════════════════════
@pytest.mark.excel
def test_guardado_global_sobre_copia_no_toca_la_plantilla():
    from openpyxl import load_workbook

    original = main.PLANTILLA_XLSX
    assert os.path.exists(original), original
    est_antes = os.stat(original)

    carpeta = tempfile.mkdtemp(prefix="anexos_global_")
    copia = os.path.join(carpeta, "plantilla_copia.xlsm")
    shutil.copyfile(original, copia)

    try:
        bd = main.cargar_bd_inicial()
        assert len(bd) > 100, f"BD leída: {len(bd)}"
        n_antes = len(bd)
        assert not any(c["codigo"] == COD_NUEVO for c in bd)

        nuevo = main.componente_nuevo(COD_NUEVO, "EQUIPO GLOBAL DE PRUEBA", "U", "77.00")
        nuevo["BT"] = 1
        bd.append(nuevo)

        respaldo_real = main._respaldo_plantilla
        main.PLANTILLA_XLSX = copia
        main._respaldo_plantilla = lambda: None      # no ensuciar backups/
        try:
            main.bd_guardar_plantilla(bd)
        finally:
            main.PLANTILLA_XLSX = original
            main._respaldo_plantilla = respaldo_real

        ws = load_workbook(copia, data_only=True)[main.HOJA_BD]
        filas = _filas_bd(ws)
        codigos = [str(r[0]).strip() for r in filas]
        assert COD_NUEVO in codigos, "el código no se guardó en la copia"
        assert len(filas) == n_antes + 1, f"filas: {n_antes} → {len(filas)}"
        assert len([c for c in codigos if c]) == len(set(codigos)), "hay duplicados"
        encabezado = [c.value for c in ws[1] if c.value not in (None, "")]
        assert len(encabezado) == 10, encabezado

        fila = filas[codigos.index(COD_NUEVO)]
        marcas = dict(zip(("MT", "MTBT", "BT", "G", "v76", "v199"), fila[4:10]))
        assert marcas["BT"] == 1, marcas
        for k in ("MT", "MTBT", "G", "v76", "v199"):
            assert marcas[k] == 0, marcas

        # · la plantilla REAL no se tocó
        est_despues = os.stat(original)
        assert est_despues.st_size == est_antes.st_size, "cambió el tamaño de la plantilla"
        assert int(est_despues.st_mtime) == int(est_antes.st_mtime), "cambió la plantilla"
    finally:
        shutil.rmtree(carpeta, ignore_errors=True)
