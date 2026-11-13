# -*- coding: utf-8 -*-
"""
T1 + T2 del Ramal — ROUNDTRIP REAL con Excel/COM (xlwings).

Casos:
  · Caso 1: Ramal SIN filas propias → el .xlsx debe salir como antes
            (SUMIF puro, sin "+literal").
  · Caso 2: Ramal CON una fila propia en ANE (código que no existe en ningún
            proyecto) y VARIACIÓN activada SOLO en el Ramal.
  · Caso 3: compatibilidad — al mismo libro del Caso 2 se le quita el
            "+literal" de la fórmula (como lo escribía la versión anterior) y
            se reimporta: no debe inventar filas propias ni filas fantasma.

Nunca se escribe sobre la plantilla real: `exportar_xlsx` copia el machote a un
temporal y todos los .xlsx de salida viven en la carpeta TEMP del proyecto
(C:\\Users\\deivi\\AppData\\Local\\Temp\\opencode\\ramal_t1t2).

Uso:
  · `python -m pytest tests/test_ramal_propio_excel.py -q -m excel`
  · `python tests/test_ramal_propio_excel.py`   → runner propio (sin pytest)
"""

import atexit
import hashlib
import os
import re
import shutil
import subprocess
import sys
import tempfile
import traceback
import types

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RAIZ not in sys.path:
    sys.path.insert(0, RAIZ)

try:                                            # consola Windows = cp1252
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

_falso = types.ModuleType("license_manager")
_falso.check_license = lambda: True
sys.modules.setdefault("license_manager", _falso)

try:                                            # pytest es opcional
    import pytest
    excel = pytest.mark.excel
except ImportError:                             # pragma: no cover
    pytest = None

    def excel(func):
        return func

from decimal import Decimal                     # noqa: E402

import main                                     # noqa: E402


# ═════════════════════════════════════════════════════════════════════
#  CARPETA DE TRABAJO (nunca la raíz del proyecto)
# ═════════════════════════════════════════════════════════════════════
CARPETA = os.path.join(
    os.environ.get("TEMP") or tempfile.gettempdir(),
    "opencode", "ramal_t1t2")
if not os.path.isdir(CARPETA):
    try:
        os.makedirs(CARPETA)
    except OSError:                              # pragma: no cover
        CARPETA = tempfile.mkdtemp(prefix="anexos_ramal_")


def _limpiar():
    """Borra solo los .xlsx que creó esta prueba (no la carpeta del usuario)."""
    try:
        for nombre in os.listdir(CARPETA):
            if nombre.startswith(("t1t2_", "~$t1t2_")):
                try:
                    os.remove(os.path.join(CARPETA, nombre))
                except OSError:
                    pass
    except OSError:
        pass


atexit.register(_limpiar)


# ═════════════════════════════════════════════════════════════════════
#  ESTADO DE PRUEBA (BD inyectada; no se lee la plantilla)
# ═════════════════════════════════════════════════════════════════════
def comp(codigo, precio="10.00", unidad="U", lineas=("MT",)):
    c = main.componente_nuevo(codigo, f"DESCRIPCION {codigo}", unidad, precio)
    for k, _, _ in main.LINEAS:
        c[k] = 1 if k in lineas else 0
    return c


BD_PRUEBA = [comp("P001"), comp("P002"), comp("K001", "100.00", "Km"),
             comp("P999", "5.00")]


def estado_nuevo(bd=None):
    e = main.Estado.__new__(main.Estado)
    e.bd = [dict(c) for c in (bd if bd is not None else BD_PRUEBA)]
    e.reset()
    return e


def estado_vacio():
    e = main.Estado.__new__(main.Estado)
    e.bd = []
    e.reset()
    return e


def agregar(base, hoja, linea, codigo, cantidad, redondeo="ROUND"):
    base["items"].setdefault(hoja, main._vacio())[linea].append(
        {"codigo": codigo, "cantidad": str(cantidad), "redondeo": redondeo})


def totales_ramal(e, hoja):
    return {k: main.Estado.subtotales(e.filas(e.ramal, hoja))[k]
            for k, _, _ in main.LINEAS}


# ═════════════════════════════════════════════════════════════════════
#  AYUDAS DE EXCEL
# ═════════════════════════════════════════════════════════════════════
def con_libro(ruta, fn):
    """Abre el .xlsx, ejecuta `fn(wb)` y cierra SIEMPRE el libro y el COM."""
    app, pythoncom = main._abrir_excel()
    wb = None
    try:
        wb = app.books.open(os.path.abspath(ruta))
        app.calculate()
        return fn(wb, app)
    finally:
        try:
            if wb is not None:
                wb.close()
        finally:
            main._cerrar_excel(app, pythoncom)


def formulas_cantidad(wb, abrev, hoja, linea):
    """Texto de las fórmulas de 'Cantidad a ejecutar' de una tabla del Ramal."""
    token = main._xl_token(abrev)
    ws = wb.sheets[main.PREFIJOS_HOJA[hoja] + token]
    tabla = main._xl_tabla(ws, main._xl_nombre_tabla(hoja, token, linea))
    if tabla is None or tabla.api.DataBodyRange is None:
        return []
    col = main._xl_col_cantidad(tabla)
    cuerpo = tabla.api.ListColumns(col).DataBodyRange
    return [str(f[0] or "") for f in main._xl_valores(cuerpo.Formula)]


def celdas_con_error(wb, abrev, hoja):
    """Celdas con #N/A, #REF!, #VALUE!, #NAME? o #DIV/0! en la hoja indicada."""
    token = main._xl_token(abrev)
    ws = wb.sheets[main.PREFIJOS_HOJA[hoja] + token]
    malas = []
    for fila in main._xl_valores(ws.used_range.value):
        for v in (fila if isinstance(fila, (list, tuple)) else [fila]):
            if isinstance(v, str) and v.startswith("#"):
                malas.append(v)
    return malas


def _procesos_excel():
    out = subprocess.run(
        ["tasklist", "/FO", "CSV", "/NH", "/FI", "IMAGENAME eq EXCEL.EXE"],
        capture_output=True, text=True, timeout=60,
    )
    filas = []
    for linea in out.stdout.splitlines():
        linea = linea.strip()
        if not linea or "INFO:" in linea.upper():
            continue
        campos = [c.strip('"') for c in linea.split('","')]
        if campos and campos[0].upper().startswith("EXCEL.EXE"):
            filas.append(campos)
    return filas


def _huella_plantilla():
    ruta = main.PLANTILLA_XLSX
    if not os.path.exists(ruta):
        return None
    h = hashlib.sha256()
    with open(ruta, "rb") as f:
        for trozo in iter(lambda: f.read(1 << 20), b""):
            h.update(trozo)
    st = os.stat(ruta)
    return (h.hexdigest(), st.st_size)


EXCEL_ANTES = None
PLANTILLA_ANTES = None


def _recordar_estado_inicial():
    global EXCEL_ANTES, PLANTILLA_ANTES
    if EXCEL_ANTES is None:
        try:
            EXCEL_ANTES = len(_procesos_excel())
        except Exception:
            EXCEL_ANTES = 0
    if PLANTILLA_ANTES is None:
        try:
            PLANTILLA_ANTES = _huella_plantilla()
        except Exception:
            PLANTILLA_ANTES = None


_recordar_estado_inicial()


# ═════════════════════════════════════════════════════════════════════
#  CASO 1 · Ramal SIN filas propias
# ═════════════════════════════════════════════════════════════════════
@excel
def test_caso1_ramal_sin_filas_propias():
    if pytest is not None:
        pytest.importorskip("xlwings")

    e = estado_nuevo()
    p0 = e.proyectos[0]
    agregar(p0, "ANE", "MT", "P001", 3)
    e.activar_hoja(p0, "ANE_VAR")
    agregar(p0, "ANE_VAR", "MT", "P002", 2)
    p1 = e.agregar_proyecto()
    # Cada proyecto con al menos una fila: el SUMIF de su tabla no queda sobre
    # una tabla vacía (situación que no es objeto de esta prueba).
    agregar(p1, "ANE", "MT", "K001", 2)

    assert e.ramal_visible() is True
    totales = totales_ramal(e, "ANE")

    ruta = os.path.join(CARPETA, "t1t2_caso1.xlsx")
    if os.path.exists(ruta):
        os.remove(ruta)
    main.exportar_xlsx(e, ruta)
    assert os.path.exists(ruta), "la exportación no generó el archivo"

    # 1) La fórmula del Ramal sigue siendo un SUMIF puro (sin "+literal").
    formulas = con_libro(ruta, lambda wb, app: formulas_cantidad(
        wb, e.ramal["abrev"], "ANE", "MT"))
    assert formulas, "no se encontraron filas en ANE - MT del Ramal"
    for f in formulas:
        assert "SUMIF" in f.upper(), f"fórmula sin SUMIF: {f!r}"
        assert not re.search(r"\)\s*\+\s*[\d.]", f), \
            f"sin filas propias no debe haber un literal sumado: {f!r}"

    # 2) Sin celdas con error.
    malas = con_libro(ruta, lambda wb, app: celdas_con_error(
        wb, e.ramal["abrev"], "ANE"))
    assert not malas, f"celdas con error en el Ramal: {malas}"

    # 3) La firma SUMIF sigue detectando el Ramal.
    def firmas(wb, app):
        grupos = main._xl_grupos_desde_wb(wb)
        return [(g["token"], main._xl_firma_ramal(g)) for g in grupos]
    pares = con_libro(ruta, firmas)
    ramal_token = main._xl_token(e.ramal["abrev"]).upper()
    firma = {main._xl_token(t).upper(): v for t, v in pares}.get(ramal_token)
    assert firma == 1000, f"firma del Ramal {firma} (esperada 1000) en {pares}"

    # 4) Roundtrip: nada propio, ninguna fila duplicada.
    e2 = estado_vacio()
    main.cargar_proyecto_desde_excel(e2, ruta)
    assert e2.ramal_presente is True
    for hoja in main.EDITABLES:
        for k, _, _ in main.LINEAS:
            # `_xl_leer_items` deja filas propias en 0 cuando la resta da 0;
            # lo que importa es que ninguna quede en positivo.
            positivas = {c: q for c, q in e2._cant_propia(hoja)[k].items() if q > 0}
            assert not positivas, \
                f"Caso 1 no debe tener filas propias en {hoja}:{k} → {positivas}"
    filas = e2.filas(e2.ramal, "ANE")["MT"]
    codigos = [f["codigo"] for f in filas]
    assert len(codigos) == len(set(codigos)), f"filas duplicadas: {codigos}"
    por_cod = {f["codigo"]: f for f in filas}
    assert por_cod["P001"]["cantidad"] == Decimal(3)
    assert totales_ramal(e2, "ANE") == totales, "los totales cambiaron al reimportar"


# ═════════════════════════════════════════════════════════════════════
#  CASO 2 · Ramal CON una fila propia y VARIACIÓN solo en el Ramal
# ═════════════════════════════════════════════════════════════════════
@excel
def test_caso2_fila_propia_y_variacion_solo_en_el_ramal():
    if pytest is not None:
        pytest.importorskip("xlwings")

    e = estado_nuevo()
    p0 = e.proyectos[0]
    agregar(p0, "ANE", "MT", "P001", 3)
    e.activar_hoja(p0, "ANE_VAR")
    agregar(p0, "ANE_VAR", "MT", "P002", 2)
    p1 = e.agregar_proyecto()
    agregar(p1, "ANE", "MT", "K001", 2)
    # Ningún proyecto tiene VARIACIÓN; el Ramal la activa por su cuenta (T1).
    e.activar_hoja(e.ramal, "VARIACION")
    # Fila propia: P999 no existe en ningún proyecto.
    agregar(e.ramal, "ANE", "MT", "P999", 5)
    assert e._cant_propia("ANE")["MT"].get("P999", Decimal(0)) > 0
    assert all(not p["hojas"]["VARIACION"] for p in e.proyectos)

    totales = totales_ramal(e, "ANE")
    ruta = os.path.join(CARPETA, "t1t2_caso2.xlsx")
    if os.path.exists(ruta):
        os.remove(ruta)
    main.exportar_xlsx(e, ruta)
    assert os.path.exists(ruta)

    # 1) La fórmula del Ramal suma el literal de la fila propia.
    formulas = con_libro(ruta, lambda wb, app: formulas_cantidad(
        wb, e.ramal["abrev"], "ANE", "MT"))
    assert len(formulas) == 3, f"se esperaban 3 filas y hay {formulas}"
    con_literal = [f for f in formulas if re.search(r"\)\s*\+\s*5(?!\d)", f)]
    assert len(con_literal) == 1, \
        f"una fila debe llevar '+5' y no se encontró en {formulas}"
    for f in formulas:
        assert "SUMIF" in f.upper(), f"fórmula sin SUMIF: {f!r}"

    # 2) Nada se rompe en Excel.
    malas = con_libro(ruta, lambda wb, app: celdas_con_error(
        wb, e.ramal["abrev"], "ANE"))
    assert not malas, f"celdas con error en el Ramal: {malas}"
    # La hoja VARIACIÓN del Ramal existe y no tiene errores.
    malas_var = con_libro(ruta, lambda wb, app: celdas_con_error(
        wb, e.ramal["abrev"], "VARIACION"))
    assert not malas_var, f"celdas con error en VARIACIÓN del Ramal: {malas_var}"

    # 3) La firma SUMIF sigue funcionando con el "+literal" añadido.
    def firmas(wb, app):
        return [(g["token"], main._xl_firma_ramal(g))
                for g in main._xl_grupos_desde_wb(wb)]
    pares = con_libro(ruta, firmas)
    firma = {main._xl_token(t).upper(): v for t, v in pares}.get(
        main._xl_token(e.ramal["abrev"]).upper())
    assert firma == 1000, f"firma del Ramal {firma} (esperada 1000) en {pares}"

    # 4) Roundtrip: la fila propia vuelve con su cantidad, sin duplicados.
    e2 = estado_vacio()
    main.cargar_proyecto_desde_excel(e2, ruta)
    propias = e2._cant_propia("ANE")["MT"]
    assert propias.get("P999") == Decimal(5), \
        f"la fila propia debía volver con 5 y volvió con {propias.get('P999')}"
    assert propias.get("P001", Decimal(0)) == 0, \
        "una fila heredada no debe convertirse en propia"
    filas = e2.filas(e2.ramal, "ANE")["MT"]
    codigos = [f["codigo"] for f in filas]
    assert codigos == ["P001", "K001", "P999"], f"filas inesperadas: {codigos}"
    assert e2._cant_propia("ANE")["MT"].get("P999", Decimal(0)) > 0

    # 5) T1 sobrevive al viaje: VARIACIÓN del Ramal sigue activa y los
    #    proyectos siguen sin ella.
    assert e2.ramal["hojas"]["VARIACION"] is True
    assert all(not p["hojas"]["VARIACION"] for p in e2.proyectos), \
        "importar no debe activar VARIACIÓN en los proyectos"

    # 6) Totales y orden SICOA.
    assert totales_ramal(e2, "ANE") == totales, "los totales cambiaron al reimportar"
    indice = {c["codigo"]: i for i, c in enumerate(e2.bd)}
    posiciones = [indice.get(c, len(indice)) for c in codigos]
    assert posiciones == sorted(posiciones), f"orden SICOA roto: {codigos}"


# ═════════════════════════════════════════════════════════════════════
#  CASO 3 · compatibilidad con la fórmula anterior (sin "+literal")
# ═════════════════════════════════════════════════════════════════════
@excel
def test_caso3_formula_vieja_sin_literal_no_inventa_filas():
    if pytest is not None:
        pytest.importorskip("xlwings")

    origen = os.path.join(CARPETA, "t1t2_caso2.xlsx")
    if not os.path.exists(origen):
        _crear_caso2_si_falta()
    assert os.path.exists(origen), "falta el libro del Caso 2"

    ruta = os.path.join(CARPETA, "t1t2_caso3.xlsx")
    if os.path.exists(ruta):
        os.remove(ruta)
    shutil.copyfile(origen, ruta)

    abrev = "RAMAL"                              # el Ramal se llama RAMAL aquí

    def quitar_literal(wb, app):
        token = main._xl_token(abrev)
        nombre = main.PREFIJOS_HOJA["ANE"] + token
        assert nombre in [s.name for s in wb.sheets], \
            f"no existe la hoja {nombre!r} en {[s.name for s in wb.sheets]}"
        ws = wb.sheets[nombre]
        tabla = main._xl_tabla(ws, main._xl_nombre_tabla("ANE", token, "MT"))
        assert tabla is not None, "no se encontró la tabla del Ramal"
        col = main._xl_col_cantidad(tabla)
        cuerpo = tabla.api.ListColumns(col).DataBodyRange
        textos = [str(f[0] or "") for f in main._xl_valores(cuerpo.Formula)]
        # Solo se quita el literal NUMÉRICO del final (`...)+5`): los SUMIF de
        # los proyectos también van unidos con «+» y deben conservarse.
        limpios = [re.sub(r"\)\s*\+\s*[\d.]+$", ")", t) for t in textos]
        quitados = sum(1 for a, b in zip(textos, limpios) if a != b)
        # Escribir primero valores rompe la columna calculada; si no, Excel
        # copia la fórmula de una fila a todas al asignar celda por celda.
        cuerpo.Value = 0
        for i, formula in enumerate(limpios, start=1):
            cuerpo.Cells(i, 1).Formula = formula
        app.calculate()
        wb.save()
        return quitados

    quitados = con_libro(ruta, quitar_literal)
    assert quitados == 1, f"se esperaba quitar 1 literal y se quitaron {quitados}"

    e2 = estado_vacio()
    main.cargar_proyecto_desde_excel(e2, ruta)
    propias = {c: q for c, q in e2._cant_propia("ANE")["MT"].items() if q > 0}
    assert not propias, f"un archivo viejo no debe traer filas propias: {propias}"
    codigos = [f["codigo"] for f in e2.filas(e2.ramal, "ANE")["MT"]]
    assert codigos == ["P001", "K001"], f"fila fantasma en el Ramal: {codigos}"
    assert e2.ramal["hojas"]["VARIACION"] is True, \
        "la VARIACIÓN del Ramal ya existía en el archivo y se conserva"


def _crear_caso2_si_falta():
    """Crea el libro del Caso 2 si algún día se ejecuta el Caso 3 solo."""
    ruta = os.path.join(CARPETA, "t1t2_caso2.xlsx")
    if not os.path.exists(ruta):
        test_caso2_fila_propia_y_variacion_solo_en_el_ramal()


# ═════════════════════════════════════════════════════════════════════
#  CIERRE: plantilla intacta y ningún EXCEL.EXE huérfano
# ═════════════════════════════════════════════════════════════════════
@excel
def test_zz_plantilla_intacta_y_sin_excel_huerfano():
    _recordar_estado_inicial()
    despues = _huella_plantilla()
    if PLANTILLA_ANTES is not None and despues is not None:
        assert despues == PLANTILLA_ANTES, \
            "¡la plantilla real cambió! " \
            f"antes={PLANTILLA_ANTES} después={despues}"
    abiertos = _procesos_excel()
    print("   procesos EXCEL.EXE:", len(abiertos))
    assert len(abiertos) <= (EXCEL_ANTES or 0), \
        f"quedaron {len(abiertos)} EXCEL.EXE abiertos (antes {EXCEL_ANTES})"


# ═════════════════════════════════════════════════════════════════════
#  RUNNER sin pytest
# ═════════════════════════════════════════════════════════════════════
def _runner():
    pruebas = [(n, f) for n, f in sorted(globals().items())
               if n.startswith("test_") and callable(f)]
    fallas, omitidas = [], []
    for nombre, fn in pruebas:
        try:
            fn()
            print(f"PASA   {nombre}")
        except ImportError as e:
            omitidas.append(nombre)
            print(f"OMITE  {nombre}  -> {e!r}")
        except Exception as e:
            fallas.append(nombre)
            print(f"FALLA  {nombre}  -> {e!r}")
            traceback.print_exc()
    print()
    print(f"RESULTADO: {len(pruebas) - len(fallas) - len(omitidas)}/{len(pruebas)} PASA"
          f" · {len(omitidas)} omitidas · {len(fallas)} fallas")
    if fallas:
        print("FALLARON:", fallas)
    return 1 if fallas else 0


if __name__ == "__main__":
    sys.exit(_runner())
