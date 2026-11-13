# -*- coding: utf-8 -*-
"""
Pruebas del Ramal renombrado: `_xl_firma_ramal` / `_xl_indice_ramal`.

El Ramal ya no se reconoce por el nombre de la hoja, sino por las fórmulas
SUMIF / SUMAR.SI de la columna "Cantidad a ejecutar" de ANE / ANE-VAR.

Este archivo se puede usar de dos formas:
  · `python tests/test_ramal_firma.py`  → runner propio, sin pytest.
  · `python -m pytest tests/test_ramal_firma.py` → si pytest está instalado
    (las pruebas de Excel quedan marcadas con `excel`).

Niveles:
  · Puras  : `_xl_indice_ramal` (no toca Excel ni widgets).
  · Excel  : `_xl_firma_ramal` sobre libros DESECHABLES creados en %TEMP%
             (nunca se abre ni se escribe la plantilla del proyecto).
  · Integr.: exportar 2 proyectos + Ramal renombrado y reimportar el .xlsx.

No modifica el código fuente: solo lee funciones de `main`.
"""

import atexit
import os
import shutil
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

    def excel(func):                            # sin pytest: marca inocua
        return func

import main                                     # noqa: E402


# ═════════════════════════════════════════════════════════════════════
#  0) limpieza de temporales
# ═════════════════════════════════════════════════════════════════════
def _limpiar_temporales():
    base = tempfile.gettempdir()
    try:
        nombres = os.listdir(base)
    except OSError:
        return
    for nombre in nombres:
        if nombre.startswith(("anexos_ramal_",)):
            shutil.rmtree(os.path.join(base, nombre), ignore_errors=True)


atexit.register(_limpiar_temporales)


# ═════════════════════════════════════════════════════════════════════
#  1) PRUEBAS PURAS: _xl_indice_ramal
# ═════════════════════════════════════════════════════════════════════
# (tokens, firmas, esperado, nota)
CASOS_INDICE = [
    ([], [], None, "sin grupos"),
    (["COM1"], [0], None, "un solo grupo sin firma"),
    (["COM1"], [1000], None, "un solo grupo con firma"),
    (["A", "B"], [0, 1000, 0], None, "longitudes distintas"),
    (["A", "B", "C"], [0], None, "firmas incompletas"),
    (["COM1", "RAMAL2", "RAMAL"], [0, 1000, 0], 1, "la firma manda sobre el nombre"),
    (["COM1", "RAMAL"], [0, 0], 1, "sin firma, se prefiere el nombre RAMAL"),
    (["RAMAL", "RAMAL2"], [0, 1000], 1, "un RAMAL vacío no gana sobre uno con firma"),
    (["RAMAL", "RAMAL2"], [1000, 1000], 0, "empate de firma resuelto por el nombre RAMAL"),
    (["A", "B"], [1000, 1000], None, "empate sin RAMAL: no se marca ninguno"),
    (["A", "B"], [0, 0], None, "sin firma y sin RAMAL: no se marca ninguno"),
    (["com1", "ramal2", "ramal"], [0, 1000, 0], 1, "minúsculas: firma manda"),
    (["RAMAL", "ramal2"], [1000, 1000], 0, "minúsculas: desempate por nombre"),
    (["ramal2", "com1"], [1000, 0], 0, "minúsculas: firma única"),
]


def test_indice_ramal():
    """Todos los casos de `_xl_indice_ramal`, incluidos los límite."""
    fallos = []
    for tokens, firmas, esperado, nota in CASOS_INDICE:
        obtenido = main._xl_indice_ramal(tokens, firmas)
        if obtenido != esperado:
            fallos.append(f"{nota}: tokens={tokens} firmas={firmas} → "
                          f"{obtenido!r}, se esperaba {esperado!r}")
    assert not fallos, "\n".join(fallos)


def test_indice_ramal_no_depende_del_orden_del_nombre():
    """El Ramal gana por firma aunque no se llame RAMAL ni esté primero."""
    assert main._xl_indice_ramal(["COM_A", "COM_B", "RAMAL2"], [0, 0, 1000]) == 2
    assert main._xl_indice_ramal(["RAMAL2", "COM_A", "COM_B"], [1000, 0, 0]) == 0


# ═════════════════════════════════════════════════════════════════════
#  2) PRUEBAS REALES: _xl_firma_ramal con Excel desechable
# ═════════════════════════════════════════════════════════════════════
ENCABEZADOS = ["CODIGO COMPONENTE", "Tipo", "DESCRIPCION", "UNIDAD",
               "Precio unitario", "Cantidad a ejecutar", "Precio total"]

SUMIF = "=SUMIF(SRC_T[CODIGO],[@[CODIGO COMPONENTE]],SRC_T[CANT])"


def _fila(codigo, cantidad=1):
    return [codigo, "Media tensión", "POSTE DE PRUEBA", "U", 10.0, cantidad, 10.0]


class _Excel:
    """App de Excel oculta y desechable para las pruebas de firma."""

    def __init__(self):
        import xlwings as xw
        self.app = xw.App(visible=False, add_book=False)
        for atributo in ("display_alerts", "screen_updating"):
            try:
                setattr(self.app, atributo, False)
            except Exception:
                pass
        self.libro = None

    def nuevo_libro(self):
        self.libro = self.app.books.add()
        # Tabla fuente para que el SUMIF sea una fórmula válida y calculable.
        src = self.libro.sheets.add(name="SRC")
        src.range("A1").value = [["CODIGO", "CANT"], ["P001", 1], ["P002", 2]]
        src.tables.add(src.range("A1:B3"), name="SRC_T")
        return self.libro

    def cerrar_libro(self):
        if self.libro is not None:
            try:
                self.libro.close()
            except Exception:
                pass
            self.libro = None

    def cerrar(self):
        self.cerrar_libro()
        try:
            self.app.quit()
        except Exception:
            pass


def _hoja(wb, nombre):
    """Devuelve (creando si hace falta) la hoja con ese nombre."""
    for ws in wb.sheets:
        if ws.name == nombre:
            return ws
    return wb.sheets.add(name=nombre)


def _crear_tabla(ws, nombre, filas, encabezados=None):
    enc = list(encabezados or ENCABEZADOS)
    ws.range("A1").value = [enc]
    if filas:
        ws.range("A2").value = [list(f) for f in filas]
    fin = f"{chr(ord('A') + len(enc) - 1)}{1 + len(filas)}"
    return ws.tables.add(ws.range(f"A1:{fin}"), name=nombre)


def _escribir_formulas(ws, columna, formulas):
    for i, formula in enumerate(formulas, start=2):
        ws.range(f"{columna}{i}").formula = formula


def _grupo(*pares):
    """grupo = {"token": ..., "hojas": {hoja: ws}} como en _xl_grupos_desde_wb."""
    return {"token": "RAMAL2", "hojas": dict(pares)}


@excel
def test_firma_sumif_en_todas_las_filas():
    xl = _Excel()
    try:
        libro = xl.nuevo_libro()
        ws = _hoja(libro, "ANE. - RAMAL2")
        _crear_tabla(ws, "A_MT_RAMAL2",
                     [_fila("P001"), _fila("P002"), _fila("P001"), _fila("P002")])
        _escribir_formulas(ws, "F", [SUMIF] * 4)
        firma = main._xl_firma_ramal(_grupo(("ANE", ws)))
        assert firma == 1000, f"firma={firma} (se esperaba 1000)"
    finally:
        xl.cerrar()


@excel
def test_firma_una_sola_sumif_entre_numeros_no_basta():
    """Caso bloqueante: 3 números + 1 SUMIF no puede ser Ramal."""
    xl = _Excel()
    try:
        libro = xl.nuevo_libro()
        ws = _hoja(libro, "ANE. - COM1")
        _crear_tabla(ws, "A_MT_COM1",
                     [_fila("P001", 1), _fila("P002", 2), _fila("P001", 3),
                      _fila("P002", 4)])
        _escribir_formulas(ws, "F", ["=1", "=2", "=3", SUMIF])
        firma = main._xl_firma_ramal(_grupo(("ANE", ws)))
        assert firma == 0, (f"firma={firma}: una SUMIF suelta no convierte "
                            f"el proyecto en Ramal")
    finally:
        xl.cerrar()


@excel
def test_firma_formulas_sin_sumif():
    """'Todas fórmulas' ya no debe valer como Ramal."""
    xl = _Excel()
    try:
        libro = xl.nuevo_libro()
        ws = _hoja(libro, "ANE. - COM1")
        _crear_tabla(ws, "A_MT_COM1",
                     [_fila("P001"), _fila("P002"), _fila("P001"), _fila("P002")])
        _escribir_formulas(ws, "F", ["=1+1"] * 4)
        firma = main._xl_firma_ramal(_grupo(("ANE", ws)))
        assert firma == 0, f"firma={firma}: fórmulas sin SUMIF no son Ramal"
    finally:
        xl.cerrar()


@excel
def test_firma_tabla_vacia():
    xl = _Excel()
    try:
        libro = xl.nuevo_libro()
        ws = _hoja(libro, "ANE. - RAMAL2")
        tabla = _crear_tabla(ws, "A_MT_RAMAL2", [_fila("P001")])
        _escribir_formulas(ws, "F", [SUMIF])
        tabla.api.ListRows(1).Delete()          # queda solo el encabezado
        firma = main._xl_firma_ramal(_grupo(("ANE", ws)))
        assert firma == 0, f"firma={firma} con la tabla sin filas"
    finally:
        xl.cerrar()


@excel
def test_firma_una_sola_fila_con_sumif():
    xl = _Excel()
    try:
        libro = xl.nuevo_libro()
        ws = _hoja(libro, "ANE. - RAMAL2")
        _crear_tabla(ws, "A_MT_RAMAL2", [_fila("P001")])
        _escribir_formulas(ws, "F", [SUMIF])
        firma = main._xl_firma_ramal(_grupo(("ANE", ws)))
        assert firma == 1000, f"firma={firma} con 1/1 filas SUMIF"
    finally:
        xl.cerrar()


@excel
def test_firma_mitad_exacta_no_basta():
    xl = _Excel()
    try:
        libro = xl.nuevo_libro()
        ws = _hoja(libro, "ANE. - RAMAL2")
        _crear_tabla(ws, "A_MT_RAMAL2", [_fila("P001", 1), _fila("P002", 2)])
        _escribir_formulas(ws, "F", [SUMIF, "=2"])
        firma = main._xl_firma_ramal(_grupo(("ANE", ws)))
        assert firma == 0, f"firma={firma}: 1 de 2 no es 'más de la mitad'"
    finally:
        xl.cerrar()


@excel
def test_firma_dos_de_tres_si_basta():
    xl = _Excel()
    try:
        libro = xl.nuevo_libro()
        ws = _hoja(libro, "ANE. - RAMAL2")
        _crear_tabla(ws, "A_MT_RAMAL2",
                     [_fila("P001", 1), _fila("P002", 2), _fila("P001", 3)])
        _escribir_formulas(ws, "F", [SUMIF, SUMIF, "=3"])
        firma = main._xl_firma_ramal(_grupo(("ANE", ws)))
        assert firma == 1000, f"firma={firma}: 2 de 3 sí es 'más de la mitad'"
    finally:
        xl.cerrar()


@excel
def test_firma_tabla_renombrada_no_se_filtra_por_nombre():
    """Una tabla renombrada a mano (sin prefijo A_/AV_) no provoca falso negativo."""
    xl = _Excel()
    try:
        libro = xl.nuevo_libro()
        ws = _hoja(libro, "ANE. - RAMAL2")
        tabla = _crear_tabla(ws, "A_MT_RAMAL2",
                             [_fila("P001"), _fila("P002"), _fila("P001")])
        tabla.name = "OTRA_COSA"
        _escribir_formulas(ws, "F", [SUMIF] * 3)
        nombre = str(tabla.name)
        assert nombre == "OTRA_COSA", f"la tabla no se renombró: {nombre!r}"
        firma = main._xl_firma_ramal(_grupo(("ANE", ws)))
        assert firma == 1000, f"firma={firma} con la tabla renombrada"
    finally:
        xl.cerrar()


@excel
def test_firma_ane_var_sumif():
    xl = _Excel()
    try:
        libro = xl.nuevo_libro()
        ws = _hoja(libro, "ANE. - VAR. - RAMAL2")
        _crear_tabla(ws, "AV_MT_RAMAL2",
                     [_fila("P001"), _fila("P002"), _fila("P001")])
        _escribir_formulas(ws, "F", [SUMIF] * 3)
        firma = main._xl_firma_ramal(_grupo(("ANE_VAR", ws)))
        assert firma == 1000, f"firma={firma} en ANE-VAR"
    finally:
        xl.cerrar()


@excel
def test_firma_suma_hojas_del_grupo():
    """ANE + ANE-VAR se acumulan en la misma firma."""
    xl = _Excel()
    try:
        libro = xl.nuevo_libro()
        ane = _hoja(libro, "ANE. - RAMAL2")
        _crear_tabla(ane, "A_MT_RAMAL2", [_fila("P001", 1), _fila("P002", 2)])
        _escribir_formulas(ane, "F", [SUMIF, SUMIF])
        var = _hoja(libro, "ANE. - VAR. - RAMAL2")
        _crear_tabla(var, "AV_MT_RAMAL2", [_fila("P001", 1), _fila("P002", 2)])
        _escribir_formulas(var, "F", [SUMIF, SUMIF])
        firma = main._xl_firma_ramal(_grupo(("ANE", ane), ("ANE_VAR", var)))
        assert firma == 1000, f"firma={firma} sumando ANE + ANE-VAR"
    finally:
        xl.cerrar()


@excel
def test_firma_encuentra_la_columna_por_nombre():
    """`_xl_col_cantidad` busca por nombre, no por la posición 6 fija."""
    xl = _Excel()
    try:
        libro = xl.nuevo_libro()
        ws = _hoja(libro, "ANE. - RAMAL2")
        enc = ["CODIGO COMPONENTE", "Tipo", "DESCRIPCION", "UNIDAD",
               "Cantidad a ejecutar", "Relleno", "Precio total"]
        _crear_tabla(ws, "A_MT_RAMAL2",
                     [["P001", "Media tensión", "POSTE", "U", 1, 0, 10.0],
                      ["P002", "Media tensión", "POSTE", "U", 2, 0, 10.0]],
                     encabezados=enc)
        _escribir_formulas(ws, "E", [SUMIF, SUMIF])
        _escribir_formulas(ws, "F", ["=0", "=0"])
        firma = main._xl_firma_ramal(_grupo(("ANE", ws)))
        assert firma == 1000, (f"firma={firma}: se leyó la columna 6 fija en vez "
                               f"de buscar 'Cantidad a ejecutar' por nombre")
    finally:
        xl.cerrar()


@excel
def test_firma_sin_hojas_editables():
    xl = _Excel()
    try:
        libro = xl.nuevo_libro()
        ws = _hoja(libro, "OC - RAMAL2")
        ws.range("A1").value = [["No."]]
        assert main._xl_firma_ramal({"token": "RAMAL2", "hojas": {}}) == 0
        assert main._xl_firma_ramal({"token": "RAMAL2",
                                     "hojas": {"OC": ws}}) == 0
    finally:
        xl.cerrar()


# ═════════════════════════════════════════════════════════════════════
#  3) INTEGRACIÓN: exportar con Ramal renombrado y reimportar
# ═════════════════════════════════════════════════════════════════════
def _estado_prueba(abreviaturas, ramal_abrev, codigos=("P001", "P002")):
    """Estado con N proyectos (ANE + ANE-VAR) y un Ramal presente."""
    e = main.Estado.__new__(main.Estado)
    e.bd = [dict(c) for c in main.cargar_bd_inicial()]
    e.reset()
    e.proyectos = []
    for i, abrev in enumerate(abreviaturas):
        base = e._base(abrev, main.PALETA[i % len(main.PALETA)])
        base["nombre"] = f"COMUNIDAD {abrev}"
        base["hojas"] = {h: h in ("ANE", "ANE_VAR") for h in main.HOJAS}
        base["visibles"] = ["ANE", "ANE_VAR"]
        base["items"] = {h: main._vacio() for h in main.EDITABLES}
        for j, codigo in enumerate(codigos):
            for hoja in base["items"]:
                base["items"][hoja]["MT"] = [
                    {"codigo": codigo, "cantidad": str(j + 1 + i),
                     "redondeo": "ROUND"}]
        e.proyectos.append(base)
    e.ramal = e._base(ramal_abrev, "#7f8c8d", True)
    e.ramal["abrev"] = ramal_abrev
    e.ramal_presente = True
    e._sincronizar_ramal()
    e.cambios = True
    return e


def _exportar(estado, nombre="ANEXO - PRUEBA - 21.10.2026.xlsx"):
    carpeta = tempfile.mkdtemp(prefix="anexos_ramal_")
    ruta = os.path.join(carpeta, nombre)
    main.exportar_xlsx(estado, ruta)
    return ruta, carpeta


def _reimportar(ruta):
    e = main.Estado.__new__(main.Estado)
    e.bd = []
    e.reset()
    main.cargar_proyecto_desde_excel(e, ruta)
    return e


@excel
def test_exportar_y_reimportar_ramal_renombrado():
    estado = _estado_prueba(["COM1", "COM2"], "RAMAL2")
    ruta, carpeta = _exportar(estado)
    try:
        # · el Excel exportado lleva la firma SUMIF en el Ramal
        from openpyxl import load_workbook
        wbf = load_workbook(ruta, data_only=False)
        hoja_ramal = main.PREFIJOS_HOJA["ANE"] + "RAMAL2"
        assert hoja_ramal in wbf.sheetnames, wbf.sheetnames
        sumifs = []
        for fila in wbf[hoja_ramal].iter_rows():
            for celda in fila:
                if isinstance(celda.value, str) and "SUMIF" in celda.value.upper():
                    sumifs.append(celda.coordinate)
        assert sumifs, f"el Ramal exportado no tiene SUMIF: {wbf.sheetnames}"

        # · al reimportar, el Ramal renombrado se reconoce
        e2 = _reimportar(ruta)
        assert e2.ramal["abrev"].upper() == "RAMAL2", e2.ramal["abrev"]
        assert e2.ramal_presente is True
        assert len(e2.proyectos) == 2, [p["abrev"] for p in e2.proyectos]
        abrevs = sorted(p["abrev"].upper() for p in e2.proyectos)
        assert abrevs == ["COM1", "COM2"], abrevs
        assert all(a != "RAMAL2" for a in abrevs), abrevs
    finally:
        shutil.rmtree(carpeta, ignore_errors=True)


@excel
def test_exportar_y_reimportar_ramal_estandar():
    """Regresión: el Ramal llamado RAMAL se sigue reconociendo."""
    estado = _estado_prueba(["COM1", "COM2"], "RAMAL")
    ruta, carpeta = _exportar(estado)
    try:
        e2 = _reimportar(ruta)
        assert e2.ramal["abrev"].upper() == "RAMAL", e2.ramal["abrev"]
        assert e2.ramal_presente is True
        assert len(e2.proyectos) == 2, [p["abrev"] for p in e2.proyectos]
        abrevs = sorted(p["abrev"].upper() for p in e2.proyectos)
        assert abrevs == ["COM1", "COM2"], abrevs
    finally:
        shutil.rmtree(carpeta, ignore_errors=True)


@excel
def test_exportar_y_reimportar_un_proyecto_con_ramal():
    """Un solo proyecto + Ramal (caso 'Excel viejo ya trae Ramal')."""
    estado = _estado_prueba(["COM1"], "RAMAL")
    ruta, carpeta = _exportar(estado)
    try:
        e2 = _reimportar(ruta)
        assert e2.ramal_presente is True, "se perdió el Ramal de un solo proyecto"
        assert e2.ramal["abrev"].upper() == "RAMAL", e2.ramal["abrev"]
        assert len(e2.proyectos) == 1, [p["abrev"] for p in e2.proyectos]
    finally:
        shutil.rmtree(carpeta, ignore_errors=True)


@excel
def test_exportar_y_reimportar_sin_ramal():
    """Dos proyectos sin Ramal: nadie debe marcarse como Ramal."""
    estado = _estado_prueba(["COM1", "COM2"], "RAMAL")
    estado.ramal_presente = False
    estado.ramal["hojas"] = {h: False for h in main.HOJAS}
    ruta, carpeta = _exportar(estado)
    try:
        e2 = _reimportar(ruta)
        assert e2.ramal_presente is False, e2.ramal["abrev"]
        abrevs = sorted(p["abrev"].upper() for p in e2.proyectos)
        assert abrevs == ["COM1", "COM2"], abrevs
    finally:
        shutil.rmtree(carpeta, ignore_errors=True)


@excel
def test_proyecto_llamado_ramal_y_ramal_renombrado():
    """Un proyecto puede llamarse RAMAL; el Ramal real es RAMAL2."""
    estado = _estado_prueba(["RAMAL", "COM1"], "RAMAL2")
    ruta, carpeta = _exportar(estado)
    try:
        e2 = _reimportar(ruta)
        assert e2.ramal["abrev"].upper() == "RAMAL2", e2.ramal["abrev"]
        assert e2.ramal_presente is True
        abrevs = sorted(p["abrev"].upper() for p in e2.proyectos)
        assert abrevs == ["COM1", "RAMAL"], abrevs
    finally:
        shutil.rmtree(carpeta, ignore_errors=True)


# ═════════════════════════════════════════════════════════════════════
#  4) runner sin pytest
# ═════════════════════════════════════════════════════════════════════
def _runner():
    pruebas = [v for k, v in sorted(globals().items())
               if k.startswith("test_") and callable(v)]
    pasan, fallan = 0, []
    for prueba in pruebas:
        try:
            prueba()
        except Exception:
            fallan.append((prueba.__name__, traceback.format_exc()))
            print(f"FALLA  {prueba.__name__}")
        else:
            pasan += 1
            print(f"PASA   {prueba.__name__}")
    print("-" * 70)
    for nombre, detalle in fallan:
        print(f"\n=== {nombre} ===")
        print(detalle)
    print(f"{pasan} pasaron, {len(fallan)} fallaron, {len(pruebas)} en total")
    return 0 if not fallan else 1


if __name__ == "__main__":
    sys.exit(_runner())
