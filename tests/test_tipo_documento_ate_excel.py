# -*- coding: utf-8 -*-
"""
Verificación con EXCEL REAL de la regla `Tipo de Documento`
(ATE solo si el código es numérico y MAYOR a UMBRAL_CODIGO_ATE = 2000).

Requiere Excel instalado (xlwings/COM) → @pytest.mark.excel.

Escenario que se exporta (proyecto ATE_T, todo en la línea MT):

    código  ANE  ANE-VAR  variación  tipo esperado
    2500     5      0         -5         ATE   (numérico > 2000 gana al OC)
    3000     0      5         +5         ATE   (numérico > 2000 gana al OTS)
    1999     4      0         -4         OC
    2000     1      5         +4         OTS   (2000 exacto NO es ATE)
    1083     2      2          0         ---   (no entra en OC/OTS/ATE)

Nunca modifica la plantilla ni main.py: `exportar_xlsx` copia el machote a
un temporal. Los .xlsx de salida van a la carpeta temporal indicada.
"""

import os
import shutil
import subprocess
import tempfile
from decimal import Decimal

import pytest

import main

pytestmark = pytest.mark.excel

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOGOTIPOS_BLANCO = 16777215

# línea → código → (cantidad ANE, cantidad ANE-VAR)
ESCENARIO = {
    "2500": ("5", "0"),
    "3000": ("1", "6"),   # aparece en ANE y ANE-VAR: sirve para el formato
    "1999": ("4", "0"),
    "2000": ("1", "5"),
    "1083": ("2", "2"),
}
ESPERADO_TIPO = {"2500": "ATE", "3000": "ATE", "1999": "OC",
                 "2000": "OTS", "1083": "---"}
ESPERADO_EN_HOJA = {"ATE": {"2500", "3000"}, "OTS": {"2000"}, "OC": {"1999"}}

_cache = {}


# ── utilidades ───────────────────────────────────────────────────────
def _carpeta_temporal():
    """Carpeta acordada para los desechables (o %TEMP% si no existe)."""
    preferida = r"C:\Users\deivi\AppData\Local\Temp\opencode"
    base = preferida if os.path.isdir(preferida) else tempfile.gettempdir()
    ruta = os.path.join(base, "ate_umbral")
    os.makedirs(ruta, exist_ok=True)
    return ruta


def _procesos_excel():
    salida = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/NH"],
        capture_output=True, text=True, errors="replace")
    return salida.stdout.upper().count("EXCEL.EXE")


def _cerrar(app, pythoncom):
    """Cierra Excel liberando ANTES los proxies COM (evita 0x800706ba al GC)."""
    import gc
    gc.collect()
    main._cerrar_excel(app, pythoncom)
    gc.collect()


def _estado():
    if "estado" in _cache:
        return _cache["estado"]
    e = main.Estado.__new__(main.Estado)
    e.bd = [main.componente_nuevo(c, f"COMPONENTE DE PRUEBA {c}", "U", "100.00")
            for c in ESCENARIO]
    for c in e.bd:
        c["MT"] = 1
    e.reset()
    b = e.proyectos[0]
    b["abrev"] = "ATE_T"
    b["nombre"] = "PRUEBA ATE"
    b["hojas"] = {h: True for h in main.HOJAS}
    b["visibles"] = list(main.HOJAS)
    b["items"] = {"ANE": main._vacio(), "ANE_VAR": main._vacio()}
    for cod, (q_ane, q_av) in ESCENARIO.items():
        if q_ane != "0":
            b["items"]["ANE"]["MT"].append(
                {"codigo": cod, "cantidad": q_ane, "redondeo": "ROUND"})
        if q_av != "0":
            b["items"]["ANE_VAR"]["MT"].append(
                {"codigo": cod, "cantidad": q_av, "redondeo": "ROUND"})
    e.cambios = True
    _cache["estado"] = e
    return e


def _exportar(nombre):
    carpeta = _carpeta_temporal()
    ruta = os.path.join(carpeta, nombre)
    main.exportar_xlsx(_estado(), ruta)
    return ruta


def _exportado():
    if "ruta" not in _cache:
        _cache["ruta"] = _exportar("ANEXO - ATE_T - umbral.xlsx")
    return _cache["ruta"]


def _indice_columna(tabla, titulo):
    """Índice (1-based) de la columna cuyo encabezado coincide con `titulo`."""
    objetivo = main._xl_normalizar_etiqueta(titulo)
    for i in range(1, int(tabla.api.ListColumns.Count) + 1):
        if main._xl_normalizar_etiqueta(
                str(tabla.api.ListColumns(i).Name or "")) == objetivo:
            return i
    raise AssertionError(f"no existe la columna {titulo!r} en {tabla.name}")


def _valor(valores, fila, col):
    """Valor de una fila de `Range.Value` sea tupla de tuplas o escalar."""
    if not isinstance(valores, tuple):
        return valores
    if valores and not isinstance(valores[0], tuple):
        return valores[col - 1]
    return valores[fila][col - 1]


def _filas_tabla(tabla):
    """[(codigo, tipo, ...)] de una tabla de documento, ya calculada en Excel."""
    rango = tabla.api.DataBodyRange
    if rango is None:
        return []
    n = int(rango.Rows.Count)
    valores = rango.Value
    col_cod = 1
    col_tipo = _indice_columna(tabla, "TIPO DE DOCUMENTO")
    out = []
    for i in range(n):
        cod = _valor(valores, i, col_cod)
        if cod in (None, ""):
            continue
        out.append((main._cod(cod), _valor(valores, i, col_tipo)))
    return out


def _tabla(ws, hoja, linea="MT"):
    return main._xl_tabla(ws, main._xl_nombre_tabla(
        hoja, main._xl_token(_estado().proyectos[0]["abrev"]), linea))


# ═════════════════════════════════════════════════════════════════════
#  0) el script de verificación que pidió el Director
# ═════════════════════════════════════════════════════════════════════
def test_script_verifica_ate_dice_todo_pasa():
    script = os.path.join(RAIZ, "herramientas", "verifica_ate.py")
    if not os.path.exists(script):
        pytest.skip("no existe herramientas/verifica_ate.py")
    antes = _procesos_excel()
    salida = subprocess.run([sys_executable(), script], cwd=RAIZ,
                            capture_output=True, text=True, errors="replace")
    print(salida.stdout)
    print(salida.stderr)
    assert "RESULTADO: TODO PASA" in salida.stdout, salida.stdout
    assert salida.returncode == 0
    assert _procesos_excel() <= antes, "el script dejó Excel abierto"


def sys_executable():
    import sys
    return sys.executable


# ═════════════════════════════════════════════════════════════════════
#  1) paridad VARIACIÓN (Excel) ↔ Estado._tipo_documento (Python)
# ═════════════════════════════════════════════════════════════════════
def test_variacion_paridad_total_con_python():
    ruta = _exportado()
    app, pythoncom = main._abrir_excel()
    wb = None
    try:
        wb = app.books.open(ruta)
        app.calculate()
        b = _estado().proyectos[0]
        ws = wb.sheets[main.Estado.titulo(_estado(), b, "VARIACION")]
        tabla = _tabla(ws, "VARIACION")
        assert tabla is not None, "no se encontró la tabla de VARIACIÓN"

        filas = _filas_tabla(tabla)
        assert len(filas) == len(ESCENARIO), filas

        discrepancias = []
        for cod, tipo_excel in filas:
            q_ane, q_av = ESCENARIO[cod]
            dif = Decimal(q_av) - Decimal(q_ane)
            tipo_py = _estado()._tipo_documento(cod, dif)
            if tipo_excel != tipo_py:
                discrepancias.append((cod, tipo_excel, tipo_py))
            assert tipo_excel == ESPERADO_TIPO[cod], \
                f"{cod}: Excel dijo {tipo_excel!r}, se esperaba {ESPERADO_TIPO[cod]!r}"
        assert not discrepancias, f"discrepancias Excel/Python: {discrepancias}"
    finally:
        try:
            if wb is not None:
                wb.close()
        finally:
            _cerrar(app, pythoncom)


# ═════════════════════════════════════════════════════════════════════
#  2) OC / OTS / ATE: cada hoja con las filas de su tipo
# ═════════════════════════════════════════════════════════════════════
def test_hojas_oc_ots_ate_solo_con_sus_filas():
    ruta = _exportado()
    app, pythoncom = main._abrir_excel()
    wb = None
    try:
        wb = app.books.open(ruta)
        app.calculate()
        e = _estado()
        b = e.proyectos[0]
        obtenido = {}
        for hoja in ("OC", "OTS", "ATE"):
            ws = wb.sheets[e.titulo(b, hoja)]
            tabla = _tabla(ws, hoja)
            assert tabla is not None, f"falta la tabla de {hoja}"
            pares = _filas_tabla(tabla)
            obtenido[hoja] = {c for c, _ in pares}
            # el Tipo de Documento de la propia hoja también debe coincidir
            for cod, tipo in pares:
                assert tipo == hoja, f"{hoja}: fila {cod} con tipo {tipo!r}"
        assert obtenido == ESPERADO_EN_HOJA, obtenido
    finally:
        try:
            if wb is not None:
                wb.close()
        finally:
            _cerrar(app, pythoncom)


# ═════════════════════════════════════════════════════════════════════
#  3) formato condicional: solo los ATE quedan en blanco
# ═════════════════════════════════════════════════════════════════════
def _colores_codigos(app, ruta):
    """{codigo: Font.Color de DisplayFormat} en la columna 1 de la tabla ANE.

    No devuelve objetos COM: se descartan antes de cerrar el libro para que
    Excel no quede con proxies vivos.
    """
    wb = None
    try:
        wb = app.books.open(ruta)
        app.calculate()
        e = _estado()
        b = e.proyectos[0]
        ws = wb.sheets[e.titulo(b, "ANE")]
        tabla = _tabla(ws, "ANE")
        assert tabla is not None
        rango = tabla.api.ListColumns(1).DataBodyRange
        n = int(rango.Rows.Count)
        colores = {}
        for i in range(1, n + 1):
            celda = rango.Cells(i, 1)
            valor = celda.Value
            if valor in (None, ""):
                continue
            colores[main._cod(valor)] = int(celda.DisplayFormat.Font.Color)
        reglas = int(rango.FormatConditions.Count)
        del rango, tabla, ws
        return colores, reglas
    finally:
        if wb is not None:
            wb.close()


def test_formato_condicional_solo_blanquea_mayores_al_umbral():
    ruta = _exportado()
    app, pythoncom = main._abrir_excel()
    try:
        colores, reglas = _colores_codigos(app, ruta)
    finally:
        main._cerrar_excel(app, pythoncom)

    assert reglas == 1, f"reglas de formato condicional: {reglas}"
    assert set(colores) == {"2500", "3000", "1999", "2000", "1083"}, colores
    assert colores["2500"] == LOGOTIPOS_BLANCO, colores
    assert colores["3000"] == LOGOTIPOS_BLANCO, colores
    for visible in ("1999", "2000", "1083"):
        assert colores[visible] != LOGOTIPOS_BLANCO, \
            f"{visible} quedó oculto en blanco: {colores[visible]}"


# ═════════════════════════════════════════════════════════════════════
#  4) reexportar / reaplicar no duplica la regla y borra la vieja
# ═════════════════════════════════════════════════════════════════════
def test_regla_vieja_se_borra_y_no_se_duplica():
    ruta = _exportado()
    app, pythoncom = main._abrir_excel()
    wb = None

    def reglas_de(ws_libro):
        """Cuenta fresca: se vuelve a pedir la tabla y su columna 1."""
        columna = _tabla(ws_libro, "ANE").api.ListColumns(1).DataBodyRange
        return int(columna.FormatConditions.Count), columna

    try:
        wb = app.books.open(ruta)
        e = _estado()
        b = e.proyectos[0]
        ws = wb.sheets[e.titulo(b, "ANE")]
        tabla = _tabla(ws, "ANE")

        # Simula un archivo viejo: la regla =A7+0=A7 blanqueaba TODA la columna.
        rango = tabla.api.ListColumns(1).DataBodyRange
        direccion = str(rango.Cells(1, 1).Address).replace("$", "")
        assert direccion == "A7", direccion
        vieja = rango.FormatConditions.Add(Type=2, Formula1=f"={direccion}+0={direccion}")
        vieja.Font.Color = LOGOTIPOS_BLANCO

        cantidad, rango = reglas_de(ws)
        assert cantidad == 2, f"la inyección de la regla vieja falló: {cantidad}"
        # con la regla vieja, el 1999 SÍ quedaba oculto (era el bug)
        assert int(rango.Cells(1, 1).DisplayFormat.Font.Color) == LOGOTIPOS_BLANCO
        del rango

        main._xl_formato_condicional_codigo_numerico(tabla)

        cantidad, rango = reglas_de(ws)
        assert cantidad == 1, \
            f"no se borró la regla vieja o se duplicó la nueva: {cantidad}"
        nueva = str(rango.FormatConditions.Item(1).Formula1)
        assert f">{main.UMBRAL_CODIGO_ATE}" in nueva.replace(" ", ""), nueva

        # los códigos ≤ umbral vuelven a verse
        celdas = {main._cod(rango.Cells(i, 1).Value):
                  int(rango.Cells(i, 1).DisplayFormat.Font.Color)
                  for i in range(1, int(rango.Rows.Count) + 1)
                  if rango.Cells(i, 1).Value not in (None, "")}
        assert celdas["1999"] != LOGOTIPOS_BLANCO, celdas
        assert celdas["2000"] != LOGOTIPOS_BLANCO, celdas
        assert celdas["2500"] == LOGOTIPOS_BLANCO, celdas
        del rango

        # reaplicar dos veces más no acumula reglas
        main._xl_formato_condicional_codigo_numerico(tabla)
        main._xl_formato_condicional_codigo_numerico(tabla)
        cantidad, rango = reglas_de(ws)
        assert cantidad == 1, cantidad
        del rango, tabla, ws
    finally:
        try:
            if wb is not None:
                wb.close()          # sin guardar: el archivo de prueba no cambia
        finally:
            _cerrar(app, pythoncom)


def test_reexportar_produce_una_sola_regla():
    primera = _exportado()
    segunda = _exportar("ANEXO - ATE_T - copia.xlsx")
    try:
        app, pythoncom = main._abrir_excel()
        try:
            _, reglas_1 = _colores_codigos(app, primera)
            _, reglas_2 = _colores_codigos(app, segunda)
        finally:
            _cerrar(app, pythoncom)
        assert reglas_1 == 1, reglas_1
        assert reglas_2 == 1, f"la segunda exportación dejó {reglas_2} reglas"
    finally:
        try:
            os.remove(segunda)
        except OSError:
            pass


# ═════════════════════════════════════════════════════════════════════
#  5) roundtrip: el .xlsx reabierto conserva códigos, cantidades y modos
# ═════════════════════════════════════════════════════════════════════
def test_roundtrip_conserva_codigos_cantidades_y_hojas():
    ruta = _exportado()
    e = main.Estado.__new__(main.Estado)
    e.bd = []
    e.reset()
    main.cargar_proyecto_desde_excel(e, ruta)

    assert len(e.proyectos) == 1, [p["abrev"] for p in e.proyectos]
    b = e.proyectos[0]
    assert b["abrev"] == "ATE_T", b["abrev"]
    for hoja in ("ANE", "ANE_VAR", "VARIACION", "OC", "OTS", "ATE"):
        assert b["hojas"][hoja] is True, (hoja, b["hojas"])

    ane = e._cant_editable(b, "ANE")["MT"]
    av = e._cant_editable(b, "ANE_VAR")["MT"]
    esperado_ane = {c: Decimal(q[0]) for c, q in ESCENARIO.items() if q[0] != "0"}
    esperado_av = {c: Decimal(q[1]) for c, q in ESCENARIO.items() if q[1] != "0"}
    assert dict(ane) == esperado_ane, ane
    assert dict(av) == esperado_av, av

    # y el tipo calculado en la app reabierta coincide con el escenario
    for cod in ESCENARIO:
        dif = esperado_av.get(cod, Decimal(0)) - esperado_ane.get(cod, Decimal(0))
        assert e._tipo_documento(cod, dif) == ESPERADO_TIPO[cod], cod


# ═════════════════════════════════════════════════════════════════════
#  6) sin procesos de Excel huérfanos
# ═════════════════════════════════════════════════════════════════════
def test_no_quedan_procesos_excel_huerfanos():
    antes = _procesos_excel()
    _exportado()
    despues = _procesos_excel()
    assert despues <= antes, \
        f"quedaron EXCEL.EXE abiertos: {antes} → {despues}"
