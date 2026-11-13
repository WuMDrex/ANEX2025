# -*- coding: utf-8 -*-
"""
Verificación del TIPO DE DATO del código en la exportación.

Contexto del arreglo: `ANE` escribía el código siempre como TEXTO mientras
`VARIACIÓN`/`OC`/`OTS`/`ATE` lo escribían como NÚMERO, así que el BUSCARX de
descripción/unidad/precio y el de "Cantidad a ejecutar" podían devolver 0 o
#N/A, y el SUMAR.SI del Ramal podía sumar 0. Ahora `_xl_codigo_celda` decide:
número si el código es numérico ASCII, texto si es alfanumérico.

Escenario (2 proyectos + Ramal, todo en la línea MT):

    código  P1 ANE  P1 ANE-VAR  P2 ANE   precio   tipo esperado
    1999       4        0          3     100.00   OC
    2000       1        5          -     200.00   OTS
    2500       5        0          2     300.00   ATE   (> 2000)
    P001       2        1          1      50.00   OTS   (texto)

Requiere Excel instalado → @pytest.mark.excel. Nunca modifica la plantilla ni
main.py: `exportar_xlsx` copia el machote a un temporal.
"""

import os
import subprocess
import sys
import tempfile

import pytest

import main

pytestmark = pytest.mark.excel

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BLANCO = 16777215

P1 = {"1999": ("4", "0"), "2000": ("1", "5"), "2500": ("5", "0"),
      "P001": ("2", "1")}
P2 = {"1999": ("3", "0"), "2500": ("2", "0"), "P001": ("1", "0")}
PRECIOS = {"1999": "100.00", "2000": "200.00", "2500": "300.00",
           "P001": "50.00"}

_cache = {}


# ── utilidades ───────────────────────────────────────────────────────
def _carpeta():
    preferida = r"C:\Users\deivi\AppData\Local\Temp\opencode"
    base = preferida if os.path.isdir(preferida) else tempfile.gettempdir()
    ruta = os.path.join(base, "codigo_tipo")
    os.makedirs(ruta, exist_ok=True)
    return ruta


def _procesos_excel():
    salida = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/NH"],
        capture_output=True, text=True, errors="replace")
    return salida.stdout.upper().count("EXCEL.EXE")


def _cerrar(app, pythoncom):
    """Cierra Excel liberando ANTES los proxies COM (evita 0x800706ba)."""
    import gc
    gc.collect()
    main._cerrar_excel(app, pythoncom)
    gc.collect()


def _estado():
    if "estado" in _cache:
        return _cache["estado"]
    e = main.Estado.__new__(main.Estado)
    e.bd = [main.componente_nuevo(c, f"COMPONENTE DE PRUEBA {c}", "U",
                                  PRECIOS[c]) for c in PRECIOS]
    for c in e.bd:
        c["MT"] = 1
    e.reset()
    for i, (esc, abrev, nombre) in enumerate(
            ((P1, "CT1", "PRUEBA CODIGO"), (P2, "CT2", "PRUEBA DOS"))):
        b = e.proyectos[0] if i == 0 else e._base(abrev, main.PALETA[1])
        if i > 0:
            e.proyectos.append(b)
        b["abrev"] = abrev
        b["nombre"] = nombre
        b["hojas"] = {h: True for h in main.HOJAS}
        b["visibles"] = list(main.HOJAS)
        b["items"] = {"ANE": main._vacio(), "ANE_VAR": main._vacio()}
        for cod, (qa, qv) in esc.items():
            if qa != "0":
                b["items"]["ANE"]["MT"].append(
                    {"codigo": cod, "cantidad": qa, "redondeo": "ROUND"})
            if qv != "0":
                b["items"]["ANE_VAR"]["MT"].append(
                    {"codigo": cod, "cantidad": qv, "redondeo": "ROUND"})
    e.ramal["abrev"] = "RAMAL"
    e.ramal["nombre"] = "RAMAL"
    e.ramal["hojas"] = {h: True for h in main.HOJAS}
    e.ramal["visibles"] = list(main.HOJAS)
    e.ramal["items"] = {"ANE": main._vacio(), "ANE_VAR": main._vacio()}
    e.cambios = True
    _cache["estado"] = e
    return e


def _exportado():
    if "ruta" not in _cache:
        ruta = os.path.join(_carpeta(), "ANEXO - CT1 - codigo_tipo.xlsx")
        if os.path.exists(ruta):
            os.remove(ruta)
        main.exportar_xlsx(_estado(), ruta)
        _cache["ruta"] = ruta
    return _cache["ruta"]


def _tabla_mt(app, wb, base, hoja):
    ws = wb.sheets[main.Estado.titulo(_estado(), base, hoja)]
    tabla = main._xl_tabla(ws, main._xl_nombre_tabla(
        hoja, main._xl_token(base["abrev"]), "MT"))
    assert tabla is not None, f"no existe la tabla MT de {hoja}"
    return ws, tabla


def _filas(ws, tabla):
    """[(código, tipo_python, fila)] de la columna A del cuerpo de la tabla."""
    enc = main._xl_encabezado(tabla)
    n = int(tabla.api.ListRows.Count)
    out = []
    for i in range(n):
        fila = enc + 1 + i
        val = ws.range(f"A{fila}").value
        if val in (None, ""):
            continue
        out.append((main._cod(val), type(val).__name__, fila, val))
    return out


# ═════════════════════════════════════════════════════════════════════
#  0) lógica pura: _xl_codigo_celda
# ═════════════════════════════════════════════════════════════════════
@pytest.mark.parametrize("entrada,esperado,tipo", [
    ("1999", 1999, int),
    ("2000", 2000, int),
    ("2500", 2500, int),
    (3000, 3000, int),
    (" 2500 ", 2500, int),
    ("00123", 123, int),
    ("P001", "P001", str),
    ("E051", "E051", str),
    ("12A", "12A", str),
    ("", "", str),
])
def test_codigo_celda_numerico_a_numero_alfanumerico_a_texto(entrada, esperado,
                                                             tipo):
    obtenido = main._xl_codigo_celda(entrada)
    assert obtenido == esperado, (entrada, obtenido, esperado)
    assert type(obtenido) is tipo, (entrada, type(obtenido), tipo)


def test_codigo_celda_no_convierte_digitos_unicode():
    """isascii(): un '²' no debe volverse número (Excel no lo leería igual)."""
    assert main._xl_codigo_celda("²") == "²"
    assert isinstance(main._xl_codigo_celda("²"), str)


# ═════════════════════════════════════════════════════════════════════
#  1) el verificador preparado por el especialista
# ═════════════════════════════════════════════════════════════════════
def test_verificador_del_especialista_dice_todo_pasa():
    script = os.path.join(RAIZ, "herramientas", "_verifica_codigo_tmp.py")
    if not os.path.exists(script):
        pytest.skip("no existe herramientas/_verifica_codigo_tmp.py")
    antes = _procesos_excel()
    salida = subprocess.run([sys.executable, script], cwd=RAIZ,
                            capture_output=True, text=True, errors="replace")
    print(salida.stdout)
    print(salida.stderr)
    # DEFECTO CONOCIDO del propio script (no de main.py): hardcodea
    # C:\Users\devi\... y en esta máquina el perfil es C:\Users\deivi, así que
    # muere en os.makedirs con PermissionError antes de probar nada.
    assert "RESULTADO: TODO PASA" in salida.stdout, (
        f"el verificador no llegó al final. returncode={salida.returncode}\n"
        f"STDOUT:\n{salida.stdout}\nSTDERR:\n{salida.stderr}")
    assert salida.returncode == 0, salida.returncode
    assert _procesos_excel() <= antes, "el verificador dejó Excel abierto"


# ═════════════════════════════════════════════════════════════════════
#  2) ANE: tipo REAL de la celda + BUSCARX de descripción/unidad/precio
# ═════════════════════════════════════════════════════════════════════
def test_ane_codigo_se_escribe_con_el_tipo_correcto_y_xlookup_resuelve():
    ruta = _exportado()
    app, pythoncom = main._abrir_excel()
    wb = None
    try:
        wb = app.books.open(ruta)
        app.calculate()
        b = _estado().proyectos[0]
        ws, tabla = _tabla_mt(app, wb, b, "ANE")

        leidos = _filas(ws, tabla)
        assert sorted(c for c, _, _, _ in leidos) == sorted(PRECIOS), leidos

        tipos = {cod: t for cod, t, _, _ in leidos}
        assert tipos["1999"] == "float", tipos
        assert tipos["2000"] == "float", tipos
        assert tipos["2500"] == "float", tipos
        assert tipos["P001"] == "str", tipos

        for cod, _, fila, _ in leidos:
            desc = ws.range(f"C{fila}").value
            uni = ws.range(f"D{fila}").value
            prec = ws.range(f"E{fila}").value
            cant = ws.range(f"F{fila}").value
            assert not (isinstance(desc, str) and desc.startswith("#")), \
                (cod, desc)
            assert desc == f"COMPONENTE DE PRUEBA {cod}", (cod, desc)
            assert uni == "U", (cod, uni)
            assert float(prec) == float(PRECIOS[cod]), (cod, prec)
            qa = P1[cod][0]
            assert float(cant) == float(qa), (cod, cant)
    finally:
        try:
            if wb is not None:
                wb.close()
        finally:
            _cerrar(app, pythoncom)


# ═════════════════════════════════════════════════════════════════════
#  3) VARIACIÓN: BUSCARX de "Cantidad a ejecutar" y cuadre de totales
# ═════════════════════════════════════════════════════════════════════
def test_variacion_cantidades_y_totales_cuadran_sin_ceros_espurios():
    ruta = _exportado()
    app, pythoncom = main._abrir_excel()
    wb = None
    try:
        wb = app.books.open(ruta)
        app.calculate()
        b = _estado().proyectos[0]
        ws, tabla = _tabla_mt(app, wb, b, "VARIACION")
        leidos = _filas(ws, tabla)
        assert sorted(c for c, _, _, _ in leidos) == sorted(PRECIOS), leidos

        suma_f = suma_i = suma_m = suma_n = suma_o = 0.0
        # P001 es texto → no es ATE aunque sea "numérico"; su variación es
        # 1 - 2 = -1 (baja), así que es OC.
        esperado_tipo = {"1999": "OC", "2000": "OTS", "2500": "ATE",
                         "P001": "OC"}
        for cod, tipo_celda, fila, _ in leidos:
            assert tipo_celda == ("str" if cod == "P001" else "float"), \
                (cod, tipo_celda)
            cant_ane = ws.range(f"D{fila}").value
            cant_av = ws.range(f"G{fila}").value
            tot_f = ws.range(f"F{fila}").value
            tot_i = ws.range(f"I{fila}").value
            tipo_doc = ws.range(f"J{fila}").value
            m = ws.range(f"M{fila}").value
            n = ws.range(f"N{fila}").value
            o = ws.range(f"O{fila}").value

            qa = float(P1[cod][0])
            qv = float(P1[cod][1])
            # lo que se rompía con el choque de tipos: esto daba 0
            assert float(cant_ane or 0) == qa, (cod, "D", cant_ane, qa)
            assert float(cant_av or 0) == qv, (cod, "G", cant_av, qv)
            assert float(tot_f or 0) == qa * float(PRECIOS[cod]), (cod, tot_f)
            assert float(tot_i or 0) == qv * float(PRECIOS[cod]), (cod, tot_i)
            assert tipo_doc == esperado_tipo[cod], (cod, tipo_doc)

            suma_f += float(tot_f or 0)
            suma_i += float(tot_i or 0)
            suma_m += float(m or 0)
            suma_n += float(n or 0)
            suma_o += float(o or 0)

        assert round(suma_f, 2) == 2200.00, suma_f
        assert round(suma_i, 2) == 1050.00, suma_i
        # (Modificado - Contratado) == OC + OTS + ATE  → no descuadra
        assert round((suma_i - suma_f), 2) == round(suma_m + suma_n + suma_o, 2), \
            (suma_i, suma_f, suma_m, suma_n, suma_o)
        assert round(suma_m, 2) == -450.00, suma_m    # 1999 (-400) + P001 (-50)
        assert round(suma_n, 2) == 800.00, suma_n
        assert round(suma_o, 2) == -1500.00, suma_o
    finally:
        try:
            if wb is not None:
                wb.close()
        finally:
            _cerrar(app, pythoncom)


# ═════════════════════════════════════════════════════════════════════
#  4) Ramal: SUMAR.SI de los dos proyectos
# ═════════════════════════════════════════════════════════════════════
def test_ramal_sumif_suma_los_dos_proyectos_y_no_da_cero():
    ruta = _exportado()
    app, pythoncom = main._abrir_excel()
    wb = None
    try:
        wb = app.books.open(ruta)
        app.calculate()
        e = _estado()
        assert e.ramal_visible(), "el escenario no generó Ramal"
        ws, tabla = _tabla_mt(app, wb, e.ramal, "ANE")
        esperado = {cod: (float(P1[cod][0]) if cod in P1 else 0.0)
                    + (float(P2[cod][0]) if cod in P2 else 0.0)
                    for cod in PRECIOS}
        obtenido = {}
        for cod, tipo_celda, fila, _ in _filas(ws, tabla):
            obtenido[cod] = float(ws.range(f"F{fila}").value or 0)
            assert tipo_celda == ("str" if cod == "P001" else "float"), \
                (cod, tipo_celda)
        assert obtenido == esperado, obtenido
        assert all(v > 0 for v in obtenido.values()), obtenido

        # ANE-VAR del Ramal: solo aparecen los códigos con fila (2000 y P001);
        # 1999 y 2500 no tienen cantidad en ninguna ANE-VAR, así que no existen
        # como fila.
        ws2, tabla2 = _tabla_mt(app, wb, e.ramal, "ANE_VAR")
        av = {cod: float(ws2.range(f"F{fila}").value or 0)
              for cod, _, fila, _ in _filas(ws2, tabla2)}
        assert av == {"2000": 5.0, "P001": 1.0}, av
    finally:
        try:
            if wb is not None:
                wb.close()
        finally:
            _cerrar(app, pythoncom)


# ═════════════════════════════════════════════════════════════════════
#  5) formato condicional: 1 regla y blanco solo para > 2000
# ═════════════════════════════════════════════════════════════════════
def test_formato_condicional_una_regla_y_blanco_solo_mayores_al_umbral():
    ruta = _exportado()
    app, pythoncom = main._abrir_excel()
    wb = None
    try:
        wb = app.books.open(ruta)
        app.calculate()
        b = _estado().proyectos[0]
        ws, tabla = _tabla_mt(app, wb, b, "ANE")
        rango = tabla.api.ListColumns(1).DataBodyRange
        assert int(rango.FormatConditions.Count) == 1, \
            int(rango.FormatConditions.Count)
        enc = main._xl_encabezado(tabla)
        colores = {}
        for i in range(int(tabla.api.ListRows.Count)):
            celda = ws.range(f"A{enc + 1 + i}")
            if celda.value in (None, ""):
                continue
            colores[main._cod(celda.value)] = int(
                celda.api.DisplayFormat.Font.Color)
        assert colores["2500"] == BLANCO, colores
        for visible in ("1999", "2000", "P001"):
            assert colores[visible] != BLANCO, (visible, colores)
    finally:
        try:
            if wb is not None:
                wb.close()
        finally:
            _cerrar(app, pythoncom)


# ═════════════════════════════════════════════════════════════════════
#  6) roundtrip: códigos exactos, sin "2500.0"
# ═════════════════════════════════════════════════════════════════════
def test_roundtrip_conserva_codigos_exactos_cantidades_hojas_y_redondeos():
    ruta = _exportado()
    e = main.Estado.__new__(main.Estado)
    e.bd = []
    e.reset()
    main.cargar_proyecto_desde_excel(e, ruta)

    assert len(e.proyectos) == 2, [p["abrev"] for p in e.proyectos]
    assert e.ramal_presente, "no se reconoció el Ramal"
    assert e.ramal["abrev"] == "RAMAL", e.ramal["abrev"]

    b = next(p for p in e.proyectos if p["abrev"] == "CT1")
    for hoja in ("ANE", "ANE_VAR", "VARIACION", "OC", "OTS", "ATE"):
        assert b["hojas"][hoja] is True, (hoja, b["hojas"])

    ane = [it["codigo"] for it in b["items"]["ANE"]["MT"]]
    av = [it["codigo"] for it in b["items"]["ANE_VAR"]["MT"]]
    assert sorted(ane) == sorted(["1999", "2000", "2500", "P001"]), ane
    assert sorted(av) == sorted(["2000", "P001"]), av
    assert not [c for c in ane + av if str(c).endswith(".0")], ane + av

    cant_ane = {it["codigo"]: main.D(it["cantidad"])
                for it in b["items"]["ANE"]["MT"]}
    cant_av = {it["codigo"]: main.D(it["cantidad"])
               for it in b["items"]["ANE_VAR"]["MT"]}
    assert cant_ane == {c: main.D(q[0]) for c, q in P1.items()
                        if q[0] != "0"}, cant_ane
    assert cant_av == {c: main.D(q[1]) for c, q in P1.items()
                       if q[1] != "0"}, cant_av

    modos = {it["redondeo"] for it in b["items"]["ANE"]["MT"]}
    assert modos == {"ROUND"}, modos


# ═════════════════════════════════════════════════════════════════════
#  7) la alerta "No Cuadra" de VARIACIÓN no debe aparecer (lógica pura)
# ═════════════════════════════════════════════════════════════════════
def test_variacion_no_descuadra_en_python_asi_que_la_alerta_no_aparece():
    """
    Mismo cálculo que `Principal.refrescar_vista` para el aviso rojo:
        diferencia = (total_i - total_f) - (OC + OTS + ATE)
    Si da 0, el widget `alerta_var` se oculta. Se calcula con Estado, sin UI.
    """
    e = _estado()
    b = e.proyectos[0]
    filas = e.filas(b, "VARIACION")
    total_f = sum((f["total_ane"] for lista in filas.values() for f in lista),
                  main.D(0))
    total_i = sum((f["total_av"] for lista in filas.values() for f in lista),
                  main.D(0))
    total_doc = sum((f["monto_doc"] for lista in filas.values() for f in lista),
                    main.D(0))
    diferencia = (total_i - total_f - total_doc).quantize(main.D("0.01"))
    assert diferencia == 0, (total_i, total_f, total_doc, diferencia)
    assert total_f == main.D("2200.00"), total_f
    assert total_i == main.D("1050.00"), total_i


# ═════════════════════════════════════════════════════════════════════
#  8) sin procesos de Excel huérfanos
# ═════════════════════════════════════════════════════════════════════
def test_no_quedan_procesos_excel_huerfanos():
    antes = _procesos_excel()
    _exportado()
    despues = _procesos_excel()
    assert despues <= antes, f"quedaron EXCEL.EXE abiertos: {antes} → {despues}"
