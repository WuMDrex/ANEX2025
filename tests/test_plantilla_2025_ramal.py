# -*- coding: utf-8 -*-
"""Ramal con la plantilla nueva `Plantilla/ANEXO SICOA 2025.xlsm`.

La plantilla cambió: las hojas base se llaman igual (`ANE. - RAMAL`, etc.) pero
se quitó `INICIO`, la BD pasó de 218 a 159 códigos y los precios quedaron como
texto. Esta prueba exporta 2 comunidades + Ramal y comprueba que:

  * se crean las hojas del Ramal y ninguna `_BASE_*`;
  * la hoja del Ramal suma las cantidades de los dos proyectos (SUMIF real);
  * no queda ningún error de fórmula;
  * al reabrir se recuperan 2 proyectos y el Ramal.

Es una exportación real (xlwings/COM, ~30 s), marcada `excel`.
"""

import os
import sys

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RAIZ not in sys.path:
    sys.path.insert(0, RAIZ)

import pytest  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

import main  # noqa: E402

SALIDA = os.path.join(os.environ.get("LOCALAPPDATA") or RAIZ, "Temp",
                      "opencode", "plantilla2025")

_cache = {}


def _estado():
    if "estado" in _cache:
        return _cache["estado"]
    bd = main.cargar_bd_inicial()
    assert bd, "cargar_bd_inicial() devolvió una lista vacía"
    cod = next(c["codigo"] for c in bd if c["MT"] == 1)
    _cache["cod"] = cod

    e = main.Estado.__new__(main.Estado)
    e.bd = bd
    e.reset()

    p1 = e.proyectos[0]
    p1["abrev"] = "COM_A"
    p1["nombre"] = "COMUNIDAD A"
    p1["items"]["ANE"]["MT"] = [{"codigo": cod, "cantidad": "2", "redondeo": "ROUND"}]

    p2 = e.agregar_proyecto()
    p2["abrev"] = "COM_B"
    p2["nombre"] = "COMUNIDAD B"
    p2["items"]["ANE"]["MT"] = [{"codigo": cod, "cantidad": "3", "redondeo": "ROUND"}]

    # El Ramal refleja las hojas activas de los proyectos (tarea 39).
    e._sincronizar_ramal()
    assert e.ramal_visible(), "el Ramal debería estar activo con 2 proyectos"
    _cache["estado"] = e
    return e


def _exportado():
    if "ruta" not in _cache:
        os.makedirs(SALIDA, exist_ok=True)
        ruta = os.path.join(SALIDA, "ANEXO - RAMAL - plantilla2025.xlsx")
        if os.path.exists(ruta):
            os.remove(ruta)
        main.exportar_xlsx(_estado(), ruta)
        _cache["ruta"] = ruta
    return _cache["ruta"]


def _filas(ws, nombre):
    tabla = next(t for t in ws.tables.values() if t.name.upper() == nombre.upper())
    from openpyxl.utils import range_boundaries
    c1, f1, c2, f2 = range_boundaries(tabla.ref)
    return [tuple(ws.cell(row=f, column=c).value for c in range(c1, c2 + 1))
            for f in range(f1 + 1, f2 + 1)]


@pytest.mark.excel
def test_ramal_suma_los_proyectos_y_vuelve_al_reabrir():
    ruta = _exportado()
    e = _cache["estado"]
    cod = _cache["cod"]

    wb = load_workbook(ruta, data_only=True)
    try:
        assert not [n for n in wb.sheetnames if n.startswith("_BASE")], wb.sheetnames
        for base in (e.proyectos[0], e.proyectos[1], e.ramal):
            titulo = e.titulo(base, "ANE")
            assert titulo in wb.sheetnames, (titulo, wb.sheetnames)

        # errores de fórmula en todo el libro
        males = []
        for ws in wb.worksheets:
            for fila in ws.iter_rows():
                for cel in fila:
                    if isinstance(cel.value, str) and cel.value.strip() in (
                            "#N/A", "#REF!", "#VALUE!", "#NAME?", "#DIV/0!"):
                        males.append((ws.title, cel.coordinate, cel.value))
        assert not males, males[:8]

        # el Ramal suma 2 + 3 del mismo código (SUMIF resuelto por Excel)
        ws_ram = wb[e.titulo(e.ramal, "ANE")]
        nombre = main._xl_nombre_tabla("ANE", main._xl_token(e.ramal["abrev"]), "MT")
        filas = _filas(ws_ram, nombre)
        fila = next((f for f in filas if str(f[0]).strip() == cod), None)
        assert fila is not None, (cod, nombre, filas)
        assert float(fila[5]) == 5.0, fila       # cantidad 2 + 3
        precio = float(fila[4])
        assert abs(float(fila[6]) - precio * 5) < 0.02, fila
        _cache["ramal_cant"] = float(fila[5])
    finally:
        wb.close()

    # roundtrip: 2 proyectos + Ramal reconocido
    e2 = main.Estado.__new__(main.Estado)
    e2.bd = []
    e2.reset()
    main.cargar_proyecto_desde_excel(e2, ruta)
    assert len(e2.proyectos) == 2, [p["abrev"] for p in e2.proyectos]
    assert sorted(p["abrev"] for p in e2.proyectos) == ["COM_A", "COM_B"], \
        [p["abrev"] for p in e2.proyectos]
    assert e2.ramal_presente, "el Ramal no se reconoció al reabrir"
    assert e2.ramal["abrev"] == "RAMAL", e2.ramal["abrev"]
    assert e2.ramal["hojas"]["ANE"], e2.ramal["hojas"]
