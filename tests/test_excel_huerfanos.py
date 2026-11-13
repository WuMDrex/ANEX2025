# -*- coding: utf-8 -*-
"""Verifica que no queden procesos EXCEL.EXE huérfanos tras las pruebas.

Usa `tasklist` (Windows) desde Python porque el shell del agente solo ejecuta
`py_compile` y `pytest`. No abre Excel: solo consulta la lista de procesos.
"""

import os
import subprocess
import sys

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RAIZ not in sys.path:
    sys.path.insert(0, RAIZ)


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


def test_sin_excel_huerfano():
    filas = _procesos_excel()
    print("   procesos EXCEL.EXE:", len(filas))
    for f in filas:
        print("   ->", f[:2])
    assert not filas, f"quedaron {len(filas)} EXCEL.EXE abiertos: {filas}"
