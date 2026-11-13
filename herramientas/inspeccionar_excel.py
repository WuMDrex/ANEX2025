# -*- coding: utf-8 -*-
"""Inspecciona un libro .xlsx/.xlsm y reporta hojas, tablas y encabezados.

Uso:
    python herramientas/inspeccionar_excel.py "ruta\al\libro.xlsm"
"""

from __future__ import annotations

import sys
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("Falta openpyxl. Instale con:  python -m pip install openpyxl")


def inspeccionar(ruta: str) -> None:
    libro = Path(ruta)
    if not libro.exists():
        sys.exit(f"No existe el archivo: {libro}")

    wb = load_workbook(libro, data_only=False, keep_vba=libro.suffix.lower() == ".xlsm")

    print(f"LIBRO: {libro.name}")
    print(f"HOJAS: {len(wb.worksheets)}\n")

    for numero, ws in enumerate(wb.worksheets, start=1):
        tablas = list(ws.tables.values())
        print(f"{numero:>2}. {ws.title!r}")
        print(f"    estado={ws.sheet_state}  dimensiones={ws.dimensions}  tablas={len(tablas)}")
        for tabla in tablas:
            columnas = [c.name for c in tabla.tableColumns]
            print(f"    - {tabla.name} {tabla.ref}")
            print(f"      columnas: {columnas}")
        print()

    print("NOMBRES DEFINIDOS")
    for nombre, definido in wb.defined_names.items():
        print(f"  {nombre} -> {definido.attr_text}")

    wb.close()


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    inspeccionar(sys.argv[1])
