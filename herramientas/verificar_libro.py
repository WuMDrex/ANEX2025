# -*- coding: utf-8 -*-
"""Verifica que el lector reproduzca exactamente los totales del libro.

Uso:
    python herramientas/verificar_libro.py "ruta\\al\\libro.xlsm"

Comprueba, para cada proyecto y para el Ramal:
  - cantidad de componentes leídos,
  - suma de las 4 secciones (MT, MT+BT, BT, G),
  - total G377 del libro,
  - diferencia.

Si la diferencia es 0.0000 en todos, la lectura es fiel al libro.
"""

from __future__ import annotations

import sys
from decimal import Decimal
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from servicios.excel_lector import leer_anexos, leer_todo  # noqa: E402


def verificar(ruta: str) -> int:
    libro = Path(ruta)
    if not libro.exists():
        print(f"No existe el libro: {libro}")
        return 2

    config = leer_todo(libro)
    anexos = leer_anexos(libro)

    print(f"LIBRO: {libro.name}")
    print(f"Componentes en CODIGOS: {len(config.componentes)}")
    print(f"Proyectos: {len(config.proyectos)}  |  Ramal: "
          f"{config.ramal.abreviacion if config.ramal else 'ninguno'}\n")

    print(f"{'ABREV.':<12}{'HOJA':<30}{'SUMA 4 SECCIONES':>20}{'TOTAL G377':>18}{'DIF.':>10}")
    print("-" * 92)

    fallos = 0
    for abreviatura, datos in anexos.items():
        suma = sum(
            (fila.total for filas in datos["secciones"].values() for fila in filas),
            Decimal("0"),
        )
        total = datos["total"]
        diferencia = suma - total
        if abs(diferencia) >= Decimal("0.01"):
            fallos += 1
        print(
            f"{abreviatura:<12}{datos['hoja']:<30}{suma:>20,.4f}{total:>18,.4f}"
            f"{diferencia:>10.4f}"
        )

    print()
    if fallos:
        print(f"RESULTADO: {fallos} diferencia(s) encontrada(s). Revisar el lector.")
        return 1
    print("RESULTADO: OK. La lectura cuadra con el libro en todos los anexos.")
    return 0


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(__doc__)
        raise SystemExit(2)
    raise SystemExit(verificar(sys.argv[1]))
