# -*- coding: utf-8 -*-
"""Prepara la plantilla del nuevo anexo a partir del libro real.

Hace dos cosas, en una COPIA (el libro original no se toca):

1. Extiende la tabla CODIGOS_SICOA con las columnas que pide el diseño:
   MT, MT + BT, BT, G, 7.6 kV, 19.9 kV y Activo (G–M).
2. Opcionalmente (--limpiar), elimina las hojas de los proyectos de ejemplo
   y deja INICIO sin proyectos, conservando las 12 plantillas ocultas,
   las macros, los dibujos y las fórmulas.

Uso:
    python herramientas/preparar_plantilla.py
    python herramientas/preparar_plantilla.py --limpiar
    python herramientas/preparar_plantilla.py --solo-codigos --forzar

Antes de escribir crea una copia de seguridad en backups/.
"""

from __future__ import annotations

import argparse
import shutil
import sys
from datetime import date
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
ORIGEN_POR_DEFECTO = RAIZ / "VARIACIÓN FINAL - COPONES - 14.10.2026.xlsm"
DESTINO_POR_DEFECTO = RAIZ / "plantillas" / "Plantilla_Anexos.xlsm"

HOJA_CODIGOS = "CODIGOS"
HOJA_INICIO = "INICIO"
TABLA_CODIGOS = "CODIGOS_SICOA"
TABLA_CONFIG = "Configuration"
TABLA_RAMAL = "Configuration_Ramal"

ENCABEZADOS_NUEVOS = ["MT", "MT + BT", "BT", "G", "7.6 kV", "19.9 kV", "Activo"]
PREFIJOS = [
    "ANE. - ",
    "ANE. - VAR. - ",
    "VARIACIÓN - ",
    "OC - ",
    "OTS - ",
    "ATE - ",
]


def _backup(origen: Path) -> Path:
    carpeta = RAIZ / "backups"
    carpeta.mkdir(parents=True, exist_ok=True)
    destino = carpeta / f"Backup_{date.today():%d.%m.%Y}{origen.suffix}"
    if not destino.exists():
        shutil.copy2(origen, destino)
        print(f"Backup creado: {destino}")
    else:
        print(f"Backup ya existente: {destino}")
    return destino


def _tabla(hoja, nombre: str):
    for tabla in hoja.api.ListObjects:
        if tabla.Name == nombre:
            return tabla
    return None


def _extender_codigos(libro) -> None:
    hoja = libro.sheets[HOJA_CODIGOS]
    if hoja.api.ProtectContents:
        try:
            hoja.api.Unprotect()
        except Exception as exc:  # noqa: BLE001
            raise RuntimeError(
                "La hoja CODIGOS está protegida con contraseña; no se puede "
                "extender la tabla."
            ) from exc

    tabla = _tabla(hoja, TABLA_CODIGOS)
    if tabla is None:
        raise RuntimeError(f"No se encontró la tabla {TABLA_CODIGOS}.")

    fila_inicio = tabla.HeaderRowRange.Row
    fila_fin = fila_inicio + tabla.Range.Rows.Count - 1
    columnas_actuales = tabla.Range.Columns.Count

    if columnas_actuales >= 13 and str(hoja.range("G1").value or "").strip() == "MT":
        print("CODIGOS ya tiene las columnas nuevas; no se vuelve a extender.")
        return

    # Las notas que estaban en H5:H6 se mueven a P1:P2 para no pisarlas.
    notas = []
    for celda in ("H5", "H6"):
        valor = hoja.range(celda).value
        if valor not in (None, ""):
            notas.append(valor)
    hoja.range("H5:H6").clear_contents()
    for indice, nota in enumerate(notas, start=1):
        hoja.range(f"P{indice}").value = nota

    tabla.Resize(hoja.range(f"A{fila_inicio}:M{fila_fin}").api)
    hoja.range(f"G{fila_inicio}:M{fila_inicio}").value = [ENCABEZADOS_NUEVOS]

    primera = fila_inicio + 1
    hoja.range(f"G{primera}:I{fila_fin}").value = True
    hoja.range(f"J{primera}:J{fila_fin}").value = False
    hoja.range(f"K{primera}:L{fila_fin}").value = False
    hoja.range(f"M{primera}:M{fila_fin}").value = True

    # G (General) solo para los dos códigos fijos del libro.
    for fila in range(primera, fila_fin + 1):
        codigo = hoja.range(f"A{fila}").value
        if codigo in (1150, 1151):
            hoja.range(f"J{fila}").value = True

    print(
        f"CODIGOS extendido: {TABLA_CODIGOS} ahora es "
        f"A{fila_inicio}:M{fila_fin} (121 componentes)."
    )


def _abreviaturas(libro) -> list[str]:
    hoja = libro.sheets[HOJA_INICIO]
    abreviaturas: list[str] = []
    for nombre in (TABLA_CONFIG, TABLA_RAMAL):
        tabla = _tabla(hoja, nombre)
        if tabla is None:
            continue
        # La columna D (Abreviación) está en la posición 3 de la tabla.
        for fila in range(1, tabla.ListRows.Count + 1):
            try:
                valor = tabla.DataBodyRange(fila, 3).Value
            except Exception:  # noqa: BLE001
                continue
            if valor not in (None, ""):
                abreviaturas.append(str(valor).strip())
    return abreviaturas


def _limpiar_proyectos(libro) -> None:
    abreviaturas = _abreviaturas(libro)
    print(f"Proyectos/ramal detectados: {abreviaturas or 'ninguno'}")

    for abreviatura in abreviaturas:
        for prefijo in PREFIJOS:
            nombre = f"{prefijo}{abreviatura}"
            if nombre in [h.name for h in libro.sheets]:
                libro.sheets[nombre].delete()
                print(f"  Hoja eliminada: {nombre}")

    hoja = libro.sheets[HOJA_INICIO]
    for nombre in (TABLA_CONFIG, TABLA_RAMAL):
        tabla = _tabla(hoja, nombre)
        if tabla is not None and tabla.ListRows.Count:
            tabla.DataBodyRange.ClearContents()
            print(f"  Tabla limpiada: {nombre}")
    hoja.range("C7").value = 0
    print("INICIO: C7 = 0 (sin comunidades). Las 12 plantillas ocultas se conservan.")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--origen", type=Path, default=ORIGEN_POR_DEFECTO)
    parser.add_argument("--destino", type=Path, default=DESTINO_POR_DEFECTO)
    parser.add_argument("--forzar", action="store_true", help="Sobrescribe el destino.")
    grupo = parser.add_mutually_exclusive_group()
    grupo.add_argument(
        "--limpiar",
        action="store_true",
        help="Elimina las hojas de ejemplo (comportamiento por defecto).",
    )
    grupo.add_argument(
        "--solo-codigos",
        action="store_true",
        help="No elimina las hojas de los proyectos de ejemplo.",
    )
    args = parser.parse_args()

    if not args.origen.exists():
        print(f"No existe el libro de origen: {args.origen}")
        return 2
    if args.destino.exists() and not args.forzar:
        print(f"El destino ya existe: {args.destino}. Use --forzar para sobrescribirlo.")
        return 2

    _backup(args.origen)
    args.destino.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(args.origen, args.destino)
    print(f"Copia de trabajo: {args.destino}")

    import xlwings as xw  # import tardío para que --help no abra Excel

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    libro = None
    try:
        libro = app.books.open(str(args.destino))
        _extender_codigos(libro)
        if not args.solo_codigos:
            _limpiar_proyectos(libro)
        libro.save()
        print("Plantilla guardada correctamente.")
        return 0
    except Exception as exc:  # noqa: BLE001
        print(f"ERROR: {exc}")
        return 1
    finally:
        if libro is not None:
            libro.close()
        app.quit()


if __name__ == "__main__":
    raise SystemExit(main())
