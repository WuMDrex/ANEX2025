# -*- coding: utf-8 -*-
"""
Mutaciones: comprueban que las pruebas de `test_ramal_firma.py` DISCRIMINAN.

No modifica `main.py`: reemplaza funciones EN MEMORIA (`main._xl_firma_ramal`,
`main._xl_indice_ramal`) por versiones "viejas" y verifica que las pruebas
correspondientes FALLAN. Si una prueba pasa con la versión vieja, esa prueba
no sirve (pasaría por casualidad).

Se ejecuta con el shim de pruebas del proyecto (NO es pytest real):
    cd tests && python -m pytest mutaciones_ramal.py
"""

import os
import shutil
import sys
import traceback

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RAIZ not in sys.path:
    sys.path.insert(0, RAIZ)
CARPETA = os.path.dirname(os.path.abspath(__file__))
if CARPETA not in sys.path:
    sys.path.insert(0, CARPETA)

try:                                            # consola Windows = cp1252
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import main                                     # noqa: E402

import test_ramal_firma as T                    # noqa: E402


# ── lectura de celdas, reutilizada por las mutaciones ───────────────
def _celdas(grupo):
    """Texto de cada celda de 'Cantidad a ejecutar' en ANE / ANE-VAR."""
    textos = []
    for hoja in main.EDITABLES:
        ws = grupo["hojas"].get(hoja)
        if ws is None:
            continue
        for tabla in ws.tables:
            try:
                col = main._xl_col_cantidad(tabla)
                cuerpo = tabla.api.ListColumns(col).DataBodyRange
            except Exception:
                continue
            if cuerpo is None:
                continue
            try:
                formulas = main._xl_valores(cuerpo.Formula)
            except Exception:
                continue
            for fila in formulas:
                textos.append(str((fila[0] if fila else "") or "").upper())
    return textos


# ── versiones "viejas" de la firma ──────────────────────────────────
def firma_vieja_todas_formulas(grupo):
    """Regla anterior: 'si todas las celdas son fórmula, es Ramal' → 1."""
    textos = _celdas(grupo)
    if textos and all(t.startswith("=") for t in textos):
        return 1
    return 0


def firma_vieja_alguna_sumif(grupo):
    """Regla laxa: 'con una sola SUMIF alcanza' → 1000."""
    textos = _celdas(grupo)
    if any("SUMIF(" in t or "SUMAR.SI(" in t for t in textos):
        return 1000
    return 0


def indice_viejo_por_nombre(tokens, firmas):
    """Regla anterior: el Ramal se reconoce SOLO por el nombre 'RAMAL'."""
    for i, token in enumerate(tokens):
        if main._xl_token(token).upper() == "RAMAL":
            return i
    return None


def _corre(nombre):
    """Devuelve (discrimina, detalle) al ejecutar la prueba con la mutación."""
    try:
        getattr(T, nombre)()
    except AssertionError as exc:
        return True, str(exc).splitlines()[0][:160]
    except Exception as exc:                     # pragma: no cover
        return False, f"error inesperado {type(exc).__name__}: {exc}"
    return False, "la prueba pasó con la versión vieja"


def test_mutaciones_discriminan():
    resultados = []
    firma_original = main._xl_firma_ramal
    indice_original = main._xl_indice_ramal
    try:
        main._xl_firma_ramal = firma_vieja_todas_formulas
        resultados.append((
            "'todas fórmulas' → 1",
            "test_firma_formulas_sin_sumif",
            _corre("test_firma_formulas_sin_sumif")))
        resultados.append((
            "'todas fórmulas' → 1",
            "test_firma_sumif_en_todas_las_filas",
            _corre("test_firma_sumif_en_todas_las_filas")))
    finally:
        main._xl_firma_ramal = firma_original

    try:
        main._xl_firma_ramal = firma_vieja_alguna_sumif
        for prueba in ("test_firma_una_sola_sumif_entre_numeros_no_basta",
                       "test_firma_mitad_exacta_no_basta"):
            resultados.append(("'alguna SUMIF' → 1000", prueba, _corre(prueba)))
    finally:
        main._xl_firma_ramal = firma_original

    try:
        main._xl_indice_ramal = indice_viejo_por_nombre
        resultados.append((
            "Ramal reconocido solo por nombre",
            "test_exportar_y_reimportar_ramal_renombrado",
            _corre("test_exportar_y_reimportar_ramal_renombrado")))
    finally:
        main._xl_indice_ramal = indice_original

    malas = 0
    print("\n" + "=" * 70)
    for mutacion, prueba, (discrimina, detalle) in resultados:
        estado = "OK  (discrimina)" if discrimina else "MALA (no discrimina)"
        print(f"{estado:<22} {prueba}\n"
              f"{'':22} mutación: {mutacion}\n"
              f"{'':22} detalle : {detalle}")
        if not discrimina:
            malas += 1
    print("=" * 70)
    print(f"{len(resultados) - malas}/{len(resultados)} mutaciones detectadas")
    assert malas == 0, f"{malas} prueba(s) no discriminan: {resultados}"


# ── sonda informativa (no falla: describe el comportamiento) ────────
def test_sonda_ramal_renombrado_sin_filas():
    print("\n---- SONDA ----")
    # 1) Ramal renombrado SIN filas: la firma no tiene nada que leer.
    try:
        estado = T._estado_prueba(["COM1", "COM2"], "RAMAL2", codigos=())
        ruta, carpeta = T._exportar(estado, "ANEXO - SIN DATOS.xlsx")
        try:
            e2 = T._reimportar(ruta)
            print(f"SONDA ramal renombrado sin filas → ramal_presente="
                  f"{e2.ramal_presente}, ramal={e2.ramal['abrev']!r}, "
                  f"proyectos={[p['abrev'] for p in e2.proyectos]}")
        finally:
            shutil.rmtree(carpeta, ignore_errors=True)
    except Exception as exc:
        print(f"SONDA ramal renombrado sin filas → ERROR {type(exc).__name__}: {exc}")
        traceback.print_exc()


if __name__ == "__main__":                      # pragma: no cover
    test_mutaciones_discriminan()
    test_sonda_ramal_renombrado_sin_filas()
