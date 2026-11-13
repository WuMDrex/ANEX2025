# -*- coding: utf-8 -*-
"""
Pruebas de MUTACIÓN del modo «0.00 SR».

No tocan `main.py`: reemplazan las funciones EN MEMORIA con `monkeypatch` y
comprueban que las expectativas de `tests/test_modo_sr.py` SE ROMPEN. Así queda
demostrado que esas pruebas fallan si el bug existe y pasan si no.
"""

import pytest

import main

EXPR = "[@[Cantidad]]*[@[Precio]]"


# ── Mutantes ────────────────────────────────────────────────────────────────

def _mut_sin_sr(expr, modo):
    """Bug viejo: el modo SR se exporta como ROUND (no se respeta)."""
    funcion = modo if modo in main.MODOS_VALIDOS else main.MODO_DEFECTO
    if funcion == "SR":
        funcion = "ROUND"
    return f"={funcion}({expr},2)"


def _mut_sr_como_funcion(expr, modo):
    """Bug: se escribe `=SR(...)`, que Excel no reconoce como función."""
    if modo == "SR":
        return f"=SR({expr},2)"
    return f"={modo}({expr},2)"


def _mut_round_primero(formula):
    """Bug de orden: ROUND se comprueba antes que ROUNDUP/ROUNDDOWN."""
    s = str(formula or "").upper()
    if "ROUND" in s:
        return "ROUND"
    if "ROUNDUP" in s:
        return "ROUNDUP"
    if "ROUNDDOWN" in s:
        return "ROUNDDOWN"
    if "*" in s:
        return "SR"
    return main.MODO_DEFECTO


def _mut_asterisco_primero(formula):
    """Bug de orden: '*' se evalúa antes que ROUND -> toda fila ROUND sería SR."""
    s = str(formula or "").upper()
    if "*" in s:
        return "SR"
    if "ROUNDUP" in s:
        return "ROUNDUP"
    if "ROUNDDOWN" in s:
        return "ROUNDDOWN"
    if "ROUND" in s:
        return "ROUND"
    return main.MODO_DEFECTO


# ── Cada mutante debe romper su prueba correspondiente ──────────────────────

def test_mutante_sr_exportado_como_round_falla_la_prueba(monkeypatch):
    """M1: si SR se exporta con ROUND, la prueba de SR debe fallar."""
    monkeypatch.setattr(main, "_xl_formula_celda_modo", _mut_sin_sr)

    with pytest.raises(AssertionError):
        assert main._xl_formula_celda_modo(EXPR, "SR") == f"=({EXPR})"
    # y además el roundtrip deja de devolver SR
    with pytest.raises(AssertionError):
        assert main._xl_modo_desde_formula(
            main._xl_formula_celda_modo(EXPR, "SR")) == "SR"


def test_mutante_sr_como_funcion_falla_la_prueba(monkeypatch):
    """M4: `=SR(...)` no puede pasar ni la igualdad ni la ausencia de ROUND."""
    monkeypatch.setattr(main, "_xl_formula_celda_modo", _mut_sr_como_funcion)

    formula = main._xl_formula_celda_modo(EXPR, "SR")
    assert "=SR(" in formula.upper()            # el mutante sí lo escribe
    with pytest.raises(AssertionError):
        assert "=SR(" not in formula.upper()
    with pytest.raises(AssertionError):
        assert formula == f"=({EXPR})"


def test_mutante_round_primero_rompe_roundup_y_rounddown(monkeypatch):
    """M2: con ROUND primero, ROUNDUP se lee como ROUND (regresión del orden)."""
    monkeypatch.setattr(main, "_xl_modo_desde_formula", _mut_round_primero)

    with pytest.raises(AssertionError):
        assert main._xl_modo_desde_formula("=ROUNDUP([@[A]]*[@[B]],2)") == "ROUNDUP"
    with pytest.raises(AssertionError):
        assert main._xl_modo_desde_formula("=ROUNDDOWN([@[A]]*[@[B]],2)") == "ROUNDDOWN"


def test_mutante_asterisco_primero_rompe_round_con_multiplicacion(monkeypatch):
    """M3: si '*' se mira antes que ROUND, una fila Redondear se leería como SR."""
    monkeypatch.setattr(main, "_xl_modo_desde_formula", _mut_asterisco_primero)

    assert main._xl_modo_desde_formula("=ROUND([@[A]]*[@[B]],2)") == "SR"
    with pytest.raises(AssertionError):
        assert main._xl_modo_desde_formula("=ROUND([@[A]]*[@[B]],2)") == "ROUND"


def test_mutante_sin_roundup_en_modos_validos_falla(monkeypatch):
    """M5: quitar ROUNDUP de MODOS_VALIDOS rompe la compatibilidad."""
    monkeypatch.setattr(main, "MODOS_VALIDOS", ("ROUND", "SR", "ROUNDDOWN"))

    with pytest.raises(AssertionError):
        assert "ROUNDUP" in main.MODOS_VALIDOS


def test_mutante_menu_con_texto_sr_falla(monkeypatch):
    """M6: el texto del menú debe ser '0.00 SR', no 'SR'."""
    monkeypatch.setattr(main, "MODOS_REDONDEO",
                        [("Redondear", "ROUND"), ("SR", "SR"),
                         ("Redondear.menos", "ROUNDDOWN")])

    with pytest.raises(AssertionError):
        assert main.MODOS_REDONDEO[1] == ("0.00 SR", "SR")


# ── Y el módulo real NO está mutado ────────────────────────────────────────

def test_main_real_no_quedo_mutado():
    """Sanidad: al terminar, `main` sigue con su comportamiento correcto."""
    assert main._xl_formula_celda_modo(EXPR, "SR") == f"=({EXPR})"
    assert main._xl_modo_desde_formula("=([@[A]]*[@[B]])") == "SR"
    assert main._xl_modo_desde_formula("=ROUND([@[A]]*[@[B]],2)") == "ROUND"
    assert "ROUNDUP" in main.MODOS_VALIDOS
    assert main.MODOS_REDONDEO[1] == ("0.00 SR", "SR")
