# -*- coding: utf-8 -*-
"""
Pruebas del modo «0.00 SR» (multiplicación SIN redondeo) en la exportación /
importación de `main.py`.

Qué se comprueba, sin Excel ni COM:

1. `_xl_formula_celda_modo(expr, modo)` escribe `=(expr)` para `SR` y
   `=FUNCION(expr,2)` para los demás modos; un modo desconocido cae en
   `MODO_DEFECTO` y NUNCA produce `=SR(...)` ni un nombre inválido.
2. Roundtrip: exportar + `_xl_modo_desde_formula` devuelve el modo original.
3. Regresión del orden de comprobación en `_xl_modo_desde_formula`:
   `ROUNDUP`/`ROUNDDOWN` contienen la cadena `ROUND` y las fórmulas llevan `*`,
   así que si el orden se rompe, una fila redondeada se leería como `SR`.
4. Casos borde: `None`, `""`, minúsculas, expresión con `*` sin `ROUND`, y
   fórmula sin `*` ni `ROUND`.
5. `_xl_formula_documento(tipo, "SR", col_var)` genera el `=IF(...)` de
   OC/OTS/ATE sin `ROUND` y con paréntesis equilibrados.
6. Las constantes del menú: 3 entradas, la segunda es `("0.00 SR", "SR")` y
   `"ROUNDUP"` sigue en `MODOS_VALIDOS` por compatibilidad.

Se importa `main` como módulo; `tests/conftest.py` pone `license_manager` falso
y Qt en offscreen, y aquí NO se llama a `main()`.
"""

import pytest

import main

EXPR = "[@[Cantidad]]*[@[Precio]]"

# Modos que la app puede escribir hoy + el que se arrastra de archivos viejos.
MODOS_ESCRIBIBLES = ["ROUND", "SR", "ROUNDDOWN", "ROUNDUP"]


def _parentesis_equilibrados(texto):
    """True si todos los paréntesis del texto abren y cierran en orden."""
    nivel = 0
    for ch in texto:
        if ch == "(":
            nivel += 1
        elif ch == ")":
            nivel -= 1
            if nivel < 0:
                return False
    return nivel == 0


# ════════════════════════════════════════════════════════════════════════════
# 1. EXPORTACIÓN POR MODO
# ════════════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("modo", MODOS_ESCRIBIBLES)
def test_exportar_modo_valido_no_genera_formula_invalida(modo):
    """Cada modo válido produce una fórmula que empieza por '=' y cierra bien."""
    formula = main._xl_formula_celda_modo(EXPR, modo)

    assert formula.startswith("="), formula
    assert "=SR(" not in formula.upper(), f"SR no es una función de Excel: {formula}"
    assert _parentesis_equilibrados(formula), formula
    assert EXPR in formula, f"la expresión se perdió: {formula}"


def test_exportar_sr_es_multiplicacion_sin_redondeo():
    """«0.00 SR» -> =(expr): sin ROUND y con la expresión entre paréntesis."""
    formula = main._xl_formula_celda_modo(EXPR, "SR")

    assert formula == f"=({EXPR})", formula
    assert formula.startswith("=("), formula
    assert "ROUND" not in formula.upper(), f"SR no debe redondear: {formula}"
    assert ",2" not in formula, f"SR no debe cuantizar: {formula}"


@pytest.mark.parametrize("modo", ["ROUND", "ROUNDDOWN", "ROUNDUP"])
def test_exportar_modos_redondeados_llaman_a_su_funcion_con_dos_decimales(modo):
    formula = main._xl_formula_celda_modo(EXPR, modo)

    assert formula == f"={modo}({EXPR},2)", formula
    assert formula.upper().startswith(f"={modo}("), formula
    assert formula.endswith(",2)"), formula


@pytest.mark.parametrize("modo", ["BASURA", "", None, "sr", "round", "  SR  "])
def test_exportar_modo_desconocido_cae_en_el_defecto(modo):
    """Un modo inválido nunca se escribe tal cual: cae en MODO_DEFECTO."""
    formula = main._xl_formula_celda_modo(EXPR, modo)
    esperado = f"={main.MODO_DEFECTO}({EXPR},2)"

    assert formula == esperado, f"modo={modo!r} dio {formula!r}"
    assert "=SR(" not in formula.upper(), formula
    assert main._xl_modo_desde_formula(formula) == main.MODO_DEFECTO, formula


def test_exportar_modo_defecto_explicito():
    """MODO_DEFECTO se escribe igual que cualquier modo válido."""
    assert (main._xl_formula_celda_modo(EXPR, main.MODO_DEFECTO)
            == main._xl_formula_celda_modo(EXPR, "ROUND"))


# ════════════════════════════════════════════════════════════════════════════
# 2. ROUNDTRIP EXPORTAR -> LEER
# ════════════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("modo", MODOS_ESCRIBIBLES)
def test_roundtrip_conserva_el_modo(modo):
    """La fórmula escrita se vuelve a leer como el mismo modo."""
    formula = main._xl_formula_celda_modo(EXPR, modo)

    assert main._xl_modo_desde_formula(formula) == modo, \
        f"{modo} -> {formula} -> {main._xl_modo_desde_formula(formula)}"


def test_roundtrip_modo_invalido_vuelve_como_defecto():
    """«BASURA» se escribe como ROUND y al leer vuelve como MODO_DEFECTO."""
    formula = main._xl_formula_celda_modo(EXPR, "BASURA")

    assert main._xl_modo_desde_formula(formula) == main.MODO_DEFECTO


def test_sr_no_se_confunde_con_redondeo_en_el_roundtrip():
    """Garantía explícita: SR -> fórmula sin ROUND -> SR."""
    formula = main._xl_formula_celda_modo(EXPR, "SR")

    assert "ROUND" not in formula.upper()
    assert main._xl_modo_desde_formula(formula) == "SR"


# ════════════════════════════════════════════════════════════════════════════
# 3. REGRESIÓN DEL ORDEN DE COMPROBACIÓN (ROUNDUP -> ROUNDDOWN -> ROUND -> *)
# ════════════════════════════════════════════════════════════════════════════

def test_leer_round_con_multiplicacion_no_es_sr():
    """Si el orden se invierte, esta fórmula (tiene '*') se leería como SR."""
    assert main._xl_modo_desde_formula("=ROUND([@[A]]*[@[B]],2)") == "ROUND"


def test_leer_rounddown_con_multiplicacion_no_es_sr():
    assert main._xl_modo_desde_formula("=ROUNDDOWN([@[A]]*[@[B]],2)") == "ROUNDDOWN"


def test_leer_roundup_con_multiplicacion_no_es_round():
    """ROUNDUP contiene la cadena 'ROUND': debe comprobarse antes."""
    assert main._xl_modo_desde_formula("=ROUNDUP([@[A]]*[@[B]],2)") == "ROUNDUP"


def test_leer_rounddown_no_es_round():
    """ROUNDDOWN también contiene 'ROUND': debe comprobarse antes."""
    assert main._xl_modo_desde_formula("=ROUNDDOWN([@[A]]*[@[B]],2)") == "ROUNDDOWN"


def test_leer_multiplicacion_sin_redondeo_es_sr():
    assert main._xl_modo_desde_formula("=([@[A]]*[@[B]])") == "SR"


# ════════════════════════════════════════════════════════════════════════════
# 4. CASOS BORDE AL LEER
# ════════════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("formula", [None, ""])
def test_leer_vacio_devuelve_el_defecto(formula):
    assert main._xl_modo_desde_formula(formula) == main.MODO_DEFECTO


def test_leer_minusculas_reconoce_round():
    assert main._xl_modo_desde_formula("=round(a*b,2)") == "ROUND"


def test_leer_minusculas_reconoce_rounddown():
    assert main._xl_modo_desde_formula("=rounddown(a*b,2)") == "ROUNDDOWN"


def test_leer_multiplicacion_sin_round_es_sr():
    assert main._xl_modo_desde_formula("=(A*B)") == "SR"


def test_leer_sin_multiplicacion_ni_round_es_defecto():
    assert main._xl_modo_desde_formula("=SUM(A1:A3)") == main.MODO_DEFECTO


def test_leer_no_texto_se_convierte_a_texto_sin_error():
    """Un valor no textual (None/numero) no debe lanzar excepción."""
    assert main._xl_modo_desde_formula(0) == main.MODO_DEFECTO
    assert main._xl_modo_desde_formula(123) == main.MODO_DEFECTO


def test_leer_minusculas_de_multiplicacion_es_sr():
    assert main._xl_modo_desde_formula("=(a*b)") == "SR"


# ════════════════════════════════════════════════════════════════════════════
# 5. FÓRMULA DE OC/OTS/ATE EN MODO SR
# ════════════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("tipo", ["OC", "OTS", "ATE"])
def test_formula_documento_sr_sin_redondeo(tipo):
    formula = main._xl_formula_documento(tipo, "SR", "CANTIDAD VARIADA")

    assert formula.startswith("=IF("), formula
    assert f'[@[Tipo de Documento]]="{tipo}"' in formula, formula
    assert "ROUND" not in formula.upper(), f"SR no debe redondear: {formula}"
    assert ",2" not in formula, formula
    assert _parentesis_equilibrados(formula), formula
    # El producto va entre paréntesis antes de la coma del "sí".
    producto = "([@[CANTIDAD VARIADA]]*[@[PRECIO UNITARIO ( Q)]])"
    assert producto in formula, formula


def test_formula_documento_sr_texto_exacto():
    """Forma exacta esperada: el producto va entre paréntesis, sin ROUND."""
    esperado = ('=IF([@[Tipo de Documento]]="ATE",'
                '([@[CANTIDAD VARIADA]]*[@[PRECIO UNITARIO ( Q)]]),"")')
    assert main._xl_formula_documento("ATE", "SR", "CANTIDAD VARIADA") == esperado


def test_formula_documento_round_conserva_su_funcion():
    formula = main._xl_formula_documento("OC", "ROUND", "CANTIDAD VARIADA")

    assert "ROUND([@[CANTIDAD VARIADA]]*[@[PRECIO UNITARIO ( Q)]],2)" in formula, formula
    assert _parentesis_equilibrados(formula), formula


def test_formula_documento_roundtrip():
    """La fórmula de OC/OTS/ATE se vuelve a leer con el modo correcto."""
    for modo in MODOS_ESCRIBIBLES:
        formula = main._xl_formula_documento("OTS", modo, "CANTIDAD VARIADA")
        assert main._xl_modo_desde_formula(formula) == modo, (modo, formula)


# ════════════════════════════════════════════════════════════════════════════
# 6. CONSTANTES DEL MENÚ
# ════════════════════════════════════════════════════════════════════════════

def test_modos_redondeo_tiene_tres_entradas():
    assert len(main.MODOS_REDONDEO) == 3, main.MODOS_REDONDEO


def test_segunda_entrada_es_sin_redondeo():
    assert main.MODOS_REDONDEO[1] == ("0.00 SR", "SR"), main.MODOS_REDONDEO


def test_primera_y_tercera_entrada_del_menu():
    assert main.MODOS_REDONDEO[0] == ("Redondear", "ROUND"), main.MODOS_REDONDEO
    assert main.MODOS_REDONDEO[2] == ("Redondear.menos", "ROUNDDOWN"), main.MODOS_REDONDEO


def test_todos_los_modos_del_menu_son_validos():
    for _texto, modo in main.MODOS_REDONDEO:
        assert modo in main.MODOS_VALIDOS, (modo, main.MODOS_VALIDOS)


def test_roundup_sigue_valido_para_archivos_viejos():
    assert "ROUNDUP" in main.MODOS_VALIDOS, main.MODOS_VALIDOS
    assert main.MODO_POR_TEXTO.get("Redondear.más") == "ROUNDUP", main.MODO_POR_TEXTO


def test_modo_defecto_es_round():
    assert main.MODO_DEFECTO == "ROUND"
    assert main.MODO_DEFECTO in main.MODOS_VALIDOS


def test_texto_del_menu_sr_se_lee_como_sr():
    assert main.MODO_POR_TEXTO.get("0.00 SR") == "SR", main.MODO_POR_TEXTO
