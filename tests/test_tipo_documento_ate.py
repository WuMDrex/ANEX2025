# -*- coding: utf-8 -*-
"""
Regla nueva de `Tipo de Documento`: un código es ATE solo si es numérico y
MAYOR a `UMBRAL_CODIGO_ATE` (2000). `2000` exacto NO es ATE.

Cubre:
  1. `Estado._tipo_documento` (lógica pura, sin Excel ni UI).
  2. La fórmula `f_j` que se escribe en VARIACIÓN/OC/OTS/ATE.
  3. Que el umbral viva en UN solo sitio (constante) y no duplicado a mano.
  4. Que la regla vieja de formato condicional (`=X+0=X`) siga contemplada.
  5. UI offscreen de `DlgComponentes` (validar_agregar + tooltip/etiqueta).

No abre Excel: las pruebas COM están en test_tipo_documento_ate_excel.py.
"""

import re
import sys

import pytest

main = pytest.importorskip("main")

ATE, OTS, OC, NADA = "ATE", "OTS", "OC", "---"
UMBRAL = main.UMBRAL_CODIGO_ATE


def tipo(codigo, diferencia):
    """`_tipo_documento` no usa `self`: se llama como función sin instancia."""
    return main.Estado._tipo_documento(None, codigo, diferencia)


# ═════════════════════════════════════════════════════════════════════
#  1) lógica pura
# ═════════════════════════════════════════════════════════════════════
def test_umbral_es_2000():
    assert UMBRAL == 2000, f"UMBRAL_CODIGO_ATE = {UMBRAL!r}"


@pytest.mark.parametrize("codigo", ["2500", "3000", "999999", "02500"])
@pytest.mark.parametrize("diferencia", [5, 0, -5])
def test_codigo_mayor_al_umbral_es_ate_sin_importar_el_signo(codigo, diferencia):
    # La rama ATE va primero: la cantidad variada no influye.
    assert tipo(codigo, diferencia) == ATE


@pytest.mark.parametrize("codigo", ["2000", "1999", "1083", "0", "0001"])
@pytest.mark.parametrize("diferencia,esperado", [(5, OTS), (0, NADA), (-5, OC)])
def test_numerico_menor_o_igual_al_umbral_sigue_la_cantidad(codigo, diferencia, esperado):
    # `2000` es el borde: si el umbral se hubiera escrito con `>=`, esta
    # prueba FALLA (devolvería ATE en vez de OTS/OC/---).
    assert tipo(codigo, diferencia) == esperado


@pytest.mark.parametrize("codigo", ["P001", "E051", "M206", "C185", "A1"])
@pytest.mark.parametrize("diferencia,esperado", [(5, OTS), (0, NADA), (-5, OC)])
def test_alfanumerico_sigue_la_cantidad(codigo, diferencia, esperado):
    # Un código con letras nunca es ATE aunque el número sea grande.
    assert tipo(codigo, diferencia) == esperado


@pytest.mark.parametrize("vacio", ["", " ", "   ", None, "\t"])
def test_codigo_vacio_no_revienta_y_da_guion(vacio):
    assert tipo(vacio, 0) == NADA


def test_decimal_y_signos_no_son_ate():
    # Excel escribiría 2500.0 como número, pero el código llega como texto:
    # el punto/signo lo saca de la rama ATE y manda la cantidad.
    assert tipo("2500.0", 5) == OTS
    assert tipo("2500.0", -5) == OC
    assert tipo("+2500", 5) == OTS
    assert tipo("-2500", 5) == OTS


def test_digitos_unicode_no_son_ate():
    # isascii() evita que "٢٥٠٠" (2500 árabe) o "²" pasen por numéricos.
    assert tipo("\u0662\u0665\u0660\u0660", 5) == OTS


def test_cantidad_en_decimal_no_rompe():
    from decimal import Decimal

    assert tipo("2500", Decimal("-3.50")) == ATE
    assert tipo("1500", Decimal("0.01")) == OTS
    assert tipo("1500", Decimal("-0.01")) == OC
    assert tipo("1500", Decimal("0")) == NADA


# ═════════════════════════════════════════════════════════════════════
#  2) fórmula de tabla
# ═════════════════════════════════════════════════════════════════════
def _formulas(col_var="CANTIDAD VARIADA"):
    return main._xl_formulas_base_tabla("ANE_T", "AV_T", "ROUND", "ROUND", col_var)


def test_formula_tipo_de_documento_exacta():
    f_j = _formulas()[8]
    esperado = (f'=IF(AND(ISNUMBER([@[No.]]),[@[No.]]>{UMBRAL}),"ATE",'
                'IF([@[CANTIDAD VARIADA]]>0,"OTS",'
                'IF([@[CANTIDAD VARIADA]]<0,"OC","---")))')
    assert f_j == esperado
    assert f">{UMBRAL}" in f_j
    assert f"={UMBRAL}" not in f_j          # no usa >=
    assert "IFERROR" not in f_j


def test_formula_usa_la_constante_del_umbral():
    # Si alguien cambia UMBRAL_CODIGO_ATE, la fórmula debe seguirlo.
    original = main.UMBRAL_CODIGO_ATE
    try:
        main.UMBRAL_CODIGO_ATE = 777
        assert ">777" in _formulas()[8]
    finally:
        main.UMBRAL_CODIGO_ATE = original
    assert ">2000" in _formulas()[8]


def test_las_otras_formulas_no_se_tocaron():
    f = _formulas()
    assert len(f) == 11
    assert f[0].startswith("=_xlfn.XLOOKUP([@[No.]],CODIGOS_SICOA[")
    assert f[9] == "=[@[CANTIDAD - V]]-[@[CANTIDAD]]"


# ═════════════════════════════════════════════════════════════════════
#  3) el umbral vive en un solo sitio
# ═════════════════════════════════════════════════════════════════════
def test_umbral_no_esta_duplicado_en_el_codigo():
    import inspect

    fuente = inspect.getsource(main)
    assert "UMBRAL_CODIGO_ATE = 2000" in fuente
    # No debe quedar el número escrito a mano en ninguna comparación/fórmula.
    assert not re.search(r"[><]=?\s*2000", fuente), \
        "hay un 2000 escrito a mano además de la constante"
    assert ">2000" not in fuente, "la fórmula no usa la constante"


# ═════════════════════════════════════════════════════════════════════
#  4) formato condicional: patrón de la regla vieja
# ═════════════════════════════════════════════════════════════════════
def test_formato_condicional_borra_la_regla_vieja_sin_tocar_otras():
    import inspect
    import textwrap

    fuente = textwrap.dedent(
        inspect.getsource(main._xl_formato_condicional_codigo_numerico))
    # La expresión nueva incluye el umbral y la comparación de la MISMA celda.
    assert "*({direccion}>{UMBRAL_CODIGO_ATE})" in fuente
    assert 'formula_vieja = f"={direccion}+0={direccion}"' in fuente
    # La retroreferencia \1 evita borrar reglas ajenas con la misma forma.
    assert r"([A-Z]{1,3}\d+)\+0=\1$" in fuente


# patrón tal como lo aplica el código (se replica para probar la semántica)
_PATRON_VIEJO = r"=?([A-Z]{1,3}\d+)\+0=\1$"


@pytest.mark.parametrize("formula,se_borra", [
    ("=A7+0=A7", True),
    ("=AB12+0=AB12", True),
    ("=A7+0=A8", False),         # compara otra celda: NO se borra
    ("=AND(A7+0=A7,A7>2000)", False),
    ("=A7>2000", False),
])
def test_patron_regla_vieja(formula, se_borra):
    normalizada = formula.replace(" ", "").upper().lstrip("=")
    assert bool(re.fullmatch(_PATRON_VIEJO, normalizada)) is se_borra


# ═════════════════════════════════════════════════════════════════════
#  5) UI offscreen de DlgComponentes
# ═════════════════════════════════════════════════════════════════════
def _estado_minimo():
    e = main.Estado.__new__(main.Estado)          # sin leer la plantilla
    e.bd = [main.componente_nuevo("P001", "POSTE DE PRUEBA", "U", "10.00")]
    e.reset()
    return e


@pytest.fixture
def dlg(qtbot):
    d = main.DlgComponentes(None, _estado_minimo())
    qtbot.addWidget(d)
    return d


def _llenar(dlg, codigo):
    dlg.cod.setText(codigo)
    dlg.desc.setText("COMPONENTE DE PRUEBA")
    dlg.uni.setText("U")
    dlg.precio.setText("10.50")
    dlg.validar_agregar()


@pytest.mark.parametrize("codigo,habilitado", [
    ("2500", True),      # numérico > umbral: ATE al exportar
    ("2000", True),      # borde: se acepta, pero NO es ATE
    ("1999", True),
    ("1083", True),
    ("P001", False),     # alfanumérico: rechazado
    ("", False),
])
def test_validar_agregar_solo_numericos(dlg, codigo, habilitado):
    _llenar(dlg, codigo)
    assert dlg.btn_add.isEnabled() is habilitado, \
        f"código {codigo!r}: Agregar={'habilitado' if dlg.btn_add.isEnabled() else 'deshabilitado'}"


def test_tooltip_y_etiqueta_mencionan_el_umbral(dlg):
    assert str(UMBRAL) in dlg.cod.toolTip()
    dlg.cod.setText("2500")
    dlg.cod_editado("2500")
    assert str(UMBRAL) in dlg.lbl_modo.text(), dlg.lbl_modo.text()


def test_agregar_rechaza_alfanumerico(dlg, monkeypatch):
    avisos = []
    monkeypatch.setattr(main, "msg", lambda *a, **k: avisos.append(a))
    antes = len(dlg.datos)
    _llenar(dlg, "P001")
    dlg.agregar()
    assert len(dlg.datos) == antes, "se agregó un código alfanumérico"
    assert avisos, "no avisó del código no válido"


def test_agregar_acepta_2000_y_nace_sin_lineas(dlg, monkeypatch):
    monkeypatch.setattr(main, "msg", lambda *a, **k: None)
    _llenar(dlg, "2000")
    dlg.agregar()
    nuevo = next(c for c in dlg.datos if c["codigo"] == "2000")
    assert all(nuevo[k] == 0 for k in ("MT", "MTBT", "BT", "G", "v76", "v199")), nuevo


def test_agregar_acepta_mayor_al_umbral(dlg, monkeypatch):
    monkeypatch.setattr(main, "msg", lambda *a, **k: None)
    _llenar(dlg, "2500")
    dlg.agregar()
    assert any(c["codigo"] == "2500" for c in dlg.datos)


# ═════════════════════════════════════════════════════════════════════
#  6) control de mutación: las pruebas de arriba SÍ discriminan
#     (se muta una COPIA de main.py en %TEMP%; el original no se toca)
# ═════════════════════════════════════════════════════════════════════
import importlib.util                      # noqa: E402
import os                                  # noqa: E402
import pathlib                             # noqa: E402

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MATRIZ = [
    ("2500", 5, ATE), ("3000", 0, ATE), ("2500", -5, ATE),
    ("2000", 5, OTS), ("2000", 0, NADA), ("2000", -5, OC),
    ("1999", 5, OTS), ("1083", -5, OC),
    ("P001", 5, OTS), ("", 0, NADA),
]


def _importar_mutado(tmp_path, nombre, cambios):
    """Copia main.py a %TEMP%, le aplica `cambios` y la importa como módulo."""
    origen = pathlib.Path(RAIZ, "main.py").read_text(encoding="utf-8")
    for viejo, nuevo in cambios:
        assert viejo in origen, f"el texto a mutar ya no existe: {viejo!r}"
        origen = origen.replace(viejo, nuevo, 1)
    destino = tmp_path / f"{nombre}.py"
    destino.write_text(origen, encoding="utf-8")
    spec = importlib.util.spec_from_file_location(nombre, destino)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[nombre] = mod
    try:
        spec.loader.exec_module(mod)
    finally:
        sys.modules.pop(nombre, None)
    return mod


def _fallos_de_la_matriz(mod):
    return [(c, d) for c, d, esperado in MATRIZ
            if mod.Estado._tipo_documento(None, c, d) != esperado]


def test_mutacion_ge_en_vez_de_gt_es_detectada(tmp_path):
    mod = _importar_mutado(tmp_path, "main_mut_ge", [
        ("int(texto) > UMBRAL_CODIGO_ATE",
         "int(texto) >= UMBRAL_CODIGO_ATE")])
    fallos = _fallos_de_la_matriz(mod)
    assert ("2000", 5) in fallos, fallos          # 2000 pasaría a ATE
    assert mod.Estado._tipo_documento(None, "2000", 5) == "ATE"


def test_mutacion_sin_umbral_es_detectada(tmp_path):
    mod = _importar_mutado(tmp_path, "main_mut_sin", [
        ("""if (texto.isascii() and texto.isdigit()
                and int(texto) > UMBRAL_CODIGO_ATE):""",
         """if texto.isascii() and texto.isdigit():""")])
    assert mod.Estado._tipo_documento(None, "1999", 5) == "ATE"
    assert ("1999", 5) in _fallos_de_la_matriz(mod)


def test_mutacion_umbral_equivocado_es_detectada(tmp_path):
    mod = _importar_mutado(tmp_path, "main_mut_umbral", [
        ("UMBRAL_CODIGO_ATE = 2000", "UMBRAL_CODIGO_ATE = 1000")])
    assert mod.Estado._tipo_documento(None, "1083", -5) == "ATE"
    assert ("1083", -5) in _fallos_de_la_matriz(mod)


def test_mutacion_de_la_formula_es_detectada(tmp_path):
    mod = _importar_mutado(tmp_path, "main_mut_formula", [
        ("[@[No.]]>{UMBRAL_CODIGO_ATE}", "[@[No.]]>={UMBRAL_CODIGO_ATE}")])
    f_j = mod._xl_formulas_base_tabla("ANE_T", "AV_T", "ROUND", "ROUND",
                                      "CANTIDAD VARIADA")[8]
    esperado = (f'=IF(AND(ISNUMBER([@[No.]]),[@[No.]]>{UMBRAL}),"ATE",'
                'IF([@[CANTIDAD VARIADA]]>0,"OTS",'
                'IF([@[CANTIDAD VARIADA]]<0,"OC","---")))')
    assert f_j != esperado, "la prueba de la fórmula no distingue > de >="
