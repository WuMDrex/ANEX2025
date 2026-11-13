# -*- coding: utf-8 -*-
"""
T1 + T2 del Ramal — lógica pura de `Estado` y UI offscreen (SIN Excel).

· T1: `VARIACION` / `OC` / `OTS` / `ATE` del Ramal se activan por separado de
  los proyectos; solo dependen del `ANE` (siempre presente) y del `ANE - VAR`
  del propio Ramal, que sigue espejándose de los proyectos.
· T2: en `ANE` / `ANE - VAR` del Ramal se pueden capturar filas de códigos que
  no existen en ningún proyecto («filas propias»); las heredadas solo cambian
  su redondeo y, si un código propio aparece en un proyecto, las cantidades
  se SUMAN (decisión de WuMDrex).

Este archivo se puede usar de dos formas:
  · `python tests/test_ramal_propio.py`          → runner propio, sin pytest.
  · `python -m pytest tests/test_ramal_propio.py` → si pytest está instalado.

No toca Excel ni el código fuente: la BD se inyecta a mano con
`Estado.__new__` (igual que `tests/test_bd_componentes.py`), así que no se abre
la plantilla. Las pruebas de UI usan `QT_QPA_PLATFORM=offscreen`.
"""

import os
import sys
import traceback

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RAIZ not in sys.path:
    sys.path.insert(0, RAIZ)

try:                                            # consola Windows = cp1252
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from decimal import Decimal                     # noqa: E402

import types                                    # noqa: E402
_falso = types.ModuleType("license_manager")
_falso.check_license = lambda: True
sys.modules.setdefault("license_manager", _falso)

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import main                                     # noqa: E402


# ═════════════════════════════════════════════════════════════════════
#  UTILIDADES DE LAS PRUEBAS
# ═════════════════════════════════════════════════════════════════════
def comp(codigo, precio="10.00", unidad="U", lineas=("MT",)):
    """Componente de BD con las marcas pedidas (el resto en 0)."""
    c = main.componente_nuevo(codigo, f"DESCRIPCION {codigo}", unidad, precio)
    for k, _, _ in main.LINEAS:
        c[k] = 1 if k in lineas else 0
    return c


# El ORDEN de esta lista es el de CODIGOS_SICOA: P001, P002, K001, P999.
BD_PRUEBA = [comp("P001"), comp("P002"), comp("K001", "100.00", "Km"),
             comp("P999", "5.00")]


def estado_nuevo(bd=None):
    """Estado real, sin leer la plantilla (la BD se inyecta a mano)."""
    e = main.Estado.__new__(main.Estado)
    e.bd = [dict(c) for c in (bd if bd is not None else BD_PRUEBA)]
    e.reset()
    return e


def agregar(base, hoja, linea, codigo, cantidad, redondeo="ROUND"):
    """Agrega una fila a la lista guardada (proyecto o Ramal)."""
    base["items"].setdefault(hoja, main._vacio())[linea].append(
        {"codigo": codigo, "cantidad": str(cantidad), "redondeo": redondeo})


def por_codigo(e, base, hoja, linea):
    return {f["codigo"]: f for f in e.filas(base, hoja)[linea]}


# ═════════════════════════════════════════════════════════════════════
#  1) T1 · el Ramal con UN proyecto no activa nada por su cuenta
# ═════════════════════════════════════════════════════════════════════
def test_un_proyecto_no_activa_hojas_del_ramal():
    e = estado_nuevo()
    p0 = e.proyectos[0]
    agregar(p0, "ANE", "MT", "P001", 3)
    e.activar_hoja(p0, "ANE_VAR")
    e.activar_hoja(p0, "VARIACION")

    assert e.ramal_visible() is False, "con un proyecto el Ramal no debe existir"
    for h in ("VARIACION", "OC", "OTS", "ATE"):
        assert e.ramal["hojas"][h] is False, \
            f"con 1 proyecto el Ramal no debe tener {h}"


# ═════════════════════════════════════════════════════════════════════
#  2) T1 · transición 1→2: se siembra UNA vez la unión de los proyectos
# ═════════════════════════════════════════════════════════════════════
def test_siembra_en_la_transicion_1_a_2():
    e = estado_nuevo()
    p0 = e.proyectos[0]
    agregar(p0, "ANE", "MT", "P001", 3)
    e.activar_hoja(p0, "ANE_VAR")
    e.activar_hoja(p0, "VARIACION")

    e.agregar_proyecto()

    assert e.ramal_visible() is True
    assert e.ramal["hojas"]["ANE"] is True
    assert e.ramal["hojas"]["ANE_VAR"] is True, "p0 la tiene: se espeja"
    assert e.ramal["hojas"]["VARIACION"] is True, "arranca con la unión de hojas"
    # La siembra copia lo que los proyectos YA tienen; no inventa hojas.
    assert e.ramal["hojas"]["OC"] is False
    assert e.ramal["hojas"]["OTS"] is False
    assert e.ramal["hojas"]["ATE"] is False


# ═════════════════════════════════════════════════════════════════════
#  3) T1 · las derivadas del Ramal ya no dependen de los proyectos
# ═════════════════════════════════════════════════════════════════════
def test_derivadas_del_ramal_son_independientes():
    e = estado_nuevo()
    p0 = e.proyectos[0]
    agregar(p0, "ANE", "MT", "P001", 3)
    e.activar_hoja(p0, "ANE_VAR")
    e.activar_hoja(p0, "VARIACION")
    e.agregar_proyecto()

    # Quitarla en el proyecto NO apaga la del Ramal.
    e.quitar_hoja(p0, "VARIACION")
    assert e.ramal["hojas"]["VARIACION"] is True, \
        "el Ramal es dueño de su VARIACIÓN"
    # Activarla en el proyecto tampoco cambia la decisión del Ramal.
    e.quitar_hoja(e.ramal, "VARIACION")
    assert e.ramal["hojas"]["VARIACION"] is False
    e.activar_hoja(p0, "VARIACION")
    assert e.ramal["hojas"]["VARIACION"] is False, \
        "activarla en el proyecto no la reactiva en el Ramal"
    # Y el Ramal la activa por sí solo (la capa Estado no pide permiso).
    e.activar_hoja(e.ramal, "VARIACION")
    assert e.ramal["hojas"]["VARIACION"] is True

    # Igual con OC / OTS / ATE.
    for h in ("OC", "OTS", "ATE"):
        e.activar_hoja(e.ramal, h)
        assert e.ramal["hojas"][h] is True
        e.quitar_hoja(e.ramal, h)
        assert e.ramal["hojas"][h] is False
        assert all(p["hojas"][h] is False for p in e.proyectos), \
            f"{h} del Ramal no debe tocar la hoja de los proyectos"


# ═════════════════════════════════════════════════════════════════════
#  4) T1 · sin ANE - VAR en los proyectos, el Ramal pierde ANE - VAR
#     y sus derivadas se apagan solas
# ═════════════════════════════════════════════════════════════════════
def test_sin_ane_var_se_apagan_las_derivadas():
    e = estado_nuevo()
    p0, p1 = e.proyectos[0], None
    agregar(p0, "ANE", "MT", "P001", 1)
    e.activar_hoja(p0, "ANE_VAR")
    p1 = e.agregar_proyecto()                    # p1 SIN ANE - VAR
    e.activar_hoja(e.ramal, "VARIACION")
    e.activar_hoja(e.ramal, "OC")
    agregar(e.ramal, "ANE_VAR", "MT", "P002", 7)  # fila propia en ANE - VAR
    assert e._cant_propia("ANE_VAR")["MT"]["P002"] == Decimal(7)

    e.quitar_hoja(p0, "ANE_VAR")                 # ningún proyecto la tiene ya

    assert e.ramal["hojas"]["ANE_VAR"] is False
    for h in main.DEPENDEN_DE_VAR:
        assert e.ramal["hojas"][h] is False, \
            f"{h} no puede quedar activa sin ANE - VAR"
    # Por diseño la hoja se descarta con sus filas (no hay dónde calcularlas).
    assert "ANE_VAR" not in e.ramal["items"], \
        "al perder ANE - VAR el Ramal descarta sus filas de esa hoja"


# ═════════════════════════════════════════════════════════════════════
#  5) T1 · activar VARIACIÓN en el Ramal con 1 proyecto y sin ANE - VAR
#     (la capa Estado lo permite; el aviso/bloqueo es de la UI)
# ═════════════════════════════════════════════════════════════════════
def test_variacion_en_el_ramal_depende_del_ane():
    e = estado_nuevo()
    p0 = e.proyectos[0]
    agregar(p0, "ANE", "MT", "P001", 3)
    e.activar_hoja(p0, "VARIACION")
    e.agregar_proyecto()                         # ningún proyecto tiene ANE - VAR

    assert e.ramal["hojas"]["ANE"] is True
    assert e.ramal["hojas"]["ANE_VAR"] is False
    assert e.ramal["hojas"]["VARIACION"] is False, \
        "la sincronización apaga las derivadas sin ANE - VAR"

    # El Estado no bloquea: el bloqueo con aviso vive en DlgEdicion.toggle.
    e.activar_hoja(e.ramal, "VARIACION")
    assert e.ramal["hojas"]["VARIACION"] is True

    # La fila del ANE del Ramal sale igualmente del proyecto.
    assert e._cant_editable(e.ramal, "ANE")["MT"]["P001"] == Decimal(3)


# ═════════════════════════════════════════════════════════════════════
#  6) T1 · `_sembrar_ramal` corre SOLO en la transición 1→2
# ═════════════════════════════════════════════════════════════════════
def test_sembrar_solo_en_la_transicion_1_a_2():
    llamadas = []
    original = main.Estado._sembrar_ramal

    def espia(self):
        llamadas.append(len(self.proyectos))
        return original(self)

    main.Estado._sembrar_ramal = espia
    try:
        e = estado_nuevo()
        agregar(e.proyectos[0], "ANE", "MT", "P001", 1)
        e.agregar_proyecto()      # 1 → 2: siembra
        e.agregar_proyecto()      # 2 → 3: NO
        e.agregar_proyecto()      # 3 → 4: NO
        e.eliminar_proyecto(e.proyectos[-1])   # 4 → 3: NO
        e.quitar_hoja(e.proyectos[0], "ANE_VAR")  # sincroniza: NO
    finally:
        main.Estado._sembrar_ramal = original

    assert llamadas == [2], f"se sembró en {llamadas} y solo debía ser en 2"


def test_no_se_reactiva_lo_que_el_usuario_apago():
    e = estado_nuevo()
    p0 = e.proyectos[0]
    agregar(p0, "ANE", "MT", "P001", 1)
    e.activar_hoja(p0, "ANE_VAR")
    e.agregar_proyecto()
    e.quitar_hoja(e.ramal, "ANE_VAR")            # el usuario la apagó
    e.quitar_hoja(e.ramal, "VARIACION")

    e.activar_hoja(p0, "OC")                     # cambia el proyecto
    e.agregar_proyecto()                         # aparece un tercero

    assert e.ramal["hojas"]["ANE_VAR"] is True, "se espeja: p0 la tiene"
    assert e.ramal["hojas"]["OC"] is False, "no se siembra dos veces"


# ═════════════════════════════════════════════════════════════════════
#  7) T2 · fila propia vs. fila heredada
# ═════════════════════════════════════════════════════════════════════
def test_propio_vs_heredado():
    e = estado_nuevo()
    p0 = e.proyectos[0]
    agregar(p0, "ANE", "MT", "P001", 3)
    e.agregar_proyecto()
    agregar(e.ramal, "ANE", "MT", "P002", 5)     # P002 no está en ningún proyecto

    assert e._cant_propia("ANE")["MT"].get("P002", Decimal(0)) > 0
    assert e.codigo_heredado("ANE", "MT", "P002") is False
    assert e._cant_propia("ANE")["MT"].get("P001", Decimal(0)) == 0
    assert e.codigo_heredado("ANE", "MT", "P001") is True

    filas = por_codigo(e, e.ramal, "ANE", "MT")
    assert filas["P002"]["propio"] is True
    assert filas["P002"]["cant_propia"] == Decimal(5)
    assert filas["P001"]["propio"] is False
    assert filas["P001"]["cant_propia"] == Decimal(0)

    # Orden SICOA: P001 (índice 0) antes que P002 (índice 1).
    assert [f["codigo"] for f in e.filas(e.ramal, "ANE")["MT"]] == ["P001", "P002"]


def test_mismas_funciones_en_ane_var():
    e = estado_nuevo()
    p0 = e.proyectos[0]
    e.activar_hoja(p0, "ANE_VAR")
    agregar(p0, "ANE_VAR", "BT", "P001", 2)
    e.agregar_proyecto()
    agregar(e.ramal, "ANE_VAR", "BT", "P999", 4)

    assert e._cant_propia("ANE_VAR")["BT"].get("P999", Decimal(0)) > 0
    assert e.codigo_heredado("ANE_VAR", "BT", "P001") is True
    filas = por_codigo(e, e.ramal, "ANE_VAR", "BT")
    assert filas["P001"]["cantidad"] == Decimal(2)
    assert filas["P999"]["cantidad"] == Decimal(4)


# ═════════════════════════════════════════════════════════════════════
#  8) T2 · si el código propio aparece en un proyecto, se SUMAN
# ═════════════════════════════════════════════════════════════════════
def test_codigo_propio_que_aparece_en_un_proyecto_se_suma():
    e = estado_nuevo()
    p0, p1 = e.proyectos[0], None
    agregar(p0, "ANE", "MT", "P001", 3)
    p1 = e.agregar_proyecto()
    agregar(e.ramal, "ANE", "MT", "P002", 5)
    assert por_codigo(e, e.ramal, "ANE", "MT")["P002"]["propio"] is True

    agregar(p1, "ANE", "MT", "P002", 2)          # ahora también está en p1

    fila = por_codigo(e, e.ramal, "ANE", "MT")["P002"]
    assert fila["cantidad"] == Decimal(7), f"debía sumar 5+2 y dio {fila['cantidad']}"
    assert fila["cant_propia"] == Decimal(5), "lo capturado en el Ramal no se pierde"
    assert fila["propio"] is False, "al venir de un proyecto ya no es propia"

    assert e._cant_proyectos("ANE")["MT"]["P002"] == Decimal(2), \
        "_cant_proyectos solo mira los proyectos"
    assert e._cant_editable(e.ramal, "ANE")["MT"]["P002"] == Decimal(7)
    assert e._cant_editable(p0, "ANE")["MT"]["P001"] == Decimal(3), \
        "el proyecto no se contamina con lo del Ramal"


# ═════════════════════════════════════════════════════════════════════
#  9) T2 · quitar el proyecto del código: la fila propia sobrevive
# ═════════════════════════════════════════════════════════════════════
def test_fila_propia_sobrevive_al_quitar_el_proyecto():
    e = estado_nuevo()
    p0 = e.proyectos[0]
    agregar(p0, "ANE", "MT", "P002", 4)          # p0 tiene P002
    e.agregar_proyecto()
    agregar(e.ramal, "ANE", "MT", "P002", 6)
    e.agregar_proyecto()                         # 3 proyectos: el Ramal sigue
    assert por_codigo(e, e.ramal, "ANE", "MT")["P002"]["cantidad"] == Decimal(10)

    e.eliminar_proyecto(p0)                      # 3 → 2: el Ramal no se destruye

    assert e.ramal_visible() is True
    fila = por_codigo(e, e.ramal, "ANE", "MT")["P002"]
    assert fila["cantidad"] == Decimal(6), \
        "la cantidad propia debe sobrevivir sin el proyecto"
    assert fila["propio"] is True, "vuelve a ser propia: ya no viene de nadie"
    assert e.codigo_heredado("ANE", "MT", "P002") is False


def test_quitar_un_proyecto_ajeno_no_toca_las_propias():
    e = estado_nuevo()
    p0 = e.proyectos[0]
    agregar(p0, "ANE", "MT", "P001", 1)
    p1 = e.agregar_proyecto()
    agregar(p1, "ANE", "MT", "P002", 9)
    agregar(e.ramal, "ANE", "MT", "P999", 3)
    e.agregar_proyecto()

    e.eliminar_proyecto(p1)

    filas = por_codigo(e, e.ramal, "ANE", "MT")
    assert "P002" not in filas, "el código del proyecto eliminado desaparece"
    assert filas["P999"]["cantidad"] == Decimal(3)
    assert filas["P999"]["propio"] is True


# ═════════════════════════════════════════════════════════════════════
#  10) T2 · Decimal y kilometraje de 3 decimales
# ═════════════════════════════════════════════════════════════════════
def test_kilometraje_tres_decimales():
    e = estado_nuevo()
    p0 = e.proyectos[0]
    agregar(p0, "ANE", "MT", "K001", Decimal("1.75"))
    e.agregar_proyecto()
    agregar(e.ramal, "ANE", "MT", "K001", Decimal("2.5"))

    fila = por_codigo(e, e.ramal, "ANE", "MT")["K001"]
    assert fila["cantidad"] == Decimal("4.25")
    assert fila["cant_propia"] == Decimal("2.5")
    # El total usa el precio * la cantidad, con el redondeo de la fila.
    assert fila["bruto"] == Decimal("425.00")
    assert fila["total"] == Decimal("425.00")


def test_cuantizar_al_importar_no_altera_lo_ya_redondeado():
    """Reproduce la resta + quantize de `_xl_leer_items` (Román: fila del Ramal)."""
    def reconstruir(leido, proyectos):
        return (Decimal(leido) - Decimal(proyectos)).quantize(Decimal("0.001"))

    casos = [("4.25", "1.75", Decimal("2.500")),
             ("3", "1", Decimal("2.000")),
             ("0.001", "0", Decimal("0.001")),
             ("10.5", "10.5", Decimal("0.000"))]
    for leido, proys, esperado in casos:
        obtenido = reconstruir(leido, proys)
        assert obtenido == esperado, f"{leido}-{proys}: {obtenido} != {esperado}"
        # Idempotencia: volver a cuantizar no cambia el valor.
        assert obtenido.quantize(Decimal("0.001")) == obtenido
        # Y el Decimal conserva los 3 decimales al pasar por texto.
        assert Decimal(str(obtenido)) == obtenido


def test_literal_de_cantidad_para_la_formula():
    lit = main._xl_literal_cantidad
    casos = [(Decimal("2.000"), "2"), (Decimal("2.5"), "2.5"),
             (Decimal("0.750"), "0.75"), (Decimal(0), "0"),
             (Decimal("1000"), "1000"), (Decimal("0.001"), "0.001"),
             (Decimal("12.345"), "12.345"), ("7", "7")]
    for valor, esperado in casos:
        assert lit(valor) == esperado, f"{valor!r} → {lit(valor)!r} != {esperado!r}"
    # Nunca en notación científica (Excel no la leería como número).
    for valor in ("1000", "1000000", "0.0001"):
        texto = lit(Decimal(valor))
        assert "E" not in texto.upper(), f"{valor} salió como {texto!r}"


# ═════════════════════════════════════════════════════════════════════
#  11) Regresión: orden SICOA, redondeo por celda y hojas de proyecto
# ═════════════════════════════════════════════════════════════════════
def test_regresion_orden_sicoa():
    e = estado_nuevo()
    lista = [{"codigo": "ZZZ"}, {"codigo": "P002"}, {"codigo": "P001"}]
    e.ordenar_por_codigo(lista)
    assert [it["codigo"] for it in lista] == ["P001", "P002", "ZZZ"], \
        "los códigos externos van al final"
    # posicion_por_codigo inserta en el hueco de SICOA
    assert e.posicion_por_codigo([{"codigo": "P001"}, {"codigo": "P999"}],
                                 "P002") == 1


def test_regresion_redondeo_por_celda():
    assert main.redondear_monto(Decimal("10.005"), "ROUND") == Decimal("10.01")
    assert main.redondear_monto(Decimal("10.005"), "ROUNDDOWN") == Decimal("10.00")
    assert main.redondear_monto(Decimal("10.004"), "ROUNDUP") == Decimal("10.01")
    assert main.redondear_monto(Decimal("10.005"), "SR") == Decimal("10.005")
    assert main.texto_total(Decimal("10.01"), "SR") == "Q 10.01 SR"
    assert main.texto_total(Decimal("10.01"), "ROUNDUP") == "Q 10.01 +"
    assert main.texto_total(Decimal("10.00"), "ROUNDDOWN") == "Q 10.00 -"
    assert main.texto_total(Decimal("10.00"), "ROUND") == "Q 10.00"
    # Los modos ofrecidos siguen siendo los tres vigentes.
    assert [m for _, m in main.MODOS_REDONDEO] == ["ROUND", "SR", "ROUNDDOWN"]


def test_regresion_filas_de_proyecto_no_traen_propio():
    e = estado_nuevo()
    p0 = e.proyectos[0]
    agregar(p0, "ANE", "MT", "P001", 3, "SR")
    fila = e.filas(p0, "ANE")["MT"][0]
    assert "propio" not in fila, "las filas de un proyecto no son «propias»"
    assert "cant_propia" not in fila
    assert fila["redondeo"] == "SR", "el modo individual de la fila se conserva"
    assert fila["total"] == Decimal("30.00")


def test_renombrar_y_uso_del_codigo_miran_el_ramal():
    e = estado_nuevo()
    agregar(e.proyectos[0], "ANE", "MT", "P001", 3)
    e.agregar_proyecto()
    agregar(e.ramal, "ANE", "MT", "P002", 5)

    usos = e.codigo_en_uso("P002")
    assert any("RAMAL" in u.upper() for u in usos), \
        f"una fila propia en uso debe aparecer en codigo_en_uso: {usos}"
    e.renombrar_codigo("P002", "P099")
    assert any(it["codigo"] == "P099" for it in e.ramal["items"]["ANE"]["MT"])
    assert e._cant_propia("ANE")["MT"].get("P099", Decimal(0)) > 0


# ═════════════════════════════════════════════════════════════════════
#  12) UI offscreen
# ═════════════════════════════════════════════════════════════════════
def _app():
    from PySide6.QtWidgets import QApplication
    return QApplication.instance() or QApplication([])


def _principal(estado):
    """Construye Principal sin que `Estado()` lea la plantilla."""
    _app()
    original = main.cargar_bd_inicial
    main.cargar_bd_inicial = lambda: []
    try:
        p = main.Principal()
    finally:
        main.cargar_bd_inicial = original
    p.estado = estado
    p.tab = None
    p.refrescar_todo()
    return p


def _cerrar(p):
    """Cierra la ventana SIN el diálogo modal de «cambios sin guardar».

    `Principal.closeEvent` abre `tres_opciones(...)` cuando hay cambios
    pendientes; en una prueba headless ese diálogo dejaría el proceso colgado.
    """
    try:
        p.estado.cambios = False          # evita el diálogo de closeEvent
        p.close()
    except Exception:
        pass


def _escenario_ui():
    """2 proyectos + Ramal con ANE / ANE - VAR, una fila heredada y una propia."""
    e = estado_nuevo()
    p0 = e.proyectos[0]
    agregar(p0, "ANE", "MT", "P001", 3)
    e.activar_hoja(p0, "ANE_VAR")
    e.agregar_proyecto()
    e.activar_hoja(e.ramal, "VARIACION")
    agregar(e.ramal, "ANE", "MT", "P999", 4)
    return e


def test_ui_panel_del_ramal_arranca_plegado_y_se_abre_con_la_flecha():
    from PySide6.QtWidgets import QApplication
    e = _escenario_ui()
    p = _principal(e)
    try:
        p.show()
        QApplication.processEvents()
        p.seleccionar_tab(e.ramal, "ANE")
        QApplication.processEvents()

        assert p._panel_ramal is False
        assert p.panel.isVisible() is False, \
            "el panel de Materiales del Ramal debe arrancar plegado"
        assert p.btn_colapsar.text() == "▼"
        assert p.panel_g.isVisible() is True, "el grupo (con la flecha) sigue visible"

        p.btn_colapsar.click()
        QApplication.processEvents()
        assert p.panel.isVisible() is True, "la flecha debe desplegar el panel"
        assert p.btn_colapsar.text() == "▲"
        assert p._panel_ramal is True

        p.btn_colapsar.click()
        QApplication.processEvents()
        assert p.panel.isVisible() is False
        assert p._panel_ramal is False

        # En un proyecto el panel sigue arrancando desplegado.
        p.seleccionar_tab(e.proyectos[0], "ANE")
        QApplication.processEvents()
        assert p._panel_visible is True
        assert p.panel.isVisible() is True
    finally:
        _cerrar(p)


def test_ui_cantidad_editable_solo_en_filas_propias():
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QApplication
    e = _escenario_ui()
    p = _principal(e)
    try:
        p.show()
        QApplication.processEvents()
        p.seleccionar_tab(e.ramal, "ANE")
        QApplication.processEvents()

        filas = e.filas(e.ramal, "ANE")["MT"]
        indice = {f["codigo"]: i for i, f in enumerate(filas)}
        tabla = p.secciones["MT"][1]

        heredada = indice["P001"]
        propia = indice["P999"]
        assert not (tabla.item(heredada, 4).flags() & Qt.ItemIsEditable), \
            "una fila heredada no debe permitir editar la cantidad"
        assert (tabla.item(propia, 4).flags() & Qt.ItemIsEditable), \
            "una fila propia debe permitir editar la cantidad"

        # Control negativo: en un proyecto la cantidad sí es editable.
        p.seleccionar_tab(e.proyectos[0], "ANE")
        QApplication.processEvents()
        tabla0 = p.secciones["MT"][1]
        assert (tabla0.item(0, 4).flags() & Qt.ItemIsEditable)
    finally:
        _cerrar(p)


def test_ui_variacion_del_ramal_sin_que_los_proyectos_la_tengan():
    from PySide6.QtWidgets import QApplication
    e = _escenario_ui()
    assert all(not pr["hojas"]["VARIACION"] for pr in e.proyectos), \
        "ningún proyecto debe tener VARIACIÓN en este escenario"
    p = _principal(e)
    try:
        p.show()
        QApplication.processEvents()
        tabs = e.tabs()
        assert (e.ramal, "VARIACION") in tabs
        assert not any(pr is not e.ramal and h == "VARIACION" for pr, h in tabs)

        p.seleccionar_tab(e.ramal, "VARIACION")
        QApplication.processEvents()
        assert p.tab == (e.ramal, "VARIACION")
        # La hoja se dibuja sin errores y con sus 15 columnas.
        assert p.secciones["MT"][1].columnCount() == 15
    finally:
        _cerrar(p)


def test_ui_ramal_sin_texto_naranja_permanente():
    from PySide6.QtWidgets import QApplication
    e = _escenario_ui()
    p = _principal(e)
    try:
        p.show()
        QApplication.processEvents()
        for hoja in ("ANE", "ANE_VAR", "VARIACION"):
            p.seleccionar_tab(e.ramal, hoja)
            QApplication.processEvents()
            assert p.lbl_info.text() == "", \
                f"{hoja} del Ramal no debe mostrar aviso permanente"
            assert p.lbl_info.isVisible() is False
    finally:
        _cerrar(p)


def test_ui_no_deja_agregar_un_codigo_heredado():
    from PySide6.QtWidgets import QApplication
    e = _escenario_ui()
    p = _principal(e)
    try:
        p.show()
        QApplication.processEvents()
        p.seleccionar_tab(e.ramal, "ANE")
        QApplication.processEvents()

        p.elegir("P001")                 # viene de un proyecto
        p.cant.setText("2")
        p.validar_agregar()
        assert p.btn_add.isEnabled() is False, \
            "un código heredado no se captura desde el Ramal"
        assert "redondeo" in p.btn_add.toolTip().lower()

        p.elegir("P999")                 # no está en ningún proyecto
        p.cant.setText("2")
        p.validar_agregar()
        assert p.btn_add.isEnabled() is True, \
            "un código que no está en ningún proyecto sí se captura"
    finally:
        _cerrar(p)


def test_ui_agregar_en_el_ramal_guarda_una_fila_propia():
    from PySide6.QtWidgets import QApplication
    e = _escenario_ui()
    p = _principal(e)
    try:
        p.show()
        QApplication.processEvents()
        p.seleccionar_tab(e.ramal, "ANE")
        QApplication.processEvents()

        # Un código heredado: `agregar` no debe tocarlo.
        p.elegir("P001")
        p.cant.setText("9")
        p.validar_agregar()
        p.agregar()
        assert e._cant_propia("ANE")["MT"].get("P001", Decimal(0)) == 0

        # Un código libre: sí entra como fila propia.
        p.elegir("P002")
        p.cant.setText("4")
        p.validar_agregar()
        p.agregar()
        assert e._cant_propia("ANE")["MT"]["P002"] == Decimal(4)
        assert e._cant_propia("ANE")["MT"].get("P002", Decimal(0)) > 0
    finally:
        _cerrar(p)


def test_ui_borrar_filas_del_ramal():
    from PySide6.QtWidgets import QApplication
    e = _escenario_ui()
    p = _principal(e)
    avisos = []
    original_msg = main.msg
    main.msg = lambda *a, **k: avisos.append(a[1] if len(a) > 1 else "")
    try:
        p.show()
        QApplication.processEvents()
        p.seleccionar_tab(e.ramal, "ANE")
        QApplication.processEvents()
        filas = e.filas(e.ramal, "ANE")["MT"]
        indice = {f["codigo"]: i for i, f in enumerate(filas)}
        tabla = p.secciones["MT"][1]

        # Heredada: se avisa y no se borra.
        tabla.clearSelection()
        tabla.selectRow(indice["P001"])
        p.eliminar_filas("MT")
        assert any("proyecto" in a.lower() for a in avisos), avisos
        assert "P001" in por_codigo(e, e.ramal, "ANE", "MT")

        # Propia: se borra sin aviso nuevo.
        avisos.clear()
        tabla.clearSelection()
        tabla.selectRow(indice["P999"])
        p.eliminar_filas("MT")
        assert not avisos, f"no debía avisar al borrar una fila propia: {avisos}"
        assert "P999" not in por_codigo(e, e.ramal, "ANE", "MT")
    finally:
        main.msg = original_msg
        _cerrar(p)


# ═════════════════════════════════════════════════════════════════════
#  RUNNER sin pytest (python tests/test_ramal_propio.py)
# ═════════════════════════════════════════════════════════════════════
def _runner():
    pruebas = [(n, f) for n, f in sorted(globals().items())
               if n.startswith("test_") and callable(f)]
    fallas = []
    for nombre, fn in pruebas:
        try:
            fn()
            print(f"PASA   {nombre}")
        except Exception as e:
            fallas.append(nombre)
            print(f"FALLA  {nombre}  -> {e!r}")
            traceback.print_exc()
    print()
    print(f"RESULTADO: {len(pruebas) - len(fallas)}/{len(pruebas)} PASA")
    if fallas:
        print("FALLARON:", fallas)
    return 1 if fallas else 0


if __name__ == "__main__":
    sys.exit(_runner())
