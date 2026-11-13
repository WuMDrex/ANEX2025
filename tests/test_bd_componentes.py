# -*- coding: utf-8 -*-
"""
Pruebas del cambio "BD de componentes nuevos + Esc con cambios".

Cubre:
  · componente_nuevo() devuelve las seis marcas en 0.
  · DlgComponentes.agregar deja el componente en 0, cambios=True y 💾 habilitado.
  · DlgReporte.agregar_bd también usa componente_nuevo (marcas en 0).
  · Principal.keyPressEvent: Esc con cambios pregunta UNA vez; sin cambios no cierra.

Se ejecutan sin pantalla (QT_QPA_PLATFORM=offscreen). NO tocan Excel.
"""

import os
import sys
import types

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RAIZ not in sys.path:
    sys.path.insert(0, RAIZ)

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

# Licencia: doble de prueba. Nunca se importa ni se modifica la real.
_falso = types.ModuleType("license_manager")
_falso.check_license = lambda: True
sys.modules.setdefault("license_manager", _falso)

import main                                     # noqa: E402

# Se capturan las referencias REALES antes de que ninguna prueba las parchee.
_EstadoReal = main.Estado
_AppReal = main.QApplication
_tres_opciones_real = main.tres_opciones

MARCA = ("MT", "MTBT", "BT", "G", "v76", "v199")


def estado_vacio():
    """Estado real pero sin abrir Excel (la BD se inyecta a mano)."""
    e = _EstadoReal.__new__(_EstadoReal)
    e.bd = []
    e.reset()
    return e


def app_qt():
    from PySide6.QtWidgets import QApplication
    return QApplication.instance() or QApplication([])


# ═════════════════════════════════════════════════════════════════════
#  1) componente_nuevo: las seis marcas en 0
# ═════════════════════════════════════════════════════════════════════
def test_componente_nuevo_marcas_en_cero():
    import main
    c = main.componente_nuevo("999999", "POSTE NUEVO", "U", "1500")
    for k in MARCA:
        assert c[k] == 0, f"{k} debería ser 0 y es {c[k]!r}"
    assert c["codigo"] == "999999"
    assert c["descripcion"] == "POSTE NUEVO"
    assert c["unidad"] == "U"
    assert c["precio"] == "1500.00", c["precio"]


def test_componente_nuevo_no_se_clasifica_por_la_descripcion():
    """Control negativo: la descripción SÍ sugiere voltaje, y aun así va en 0.

    Si `componente_nuevo` reutilizara `clasificar_componente`, v76 quedaría en 1
    y esta prueba fallaría; así se comprueba que la prueba discrimina.
    """
    import main
    texto = "TRANSFORMADOR 13.2 kV / 34.5 kV PARA POSTE"

    sugerido = main.clasificar_componente("999999", texto)
    assert sugerido["v76"] == 1 or sugerido["MT"] == 1, \
        "control inválido: clasificar_componente no marca nada para este texto"

    nuevo = main.componente_nuevo("999999", texto, "U", "10")
    for k in MARCA:
        assert nuevo[k] == 0, f"{k} debería ser 0 y es {nuevo[k]!r}"


def test_componente_nuevo_precio_vacio_es_cero():
    import main
    c = main.componente_nuevo("1", "X", "U", "")
    assert c["precio"] == "0.00", c["precio"]
    for k in MARCA:
        assert c[k] == 0


# ═════════════════════════════════════════════════════════════════════
#  2) DlgComponentes.agregar / DlgReporte.agregar_bd
# ═════════════════════════════════════════════════════════════════════
def test_dlg_componentes_agregar_marcas_en_cero():
    import main
    app_qt()
    e = estado_vacio()
    d = main.DlgComponentes(None, e)
    try:
        antes = len(d.datos)
        assert d.cambios is False
        assert d.btn_guardar.isEnabled() is False

        d.cod.setText("999999")
        d.desc.setText("EQUIPO NUEVO DE PRUEBA")
        d.uni.setText("U")
        d.precio.setText("10.00")
        d.agregar()

        assert len(d.datos) == antes + 1, (antes, len(d.datos))
        nuevo = d.datos[-1]
        assert nuevo["codigo"] == "999999"
        for k in MARCA:
            assert nuevo[k] == 0, f"{k} debería ser 0 y es {nuevo[k]!r}"
        assert d.cambios is True, "agregar() no marcó cambios"
        assert d.btn_guardar.isEnabled() is True, "💾 quedó deshabilitado"
    finally:
        d.deleteLater()


def test_dlg_componentes_agregar_rechaza_codigo_no_numerico():
    import main
    app_qt()
    d = main.DlgComponentes(None, estado_vacio())
    try:
        antes = len(d.datos)
        d.cod.setText("ABC-1")
        d.desc.setText("EQUIPO")
        d.uni.setText("U")
        d.precio.setText("5")
        # msg() abre un QMessageBox modal: se neutraliza para no bloquear.
        original = main.msg
        main.msg = lambda *a, **k: None
        try:
            d.agregar()
        finally:
            main.msg = original
        assert len(d.datos) == antes, "agregó un código no numérico"
        assert d.cambios is False
    finally:
        d.deleteLater()


def test_dlg_reporte_agregar_bd_marcas_en_cero():
    app_qt()
    e = estado_vacio()
    d = main.DlgReporte(None, {"totales": [], "externos": [], "actualizadas": []}, e)
    try:
        from PySide6.QtWidgets import QPushButton
        btn = QPushButton("Agregar a la BD")
        d.agregar_bd(btn, "888888", "COMPONENTE EXTERNO", "U", "25.50")

        c = e.comp("888888")
        assert c is not None, "no entró a la BD"
        for k in MARCA:
            assert c[k] == 0, f"{k} debería ser 0 y es {c[k]!r}"
        assert e.cambios is True
        assert btn.isEnabled() is False
        assert "Agregado" in btn.text()
    finally:
        d.deleteLater()


# ═════════════════════════════════════════════════════════════════════
#  5) Esc en la ventana principal
# ═════════════════════════════════════════════════════════════════════
class _AppFalsa(_AppReal):
    """QApplication con quit() contado; instance() y el resto quedan intactos."""

    salidas = 0

    @staticmethod
    def quit():
        _AppFalsa.salidas += 1


def _fabricar_espia():
    class Espia(main.Principal):
        veces_close = 0

        def closeEvent(self, ev):
            Espia.veces_close += 1
            super().closeEvent(ev)

    return Espia


def _escenario(cambios, respuesta):
    """Construye la ventana con los parches puestos; devuelve utilidades."""
    from PySide6.QtTest import QTest
    from PySide6.QtCore import Qt

    app_qt()
    estado = estado_vacio()
    estado.cambios = cambios

    main.Estado = lambda: estado
    main.tres_opciones = lambda *a, **k: respuesta
    main.QApplication = _AppFalsa
    _AppFalsa.salidas = 0

    try:
        Espia = _fabricar_espia()
        Espia.veces_close = 0
        w = Espia()
        w.show()
        app_qt().processEvents()
    except Exception:
        main.Estado = _EstadoReal
        main.tres_opciones = _tres_opciones_real
        main.QApplication = _AppReal
        raise

    def enviar():
        QTest.keyClick(w, Qt.Key_Escape)
        app_qt().processEvents()

    def restaurar():
        w.hide()
        w.deleteLater()
        app_qt().processEvents()
        main.Estado = _EstadoReal
        main.tres_opciones = _tres_opciones_real
        main.QApplication = _AppReal

    return w, Espia, enviar, restaurar


def test_esc_sin_cambios_no_cierra():
    w, Espia, enviar, restaurar = _escenario(cambios=False, respuesta=2)
    try:
        enviar()
        assert Espia.veces_close == 0, \
            f"Esc sin cambios llamó closeEvent {Espia.veces_close} vez/veces"
        assert _AppFalsa.salidas == 0
    finally:
        restaurar()


def test_esc_con_cambios_pregunta_una_sola_vez():
    w, Espia, enviar, restaurar = _escenario(cambios=True, respuesta=2)
    try:
        enviar()
        assert Espia.veces_close == 1, \
            f"closeEvent se llamó {Espia.veces_close} vez/veces (se esperaba 1)"
        assert _AppFalsa.salidas == 0, "Cancelar no debe salir del programa"
    finally:
        restaurar()


def test_esc_con_cambios_salir_sin_guardar_sale_una_vez():
    w, Espia, enviar, restaurar = _escenario(cambios=True, respuesta=1)
    try:
        enviar()
        assert Espia.veces_close == 1, \
            f"closeEvent se llamó {Espia.veces_close} vez/veces (se esperaba 1)"
        assert _AppFalsa.salidas == 1, \
            f"QApplication.quit se llamó {_AppFalsa.salidas} vez/veces (se esperaba 1)"
    finally:
        restaurar()
