"""Configuración común de las pruebas de ANEXOS.

· Sustituye `license_manager` por un módulo falso ANTES de importar `main.py`
  (el real no se toca ni se modifica).
· Fuerza Qt en modo offscreen: nunca se abre una ventana visible.
"""

import os
import sys
import types

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RAIZ not in sys.path:
    sys.path.insert(0, RAIZ)

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

_falso = types.ModuleType("license_manager")
_falso.check_license = lambda: True
_falso.__doc__ = "Doble de prueba: no valida licencia."
sys.modules.setdefault("license_manager", _falso)


def pytest_configure(config):
    """Registra el marcador `excel` para que pytest no avise por él."""
    config.addinivalue_line(
        "markers", "excel: requiere Excel instalado (xlwings/COM)")
