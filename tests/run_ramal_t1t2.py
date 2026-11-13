# -*- coding: utf-8 -*-
"""
Un solo comando para verificar T1 + T2 del Ramal (Anexos · main.py).

Hace, en este orden:
  A) `python -m py_compile main.py`
  B) `python herramientas/cambios.py main.py backups/Backup_11.11.2026_104046_main.py`
     (esperado: 7 agregadas / 24 modificadas / 0 eliminadas)
  C) Lógica pura + UI offscreen  → ejecuta `tests/test_ramal_propio.py`
  D) [opcional] Roundtrip real con Excel/COM → `tests/test_ramal_propio_excel.py`

Uso (desde la raíz del proyecto):
    python tests\\run_ramal_t1t2.py            # A + B + C
    python tests\\run_ramal_t1t2.py --excel    # A + B + C + D
    python tests\\run_ramal_t1t2.py --solo-pruebas
"""

import os
import subprocess
import sys
import traceback

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CARPETA_TESTS = os.path.dirname(os.path.abspath(__file__))
BACKUP = os.path.join(RAIZ, "backups", "Backup_11.11.2026_104046_main.py")
PY = sys.executable or "python"

for ruta in (RAIZ, CARPETA_TESTS):
    if ruta not in sys.path:
        sys.path.insert(0, ruta)

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

ARGS = [a.lower() for a in sys.argv[1:]]
SOLO_PRUEBAS = "--solo-pruebas" in ARGS
CON_EXCEL = "--excel" in ARGS


def _paso(titulo):
    print()
    print("=" * 70)
    print(titulo)
    print("=" * 70)


def _correr(cmd):
    print("$", " ".join(cmd))
    r = subprocess.run(cmd, cwd=RAIZ, capture_output=True, text=True)
    if r.stdout:
        print(r.stdout.rstrip())
    if r.stderr:
        print("STDERR:", r.stderr.rstrip())
    print(f"-> código de salida: {r.returncode}")
    return r.returncode


# ═════════════════════════════════════════════════════════════════════
#  A · sintaxis
# ═════════════════════════════════════════════════════════════════════
fallas = []
if not SOLO_PRUEBAS:
    _paso("A · sintaxis (py_compile main.py)")
    if _correr([PY, "-m", "py_compile", "main.py"]) != 0:
        fallas.append("py_compile")

    # ═════════════════════════════════════════════════════════════════
    #  B · aislamiento del cambio
    # ═════════════════════════════════════════════════════════════════
    _paso("B · aislamiento (cambios.py contra el backup)")
    if not os.path.exists(BACKUP):
        print("NO EXISTE el backup:", BACKUP)
        fallas.append("cambios.py (backup ausente)")
    elif _correr([PY, os.path.join("herramientas", "cambios.py"),
                  "main.py", BACKUP]) != 0:
        fallas.append("cambios.py")

# ═════════════════════════════════════════════════════════════════════
#  C · lógica pura + UI offscreen
# ═════════════════════════════════════════════════════════════════════
_paso("C · lógica pura de Estado + UI offscreen")
try:
    import test_ramal_propio as T
    pruebas = [(n, f) for n, f in sorted(vars(T).items())
               if n.startswith("test_") and callable(f)]
    ok = 0
    for nombre, fn in pruebas:
        try:
            fn()
            ok += 1
            print(f"PASA   {nombre}")
        except Exception as e:
            fallas.append(nombre)
            print(f"FALLA  {nombre}  -> {e!r}")
            traceback.print_exc()
    print(f"\nC: {ok}/{len(pruebas)} PASA")
except Exception as e:
    fallas.append("import test_ramal_propio")
    print("No se pudo ejecutar la suite de lógica:", repr(e))
    traceback.print_exc()

# ═════════════════════════════════════════════════════════════════════
#  D · Excel real (opcional)
# ═════════════════════════════════════════════════════════════════════
if CON_EXCEL:
    _paso("D · roundtrip real con Excel/COM")
    try:
        import test_ramal_propio_excel as X
        pruebas = [(n, f) for n, f in sorted(vars(X).items())
                   if n.startswith("test_") and callable(f)]
        ok = 0
        for nombre, fn in pruebas:
            try:
                fn()
                ok += 1
                print(f"PASA   {nombre}")
            except Exception as e:
                fallas.append(nombre)
                print(f"FALLA  {nombre}  -> {e!r}")
                traceback.print_exc()
        print(f"\nD: {ok}/{len(pruebas)} PASA")
    except Exception as e:
        fallas.append("import test_ramal_propio_excel")
        print("No se pudo ejecutar la suite de Excel:", repr(e))
        traceback.print_exc()

# ═════════════════════════════════════════════════════════════════════
#  RESUMEN
# ═════════════════════════════════════════════════════════════════════
_paso("RESUMEN")
if fallas:
    print("FALLA ·", len(fallas), "problema(s):")
    for f in fallas:
        print("  ·", f)
    sys.exit(1)
print("PASA · todo lo ejecutado terminó sin fallas")
sys.exit(0)
