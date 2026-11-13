"""Corredor mínimo de pruebas (sustituto temporal de pytest)."""

import importlib.util
import os
import sys
import time
import traceback


def _archivos(objetivos):
    rutas = []
    for obj in objetivos:
        if os.path.isdir(obj):
            for nombre in sorted(os.listdir(obj)):
                if nombre.startswith("test_") and nombre.endswith(".py"):
                    rutas.append(os.path.join(obj, nombre))
        elif obj.endswith(".py") and os.path.exists(obj):
            rutas.append(obj)
    return rutas


def _cargar(ruta):
    ruta = os.path.abspath(ruta)
    nombre = os.path.splitext(os.path.basename(ruta))[0]
    carpeta = os.path.dirname(ruta)
    if carpeta not in sys.path:
        sys.path.insert(0, carpeta)
    spec = importlib.util.spec_from_file_location(nombre, ruta)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[nombre] = mod
    spec.loader.exec_module(mod)
    return mod


def main(args=None):
    args = list(sys.argv[1:] if args is None else args)
    objetivos = [a for a in args if not a.startswith("-")] or ["."]
    archivos = _archivos(objetivos)
    if not archivos:
        print("No se encontraron archivos test_*.py en:", objetivos)
        return 2

    pasaron = fallaron = omitidos = 0
    fallos = []
    t0 = time.time()

    for ruta in archivos:
        print(f"\n===== {os.path.basename(ruta)} =====")
        try:
            mod = _cargar(ruta)
        except Exception:
            print("ERROR al importar el archivo:")
            traceback.print_exc()
            fallaron += 1
            fallos.append((ruta, "<import>", traceback.format_exc()))
            continue

        pruebas = [(n, getattr(mod, n)) for n in dir(mod)
                   if n.startswith("test") and callable(getattr(mod, n))
                   and getattr(getattr(mod, n), "__module__", None) == mod.__name__]
        pruebas.sort()

        for nombre, fn in pruebas:
            if getattr(fn, "_es_fixture", False):
                continue
            if fn.__code__.co_argcount:
                print(f"OMITIDA {nombre} (necesita fixtures: {fn.__code__.co_varnames[:fn.__code__.co_argcount]})")
                omitidos += 1
                continue
            inicio = time.time()
            try:
                fn()
            except Exception as e:
                if type(e).__name__ == "Skipped":
                    print(f"OMITIDA {nombre}: {e}")
                    omitidos += 1
                    continue
                fallaron += 1
                tb = traceback.format_exc()
                print(f"FALLA   {nombre}  ({time.time() - inicio:.1f}s)")
                print(tb)
                fallos.append((ruta, nombre, tb))
            else:
                pasaron += 1
                print(f"PASA    {nombre}  ({time.time() - inicio:.1f}s)")

    print("\n" + "=" * 60)
    print(f"RESUMEN: {pasaron} pasaron, {fallaron} fallaron, {omitidos} omitidas "
          f"({time.time() - t0:.1f}s)")
    if fallos:
        print("\nFALLOS:")
        for ruta, nombre, _ in fallos:
            print(f"  - {os.path.basename(ruta)}::{nombre}")
    return 1 if fallaron else 0
