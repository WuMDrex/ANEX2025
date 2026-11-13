"""
Lista las funciones y clases que cambiaron respecto a un respaldo.

Uso:
    python herramientas/cambios.py main.py [respaldo]

Sin el segundo argumento se compara contra el respaldo más reciente de
`backups/`. Compara el código fuente de cada función, método y bloque de
atributos de clase, e informa AGREGADAS, MODIFICADAS y ELIMINADAS. Sirve para
comprobar que una cirugía tocó solo las funciones del plan aprobado.
"""
import ast
import os
import sys

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _ruta(valor):
    return valor if os.path.isabs(valor) else os.path.join(RAIZ, valor)


def _respaldo_reciente():
    carpeta = os.path.join(RAIZ, "backups")
    if not os.path.isdir(carpeta):
        return None
    candidatos = [os.path.join(carpeta, n) for n in os.listdir(carpeta)
                  if n.lower().endswith(".py")]
    return max(candidatos, key=os.path.getmtime) if candidatos else None


def _piezas(ruta):
    """Devuelve {nombre: (linea, fuente)} de funciones, métodos y atributos."""
    with open(ruta, encoding="utf-8") as f:
        texto = f.read()
    salida = {}
    for nodo in ast.parse(texto).body:
        if isinstance(nodo, (ast.FunctionDef, ast.AsyncFunctionDef)):
            salida[nodo.name] = (nodo.lineno, ast.get_source_segment(texto, nodo) or "")
        elif isinstance(nodo, ast.ClassDef):
            resto = [n for n in nodo.body
                     if not isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))]
            if resto:
                fuente = "\n".join(ast.get_source_segment(texto, n) or "" for n in resto)
                salida[nodo.name + ".<atributos>"] = (nodo.lineno, fuente)
            for hijo in nodo.body:
                if isinstance(hijo, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    salida["%s.%s" % (nodo.name, hijo.name)] = (
                        hijo.lineno, ast.get_source_segment(texto, hijo) or "")
    return salida


def main():
    if len(sys.argv) < 2:
        print("Uso: python herramientas/cambios.py <archivo> [respaldo]")
        return 2
    actual = _ruta(sys.argv[1])
    respaldo = _ruta(sys.argv[2]) if len(sys.argv) > 2 else _respaldo_reciente()
    if not respaldo or not os.path.isfile(respaldo):
        print("No se encontró el respaldo para comparar.")
        return 1

    antes = _piezas(respaldo)
    ahora = _piezas(actual)
    agregadas = [n for n in ahora if n not in antes]
    eliminadas = [n for n in antes if n not in ahora]
    modificadas = [n for n in ahora if n in antes and ahora[n][1].rstrip() != antes[n][1].rstrip()]

    print("Archivo   :", os.path.relpath(actual, RAIZ))
    print("Respaldo  :", os.path.relpath(respaldo, RAIZ))
    print("")
    print("AGREGADAS   (%d): %s" % (len(agregadas), ", ".join(sorted(agregadas)) or "-"))
    print("MODIFICADAS (%d): %s" % (len(modificadas), ", ".join(sorted(modificadas)) or "-"))
    print("ELIMINADAS  (%d): %s" % (len(eliminadas), ", ".join(sorted(eliminadas)) or "-"))
    print("")
    print("FUNCIONES TOCADAS (%d):" % (len(agregadas) + len(modificadas) + len(eliminadas)))
    for nombre in sorted(agregadas):
        print("  + %-45s linea %d" % (nombre, ahora[nombre][0]))
    for nombre in sorted(modificadas):
        print("  ~ %-45s linea %d" % (nombre, ahora[nombre][0]))
    for nombre in sorted(eliminadas):
        print("  - %s" % nombre)
    return 0


if __name__ == "__main__":
    sys.exit(main())
