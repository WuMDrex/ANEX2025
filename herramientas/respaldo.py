"""
Copia de seguridad de un archivo antes de editarlo.

Uso:
    python herramientas/respaldo.py main.py

Crea `backups/Backup_dd.mm.aaaa_HHMMSS_<archivo>.py` y muestra la ruta creada.
Los respaldos NUNCA se editan: son el punto de retorno de la cirugía.
"""
import os
import shutil
import sys
from datetime import datetime

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def main():
    if len(sys.argv) < 2:
        print("Uso: python herramientas/respaldo.py <archivo>")
        return 2
    origen = sys.argv[1]
    if not os.path.isabs(origen):
        origen = os.path.join(RAIZ, origen)
    if not os.path.isfile(origen):
        print("No existe el archivo:", origen)
        return 1
    carpeta = os.path.join(RAIZ, "backups")
    os.makedirs(carpeta, exist_ok=True)
    nombre = os.path.splitext(os.path.basename(origen))[0]
    destino = os.path.join(carpeta, "Backup_%s_%s.py" % (datetime.now().strftime("%d.%m.%Y_%H%M%S"), nombre))
    shutil.copyfile(origen, destino)
    print("Respaldo creado:", destino)
    return 0


if __name__ == "__main__":
    sys.exit(main())
