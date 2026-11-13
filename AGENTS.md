# Reglas compartidas del equipo WuMDrex

Estas reglas las leen TODOS los agentes (OpenCode carga este archivo automáticamente).
Los prompts individuales solo añaden lo específico de cada rol.

## Contexto
- Desarrollador: **WuMDrex**. Lenguaje: **Python** (PySide6, openpyxl, xlwings, pandas, PyInstaller).
- Plataforma de destino: **Windows** (Excel por COM con xlwings/pythoncom, AppUserModelID, .exe).
- Cada programa vive en su propia carpeta y tiene UN archivo principal monolítico:
  - **Anexos** (gestor de hojas Excel): `RAIZ.py`  ← ajusta si tu archivo se llama distinto
  - **Poste a Poste** (captura de postes + vista CAD): `main.py`  ← ajusta si tu archivo se llama distinto
- Antes de empezar, el Director identifica sobre QUÉ programa se trabaja. Nunca mezclar los dos.

## Reglas inquebrantables
1. **Cirugía:** prohibido modificar, formatear, reordenar o eliminar código que no esté directamente
   relacionado con la tarea. Nada de "ya que estoy". Si ves algo mal fuera del alcance, se REPORTA
   como observación; no se arregla.
2. **Backup obligatorio:** antes de la primera edición de una tarea se ejecuta
   `python herramientas/respaldo.py <archivo_principal>`. Nunca se editan backups.
3. **Idioma:** variables, funciones, comentarios, docstrings y textos de interfaz en **español**.
4. **Hilo principal sagrado:** nada lento en la UI. Excel (xlwings/COM), openpyxl sobre archivos
   grandes, red, `time.sleep` y bucles largos van en `QThread` + `Signal` (patrón `ActualizarWorker`).
5. **Banderas de estado:** respetar `actualizando`, `_cargando` y `blockSignals`. Toda bandera que se
   activa debe desactivarse en TODAS las rutas de salida (usar `try/finally` en código nuevo).
6. **UI separada de datos:** la lógica de negocio (cálculos, lectura/escritura de Excel) no depende de
   widgets. Sin dependencias circulares entre módulos.
7. **Dinero con `Decimal`**, nunca `float`, y con el redondeo ya definido (`redondear_monto`).
8. **Licencia intocable:** `license_manager` / `check_license()` no se quita, no se mueve, no se evita.
9. **Recursos empaquetables:** archivos de solo lectura (logo, plantilla, machote) con
   `getattr(sys, "_MEIPASS", BASE_DIR)`. Datos que se ESCRIBEN (backups, configuración) junto al .exe,
   nunca dentro de `_MEIPASS` (se borra al cerrar).

## Protocolo de construcción (quien edite código debe seguirlo)
1. `python herramientas/respaldo.py <archivo>`  → anotar la ruta del backup creado.
2. Editar SOLO lo indicado en el plan (funciones/clases permitidas).
3. `python -m py_compile <archivo>` → debe pasar.
4. `python herramientas/cambios.py <archivo>` → copiar la salida en el reporte. Si aparece una función
   que no estaba en el plan: revertirla antes de reportar.
5. Reportar con el formato de la sección siguiente.

## Formato de reporte de un subagente (siempre este orden)
- **Resultado:** HECHO / BLOQUEADO / PARCIAL
- **Backup:** ruta
- **Funciones tocadas:** salida de `cambios.py`
- **Verificación:** qué comprobaste y cómo (comando + resultado real; si no pudiste, di "NO VERIFICADO")
- **Observaciones fuera de alcance:** problemas detectados que NO se tocaron
- **Dudas:** si las hay, como opciones A / B / C con pros y contras, marcando cuál recomiendas

## Estilo observado en el código de WuMDrex (mantenerlo)
- Secciones separadas con banners `# ═════…` y un título en MAYÚSCULAS.
- Docstrings que explican el POR QUÉ (reglas de negocio de Excel, límites de nombres, etc.).
- Marcadores `TODO-EXCEL` para lo que depende de la plantilla real.
- Constantes en MAYÚSCULAS arriba del archivo; helpers pequeños (`msg`, `confirmar`, `boton`, `norm`).
- Tema día/noche centralizado en una función de QSS; atajos registrados en UNA tabla que alimenta la
  ventana de ayuda (F1). Si se agrega un atajo, se agrega a esa tabla y se revisa que no choque.
- Mensajes al usuario: cortos, en español, con `QMessageBox` en botones "Sí/No/Aceptar" en español.
