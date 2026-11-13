# Pruebas de T1 + T2 del Ramal (Anexos · `main.py`)

Verifican los dos cambios del Ramal contra `backups/Backup_11.11.2026_104046_main.py`.
**No** tocan `main.py`, **no** escriben en `Plantilla/ANEXO SICOA 2025.xlsm` y todos los
`.xlsx` de salida viven en `%TEMP%\opencode\ramal_t1t2\`.

## Comandos (desde la raíz del proyecto)

```bat
:: A + B + C (sintaxis, aislamiento, lógica pura + UI offscreen)
python tests\run_ramal_t1t2.py

:: Todo lo anterior + D (roundtrip real con Excel/COM)
python tests\run_ramal_t1t2.py --excel

:: Alternativa con pytest (los 4 de Excel llevan el marcador `excel`)
python -m pytest tests\test_ramal_propio.py -q
python -m pytest tests\test_ramal_propio_excel.py -q -m excel

:: Sólo las suites de este cambio
python tests\test_ramal_propio.py
python tests\test_ramal_propio_excel.py
```

## Qué debe salir

| Paso | Esperado |
|------|----------|
| A · `py_compile main.py` | código de salida **0** |
| B · `cambios.py main.py backups\Backup_11.11.2026_104046_main.py` | **7 agregadas / 24 modificadas / 0 eliminadas** |
| C · `tests\test_ramal_propio.py` | **26/26 PASA** (19 de lógica + 7 de UI offscreen) |
| D · `tests\test_ramal_propio_excel.py` | **4/4 PASA** (Casos 1-3 + plantilla intacta / 0 `EXCEL.EXE`) |

## Cobertura

**`tests/test_ramal_propio.py`** (sin Excel: la BD se inyecta con `Estado.__new__`)

- **T1** — el Ramal con 1 proyecto no activa nada; la siembra ocurre **una sola vez**
  (transición 1→2); `VARIACION`/`OC`/`OTS`/`ATE` del Ramal son independientes de los
  proyectos; sin `ANE - VAR` en ningún proyecto se apagan solas; `Estado` no bloquea
  (el aviso vive en `DlgEdicion.toggle`).
- **T2** — `_cant_propia(...) > 0` / `codigo_heredado` / `filas()["propio"/"cant_propia"]`; si un
  código propio aparece después en un proyecto **se suman** las cantidades; la fila
  propia sobrevive al quitar el proyecto; kilometraje con 3 decimales; `_xl_literal_cantidad`
  sin notación científica; cuantización `0.001`.
- **Regresión** — orden SICOA, `redondear_monto`/`texto_total`/`MODOS_REDONDEO`, y que las
  filas de un proyecto **no** traigan las claves `propio`/`cant_propia`.
- **UI offscreen** — el panel del Ramal arranca plegado y la flecha lo abre; la cantidad
  sólo es editable en filas propias; `VARIACIÓN` del Ramal sin que los proyectos la
  tengan (15 columnas); sin texto naranja permanente; no se captura un código heredado;
  `agregar` guarda filas propias y `eliminar_filas` avisa y respeta las heredadas.

**`tests/test_ramal_propio_excel.py`** (marcador `excel`, requiere Excel instalado)

- **Caso 1** sin filas propias → `SUMIF` puro, sin `+literal`, sin celdas `#...`,
  firma del Ramal = 1000, roundtrip sin filas propias ni duplicados.
- **Caso 2** una fila propia + `VARIACION` sólo en el Ramal → exactamente un `+5`,
  la hoja `VARIACIÓN - RAMAL` existe sin errores, roundtrip devuelve `P999 = 5`, la
  `VARIACION` del Ramal sigue activa y los proyectos siguen sin ella.
- **Caso 3** al mismo libro se le quita el `+literal` (formato anterior) → no inventa
  filas propias ni filas fantasma (`["P001","K001"]`).
- **Cierre** hash/tamaño de la plantilla sin cambios y ningún `EXCEL.EXE` huérfano.

## Notas

- `_xl_leer_items` deja en el Ramal filas propias con cantidad `0.000` cuando la resta
  `leído − proyectos` da cero; por eso las pruebas comprueban «ninguna propia **en
  positivo**», no «diccionario vacío». `_cant_propia(...) > 0` ya aplica ese criterio.
- `Principal.closeEvent` abre un diálogo modal si hay cambios pendientes; las pruebas
  cierran con el helper `_cerrar()`, que pone `estado.cambios = False` antes de `close()`.
- Ninguna prueba abre una ventana visible: `QT_QPA_PLATFORM=offscreen` (lo fija
  `tests/conftest.py` y cada archivo).
