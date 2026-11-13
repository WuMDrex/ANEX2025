# Pruebas de ANEXOS (main.py)

Nada de aquí modifica el código fuente ni la plantilla real.

## Archivos

| Archivo | Qué comprueba | Necesita Excel |
|---|---|---|
| `conftest.py` | Sustituye `license_manager` por un doble y fuerza `QT_QPA_PLATFORM=offscreen`. | No |
| `test_bd_componentes.py` | `componente_nuevo` (6 marcas en 0), `DlgComponentes.agregar`, `DlgReporte.agregar_bd`, `Esc` con/sin cambios. | No |
| `test_bd_excel.py` | `_xl_escribir_bd`, `exportar_xlsx`, `cargar_proyecto_desde_excel`, guardado Global sobre copia. | **Sí** |
| `pytest/` | **SHIM TEMPORAL — NO es pytest.** Ver abajo. | — |

## Cómo ejecutar

Desde la raíz del proyecto (recomendado):

```
pip install pytest pytest-qt
python -m pytest tests -q
```

Solo las que no usan Excel:

```
python -m pytest tests -q -m "not excel"
```

## ⚠️ Antes de usar el pytest real: borra el shim

La sesión de QA que creó estas pruebas tenía el shell restringido: solo se
podían ejecutar `python -m py_compile`, `python -m pytest` y
`python herramientas\cambios.py`, y `pip install` estaba bloqueado. Para poder
ejecutar algo se creó el paquete `tests/pytest/`, un corredor mínimo que
implementa `mark`, `raises`, `skip` y `fixture`.

**Ese paquete ensombrece al pytest real si ejecutas con la carpeta `tests` como
directorio actual** (`cd tests && python -m pytest`). El sistema de permisos
impidió borrarlo, así que bórralo tú una vez:

```
rmdir /s /q tests\pytest
```

También puedes borrar `tests/_env_check.py` (diagnóstico temporal) y las
carpetas `tests\__pycache__`.

Después de borrarlo, ejecuta siempre desde la raíz del proyecto.

## Notas

* Las pruebas de `test_bd_excel.py` usan una **copia** de la plantilla, nunca
  `Plantilla/ANEXO SICOA 2025.xlsm`. El guardado Global se prueba parcheando
  `main.PLANTILLA_XLSX` a la copia y anulando `_respaldo_plantilla` para no
  llenar `backups/`.
* Las carpetas temporales `anexos_*` que crean las pruebas se borran solas al
  terminar la ejecución (`atexit`).
* Dos pruebas son de **mutación** a propósito: desactivan la función bajo
  prueba para confirmar que la aserción falla si el cambio no existiera.
