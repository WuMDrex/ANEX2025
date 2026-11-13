# Anexos PROASE — proyecto nuevo

Aplicación de escritorio (PySide6) para armar y controlar el anexo de precios
de proyectos de distribución eléctrica sobre un libro Excel con macros.

> `RAIZ.py` es **solo referencia de estilo**. No se modifica ni se reutiliza su lógica
> de postes: este es un programa distinto.

## Estado

| Fase | Estado |
|---|---|
| Auditoría del libro real | ✅ Terminada |
| Núcleo puro (texto, montos, validaciones, tema) | ✅ Terminado |
| Lector Excel (solo lectura) | ✅ Validado: cuadra con el libro en los 7 anexos |
| Ventana Principal + Componentes + Configuración/Edición/Visualización | ✅ Terminadas |
| Estado en memoria (`nucleo/estado.py`) | ✅ Terminado |
| Escritura del Excel (xlwings) | ✅ Terminada y probada con Excel real |
| Ramal automático | ✅ Terminado |
| Guardado de Componentes en `CODIGOS` (Local/Global) | ✅ Terminado y probado con Excel real |
| Ayuda `F1` y atajos | ✅ Terminada |
| Pruebas automáticas | ✅ 73 pruebas (~24 s, 2 saltadas); integración Excel aparte |
| Pendiente | Probar el ciclo abrir→editar→guardar con el libro real |

## Estructura

```
ANEXOS/
├── MEMORY.md                  # backlog y bitácora del Director
├── LEEME.md
├── main.py                    # punto de entrada
├── requirements.txt
├── backups/                   # Backup_DD.MM.YYYY.py / .xlsm (no se editan)
├── plantillas/
│   └── Plantilla_Anexos.xlsm  # plantilla limpia (conserva macros y dibujos)
├── docs/
│   ├── AUDITORIA_EXCEL.md     # hallazgos del .xlsm real
│   └── vba_Congfi.bas         # VBA extraído (solo lectura, protegido)
├── nucleo/                    # lógica pura, sin PySide6 ni Excel
│   ├── constantes.py          # colores, hojas, secciones, atajos
│   ├── texto.py               # normalización, búsqueda, nombres válidos
│   ├── monto_letras.py        # total en letras (réplica del VBA)
│   ├── validaciones.py        # cantidades (km/entero), Quetzales
│   ├── modelos.py             # Componente, FilaConcepto, Proyecto, HojaEstado, Configuración
│   ├── estado.py              # EstadoApp: proyectos, Ramal, filas, totales
│   ├── licencia.py            # gancho de licencia (D6 = A: desactivado)
│   └── tema.py                # QSS día / noche
├── servicios/
│   ├── excel_lector.py        # lectura del libro (solo lectura, validada)
│   ├── generador_hojas.py     # copia/renombra hojas y tablas, fórmula del Ramal
│   ├── excel_escritor.py      # escritura con xlwings (respaldo, recálculo, totales)
│   └── excel_componentes.py   # guardado de CODIGOS (altas, cambios y bajas)
├── pruebas/                   # unittest (el núcleo no requiere PySide6)
├── herramientas/
│   ├── inspeccionar_excel.py  # auditoría de cualquier .xlsx/.xlsm
│   └── verificar_libro.py     # comprueba que la lectura cuadre con G377
└── ui/                        # ventanas PySide6
    ├── ventana_base.py
    ├── ventana_principal.py
    ├── ventana_componentes.py
    ├── ventana_configuracion.py
    ├── ventana_edicion.py
    ├── ventana_visualizacion.py
    ├── ventana_ayuda.py
    ├── dialogo_seleccion.py
    └── dialogo_cantidad.py
```

## Cómo ejecutar la aplicación

```powershell
python -m pip install -r requirements.txt
python main.py
```

## Cómo correr las pruebas

```powershell
python -m unittest discover -s pruebas -v
```

Las pruebas del núcleo no necesitan PySide6 ni Excel. Las de
`pruebas/test_excel_lector.py` se ejecutan solo si el `.xlsm` está en la raíz
y comprueban que la suma de las 4 secciones cuadre con el total del libro.

### Pruebas de integración con Excel (opcionales)

Generan libros reales y los vuelven a leer. Necesitan Excel instalado:

```powershell
$env:ANEXOS_PROBAR_EXCEL = "1"
python -m unittest pruebas.test_excel_escritor -v      # exportación completa (~33 s)
python -m unittest pruebas.test_excel_componentes -v   # alta/cambio/baja en CODIGOS (~128 s)
Remove-Item Env:\ANEXOS_PROBAR_EXCEL
```

## Cómo verificar la lectura del libro

```powershell
python herramientas/verificar_libro.py "VARIACIÓN FINAL - COPONES - 14.10.2026.xlsm"
```

Imprime, por cada anexo, la suma de las 4 secciones contra el total `G377`.
Todas las diferencias deben ser `0.0000`.

## Cómo inspeccionar un libro

```powershell
python herramientas/inspeccionar_excel.py "VARIACIÓN FINAL - COPONES - 14.10.2026.xlsm"
```

## Guardado de componentes (`CODIGOS`)

La ventana **Componentes** (`F2`) permite:

- **Guardar Local:** escribe en el libro abierto (el anexo que está editando).
- **Guardar Global:** escribe en `plantillas/Plantilla_Anexos.xlsm` para que los
  proyectos nuevos usen la lista actualizada.

En ambos casos se crea un respaldo antes de tocar el libro. El de la plantilla se
guarda como `backups/Backup_DD.MM.YYYY_Plantilla.xlsm`; el local, junto al libro
(`Backup_DD.MM.YYYY.xlsm`). La escritura se hace en un hilo aparte, con xlwings,
para no bloquear la interfaz.

La sincronización es por **código**: los códigos existentes se actualizan en su
fila (columnas B–D y G–M), los nuevos se agregan al final con su fórmula
`=D<fila>*E<fila>`, y las bajas solo eliminan componentes manuales que ya no
están en la tabla.

## Ayuda y atajos

`F1` abre la ayuda en cualquier ventana. Incluye los atajos vigentes:

| Tecla | Acción |
|---|---|
| `F1` | Ayuda |
| `F2` | Gestión de Componentes |
| `Ctrl + S` | Guardar / exportar el anexo |
| `Ctrl + Espacio` | Enfocar la búsqueda |
| `Supr` | Eliminar la fila seleccionada (pide confirmación) |
| `Alt + B` | Tema día / noche |
| `Esc` | Cerrar la ventana actual |

## Decisiones (D1–D7)

| # | Decisión | Estado |
|---|---|---|
| D1 | Escritura con **xlwings** (conserva macros y dibujos) | ✅ Aplicada |
| D2 | Extender `CODIGOS` con columnas de línea/voltaje/activo al final | ✅ Aplicada |
| D3 | Plantilla = copia limpia de COPONES (`plantillas/Plantilla_Anexos.xlsm`) | ✅ Aplicada |
| D4 | Estructura **modular** (`main.py` + `nucleo/` + `ui/` + `servicios/`) | ✅ Aplicada |
| D5 | Nomenclatura de hojas: respetar el libro real (`ANE. - X`) | ✅ Aplicada (A) |
| D6 | Licencia: gancho **desactivado** (`nucleo/licencia.py`) | ✅ Aplicada (A) |
| D7 | Total en letras: como el libro (`MILLON`, sin tilde, sin `DE`) | ✅ Aplicada (A) |

## Notas confirmadas del libro real

1. **Abreviación máxima: 17 caracteres.** El prefijo más largo (`ANE. - VAR. - `, 14)
   más 18 daría 32 y Excel rechaza nombres de más de 31. El límite de 18 del diseño
   aplica solo si el prefijo fuese `ANE - VAR - ` (12). `validar_abreviacion` ya lo controla.
2. **El total en letras tiene dos variantes.** La macro del libro escribe `MILLON`
   (sin tilde) y nunca pone `DE`. El diseño pide `MILLÓN` y `UN MILLÓN DE QUETZALES`.
   `convertir_monto_quetzales` tiene las banderas `con_tildes` y `usar_de_en_millon`;
   por defecto usa la variante del libro (D7 = A).
3. **El precio real lo calcula Excel** desde `CODIGOS` (VLOOKUP): el estado en memoria
   solo aporta el código y la cantidad.
4. **Al copiar hojas**, Excel agrega sufijos a los nombres de tabla; el renombrado debe
   hacerse por índice usando el nombre base.
5. **El código 1083 y los códigos ≥ 2000 se suman al total** aunque el libro deje su
   columna de código en blanco. La app los conserva al leer y al guardar.
6. **La tabla General es fija**: solo `1150` (Diseño del Plano Final) y
   `1151` (Rótulo de la Obra). Si el diseño quiere meter cualquier componente en General,
   hay que cambiar fórmulas del libro.

## Reglas de seguridad

1. **Nunca** guardar el libro con `openpyxl` (pierde dibujos).
2. **No reordenar** las filas de `CODIGOS`. Los códigos existentes se actualizan en su
   fila; las bajas solo eliminan filas de componentes manuales.
3. Las fórmulas de `VARIACIÓN`, `OC`, `OTS` y `ATE` las genera la app al exportar; **no** editarlas a mano en el Excel final.
4. Antes de cualquier cambio estructural, crear `Backup_DD.MM.YYYY.xlsm`; antes de tocar
   el punto de entrada, crear `Backup_DD.MM.YYYY.py`. Los respaldos no se editan.
5. Mantener todos los strings, variables y comentarios en español.
