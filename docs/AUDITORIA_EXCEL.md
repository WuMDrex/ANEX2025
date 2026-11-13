# Auditoría del libro real — `VARIACIÓN FINAL - COPONES - 14.10.2026.xlsm`

> Fecha: 19.10.2026 · Herramienta: `openpyxl` (lectura) + `oletools` (VBA) · Libro: 3.89 MB

## 1. Resumen estructural

| Dato | Valor |
|---|---|
| Hojas totales | **56** |
| Entidades | **7**: 6 proyectos + 1 Ramal |
| Proyectos | `CHAILA`, `ARROLLO`, `VENTANA`, `SAN_PEDRO`, `COTIJA`, `PLAYITAS` |
| Ramal | `COPONES` |
| Hojas por entidad | **6** |
| Plantillas ocultas | **12** (6 `*- COMUNIDAD` + 6 `*- RAMAL`) |
| Componentes en `CODIGOS` | **121** (filas 2–122) |
| Tablas de Excel | **196** |
| Macros VBA | Sí (`vbaProject.bin`) |
| Dibujos/imágenes | 56 dibujos + 1 imagen (logos, firmas) |

### Hojas por entidad (nombres reales)

```
ANE. - <ABREV>            (oculta en los proyectos)
ANE. - VAR. - <ABREV>     (visible)
VARIACIÓN - <ABREV>       (visible)
OC - <ABREV>              (visible)
OTS - <ABREV>             (visible)
ATE - <ABREV>             (visible)
```

> **Diferencia con el diseño:** el libro usa **punto** (`ANE. - COPONES`), no `ANE - COPONES`.
> También usa tilde en `VARIACIÓN`. La app debe respetar el libro real o renombrar todas las
> hojas y reescribir fórmulas (riesgo alto).

### Hojas de configuración

- `INICIO`
  - Tabla `Configuration` (`B9:L15`): `No.`, `Proyecto`, `Abreviación`, `Abrevi_Ant`, `Ane.`, `Ane. Var`, `Variación`, `OC`, `OTS`, `ATE`, `Exportar`.
  - Tabla `Configuration_Ramal` (`B3:L4`): las mismas columnas para el Ramal.
  - `C7` = cantidad de comunidades (validación entera).
  - `A1` = `Datos del Proyecto (Línea 7.6)`.
- `CODIGOS`
  - Tabla `CODIGOS_SICOA` (`A1:F122`): `Codigo Componente`, `Descripción del componente`, `Unidad de medida`, `Precio unitario`, `Factor de ponderación`, `Precio ponderado`.
  - **No tiene** columnas de línea (MT/MT+BT/BT/G), voltaje (7.6/19.9), activo ni orden. Esas columnas del diseño **no existen** en el libro.

## 2. Estructura de las hojas ANE / ANE-VAR

Cada hoja tiene **4 tablas**:

| Sección | Tabla | Filas | Contenido |
|---|---|---|---|
| MEDIA TENSIÓN | `A_MT_<ABREV>` / `AV_MT_<ABREV>` | 121 | Todos los componentes |
| MEDIA Y BAJA TENSIÓN | `A_MT_BT_...` | 121 | Todos los componentes |
| BAJA TENSIÓN | `A_BT_...` | 121 | Todos los componentes |
| GENERAL | `A_G_...` | **2** | Solo códigos `1150` y `1151` |

Columnas: `CODIGO COMPONENTE`, `Tipo`, `Descripción del componente`, `Unidad de medida`, `Precio unitario`, `Cantidad a ejecutar`, `Precio total`.

### Fórmulas clave

- `A7` = `=IF(OR(CODIGOS!A2=1083,CODIGOS!A2>=2000),"",CODIGOS!A2)`
- `C` = `=_xlfn.XLOOKUP([@[CODIGO COMPONENTE]],CODIGOS_SICOA[Codigo Componente],CODIGOS_SICOA[Descripción del componente],"")`
- `D` = `=_xlfn.XLOOKUP([@[CODIGO COMPONENTE]],CODIGOS_SICOA[Codigo Componente],CODIGOS_SICOA[Unidad de medida],"")`
- `E` = `=_xlfn.XLOOKUP([@[CODIGO COMPONENTE]],CODIGOS_SICOA[Codigo Componente],CODIGOS_SICOA[Precio unitario],0)`
- `G` = `=ROUND([@[Precio unitario]]*[@[Cantidad a ejecutar]],2)` (o `ROUNDUP`/`ROUNDDOWN` según el modo de esa fila; con el modo **«0.00 SR»** se escribe la multiplicación exacta sin función: `=([@[Precio unitario]]*[@[Cantidad a ejecutar]])`)
- `G377` = `=SUM(A_BT_...[Precio total], A_MT_BT_...[Precio total], A_MT_...[Precio total], A_G_...[Precio total])`
- `A379` = `=ConvertirMontoQuetzales(G377)` (macro VBA)
- El Ramal calcula su cantidad con `=SUMIF(A_MT_PROY1[CODIGO COMPONENTE],[@[CODIGO COMPONENTE]],A_MT_PROY1[Cantidad a ejecutar]) + ...` por cada proyecto activo.

> **Cada fila de la tabla está amarrada a una fila fija de `CODIGOS`.**
> Si se reordena, inserta o elimina un componente en `CODIGOS`, **todas** las hojas se desalinean.
> El orden de `CODIGOS` es intocable; los componentes nuevos solo se pueden **agregar al final**.

### Regla de exclusión

La fórmula `IF(OR(código=1083, código>=2000),"",código)` deja en blanco:

- `1083` — Estructura tipo XI (13.8 kV).
- `2001`, `2002`, `2003` — Postes de compuesto P.R.F.V.

Son códigos **de ATE / trabajos adicionales**, no del ANE normal.

## 3. VARIACIÓN, OC, OTS, ATE

- `VARIACIÓN` tiene 15 columnas: contratado, a modificar, variaciones, `Tipo de Documento`, `CANTIDAD VARIADA`, `OC`, `OTS` y `ATE`; `CANT TOTAL` fue eliminada.
- `J` (Tipo de Documento) usa la fórmula exacta:
  `=IF(AND(ISNUMBER([@[No.]]),[@[No.]]>2000),"ATE",IF([@[CANTIDAD VARIADA]]>0,"OTS",IF([@[CANTIDAD VARIADA]]<0,"OC","---")))`
  - código numérico **MAYOR a 2000** → `ATE`, sin importar el signo de la cantidad variada
  - código numérico **≤ 2000** (incluido el 2000 exacto) → sigue la regla de la cantidad variada
  - cantidad variada > 0 → `OTS`
  - cantidad variada < 0 → `OC`
  - en otro caso → `---`
- El umbral vive en una sola constante (`UMBRAL_CODIGO_ATE = 2000`) que comparten la fórmula de Excel y `Estado._tipo_documento`; el 2000 exacto no es ATE.
- El formato condicional de la columna del código (`CODIGO COMPONENTE` en `ANE`/`ANE-VAR`, `No.` en `VARIACIÓN`/`OC`/`OTS`/`ATE`) usa `=(A+0=A)*(A>2000)` con fuente blanca: solo se ocultan los códigos ATE (> umbral). Un código numérico ≤ 2000 se ve normalmente aunque toda la columna sea numérica. La expresión se arma solo con operadores aritméticos y de comparación (`+`, `*`, `>`) y no con nombres de función, porque las funciones se traducen según el idioma de Excel y los operadores no: así la regla se evalúa igual en Excel en español o en inglés.
- La app normaliza el encabezado de tabla y el rótulo manual de la columna `K` a `CANTIDAD VARIADA` (un solo espacio) antes de escribir las fórmulas.
- `OC - <ABREV>`, `OTS - <ABREV>` y `ATE - <ABREV>` son **copias con fórmulas** de la variación; cada hoja recibe solo las filas de su tipo y se recalcula sola.
- El monto de `F`/`I` y de la columna de importe de `OC`/`OTS`/`ATE` usa el **mismo modo por fila** que `ANE`/`ANE-VAR`: `ROUND`/`ROUNDDOWN` (y `ROUNDUP` heredado) escriben la función con `,2`, y **«0.00 SR»** escribe el producto sin redondear, envuelto en paréntesis: `=([@[CANTIDAD]]*[@[PRECIO UNITARIO (Q)]])` en `F`, `=([@[CANTIDAD - V]]*[@[PRECIO UNITARIO ( Q )]])` en `I` y `([@[CANTIDAD VARIADA]]*[@[PRECIO UNITARIO ( Q)]])` dentro del `IF([@[Tipo de Documento]]="ATE",…,"")` en las hojas de documento.
- Al reimportar un `.xlsx` la app deduce el modo leyendo esa fórmula: `ROUNDUP`/`ROUNDDOWN`/`ROUND` devuelven su propio modo y una fórmula con `*` que no contiene `ROUND` se interpreta como **«0.00 SR»**. Por eso no debe editarse a mano el `Precio total` para dejarlo en un producto suelto si no se quiere que pase a «0.00 SR».
- `VARIACIÓN` totaliza en `F376` (contratado), `I376` (variaciones), `M376/N376/O376` (OC/OTS/ATE) y `B381:C383` (costo total, costo por variaciones, diferencia).

> **Consecuencia para el botón ⟳ Actualizar:** la app genera las fórmulas de OC/OTS/ATE al
> exportar; al reabrir un `.xlsx` se recalculan solas (`Application.Calculate`). No se editan
> a mano en el archivo final.

## 4. Macros VBA encontradas

| Macro | Qué hace |
|---|---|
| `ConvertirMontoQuetzales` | Total en letras. Réplica exacta en `nucleo/monto_letras.py`. |
| `GenerarPlantillaProyectos` | Crea/elimina hojas de comunidades y Ramal a partir de las plantillas ocultas, renombra tablas y acomoda la visibilidad. |
| `RenombrarYMostrarProyectos` | Renombra hojas y tablas al cambiar la abreviación y actualiza el nombre del proyecto en `A1`. |
| `SincronizarTablasConExclusión` | Sincroniza filas entre tablas y oculta/filtra las que no tienen cantidad. |
| `ActualizarFormulasRamal` | Reescribe las fórmulas `SUM` del Ramal según los proyectos activos. |
| `Filtrar_ANE`, `Filtrar_ANE_VAR`, `Filtrar_VARIACION`, `Filtrar_OC`, `Filtrar_OTS`, `Filtrar_ATE` | Filtros de las tablas. |

> La app puede **llamar estas macros** por `xlwings` en lugar de reimplementarlas. Es la vía
> de menor riesgo para no romper el libro.

## 5. Hallazgo crítico: `openpyxl` NO puede guardar este libro

Al leerlo, `openpyxl` advierte:

> `DrawingML support is incomplete... Shapes and drawings will be lost.`

El libro tiene **56 dibujos** (logos, firmas, membretes) y una imagen. Si se guarda con
`openpyxl`, **se pierden los dibujos** y el archivo deja de verse igual. Además, las macros
solo se conservan con `keep_vba=True`, pero eso no recupera los dibujos.

### Opciones

| Opción | Ventaja | Riesgo |
|---|---|---|
| **A. `xlwings` para escribir** (Excel instalado) | Conserva macros, dibujos, fórmulas y formato. Permite llamar las macros existentes. | Requiere Excel; más lento; hay que liberar el proceso. |
| **B. `openpyxl` para escribir** | Rápido, sin Excel. | **Pierde dibujos/firmas**. No recomendado para este libro. |
| **C. Híbrido (recomendado)** | `openpyxl` solo para **leer** (`data_only=True`); `xlwings` para **escribir/guardar**. | Dos dependencias, pero es el único camino seguro. |

## 6. Otros hallazgos que afectan el diseño

1. **No existe la BD con casillas MT/MT+BT/BT/G ni voltajes.** El libro repite los 121 componentes
   en las 3 tablas (MT, MT+BT, BT) y solo 2 en General. Las casillas del diseño son un filtro
   **de la app**, no una columna del Excel.
2. **El filtro de Voltaje no tiene datos de origen.** El libro es de "Línea 7.6". Hay que decidir
   si se marca a mano, si se deduce de la descripción (transformadores) o si se agrega una hoja
   nueva de configuración.
3. **La hoja General está fija** a `1150` (Diseño del Plano Final) y `1151` (Rótulo de la Obra).
   El diseño dice que cualquier componente puede ir a General; el libro no lo permite sin cambiar fórmulas.
4. **La columna `Exportar` de `INICIO`** decide qué hojas se exportan. La app debe respetarla.
5. **`Abrevi_Ant`** guarda la abreviación anterior y es la base para renombrar sin perder referencias.
6. **El nombre del proyecto** se escribe en `A1` de la hoja `ANE. - <ABREV>`; las demás hojas lo
   referencian (`A1 = 'ANE. - X'!$A$1`) o lo muestran en sus propios títulos.
7. **`C7` de `INICIO`** es la cantidad de comunidades; el Ramal se activa con ≥ 2.
8. **Los códigos `>= 2000` y `1083`** se excluyen del ANE; si la app los muestra en la búsqueda
   normal, el usuario podría agregarlos y el Excel los dejaría en blanco. Hay que marcarlos como ATE.
   Nota de la app: la aplicación usa `> 2000` estricto, por lo que `2000` exacto y `1083` no se
   clasifican como ATE.

## 7. Validación con el libro real (19.10.2026)

El lector `servicios/excel_lector.py` se probó contra el libro y **reproduce exactamente**
el total `G377` de cada hoja `ANE. - VAR. - <ABREV>`:

| Abreviatura | Suma de las 4 secciones | Total `G377` | Diferencia |
|---|---:|---:|---:|
| CHAILA | 1,910,469.6066 | 1,910,469.6066 | 0.0000 |
| ARROLLO | 1,143,145.4481 | 1,143,145.4481 | 0.0000 |
| VENTANA | 1,213,675.2099 | 1,213,675.2099 | 0.0000 |
| SAN_PEDRO | 2,530,230.3310 | 2,530,230.3310 | 0.0000 |
| COTIJA | 783,315.3872 | 783,315.3872 | 0.0000 |
| PLAYITAS | 430,692.7022 | 430,692.7022 | 0.0000 |
| COPONES (Ramal) | 8,011,528.6850 | 8,011,528.6850 | 0.0000 |

Para reproducirlo: `python herramientas/verificar_libro.py "<libro.xlsm>"`.

### Hallazgos finos confirmados

1. **El código 1083 SÍ se suma.** La fórmula `IF(código=1083 O código>=2000, "", código)`
   deja la columna del código en blanco, pero la fila **sigue contando** en el total.
   En COPONES hay 5 unidades de la 1083 en MT y 11 en MT+BT (Q117,600 en total).
   El lector recupera el código por la **posición fija** de la fila dentro de la tabla.
   La app no debe perder estas filas al leer ni al guardar.
2. **La tabla General guarda la descripción en la columna `Tipo`.**
   La columna `Descripción del componente` viene vacía; el lector lo compensa.
   Sus únicos códigos son `1150` y `1151`.
3. **Cuidado con el sufijo de las tablas.** La sección `BT` (`AV_BT_<ABREV>`) y la
   `MT+BT` (`AV_MT_BT_<ABREV>`) se confunden si se busca por sufijo. Hay que usar el
   nombre exacto de la tabla.
4. **Las banderas de `INICIO` deciden qué hojas existen y se muestran.**
   - `Ane.` / `Ane. Var` / `Variación` / `OC` / `OTS` / `ATE` (columnas F–K).
   - La hoja `ANE. - X` es interna: siempre oculta.
   - Proyectos: solo `ANE. - VAR. - X` visible.
   - Ramal: `ANE. - VAR.` + `VARIACIÓN` + `OC` + `OTS` + `ATE` visibles.
   - `Exportar` (columna L) es aparte.
5. **`xlwings` funciona**: abre el libro oculto, **recalcula** (`app.calculate()`),
   lee los totales y lo cierra **sin guardar**. Es la vía segura para escritura.

## 8. Lo que falta para cerrar la especificación

| # | Falta | Por qué |
|---|---|---|
| 1 | **Plantilla limpia** (sin datos de COPONES) | El libro actual es un proyecto real, no una plantilla. |
| 2 | **Decidir dónde viven las casillas de línea/voltaje** | No existen en `CODIGOS`. |
| 3 | **Decidir el motor de escritura** | `openpyxl` pierde dibujos. |
| 4 | **Confirmar nomenclatura** (`ANE. -` vs `ANE -`) | Cambiarla obliga a reescribir fórmulas. |
| 5 | **Confirmar tipografía del total en letras** | El VBA no usa tildes ni "DE"; el diseño sí. |
| 6 | **Confirmar licencia** | `RAIZ.py` usa `license_manager.check_license()`. |
| 7 | **Confirmar estructura de código** | Modular vs monolito `main.py`. |
