# Diseño del nuevo software – Especificación gráfica y lógica

> Basado en tus 5 pantallas dibujadas, tu descripción y el estilo de `RAIZ.py` (PySide6).
> Los puntos marcados con **⚠ DECISIÓN** son cosas que detecté que faltan o se contradicen, y necesito que las confirmes.

---

## 0. Resumen de pantallas

| # | Ventana | Cómo se abre | Cómo se sale |
|---|---|---|---|
| 1 | **Principal** | Al iniciar | `Esc` o `X` → pregunta si desea guardar |
| 2 | **Configuración** | ⚙ (tuerca) de la Principal | `Esc` / `X` → regresa a Principal |
| 3 | **Edición** | Botón de la **Abreviación** dentro de Configuración | `Esc` / `X` → regresa a Configuración. Botón 👁 → va a Visualización |
| 4 | **Visualización** | 👁 (ojo) de la Principal, o desde Edición | `Esc` / `X` → regresa a donde venía. Botón de Abreviación → va a Edición |
| 5 | **Componentes** | Botón 4 (capas) de la Principal | `Esc` / `X` → pregunta si desea guardar |
| – | **Proyecto** | Botón "Proyecto" de la Principal | Menú: Nuevo / Abrir |
| – | **Actualizar** | Botón 5 (flechas circulares) | Animación mientras trabaja |

### Navegación

```
                 ┌────────────── 👁 ──────────────┐
                 ▼                                │
 PRINCIPAL ──⚙──► CONFIGURACIÓN ──(Abreviación)──► EDICIÓN
     │                                              ▲  │
     │                                              │  👁
     ├──👁──► VISUALIZACIÓN ──(Abreviación)─────────┘  │
     │             ▲───────────────────────────────────┘
     └──▣──► COMPONENTES
```

---

## 1. Diseño visual (tomado de `RAIZ.py`)

**Fuente:** Segoe UI, 8 pt. Títulos 14 pt negrita. Subtítulos 10 pt negrita.
**Esquinas:** botones/inputs radio 4 px, grupos radio 8 px. **Bordes de input:** 2 px.

### Colores por tema

| Elemento | 🌙 Noche | ☀️ Día |
|---|---|---|
| Fondo ventana | `#2c3e50` | `#f5f6fa` |
| Texto | `#ecf0f1` | `#2c3e50` |
| Fondo inputs / listas / tablas | `#34495e` | `#ffffff` |
| Borde | `#7f8c8d` | `#bdc3c7` |
| Fondo grupo | `#34495e` | `#ffffff` |
| Borde grupo | `#7f8c8d` | `#e0e0e0` |
| Botón tipo "radio" (Abreviación, ANE…) | `#34495e` | `#d6eaf8` |
| Hover radio | `#7f8c8d` | `#aed6f1` |
| Selección fila | `#2980b9` | `#3498db` |
| Foco input | `#3498db` | `#3498db` |

### Colores de botones (iguales en ambos temas)

| Uso | Normal | Hover |
|---|---|---|
| Primario (Agregar, Guardar, Regresar) | `#3498db` | `#2980b9` |
| Éxito (Nuevo, Agregar material, Línea separadora de sección) | `#27ae60` | `#229954` |
| Confirmar / Excel | `#16a085` | `#138d75` |
| Peligro (Papelera, Salir, Eliminar) | `#e74c3c` | `#c0392b` |
| Advertencia | `#f39c12` | `#e67e22` |
| Deshabilitado | `#95a5a6` | – |
| Seleccionado / activo (ANE, ANE-VAR…) | `#3498db` con borde `#21618c`, texto blanco negrita | – |

### Botón día/noche
Igual que RAIZ: en modo noche muestra ☀️ (clic para pasar a día), en modo día muestra 🌙. Atajo `Alt+B`. El cambio de tema se aplica a **todas** las ventanas abiertas (usar una sola función `aplicar_tema()` compartida).

### Barra de herramientas de la Principal (de izquierda a derecha)

| Posición | Icono propuesto | Función |
|---|---|---|
| 1 | ⚙ | Configuración |
| 2 | `Proyecto` (texto) | Nuevo / Abrir |
| 3 | 👁 | Visualización |
| 4 | ◈ (capas) | Componentes |
| 5 | ⟳ | Actualizar (gira mientras trabaja) |
| 6 | 💾 | Guardar (copia de la plantilla) |
| Derecha | `Voltaje: [ ▼ ]` | Filtro de voltaje |

---

## 2. Tamaños de ventana propuestos

Todas redimensionables. Si el contenido no cabe, se envuelve en `QScrollArea` (barra vertical; horizontal solo si es necesario). Para `Esc`: `QShortcut(QKeySequence("Esc"), self)`.

| Ventana | Tamaño inicial | Mínimo | Notas |
|---|---|---|---|
| **Principal** | 820 × 780 | **820 × 780** (el mínimo = el tamaño inicial) | Se puede extender hasta pantalla completa; ver 2.1 |
| **Configuración** | 480 × 420 | 420 × 320 | Lista de proyectos con scroll |
| **Edición** | 760 × 460 | 700 × 340 | Cada proyecto ≈ 110 px de alto; scroll vertical |
| **Visualización** | 700 × 340 | 640 × 260 | Cada proyecto ≈ 36 px de alto |
| **Componentes** | 980 × 640 | 860 × 500 | La tabla inferior ocupa el espacio sobrante (ahora con 2 columnas más de voltaje) |
| Diálogo confirmación | 380 × 160 | – | Fijo |
| Diálogo guardar Local / Global | 440 × 200 | – | Fijo |
| Menú Proyecto (Nuevo / Abrir) | 360 × 200 | – | Fijo |
| Selector de color (pincel) | Estándar `QColorDialog` | – | – |
| Popup de ayuda de búsqueda | Ancho = campo, alto máx. 8 filas (~200 px) | – | Con scroll si hay más |

### Anchos de columnas de la tabla de la Principal (a 820 px)

| Columna | Ancho |
|---|---|
| Código componente | 95 |
| Descripción del componente | Flexible (mín. 220) |
| Unidad | 80 |
| Precio Unitario | 105 |
| Cantidad a ejecutar | 95 |
| Precio Total | 115 |

### 2.1 La Principal al agrandarse (hasta pantalla completa)

| Elemento | Comportamiento al agrandar |
|---|---|
| Ventana | Mínimo 820 × 780; botón maximizar habilitado |
| Columna Descripción | Absorbe todo el ancho extra; las demás columnas mantienen su ancho |
| Campo Componente | Se estira con la ventana |
| Popup de búsqueda | Su ancho sigue al campo y su alto crece: muestra **más filas** (≈ las que quepan, en vez de 8 fijas) |
| Tablas | Crecen en alto; se muestran más filas antes de aparecer el scroll |
| Pie (Total / Total en letras) | Siempre visible abajo |

> **Mi interpretación:** tomé el tamaño propuesto de la Principal como su **mínimo**, y que al extenderla la pantalla "se relacione con las búsquedas", es decir, que el popup de búsqueda muestre más resultados. Si te referías a otra cosa (por ejemplo, que el mínimo sea el tamaño de la ventana de Visualización), dime y lo cambio.

---

## 3. Ventana PRINCIPAL

### 3.1 Wireframe

| | | | | | |
|---|---|---|---|---|---|
| ⚙ `Proyecto` 👁 ◈ ⟳ 💾 | | | | **Voltaje:** `[Línea 7.6 kV ▼]` | |
| **Materiales:** | | | | | **▲** (colapsar panel) |
| `Código` | `Componente` (ocupa el resto de la fila) | | | | |
| `Cantidad` | **Línea:** `[ ▼ ]` | | `Agregar` | | |

Debajo, las 4 tablas (cada una **solo se muestra si tiene al menos un elemento**):

| Sección | Código | Descripción del componente | Unidad | Precio Unitario | Cantidad a ejecutar | Precio Total |
|---|---|---|---|---|---|---|
| **MEDIA TENSION** | … | … | … | Q 0.00 | … | Q 0.00 |
| **MEDIA Y BAJA TENSION** | … | … | … | Q 0.00 | … | Q 0.00 |
| **BAJA TENSION** | … | … | … | Q 0.00 | … | Q 0.00 |
| **GENERAL** | … | … | … | Q 0.00 | … | Q 0.00 |

Pie:

| Elemento | Contenido |
|---|---|
| **TOTAL:** | Q 0.00 (**negrita**, suma de los Precio Total de las 4 tablas) |
| Total en letras: | `Total en letras: UN MILLÓN QUINIENTOS TREINTA Y CINCO MIL NOVECIENTOS QUINCE QUETZALES CON CUARENTA Y OCHO CENTAVOS DE QUETZAL` |
| Pestañas de hojas (abajo izquierda) | `ANE - COMUNIDAD_1` · `ANE - VAR - COMUNIDAD_1` · … (con el color elegido con el pincel) |

### 3.2 Campos de la barra superior

| Campo | Tipo | Opciones / regla |
|---|---|---|
| Voltaje | Lista desplegable | `Línea 7.6 kV`, `Línea 19.9 kV` — aplica un filtro |
| Línea | Lista desplegable | `Media tensión`, `Media + baja tensión`, `Baja tensión`, `General` → define **a qué tabla** va el componente |
| Código | Texto con búsqueda | Busca por código en la BD |
| Componente | Texto con búsqueda | Busca por descripción en la BD |
| Cantidad | Texto con validación | Ver 3.3 |
| Agregar | Botón | Se habilita solo si Código, Cantidad y Línea son válidos |

### 3.3 Validación de Cantidad

| Unidad del componente | Entrada | Resultado |
|---|---|---|
| `Kilómetro` | `2` | Se guarda `2.000` |
| `Kilómetro` | `2.5` | Se guarda `2.500` |
| `Kilómetro` | `2.55` | Se guarda `2.550` (se completa con 0) |
| `Kilómetro` | `2.555` | Se guarda `2.555` |
| `Kilómetro` | `2.5555` | **Bloqueado** (no deja escribir el 4.º decimal) |
| Cualquier otra | `7` | Se guarda `7` |
| Cualquier otra | `7.5` | **Bloqueado** (solo enteros) |
| Cualquiera | `0` o vacío | Agregar deshabilitado |

> Implementación sugerida: `QRegularExpressionValidator` que cambia según la unidad detectada: `^\d{1,9}(\.\d{0,3})?$` para Kilómetro y `^\d{1,9}$` para el resto. Comparar la unidad **sin acentos y sin importar mayúsculas** (`kilometro`, `Kilómetro`, `KILOMETRO`).

### 3.4 Autocompletado tipo VSCode

| Aspecto | Comportamiento |
|---|---|
| Aparece | Al escribir ≥ 1 carácter en **Código** o **Componente** |
| Posición | Justo debajo del campo, sin quitar el foco |
| Columnas | `Código` · `Descripción` |
| Coincidencia | "Contiene", sin acentos, sin importar mayúsculas |
| Máximo visible | 8 filas en tamaño mínimo; crece con la ventana (ver 2.1). Scroll si hay más |
| Navegación | `↑` `↓` mover · `Enter`/`Tab` seleccionar · `Esc` cierra solo el popup |
| Al seleccionar | Se autocompletan **ambos** campos (Código y Componente) y la Unidad se lee de la BD |
| Filtro aplicado | Solo componentes con la casilla activa (1) para la **Línea** elegida **y** compatibles con el **Voltaje** elegido (ver 7.5: voltaje marcado = ese; ninguno marcado = ambos) |

#### Búsqueda total con `Ctrl + Espacio`

| Aspecto | Comportamiento |
|---|---|
| Atajo | `Ctrl + Espacio` estando en **Código** o **Componente** |
| Qué muestra | **Todos** los componentes de la BD (Código · Descripción), sin filtrar por Línea ni Voltaje, aunque el campo esté vacío |
| Al seguir escribiendo | Filtra la lista (igual que el autocompletado normal) |
| Navegación y selección | `↑` `↓`, `Enter`/`Tab`, `Esc` cierra el popup |
| Componentes que no corresponden a la Línea/Voltaje actual | **Propuesta:** se muestran en gris con una etiqueta; al elegirlos, preguntar *"Este componente no está habilitado para <Línea>. ¿Agregarlo de todas formas?"* |
| También disponible en | Ventana **Componentes** (campos Código y Componente) |

> Nota: en algunos Windows con varios idiomas de teclado, `Ctrl + Espacio` puede estar asignado a otra función del sistema. Si pasa, se puede dejar un segundo atajo (por ejemplo `F4`).

### 3.5 Reglas de la tabla

| Regla | Detalle |
|---|---|
| Tabla vacía | Se oculta la tabla **y** su encabezado (incluido el título de la sección) |
| Primer elemento | La tabla y el encabezado aparecen |
| Último elemento eliminado | Se vuelven a ocultar |
| Precio Total | `Precio Unitario × Cantidad` (fórmula en el Excel) |
| Formato moneda | Quetzales, igual al del Excel: `Q#,##0.00` |
| Total general | Negrita |
| Componente que **ya está** en la misma tabla (Línea) | Se **modifica** la fila existente (la cantidad nueva reemplaza a la anterior) |
| Componente que ya está en **otra** tabla (otra Línea) | Se **agrega** como fila nueva en la tabla de la Línea elegida |
| Eliminar fila | Seleccionar + `Supr` |
| Editar cantidad | Doble clic en la celda (como Excel) |
| Total en letras | Ver sección 8 |

### 3.6 Botón ⟳ Actualizar

El programa **controla hojas de Excel**, y el Excel puede haberse modificado fuera del programa. Este botón lo revisa.

| Paso | Qué hace |
|---|---|
| 0 | El icono ⟳ **gira** mientras trabaja (proceso en segundo plano para no congelar la ventana); la ventana principal queda bloqueada |
| 1 | **Verifica los totales:** compara el total de cada hoja con el último valor conocido por el programa. Si cambió, lo reporta |
| 2 | **Busca componentes externos:** revisa las tablas de `ANE` y `ANE - VAR` y detecta códigos agregados directamente en el Excel que **no existen en la base de datos** |
| 3 | **Actualiza** las hojas dependientes que existan: `VARIACION`, `OC`, `OTS` y `ATE` |
| 4 | Muestra un resumen al terminar y detiene la animación |

Resumen / reporte (ventana pequeña, 560 × 360, con scroll):

| Sección del reporte | Contenido |
|---|---|
| Totales | Hoja · total anterior · total actual · diferencia |
| Componentes no registrados | Hoja · fila · código · descripción · botón por fila: **Agregar a la BD** / **Ignorar** |
| Hojas actualizadas | Lista de hojas que se refrescaron |

| Detalle a confirmar | Propuesta |
|---|---|
| Cómo conocer el "total anterior" | Guardar el último total de cada hoja al guardar con el programa (archivo auxiliar o hoja oculta del Excel) |
| Componente externo no está en la BD | No se borra del Excel; se avisa, y si el usuario elige **Agregar a la BD** se pide Unidad, Precio y casillas (Línea y Voltaje) |
| Si no hay cambios | Igual actualiza las hojas dependientes, y avisa *"Sin cambios detectados"* |
| Hojas que no existen en el proyecto | Se omiten |

---

## 4. Ventana CONFIGURACIÓN

### 4.1 Wireframe

| | | | |
|---|---|---|---|
| **Configuración** | | | |
| **Proyectos** | | | ▲ scroll ▼ |
| `COMUNIDAD_1` | 🖌 | 🗑 | |
| `COMUNIDAD_2` | 🖌 | 🗑 | |
| `COMUNIDAD_3` (nuevo) | ➕ | | |
| **Ramal** *(solo si hay ≥ 2 proyectos)* | | | |
| `RAMAL` | 🖌 | | |

### 4.2 Botones

| Botón | Acción |
|---|---|
| `Abreviación` (nombre) | Abre la ventana **Edición** de ese proyecto |
| 🖌 Pincel | Selector de color: cambia el color de la pestaña de la hoja. Se refleja en: el botón de Abreviación, las pestañas de la Principal y la pestaña del Excel |
| 🗑 Papelera | Confirmación: **"Se eliminará todo lo relacionado al proyecto"**. Si es el **único** proyecto → solo lo reinicia (elimina hojas extras, deja solo ANE vacío y nombre `COMUNIDAD_1`) |
| ➕ Más | Crea un nuevo proyecto con nombre `COMUNIDAD_X` (siguiente número libre). Aparece una fila nueva con 🖌 y 🗑 y otra fila con ➕ |

### 4.3 Reglas del nombre (Abreviación)

| Regla | Ejemplo |
|---|---|
| Nombre por defecto | `COMUNIDAD_1`, `COMUNIDAD_2`… (primer número no usado) |
| Sin espacios: se cambian por `_` | `Mi Aldea` → `Mi_Aldea` |
| No repetidos (sin importar mayúsculas) | Si ya existe → mensaje y no se acepta hasta que sea distinto |
| Caracteres prohibidos de Excel | `\ / ? * [ ] :` → se rechazan |
| Longitud máxima | **18 caracteres** (el nombre de hoja de Excel admite 31 y el prefijo más largo, `ANE - VAR - `, usa 12) |
| No vacío | Si se deja vacío, vuelve al nombre anterior |

### 4.4 Cuándo aparece la sección Ramal

| Proyectos existentes | Sección Ramal |
|---|---|
| 1 | Oculta (inactiva) |
| ≥ 2 | Visible |
| Se pasa de 2 a 1 (por eliminar) | Se elimina el Ramal y sus hojas → avisar al usuario antes |

---

## 5. Ventana EDICIÓN

### 5.1 Wireframe

| | | | | | | | |
|---|---|---|---|---|---|---|---|
| **Edición** 👁 | | | | | | | |
| **Proyectos** | | | | | | | ▲ ▼ |
| `Abreviación` | `ANE` | `ANE - VAR` | `VARIACION` | `OC` | `OTS` | `ATE` | |
| `Nombre del Proyecto` (campo de texto amplio) | | | | | | | |
| `Abreviación` | `ANE` | `ANE - VAR` | `VARIACION` | `OC` | `OTS` | `ATE` | |
| `Nombre del Proyecto` | | | | | | | |
| **Ramal** | | | | | | | |
| `Abreviación` | `ANE` | `ANE - VAR` | `VARIACION` | `OC` | `OTS` | `ATE` | |
| `Nombre del Ramal` | | | | | | | |

### 5.2 Campos

| Campo | Uso |
|---|---|
| Abreviación | Campo editable con las reglas de 4.3 |
| Nombre del Proyecto | Texto libre; se usa para la exportación al Excel (celda de título) |
| ANE, ANE - VAR, VARIACION, OC, OTS, ATE | Botones tipo *toggle*: marcado = la hoja existe |
| 👁 | Va a Visualización |

> Cualquier cambio de Abreviación o Nombre debe **actualizar las referencias en el Excel** (nombres de hojas, fórmulas que apuntan a ellas, celdas de título). **⚠ DECISIÓN:** al renombrar una hoja con openpyxl, las fórmulas de otras hojas **no** se actualizan solas; hay que reescribirlas (o usar xlwings, que sí lo hace). Hay que revisarlo con tu Excel.

### 5.3 Dependencias entre hojas (Proyecto)

| Hoja | ¿Se puede desmarcar? | Al desmarcar |
|---|---|---|
| **ANE** | ❌ Nunca | – |
| **ANE - VAR** | ✅ | Si hay `VARIACION`, `OC`, `OTS` o `ATE` activas: *"Si elimina esta hoja eliminará también las siguientes: …"* → Aceptar elimina todas |
| **VARIACION** | ✅ | *"¿Desea eliminar la hoja?"* |
| **OC** | ✅ | *"¿Desea eliminar la hoja?"* |
| **OTS** | ✅ | *"¿Desea eliminar la hoja?"* |
| **ATE** | ✅ | *"¿Desea eliminar la hoja?"* |

**Al marcar (crear) hojas dependientes (decidido):** si el usuario marca `VARIACION`, `OC`, `OTS` o `ATE` y `ANE - VAR` no está activa, el programa la **activa automáticamente** y lo informa: *"OC requiere ANE - VAR, se activó automáticamente"*.

| Pendiente menor | Pregunta |
|---|---|
| ¿`ATE` depende solo de `ANE - VAR`, o también de `OC`/`OTS`? | Asumo que solo de `ANE - VAR` (igual que las demás) |

### 5.4 Ramal

| Aspecto | Regla |
|---|---|
| Hoja ANE | **Sí** se puede desmarcar (todas se pueden) |
| Origen de los datos | Suma de renglones de cada proyecto |
| Visibilidad | Solo si hay ≥ 2 proyectos |
| Nombre de hoja | **⚠ DECISIÓN:** propongo `RAMAL` por defecto, editable, con las mismas reglas de nombre |
| Edición de componentes | **Solo lectura** (es suma de los proyectos, no se agrega manualmente) |

### 5.5 Detección del Ramal y de hojas existentes

Como el programa trabaja sobre Excel que pueden venir ya armados, al **abrir** un archivo (y al pasar a ≥ 2 proyectos) se revisa lo que ya existe.

| Situación | Qué hace el programa |
|---|---|
| Hay ≥ 2 proyectos y **ya existe** la hoja Ramal con la misma configuración (mismas hojas activas y mismos proyectos que suma) | La **reutiliza**, sin duplicarla |
| Hay ≥ 2 proyectos y existe, pero con **otra configuración** (le faltan o le sobran proyectos/hojas) | Avisa y propone ajustarla a la configuración actual |
| Hay ≥ 2 proyectos y **no existe** | La crea con las hojas necesarias |
| Se baja a 1 proyecto | Se avisa y se eliminan las hojas del Ramal |
| Nombres de hoja ya existentes en el Excel (`ANE - X`, `ANE - VAR - X`…) | Se reconocen como proyectos; la parte final del nombre es la Abreviación |

---

## 6. Ventana VISUALIZACIÓN

### 6.1 Wireframe

| | | | | | | | |
|---|---|---|---|---|---|---|---|
| **Visualización** 🌙/☀️ | | | | | | | |
| **Proyectos** | | | | | | | ▲ ▼ |
| `Abreviación` | `ANE` | `ANE - VAR` | `VARIACION` | `OC` | `OTS` | `ATE` | |
| `Abreviación` | `ANE` | `ANE - VAR` | `VARIACION` | `OC` | *(oculto)* | `ATE` | |
| **Ramal** | | | | | | | |
| `Abreviación` | `ANE` | `ANE - VAR` | | | | | |

### 6.2 Reglas

| Regla | Detalle |
|---|---|
| Botones mostrados | **Solo** los de hojas que existen en ese proyecto |
| Botón de hoja | Alterna si esa hoja se **muestra como pestaña** en la Principal (marcado = visible) |
| Botón `Abreviación` | Va a Edición |
| Botón día/noche | Cambia el tema de **todos** los formularios |
| `Esc` / `X` | Regresa a donde venía (Principal, o Edición si vino de ahí) |
| Alineación | Mantener las columnas en posiciones fijas (en tu dibujo, cuando falta OTS queda hueco, y eso se ve bien: conservarlo) |

---

## 7. Ventana COMPONENTES

### 7.1 Wireframe

| | | | | | | | |
|---|---|---|---|---|---|---|---|
| **Componentes** 💾 | | | | | | | |
| `Código` | `Componente` (ancho completo) | | | | | | |
| `Unidad` | `Precio Unitario` | `Agregar` | | | | | |

Tabla inferior:

| Código componente | Descripción del componente | Unidad | Precio Unitario | MT | MT + BT | BT | G | 7.6 kV | 19.9 kV |
|---|---|---|---|---|---|---|---|---|---|
| 1001 | Poste … | Unidad | Q 1,250.00 | ☑ | ☐ | ☐ | ☐ | ☐ | ☐ |
| 1002 | Conductor … | Kilómetro | Q 8,400.00 | ☑ | ☐ | ☐ | ☐ | ☑ | 🔒 |
| 1003 | Transformador … | Unidad | Q 9,800.00 | ☑ | ☑ | ☐ | ☐ | 🔒 | ☑ |

*(Las 2 columnas de voltaje son nuevas. Las columnas de línea siguen siendo MT / MT+BT / BT / G. 🔒 = casilla deshabilitada porque la otra ya está marcada.)*

### 7.2 Modos de los campos superiores

| Situación | Código | Componente | Unidad | Precio Unitario | Agregar |
|---|---|---|---|---|---|
| Se escribe/elige un **código existente** | Editable | Autocompleta | 🔒 bloqueado | 🔒 bloqueado | 🔒 bloqueado |
| Se escribe un **código o componente nuevo** | Solo números | Texto | ✅ desbloqueado | ✅ desbloqueado | Se activa al llenar todo y validar |
| Se **selecciona una fila** de la tabla | Autocompleta | Autocompleta | – | – | – |

### 7.3 Validaciones al agregar

| Campo | Regla |
|---|---|
| Código | Solo dígitos, único |
| Componente | No vacío, único (sin importar mayúsculas/acentos) |
| Unidad | Texto, no vacío |
| Precio Unitario | Número ≥ 0 con 2 decimales; se muestra en formato `Q#,##0.00` |
| Casillas MT/MT+BT/BT/G (y voltaje: ver 7.5) | Pueden quedar **todas desmarcadas** (decidido). Consecuencia: ese componente no aparecerá en el autocompletado normal de la Principal; solo se verá aquí y con `Ctrl + Espacio` |

### 7.4 Tabla

| Acción | Resultado |
|---|---|
| Clic en fila | Autocompleta Código y Componente arriba |
| Editar texto de código/descripción | Cambia en la **BD principal** (aplica a todo) |
| `Supr` con fila seleccionada | Elimina la fila (con confirmación) |
| Arrastrar fila | Cambia el orden; el orden se refleja en el Excel |
| Casilla marcada | `1` en el Excel (activo) |
| Casilla desmarcada | `0` en el Excel (inactivo) |
| Orden por defecto | El del Excel |

**⚠ Puntos que conviene cuidar:**

| Riesgo | Propuesta |
|---|---|
| Arrastrar filas mientras hay un filtro de búsqueda activo (el orden visible no es el real) | Permitir arrastrar **solo sin filtro** |
| Cambiar el código de un componente que ya está usado en las tablas de un proyecto | Actualizar también las referencias en esos proyectos |
| Eliminar un componente que ya está en un proyecto | Advertir con la lista de proyectos que lo usan (en RAIZ ya haces algo similar: "En uso") |
| Distinguir "estoy buscando" de "estoy editando" cuando ya hay código existente en el campo | Editar el código/componente **solo desde la fila seleccionada de la tabla** y que los campos de arriba, con código existente, sean de búsqueda |

### 7.5 Columnas de Voltaje (7.6 kV / 19.9 kV)

Sirven para el filtro de **Voltaje** de la Principal.

| 7.6 kV | 19.9 kV | Significado | ¿Aparece con Voltaje = 7.6? | ¿Aparece con Voltaje = 19.9? |
|---|---|---|---|---|
| ☐ | ☐ | Sirve para **ambas** líneas | ✅ | ✅ |
| ☑ | ☐ | Solo para 7.6 kV | ✅ | ❌ |
| ☐ | ☑ | Solo para 19.9 kV | ❌ | ✅ |
| ☑ | ☑ | **No permitido** | – | – |

| Regla | Detalle |
|---|---|
| Exclusión mutua | Si una casilla está marcada, la otra queda **deshabilitada** (no se puede marcar). Para cambiar, primero se desmarca la actual |
| Valor por defecto (componente nuevo) | Ambas desmarcadas (= sirve para las dos líneas) |
| Excel | Dos columnas con `1` / `0`: `0,0` = ambas · `1,0` = 7.6 kV · `0,1` = 19.9 kV. Nunca `1,1` |
| Validación al leer el Excel | Si una fila trae `1,1`, avisar al usuario y corregirla a `0,0` (o pedir elegir) |
| Independencia | Las columnas de voltaje **no** reemplazan a MT / MT+BT / BT / G; ambos filtros se aplican juntos (línea **y** voltaje) |

### 7.6 Guardar (💾)

| Paso | Detalle |
|---|---|
| 1 | Detecta si hay cambios; si no, el botón queda deshabilitado |
| 2 | Pregunta: **¿Guardar de forma Local o Global?** |
| Local | Solo afecta a este proyecto |
| Global | Modifica la **plantilla de Excel** para proyectos nuevos |
| Seguridad | Antes de modificar la plantilla global, crear una copia de respaldo (`plantilla_backup_FECHA.xlsx`) |
| `Esc` / `X` con cambios | *"¿Desea guardar los cambios?"* → Sí / No / Cancelar |

---

## 8. Total en letras

Formato: `Total en letras: <ENTERO EN LETRAS> QUETZALES CON <CENTAVOS EN LETRAS> CENTAVOS DE QUETZAL`

| Total | Resultado |
|---|---|
| Q 1,535,915.48 | `UN MILLÓN QUINIENTOS TREINTA Y CINCO MIL NOVECIENTOS QUINCE QUETZALES CON CUARENTA Y OCHO CENTAVOS DE QUETZAL` |
| Q 1.00 | `UN QUETZAL CON CERO CENTAVOS DE QUETZAL` *(singular: ⚠ confirmar)* |
| Q 21,000.50 | `VEINTIUN MIL QUETZALES CON CINCUENTA CENTAVOS DE QUETZAL` |
| Q 0.00 | `CERO QUETZALES CON CERO CENTAVOS DE QUETZAL` |

Reglas:

| Regla | Detalle |
|---|---|
| Apócope | `UN` delante de sustantivo (`VEINTIÚN MIL`, `UN MILLÓN`, `TREINTA Y UN MIL`) |
| `CIEN` vs `CIENTO` | `CIEN QUETZALES` pero `CIENTO UN QUETZALES` |
| Millones | `UN MILLÓN` / `DOS MILLONES`; `MILLÓN DE QUETZALES` cuando no hay resto (`UN MILLÓN DE QUETZALES`) |
| Redondeo | Redondear a 2 decimales **antes** de separar entero y centavos (`round(total, 2)`), para evitar `0.999` → `99 centavos` |
| Mayúsculas | Todo en mayúsculas, con tildes |
| Ejemplo tuyo | Tenía "CINTO" → es **CINCO** (lo corregí arriba) |

---

## 9. Atajos de teclado

| Tecla | Acción |
|---|---|
| `Esc` | Retroceder / cerrar (en Principal y Componentes pregunta si guardar) |
| `Esc` con popup de búsqueda abierto | Cierra solo el popup |
| `Alt + B` | Día / Noche (igual que RAIZ) |
| `Supr` | Eliminar fila seleccionada |
| `Ctrl + Espacio` | Búsqueda con **todos** los componentes (en Código o Componente) |
| `Ctrl + S` | Guardar *(sugerido)* |
| `F1` | Ayuda con la lista de atajos *(como RAIZ)* |

---

## 10. Cosas que faltaban o detecté al revisar

| # | Observación | Sugerencia |
|---|---|---|
| 1 | **El Excel no llegó** en los archivos (solo 5 imágenes y `RAIZ.py`) | Súbelo para revisar: nombres de hojas, tablas, formato de moneda, fórmulas del total, columnas 1/0, filtros |
| 2 | ✅ **Resuelto:** el Voltaje filtra los componentes con 2 columnas nuevas (7.6 kV / 19.9 kV). Ninguna marcada = sirve para ambas; si una está marcada, la otra se bloquea | **Pendiente menor:** si ya hay componentes agregados en las tablas y el usuario cambia el Voltaje, ¿se dejan tal cual, o se avisa de los que no corresponden al nuevo voltaje? Propongo avisar sin borrar |
| 3 | **Línea** tiene 4 opciones y el Excel tiene 4 tablas; pero las casillas de Componentes se llaman MT, MT+BT, BT, G | Mapa claro: MT↔Media Tensión, MT+BT↔Media y Baja, BT↔Baja, G↔General |
| 4 | ✅ **Resuelto:** Actualizar verifica totales, detecta componentes externos al Excel que no estén en la BD y actualiza VARIACION, OC, OTS y ATE | Ver 3.6; queda por definir dónde guardar el "total anterior" |
| 5 | Botón **Proyecto → Nuevo/Abrir** | Si hay cambios sin guardar, preguntar antes de cerrar el actual |
| 6 | ✅ **Resuelto:** misma tabla → se modifica la fila; otra tabla → se agrega fila nueva | Ver 3.5. Asumo que "modificar" = reemplazar la cantidad (no sumarla) |
| 7 | Hojas **VARIACION, OC, OTS, ATE** y las del **Ramal** | Entiendo que son de solo lectura (calculadas). El panel de agregar debe quedar deshabilitado al estar en ellas |
| 8 | Hojas ocultas de Visualización | Si todas se ocultan, la Principal no tendría pestañas. Mantener **ANE siempre visible** como mínimo |
| 9 | Triángulo ▲ de la Principal (junto a "Materiales") | Lo interpreto como *colapsar el panel de captura* para ganar espacio. ¿Correcto? |
| 10 | Títulos en tus imágenes | "Componenetes" → *Componentes* (typo en el dibujo) |
| 11 | Escape cierra la Principal | Aunque pida confirmación, puede cerrarse sin querer. Sugiero que en Principal `Esc` pregunte *"¿Desea salir?"* y, si hay cambios, *"¿Guardar?"* |
| 12 | Cambiar abreviación con hojas ya creadas | Recordar que el nombre de hoja = `ANE - <ABREVIACIÓN>`; renombrar hojas y reescribir las fórmulas que las referencian |
| 13 | Guardar en Excel con tablas ocultas | En el Excel final, las tablas vacías deben ocultarse (filas ocultas o no escribirse), y el Total debe seguir sumando bien |
| 14 | Licencia | En RAIZ validas `check_license()` al inicio; recomiendo conservarlo |

---

## 11. Orden sugerido de desarrollo

| Fase | Contenido |
|---|---|
| 1 | Tema día/noche + clase base de ventana (`Esc`, X, tamaños, scroll) |
| 2 | Componentes (BD, tabla, casillas, drag, guardar local/global) |
| 3 | Principal (búsqueda con popup, validación de cantidad, tablas dinámicas, total y total en letras) |
| 4 | Configuración + Edición + Visualización (proyectos, hojas, dependencias, color) |
| 5 | Exportación al Excel (copia de plantilla, nombres de hoja, fórmulas, formato Quetzales) |
| 6 | Ramal (suma de proyectos) |
| 7 | Botón Actualizar, atajos, ayuda F1 |
