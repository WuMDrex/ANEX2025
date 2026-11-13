---
name: new-feature
description: Ayuda a diseñar e implementar una nueva funcionalidad en el código existente. Úsala cuando se pida agregar nuevas características, módulos o capacidades al proyecto. No la uses para explicar código básico, solo corregir errores o refactorizar sin agregar valor.
---

# Crear Nueva Funcionalidad

## Objetivo
Guiar al desarrollador en la planificación, diseño e implementación estructurada de una nueva característica dentro del código actual.

## Entradas
- Descripción clara de la funcionalidad que se desea agregar.
- Archivos, módulos o clases donde se espera integrar (opcional).
- Restricciones técnicas o librerías específicas que se deben utilizar (opcional).

Si la descripción es demasiado ambigua o vaga, pide detalles específicos (como el resultado esperado o el formato de datos) antes de escribir código.
Si no se indican los archivos de destino, explora el proyecto brevemente y sugiere el mejor lugar para la integración.

## Procedimiento
1. Analiza los requerimientos de la funcionalidad solicitada.
2. Identifica dónde debe integrarse leyendo el contexto de los archivos relevantes.
3. Presenta un breve plan de implementación (qué archivos se crearán o modificarán, y la lógica general).
4. Proporciona el código necesario, fragmentado de forma lógica y modular.
5. Explica brevemente cómo el nuevo código interactúa con el sistema que ya existe.
6. Propón un método rápido o un caso de prueba para verificar que la funcionalidad hace lo que debe.

## Límites
- No reescribas módulos enteros si no es estrictamente necesario para la nueva funcionalidad.
- No entregues un bloque masivo de código sin haber presentado el plan de implementación primero.
- Asegúrate de incluir las importaciones (imports) necesarias si utilizas nuevas librerías.
- Respeta el estilo de código, la arquitectura y el idioma de los comentarios del proyecto actual.

## Formato de salida
### Plan de Implementación
### Código Propuesto
### Integración y Contexto
### Cómo Probarlo

## Comprobación final
Confirma que el código propuesto resuelve exactamente la funcionalidad pedida, que no rompe la lógica previa del archivo y que has incluido instrucciones claras sobre cómo probarlo.