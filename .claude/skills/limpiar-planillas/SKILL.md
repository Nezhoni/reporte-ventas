---
name: limpiar-planillas
description: Reglas para leer y limpiar planillas de negocios chilenos (CSV o Excel) con formatos distintos, nombres mal escritos, totales a mano y duplicados. Úsala cuando haya que unir o limpiar planillas de varias fuentes.
---

# Limpiar planillas desordenadas

## Lectura
- Detecta el separador del CSV (`,` `;` tab) en vez de suponerlo. Lee con encoding `utf-8-sig`.
- Busca la fila de encabezados en las primeras 15 filas (puede haber títulos arriba).
- Traduce nombres de columnas con una tabla de sinónimos (ej.: "Plato", "Producto", "Item" → producto).

## Limpieza
- Fechas: acepta `AAAA-MM-DD`, `DD/MM/AAAA`, `DD-MM-AAAA` y fechas de Excel.
- Números: quita `$` y espacios; `9.900` es nueve mil novecientos; acepta palabras ("tres").
- Nombres: compara sin tildes, sin mayúsculas y sin símbolos contra el catálogo oficial; si no hay coincidencia exacta usa similitud (difflib, umbral 0.75). Si nada coincide, conserva el nombre y regístralo como "producto desconocido".
- Elimina filas vacías y filas de "TOTAL" escritas a mano (el total se recalcula).

## Registro (obligatorio)
Cada corrección se guarda con: archivo, fila, problema, acción. Agrupa las correcciones repetidas ("14 filas: 'CHORILLANA' unificado como 'Chorrillana'") para que el registro sea legible.

## Verificación
Prueba siempre con un archivo de cada tipo de desorden y confirma que cada problema aparece en el registro.
