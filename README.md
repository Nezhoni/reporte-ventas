# Reporte semanal automático de ventas · Cocina del Valle

> **Datos de ejemplo.** "Cocina del Valle" es una cadena ficticia y todas las
> cifras de este proyecto se generaron para demostración. No corresponden a
> ningún negocio real.

Proyecto de portafolio: un script que toma las planillas de ventas que envían
cada semana 3 sucursales de una cadena de restaurantes, las limpia, las une y
entrega un Excel listo para el dueño, con resumen, ranking de productos, qué
insumos pedir en cada sucursal y el registro de cada corrección hecha a los datos.

## El problema

Cada lunes, las sucursales de Providencia, Ñuñoa y Maipú envían sus ventas de
la semana, cada una en su propio formato:

| Sucursal | Cómo llega | Qué trae mal |
|---|---|---|
| Providencia | CSV con `;`, una fila por plato y día | títulos sobre el encabezado, montos como `$12.900`, nombres en MAYÚSCULAS, una fila copiada dos veces, un TOTAL sumado a mano |
| Ñuñoa | Excel | columnas con otros nombres (`Item`, `Cant.`), fechas como número (`46314`), tipeos (`Chorillana`, `lomo saltao`), abreviaciones (`Chorri`, `Lomo p.`, `Emp. pino`), cantidades en palabras (`tres`), filas vacías |
| Maipú | CSV con tabulador, exportado de la caja (una fila por pedido) | dos formatos de fecha mezclados, filas idénticas que son pedidos distintos, un producto que no está en el catálogo (`Promo almuerzo`) |

Hoy una persona tarda unas 3 horas en unir esto a mano, y se le cuelan errores.
Además, nadie sabe a tiempo qué insumos se están acabando en cada local.

## La solución

Un comando lee las planillas **tal como llegan** (sin tocarlas), y en unos
segundos:

1. **Limpia y une** los datos: detecta el formato de cada archivo, convierte
   fechas y montos, unifica los nombres contra el catálogo, elimina filas
   vacías, totales a mano y duplicados (solo donde un duplicado es un error).
2. **Registra cada corrección** con archivo, fila, problema y acción, separando
   los **errores reales** de las simples **conversiones de formato**. Nada se
   cambia en silencio.
3. **Calcula** el resumen semanal comparado con la semana anterior, el ranking
   de productos y la **reposición de insumos por sucursal**, según las recetas
   y el stock de cada local.
4. **Entrega un Excel** con 5 hojas, listo para el cliente.

## Antes y después

Con los datos de ejemplo de la semana del 19 al 25 de octubre de 2026:

| | Antes (a mano) | Después (con el script) |
|---|---|---|
| Tiempo para armar el reporte | unas 3 horas | **1,2 segundos** (medido) |
| Total de Providencia | $6.477.300, sumado a mano con una fila copiada dos veces | **$6.431.800**, y un aviso con la diferencia de $45.500 |
| Fila duplicada (Sopaipillas, 22/10) | se suma dos veces sin que nadie lo note | se elimina y queda registrada: fila 51, igual a la fila 50 |
| Pedidos iguales de Maipú | riesgo de borrarlos creyendo que son duplicados | se conservan: son clientes distintos |
| `Chorri`, `Chorillana`, `CHORRILLANA` | parecen tres productos distintos | los tres son **Chorrillana** |
| `Emp.` (¿de pino o de queso?) | alguien adivina | no se adivina: queda marcado para revisión |
| ¿Qué se corrigió? | no queda registro | **934 correcciones** anotadas: **37 errores reales** y 897 conversiones de formato |
| ¿De dónde sale una cifra? | hay que buscarla en 6 planillas | cada venta dice su archivo y fila de origen |
| ¿Qué pedir? | se revisa a ojo, si hay tiempo | por sucursal: **2 URGENTE y 3 Pedir** de 54 combinaciones, con la cantidad sugerida |

Y así se ve una fila antes y después de la limpieza:

| Archivo original | Fila tal como llega | Fila limpia |
|---|---|---|
| `providencia_2026-10-19.csv`, línea 4 | `19/10/2026;CHORRILLANA;6;$12.900;$77.400` | 19/10/2026 · Providencia · Chorrillana · 6 · $77.400 |
| `nunoa_2026-10-12.xlsx`, fila 43 | `46310 \| Chorri \| seis \| 77400` | 15/10/2026 · Ñuñoa · Chorrillana · 6 · $77.400 |
| `maipu_2026-10-19.csv`, línea 31 | `2026-10-19 → Promo almuerzo → 1 → 6900` | 19/10/2026 · Maipú · Promo almuerzo (sin categoría, avisado) · 1 · $6.900 |

## Cómo ejecutarlo

Requisitos: Python 3 con `pandas`, `openpyxl` y `pytest`.

```bash
pip install pandas openpyxl pytest   # si no los tienes

python3 crear_datos_ejemplo.py   # crea datos_semana/ con planillas de ejemplo desordenadas
python3 generar_reporte.py       # crea salida/reporte_semanal.xlsx
python3 -m pytest -q             # ejecuta los tests
```

Salida de `python3 generar_reporte.py` con los datos de ejemplo:

```
Reporte semanal de ventas · Cocina del Valle (DATOS DE EJEMPLO)
Archivos de ventas leídos: 6 (más inventario.xlsx)
Filas de ventas limpias:   1598
Correcciones registradas:  934 (33 líneas en 'Revisión de datos')
  errores reales:          37
  conversiones de formato: 897
Semana del reporte:        19/10/2026 al 25/10/2026
Ventas totales:            $15.637.000 (+6,3 % vs semana anterior)
Insumos:                   2 URGENTE, 3 Pedir, 49 OK
Reporte guardado en:       salida/reporte_semanal.xlsx
```

Detalles:

- `crear_datos_ejemplo.py` no sobrescribe `datos_semana/` si ya tiene archivos,
  porque son la entrada original. Para generarlos en otro lugar:
  `python3 crear_datos_ejemplo.py --carpeta otra_carpeta/`.
- `generar_reporte.py` solo lee `datos_semana/`; nunca modifica esos archivos.
  Acepta `--datos` y `--salida` para usar otras rutas.
- La carpeta `salida/` no se sube al repositorio: el Excel se vuelve a generar
  con el comando.

## Qué contiene el Excel

1. **Resumen**: ventas totales, venta promedio diaria, unidades y producto
   estrella, comparados con la semana anterior; tabla por sucursal (variación
   en verde si sube, en rojo si baja); gráfico de ventas por sucursal y gráfico
   de ventas diarias.
2. **Ranking de productos**: ordenados por ventas, con su % del total y cuántos
   puestos subieron o bajaron respecto de la semana anterior.
3. **Reposición de insumos**: cada sucursal tiene su propio stock. Para cada
   sucursal e insumo muestra cuántos días alcanza y cuánto pedir: URGENTE
   (rojo) si alcanza para menos de 2 días, Pedir (amarillo) si alcanza para
   menos de 5, OK (verde) en otro caso. Arriba resume las alertas por sucursal.
4. **Revisión de datos**: cada corrección, con archivo, filas, problema y
   acción, en dos grupos: primero los **errores corregidos** (duplicados,
   totales que no cuadran, tipeos, abreviaciones, productos desconocidos,
   cantidades en palabras) y después las **conversiones de formato** (fechas,
   montos con `$`, mayúsculas, filas vacías o de título). Arriba dice cuántos
   errores reales hubo. Las repetidas se agrupan ("91 filas: ...").
5. **Datos limpios**: todas las ventas ya limpias, con filtros y el archivo y la
   fila de origen de cada una, para poder rastrear cualquier cifra.

## Explicación para el cliente

**¿Cómo sé que no inventa nada?** Ningún dato se cambia en silencio. Cada
corrección queda anotada en "Revisión de datos" con el archivo y la fila
exacta, y cada venta de "Datos limpios" dice de qué archivo y fila viene. Si
una cifra parece rara, se puede buscar en la planilla original.

**¿Por qué borra una fila repetida en Providencia, pero no en Maipú?**
Providencia anota una fila por plato y día: si una fila aparece dos veces, es
un error de copia. La caja de Maipú anota una fila por pedido: dos filas
iguales son dos clientes que pidieron lo mismo, y borrarlas sería perder ventas.

**¿Y si el total que anotó la sucursal no coincide?** El reporte usa la suma
real de las filas y avisa en "Revisión de datos" cuánto difiere el total
escrito a mano, para que se pueda aclarar con la sucursal.

**¿Qué hace con los nombres mal escritos?** Los compara con el catálogo oficial.
`CHORRILLANA` y `Chorillana` pasan a "Chorrillana"; `Chorri` y `Lomo p.` se
reconocen como abreviaciones. Si una abreviación puede ser dos platos (`Emp.`
puede ser de pino o de queso), **no adivina**: la deja marcada para que una
persona decida.

**¿Cómo calcula qué pedir?** Con las recetas sabe cuánto de cada insumo lleva
cada plato. Multiplica por lo que vendió cada sucursal en la semana y obtiene
su consumo diario, que compara con el stock de esa misma sucursal. Si el stock
alcanza para pocos días, avisa y sugiere pedir lo que falta para cubrir una
semana completa más 2 días de margen por si el pedido se demora (consumo de
9 días − stock actual). La cantidad se redondea hacia arriba (a 0,1 en kilos y
litros, a número entero en unidades), porque es mejor pedir un poco de más que
quedarse corto.

**¿Qué necesita de nosotros?** Que el catálogo y las recetas del inventario
estén al día. Por ejemplo, "Promo almuerzo" se vende en Maipú pero no está en
el catálogo: el reporte la incluye en las ventas, pero no descuenta insumos y
lo avisa en la hoja de reposición.

## Estructura

```
crear_datos_ejemplo.py   genera datos_semana/ (datos ficticios y desordenados)
limpieza.py              lee y limpia las planillas; Registro de correcciones
calculos.py              resumen, ranking y reposición por sucursal (sin nada de Excel)
generar_reporte.py       arma el Excel final
datos_semana/            planillas de ejemplo de 2 semanas e inventario
tests/
  test_limpieza.py       un test por cada tipo de desorden
  test_integracion.py    limpia datos_semana/ y compara con las ventas reales
  test_calculos.py       cálculos con tablas pequeñas calculadas a mano
  test_reporte.py        reabre el Excel y revisa cifras, formato y diseño
```

## Cómo se comprobó

- **143 tests** con pytest. El test de integración compara, fecha por fecha,
  sucursal por sucursal y producto por producto, lo que sale de la limpieza con
  las ventas "verdaderas" con que se generaron los datos de ejemplo.
- Se comprobó que los tests detectan errores reales: si Providencia no borrara
  su fila duplicada, o si Maipú borrara sus pedidos idénticos, los tests fallan.
- Un agente revisor independiente revisó el reporte cuatro veces, recalculando
  a mano desde los archivos originales y sin usar el código del proyecto: las
  ventas por sucursal, productos puntuales y las 54 combinaciones de sucursal e
  insumo de la reposición. Todo coincidió. Su hallazgo válido (el TOTAL anotado
  a mano por Providencia no coincidía con la suma real y el reporte no lo
  avisaba) se corrigió y ahora queda registrado.

## Límites conocidos

- Los formatos de cada sucursal (`producto_dia` o `pedido`) se configuran en
  `limpieza.py` (`SUCURSALES`). Una sucursal nueva necesita agregarse ahí.
- Las abreviaciones que no se deducen del nombre (por ejemplo, "BL" para
  "Barros Luco") se agregan a mano en la tabla `ABREVIACIONES` de `limpieza.py`.
- El reporte compara la semana más reciente con la anterior que haya en
  `datos_semana/`.
- Los días de margen de la reposición se cambian en `DIAS_MARGEN` de `calculos.py`.
