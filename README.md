# Reporte semanal automático de ventas · Cocina del Valle

> **Datos de ejemplo.** "Cocina del Valle" es una cadena ficticia y todas las
> cifras de este proyecto se generaron para demostración. No corresponden a
> ningún negocio real.

Proyecto de portafolio: un script que toma las planillas de ventas que envían
cada semana 3 sucursales de una cadena de restaurantes, las limpia, las une y
entrega un Excel listo para el dueño, con resumen, ranking de productos, qué
insumos pedir y el registro de cada corrección hecha a los datos.

## El problema

Cada lunes, las sucursales de Providencia, Ñuñoa y Maipú envían sus ventas de
la semana, cada una en su propio formato:

| Sucursal | Cómo llega | Qué trae mal |
|---|---|---|
| Providencia | CSV con `;`, una fila por plato y día | títulos sobre el encabezado, montos como `$12.900`, nombres en MAYÚSCULAS, una fila copiada dos veces, un TOTAL sumado a mano |
| Ñuñoa | Excel | columnas con otros nombres (`Item`, `Cant.`), fechas como número (`46314`), tipeos (`Chorillana`, `lomo saltao`), abreviaciones (`Chorri`, `Lomo p.`, `Emp. pino`), cantidades en palabras (`tres`), filas vacías |
| Maipú | CSV con tabulador, exportado de la caja (una fila por pedido) | dos formatos de fecha mezclados, filas idénticas que son pedidos distintos, un producto que no está en el catálogo (`Promo almuerzo`) |

Hoy una persona tarda unas 3 horas en unir esto a mano, y se le cuelan errores.
Por ejemplo, en los datos de ejemplo el TOTAL escrito a mano de Providencia
($6.477.300) estaba inflado en $45.500 por la fila copiada dos veces; el total
correcto es $6.431.800.

## Cómo usarlo

Requisitos: Python 3 con `pandas`, `openpyxl` y `pytest`.

```bash
python3 crear_datos_ejemplo.py   # crea datos_semana/ con planillas de ejemplo desordenadas
python3 generar_reporte.py       # crea salida/reporte_semanal.xlsx
python3 -m pytest -q             # ejecuta los tests
```

`crear_datos_ejemplo.py` no sobrescribe `datos_semana/` si ya tiene archivos,
porque son la entrada original. Para regenerarlos en otro lugar:
`python3 crear_datos_ejemplo.py --carpeta otra_carpeta/`.

`generar_reporte.py` solo lee `datos_semana/`; nunca modifica esos archivos.
También acepta `--datos` y `--salida` para usar otras rutas.

## Qué contiene el Excel

1. **Resumen**: ventas totales, venta promedio diaria, unidades y producto
   estrella, comparados con la semana anterior; tabla por sucursal (variación
   en verde si sube, en rojo si baja); gráfico de ventas por sucursal y gráfico
   de ventas diarias.
2. **Ranking de productos**: ordenados por ventas, con su % del total y cuántos
   puestos subieron o bajaron respecto de la semana anterior.
3. **Reposición de insumos**: cuántos días alcanza cada insumo y cuánto pedir.
   URGENTE (rojo) si alcanza para menos de 2 días, Pedir (amarillo) si alcanza
   para menos de 5, OK (verde) en otro caso.
4. **Revisión de datos**: cada corrección hecha a los datos, con archivo, filas,
   problema y acción. Las repetidas se agrupan ("91 filas: ...").
5. **Datos limpios**: todas las ventas ya limpias, con filtros y el archivo y la
   fila de origen de cada una, para poder rastrear cualquier cifra.

## Explicación para el cliente

**¿Qué hace?** Lee las planillas tal como las mandan las sucursales, corrige los
errores conocidos y arma el reporte en segundos. Lo que antes eran 3 horas de
copiar y pegar pasa a ser ejecutar un comando.

**¿Cómo sé que no inventa nada?** Ningún dato se cambia en silencio. Cada
corrección queda anotada en la hoja "Revisión de datos" con el archivo y la
fila exacta, y cada venta de "Datos limpios" dice de qué archivo y fila viene.
Si una cifra parece rara, se puede buscar en la planilla original.

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
cada plato. Multiplica por lo vendido en la semana y obtiene el consumo
diario. Si el stock alcanza para pocos días, avisa y sugiere pedir lo que falta
para cubrir una semana completa más 2 días de margen por si el pedido se demora
(consumo de 9 días − stock actual). La cantidad se redondea hacia arriba (a 0,1
en kilos y litros, a número entero en unidades), porque es mejor pedir un poco
de más que quedarse corto.

**¿Qué necesita de nosotros?** Que el catálogo y las recetas del inventario
estén al día. Por ejemplo, "Promo almuerzo" se vende en Maipú pero no está en
el catálogo: el reporte la incluye en las ventas, pero no descuenta insumos y
lo avisa en la hoja de reposición.

## Estructura

```
crear_datos_ejemplo.py   genera datos_semana/ (datos ficticios y desordenados)
limpieza.py              lee y limpia las planillas; Registro de correcciones
calculos.py              resumen, ranking y reposición (sin nada de Excel)
generar_reporte.py       arma el Excel final
tests/
  test_limpieza.py       un test por cada tipo de desorden
  test_integracion.py    limpia datos_semana/ y compara con las ventas reales
  test_calculos.py       cálculos con tablas pequeñas calculadas a mano
  test_reporte.py        reabre el Excel y revisa cifras, formato y diseño
```

## Cómo se comprobó

- **132 tests** con pytest. El test de integración compara, fecha por fecha,
  sucursal por sucursal y producto por producto, lo que sale de la limpieza con
  las ventas "verdaderas" con que se generaron los datos de ejemplo.
- Se comprobó que los tests detectan errores reales: si Providencia no borrara
  su fila duplicada, o si Maipú borrara sus pedidos idénticos, los tests fallan.
- Un agente revisor independiente recalculó a mano, desde los archivos
  originales y sin usar el código del proyecto, la venta semanal de Maipú, las
  ventas de Empanada de pino y el pedido de Huevos. Todo coincidió. Su único
  hallazgo válido, que el TOTAL anotado a mano por Providencia no coincidía con
  la suma real y el reporte no lo avisaba, se corrigió y ahora queda registrado.

## Límites conocidos

- Los formatos de cada sucursal (`producto_dia` o `pedido`) se configuran en
  `limpieza.py` (`SUCURSALES`). Una sucursal nueva necesita agregarse ahí.
- Las abreviaciones que no se deducen del nombre (por ejemplo, "BL" para
  "Barros Luco") se agregan a mano en la tabla `ABREVIACIONES` de `limpieza.py`.
- El reporte compara la semana más reciente con la anterior que haya en
  `datos_semana/`.
