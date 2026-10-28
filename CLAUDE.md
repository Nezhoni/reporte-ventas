# Proyecto: reporte semanal automático de ventas

## Objetivo
Proyecto de portafolio. Una cadena ficticia de restaurantes ("Cocina del Valle", 3 sucursales) envía cada semana sus ventas en planillas desordenadas. Un script las limpia, las une y genera un Excel con resumen, ranking de productos, reposición de insumos y registro de correcciones.

## Reglas
- Todo en español: código, comentarios, mensajes y el Excel final.
- Python 3 con pandas y openpyxl. No agregues otras librerías sin preguntarme.
- Nunca modifiques los archivos de `datos_semana/`. Son la entrada original.
- Toda corrección a los datos se registra (archivo, fila, problema, acción). Nada se cambia en silencio.
- Los datos son de ejemplo y deben decirlo en el reporte y en el README.

## Comandos
- Crear datos de ejemplo: `python3 crear_datos_ejemplo.py`
- Generar reporte: `python3 generar_reporte.py`
- Tests: `python3 -m pytest -q`

## Cómo trabajar
- Antes de un cambio grande, propón un plan y espera mi aprobación.
- Avanza en pasos pequeños. Después de cada cambio, ejecuta los tests y el script y muéstrame la salida.
- No digas "funciona" sin mostrar evidencia: el comando que corriste y lo que devolvió.
- Si no entiendo algo, explícamelo en simple. Debo poder explicárselo a un cliente.

## Errores conocidos (evítalos)
- Los sistemas de caja pueden tener filas idénticas legítimas (pedidos distintos). Solo trata como duplicado una fila repetida cuando el formato es "una fila por producto y día".
- Los montos chilenos usan punto de miles (`$9.900` = 9900). No lo confundas con decimales.
- Los datos de inventario de ejemplo deben ser realistas: pocas alertas urgentes, no todas.
