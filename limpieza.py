"""
Lectura y limpieza de las planillas de ventas de cada sucursal.

ESQUELETO DEL PASO 2: solo define los nombres de las funciones que prueban
los tests. La lógica se escribe en el Paso 3.
"""

PENDIENTE = "Pendiente: se implementa en el Paso 3"


class Registro:
    """Junta las correcciones hechas a los datos: archivo, fila, problema, acción."""

    def __init__(self):
        self.correcciones = []

    def agregar(self, archivo, fila, problema, accion):
        raise NotImplementedError(PENDIENTE)

    def agrupado(self):
        raise NotImplementedError(PENDIENTE)


def detectar_separador(texto):
    raise NotImplementedError(PENDIENTE)


def buscar_encabezado(filas, max_filas=15):
    raise NotImplementedError(PENDIENTE)


def normalizar_columnas(columnas):
    raise NotImplementedError(PENDIENTE)


def leer_planilla(ruta, registro):
    raise NotImplementedError(PENDIENTE)


def leer_fecha(valor):
    raise NotImplementedError(PENDIENTE)


def leer_numero(valor):
    raise NotImplementedError(PENDIENTE)


def emparejar_producto(nombre, catalogo):
    raise NotImplementedError(PENDIENTE)


def quitar_filas_basura(tabla, archivo, registro):
    raise NotImplementedError(PENDIENTE)


def limpiar_tabla(tabla, archivo, catalogo, registro):
    raise NotImplementedError(PENDIENTE)


def quitar_duplicados(tabla, formato, archivo, registro):
    raise NotImplementedError(PENDIENTE)
