"""
Cálculos del reporte: resumen por sucursal, ranking de productos y reposición de insumos.

Trabaja solo con tablas de pandas; no sabe nada de Excel. Recibe las ventas ya
limpias (la tabla que entrega limpieza.cargar_ventas).
"""
import math
from datetime import timedelta

import pandas as pd

from limpieza import SUCURSALES

TOTAL = "Total cadena"
NOMBRES_DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]

# Reposición: días que alcanza el stock con el consumo de la semana actual.
DIAS_URGENTE = 2   # menos de 2 días: URGENTE
DIAS_PEDIR = 5     # menos de 5 días: Pedir
DIAS_A_CUBRIR = 7  # la sugerencia cubre el consumo de una semana...
DIAS_MARGEN = 2    # ...más 2 días de margen por si el pedido se demora
ORDEN_ESTADOS = {"URGENTE": 0, "Pedir": 1, "Sin dato de stock": 2, "OK": 3}


def cargar_inventario(ruta):
    """Recetas (producto, insumo, cantidad) y stock por sucursal (sucursal, insumo, unidad, stock)."""
    recetas = pd.read_excel(ruta, sheet_name="Recetas").rename(
        columns={"Producto": "producto", "Insumo": "insumo", "Cantidad por unidad": "cantidad"})
    stock = pd.read_excel(ruta, sheet_name="Stock").rename(
        columns={"Sucursal": "sucursal", "Insumo": "insumo", "Unidad": "unidad",
                 "Stock al cierre del domingo": "stock"})
    return recetas, stock


def semanas(ventas):
    """(lunes de la semana actual, lunes de la semana anterior o None)."""
    lista = sorted(ventas["semana"].unique())
    return lista[-1], (lista[-2] if len(lista) > 1 else None)


def de_semana(ventas, semana):
    return ventas[ventas["semana"] == semana]


def variacion(actual, anterior):
    """Cambio porcentual (0.1 = +10 %). None si no hay con qué comparar."""
    if anterior is None or anterior == 0:
        return None
    return (actual - anterior) / anterior


def producto_estrella(ventas):
    if ventas.empty:
        return None
    return ventas.groupby("producto")["monto"].sum().idxmax()


def promedio_diario(ventas):
    """Venta promedio por día con ventas registradas."""
    dias = ventas["fecha"].nunique()
    return round(ventas["monto"].sum() / dias) if dias else 0


# --- Resumen ------------------------------------------------------------------

def resumen_sucursales(ventas, actual, anterior):
    """Una fila por sucursal y una de total: semana actual vs anterior.

    No incluye ticket promedio porque no todas las sucursales registran pedidos individuales.
    """
    presentes = set(ventas["sucursal"])
    orden = [nombre for nombre, _ in SUCURSALES.values() if nombre in presentes]
    orden += sorted(presentes - set(orden))

    filas = []
    for sucursal in orden + [TOTAL]:
        datos = ventas if sucursal == TOTAL else ventas[ventas["sucursal"] == sucursal]
        act = de_semana(datos, actual)
        ant = de_semana(datos, anterior) if anterior is not None else None
        ventas_act = int(act["monto"].sum())
        ventas_ant = int(ant["monto"].sum()) if ant is not None else None
        filas.append({
            "sucursal": sucursal,
            "ventas": ventas_act,
            "ventas_anterior": ventas_ant,
            "variacion": variacion(ventas_act, ventas_ant),
            "unidades": int(act["cantidad"].sum()),
            "unidades_anterior": int(ant["cantidad"].sum()) if ant is not None else None,
            "dias_con_ventas": act["fecha"].nunique(),
            "promedio_diario": promedio_diario(act),
            "promedio_diario_anterior": promedio_diario(ant) if ant is not None else None,
            "producto_estrella": producto_estrella(act),
        })
    return pd.DataFrame(filas)


def ventas_diarias(ventas, actual, anterior):
    """Ventas de la cadena por día de la semana: actual vs anterior."""
    por_fecha = ventas.groupby("fecha")["monto"].sum()
    filas = []
    for i, dia in enumerate(NOMBRES_DIAS):
        fecha = actual + timedelta(days=i)
        filas.append({
            "dia": dia,
            "fecha": fecha,
            "ventas": int(por_fecha.get(fecha, 0)),
            "ventas_anterior": (int(por_fecha.get(anterior + timedelta(days=i), 0))
                                if anterior is not None else None),
        })
    return pd.DataFrame(filas)


# --- Ranking ------------------------------------------------------------------

def ranking_productos(ventas, actual, anterior):
    """Productos de la semana actual ordenados por monto, con su posición de la semana anterior."""
    def por_producto(semana):
        tabla = (de_semana(ventas, semana)
                 .groupby(["producto", "categoria"], as_index=False)[["cantidad", "monto"]].sum()
                 .sort_values(["monto", "producto"], ascending=[False, True])
                 .reset_index(drop=True))
        tabla["posicion"] = range(1, len(tabla) + 1)
        return tabla

    ranking = por_producto(actual).rename(columns={"cantidad": "unidades"})
    ranking["porcentaje"] = ranking["monto"] / ranking["monto"].sum()
    posiciones_anteriores = {}
    if anterior is not None:
        tabla_anterior = por_producto(anterior)
        posiciones_anteriores = dict(zip(tabla_anterior["producto"], tabla_anterior["posicion"]))
    ranking["posicion_anterior"] = pd.array(
        [posiciones_anteriores.get(p) for p in ranking["producto"]], dtype="Int64")
    return ranking[["posicion", "producto", "categoria", "unidades", "monto",
                    "porcentaje", "posicion_anterior"]]


# --- Reposición ---------------------------------------------------------------

def estado_stock(dias):
    if dias < DIAS_URGENTE:
        return "URGENTE"
    if dias < DIAS_PEDIR:
        return "Pedir"
    return "OK"


def redondear_hacia_arriba(valor, unidad):
    """Unidades enteras para 'unidad'; un decimal para kg y litros."""
    decimales = 0 if unidad == "unidad" else 1
    factor = 10 ** decimales
    return math.ceil(round(valor * factor, 6)) / factor


def consumo_por_insumo(ventas, recetas, actual):
    """Σ (unidades vendidas × cantidad de la receta) en la semana actual, por sucursal e insumo."""
    vendidas = de_semana(ventas, actual).groupby(["sucursal", "producto"], as_index=False)["cantidad"].sum()
    uso = vendidas.merge(recetas.rename(columns={"cantidad": "por_unidad"}), on="producto")
    uso["consumo"] = uso["cantidad"] * uso["por_unidad"]
    return uso.groupby(["sucursal", "insumo"])["consumo"].sum()


def reposicion(ventas, recetas, stock, actual):
    """Qué insumos pedir en cada sucursal, según su consumo de la semana actual y su stock al cierre.

    Cantidad sugerida = consumo de 9 días (7 + 2 de margen) − stock actual, nunca
    negativa. Solo se sugiere para URGENTE y Pedir; en OK queda vacía para no confundir.
    """
    consumo = consumo_por_insumo(ventas, recetas, actual)
    con_stock = set(zip(stock["sucursal"], stock["insumo"]))
    sin_stock = [clave for clave in consumo.index if clave not in con_stock]
    tabla = pd.concat([
        stock[["sucursal", "insumo", "unidad", "stock"]],
        pd.DataFrame({"sucursal": [s for s, _ in sin_stock], "insumo": [i for _, i in sin_stock],
                      "unidad": "", "stock": float("nan")}),
    ], ignore_index=True)

    filas = []
    for fila in tabla.itertuples():
        consumo_semana = float(consumo.get((fila.sucursal, fila.insumo), 0))
        consumo_diario = consumo_semana / 7
        if pd.isna(fila.stock):
            dias, estado, sugerida = None, "Sin dato de stock", None
        elif consumo_diario == 0:
            dias, estado, sugerida = None, "OK", None
        else:
            dias = fila.stock / consumo_diario
            estado = estado_stock(dias)
            necesario = consumo_diario * (DIAS_A_CUBRIR + DIAS_MARGEN) - fila.stock
            sugerida = None if estado == "OK" else redondear_hacia_arriba(max(0, necesario), fila.unidad)
        filas.append({
            "sucursal": fila.sucursal, "insumo": fila.insumo, "unidad": fila.unidad, "stock": fila.stock,
            "consumo_semana": consumo_semana, "consumo_diario": consumo_diario,
            "dias_cobertura": dias, "estado": estado, "cantidad_sugerida": sugerida,
        })

    resultado = pd.DataFrame(filas)
    resultado["_orden"] = resultado["estado"].map(ORDEN_ESTADOS)
    return (resultado.sort_values(["_orden", "dias_cobertura", "sucursal", "insumo"], na_position="last")
            .drop(columns="_orden").reset_index(drop=True))


def productos_sin_receta(ventas, recetas, actual):
    """Productos vendidos en la semana actual que no descuentan insumos por no tener receta."""
    vendidas = de_semana(ventas, actual).groupby("producto", as_index=False)["cantidad"].sum()
    sin_receta = vendidas[~vendidas["producto"].isin(recetas["producto"])]
    return sin_receta.rename(columns={"cantidad": "unidades"}).reset_index(drop=True)
