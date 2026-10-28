"""
Tests de los cálculos del reporte, con tablas pequeñas cuyos resultados están
calculados a mano en los comentarios.
"""
from datetime import date, timedelta

import pandas as pd
import pytest

from calculos import (
    estado_stock,
    productos_sin_receta,
    ranking_productos,
    redondear_hacia_arriba,
    reposicion,
    resumen_sucursales,
    semanas,
    ventas_diarias,
)

ANTERIOR, ACTUAL = date(2026, 10, 12), date(2026, 10, 19)


def venta(fecha, sucursal, producto, cantidad, monto, categoria="Platos de fondo"):
    return {
        "fecha": fecha, "semana": fecha - timedelta(days=fecha.weekday()),
        "sucursal": sucursal, "producto": producto, "categoria": categoria,
        "cantidad": cantidad, "monto": monto, "archivo": "x.csv", "fila_origen": 2,
    }


@pytest.fixture
def ventas():
    return pd.DataFrame([
        # Semana anterior: Providencia $36.700 (3 un.), Maipú $12.900 (1 un.)
        venta(date(2026, 10, 12), "Providencia", "Chorrillana", 2, 25800),
        venta(date(2026, 10, 13), "Providencia", "Lomo saltado", 1, 10900),
        venta(date(2026, 10, 12), "Maipú", "Chorrillana", 1, 12900),
        # Semana actual: Providencia $58.300 (7 un., 2 días), Maipú $10.900 (1 un., 1 día)
        venta(date(2026, 10, 19), "Providencia", "Chorrillana", 3, 38700),
        venta(date(2026, 10, 20), "Providencia", "Pisco sour", 4, 19600, "Bebidas"),
        venta(date(2026, 10, 19), "Maipú", "Lomo saltado", 1, 10900),
    ])


def test_identifica_semana_actual_y_anterior(ventas):
    assert semanas(ventas) == (ACTUAL, ANTERIOR)


def test_semanas_con_una_sola_semana(ventas):
    assert semanas(ventas[ventas["semana"] == ACTUAL]) == (ACTUAL, None)


# --- Resumen por sucursal -----------------------------------------------------

def test_resumen_por_sucursal_y_total(ventas):
    resumen = resumen_sucursales(ventas, ACTUAL, ANTERIOR)
    assert list(resumen["sucursal"]) == ["Providencia", "Maipú", "Total cadena"]

    prov = resumen.iloc[0]
    assert prov["ventas"] == 58300
    assert prov["ventas_anterior"] == 36700
    assert prov["variacion"] == pytest.approx(21600 / 36700)   # +58,9 %
    assert prov["unidades"] == 7
    assert prov["unidades_anterior"] == 3
    assert prov["dias_con_ventas"] == 2
    assert prov["promedio_diario"] == 29150                   # 58.300 / 2 días
    assert prov["promedio_diario_anterior"] == 18350          # 36.700 / 2 días
    assert prov["producto_estrella"] == "Chorrillana"         # $38.700 > $19.600

    maipu = resumen.iloc[1]
    assert maipu["variacion"] == pytest.approx(-2000 / 12900)  # baja 15,5 %
    assert maipu["producto_estrella"] == "Lomo saltado"

    total = resumen.iloc[2]
    assert total["ventas"] == 69200
    assert total["ventas_anterior"] == 49600
    assert total["promedio_diario"] == 34600                  # 69.200 / 2 días (19 y 20)
    assert total["producto_estrella"] == "Chorrillana"


def test_resumen_no_incluye_ticket_promedio(ventas):
    # No todas las sucursales registran pedidos individuales.
    columnas = " ".join(resumen_sucursales(ventas, ACTUAL, ANTERIOR).columns)
    assert "ticket" not in columnas


def test_variacion_vacia_si_no_hay_semana_anterior(ventas):
    solo_actual = ventas[ventas["semana"] == ACTUAL]
    resumen = resumen_sucursales(solo_actual, ACTUAL, None)
    assert resumen["variacion"].isna().all()
    assert resumen["ventas_anterior"].isna().all()


def test_sucursal_nueva_sin_ventas_la_semana_anterior(ventas):
    con_nunoa = pd.concat([ventas, pd.DataFrame([venta(date(2026, 10, 21), "Ñuñoa", "Chorrillana", 1, 12900)])])
    resumen = resumen_sucursales(con_nunoa, ACTUAL, ANTERIOR).set_index("sucursal")
    assert resumen.loc["Ñuñoa", "ventas_anterior"] == 0
    assert pd.isna(resumen.loc["Ñuñoa", "variacion"])  # no se divide por cero


def test_ventas_diarias_compara_cada_dia_con_la_semana_anterior(ventas):
    diarias = ventas_diarias(ventas, ACTUAL, ANTERIOR)
    assert list(diarias["dia"]) == ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    assert list(diarias["ventas"])[:3] == [49600, 19600, 0]            # 38.700 + 10.900 el lunes
    assert list(diarias["ventas_anterior"])[:3] == [38700, 10900, 0]   # 25.800 + 12.900 el lunes


# --- Ranking de productos -----------------------------------------------------

def test_ranking_ordena_por_monto_de_la_semana_actual(ventas):
    ranking = ranking_productos(ventas, ACTUAL, ANTERIOR)
    assert list(ranking["producto"]) == ["Chorrillana", "Pisco sour", "Lomo saltado"]
    assert list(ranking["posicion"]) == [1, 2, 3]
    assert list(ranking["unidades"]) == [3, 4, 1]
    assert list(ranking["monto"]) == [38700, 19600, 10900]
    assert ranking["porcentaje"].tolist() == pytest.approx([38700 / 69200, 19600 / 69200, 10900 / 69200])
    assert ranking["porcentaje"].sum() == pytest.approx(1)


def test_ranking_muestra_posicion_de_la_semana_anterior(ventas):
    # Semana anterior: Chorrillana $38.700 (1.º), Lomo saltado $10.900 (2.º), Pisco sour no se vendió.
    ranking = ranking_productos(ventas, ACTUAL, ANTERIOR).set_index("producto")
    assert ranking.loc["Chorrillana", "posicion_anterior"] == 1
    assert ranking.loc["Lomo saltado", "posicion_anterior"] == 2
    assert pd.isna(ranking.loc["Pisco sour", "posicion_anterior"])


# --- Reposición de insumos ----------------------------------------------------

RECETAS = pd.DataFrame([
    {"producto": "Chorrillana", "insumo": "Papas", "cantidad": 0.25},
    {"producto": "Chorrillana", "insumo": "Huevos", "cantidad": 2},
    {"producto": "Pisco sour", "insumo": "Limón", "cantidad": 0.05},
])
STOCK = pd.DataFrame([
    {"sucursal": "Providencia", "insumo": "Papas", "unidad": "kg", "stock": 0.5},
    {"sucursal": "Providencia", "insumo": "Huevos", "unidad": "unidad", "stock": 12},
    {"sucursal": "Providencia", "insumo": "Aceite", "unidad": "litro", "stock": 3},
    {"sucursal": "Maipú", "insumo": "Papas", "unidad": "kg", "stock": 0.5},
    {"sucursal": "Maipú", "insumo": "Huevos", "unidad": "unidad", "stock": 40},
    {"sucursal": "Ñuñoa", "insumo": "Limón", "unidad": "kg", "stock": 10},
])


@pytest.fixture
def ventas_insumos():
    return pd.DataFrame([
        venta(date(2026, 10, 14), "Providencia", "Chorrillana", 7, 90300),   # semana anterior: no cuenta
        venta(date(2026, 10, 20), "Providencia", "Chorrillana", 6, 77400),
        venta(date(2026, 10, 22), "Providencia", "Chorrillana", 4, 51600),   # Providencia: 10 en la semana
        venta(date(2026, 10, 22), "Maipú", "Chorrillana", 4, 51600),         # Maipú: 4 en la semana
        venta(date(2026, 10, 23), "Ñuñoa", "Pisco sour", 14, 68600, "Bebidas"),
        venta(date(2026, 10, 23), "Maipú", "Promo almuerzo", 5, 34500, "Sin categoría"),
    ])


def fila(tabla, sucursal, insumo):
    return tabla.set_index(["sucursal", "insumo"]).loc[(sucursal, insumo)]


def test_reposicion_calcula_consumo_cobertura_estado_y_sugerencia(ventas_insumos):
    tabla = reposicion(ventas_insumos, RECETAS, STOCK, ACTUAL)

    papas = fila(tabla, "Providencia", "Papas")  # 10 chorrillanas × 0,25 kg = 2,5 kg
    assert papas["consumo_semana"] == pytest.approx(2.5)
    assert papas["consumo_diario"] == pytest.approx(2.5 / 7)
    assert papas["dias_cobertura"] == pytest.approx(0.5 / (2.5 / 7))  # 1,4 días
    assert papas["estado"] == "URGENTE"
    # 9 días (7 + 2 de margen): 2,5 / 7 × 9 = 3,21 kg − 0,5 = 2,71 → 2,8 kg
    assert papas["cantidad_sugerida"] == pytest.approx(2.8)

    huevos = fila(tabla, "Providencia", "Huevos")  # 10 × 2 = 20 huevos; 12 alcanzan para 4,2 días
    assert huevos["consumo_semana"] == pytest.approx(20)
    assert huevos["estado"] == "Pedir"
    assert huevos["cantidad_sugerida"] == 14                         # 20 / 7 × 9 = 25,7 − 12 = 13,7 → 14


def test_cada_sucursal_usa_su_consumo_y_su_stock(ventas_insumos):
    # Mismo stock de papas (0,5 kg), distinto consumo: Providencia URGENTE, Maipú solo Pedir.
    tabla = reposicion(ventas_insumos, RECETAS, STOCK, ACTUAL)
    papas_maipu = fila(tabla, "Maipú", "Papas")                       # 4 × 0,25 = 1 kg
    assert papas_maipu["consumo_semana"] == pytest.approx(1.0)
    assert papas_maipu["dias_cobertura"] == pytest.approx(3.5)       # 0,5 / (1 / 7)
    assert papas_maipu["estado"] == "Pedir"
    assert papas_maipu["cantidad_sugerida"] == pytest.approx(0.8)     # 1 / 7 × 9 = 1,29 − 0,5 = 0,79 → 0,8
    assert fila(tabla, "Providencia", "Papas")["estado"] == "URGENTE"

    huevos_maipu = fila(tabla, "Maipú", "Huevos")                     # 4 × 2 = 8 huevos; 40 alcanzan 35 días
    assert huevos_maipu["consumo_semana"] == pytest.approx(8)
    assert huevos_maipu["estado"] == "OK"


def test_insumo_ok_no_tiene_sugerencia_ni_negativa(ventas_insumos):
    limon = fila(reposicion(ventas_insumos, RECETAS, STOCK, ACTUAL), "Ñuñoa", "Limón")
    assert limon["consumo_semana"] == pytest.approx(0.7)             # 14 × 0,05 kg
    assert limon["estado"] == "OK"
    assert pd.isna(limon["cantidad_sugerida"])                       # 0,7 − 10 sería −9,3


def test_insumo_ok_que_no_completa_7_dias_tampoco_tiene_sugerencia():
    # 7 chorrillanas = 1,75 kg de papas; 1,5 kg alcanzan 6 días: OK, aunque falte para 7.
    ventas = pd.DataFrame([venta(date(2026, 10, 20), "Providencia", "Chorrillana", 7, 90300)])
    stock = pd.DataFrame([{"sucursal": "Providencia", "insumo": "Papas", "unidad": "kg", "stock": 1.5}])
    papas = fila(reposicion(ventas, RECETAS, stock, ACTUAL), "Providencia", "Papas")
    assert papas["dias_cobertura"] == pytest.approx(6)
    assert papas["estado"] == "OK"
    assert pd.isna(papas["cantidad_sugerida"])


def test_insumo_sin_consumo_queda_ok_sin_cobertura(ventas_insumos):
    aceite = fila(reposicion(ventas_insumos, RECETAS, STOCK, ACTUAL), "Providencia", "Aceite")
    assert aceite["consumo_semana"] == 0
    assert pd.isna(aceite["dias_cobertura"])
    assert aceite["estado"] == "OK"
    assert pd.isna(aceite["cantidad_sugerida"])


def test_reposicion_ordena_urgentes_primero_y_luego_por_dias(ventas_insumos):
    tabla = reposicion(ventas_insumos, RECETAS, STOCK, ACTUAL)
    assert list(zip(tabla["estado"], tabla["sucursal"], tabla["insumo"])) == [
        ("URGENTE", "Providencia", "Papas"),   # 1,4 días
        ("Pedir", "Maipú", "Papas"),           # 3,5 días
        ("Pedir", "Providencia", "Huevos"),    # 4,2 días
        ("OK", "Maipú", "Huevos"),             # 35 días
        ("OK", "Ñuñoa", "Limón"),              # 100 días
        ("OK", "Providencia", "Aceite"),       # sin consumo
    ]


def test_insumo_de_receta_sin_stock_registrado_se_marca(ventas_insumos):
    recetas = pd.concat([RECETAS, pd.DataFrame([{"producto": "Pisco sour", "insumo": "Hielo", "cantidad": 0.1}])])
    hielo = fila(reposicion(ventas_insumos, recetas, STOCK, ACTUAL), "Ñuñoa", "Hielo")
    assert hielo["estado"] == "Sin dato de stock"
    assert hielo["consumo_semana"] == pytest.approx(1.4)              # 14 × 0,1


def test_productos_vendidos_sin_receta(ventas_insumos):
    sin_receta = productos_sin_receta(ventas_insumos, RECETAS, ACTUAL)
    assert sin_receta.to_dict("records") == [{"producto": "Promo almuerzo", "unidades": 5}]


@pytest.mark.parametrize("dias, estado", [
    (0.5, "URGENTE"), (1.99, "URGENTE"), (2, "Pedir"), (4.99, "Pedir"), (5, "OK"), (30, "OK"),
])
def test_umbrales_de_estado(dias, estado):
    assert estado_stock(dias) == estado


@pytest.mark.parametrize("valor, unidad, esperado", [
    (2.01, "kg", 2.1), (2.0, "kg", 2.0), (7.2, "unidad", 8), (8, "unidad", 8), (0, "litro", 0),
])
def test_sugerencia_se_redondea_hacia_arriba(valor, unidad, esperado):
    # Mejor pedir un poco más que quedarse corto.
    assert redondear_hacia_arriba(valor, unidad) == pytest.approx(esperado)
