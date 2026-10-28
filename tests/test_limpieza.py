"""
Tests de limpieza: uno por cada tipo de desorden que traen las planillas.

Usan datos pequeños escritos aquí mismo, para no depender de datos_semana/.
"""
from datetime import date, datetime

import pandas as pd
import pytest
from openpyxl import Workbook

from limpieza import (
    Registro,
    buscar_encabezado,
    detectar_separador,
    emparejar_producto,
    leer_fecha,
    leer_numero,
    leer_planilla,
    limpiar_tabla,
    normalizar_columnas,
    quitar_duplicados,
    quitar_filas_basura,
    totales_a_mano,
    verificar_totales_a_mano,
)

CATALOGO = {
    "Chorrillana": {"categoria": "Platos de fondo", "precio": 12900},
    "Lomo a lo pobre": {"categoria": "Platos de fondo", "precio": 11900},
    "Lomo saltado": {"categoria": "Platos de fondo", "precio": 10900},
    "Pastel de choclo": {"categoria": "Platos de fondo", "precio": 9900},
    "Empanada de pino": {"categoria": "Entradas", "precio": 2900},
    "Empanada de queso": {"categoria": "Entradas", "precio": 2500},
}


def tabla(filas, columnas=("fecha", "producto", "cantidad", "total"), primera_fila=2):
    """Arma una tabla como la que entrega leer_planilla (fila 1 = encabezado)."""
    df = pd.DataFrame(filas, columns=list(columnas))
    df["fila_origen"] = range(primera_fila, primera_fila + len(df))
    return df


def problemas(registro):
    """Todo el texto del registro en minúsculas, para buscar palabras clave."""
    return " | ".join(f"{c['problema']} {c['accion']}" for c in registro.correcciones).lower()


def correcciones_de(registro, fila):
    """Texto de las correcciones de una fila, en minúsculas."""
    return " | ".join(
        f"{c['problema']} {c['accion']}" for c in registro.correcciones if c["fila"] == fila
    ).lower()


# --- Separador del CSV --------------------------------------------------------

@pytest.mark.parametrize("texto, esperado", [
    ("fecha,producto,cantidad\n2026-10-19,Chorrillana,2\n", ","),
    ("Fecha;Plato;Cantidad\n19/10/2026;CHORRILLANA;2\n", ";"),
    ("fecha\tproducto\tcantidad\n2026-10-19\tChorrillana\t2\n", "\t"),
    # El título tiene comas, pero los datos van separados por ';'.
    ("Ventas del 19/10, sucursal Providencia\nFecha;Plato;Cantidad\n19/10/2026;CHORRILLANA;2\n", ";"),
])
def test_detecta_separador(texto, esperado):
    assert detectar_separador(texto) == esperado


# --- Fila de encabezados ------------------------------------------------------

def test_encuentra_encabezado_bajo_filas_de_titulo():
    filas = [
        ["Cocina del Valle - Sucursal Providencia"],
        ["Ventas del 19/10/2026 al 25/10/2026"],
        ["Fecha", "Plato", "Cantidad", "Precio unitario", "Total"],
        ["19/10/2026", "CHORRILLANA", "6", "$12.900", "$77.400"],
    ]
    assert buscar_encabezado(filas) == 2


def test_encuentra_encabezado_en_la_primera_fila():
    filas = [["fecha", "producto", "cantidad"], ["2026-10-19", "Chorrillana", "2"]]
    assert buscar_encabezado(filas) == 0


def test_sin_encabezado_reconocible_da_error_claro():
    with pytest.raises(ValueError, match="encabezado"):
        buscar_encabezado([["a", "b"], ["1", "2"]])


# --- Nombres de columnas ------------------------------------------------------

@pytest.mark.parametrize("original, esperado", [
    ("Fecha", "fecha"), ("Día", "fecha"), ("fecha", "fecha"),
    ("Plato", "producto"), ("Item", "producto"), ("Producto", "producto"),
    ("Cantidad", "cantidad"), ("Cant.", "cantidad"),
    ("Precio unitario", "precio"), ("precio", "precio"),
    ("Total", "total"),
])
def test_traduce_sinonimos_de_columnas(original, esperado):
    assert normalizar_columnas([original]) == [esperado]


def test_columna_desconocida_se_conserva():
    assert normalizar_columnas(["Garzón"]) == ["Garzón"]


# --- Lectura de archivos ------------------------------------------------------

def test_lee_csv_con_titulos_arriba_y_punto_y_coma(tmp_path):
    ruta = tmp_path / "providencia.csv"
    ruta.write_text(
        "Cocina del Valle - Sucursal Providencia\n"
        "Ventas del 19/10/2026 al 25/10/2026\n"
        "Fecha;Plato;Cantidad;Precio unitario;Total\n"
        "19/10/2026;CHORRILLANA;6;$12.900;$77.400\n",
        encoding="utf-8-sig",
    )
    registro = Registro()
    df = leer_planilla(ruta, registro)
    assert list(df.columns) == ["fecha", "producto", "cantidad", "precio", "total", "fila_origen"]
    assert df.iloc[0]["producto"] == "CHORRILLANA"
    assert df.iloc[0]["fila_origen"] == 4  # línea 4 del archivo original
    assert "título" in problemas(registro)


def test_lee_excel_y_conserva_numero_de_fila_original(tmp_path):
    ruta = tmp_path / "nunoa.xlsx"
    libro = Workbook()
    hoja = libro.active
    hoja.append(["Día", "Item", "Cant.", "Total"])
    hoja.append([46314, "Chorri", "tres", 38700])
    hoja.append([])
    hoja.append([46314, "Lomo p.", 2, 23800])
    libro.save(ruta)

    df = leer_planilla(ruta, Registro())
    assert list(df.columns) == ["fecha", "producto", "cantidad", "total", "fila_origen"]
    assert list(df["fila_origen"]) == [2, 3, 4]
    assert df[df["producto"] == "Lomo p."].iloc[0]["fila_origen"] == 4


# --- Fechas -------------------------------------------------------------------

@pytest.mark.parametrize("valor", [
    "2026-10-19",            # AAAA-MM-DD
    "19/10/2026",            # DD/MM/AAAA
    "19-10-2026",            # DD-MM-AAAA
    46314,                   # número de serie de Excel
    " 19/10/2026 ",
    datetime(2026, 10, 19),  # celda de fecha de Excel
])
def test_lee_formatos_de_fecha(valor):
    assert leer_fecha(valor) == date(2026, 10, 19)


def test_fecha_se_lee_dia_primero():
    # 5 de octubre, no 10 de mayo.
    assert leer_fecha("05/10/2026") == date(2026, 10, 5)


@pytest.mark.parametrize("valor", ["", None, float("nan"), "ayer", "32/10/2026"])
def test_fecha_invalida_devuelve_none(valor):
    assert leer_fecha(valor) is None


# --- Números ------------------------------------------------------------------

@pytest.mark.parametrize("valor, esperado", [
    ("$9.900", 9900), ("$ 9.900", 9900), ("$6.477.300", 6477300), ("9900", 9900),
    (" 12 ", 12), ("12", 12), (12, 12), (12.0, 12),
    ("tres", 3), ("Tres", 3), ("diez", 10),
])
def test_lee_numeros(valor, esperado):
    assert leer_numero(valor) == esperado


def test_punto_es_separador_de_miles_no_decimal():
    assert leer_numero("1.500") == 1500


@pytest.mark.parametrize("valor", ["", None, float("nan"), "muchos"])
def test_numero_invalido_devuelve_none(valor):
    assert leer_numero(valor) is None


# --- Nombres de productos -----------------------------------------------------

@pytest.mark.parametrize("escrito, producto, tipo", [
    ("Chorrillana", "Chorrillana", "exacto"),
    ("CHORRILLANA", "Chorrillana", "formato"),
    ("  chorrillana ", "Chorrillana", "formato"),
    ("EMPANADA DE PINO", "Empanada de pino", "formato"),
    ("Pástel de choclo", "Pastel de choclo", "formato"),
    ("Chorillana", "Chorrillana", "tipeo"),
    ("lomo saltao", "Lomo saltado", "tipeo"),
    ("Chorri", "Chorrillana", "abreviación"),
    ("Lomo p.", "Lomo a lo pobre", "abreviación"),
    ("Emp. pino", "Empanada de pino", "abreviación"),
])
def test_empareja_con_el_catalogo(escrito, producto, tipo):
    assert emparejar_producto(escrito, CATALOGO) == (producto, tipo)


@pytest.mark.parametrize("escrito", ["Emp.", "Lomo"])
def test_abreviacion_ambigua_no_se_adivina(escrito):
    # "Emp." puede ser de pino o de queso; "Lomo" puede ser a lo pobre o saltado.
    assert emparejar_producto(escrito, CATALOGO) == (None, "desconocido")


def test_producto_fuera_de_catalogo_es_desconocido():
    assert emparejar_producto("Promo almuerzo", CATALOGO) == (None, "desconocido")


# --- Filas vacías y TOTAL a mano ----------------------------------------------

def test_elimina_filas_vacias_y_fila_total():
    df = tabla([
        ["19/10/2026", "CHORRILLANA", "6", "$77.400"],
        [None, None, None, None],
        ["", "", "", ""],
        ["", "TOTAL", "", "$77.400"],
    ])
    registro = Registro()
    limpio = quitar_filas_basura(df, "providencia.csv", registro)
    assert list(limpio["fila_origen"]) == [2]
    assert sorted(c["fila"] for c in registro.correcciones) == [3, 4, 5]
    assert "vacía" in correcciones_de(registro, 3)
    assert "total" in correcciones_de(registro, 5)


def test_total_a_mano_que_no_coincide_con_la_suma_limpia_se_registra():
    original = tabla([
        ["22/10/2026", "SOPAIPILLAS", "13", "$45.500"],
        ["22/10/2026", "SOPAIPILLAS", "13", "$45.500"],   # copiada dos veces
        ["", "TOTAL", "", "$91.000"],                      # sumado a mano, con el duplicado
    ])
    totales = totales_a_mano(original)
    assert totales == [(4, 91000)]

    registro = Registro()
    quitar_filas_basura(original, "providencia.csv", registro)  # anota "fila de TOTAL ... eliminada"
    limpia = tabla([[date(2026, 10, 22), "Sopaipillas", 13, 45500]],
                   columnas=("fecha", "producto", "cantidad", "monto"))
    verificar_totales_a_mano(totales, limpia, "providencia.csv", registro)
    # Una sola línea para la fila 4: la eliminación y la diferencia juntas.
    assert registro.correcciones == [{
        "archivo": "providencia.csv",
        "fila": 4,
        "problema": "fila de TOTAL escrita a mano ($91.000) que no coincide con la suma de las filas "
                    "limpias ($45.500); diferencia de $45.500",
        "accion": "eliminada; se usa la suma de las filas limpias",
    }]


def test_total_a_mano_que_coincide_no_agrega_correccion():
    limpia = tabla([[date(2026, 10, 22), "Sopaipillas", 13, 45500]],
                   columnas=("fecha", "producto", "cantidad", "monto"))
    registro = Registro()
    verificar_totales_a_mano([(3, 45500)], limpia, "providencia.csv", registro)
    assert registro.correcciones == []


# --- Conversión de cada fila --------------------------------------------------

def test_limpiar_tabla_convierte_y_registra_cada_correccion():
    df = tabla([
        ["19/10/2026", "CHORRILLANA", "6", "$77.400"],     # fila 2: mayúsculas, monto con $
        [46314, "Chorri", "tres", 38700],                  # fila 3: fecha Excel, abreviación, palabras
        ["2026-10-19", "lomo saltao", 2, 21800],           # fila 4: tipeo
        ["19-10-2026", "Promo almuerzo", 1, 6900],         # fila 5: fuera del catálogo
    ])
    registro = Registro()
    limpio = limpiar_tabla(df, "x.csv", CATALOGO, registro)

    assert list(limpio["fecha"]) == [date(2026, 10, 19)] * 4
    assert list(limpio["producto"]) == ["Chorrillana", "Chorrillana", "Lomo saltado", "Promo almuerzo"]
    assert list(limpio["categoria"]) == ["Platos de fondo"] * 3 + ["Sin categoría"]
    assert list(limpio["cantidad"]) == [6, 3, 2, 1]
    assert list(limpio["monto"]) == [77400, 38700, 21800, 6900]
    assert list(limpio["fila_origen"]) == [2, 3, 4, 5]

    assert "mayúsculas" in correcciones_de(registro, 2)
    assert "$" in correcciones_de(registro, 2)
    assert "excel" in correcciones_de(registro, 3)
    assert "abreviación" in correcciones_de(registro, 3)
    assert "palabras" in correcciones_de(registro, 3)
    assert "tipeo" in correcciones_de(registro, 4)
    assert "desconocido" in correcciones_de(registro, 5)


def test_fila_sin_nada_que_corregir_no_se_registra():
    df = tabla([["2026-10-19", "Chorrillana", 2, 25800]])
    registro = Registro()
    limpiar_tabla(df, "x.csv", CATALOGO, registro)
    assert registro.correcciones == []


def test_monto_se_calcula_con_precio_unitario():
    df = tabla([["2026-10-19", "Chorrillana", 2, 12900]],
               columnas=("fecha", "producto", "cantidad", "precio"))
    limpio = limpiar_tabla(df, "maipu.csv", CATALOGO, Registro())
    assert limpio.iloc[0]["monto"] == 25800


def test_monto_que_no_calza_con_el_catalogo_se_registra():
    # 2 chorrillanas deberían ser $25.800 según el catálogo.
    df = tabla([["2026-10-19", "Chorrillana", 2, 20000]])
    registro = Registro()
    limpio = limpiar_tabla(df, "x.csv", CATALOGO, registro)
    assert limpio.iloc[0]["monto"] == 20000  # se respeta lo que dice la planilla
    assert "no calza" in correcciones_de(registro, 2)


def test_fila_con_fecha_o_cantidad_ilegible_se_descarta_y_registra():
    df = tabla([
        ["ayer", "Chorrillana", 1, 12900],
        ["2026-10-19", "Chorrillana", "muchos", 12900],
        ["2026-10-19", "Chorrillana", 1, 12900],
    ])
    registro = Registro()
    limpio = limpiar_tabla(df, "x.csv", CATALOGO, registro)
    assert list(limpio["fila_origen"]) == [4]
    assert "descartada" in correcciones_de(registro, 2)
    assert "descartada" in correcciones_de(registro, 3)


# --- Duplicados ---------------------------------------------------------------

def ventas_repetidas():
    return tabla([
        [date(2026, 10, 22), "Sopaipillas", 13, 45500],
        [date(2026, 10, 22), "Sopaipillas", 13, 45500],
        [date(2026, 10, 22), "Pisco sour", 4, 19600],
    ], columnas=("fecha", "producto", "cantidad", "monto"), primera_fila=50)


def test_borra_duplicado_en_formato_producto_por_dia():
    registro = Registro()
    limpio = quitar_duplicados(ventas_repetidas(), "producto_dia", "providencia.csv", registro)
    assert list(limpio["fila_origen"]) == [50, 52]
    assert [c["fila"] for c in registro.correcciones] == [51]
    assert "duplicada" in correcciones_de(registro, 51)


def test_conserva_filas_identicas_en_formato_por_pedido():
    # En la caja de Maipú dos pedidos iguales son dos ventas reales.
    registro = Registro()
    limpio = quitar_duplicados(ventas_repetidas(), "pedido", "maipu.csv", registro)
    assert list(limpio["fila_origen"]) == [50, 51, 52]
    assert registro.correcciones == []


def test_formato_desconocido_da_error():
    with pytest.raises(ValueError):
        quitar_duplicados(ventas_repetidas(), "otro", "x.csv", Registro())


# --- Registro de correcciones -------------------------------------------------

def test_registro_guarda_archivo_fila_problema_y_accion():
    registro = Registro()
    registro.agregar("providencia.csv", 95, "fila TOTAL escrita a mano", "eliminada")
    assert registro.correcciones == [{
        "archivo": "providencia.csv",
        "fila": 95,
        "problema": "fila TOTAL escrita a mano",
        "accion": "eliminada",
    }]


def test_registro_actualiza_una_correccion_ya_anotada():
    registro = Registro()
    registro.agregar("a.csv", 5, "fila de TOTAL escrita a mano", "eliminada")
    registro.actualizar("a.csv", 5, "fila de TOTAL escrita a mano", "problema completo", "acción completa")
    assert registro.correcciones == [
        {"archivo": "a.csv", "fila": 5, "problema": "problema completo", "accion": "acción completa"}
    ]


def test_registro_actualizar_agrega_si_no_habia_nada_anotado():
    registro = Registro()
    registro.actualizar("a.csv", 5, "no existe", "problema nuevo", "acción nueva")
    assert registro.correcciones == [
        {"archivo": "a.csv", "fila": 5, "problema": "problema nuevo", "accion": "acción nueva"}
    ]


def test_registro_agrupa_correcciones_repetidas():
    registro = Registro()
    for fila in (15, 4, 9):
        registro.agregar("nunoa.xlsx", fila, "nombre con tipeo: 'Chorillana'",
                         "unificado como 'Chorrillana'")
    registro.agregar("nunoa.xlsx", 7, "fila vacía", "eliminada")

    grupos = registro.agrupado()
    assert len(grupos) == 2
    assert grupos[0] == {
        "archivo": "nunoa.xlsx",
        "cantidad": 3,
        "filas": "4, 9, 15",
        "problema": "nombre con tipeo: 'Chorillana'",
        "accion": "unificado como 'Chorrillana'",
    }


def test_registro_agrupado_resume_listas_largas_de_filas():
    registro = Registro()
    for fila in range(2, 92):
        registro.agregar("providencia.csv", fila, "monto con '$' y punto de miles",
                         "convertido a número")
    grupo = registro.agrupado()[0]
    assert grupo["cantidad"] == 90
    assert len(grupo["filas"]) < 60  # legible, no una lista de 90 números
