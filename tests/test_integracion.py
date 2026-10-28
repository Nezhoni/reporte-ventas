"""
Test de integración: limpia los archivos reales de datos_semana/ y comprueba que
el resultado coincide con las ventas "verdaderas" con que se generaron, y que
cada desorden sembrado aparece en el registro.

Solo lee datos_semana/; no modifica nada.
"""
from collections import Counter
from pathlib import Path

import pytest

from crear_datos_ejemplo import PRECIOS, PRODUCTO_FUERA_DE_CATALOGO, generar_ventas_reales
from limpieza import Registro, cargar_catalogo, cargar_ventas

CARPETA = Path(__file__).resolve().parent.parent / "datos_semana"


@pytest.fixture(scope="module")
def resultado():
    registro = Registro()
    catalogo = cargar_catalogo(CARPETA / "inventario.xlsx")
    ventas = cargar_ventas(CARPETA, catalogo, registro)
    return ventas, registro


def correcciones(registro, archivo, palabra):
    """Correcciones de un archivo cuyo problema o acción contiene la palabra."""
    return [
        c for c in registro.correcciones
        if c["archivo"] == archivo and palabra in f"{c['problema']} {c['accion']}".lower()
    ]


# --- Las cifras coinciden con las ventas verdaderas -----------------------------

def test_unidades_y_montos_coinciden_con_las_ventas_reales(resultado):
    ventas, _ = resultado
    conocidas = ventas[ventas["producto"] != PRODUCTO_FUERA_DE_CATALOGO[0]]

    esperado_unidades, esperado_monto = Counter(), Counter()
    for v in generar_ventas_reales():
        clave = (v["fecha"], v["sucursal"], v["producto"])
        esperado_unidades[clave] += v["cantidad"]
        esperado_monto[clave] += v["cantidad"] * PRECIOS[v["producto"]]

    obtenido_unidades, obtenido_monto = Counter(), Counter()
    for fila in conocidas.itertuples():
        clave = (fila.fecha, fila.sucursal, fila.producto)
        obtenido_unidades[clave] += fila.cantidad
        obtenido_monto[clave] += fila.monto

    assert obtenido_unidades == esperado_unidades
    assert obtenido_monto == esperado_monto


def test_cada_fila_conserva_archivo_y_fila_de_origen(resultado):
    ventas, _ = resultado
    assert ventas["archivo"].notna().all()
    assert (ventas["fila_origen"] > 1).all()
    assert set(ventas["sucursal"]) == {"Providencia", "Ñuñoa", "Maipú"}


# --- Cada desorden sembrado aparece en el registro --------------------------------

@pytest.mark.parametrize("semana", ["2026-10-12", "2026-10-19"])
def test_providencia_registra_titulos_montos_mayusculas_y_total(resultado, semana):
    _, registro = resultado
    archivo = f"providencia_{semana}.csv"
    assert len(correcciones(registro, archivo, "título")) == 2
    assert correcciones(registro, archivo, "$")
    assert correcciones(registro, archivo, "mayúsculas")
    assert len(correcciones(registro, archivo, "fila de total")) == 1


def test_providencia_registra_la_fila_duplicada_exacta(resultado):
    _, registro = resultado
    duplicadas = correcciones(registro, "providencia_2026-10-19.csv", "duplicada")
    assert [(c["fila"], c["problema"]) for c in duplicadas] == [
        (51, "fila duplicada (igual a la fila 50)")
    ]
    assert not correcciones(registro, "providencia_2026-10-12.csv", "duplicada")


def test_providencia_avisa_que_su_total_a_mano_no_coincide(resultado):
    # El TOTAL a mano de la semana actual incluye la fila duplicada de Sopaipillas (13 × $3.500).
    _, registro = resultado
    avisos = correcciones(registro, "providencia_2026-10-19.csv", "no coincide")
    assert [(c["fila"], c["problema"]) for c in avisos] == [(
        96,
        "el TOTAL escrito a mano ($6.477.300) no coincide con la suma de las filas limpias "
        "($6.431.800); diferencia de $45.500",
    )]
    # En la semana anterior no había duplicado y el total a mano sí coincide.
    assert not correcciones(registro, "providencia_2026-10-12.csv", "no coincide")


@pytest.mark.parametrize("semana", ["2026-10-12", "2026-10-19"])
def test_nunoa_registra_fechas_excel_palabras_y_filas_vacias(resultado, semana):
    _, registro = resultado
    archivo = f"nunoa_{semana}.xlsx"
    assert correcciones(registro, archivo, "excel")
    assert len(correcciones(registro, archivo, "palabras")) == 4
    assert len(correcciones(registro, archivo, "vacía")) == 2


def test_nunoa_registra_tipeos_y_abreviaciones(resultado):
    _, registro = resultado
    texto = " | ".join(
        f"{c['problema']} {c['accion']}" for c in registro.correcciones if c["archivo"].startswith("nunoa")
    )
    for escrito, oficial in [("Chorillana", "Chorrillana"), ("lomo saltao", "Lomo saltado")]:
        assert f"nombre con tipeo: '{escrito}' unificado como '{oficial}'" in texto
    for escrito, oficial in [("Chorri", "Chorrillana"), ("Lomo p.", "Lomo a lo pobre"),
                             ("Emp. pino", "Empanada de pino")]:
        assert f"nombre abreviado: '{escrito}' abreviación expandida a '{oficial}'" in texto


def test_nunoa_abreviacion_queda_en_su_fila_original(resultado):
    # Fila 43 de nunoa_2026-10-12.xlsx: ('Chorri', 'seis').
    _, registro = resultado
    fila_43 = [c for c in registro.correcciones
               if c["archivo"] == "nunoa_2026-10-12.xlsx" and c["fila"] == 43]
    texto = " | ".join(c["problema"] for c in fila_43)
    assert "nombre abreviado: 'Chorri'" in texto
    assert "cantidad escrita en palabras: 'seis'" in texto


@pytest.mark.parametrize("semana", ["2026-10-12", "2026-10-19"])
def test_maipu_registra_fechas_y_producto_desconocido_sin_borrar_pedidos(resultado, semana):
    ventas, registro = resultado
    archivo = f"maipu_{semana}.csv"
    assert correcciones(registro, archivo, "dd-mm-aaaa")
    assert len(correcciones(registro, archivo, "desconocido")) == 5
    assert not correcciones(registro, archivo, "duplicada")
    # Todas las líneas de pedidos del archivo siguen ahí (las filas idénticas no se borran).
    lineas = (CARPETA / archivo).read_text(encoding="utf-8-sig").strip().splitlines()
    assert len(ventas[ventas["archivo"] == archivo]) == len(lineas) - 1


def test_no_hay_filas_descartadas_ni_montos_que_no_calcen(resultado):
    # Los datos de ejemplo no traen fechas ni cantidades ilegibles ni precios mal.
    _, registro = resultado
    texto = " | ".join(c["problema"] for c in registro.correcciones)
    assert "ilegible" not in texto
    assert "no calza" not in texto
