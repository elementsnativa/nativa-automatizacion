"""
costos_excel.py
---------------
Lee la hoja 'costos de venta' del Excel y mapea
títulos de productos Shopify a su costo unitario vigente.
"""

import openpyxl
from datetime import datetime
import re


# ─── Mapeo: título Shopify → categoría de costo ──────────────────────────────
# El orden importa: más específico primero

MAPEO_PRODUCTOS = [
    # Cinturón
    (r"cintur[oó]n",                "CINTURON"),
    # Compress
    (r"compress.*manga larga",      "COMPRESS MANGA LARGA"),
    (r"compress.*manga corta",      "COMPRESS MANGA CORTA"),
    (r"compress",                   "COMPRESS MANGA CORTA"),
    # Calza
    (r"calza.*larga",               "CALZA LARGA"),
    (r"calza.*corta",               "CALZA CORTA"),
    (r"calza",                      "CALZA LARGA"),
    # Peto
    (r"peto",                       "PETO"),
    # Buzo / jogger / baggy
    (r"buzo.*baggy",                "BUZO BAGGY"),
    (r"buzo.*jogger",               "BUZO JOGGER"),
    (r"buzo",                       "BUZO"),
    # Short
    (r"short.*chino",               "SHORT CHINO"),
    (r"short",                      "SHORT DEPORTIVO"),
    # Poleron french terry
    (r"french terry",               "Poleron french terry"),
    (r"hoodie.*french",             "Poleron french terry"),
    # Poleron boxy
    (r"poleron.*boxy",              "POLERON BOXY"),
    (r"hoodie.*boxy",               "POLERON BOXY"),
    # Poleron con cierre / zip
    (r"poleron.*cierre",            "POLERON CON CIERRE"),
    (r"4ter.*zip",                  "4TER ZIP"),
    (r"quarter.*zip",               "4TER ZIP"),
    # Poleron (genérico)
    (r"poleron.*minimal",           "Poleron minimal"),
    (r"poleron.*estampado",         "Poleron estampado"),
    (r"hoodie.*minimal",            "Poleron minimal"),
    (r"hoodie",                     "Poleron minimal"),
    (r"poleron",                    "Poleron minimal"),
    # Polera manga larga/corta mujer
    (r"polera.*manga larga.*mujer", "POLERA MANGA LARGA MUJER"),
    (r"polera.*manga corta.*mujer", "POLERA MANGA CORTA MUJER"),
    # Polera oversize / boxy
    (r"polera.*oversize",           "POLERA OVERSIZE"),
    (r"polera.*boxy",               "POLERA BOXY"),
    # Polera (genérica)
    (r"polera.*minimal",            "Polera Minimal"),
    (r"polera.*estampado",          "Polera estampado"),
    (r"polera",                     "Polera Minimal"),
    # Musculosa / stringer / dryfit
    (r"musculosa",                  "MUSCULOSA"),
    (r"stringer",                   "MUSCULOSA"),
    (r"dryfit",                     "MUSCULOSA"),
]


def clasificar_producto(titulo: str) -> str:
    """Mapea un título de Shopify a la categoría de costo del Excel."""
    t = titulo.lower()
    # Quitar tildes
    t = t.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    for patron, categoria in MAPEO_PRODUCTOS:
        if re.search(patron, t):
            return categoria
    return None  # No mapeado


# ─── Lectura de costos desde Excel ───────────────────────────────────────────

def leer_costos(archivo_excel: str) -> list:
    """
    Lee todos los registros de la hoja 'costos de venta'.
    Retorna lista de dicts: [{producto, fecha, costo}]
    """
    wb = openpyxl.load_workbook(archivo_excel, data_only=True)
    ws = wb["costos de venta"]

    costos = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        producto = row[0]
        fecha    = row[1]
        costo    = row[2]

        if not producto or not costo:
            continue

        # Normalizar fecha
        if isinstance(fecha, datetime):
            fecha_dt = fecha
        else:
            fecha_dt = None  # sin fecha → costo histórico sin fecha definida

        costos.append({
            "producto": str(producto).strip(),
            "fecha":    fecha_dt,
            "costo":    float(costo),
        })

    return costos


def obtener_tabla_costos(archivo_excel: str, fecha_referencia: datetime) -> dict:
    """
    Para cada categoría de producto, devuelve el costo vigente
    más reciente a la fecha_referencia.

    Si hay registros sin fecha, se usan como fallback.

    Retorna dict: {categoria: costo}
    """
    registros = leer_costos(archivo_excel)

    # Separar con fecha y sin fecha
    con_fecha   = [r for r in registros if r["fecha"] is not None]
    sin_fecha   = [r for r in registros if r["fecha"] is None]

    tabla = {}

    # Procesar registros con fecha: para cada producto, el más reciente <= fecha_referencia
    por_producto = {}
    for r in con_fecha:
        prod = r["producto"]
        if r["fecha"] <= fecha_referencia:
            if prod not in por_producto or r["fecha"] > por_producto[prod]["fecha"]:
                por_producto[prod] = r

    for prod, r in por_producto.items():
        tabla[prod] = r["costo"]

    # Fallback: registros sin fecha (si el producto no tiene registro con fecha)
    for r in sin_fecha:
        if r["producto"] not in tabla:
            tabla[r["producto"]] = r["costo"]

    return tabla


def obtener_costo_producto(titulo_shopify: str, tabla_costos: dict) -> float:
    """
    Dado un título de Shopify y la tabla de costos vigentes,
    devuelve el costo unitario o 0 si no se encuentra.
    """
    categoria = clasificar_producto(titulo_shopify)
    if not categoria:
        return 0.0
    return tabla_costos.get(categoria, 0.0)


# ─── Test ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import os
    from dotenv import load_dotenv
    load_dotenv()

    archivo = os.getenv("EXCEL_FILE", "GESTION FINAN PY.xlsx")
    fecha   = datetime(2026, 3, 1)

    print(f"Costos vigentes al {fecha.strftime('%d/%m/%Y')}:")
    tabla = obtener_tabla_costos(archivo, fecha)
    for prod, costo in sorted(tabla.items()):
        print(f"  {prod:35s} → ${costo:,.0f} CLP")

    print("\nTest clasificación:")
    titulos_test = [
        "BUZO PERFECT FIT NEGRO",
        "CINTURÓN DE LEVANTAMIENTO NATIVA",
        "HOODIE ZONE™ BOXY FIT NEGRO",
        "HOODIE FRENCH TERRY NEGRO",
        "POLERA OVERSIZE BLANCA",
        "CALZA DEPORTIVA LARGA",
        "Compress manga larga mujer",
    ]
    for t in titulos_test:
        cat   = clasificar_producto(t)
        costo = obtener_costo_producto(t, tabla)
        print(f"  {t:40s} → {str(cat):30s} → ${costo:,.0f}")
