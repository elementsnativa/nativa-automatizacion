"""
shopify_client.py
-----------------
Cliente Shopify para Nativa Elements.
Usa Admin API Token (SHOPIFY_ADMIN_TOKEN en .env).
"""

import os
import requests
from datetime import datetime, timezone
from dotenv import load_dotenv
from costos_excel import obtener_tabla_costos, obtener_costo_producto

load_dotenv()

STORE_URL   = os.getenv("SHOPIFY_STORE_URL")
ADMIN_TOKEN = os.getenv("SHOPIFY_ADMIN_TOKEN")
API_VERSION = "2025-01"


# ─── Autenticación ────────────────────────────────────────────────────────────

def _headers() -> dict:
    if not ADMIN_TOKEN:
        raise RuntimeError("Falta SHOPIFY_ADMIN_TOKEN en .env")
    return {
        "X-Shopify-Access-Token": ADMIN_TOKEN,
        "Content-Type": "application/json",
    }


def _base_url() -> str:
    return f"https://{STORE_URL}/admin/api/{API_VERSION}"


# ─── Paginación genérica ──────────────────────────────────────────────────────

def _paginar(endpoint: str, key: str, params: dict = None) -> list:
    """Recorre todas las páginas de un endpoint REST y devuelve lista completa."""
    url = f"{_base_url()}/{endpoint}.json"
    params = params or {}
    params.setdefault("limit", 250)
    resultados = []

    while url:
        resp = requests.get(url, headers=_headers(), params=params, timeout=30)
        if resp.status_code != 200:
            raise RuntimeError(f"Error Shopify API [{endpoint}]: {resp.status_code} - {resp.text}")
        resultados.extend(resp.json().get(key, []))

        # Link header para siguiente página
        link = resp.headers.get("Link", "")
        url = None
        params = {}  # params ya van en la URL del link
        for parte in link.split(","):
            if 'rel="next"' in parte:
                url = parte.split(";")[0].strip().strip("<>")
                break

    return resultados


# ─── Órdenes ──────────────────────────────────────────────────────────────────

def obtener_ordenes(fecha_desde: str = None, fecha_hasta: str = None) -> list:
    """
    Devuelve todas las órdenes (pagadas + pendientes) en el rango de fechas.
    fecha_desde / fecha_hasta: formato ISO "2026-01-01T00:00:00Z"
    """
    params = {"status": "any", "financial_status": "any"}
    if fecha_desde:
        params["created_at_min"] = fecha_desde
    if fecha_hasta:
        params["created_at_max"] = fecha_hasta

    return _paginar("orders", "orders", params)


def obtener_ordenes_mes(año: int, mes: int) -> list:
    """Órdenes de un mes específico."""
    desde = f"{año}-{mes:02d}-01T00:00:00Z"
    if mes == 12:
        hasta = f"{año+1}-01-01T00:00:00Z"
    else:
        hasta = f"{año}-{mes+1:02d}-01T00:00:00Z"
    return obtener_ordenes(desde, hasta)


# ─── Reembolsos ───────────────────────────────────────────────────────────────

def obtener_reembolsos_orden(order_id: int) -> list:
    """Reembolsos de una orden específica."""
    return _paginar(f"orders/{order_id}/refunds", "refunds")


def obtener_reembolsos_mes(año: int, mes: int) -> list:
    """
    Devuelve todos los reembolsos del mes junto con el order_id de origen.
    Itera sobre órdenes que tengan reembolsos.
    """
    ordenes = obtener_ordenes_mes(año, mes)
    reembolsos = []
    for orden in ordenes:
        if orden.get("refunds"):
            for ref in orden["refunds"]:
                ref["order_id"] = orden["id"]
                ref["order_name"] = orden["name"]
                reembolsos.append(ref)
    return reembolsos


# ─── Productos ────────────────────────────────────────────────────────────────

def obtener_productos() -> list:
    """Todos los productos con variantes, precios y costos."""
    return _paginar("products", "products")


def obtener_productos_dict() -> dict:
    """Devuelve dict {variant_id: {titulo, sku, precio, costo, inventario}}"""
    productos = obtener_productos()
    resultado = {}
    for prod in productos:
        for var in prod.get("variants", []):
            resultado[var["id"]] = {
                "producto":     prod["title"],
                "variante":     var["title"],
                "sku":          var.get("sku", ""),
                "precio":       float(var.get("price", 0)),
                "costo":        float(var.get("cost", 0) or 0),
                "inventario":   var.get("inventory_quantity", 0),
            }
    return resultado


# ─── Inventario ───────────────────────────────────────────────────────────────

def obtener_locations() -> list:
    """Ubicaciones/bodegas de la tienda."""
    return _paginar("locations", "locations")


def obtener_inventory_levels(location_id: int = None) -> list:
    """Niveles de inventario por location."""
    params = {}
    if location_id:
        params["location_ids"] = location_id
    return _paginar("inventory_levels", "inventory_levels", params)


# ─── Resumen financiero ───────────────────────────────────────────────────────

def resumen_financiero_mes(año: int, mes: int, archivo_excel: str = None) -> dict:
    """
    Calcula el resumen financiero de un mes desde Shopify.

    Retorna:
    {
        "venta_bruta":   float,  # suma de subtotal_price de órdenes pagadas
        "descuentos":    float,  # suma de total_discounts
        "reembolsos":    float,  # suma de monto reembolsado (FUENTE DE VERDAD)
        "venta_neta":    float,  # venta_bruta - descuentos - reembolsos
        "impuestos":     float,  # suma de total_tax
        "n_ordenes":     int,
        "n_reembolsos":  int,
        "costo_neto":    float,  # suma de (costo_unitario * cantidad) por línea
        "por_producto":  dict,   # desglose por producto
    }
    """
    ordenes = obtener_ordenes_mes(año, mes)

    # Costos desde Excel (fuente de verdad)
    if archivo_excel is None:
        archivo_excel = os.getenv("EXCEL_FILE", "GESTION FINAN PY.xlsx")
    fecha_ref = datetime(año, mes, 1)
    tabla_costos = obtener_tabla_costos(archivo_excel, fecha_ref)

    venta_bruta  = 0.0
    descuentos   = 0.0
    reembolsos   = 0.0
    impuestos    = 0.0
    costo_neto   = 0.0
    n_reembolsos = 0
    por_producto = {}

    for orden in ordenes:
        estado_pago = orden.get("financial_status", "")
        if estado_pago not in ("paid", "partially_refunded", "refunded", "partially_paid"):
            continue

        venta_bruta += float(orden.get("subtotal_price", 0))
        descuentos  += float(orden.get("total_discounts", 0))
        impuestos   += float(orden.get("total_tax", 0))

        # Reembolsos desde Shopify (fuente de verdad, NO del banco)
        for ref in orden.get("refunds", []):
            n_reembolsos += 1
            for tx in ref.get("refund_line_items", []):
                monto_ref = float(tx.get("subtotal", 0))
                reembolsos += monto_ref

        # Costo y desglose por producto (costos desde Excel)
        for item in orden.get("line_items", []):
            cant        = int(item.get("quantity", 0))
            precio_item = float(item.get("price", 0)) * cant
            titulo      = item.get("title", "Desconocido")

            costo_unit = obtener_costo_producto(titulo, tabla_costos)
            costo_item = costo_unit * cant
            costo_neto += costo_item

            if titulo not in por_producto:
                por_producto[titulo] = {
                    "unidades": 0, "venta_bruta": 0, "costo": 0
                }
            por_producto[titulo]["unidades"]    += cant
            por_producto[titulo]["venta_bruta"] += precio_item
            por_producto[titulo]["costo"]       += costo_item

    venta_neta = venta_bruta - descuentos - reembolsos

    return {
        "año":          año,
        "mes":          mes,
        "venta_bruta":  round(venta_bruta, 0),
        "descuentos":   round(descuentos, 0),
        "reembolsos":   round(reembolsos, 0),
        "venta_neta":   round(venta_neta, 0),
        "impuestos":    round(impuestos, 0),
        "n_ordenes":    len([o for o in ordenes if o.get("financial_status") in
                             ("paid","partially_refunded","refunded","partially_paid")]),
        "n_reembolsos": n_reembolsos,
        "costo_neto":   round(costo_neto, 0),
        "por_producto": por_producto,
    }


# ─── Test rápido ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("Probando conexión con Shopify...")
    print(f"Store: {STORE_URL}")

    try:
        token = _get_access_token()
        print(f"Token obtenido: {token[:20]}...")

        print("\nObteniendo órdenes de marzo 2026...")
        resumen = resumen_financiero_mes(2026, 3)
        print(f"  Órdenes pagadas:  {resumen['n_ordenes']}")
        print(f"  Venta bruta:      ${resumen['venta_bruta']:,.0f} CLP")
        print(f"  Descuentos:       ${resumen['descuentos']:,.0f} CLP")
        print(f"  Reembolsos:       ${resumen['reembolsos']:,.0f} CLP")
        print(f"  Venta neta:       ${resumen['venta_neta']:,.0f} CLP")
        print(f"  Costo neto:       ${resumen['costo_neto']:,.0f} CLP")
        margen = resumen['venta_neta'] - resumen['costo_neto']
        pct    = (margen / resumen['venta_neta'] * 100) if resumen['venta_neta'] else 0
        print(f"  Margen bruto:     ${margen:,.0f} CLP ({pct:.1f}%)")
        print(f"\nTop productos (por venta bruta):")
        for prod, datos in sorted(resumen['por_producto'].items(),
                                   key=lambda x: x[1]['venta_bruta'], reverse=True)[:8]:
            print(f"  {prod[:45]:45s} {datos['unidades']:4d} uds  ${datos['venta_bruta']:>12,.0f}  costo ${datos['costo']:>10,.0f}")

    except Exception as e:
        print(f"Error: {e}")
