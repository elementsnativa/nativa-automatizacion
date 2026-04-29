"""
SYNC VENTAS SHOPIFY → GESTION FINAN PY.xlsx
============================================
Extrae órdenes del mes desde la API de Shopify, construye el mismo formato
que usaba el CSV de Shopify Analytics, y llama a importar_ventas_csv()
para escribir en la hoja 'venta' y luego actualizar el EERR.

USO:
    python sync_shopify_ventas.py              # Mes actual
    python sync_shopify_ventas.py --mes 3      # Marzo 2026
    python sync_shopify_ventas.py --mes 3 --año 2026
"""

import os
import sys
import argparse
import tempfile
import csv
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

DIRECTORIO = Path(__file__).parent.parent
EXCEL      = DIRECTORIO / "GESTION FINAN PY.xlsx"

MESES_ES = {
    1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
    7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"
}


def extraer_ventas_shopify(año: int, mes: int) -> list[dict]:
    """
    Descarga las órdenes del mes desde la API de Shopify y devuelve
    una lista de filas con el mismo esquema del CSV de Analytics:
      Month, Product title, Product vendor, Product type,
      Net items sold, Gross sales, Discounts, Returns, Net sales, Taxes, Total sales
    """
    from shopify_client import obtener_ordenes_mes

    print(f"  Descargando órdenes de Shopify: {MESES_ES[mes]} {año}...")
    ordenes = obtener_ordenes_mes(año, mes)
    print(f"  {len(ordenes)} órdenes obtenidas")

    estados_validos = {"paid", "partially_refunded", "refunded", "partially_paid"}
    mes_fecha = datetime(año, mes, 1).strftime("%Y-%m-%d")

    # Acumular por producto
    por_producto: dict[str, dict] = {}

    for orden in ordenes:
        if orden.get("financial_status") not in estados_validos:
            continue

        descuento_orden    = float(orden.get("total_discounts", 0))
        n_items_orden      = sum(int(i.get("quantity", 0)) for i in orden.get("line_items", []))
        descuento_por_item = descuento_orden / n_items_orden if n_items_orden else 0

        # Reembolsos: monto por producto
        reembolsos_por_linea: dict[str, float] = {}
        for ref in orden.get("refunds", []):
            for rli in ref.get("refund_line_items", []):
                titulo_ref = rli.get("line_item", {}).get("title", "Desconocido")
                reembolsos_por_linea[titulo_ref] = (
                    reembolsos_por_linea.get(titulo_ref, 0) + float(rli.get("subtotal", 0))
                )

        for item in orden.get("line_items", []):
            titulo    = item.get("title", "Desconocido")
            vendor    = item.get("vendor", "")
            tipo      = item.get("product_type", "")
            cant      = int(item.get("quantity", 0))
            precio_u  = float(item.get("price", 0))
            impuesto  = float(item.get("total_discount", 0))  # tax por línea no está en REST v1
            gross     = precio_u * cant
            desc_item = descuento_por_item * cant
            ret_item  = reembolsos_por_linea.get(titulo, 0)
            net_sales = gross - desc_item - ret_item

            key = (titulo, vendor, tipo)
            if key not in por_producto:
                por_producto[key] = {
                    "Month":            mes_fecha,
                    "Product title":    titulo,
                    "Product vendor":   vendor,
                    "Product type":     tipo,
                    "Net items sold":   0,
                    "Gross sales":      0.0,
                    "Discounts":        0.0,
                    "Returns":          0.0,
                    "Net sales":        0.0,
                    "Taxes":            0.0,
                    "Total sales":      0.0,
                }
            por_producto[key]["Net items sold"] += cant
            por_producto[key]["Gross sales"]    += round(gross, 0)
            por_producto[key]["Discounts"]      += round(desc_item, 0)
            por_producto[key]["Returns"]        += round(ret_item, 0)
            por_producto[key]["Net sales"]      += round(net_sales, 0)
            por_producto[key]["Total sales"]    += round(net_sales, 0)

    filas = list(por_producto.values())
    print(f"  {len(filas)} productos únicos encontrados")
    return filas


def guardar_csv_temporal(filas: list[dict]) -> str:
    """Guarda las filas en un CSV temporal y retorna la ruta."""
    campos = [
        "Month", "Product title", "Product vendor", "Product type",
        "Net items sold", "Gross sales", "Discounts", "Returns",
        "Net sales", "Taxes", "Total sales",
    ]
    fd, ruta = tempfile.mkstemp(suffix=".csv", prefix="shopify_ventas_")
    with os.fdopen(fd, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=campos)
        writer.writeheader()
        for fila in filas:
            writer.writerow({k: fila.get(k, 0) for k in campos})
    return ruta


def main():
    parser = argparse.ArgumentParser(description="Sync ventas Shopify → Excel")
    parser.add_argument("--mes",  type=int, default=None)
    parser.add_argument("--año",  type=int, default=None)
    args = parser.parse_args()

    hoy  = datetime.today()
    mes  = args.mes  or hoy.month
    año  = args.año  or hoy.year

    print(f"\n{'='*60}")
    print(f"  SYNC VENTAS SHOPIFY → EERR")
    print(f"  Período: {MESES_ES[mes]} {año}")
    print(f"{'='*60}\n")

    if not EXCEL.exists():
        print(f"[ERROR] No existe: {EXCEL}")
        sys.exit(1)

    # 1. Extraer desde API
    try:
        filas = extraer_ventas_shopify(año, mes)
    except Exception as e:
        print(f"[ERROR] Shopify API: {e}")
        sys.exit(1)

    if not filas:
        print("[WARN] No hay ventas para el período. No se actualiza el Excel.")
        return

    # 2. CSV temporal → importar_ventas_csv (reutiliza toda la lógica existente)
    ruta_csv = guardar_csv_temporal(filas)
    print(f"\n  CSV temporal: {ruta_csv}")

    try:
        from importar_ventas import importar_ventas_csv

        # importar_ventas_csv ya actualiza hoja 'venta' Y el EERR en un solo paso
        print("\n[PASO 1] Escribiendo en hoja 'venta' y actualizando EERR...")
        importar_ventas_csv(ruta_csv, str(EXCEL))

    finally:
        os.unlink(ruta_csv)

    print(f"\n{'='*60}")
    print(f"  [OK] EERR actualizado con ventas Shopify: {MESES_ES[mes]} {año}")
    print(f"  Ingresos fila 3, Devoluciones fila 5, Costo neto fila 7")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
