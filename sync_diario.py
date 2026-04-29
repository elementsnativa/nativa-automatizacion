"""
sync_diario.py
--------------
Ejecutado automáticamente cada día a las 9:00 AM via launchd.
1. Sincroniza ventas del mes actual desde Shopify → hoja 'venta' + EERR
2. Actualiza inventario Shopify → BALANCE ESTIMADO (Existencias)
"""

import sys
import os
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

DIRECTORIO = Path(__file__).parent
LOG_DIR    = DIRECTORIO / "logs"
LOG_DIR.mkdir(exist_ok=True)
LOG_FILE   = LOG_DIR / f"sync_diario_{datetime.now().strftime('%Y%m')}.log"

MESES_ES = {
    1:"enero", 2:"febrero", 3:"marzo", 4:"abril", 5:"mayo", 6:"junio",
    7:"julio", 8:"agosto", 9:"septiembre", 10:"octubre", 11:"noviembre", 12:"diciembre",
}


def log(msg: str):
    ts  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    linea = f"[{ts}] {msg}"
    print(linea)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(linea + "\n")


def main():
    hoy = datetime.now()
    mes, año = hoy.month, hoy.year

    log("=" * 60)
    log(f"SYNC DIARIO NATIVA ELEMENTS — {hoy.strftime('%d/%m/%Y %H:%M')}")
    log("=" * 60)

    errores = []

    # ── 1. Cartola Banco de Chile ────────────────────────────────────
    log(f"\n[1/6] Descargando cartola Banco de Chile ({hoy.strftime('%d/%m/%Y')})...")
    try:
        from descargar_cartola import descargar_cartola, procesar_cartola
        archivo = descargar_cartola(headless=True, mes=mes)
        if archivo:
            procesar_cartola(archivo)
            log(f"  ✓ Cartola descargada y procesada: {archivo.name}")
        else:
            msg = "  [WARN] No se pudo descargar la cartola (ver debug_error.png)"
            log(msg)
            errores.append(msg)
    except Exception as e:
        msg = f"  [ERROR] Cartola: {e}"
        log(msg)
        errores.append(msg)
        import traceback
        log(traceback.format_exc())

    # ── 2. Ventas del mes ────────────────────────────────────────────
    log(f"\n[2/6] Sincronizando ventas {MESES_ES[mes]} {año}...")
    try:
        from sync_shopify_ventas import extraer_ventas_shopify, guardar_csv_temporal
        from importar_ventas import importar_ventas_csv

        filas = extraer_ventas_shopify(año, mes)
        if filas:
            ruta_csv = guardar_csv_temporal(filas)
            try:
                importar_ventas_csv(
                    ruta_csv,
                    str(DIRECTORIO / os.getenv("EXCEL_FILE", "GESTION FINAN PY.xlsx")),
                    mes_objetivo=mes,
                    año_objetivo=año,
                )
                log(f"  ✓ {len(filas)} productos sincronizados en hoja 'venta' + EERR")
            finally:
                os.unlink(ruta_csv)
        else:
            log("  [WARN] Sin ventas para el mes actual")
    except Exception as e:
        msg = f"  [ERROR] Ventas: {e}"
        log(msg)
        errores.append(msg)
        import traceback
        log(traceback.format_exc())

    # ── 3. Inventario BALANCE ESTIMADO ──────────────────────────────
    log(f"\n[3/6] Actualizando inventario Shopify → BALANCE ESTIMADO...")
    try:
        from inventario_shopify import actualizar_balance_inventario
        resumen = actualizar_balance_inventario()
        if resumen:
            log(f"  ✓ Existencias actualizadas: ${resumen['valor_total']:,.0f} CLP")
            log(f"    {resumen['n_variantes']} variantes con stock")
        else:
            log("  [WARN] No se obtuvo resumen de inventario")
    except Exception as e:
        msg = f"  [ERROR] Inventario BALANCE: {e}"
        log(msg)
        errores.append(msg)
        import traceback
        log(traceback.format_exc())

    # ── 4. Inventario Detallado ──────────────────────────────────────
    log(f"\n[4/6] Generando hoja 'Inventario Detallado'...")
    try:
        from inventario_detallado import generar_hoja_inventario
        r = generar_hoja_inventario()
        if r:
            log(f"  ✓ Inventario Detallado: {r['variantes']} variantes — valor ${r['valor']:,.0f} CLP")
    except Exception as e:
        msg = f"  [ERROR] Inventario Detallado: {e}"
        log(msg)
        errores.append(msg)
        import traceback
        log(traceback.format_exc())

    # ── 5. Ingresos por envíos ───────────────────────────────────────
    log(f"\n[5/7] Actualizando 'Ingresos Envíos' + EERR fila 4...")
    try:
        from sync_envios import sync_envios
        r = sync_envios()
        if r:
            log(f"  ✓ Envíos: {r['meses']} meses — neto ${r['total_neto']:,.0f} CLP")
    except Exception as e:
        msg = f"  [ERROR] Envíos: {e}"
        log(msg)
        errores.append(msg)
        import traceback
        log(traceback.format_exc())

    # ── 6. Reembolsos Mensuales ──────────────────────────────────────
    log(f"\n[6/7] Actualizando hoja 'Reembolsos Mensuales'...")
    try:
        from sync_reembolsos import sync_reembolsos
        r = sync_reembolsos()
        if r:
            log(f"  ✓ Reembolsos: {r['meses']} meses — neto total ${r['total_neto']:,.0f} CLP")
    except Exception as e:
        msg = f"  [ERROR] Reembolsos: {e}"
        log(msg)
        errores.append(msg)
        import traceback
        log(traceback.format_exc())

    # ── 6. Sincronizar hoja Ordenes ─────────────────────────────────
    log(f"\n[7/7] Sincronizando hoja 'Ordenes'...")
    try:
        from sync_ordenes import sync_ordenes
        r = sync_ordenes()
        if r:
            log(f"  ✓ Ordenes: {r['nuevas']} nuevos pedidos — {r['filas_agregadas']} filas")
        else:
            log("  [WARN] Sin resultado del sync de órdenes")
    except Exception as e:
        msg = f"  [ERROR] Ordenes: {e}"
        log(msg)
        errores.append(msg)
        import traceback
        log(traceback.format_exc())

    # ── Resultado ────────────────────────────────────────────────────
    log("\n" + "=" * 60)
    if errores:
        log(f"SYNC COMPLETADO CON {len(errores)} ERROR(ES)")
        for e in errores:
            log(f"  → {e}")
    else:
        log("SYNC COMPLETADO EXITOSAMENTE")
    log("=" * 60 + "\n")

    # ── Backup automático a GitHub ────────────────────────────────────
    import subprocess
    try:
        directorio = str(DIRECTORIO)
        subprocess.run(['git', '-C', directorio, 'add', '-A'], check=True, capture_output=True)
        subprocess.run(['git', '-C', directorio, 'commit', '-m',
                        f'sync: {hoy.strftime("%Y-%m-%d")}'],
                       check=True, capture_output=True)
        subprocess.run(['git', '-C', directorio, 'push'], check=True, capture_output=True)
        log("  ✓ Código respaldado en GitHub")
    except subprocess.CalledProcessError:
        log("  [INFO] GitHub: sin cambios nuevos en el código")


if __name__ == "__main__":
    main()
