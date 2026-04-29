"""
DESCARGA AUTOMÁTICA DE CARTOLA - BANCO DE CHILE (Portal Empresas)
Usa Playwright para iniciar sesión y descargar la cartola XLS.
Luego ejecuta el procesador de cartolas para actualizar GESTION FINAN PY.xlsx

USO:
    python descargar_cartola.py              # Descarga el mes actual
    python descargar_cartola.py --mes 3      # Descarga marzo
    python descargar_cartola.py --headless false  # Abre el navegador visible (debug)
"""

import os
import sys
import time
import glob
import shutil
import argparse
import tempfile
import importlib.util
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

# ── Cargar credenciales desde .env ────────────────────────────────────────────
load_dotenv()
RUT    = os.getenv("BANCO_RUT")
CLAVE  = os.getenv("BANCO_CLAVE")

if not RUT or not CLAVE:
    print("[ERROR] Faltan BANCO_RUT o BANCO_CLAVE en el archivo .env")
    sys.exit(1)

# ── Configuración ──────────────────────────────────────────────────────────────
DIRECTORIO_TRABAJO = Path(__file__).parent
EXCEL_GESTION      = DIRECTORIO_TRABAJO / "GESTION FINAN PY.xlsx"
CARPETA_CARTOLAS   = DIRECTORIO_TRABAJO  # las cartolas se guardan aquí

URL_PORTAL = (
    "https://portalempresas.bancochile.cl/mibancochile-web/front/empresa/"
    "index.html#/movimientos-cuentas/movimientos"
)

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo",  6: "Junio",   7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}


# ── Descarga via Playwright ────────────────────────────────────────────────────
def descargar_cartola(headless: bool = True, mes: int = None) -> Path | None:
    """
    Inicia sesión en el portal empresas BancoChile y descarga la cartola XLS.
    Retorna la ruta al archivo descargado, o None si falló.
    """
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

    if mes is None:
        mes = datetime.today().month
    anio = datetime.today().year
    nombre_mes = MESES_ES.get(mes, str(mes))

    print(f"\n{'='*60}")
    print(f"  DESCARGA CARTOLA BANCO DE CHILE")
    print(f"  Período: {nombre_mes} {anio}")
    print(f"{'='*60}\n")

    carpeta_descarga = tempfile.mkdtemp(prefix="cartola_bch_")

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=headless,
            args=["--no-sandbox", "--disable-dev-shm-usage"]
        )
        context = browser.new_context(
            accept_downloads=True,
            viewport={"width": 1280, "height": 800}
        )
        page = context.new_page()

        try:
            # ── PASO 1: Ir al portal ─────────────────────────────────────────
            print("[1/6] Abriendo portal Banco de Chile empresas...")
            # SPA Angular: usar 'load' en vez de 'networkidle'
            page.goto(URL_PORTAL, wait_until="load", timeout=30000)
            time.sleep(4)

            # ── PASO 2: Ingresar RUT ─────────────────────────────────────────
            print("[2/6] Ingresando RUT...")
            rut_input = page.wait_for_selector(
                "input[name='rut'], input[id*='rut'], input[placeholder*='RUT'], "
                "input[placeholder*='rut'], input[placeholder*='Rut']",
                timeout=15000
            )
            rut_input.click()
            rut_input.fill(RUT)
            time.sleep(0.5)

            # ── PASO 3: Ingresar clave ───────────────────────────────────────
            print("[3/6] Ingresando clave...")
            clave_input = page.wait_for_selector(
                "input[type='password']",
                timeout=10000
            )
            clave_input.click()
            clave_input.fill(CLAVE)
            time.sleep(0.5)

            # ── PASO 4: Click en Ingresar ────────────────────────────────────
            print("[4/6] Iniciando sesión...")
            boton = page.wait_for_selector(
                "button[type='submit'], button:has-text('Ingresar'), input[type='submit']",
                timeout=10000
            )
            boton.click()

            # Esperar respuesta de login — SPA no dispara networkidle estable
            # Esperamos que desaparezca el formulario de login o aparezca el dashboard
            time.sleep(5)
            try:
                page.wait_for_selector(
                    "input[type='password']",
                    state="hidden",
                    timeout=15000
                )
            except PWTimeout:
                pass  # continuar de todas formas
            time.sleep(3)

            print("       ✓ Sesión iniciada correctamente")

            # ── PASO 5: Ir a Saldo y Movimientos ─────────────────────────────
            print("[5/6] Navegando a Saldos y Movimientos...")

            # Intentar hacer click en "IR A SALDO Y MOVIMIENTOS" o link equivalente
            selectores_movimientos = [
                "button:has-text('IR A SALDO Y MOVIMIENTOS')",
                "button:has-text('Saldo y Movimientos')",
                "button:has-text('Saldos y Movimientos')",
                "a:has-text('IR A SALDO Y MOVIMIENTOS')",
                "a:has-text('Saldo y Movimientos')",
                "a:has-text('Saldos y Movimientos')",
                "[href*='movimientos']",
                "button:has-text('Movimientos')",
                "a:has-text('Movimientos')",
            ]

            encontrado = False
            for selector in selectores_movimientos:
                try:
                    btn = page.wait_for_selector(selector, timeout=5000)
                    if btn:
                        btn.click()
                        page.wait_for_load_state("networkidle", timeout=15000)
                        time.sleep(3)
                        print(f"       ✓ Click en: {selector}")
                        encontrado = True
                        break
                except PWTimeout:
                    continue

            if not encontrado:
                # Navegar directamente por URL
                page.goto(URL_PORTAL, wait_until="load", timeout=20000)
                time.sleep(4)
                page.screenshot(path=str(DIRECTORIO_TRABAJO / "debug_movimientos.png"))
                print("       ⚠ No se encontró el botón, screenshot guardado: debug_movimientos.png")

            time.sleep(2)

            # ── PASO 6: Descargar XLS ────────────────────────────────────────
            print("[6/6] Descargando cartola XLS...")

            # Botón de descarga — id estable confirmado por inspección
            boton_descarga = page.wait_for_selector("#descargar-btn", timeout=15000)
            if not boton_descarga:
                print("[ERROR] No se encontró #descargar-btn en la página.")
                page.screenshot(path=str(DIRECTORIO_TRABAJO / "debug_descarga.png"))
                return None

            print("       ✓ Botón #descargar-btn encontrado")
            boton_descarga.click()
            time.sleep(1)  # esperar que se abra el menú desplegable

            # Seleccionar opción "Excel" del menú mat-menu
            selectores_excel = [
                "button:has-text('Excel')",
                "[mat-menu-item]:has-text('Excel')",
                ".mat-menu-item:has-text('Excel')",
                "button.mat-menu-item:has-text('Excel')",
                "a:has-text('Excel')",
            ]

            opcion_excel = None
            for sel in selectores_excel:
                try:
                    opcion_excel = page.wait_for_selector(sel, timeout=5000)
                    if opcion_excel:
                        print(f"       ✓ Opción Excel encontrada: {sel}")
                        break
                except PWTimeout:
                    continue

            if not opcion_excel:
                print("[ERROR] No se encontró la opción 'Excel' en el menú.")
                page.screenshot(path=str(DIRECTORIO_TRABAJO / "debug_menu.png"))
                print("       Screenshot: debug_menu.png")
                return None

            # Capturar la descarga al hacer click en Excel
            with page.expect_download(timeout=30000) as dl_info:
                opcion_excel.click()

            descarga = dl_info.value
            archivo_temp = descarga.path()

            # Determinar nombre final
            nombre_final = f"cartola_{anio}_{mes:02d}.xls"
            destino = CARPETA_CARTOLAS / nombre_final
            shutil.copy(archivo_temp, destino)

            print(f"\n       ✓ Cartola guardada: {nombre_final}")
            return destino

        except Exception as e:
            print(f"\n[ERROR] {type(e).__name__}: {e}")
            try:
                page.screenshot(path=str(DIRECTORIO_TRABAJO / "debug_error.png"))
                print("       Screenshot guardado: debug_error.png")
            except Exception:
                pass
            return None
        finally:
            context.close()
            browser.close()


# ── Procesar cartola descargada ────────────────────────────────────────────────
def procesar_cartola(archivo_xls: Path):
    """Carga el ProcesadorCartolas y actualiza GESTION FINAN PY.xlsx"""
    print(f"\n{'='*60}")
    print(f"  PROCESANDO CARTOLA → EXCEL")
    print(f"{'='*60}")

    script = DIRECTORIO_TRABAJO / "automatizacion cartolas.py"
    spec = importlib.util.spec_from_file_location("automatizacion_cartolas", script)
    mod  = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)

    procesador = mod.ProcesadorCartolas(
        archivo_gestion=str(EXCEL_GESTION),
        archivo_memoria=str(DIRECTORIO_TRABAJO / "memoria_clasificaciones.json")
    )
    procesador.procesar_cartola(str(archivo_xls), actualizar_excel=True)
    print("\n[OK] GESTION FINAN PY.xlsx actualizado (FCL + EERR)")


# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="Descarga y procesa cartola Banco de Chile")
    parser.add_argument("--mes",      type=int, default=None,  help="Mes a descargar (1-12). Default: mes actual")
    parser.add_argument("--headless", type=str, default="true", help="'false' para ver el navegador (debug)")
    parser.add_argument("--solo-procesar", type=str, default=None,
                        help="Ruta a un XLS ya descargado — saltea el login y solo procesa")
    args = parser.parse_args()

    headless = args.headless.lower() != "false"

    # Modo solo-procesar (sin descarga)
    if args.solo_procesar:
        archivo = Path(args.solo_procesar)
        if not archivo.exists():
            print(f"[ERROR] No existe: {archivo}")
            sys.exit(1)
        procesar_cartola(archivo)
        return

    # Flujo completo: descargar + procesar
    archivo = descargar_cartola(headless=headless, mes=args.mes)

    if archivo:
        procesar_cartola(archivo)
        print(f"\n{'='*60}")
        print("  PROCESO COMPLETADO EXITOSAMENTE")
        print(f"  Cartola: {archivo.name}")
        print(f"  Excel:   GESTION FINAN PY.xlsx actualizado")
        print(f"{'='*60}\n")
    else:
        print("\n[FALLO] No se pudo completar la descarga.")
        print("  Opciones:")
        print("  1. Ejecuta con --headless false para ver qué ocurre")
        print("  2. Descarga la cartola manualmente y usa --solo-procesar cartola.xls")
        sys.exit(1)


if __name__ == "__main__":
    main()
