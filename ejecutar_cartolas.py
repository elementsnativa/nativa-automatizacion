"""
MÓDULO 1: IMPORTACIÓN DE MOVIMIENTOS DE CARTOLAS
Importa una cartola y actualiza cartolas cta cte, EERR y FCL.

INSTRUCCIONES:
1. Configura ARCHIVO_EXCEL y ARCHIVO_CARTOLA abajo
2. Ejecuta este archivo (Run)
"""

import os
import importlib.util

spec = importlib.util.spec_from_file_location("automatizacion_cartolas", "automatizacion cartolas.py")
automatizacion_cartolas = importlib.util.module_from_spec(spec)
spec.loader.exec_module(automatizacion_cartolas)
ProcesadorCartolas = automatizacion_cartolas.ProcesadorCartolas

# ============================================================================
# CONFIGURACIÓN
# ============================================================================
ARCHIVO_EXCEL = 'GESTION FINAN PY.xlsx'
ARCHIVO_CARTOLA = 'cartola (4).xls'  # Cambia según la cartola a importar
# ============================================================================

def main():
    print("=" * 80)
    print("MÓDULO 1: IMPORTACIÓN DE CARTOLAS")
    print("=" * 80)
    if not os.path.exists(ARCHIVO_EXCEL):
        print(f"[ERROR] No se encontró: {ARCHIVO_EXCEL}")
        return
    if not os.path.exists(ARCHIVO_CARTOLA):
        print(f"[ERROR] No se encontró: {ARCHIVO_CARTOLA}")
        return
    procesador = ProcesadorCartolas(archivo_gestion=ARCHIVO_EXCEL)
    procesador.procesar_cartola(ARCHIVO_CARTOLA, actualizar_excel=True)
    print("\n[OK] Cartolas actualizadas correctamente")

if __name__ == "__main__":
    main()
