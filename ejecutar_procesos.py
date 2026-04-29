"""
ARCHIVO PRINCIPAL MODULARIZADO
Ejecuta procesos de importación de cartolas y ventas mensuales

INSTRUCCIONES:
1. Configura los archivos a importar en las variables de abajo
2. Ejecuta este archivo (Run)
3. El script procesará ambos archivos y actualizará el Excel
"""

import os
import importlib.util

# Importar módulo con espacio en el nombre
spec_cartolas = importlib.util.spec_from_file_location("automatizacion_cartolas", "automatizacion cartolas.py")
automatizacion_cartolas = importlib.util.module_from_spec(spec_cartolas)
spec_cartolas.loader.exec_module(automatizacion_cartolas)
ProcesadorCartolas = automatizacion_cartolas.ProcesadorCartolas

# Importar módulo de ventas
from importar_ventas import importar_ventas_csv

# ============================================================================
# CONFIGURACIÓN - MODIFICA ESTAS VARIABLES SEGÚN NECESITES
# ============================================================================

# Archivo Excel principal
ARCHIVO_EXCEL = 'GESTION FINAN PY.xlsx'

# Archivo de cartola a importar (None si no quieres importar cartolas)
# Ejemplos: 'cartola (9).xls', 'cartola (10).xls', 'cartola enero.xls'
ARCHIVO_CARTOLA = "cartola (13).xls"  # Cambia a 'cartola (9).xls' si quieres importar cartolas
# NOTA: Si quieres importar cartolas, descomenta y cambia la línea de abajo:
# ARCHIVO_CARTOLA = 'cartola (9).xls'

# Archivo CSV de ventas a importar (None si no quieres importar ventas)
# Ejemplos: 'ventas_productos_enero_2026_con_mes.csv', 'ventas_productos_febrero_2026_con_mes.csv'
ARCHIVO_VENTAS_CSV = 'ventaaaa.csv'

# ============================================================================
# NO MODIFICAR NADA DEBAJO DE ESTA LÍNEA
# ============================================================================

def main():
    print("=" * 80)
    print("EJECUTOR PRINCIPAL - PROCESOS DE AUTOMATIZACIÓN")
    print("=" * 80)
    
    # Verificar que el archivo Excel existe
    if not os.path.exists(ARCHIVO_EXCEL):
        print(f"\n[ERROR] No se encontró el archivo Excel: {ARCHIVO_EXCEL}")
        print("Por favor, verifica que el archivo existe en el directorio actual.")
        return
    
    # Procesar cartola si se especificó
    if ARCHIVO_CARTOLA:
        print(f"\n{'='*80}")
        print("PROCESO 1: IMPORTACIÓN DE CARTOLAS")
        print(f"{'='*80}")
        
        if not os.path.exists(ARCHIVO_CARTOLA):
            print(f"[ERROR] No se encontró el archivo de cartola: {ARCHIVO_CARTOLA}")
            print("Saltando importación de cartolas...")
        else:
            try:
                procesador = ProcesadorCartolas(archivo_gestion=ARCHIVO_EXCEL)
                procesador.procesar_cartola(ARCHIVO_CARTOLA, actualizar_excel=True)
                print("\n[OK] Cartola procesada correctamente")
            except Exception as e:
                print(f"\n[ERROR] Error al procesar cartola: {e}")
                import traceback
                traceback.print_exc()
    else:
        print(f"\n[INFO] No se especificó archivo de cartola - saltando importación de cartolas")
    
    # Procesar ventas si se especificó
    if ARCHIVO_VENTAS_CSV:
        print(f"\n{'='*80}")
        print("PROCESO 2: IMPORTACIÓN DE VENTAS MENSUALES")
        print(f"{'='*80}")
        
        if not os.path.exists(ARCHIVO_VENTAS_CSV):
            print(f"[ERROR] No se encontró el archivo CSV de ventas: {ARCHIVO_VENTAS_CSV}")
            print("Saltando importación de ventas...")
        else:
            try:
                importar_ventas_csv(ARCHIVO_VENTAS_CSV, ARCHIVO_EXCEL, mes_objetivo=4, año_objetivo=2026)
                print("\n[OK] Ventas importadas correctamente")
            except Exception as e:
                print(f"\n[ERROR] Error al importar ventas: {e}")
                import traceback
                traceback.print_exc()
    else:
        print(f"\n[INFO] No se especificó archivo CSV de ventas - saltando importación de ventas")
    
    print(f"\n{'='*80}")
    print("[OK] PROCESO COMPLETADO")
    print(f"{'='*80}")
    print(f"\nArchivo Excel actualizado: {ARCHIVO_EXCEL}")
    print("\nPuedes abrir el archivo Excel para verificar los cambios.")

if __name__ == "__main__":
    main()
