"""
Script para restaurar el archivo Excel desde el backup más reciente
EJECUTA ESTE SCRIPT DESPUÉS DE CERRAR EL ARCHIVO EXCEL
"""

import shutil
import os

archivo_principal = 'GESTION FINAN PY.xlsx'
backups = [f for f in os.listdir('.') if 'backup' in f.lower() and f.endswith('.xlsx')]

if not backups:
    print("[ERROR] No se encontraron backups")
else:
    # Ordenar por fecha de modificación (más reciente primero)
    backups.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    backup_mas_reciente = backups[0]
    
    print(f"Restaurando desde: {backup_mas_reciente}")
    print("Asegúrate de que el archivo Excel esté CERRADO")
    
    try:
        # Eliminar el archivo corrupto si existe
        if os.path.exists(archivo_principal):
            os.remove(archivo_principal)
        
        # Copiar el backup
        shutil.copy2(backup_mas_reciente, archivo_principal)
        print(f"[OK] Archivo restaurado: {archivo_principal}")
        print("Ahora puedes abrir el archivo Excel")
    except PermissionError:
        print("[ERROR] El archivo está abierto. Por favor, CIERRA el archivo Excel y ejecuta este script nuevamente")
    except Exception as e:
        print(f"[ERROR] {e}")
