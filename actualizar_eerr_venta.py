"""
Ejecutar DESPUÉS de editar costos (columna K) u otros datos en la hoja 'venta'.
Recalcula y escribe en el EERR: ingresos (fila 3), rembolsos (5), costo neto (7).

1. Guarda y cierra GESTION FINAN PY.xlsx en Excel.
2. Ajusta MES y AÑO abajo si hace falta.
3. python actualizar_eerr_venta.py
"""
from importar_ventas import actualizar_eerr_desde_hoja_venta

ARCHIVO_EXCEL = "GESTION FINAN PY.xlsx"
MES = 1   # marzo
AÑO = 2026

if __name__ == "__main__":
    actualizar_eerr_desde_hoja_venta(ARCHIVO_EXCEL, mes=MES, año=AÑO)
