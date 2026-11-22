# -*- coding: utf-8 -*-
"""
Genera ejemplos/plantilla_ejemplo.xlsx: una plantilla con datos ficticios
que cumple el esquema config/esquema.ejemplo.yaml.

    python ejemplos/generar_plantilla.py

Trae a proposito un par de casos que el validador resuelve con
advertencias (una sede ambigua y un contrato sin coordenadas), pero
ningun error bloqueante, para que se pueda generar el SQL.
"""
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill

SALIDA = Path(__file__).resolve().parent / "plantilla_ejemplo.xlsx"
ENCABEZADO = PatternFill("solid", fgColor="DDE4F0")


def hoja(wb, nombre, encabezados, filas, grupo=None):
    ws = wb.create_sheet(nombre)
    if grupo:
        ws.append(grupo)
        for celda in ws[1]:
            celda.font = Font(bold=True, italic=True)
    ws.append(encabezados)
    for celda in ws[ws.max_row]:
        celda.font = Font(bold=True)
        celda.fill = ENCABEZADO
    for fila in filas:
        ws.append(fila)
    for col in ws.columns:
        ancho = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(ancho + 2, 10), 40)


def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    hoja(wb, "TIPOS DE SERVICIO", ["nombre", "codigo", "penalidad", "meses_permanencia"], [
        ["Fibra Hogar", "F", 150000, 12],
        ["Radio Rural", "R", 0, 0],
        ["Fibra Empresarial", "E", 300000, 24],
    ])

    hoja(wb, "CATEGORIAS", ["nombre", "descripcion", "activo"], [
        ["Planes", "Planes mensuales", True],
        ["Equipos", "Equipos entregados en comodato", True],
        ["Instalación", "Cargos de única vez", True],
    ])

    hoja(wb, "PRODUCTOS", ["nombre", "referencia", "activo", "maneja_seriales", "se_vende", "se_compra", "precio",
                           "categoria", "unidad", "velocidad_mbps"], [
        ["Plan Fibra 300", "PF300", True, False, True, False, 89900, "Planes", "mes", 300],
        ["Plan Fibra 600", "PF600", True, False, True, False, 119900, "Planes", "mes", 600],
        ["Plan Radio 20", "PR20", True, False, True, False, 59900, "Planes", "mes", 20],
        ["Router WiFi 6", "RW6", True, True, True, True, 0, "Equipos", "unidad", None],
        ["Instalación estándar", "INST", True, False, True, False, 50000, "Instalación", "servicio", None],
    ])

    hoja(wb, "SEDES", ["tipo_documento", "consecutivo", "prefijo", "sufijo", "numero_inicial", "numero_actual",
                       "digitos", "activo", "sede", "ciudad"], [
        [1, "CONS MEDELLIN FIBRA", "MDF", None, 1, 1, 5, True, "MEDELLIN FIBRA", "Medellín"],
        [1, "CONS MEDELLIN RADIO", "MDR", None, 1, 1, 5, True, "MEDELLIN RADIO", "Medellín"],
        [1, "CONS ENVIGADO", "ENV", None, 100, 100, 5, True, "ENVIGADO CENTRO", "Envigado"],
    ])

    hoja(wb, "PERSONAS",
         ["tipo_identificacion", "documento", "nombres", "apellidos", "email",
          "plazo_pago", "regimen", "empleado_activo", "almacen", "cargo"], [
             [1, "43000001", "Laura", "Restrepo", "laura@example.com", 30, 1, None, None, None],
             [1, "71000002", "Andrés", "Zapata", "andres@example.com", 30, 1, None, None, None],
             [3, "900123456", "Panadería La Espiga S.A.S.", None, "contacto@example.com", 15, 2, None, None, None],
             [1, "1036000004", "Camila", "Ospina", "camila@example.com", None, None, True, None, "Técnica"],
         ],
         grupo=[None, None, None, None, None, "CLIENTE", None, "EMPLEADO", None, None])

    hoja(wb, "CONTRATOS", ["codigo_servicio", "cliente", "numero_contrato", "sede", "tipo_servicio", "ciudad",
                           "direccion", "dia_corte", "latitud", "longitud", "facturacion", "ip_fija", "nodo"], [
        ["laura.restrepo", "43000001", None, "MEDELLIN FIBRA", "Fibra Hogar", "Medellín", "Cra 43A # 1-50", 5,
         6.2087, -75.5679, "Factura electrónica", None, "NODO-POBLADO"],
        # La sede 'Medellín' es ambigua (hay dos); se resuelve por el tipo de servicio 'Radio'.
        ["andres.zapata", "71000002", None, "Medellín", "Radio Rural", "Medellín", "Vereda El Plan, finca 3", 5,
         6.1801, -75.6102, "Cuenta de cobro", None, None],
        # Sin coordenadas: advertencia, se asigna 0.
        ["espiga.principal", "900123456", None, "ENVIGADO CENTRO", "Fibra Empresarial", "Envigado",
         "Cl 38 Sur # 43-20", 15, None, None, "Factura electrónica", "192.0.2.10", "NODO-ENVIGADO"],
    ])

    hoja(wb, "DETALLE CONTRATOS", ["codigo_servicio", "producto", "cantidad", "precio"], [
        ["laura.restrepo", "Plan Fibra 600", 1, 119900],
        ["laura.restrepo", "Router WiFi 6", 1, 0],
        ["laura.restrepo", "Instalación estándar", 1, 50000],
        ["andres.zapata", "Plan Radio 20", 1, 59900],
        ["espiga.principal", "Plan Fibra 600", 2, 109900],
    ])

    wb.save(SALIDA)
    print(f"Plantilla creada: {SALIDA}")


if __name__ == "__main__":
    main()
