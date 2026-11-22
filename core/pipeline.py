# -*- coding: utf-8 -*-
"""
pipeline.py
Orquesta la validacion completa de la plantilla, hoja por hoja, en el
orden correcto de dependencias (catalogos primero, luego lo que los usa).
"""
from core import esquema as esquema_mod
from core import validators as v
from core.catalogs import cargar_ciudades
from core.excel_loader import load_workbook, norm_key


def ejecutar_validacion(path_excel, conn=None, esquema=None):
    """
    esquema: Esquema ya cargado; si no se da, se usa el que resuelva
    core.esquema.ruta_esquema().
    conn: conexion psycopg2 OPCIONAL a una base destino ya poblada. Si se
    provee, ademas de la validacion normal (siempre offline), se corre
    una reconciliacion de matching tolerante de PRODUCTOS y PROYECTOS
    contra los catalogos reales (ver validators.reconciliar_*).
    """
    E = esquema or esquema_mod.cargar()
    wb = load_workbook(path_excel)
    ciudades_idx = cargar_ciudades(E)

    findings = []

    filas_tipo_contrato = v.load(E, wb, "tipos_contrato")
    f, tipo_contrato_index = v.validar_tipo_contrato(E, filas_tipo_contrato)
    findings += f

    filas_categorias = v.load(E, wb, "categorias")
    f, categorias_index = v.validar_categorias_productos(E, filas_categorias)
    findings += f

    filas_productos = v.load(E, wb, "productos")

    # PRODUCTOS POR PLAN: fuera de alcance por ahora. planes_norm y
    # materiales_norm quedan vacios -> 'es_principal' y 'cargo_unico' no se
    # infieren todavia y quedan en False.
    planes_norm, materiales_norm = set(), set()

    f, productos_index = v.validar_productos(E, filas_productos, categorias_index, planes_norm)
    findings += f

    filas_consec_proy = v.load(E, wb, "consecutivos_proyectos")
    f, consecutivos_index, proyectos_index = v.validar_consecutivo_y_proyectos(E, filas_consec_proy, ciudades_idx)
    findings += f

    filas_personas = v.load(E, wb, "personas")
    f, personas_index = v.validar_personas(E, filas_personas)
    findings += f

    filas_contratos = v.load(E, wb, "contratos")
    f, contratos_index = v.validar_contratos(E, filas_contratos, personas_index, proyectos_index, tipo_contrato_index,
                                             ciudades_idx, consecutivos_index=consecutivos_index)
    findings += f

    filas_prod_contratos = v.load(E, wb, "productos_contratos")
    convencion = v.detectar_convencion_join(E, filas_contratos, filas_prod_contratos)
    f = v.validar_productos_contratos(E, filas_prod_contratos, convencion, filas_contratos, contratos_index,
                                      productos_index, planes_norm, materiales_norm)
    findings += f

    reconciliado_con_bd = False
    if conn is not None:
        from core.db import cargar_catalogo_productos, cargar_catalogo_proyectos
        findings += v.reconciliar_productos_con_bd(E, productos_index, cargar_catalogo_productos(conn, E))
        findings += v.reconciliar_proyectos_con_bd(E, proyectos_index, cargar_catalogo_proyectos(conn, E))
        reconciliado_con_bd = True

    contexto = {
        "esquema": E,
        "reconciliado_con_bd": reconciliado_con_bd,
        "convencion_join_productos_contratos": convencion,
        "tipo_contrato_index": tipo_contrato_index,
        "categorias_index": categorias_index,
        "productos_index": productos_index,
        "consecutivos_index": consecutivos_index,
        "proyectos_index": proyectos_index,
        "personas_index": personas_index,
        "contratos_index": contratos_index,
        "materiales_norm": materiales_norm,
        # filas crudas, necesarias para la generacion de SQL (joins por fila)
        "_filas_contratos": filas_contratos,
        "_filas_prod_contratos": filas_prod_contratos,
        "_wb": wb,
    }
    return findings, contexto


def resumen(findings):
    total = len(findings)
    ok = sum(1 for f in findings if f.severidad == "OK")
    adv = sum(1 for f in findings if f.severidad == "ADVERTENCIA")
    err = sum(1 for f in findings if f.severidad == "ERROR")
    return {"total": total, "OK": ok, "ADVERTENCIA": adv, "ERROR": err}
