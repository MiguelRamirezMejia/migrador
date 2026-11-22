# -*- coding: utf-8 -*-
"""
catalogs.py
Catalogos fijos que no dependen de ninguna base de datos. Hoy: el
catalogo de ciudades, que es identico en la base destino y se empaqueta
como CSV (columnas 'id' y 'nombre') en vez de requerir una conexion
solo para esto. La ruta del CSV la define el esquema
(tablas.ciudad.catalogo_csv).
"""
import csv

from core.esquema import ruta_relativa
from core.excel_loader import norm_key


def cargar_ciudades(esquema=None, csv_path=None):
    """Devuelve dict {nombre_normalizado: (id, nombre_original)}."""
    if csv_path is None:
        csv_path = esquema.opcion("ciudad", "catalogo_csv")
    path = ruta_relativa(csv_path)
    out = {}
    with open(path, encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            nombre = row.get("nombre")
            id_ = row.get("id")
            if nombre:
                out[norm_key(nombre)] = (id_, nombre)
    return out


def resolver_ciudad(ciudades_idx, nombre_excel):
    """Devuelve (id, nombre_db) o (None, None) si no matchea exacto (tolerante a tildes/mayus)."""
    if not nombre_excel:
        return None, None
    clave = norm_key(nombre_excel)
    if clave in ciudades_idx:
        return ciudades_idx[clave]
    return None, None
