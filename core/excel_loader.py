# -*- coding: utf-8 -*-
"""
excel_loader.py
Carga las hojas de la plantilla de migracion y las normaliza a listas de
diccionarios {nombre_columna: valor}, usando el nombre real del header
(no el indice de columna, porque el orden de columnas puede variar
entre archivos).
"""
import re
import openpyxl


class ExcelLoadError(Exception):
    pass


def _norm_header(h):
    if h is None:
        return None
    return re.sub(r"\s+", " ", str(h).strip()).lower()


def _is_header_row(row, marcadores):
    """Detecta filas de encabezado duplicadas incrustadas a mitad de hoja."""
    for v in row:
        if v is None:
            continue
        if _norm_header(v) in marcadores:
            return True
    return False


def load_workbook(path):
    try:
        return openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        raise ExcelLoadError(f"No se pudo abrir el archivo: {e}")


def load_sheet_as_dicts(wb, sheet_name, header_row_index=0, marcadores=frozenset()):
    """
    Devuelve (headers_originales, filas) donde filas es una lista de
    dicts {header_original: valor}, con numero de fila real de Excel
    (1-indexed) guardado en '_excel_row'.
    Filtra filas vacias y encabezados duplicados incrustados (filas que
    contienen alguno de los 'marcadores' de encabezado del esquema).
    """
    if sheet_name not in wb.sheetnames:
        raise ExcelLoadError(f"La hoja '{sheet_name}' no existe en el archivo.")
    ws = wb[sheet_name]

    rows_iter = ws.iter_rows(values_only=True)
    headers = None
    out = []
    for i, row in enumerate(rows_iter):
        if i < header_row_index:
            continue
        if headers is None:
            headers = [h for h in row]
            continue
        if all(c is None for c in row):
            continue
        if _is_header_row(row, marcadores):
            continue
        d = {}
        for col_idx, h in enumerate(headers):
            if h is None:
                continue
            d[str(h).strip()] = row[col_idx] if col_idx < len(row) else None
        d["_excel_row"] = header_row_index + i + 1
        out.append(d)
    return headers, out


def norm_str(v):
    """Normaliza un valor a texto limpio (trim), o None si esta vacio."""
    if v is None:
        return None
    s = str(v).strip()
    return s if s != "" else None


def norm_key(v):
    """Normaliza un texto para comparaciones tolerantes (mayus/tildes/espacios)."""
    if v is None:
        return None
    s = str(v).strip().upper()
    s = re.sub(r"\s+", " ", s)
    trans = str.maketrans("ÁÉÍÓÚÑáéíóúñ", "AEIOUNaeioun")
    return s.translate(trans)
