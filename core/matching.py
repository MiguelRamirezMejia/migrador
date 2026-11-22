# -*- coding: utf-8 -*-
"""
matching.py
Algoritmo de matching tolerante para resolver nombres de productos y
proyectos del Excel contra un catalogo real ya existente en una base
de datos (poblada), cuando los nombres no coinciden EXACTO por
diferencias de formato (guion bajo vs espacio, sufijo de ciudad, etc).

Resolucion en 3 pasos, nunca adivina si queda ambiguo:
  1. Match EXACTO normalizado (ignora mayus/acentos/separadores) -> si
     hay 1 solo candidato, resuelto con confianza alta.
  2. Si hay >1 candidato exacto normalizado, o 0, intenta match
     TOLERANTE por tokens (subset de tokens con tolerancia de 1
     caracter), filtrando SIEMPRE por precio si viene dado (precio
     distinto = nunca es el mismo producto, es un filtro duro).
  3. Si sigue ambiguo, se desambigua por un campo de contexto adicional
     (ej: proyecto/ciudad) si se provee. Si no se puede resolver a un
     unico candidato, se devuelve AMBIGUO (nunca se adivina).
"""
import re
from dataclasses import dataclass, field

from core.excel_loader import norm_key


def tokens(s):
    if s is None:
        return []
    return [t for t in re.split(r"[^A-Za-z0-9]+", str(s).upper()) if t]


def token_matches(t_a, t_b):
    if t_a == t_b:
        return True
    if t_a in t_b or t_b in t_a:
        return abs(len(t_a) - len(t_b)) <= 1
    return False


def es_subset_tolerante(tokens_buscado, tokens_candidato):
    usados = set()
    for tb in tokens_buscado:
        encontrado = False
        for i, tc in enumerate(tokens_candidato):
            if i in usados:
                continue
            if token_matches(tb, tc):
                usados.add(i)
                encontrado = True
                break
        if not encontrado:
            return False
    return True


ESTADO_EXACTO = "exacto"
ESTADO_TOLERANTE = "tolerante"
ESTADO_AMBIGUO = "ambiguo"
ESTADO_SIN_MATCH = "sin_match"


@dataclass
class ResultadoMatch:
    estado: str
    candidato: dict | None = None
    candidatos_ambiguos: list = field(default_factory=list)


def resolver(nombre_buscado, candidatos, precio_buscado=None, contexto_extra=None):
    """
    candidatos: lista de dicts con al menos {'id', 'nombre'}; opcionalmente
    'precio' y 'extra_tokens' (lista de tokens de contexto, ej palabras
    del nombre de proyecto/ciudad al que pertenece el candidato).
    contexto_extra: texto (ej nombre de proyecto de la fila origen) para
    desambiguar por solape de tokens con 'extra_tokens' de cada candidato.
    """
    clave_buscada = norm_key(nombre_buscado)

    exactos = [c for c in candidatos if norm_key(c["nombre"]) == clave_buscada]
    if precio_buscado is not None:
        exactos_precio = [c for c in exactos if c.get("precio") == precio_buscado]
        if exactos_precio:
            exactos = exactos_precio
    if len(exactos) == 1:
        return ResultadoMatch(ESTADO_EXACTO, exactos[0])
    if len(exactos) > 1:
        resuelto = _desambiguar_por_contexto(exactos, contexto_extra)
        if resuelto:
            return ResultadoMatch(ESTADO_EXACTO, resuelto)
        return ResultadoMatch(ESTADO_AMBIGUO, candidatos_ambiguos=exactos)

    toks_buscado = tokens(nombre_buscado)
    tolerantes = [c for c in candidatos if es_subset_tolerante(toks_buscado, tokens(c["nombre"]))]
    if precio_buscado is not None:
        con_precio = [c for c in tolerantes if c.get("precio") == precio_buscado]
        if con_precio:
            tolerantes = con_precio
        elif tolerantes:
            # Habia candidatos por nombre pero NINGUNO con el precio correcto
            # -> nunca se usa uno con precio distinto, se descarta todo.
            tolerantes = []

    if len(tolerantes) == 1:
        return ResultadoMatch(ESTADO_TOLERANTE, tolerantes[0])
    if len(tolerantes) > 1:
        resuelto = _desambiguar_por_contexto(tolerantes, contexto_extra)
        if resuelto:
            return ResultadoMatch(ESTADO_TOLERANTE, resuelto)
        return ResultadoMatch(ESTADO_AMBIGUO, candidatos_ambiguos=tolerantes)

    return ResultadoMatch(ESTADO_SIN_MATCH)


def _desambiguar_por_contexto(candidatos, contexto_extra):
    if not contexto_extra:
        return None
    toks_contexto = set(tokens(contexto_extra))
    con_overlap = [c for c in candidatos if set(c.get("extra_tokens") or tokens(c["nombre"])) & toks_contexto]
    if len(con_overlap) == 1:
        return con_overlap[0]
    return None
