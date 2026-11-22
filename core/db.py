# -*- coding: utf-8 -*-
"""
db.py
Conexion OPCIONAL a una base destino ya poblada, usada unicamente para
el cruce de matching tolerante de PRODUCTOS y PROYECTOS (ver
matching.py). La herramienta sigue funcionando 100% sin esto - esta
conexion solo se usa si el usuario la activa explicitamente en la
interfaz, para el caso de migrar/reconciliar contra una base que ya
tiene datos. Las consultas son de solo lectura.
"""
import psycopg2

from core.matching import tokens


class DbConnectionError(Exception):
    pass


def connect(host, port, user, password, dbname):
    try:
        return psycopg2.connect(host=host, port=port, user=user, password=password, dbname=dbname)
    except Exception as e:
        raise DbConnectionError(str(e))


def _id(E, ent):
    return E.col(ent, "id") if E.tiene_col(ent, "id") else "id"


def cargar_catalogo_productos(conn, E):
    cur = conn.cursor()
    cur.execute(f"SELECT {_id(E, 'producto')}, {E.col('producto', 'nombre')}, {E.col('producto', 'precio_venta')} "
                f"FROM {E.tabla('producto')}")
    out = [{"id": r[0], "nombre": r[1], "precio": float(r[2]) if r[2] is not None else None} for r in cur.fetchall()]
    cur.close()
    return out


def cargar_catalogo_proyectos(conn, E):
    cur = conn.cursor()
    cur.execute(f"""
        SELECT p.{_id(E, 'proyecto')}, p.{E.col('proyecto', 'nombre')}, c.{E.col('ciudad', 'nombre')}
        FROM {E.tabla('proyecto')} p
        LEFT JOIN {E.tabla('ciudad')} c ON c.{_id(E, 'ciudad')} = p.{E.col('proyecto', 'ciudad')}
    """)
    out = []
    for id_, nombre, ciudad in cur.fetchall():
        out.append({"id": id_, "nombre": nombre, "extra_tokens": tokens(ciudad) if ciudad else []})
    cur.close()
    return out
