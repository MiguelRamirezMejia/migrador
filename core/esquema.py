# -*- coding: utf-8 -*-
"""
esquema.py
Carga el esquema de migracion: la descripcion de la plantilla Excel
(hojas, fila de encabezado y nombres de columna aceptados) y de la base
de datos destino (tablas, columnas, valores fijos y columnas extra).

El motor (validators, sql_generator, db) no conoce ningun nombre de
tabla, columna u hoja: todo sale de aqui. Asi el mismo motor sirve para
cualquier sistema destino cambiando solo el archivo YAML.

Orden de busqueda del archivo de esquema:
  1. Variable de entorno MIGRADOR_ESQUEMA (ruta a un .yaml).
  2. privado/esquema.yaml (configuracion propia, fuera de git).
  3. config/esquema.ejemplo.yaml (esquema de ejemplo incluido en el repo).
"""
import os
import re
from pathlib import Path

import yaml

RAIZ = Path(__file__).resolve().parent.parent
ESQUEMA_PRIVADO = RAIZ / "privado" / "esquema.yaml"
ESQUEMA_EJEMPLO = RAIZ / "config" / "esquema.ejemplo.yaml"

# Los nombres de tabla y columna se interpolan en SQL, asi que solo se
# aceptan identificadores simples (opcionalmente con esquema: "ventas.clientes").
_IDENT = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*(\.[A-Za-z_][A-Za-z0-9_]*)?$")

TIPOS_VALOR = {"texto", "numero", "booleano"}


class EsquemaError(Exception):
    pass


def _ident(valor, donde):
    if not isinstance(valor, str) or not _IDENT.match(valor):
        raise EsquemaError(f"Identificador SQL no valido en {donde}: {valor!r}")
    return valor


class Esquema:
    def __init__(self, data, origen):
        if not isinstance(data, dict):
            raise EsquemaError("El esquema debe ser un diccionario YAML.")
        self.origen = Path(origen)
        self.nombre = data.get("nombre", self.origen.stem)
        self.etiquetas = data.get("etiquetas", {})
        self._plantilla = data.get("plantilla", {})
        self._tablas = data.get("tablas", {})
        self.reglas = data.get("reglas", {})
        self._validar()

    # ------------------------------------------------------------- plantilla
    def _hoja_cfg(self, clave):
        hojas = self._plantilla.get("hojas", {})
        if clave not in hojas:
            raise EsquemaError(f"El esquema no define la hoja '{clave}' (plantilla.hojas.{clave}).")
        return hojas[clave]

    def hoja(self, clave):
        """Nombre real de la hoja en el Excel."""
        return self._hoja_cfg(clave)["nombre"]

    def fila_encabezado(self, clave):
        """Indice 0-based de la fila de encabezado (el YAML usa 1-based, como Excel)."""
        return int(self._hoja_cfg(clave).get("fila_encabezado", 1)) - 1

    def cols(self, clave, campo):
        """Nombres de columna aceptados en el Excel para un campo logico."""
        columnas = self._hoja_cfg(clave).get("columnas", {})
        alias = columnas.get(campo, campo)
        return tuple(alias) if isinstance(alias, list) else (alias,)

    @property
    def marcadores_encabezado(self):
        return {str(m).strip().lower() for m in self._plantilla.get("marcadores_encabezado", [])}

    # ------------------------------------------------------------- tablas
    def _tabla_cfg(self, ent):
        if ent not in self._tablas:
            raise EsquemaError(f"El esquema no define la tabla '{ent}' (tablas.{ent}).")
        return self._tablas[ent]

    def tiene_tabla(self, ent):
        return ent in self._tablas

    def tabla(self, ent):
        return self._tabla_cfg(ent)["tabla"]

    def col(self, ent, campo):
        cols = self._tabla_cfg(ent).get("columnas", {})
        if campo not in cols:
            raise EsquemaError(f"El esquema no define la columna '{campo}' de la tabla '{ent}'.")
        return cols[campo]

    def tiene_col(self, ent, campo):
        return campo in self._tabla_cfg(ent).get("columnas", {})

    def extras(self, ent):
        """Columnas que pasan del Excel a la tabla sin reglas propias."""
        return self._tabla_cfg(ent).get("extras", [])

    def fijos(self, ent):
        """{columna: literal SQL} que se insertan siempre con el mismo valor."""
        return self._tabla_cfg(ent).get("fijos", {})

    def opcion(self, ent, clave, defecto=None):
        return self._tabla_cfg(ent).get(clave, defecto)

    def etiqueta(self, clave, defecto):
        return self.etiquetas.get(clave, defecto)

    # ------------------------------------------------------------- validacion
    def _validar(self):
        for ent, cfg in self._tablas.items():
            _ident(cfg.get("tabla"), f"tablas.{ent}.tabla")
            for campo, col in cfg.get("columnas", {}).items():
                _ident(col, f"tablas.{ent}.columnas.{campo}")
            for col in cfg.get("fijos", {}):
                _ident(col, f"tablas.{ent}.fijos")
            for i, extra in enumerate(cfg.get("extras", [])):
                _ident(extra.get("columna"), f"tablas.{ent}.extras[{i}].columna")
                if extra.get("tipo", "texto") not in TIPOS_VALOR | {"busqueda"}:
                    raise EsquemaError(f"Tipo no valido en tablas.{ent}.extras[{i}]: {extra.get('tipo')!r}")
                if extra.get("tipo") == "busqueda":
                    _ident(extra.get("tabla_busqueda"), f"tablas.{ent}.extras[{i}].tabla_busqueda")
                    _ident(extra.get("campo_busqueda", "nombre"), f"tablas.{ent}.extras[{i}].campo_busqueda")
        for hoja, cfg in self._plantilla.get("hojas", {}).items():
            if not cfg.get("nombre"):
                raise EsquemaError(f"Falta plantilla.hojas.{hoja}.nombre")


def ruta_esquema():
    env = os.environ.get("MIGRADOR_ESQUEMA")
    if env:
        return Path(env)
    if ESQUEMA_PRIVADO.exists():
        return ESQUEMA_PRIVADO
    return ESQUEMA_EJEMPLO


def cargar(ruta=None):
    ruta = Path(ruta) if ruta else ruta_esquema()
    if not ruta.exists():
        raise EsquemaError(f"No existe el archivo de esquema: {ruta}")
    with open(ruta, encoding="utf-8") as f:
        return Esquema(yaml.safe_load(f), ruta)


def ruta_relativa(ruta):
    """Resuelve una ruta del esquema relativa a la raiz del proyecto."""
    p = Path(ruta)
    return p if p.is_absolute() else RAIZ / p
