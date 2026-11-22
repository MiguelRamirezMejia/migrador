# -*- coding: utf-8 -*-
from dataclasses import dataclass

OK = "OK"
ADVERTENCIA = "ADVERTENCIA"
ERROR = "ERROR"


@dataclass
class Finding:
    hoja: str
    fila: int
    entidad: str
    severidad: str
    mensaje: str
    referencia: str = ""
