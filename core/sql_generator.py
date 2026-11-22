# -*- coding: utf-8 -*-
"""
sql_generator.py
Genera el SQL de migracion a partir de la plantilla ya validada.

Estrategia de IDs: en vez de asignar IDs numericos nosotros (lo que
exigiria conocer el MAX(id) real de la base destino), cada FK se
resuelve con una SUBCONSULTA por llave de negocio (nombre, documento,
codigo de servicio, etc.) contra la base en el momento en que el script
se ejecuta. Esto hace que el .sql generado sea autocontenido y no
dependa de una conexion en el momento de generarlo - solo al correrlo.

Excepcion: el catalogo de ciudades es fijo e identico en la base
destino (se empaqueta como CSV), asi que para esa FK se usa el ID real
ya conocido, no una subconsulta.

Todos los nombres de tablas y columnas salen del esquema.
"""
import re

from core.excel_loader import norm_str, norm_key
from core.validators import campo, mapa_contratos, codigos_destino


# ---------------------------------------------------------------------
# Helpers de formato SQL
# ---------------------------------------------------------------------
def esc(s):
    return str(s).replace("'", "''")


def lit_str(v):
    if v is None:
        return "NULL"
    s = str(v).strip()
    if s == "":
        return "NULL"
    return f"'{esc(s)}'"


def lit_num(v):
    if v is None or str(v).strip() == "":
        return "NULL"
    return str(v).strip()


def lit_bool(v):
    if v is None:
        return "NULL"
    s = str(v).strip().lower()
    if s in ("true", "1", "si", "sí", "x", "verdadero"):
        return "TRUE"
    if s in ("false", "0", "no", "falso"):
        return "FALSE"
    return "NULL"


LITERALES = {"texto": lit_str, "numero": lit_num, "booleano": lit_bool}


def comentario(texto):
    return f"-- {texto}"


def idcol(E, ent):
    """Columna de llave primaria de una tabla (por defecto 'id')."""
    return E.col(ent, "id") if E.tiene_col(ent, "id") else "id"


def sub_id(E, ent, campo_logico, valor):
    """Subconsulta por llave de negocio: (SELECT id FROM tabla WHERE campo = 'valor')."""
    if valor is None or str(valor).strip() == "":
        return "NULL"
    return f"(SELECT {idcol(E, ent)} FROM {E.tabla(ent)} WHERE {E.col(ent, campo_logico)} = {lit_str(valor)})"


def valores_extras(extras_cfg, valores):
    """[(columna, literal_sql)] para las columnas extra de una tabla."""
    pares = []
    for extra in extras_cfg:
        col = extra["columna"]
        v = valores.get(col)
        tipo = extra.get("tipo", "texto")
        if tipo == "busqueda":
            if v is None or str(v).strip() == "":
                sql = "NULL"
            else:
                sql = f"(SELECT id FROM {extra['tabla_busqueda']} WHERE {extra.get('campo_busqueda', 'nombre')} = {lit_str(v)})"
        else:
            sql = LITERALES[tipo](v)
        pares.append((col, sql))
    return pares


def insert_select(tabla, pares, resto):
    """INSERT INTO tabla (cols) SELECT valores <resto>;"""
    cols = ", ".join(c for c, _ in pares)
    vals = ", ".join(v for _, v in pares)
    return f"INSERT INTO {tabla} ({cols})\nSELECT {vals}\n{resto};"


# ---------------------------------------------------------------------
# 1. Tipos de contrato
# ---------------------------------------------------------------------
def tipo_contrato_sql(E, tipo_contrato_index):
    ent = "tipo_contrato"
    T = E.tabla(ent)
    out = []
    for info in tipo_contrato_index.values():
        pares = [(E.col(ent, "nombre"), lit_str(info["nombre"])),
                 (E.col(ent, "identificador"), lit_str(info["identificador"]))]
        pares += valores_extras(E.extras(ent), info["extras"])
        pares += list(E.fijos(ent).items())
        out.append(comentario(f"TIPO CONTRATO: {info['nombre']}"))
        out.append(insert_select(T, pares,
                                 f"WHERE NOT EXISTS (SELECT 1 FROM {T} WHERE {E.col(ent, 'identificador')} = {lit_str(info['identificador'])})"))
    return out


# ---------------------------------------------------------------------
# 2. Categorias de productos
# ---------------------------------------------------------------------
def categoria_sql(E, categorias_index):
    ent = "categoria"
    T = E.tabla(ent)
    out = []
    for info in categorias_index.values():
        pares = [(E.col(ent, "nombre"), lit_str(info["nombre"])),
                 (E.col(ent, "descripcion"), lit_str(info.get("descripcion"))),
                 (E.col(ent, "activo"), f"COALESCE({lit_bool(info.get('activo'))}, TRUE)")]
        pares += list(E.fijos(ent).items())
        out.append(comentario(f"CATEGORIA PRODUCTO: {info['nombre']}"))
        out.append(insert_select(T, pares,
                                 f"WHERE NOT EXISTS (SELECT 1 FROM {T} WHERE {E.col(ent, 'nombre')} = {lit_str(info['nombre'])})"))
    return out


# ---------------------------------------------------------------------
# 3. Productos
# ---------------------------------------------------------------------
def producto_sql(E, productos_index):
    ent = "producto"
    T = E.tabla(ent)
    out = []
    for info in productos_index.values():
        if info.get("id_real"):
            out.append(comentario(f"PRODUCTO: {info['nombre']} -> ya existe en BD real como "
                                  f"'{info['nombre_real']}' (id={info['id_real']}), se reutiliza, no se crea duplicado."))
            continue
        pares = [
            (E.col(ent, "nombre"), lit_str(info["nombre"])),
            (E.col(ent, "referencia"), lit_str(info["referencia"])),
            (E.col(ent, "activo"), f"COALESCE({lit_bool(info.get('activo'))}, TRUE)"),
            (E.col(ent, "tiene_seriales"), f"COALESCE({lit_bool(info.get('tiene_seriales'))}, FALSE)"),
            (E.col(ent, "se_puede_vender"), lit_bool(info.get("se_puede_vender"))),
            (E.col(ent, "se_puede_comprar"), lit_bool(info.get("se_puede_comprar"))),
            (E.col(ent, "precio_venta"), lit_num(info.get("precio_venta"))),
            (E.col(ent, "categoria"), sub_id(E, "categoria", "nombre", info.get("categoria_producto"))),
        ]
        pares += valores_extras(E.extras(ent), info["extras"])
        if E.tiene_col(ent, "es_principal"):
            pares.append((E.col(ent, "es_principal"), lit_bool(info.get("es_principal"))))
        pares += list(E.fijos(ent).items())
        out.append(comentario(f"PRODUCTO: {info['nombre']}"))
        out.append(insert_select(T, pares,
                                 f"WHERE NOT EXISTS (SELECT 1 FROM {T} WHERE {E.col(ent, 'nombre')} = {lit_str(info['nombre'])})"))
    return out


# ---------------------------------------------------------------------
# 4. Consecutivos + proyectos
# ---------------------------------------------------------------------
def consecutivo_y_proyecto_sql(E, consecutivos_index, proyectos_index):
    T_CONS, T_PROY = E.tabla("consecutivo"), E.tabla("proyecto")
    out = []
    for info in consecutivos_index.values():
        ent = "consecutivo"
        pares = [
            (E.col(ent, "nombre"), lit_str(info["nombre"])),
            (E.col(ent, "prefijo"), lit_str(info["prefijo"])),
            (E.col(ent, "sufijo"), lit_str(info.get("sufijo"))),
            (E.col(ent, "numero_inicial"), lit_num(info.get("numero_inicial"))),
            (E.col(ent, "numero_actual"), lit_num(info.get("numero_actual"))),
            (E.col(ent, "numero_de_digitos"), lit_num(info.get("numero_de_digitos"))),
            (E.col(ent, "activo"), f"COALESCE({lit_bool(info.get('activo'))}, TRUE)"),
            (E.col(ent, "tipo_documento"), lit_num(info.get("tipo_documento"))),
        ]
        pares += valores_extras(E.extras(ent), info["extras"])
        pares += list(E.fijos(ent).items())
        out.append(comentario(f"CONSECUTIVO: {info['nombre']}"))
        out.append(insert_select(T_CONS, pares,
                                 f"WHERE NOT EXISTS (SELECT 1 FROM {T_CONS} WHERE {E.col(ent, 'nombre')} = {lit_str(info['nombre'])})"))

    for info in proyectos_index.values():
        if info.get("id_real"):
            out.append(comentario(f"PROYECTO: {info['nombre']} -> ya existe en BD real como "
                                  f"'{info['nombre_real']}' (id={info['id_real']}), se reutiliza, no se crea duplicado."))
            continue
        ent = "proyecto"
        consecutivo_nombre = consecutivos_index.get(info["consecutivo_nombre_norm"], {}).get("nombre")
        pares = [
            (E.col(ent, "nombre"), lit_str(info["nombre"])),
            (E.col(ent, "ciudad"), lit_num(info.get("ciudad_id"))),
            (E.col(ent, "consecutivo"), sub_id(E, "consecutivo", "nombre", consecutivo_nombre)),
        ]
        pares += list(E.fijos(ent).items())
        out.append(comentario(f"PROYECTO: {info['nombre']}"))
        out.append(insert_select(T_PROY, pares,
                                 f"WHERE NOT EXISTS (SELECT 1 FROM {T_PROY} WHERE {E.col(ent, 'nombre')} = {lit_str(info['nombre'])})"))
    return out


# ---------------------------------------------------------------------
# 5. Personas -> persona + cliente + empleado (+ almacen automatico)
# ---------------------------------------------------------------------
def _slug_almacen(nombres, apellidos):
    base = f"{nombres or ''} {apellidos or ''}".strip().upper()
    base = re.sub(r"[^A-Z0-9]+", "_", base).strip("_")
    return base or "EMPLEADO"


def _desde_persona(E, doc, alias_dest, tabla_dest, col_persona_dest):
    """FROM persona p WHERE p.documento = doc AND NOT EXISTS (fila destino de esa persona)."""
    P_ID, P_DOC = idcol(E, "persona"), E.col("persona", "numero_documento")
    return (f"FROM {E.tabla('persona')} p WHERE p.{P_DOC} = {lit_str(doc)}\n"
            f"AND NOT EXISTS (SELECT 1 FROM {tabla_dest} {alias_dest} WHERE {alias_dest}.{col_persona_dest} = p.{P_ID})")


def personas_sql(E, personas_index):
    T_PER = E.tabla("persona")
    P_ID = idcol(E, "persona")
    out = []
    for doc, info in personas_index.items():
        ent = "persona"
        pares = [
            (E.col(ent, "tipo_identificacion"), lit_num(info["tipo_identificacion"])),
            (E.col(ent, "numero_documento"), lit_str(doc)),
            (E.col(ent, "nombres"), lit_str(info["nombres"])),
            (E.col(ent, "apellidos"), lit_str(info.get("apellidos"))),
            (E.col(ent, "email"), lit_str(info.get("email"))),
        ]
        pares += list(E.fijos(ent).items())
        out.append(comentario(f"PERSONA: {doc} - {info['nombres']} {info.get('apellidos') or ''}"))
        out.append(insert_select(T_PER, pares,
                                 f"WHERE NOT EXISTS (SELECT 1 FROM {T_PER} WHERE {E.col(ent, 'numero_documento')} = {lit_str(doc)})"))

        if info["es_cliente"]:
            ent = "cliente"
            T_CLI = E.tabla(ent)
            pares = []
            if E.tiene_col(ent, "id"):  # el cliente comparte el id de la persona
                pares.append((E.col(ent, "id"), f"p.{P_ID}"))
            pares.append((E.col(ent, "persona"), f"p.{P_ID}"))
            pares += list(E.fijos(ent).items())
            out.append(comentario(f"  -> CLIENTE para {doc}"))
            out.append(insert_select(T_CLI, pares, _desde_persona(E, doc, "fc", T_CLI, E.col(ent, "persona"))))

        if info["es_empleado"]:
            if info["almacen_generado"]:
                ent = "almacen"
                T_ALM = E.tabla(ent)
                sufijo = E.opcion(ent, "sufijo", "_01")
                nombre_almacen = _slug_almacen(info["nombres"], info.get("apellidos")) + sufijo
                pares = [(E.col(ent, "nombre"), lit_str(nombre_almacen)),
                         (E.col(ent, "nombre_corto"), lit_str(nombre_almacen[:3] + sufijo))]
                pares += list(E.fijos(ent).items())
                out.append(comentario(f"  -> Almacen automatico para empleado {doc}"))
                out.append(insert_select(T_ALM, pares,
                                         f"WHERE NOT EXISTS (SELECT 1 FROM {T_ALM} WHERE {E.col(ent, 'nombre')} = {lit_str(nombre_almacen)})"))
                if E.tiene_tabla("ubicacion"):
                    U, A_ID, A_NOM = "ubicacion", idcol(E, "almacen"), E.col("almacen", "nombre")
                    T_UBI = E.tabla(U)
                    for ubic in E.opcion(U, "filas", []):
                        pares = [(E.col(U, "almacen"), f"a.{A_ID}"),
                                 (E.col(U, "nombre"), lit_str(ubic["nombre"])),
                                 (E.col(U, "tipo"), lit_num(ubic.get("tipo"))),
                                 (E.col(U, "disponible_venta"), lit_bool(ubic.get("disponible_venta")))]
                        pares += list(E.fijos(U).items())
                        out.append(insert_select(T_UBI, pares,
                                                 f"FROM {T_ALM} a WHERE a.{A_NOM} = {lit_str(nombre_almacen)}\n"
                                                 f"AND NOT EXISTS (SELECT 1 FROM {T_UBI} u WHERE u.{E.col(U, 'almacen')} = a.{A_ID} "
                                                 f"AND u.{E.col(U, 'nombre')} = {lit_str(ubic['nombre'])})"))
                almacen_sub = sub_id(E, "almacen", "nombre", nombre_almacen)
            else:
                almacen_sub = "NULL"

            ent = "empleado"
            T_EMP = E.tabla(ent)
            pares = []
            if E.tiene_col(ent, "id"):
                pares.append((E.col(ent, "id"), f"p.{P_ID}"))
            pares += [(E.col(ent, "persona"), f"p.{P_ID}"),
                      (E.col(ent, "activo"), "TRUE"),
                      (E.col(ent, "almacen"), almacen_sub),
                      (E.col(ent, "almacen_generado"), lit_bool(info["almacen_generado"]))]
            pares += list(E.fijos(ent).items())
            out.append(comentario(f"  -> EMPLEADO para {doc}"))
            out.append(insert_select(T_EMP, pares, _desde_persona(E, doc, "ge", T_EMP, E.col(ent, "persona"))))
    return out


# ---------------------------------------------------------------------
# 6. Contratos (+ informacion de instalacion)
# ---------------------------------------------------------------------
def _resolver_consecutivo_contrato(info, consecutivos_index, proyectos_index, tipo_contrato_index, contadores):
    """
    Si la fila ya trae el consecutivo de contrato literal, se usa tal cual.
    Si no, se genera con la formula: identificador del tipo de contrato +
    prefijo + numero incremental (rellenado con ceros) + sufijo, llevando
    un contador en memoria por consecutivo, iniciado desde numero_actual.
    """
    if info.get("consecutivo_contrato"):
        return lit_str(info["consecutivo_contrato"]), None

    proyecto_norm = norm_key(info["proyecto"]) if info.get("proyecto") else None
    proy = proyectos_index.get(proyecto_norm) if proyecto_norm else None
    if not proy:
        return "NULL", "no se pudo generar el consecutivo de contrato: proyecto no resuelto"

    consec = consecutivos_index.get(proy["consecutivo_nombre_norm"])
    tipo = tipo_contrato_index.get(norm_key(info["tipo_contrato"])) if info.get("tipo_contrato") else None
    if not consec or not tipo:
        return "NULL", "no se pudo generar el consecutivo de contrato: falta consecutivo o tipo de contrato"

    clave = proy["consecutivo_nombre_norm"]
    if clave not in contadores:
        contadores[clave] = int(consec.get("numero_actual") or consec.get("numero_inicial") or 1)
    numero = contadores[clave]
    contadores[clave] += 1

    digitos = int(consec.get("numero_de_digitos") or 0)
    numero_str = str(numero).zfill(digitos) if digitos > 0 else str(numero)
    valor = f"{tipo['identificador']}{consec['prefijo']}{numero_str}{consec.get('sufijo') or ''}"
    return lit_str(valor), None


def contratos_sql(E, contratos_index, consecutivos_index, proyectos_index, tipo_contrato_index):
    ent = "contrato"
    T = E.tabla(ent)
    C_ID, C_COD = idcol(E, ent), E.col(ent, "codigo_servicio")
    COD = E.etiqueta("codigo_servicio", "Codigo de servicio")
    out = []
    contadores = {}
    notas = []

    for codigo, info in contratos_index.items():
        consec_sql, nota = _resolver_consecutivo_contrato(info, consecutivos_index, proyectos_index, tipo_contrato_index, contadores)
        if nota:
            notas.append(f"{codigo}: {nota}")

        cliente_sub = (f"(SELECT c.{idcol(E, 'cliente')} FROM {E.tabla('cliente')} c "
                       f"JOIN {E.tabla('persona')} p ON p.{idcol(E, 'persona')} = c.{E.col('cliente', 'persona')} "
                       f"WHERE p.{E.col('persona', 'numero_documento')} = {lit_str(info['cliente_doc'])})")
        proy_info = proyectos_index.get(norm_key(info["proyecto"])) if info.get("proyecto") else None
        nombre_proyecto_sql = (proy_info.get("nombre_real") or proy_info["nombre"]) if proy_info else info.get("proyecto")

        pares = [
            (E.col(ent, "cliente"), cliente_sub),
            (E.col(ent, "proyecto"), sub_id(E, "proyecto", "nombre", nombre_proyecto_sql)),
            (E.col(ent, "tipo_contrato"), sub_id(E, "tipo_contrato", "nombre", info.get("tipo_contrato"))),
            (C_COD, lit_str(codigo)),
            (E.col(ent, "consecutivo"), consec_sql),
            (E.col(ent, "factura_electronica"), lit_bool(info.get("factura_electronica"))),
            (E.col(ent, "latitud"), lit_num(info.get("latitud"))),
            (E.col(ent, "longitud"), lit_num(info.get("longitud"))),
        ]
        pares += list(E.fijos(ent).items())
        out.append(comentario(f"CONTRATO {COD}={codigo} cliente={info['cliente_doc']}"))
        out.append(insert_select(T, pares, f"WHERE NOT EXISTS (SELECT 1 FROM {T} WHERE {C_COD} = {lit_str(codigo)})"))

        instalacion = info.get("instalacion") or {}
        if E.tiene_tabla("instalacion") and any(instalacion.values()):
            I = "instalacion"
            T_INS, I_CONT = E.tabla(I), E.col(I, "contrato")
            pares = [(I_CONT, f"c.{C_ID}")]
            pares += valores_extras(E.extras(I), instalacion)
            pares += list(E.fijos(I).items())
            out.append(comentario(f"  -> Informacion de instalacion para {codigo}"))
            out.append(insert_select(T_INS, pares,
                                     f"FROM {T} c WHERE c.{C_COD} = {lit_str(codigo)}\n"
                                     f"AND NOT EXISTS (SELECT 1 FROM {T_INS} i WHERE i.{I_CONT} = c.{C_ID})"))
    return out, notas


# ---------------------------------------------------------------------
# 7. Productos de cada contrato (detalle)
# ---------------------------------------------------------------------
def detalle_contratos_sql(E, filas, convencion, filas_contratos, materiales_norm, productos_index=None):
    productos_index = productos_index or {}
    ent = "detalle_contrato"
    T = E.tabla(ent)
    T_CON, C_ID, C_COD = E.tabla("contrato"), idcol(E, "contrato"), E.col("contrato", "codigo_servicio")
    D_CONT, D_PROD = E.col(ent, "contrato"), E.col(ent, "producto")
    COD = E.etiqueta("codigo_servicio", "Codigo de servicio")
    mapa = mapa_contratos(E, convencion, filas_contratos)
    out = []

    for row in filas:
        producto = norm_str(campo(E, row, "productos_contratos", "producto"))
        cantidad = campo(E, row, "productos_contratos", "cantidad")
        precio = campo(E, row, "productos_contratos", "precio_venta")
        if not producto or cantidad is None or precio is None:
            continue

        cargo_unico = norm_key(producto) in materiales_norm
        prod_info = productos_index.get(norm_key(producto))
        nombre_producto_sql = (prod_info.get("nombre_real") or prod_info["nombre"]) if prod_info else producto
        prod_sub = sub_id(E, "producto", "nombre", nombre_producto_sql)

        for codigo in codigos_destino(E, row, convencion, mapa):
            if not codigo:
                continue
            pares = [(D_CONT, f"c.{C_ID}"), (D_PROD, prod_sub),
                     (E.col(ent, "precio_venta"), lit_num(precio)),
                     (E.col(ent, "cantidad"), lit_num(cantidad))]
            if E.tiene_col(ent, "cargo_unico"):
                pares.append((E.col(ent, "cargo_unico"), lit_bool(cargo_unico)))
            pares += list(E.fijos(ent).items())
            out.append(comentario(f"DETALLE CONTRATO {COD}={codigo}: {producto}"))
            out.append(insert_select(T, pares,
                                     f"FROM {T_CON} c WHERE c.{C_COD} = {lit_str(codigo)}\n"
                                     f"AND NOT EXISTS (\n"
                                     f"  SELECT 1 FROM {T} d\n"
                                     f"  WHERE d.{D_CONT} = c.{C_ID} AND d.{D_PROD} = {prod_sub}\n"
                                     f")"))
    return out


# ---------------------------------------------------------------------
# Orquestador
# ---------------------------------------------------------------------
def generar_sql_completo(contexto):
    E = contexto["esquema"]

    partes = [
        "-- =====================================================================",
        f"-- SQL de migracion generado por Migrador (esquema: {E.nombre})",
        "-- Resuelve FKs por subconsulta de llave de negocio (nombre/documento/codigo),",
        "-- por eso es seguro correrlo contra una base que ya tenga datos parciales.",
        "-- =====================================================================",
        "",
        "BEGIN;",
        "",
    ]

    def seccion(titulo, lineas):
        partes.append(f"-- ===== {titulo} =====")
        partes.extend(lineas)
        partes.append("")

    seccion(E.hoja("tipos_contrato"), tipo_contrato_sql(E, contexto["tipo_contrato_index"]))
    seccion(E.hoja("categorias"), categoria_sql(E, contexto["categorias_index"]))
    seccion(E.hoja("productos"), producto_sql(E, contexto["productos_index"]))
    seccion(E.hoja("consecutivos_proyectos"),
            consecutivo_y_proyecto_sql(E, contexto["consecutivos_index"], contexto["proyectos_index"]))
    seccion(E.hoja("personas"), personas_sql(E, contexto["personas_index"]))
    sql_contratos, notas = contratos_sql(E, contexto["contratos_index"], contexto["consecutivos_index"],
                                         contexto["proyectos_index"], contexto["tipo_contrato_index"])
    seccion(E.hoja("contratos"), sql_contratos)
    seccion(E.hoja("productos_contratos"),
            detalle_contratos_sql(E, contexto["_filas_prod_contratos"], contexto["convencion_join_productos_contratos"],
                                  contexto["_filas_contratos"], materiales_norm=contexto.get("materiales_norm", set()),
                                  productos_index=contexto["productos_index"]))

    partes.append("COMMIT;")

    if notas:
        partes.append("")
        partes.append("-- ===== ADVERTENCIAS DE GENERACION (revisar manualmente) =====")
        for n in notas:
            partes.append(f"-- {n}")

    return "\n".join(partes)
