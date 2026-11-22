# -*- coding: utf-8 -*-
"""
validators.py
Motor de validacion de la plantilla de migracion. Cada funcion valida
una hoja y devuelve (findings, indice_resuelto) para que las hojas
dependientes puedan cruzar informacion (ej: CONTRATOS necesita saber
que proyectos y tipos de contrato ya son validos).

Ningun nombre de hoja, columna o tabla esta escrito aqui: todos salen
del esquema (core/esquema.py). Las claves logicas que usa el motor son:
  hojas:  tipos_contrato, categorias, productos, productos_por_plan,
          consecutivos_proyectos, personas, contratos, productos_contratos
  tablas: tipo_contrato, categoria, producto, consecutivo, proyecto,
          ciudad, persona, cliente, empleado, almacen, ubicacion,
          contrato, instalacion, detalle_contrato
"""
from collections import defaultdict

from core.excel_loader import load_sheet_as_dicts, norm_str, norm_key
from core.finding import Finding, OK, ADVERTENCIA, ERROR
from core.matching import resolver, ESTADO_EXACTO, ESTADO_TOLERANTE, ESTADO_AMBIGUO

# Convenciones posibles para relacionar PRODUCTOS CONTRATOS con CONTRATOS.
CONV_CODIGO = "codigo_servicio"
CONV_DOCUMENTO = "documento_cliente"
CONV_CONSECUTIVO = "consecutivo_contrato"


def get(row, *names):
    """Busca un valor en la fila probando varios nombres de columna posibles."""
    for n in names:
        if n in row:
            v = row[n]
            if v is not None:
                return v
    return None


def campo(E, row, hoja, nombre):
    """Valor de un campo logico de la hoja, probando todos sus alias del esquema."""
    return get(row, *E.cols(hoja, nombre))


def etiqueta_col(E, hoja, nombre):
    """Nombre de columna a mostrar en los mensajes (el primer alias)."""
    return E.cols(hoja, nombre)[0]


def load(E, wb, hoja):
    _, filas = load_sheet_as_dicts(wb, E.hoja(hoja), header_row_index=E.fila_encabezado(hoja),
                                   marcadores=E.marcadores_encabezado)
    return filas


def leer_extras(row, extras):
    """{columna_bd: valor} para las columnas extra que pasan directo del Excel."""
    out = {}
    for extra in extras:
        alias = extra["excel"]
        alias = tuple(alias) if isinstance(alias, list) else (alias,)
        v = get(row, *alias)
        out[extra["columna"]] = norm_str(v) if extra.get("tipo", "texto") in ("texto", "busqueda") else v
    return out


# ---------------------------------------------------------------------
# TIPOS DE CONTRATO
# ---------------------------------------------------------------------
def validar_tipo_contrato(E, filas):
    H, T = E.hoja("tipos_contrato"), E.tabla("tipo_contrato")
    findings = []
    index = {}
    vistos_identificador = {}
    for row in filas:
        fila = row["_excel_row"]
        nombre = norm_str(campo(E, row, "tipos_contrato", "nombre"))
        identificador = norm_str(campo(E, row, "tipos_contrato", "identificador"))
        if not nombre:
            findings.append(Finding(H, fila, T, ERROR, f"Falta '{etiqueta_col(E, 'tipos_contrato', 'nombre')}' (obligatorio)."))
            continue
        if not identificador:
            findings.append(Finding(H, fila, T, ERROR, f"Falta '{etiqueta_col(E, 'tipos_contrato', 'identificador')}' (obligatorio).", nombre))
            continue
        clave = norm_key(nombre)
        if identificador in vistos_identificador:
            findings.append(Finding(H, fila, T, ADVERTENCIA,
                                    f"El identificador '{identificador}' ya fue usado por otro tipo de contrato ('{vistos_identificador[identificador]}').", nombre))
        vistos_identificador[identificador] = nombre
        index[clave] = {
            "nombre": nombre, "identificador": identificador,
            "extras": leer_extras(row, E.extras("tipo_contrato")),
        }
        findings.append(Finding(H, fila, T, OK, "OK", nombre))
    return findings, index


# ---------------------------------------------------------------------
# CATEGORIAS DE PRODUCTOS
# ---------------------------------------------------------------------
def validar_categorias_productos(E, filas):
    H, T = E.hoja("categorias"), E.tabla("categoria")
    findings = []
    index = {}
    for row in filas:
        fila = row["_excel_row"]
        nombre = norm_str(campo(E, row, "categorias", "nombre"))
        if not nombre:
            findings.append(Finding(H, fila, T, ERROR, f"Falta '{etiqueta_col(E, 'categorias', 'nombre')}' (obligatorio)."))
            continue
        index[norm_key(nombre)] = {"nombre": nombre, "descripcion": campo(E, row, "categorias", "descripcion"),
                                   "activo": campo(E, row, "categorias", "activo")}
        findings.append(Finding(H, fila, T, OK, "OK", nombre))
    return findings, index


# ---------------------------------------------------------------------
# PRODUCTOS POR PLAN  (relacion producto principal <-> complemento)
# ---------------------------------------------------------------------
def validar_productos_por_plan(E, filas, productos_nombres_norm):
    H, H_PROD = E.hoja("productos_por_plan"), E.hoja("productos")
    col_p, col_c = etiqueta_col(E, "productos_por_plan", "principal"), etiqueta_col(E, "productos_por_plan", "complemento")
    findings = []
    planes = set()
    materiales = set()
    for row in filas:
        fila = row["_excel_row"]
        plan = norm_str(campo(E, row, "productos_por_plan", "principal"))
        material = norm_str(campo(E, row, "productos_por_plan", "complemento"))
        if not plan or not material:
            findings.append(Finding(H, fila, "productos_por_plan", ERROR, f"Faltan '{col_p}' y/o '{col_c}' (ambos obligatorios)."))
            continue
        plan_k, mat_k = norm_key(plan), norm_key(material)
        if plan_k not in productos_nombres_norm:
            findings.append(Finding(H, fila, "productos_por_plan", ERROR,
                                    f"El producto principal '{plan}' no existe (nombre exacto) en la hoja {H_PROD}."))
        if mat_k not in productos_nombres_norm:
            findings.append(Finding(H, fila, "productos_por_plan", ERROR,
                                    f"El complemento '{material}' no existe (nombre exacto) en la hoja {H_PROD}."))
        if plan_k in productos_nombres_norm and mat_k in productos_nombres_norm:
            planes.add(plan_k)
            materiales.add(mat_k)
            findings.append(Finding(H, fila, "productos_por_plan", OK, "OK", f"{plan} -> {material}"))
    return findings, planes, materiales


# ---------------------------------------------------------------------
# PRODUCTOS
# ---------------------------------------------------------------------
def validar_productos(E, filas, categorias_index, planes_norm):
    H, T, H_CAT = E.hoja("productos"), E.tabla("producto"), E.hoja("categorias")
    findings = []
    index = {}
    vistos = {}
    for row in filas:
        fila = row["_excel_row"]

        def c(nombre):
            return campo(E, row, "productos", nombre)

        nombre = norm_str(c("nombre"))
        if not nombre:
            findings.append(Finding(H, fila, T, ERROR, f"Falta '{etiqueta_col(E, 'productos', 'nombre')}' (obligatorio)."))
            continue
        clave = norm_key(nombre)
        faltantes = []
        for f in ["activo", "se_puede_vender", "se_puede_comprar"]:
            if c(f) is None:
                faltantes.append(etiqueta_col(E, "productos", f))
        referencia = norm_str(c("referencia"))
        if not referencia:
            faltantes.append(etiqueta_col(E, "productos", "referencia"))
        if faltantes:
            findings.append(Finding(H, fila, T, ERROR, f"Faltan campos obligatorios: {', '.join(faltantes)}.", nombre))
            continue

        categoria = norm_str(c("categoria"))
        if categoria and norm_key(categoria) not in categorias_index:
            findings.append(Finding(H, fila, T, ADVERTENCIA,
                                    f"La categoria '{categoria}' no existe en la hoja {H_CAT} (se creara si continuas).", nombre))

        if clave in vistos:
            findings.append(Finding(H, fila, T, ADVERTENCIA,
                                    f"Nombre de producto duplicado (ya aparecio en la fila {vistos[clave]}).", nombre))

        vistos[clave] = fila
        index[clave] = {
            "fila": fila,
            "nombre": nombre,
            "referencia": referencia,
            "activo": c("activo"),
            "tiene_seriales": c("tiene_seriales"),
            "se_puede_vender": c("se_puede_vender"),
            "se_puede_comprar": c("se_puede_comprar"),
            "precio_venta": c("precio_venta"),
            "categoria_producto": categoria,
            "es_principal": clave in planes_norm,
            "extras": leer_extras(row, E.extras("producto")),
        }
        findings.append(Finding(H, fila, T, OK, "OK", nombre))
    return findings, index


# ---------------------------------------------------------------------
# CONSECUTIVOS Y PROYECTOS  (una fila crea un consecutivo + un proyecto)
# ---------------------------------------------------------------------
def validar_consecutivo_y_proyectos(E, filas, ciudades_idx):
    H = E.hoja("consecutivos_proyectos")
    T_CONS, T_PROY = E.tabla("consecutivo"), E.tabla("proyecto")
    T_AMBAS = f"{T_CONS}/{T_PROY}"
    cat_ciudades = f"{E.etiqueta('catalogo_ciudades', 'catalogo de ciudades')} ({E.tabla('ciudad')})"
    findings = []
    consecutivos = {}
    proyectos = {}
    for row in filas:
        fila = row["_excel_row"]

        def c(nombre):
            return campo(E, row, "consecutivos_proyectos", nombre)

        def lbl(nombre):
            return etiqueta_col(E, "consecutivos_proyectos", nombre)

        nombre_consec = norm_str(c("consecutivo"))
        prefijo = norm_str(c("prefijo"))
        activo = c("activo")
        tipo_documento = c("tipo_documento")
        nombre_proyecto = norm_str(c("proyecto"))
        ciudad = norm_str(c("ciudad"))

        faltantes = []
        if not nombre_consec:
            faltantes.append(f"{lbl('consecutivo')} (consecutivo)")
        if not prefijo:
            faltantes.append(lbl("prefijo"))
        if activo is None:
            faltantes.append(lbl("activo"))
        if tipo_documento is None:
            faltantes.append(lbl("tipo_documento"))
        if not nombre_proyecto:
            faltantes.append(lbl("proyecto"))
        if faltantes:
            findings.append(Finding(H, fila, T_AMBAS, ERROR, f"Faltan campos obligatorios: {', '.join(faltantes)}."))
            continue

        ciudad_id = None
        if ciudad:
            ciudad_id, _ = _resolver_ciudad_local(ciudades_idx, ciudad)
            if ciudad_id is None:
                findings.append(Finding(H, fila, T_PROY, ERROR,
                                        f"La ciudad '{ciudad}' no existe en el {cat_ciudades}. No se autocrea.", nombre_proyecto))

        consecutivos[norm_key(nombre_consec)] = {
            "nombre": nombre_consec, "prefijo": prefijo, "sufijo": c("sufijo"),
            "numero_inicial": c("numero_inicial"), "numero_actual": c("numero_actual"),
            "numero_de_digitos": c("numero_de_digitos"),
            "activo": activo, "tipo_documento": tipo_documento,
            "extras": leer_extras(row, E.extras("consecutivo")),
        }
        proyectos[norm_key(nombre_proyecto)] = {
            "nombre": nombre_proyecto, "consecutivo_nombre_norm": norm_key(nombre_consec),
            "ciudad": ciudad, "ciudad_id": ciudad_id,
        }
        findings.append(Finding(H, fila, T_AMBAS, OK, "OK", nombre_proyecto))
    return findings, consecutivos, proyectos


def _resolver_ciudad_local(ciudades_idx, nombre):
    from core.catalogs import resolver_ciudad
    return resolver_ciudad(ciudades_idx, nombre)


# ---------------------------------------------------------------------
# PERSONAS  (una fila puede ser persona, cliente y/o empleado)
# ---------------------------------------------------------------------
def validar_personas(E, filas):
    H = E.hoja("personas")
    T_PER, T_CLI, T_EMP = E.tabla("persona"), E.tabla("cliente"), E.tabla("empleado")
    reglas_cliente = E.reglas.get("personas", {}).get("cliente", {})
    detecta_cliente = reglas_cliente.get("se_detecta_por", [])
    obligatorios_cliente = reglas_cliente.get("obligatorios", [])
    bodega_automatica = E.reglas.get("personas", {}).get("empleado", {}).get("bodega_automatica", False)

    findings = []
    index = {}
    vistos_doc = {}
    for row in filas:
        fila = row["_excel_row"]

        def c(nombre):
            return campo(E, row, "personas", nombre)

        def lbl(nombre):
            return etiqueta_col(E, "personas", nombre)

        tipo_id = c("tipo_identificacion")
        numero_doc = norm_str(c("numero_documento"))
        nombres = norm_str(c("nombres"))

        faltantes = []
        if tipo_id is None:
            faltantes.append(lbl("tipo_identificacion"))
        if not numero_doc:
            faltantes.append("numero_documento")
        if not nombres:
            faltantes.append(lbl("nombres"))
        if faltantes:
            findings.append(Finding(H, fila, T_PER, ERROR, f"Faltan campos obligatorios: {', '.join(faltantes)}."))
            continue

        if numero_doc in vistos_doc:
            findings.append(Finding(H, fila, T_PER, ERROR,
                                    f"Documento duplicado (ya aparecio en la fila {vistos_doc[numero_doc]}).", numero_doc))
            continue
        vistos_doc[numero_doc] = fila

        email = norm_str(c("email"))
        es_cliente = any(c(f) is not None for f in detecta_cliente)
        if es_cliente:
            faltantes_cliente = [lbl(f) for f in obligatorios_cliente if norm_str(c(f)) is None]
            if faltantes_cliente:
                findings.append(Finding(H, fila, T_CLI, ERROR,
                                        f"Es cliente pero faltan campos obligatorios: {', '.join(faltantes_cliente)}.", numero_doc))

        activo_empleado = c("empleado_activo")
        almacen = c("almacen")
        cargo = c("cargo")
        es_empleado = cargo is not None or almacen is not None
        almacen_generado = bool(bodega_automatica) and es_empleado and not almacen
        if es_empleado and activo_empleado is None:
            findings.append(Finding(H, fila, T_EMP, ERROR,
                                    f"Es empleado pero falta el campo '{lbl('empleado_activo')}' (obligatorio).", numero_doc))

        index[numero_doc] = {
            "fila": fila, "nombres": nombres, "apellidos": norm_str(c("apellidos")),
            "tipo_identificacion": tipo_id, "email": email,
            "es_cliente": es_cliente, "es_empleado": es_empleado, "almacen_generado": almacen_generado,
        }
        findings.append(Finding(H, fila, T_PER, OK, "OK", numero_doc))
    return findings, index


# ---------------------------------------------------------------------
# CONTRATOS
# ---------------------------------------------------------------------
def _candidatos_proyecto_por_ciudad(proyectos_index, nombre_corto):
    """Si 'nombre_corto' no matchea exacto un proyecto, busca proyectos cuya
    ciudad coincida con ese nombre, para dar un mensaje de error util."""
    clave = norm_key(nombre_corto)
    return [p["nombre"] for p in proyectos_index.values() if norm_key(p.get("ciudad") or "") == clave]


def _resolver_proyecto_ambiguo_por_consecutivo(candidatos, proyectos_index, consecutivos_index, consecutivo_contrato):
    """
    Si hay varios proyectos candidatos por ciudad (ej 'NORTE RADIO' y
    'NORTE FIBRA'), usa el consecutivo de contrato real de la fila (ej
    'FNF162') para ver cual prefijo de consecutivo contiene - eso es
    evidencia real, no una suposicion. Si exactamente UN candidato matchea,
    lo devuelve resuelto. Si no, devuelve None (sigue ambiguo).
    """
    if not consecutivo_contrato:
        return None
    consec_norm = norm_key(consecutivo_contrato)
    matches = []
    for nombre_proy in candidatos:
        proy = proyectos_index.get(norm_key(nombre_proy))
        if not proy:
            continue
        consec = consecutivos_index.get(proy["consecutivo_nombre_norm"])
        if not consec or not consec.get("prefijo"):
            continue
        prefijo_norm = norm_key(consec["prefijo"])
        if prefijo_norm and prefijo_norm in consec_norm:
            matches.append(nombre_proy)
    if len(matches) == 1:
        return matches[0]
    return None


def _resolver_proyecto_ambiguo_por_palabras(candidatos, tipo_contrato_nombre, grupos_palabras):
    """
    Si el proyecto es ambiguo por ciudad, usa el tipo de contrato de la
    misma fila: busca a que grupo de palabras clave del esquema pertenece
    su nombre (ej 'Fibra' -> grupo fibra) y se queda con el unico proyecto
    candidato cuyo nombre contenga una palabra de ese mismo grupo.
    Se usa el NOMBRE del tipo de contrato y no un flag de la hoja, porque
    los flags pueden venir mal cargados en el Excel.
    """
    if not tipo_contrato_nombre:
        return None
    tokens_tipo = set(norm_key(tipo_contrato_nombre).split())
    palabras_esperadas = None
    for grupo in grupos_palabras:
        grupo_norm = {norm_key(p) for p in grupo}
        if tokens_tipo & grupo_norm:
            palabras_esperadas = grupo_norm
            break
    if not palabras_esperadas:
        return None
    matches = [nombre_proy for nombre_proy in candidatos if set(norm_key(nombre_proy).split()) & palabras_esperadas]
    return matches[0] if len(matches) == 1 else None


def _interpretar_texto_bool(valor, regla):
    """Convierte un texto libre en booleano segun palabras clave del esquema.
    Devuelve (bool|None, reconocido)."""
    s = norm_key(valor)
    if not s:
        return None, False
    if any(norm_key(p) in s for p in regla.get("falso_si_contiene", [])):
        return False, True
    if any(norm_key(p) in s for p in regla.get("verdadero_si_contiene", [])):
        return True, True
    return None, False


def validar_contratos(E, filas, personas_index, proyectos_index, tipo_contrato_index, ciudades_idx, consecutivos_index=None):
    consecutivos_index = consecutivos_index or {}
    H, T = E.hoja("contratos"), E.tabla("contrato")
    H_PER, H_PROY, H_TIPO = E.hoja("personas"), E.hoja("consecutivos_proyectos"), E.hoja("tipos_contrato")
    COD = E.etiqueta("codigo_servicio", "Codigo de servicio")
    cat_ciudades = f"{E.etiqueta('catalogo_ciudades', 'catalogo de ciudades')} ({E.tabla('ciudad')})"
    reglas = E.reglas.get("contratos", {})
    obligatorios = reglas.get("obligatorios", [])
    regla_factura = reglas.get("factura_electronica", {})
    grupos_palabras = reglas.get("desambiguar_proyecto_por_palabras", [])
    extras_instalacion = E.extras("instalacion") if E.tiene_tabla("instalacion") else []

    findings = []
    index = {}
    vistos_codigo = defaultdict(list)
    for row in filas:
        fila = row["_excel_row"]

        def c(nombre):
            return campo(E, row, "contratos", nombre)

        codigo = norm_str(c("codigo_servicio"))
        cliente_doc = norm_str(c("cliente"))

        faltantes = [etiqueta_col(E, "contratos", f) for f in obligatorios if c(f) is None]
        if faltantes:
            findings.append(Finding(H, fila, T, ERROR, f"Faltan campos obligatorios: {', '.join(faltantes)}.", codigo or cliente_doc))

        latitud = c("latitud")
        longitud = c("longitud")
        if latitud is None or longitud is None:
            findings.append(Finding(H, fila, T, ADVERTENCIA,
                                    "Falta latitud y/o longitud - se asignara 0 por defecto.", codigo or cliente_doc))
            latitud = latitud if latitud is not None else 0
            longitud = longitud if longitud is not None else 0

        if cliente_doc and cliente_doc not in personas_index:
            findings.append(Finding(H, fila, T, ERROR,
                                    f"El cliente (documento '{cliente_doc}') no existe en la hoja {H_PER}.", codigo))
        elif cliente_doc and not personas_index[cliente_doc]["es_cliente"]:
            findings.append(Finding(H, fila, T, ERROR,
                                    f"El documento '{cliente_doc}' existe en {H_PER} pero no esta marcado como cliente.", codigo))

        proyecto = norm_str(c("proyecto"))
        proyecto_resuelto = proyecto
        if proyecto and norm_key(proyecto) not in proyectos_index:
            candidatos = _candidatos_proyecto_por_ciudad(proyectos_index, proyecto)
            consecutivo_contrato_fila = norm_str(c("consecutivo_contrato"))

            if len(candidatos) > 1:
                tipo_contrato_fila = norm_str(c("tipo_contrato"))
                resuelto = _resolver_proyecto_ambiguo_por_consecutivo(
                    candidatos, proyectos_index, consecutivos_index, consecutivo_contrato_fila)
                fuente = f"el prefijo de su consecutivo de contrato '{consecutivo_contrato_fila}'"
                if not resuelto:
                    resuelto = _resolver_proyecto_ambiguo_por_palabras(candidatos, tipo_contrato_fila, grupos_palabras)
                    fuente = f"su tipo de contrato '{tipo_contrato_fila}'"
                if resuelto:
                    proyecto_resuelto = resuelto
                    findings.append(Finding(H, fila, T, ADVERTENCIA,
                                            f"Proyecto '{proyecto}' era ambiguo ({', '.join(candidatos)}), resuelto a "
                                            f"'{resuelto}' por {fuente}.", codigo))
                else:
                    findings.append(Finding(H, fila, T, ERROR,
                                            f"El proyecto '{proyecto}' es AMBIGUO: existen {len(candidatos)} proyectos para "
                                            f"esa ciudad ({', '.join(candidatos)}), y ni el consecutivo de contrato ni el tipo de "
                                            f"contrato de la fila permiten distinguir cual. Escribe el nombre completo en el Excel.", codigo))
            elif len(candidatos) == 1:
                findings.append(Finding(H, fila, T, ERROR,
                                        f"El proyecto '{proyecto}' no existe tal cual, pero hay un proyecto similar "
                                        f"por ciudad: '{candidatos[0]}'. Corrige el nombre en el Excel.", codigo))
            else:
                findings.append(Finding(H, fila, T, ERROR,
                                        f"El proyecto '{proyecto}' no existe en la hoja {H_PROY}.", codigo))

        tipo_contrato = norm_str(c("tipo_contrato"))
        if tipo_contrato and norm_key(tipo_contrato) not in tipo_contrato_index:
            findings.append(Finding(H, fila, T, ERROR,
                                    f"El tipo de contrato '{tipo_contrato}' no existe en la hoja {H_TIPO}.", codigo))

        ciudad_servicio = norm_str(c("ciudad_servicio"))
        if ciudad_servicio:
            ciudad_id, _ = _resolver_ciudad_local(ciudades_idx, ciudad_servicio)
            if ciudad_id is None:
                findings.append(Finding(H, fila, T, ERROR,
                                        f"La ciudad de servicio '{ciudad_servicio}' no existe en el {cat_ciudades}. No se autocrea.", codigo))

        factura_raw = c("factura_electronica")
        factura_electronica = None
        if factura_raw is not None:
            factura_electronica, reconocido = _interpretar_texto_bool(factura_raw, regla_factura)
            if not reconocido:
                findings.append(Finding(H, fila, T, ADVERTENCIA,
                                        f"Valor de '{etiqueta_col(E, 'contratos', 'factura_electronica')}' no reconocido: '{factura_raw}'.", codigo))

        if codigo:
            vistos_codigo[codigo].append((fila, cliente_doc))

        if not faltantes and codigo:
            index[codigo] = {
                "fila": fila, "cliente_doc": cliente_doc, "proyecto": proyecto_resuelto,
                "tipo_contrato": tipo_contrato, "factura_electronica": factura_electronica,
                "consecutivo_contrato": norm_str(c("consecutivo_contrato")),
                "latitud": latitud, "longitud": longitud,
                "instalacion": leer_extras(row, extras_instalacion),
            }
            findings.append(Finding(H, fila, T, OK, "OK", codigo))

    for codigo, apariciones in vistos_codigo.items():
        if len(apariciones) > 1:
            docs = {doc for _, doc in apariciones}
            filas_str = ", ".join(str(f) for f, _ in apariciones)
            if len(docs) > 1:
                findings.append(Finding(H, apariciones[0][0], T, ADVERTENCIA,
                                        f"{COD} '{codigo}' repetido en filas {filas_str} con clientes DISTINTOS "
                                        "(puede ser legitimo - revisar manualmente, no se bloquea).", codigo))
            else:
                findings.append(Finding(H, apariciones[0][0], T, ADVERTENCIA,
                                        f"{COD} '{codigo}' repetido en filas {filas_str} para el MISMO cliente.", codigo))
    return findings, index


# ---------------------------------------------------------------------
# PRODUCTOS CONTRATOS  (la forma de relacionarse con CONTRATOS varia)
# ---------------------------------------------------------------------
def detectar_convencion_join(E, contratos_filas, productos_contratos_filas):
    """
    Devuelve CONV_CODIGO si PRODUCTOS CONTRATOS tiene su propia columna de
    codigo de servicio; si no, prueba 'id contrato' contra el documento
    del cliente y contra el consecutivo de contrato de CONTRATOS, y
    devuelve la convencion que mas coincidencias tenga.
    """
    if productos_contratos_filas and any(a in productos_contratos_filas[0] for a in E.cols("productos_contratos", "codigo_servicio")):
        return CONV_CODIGO

    docs = {norm_str(campo(E, r, "contratos", "cliente")) for r in contratos_filas}
    consecutivos = {norm_str(campo(E, r, "contratos", "consecutivo_contrato")) for r in contratos_filas}

    def id_contrato(r):
        return norm_str(campo(E, r, "productos_contratos", "id_contrato"))

    match_doc = sum(1 for r in productos_contratos_filas if id_contrato(r) in docs)
    match_consec = sum(1 for r in productos_contratos_filas if id_contrato(r) in consecutivos)

    return CONV_CONSECUTIVO if match_consec >= match_doc else CONV_DOCUMENTO


def mapa_contratos(E, convencion, contratos_filas):
    """Indice para pasar del 'id contrato' de PRODUCTOS CONTRATOS al codigo de servicio."""
    if convencion == CONV_DOCUMENTO:
        doc_a_codigos = defaultdict(list)
        for r in contratos_filas:
            doc = norm_str(campo(E, r, "contratos", "cliente"))
            codigo = norm_str(campo(E, r, "contratos", "codigo_servicio"))
            if doc and codigo:
                doc_a_codigos[doc].append(codigo)
        return doc_a_codigos
    if convencion == CONV_CONSECUTIVO:
        consec_a_codigo = {}
        for r in contratos_filas:
            consec = norm_str(campo(E, r, "contratos", "consecutivo_contrato"))
            codigo = norm_str(campo(E, r, "contratos", "codigo_servicio"))
            if consec and codigo:
                consec_a_codigo[consec] = codigo
        return consec_a_codigo
    return {}


def codigos_destino(E, row, convencion, mapa):
    """Codigos de servicio de los contratos a los que pertenece una fila de PRODUCTOS CONTRATOS."""
    if convencion == CONV_CODIGO:
        codigo = norm_str(campo(E, row, "productos_contratos", "codigo_servicio"))
        return [codigo] if codigo else []
    id_contrato = norm_str(campo(E, row, "productos_contratos", "id_contrato"))
    if convencion == CONV_DOCUMENTO:
        return mapa.get(id_contrato, [])
    codigo = mapa.get(id_contrato)
    return [codigo] if codigo else []


def validar_productos_contratos(E, filas, convencion, contratos_filas, contratos_index, productos_index, planes_norm, materiales_norm):
    H, T, H_PROD = E.hoja("productos_contratos"), E.tabla("detalle_contrato"), E.hoja("productos")
    findings = []
    mapa = mapa_contratos(E, convencion, contratos_filas)

    for row in filas:
        fila = row["_excel_row"]
        producto = norm_str(campo(E, row, "productos_contratos", "producto"))
        cantidad = campo(E, row, "productos_contratos", "cantidad")
        precio = campo(E, row, "productos_contratos", "precio_venta")

        faltantes = []
        if not producto:
            faltantes.append(etiqueta_col(E, "productos_contratos", "producto"))
        if cantidad is None:
            faltantes.append(etiqueta_col(E, "productos_contratos", "cantidad"))
        if precio is None:
            faltantes.append(etiqueta_col(E, "productos_contratos", "precio_venta"))
        if faltantes:
            findings.append(Finding(H, fila, T, ERROR, f"Faltan campos obligatorios: {', '.join(faltantes)}."))
            continue

        clave_prod = norm_key(producto)
        if clave_prod not in productos_index:
            findings.append(Finding(H, fila, T, ERROR,
                                    f"El producto '{producto}' no existe (nombre exacto) en la hoja {H_PROD}.", producto))
            continue

        destinos = codigos_destino(E, row, convencion, mapa)
        if not destinos:
            findings.append(Finding(H, fila, T, ERROR,
                                    f"No se pudo relacionar esta fila con ningun contrato (convencion detectada: {convencion}).", producto))
            continue

        cargo_unico = clave_prod in materiales_norm and clave_prod not in planes_norm
        findings.append(Finding(H, fila, T, OK,
                                f"OK (-> {len(destinos)} contrato(s), cargo_unico={cargo_unico})", producto))

    return findings


# ---------------------------------------------------------------------
# Reconciliacion OPCIONAL contra una base destino ya poblada (matching
# tolerante) - solo se corre si el usuario conecta una BD en la interfaz.
# Actualiza productos_index/proyectos_index con 'id_real'/'nombre_real'
# cuando encuentra un match con confianza (exacto o tolerante unico),
# para que el generador de SQL reutilice la fila real en vez de crear
# un duplicado casi-identico.
# ---------------------------------------------------------------------
def _to_float(v):
    if v is None:
        return None
    try:
        return float(str(v).replace(",", "."))
    except (TypeError, ValueError):
        return None


def reconciliar_productos_con_bd(E, productos_index, catalogo_real):
    H, T = E.hoja("productos"), f"{E.tabla('producto')} (reconciliacion BD)"
    findings = []
    for info in productos_index.values():
        precio = _to_float(info.get("precio_venta"))
        resultado = resolver(info["nombre"], catalogo_real, precio_buscado=precio, contexto_extra=info.get("categoria_producto"))
        if resultado.estado in (ESTADO_EXACTO, ESTADO_TOLERANTE):
            info["id_real"] = resultado.candidato["id"]
            info["nombre_real"] = resultado.candidato["nombre"]
            if resultado.estado == ESTADO_TOLERANTE:
                findings.append(Finding(H, info.get("fila", 0), T, ADVERTENCIA,
                                        f"'{info['nombre']}' no coincide exacto pero SI coincide (tolerante, mismo precio) "
                                        f"con el producto ya existente '{resultado.candidato['nombre']}' en la base real - "
                                        "se reutilizara ese en vez de crear uno duplicado.", info["nombre"]))
        elif resultado.estado == ESTADO_AMBIGUO:
            nombres = ", ".join(c["nombre"] for c in resultado.candidatos_ambiguos)
            findings.append(Finding(H, info.get("fila", 0), T, ERROR,
                                    f"'{info['nombre']}' es AMBIGUO contra el catalogo real: {nombres}. "
                                    "No se puede decidir cual reutilizar sin revision manual.", info["nombre"]))
        # ESTADO_SIN_MATCH: no hay nada parecido en la BD real -> se creara como producto nuevo, no es error.
    return findings


def reconciliar_proyectos_con_bd(E, proyectos_index, catalogo_real):
    H, T = E.hoja("consecutivos_proyectos"), f"{E.tabla('proyecto')} (reconciliacion BD)"
    findings = []
    for info in proyectos_index.values():
        resultado = resolver(info["nombre"], catalogo_real, contexto_extra=info.get("ciudad"))
        if resultado.estado in (ESTADO_EXACTO, ESTADO_TOLERANTE):
            info["id_real"] = resultado.candidato["id"]
            info["nombre_real"] = resultado.candidato["nombre"]
            if resultado.estado == ESTADO_TOLERANTE:
                findings.append(Finding(H, 0, T, ADVERTENCIA,
                                        f"Proyecto '{info['nombre']}' no coincide exacto pero SI coincide (tolerante) "
                                        f"con el proyecto ya existente '{resultado.candidato['nombre']}' en la base real - "
                                        "se reutilizara ese en vez de crear uno duplicado.", info["nombre"]))
        elif resultado.estado == ESTADO_AMBIGUO:
            nombres = ", ".join(c["nombre"] for c in resultado.candidatos_ambiguos)
            findings.append(Finding(H, 0, T, ERROR,
                                    f"Proyecto '{info['nombre']}' es AMBIGUO contra el catalogo real: {nombres}. "
                                    "No se puede decidir cual reutilizar sin revision manual.", info["nombre"]))
    return findings
