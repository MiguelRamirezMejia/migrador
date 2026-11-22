-- =====================================================================
-- Base de datos de EJEMPLO (PostgreSQL) que corresponde a
-- config/esquema.ejemplo.yaml. Sirve para probar el flujo completo:
--
--   createdb migrador_demo
--   psql -d migrador_demo -f ejemplos/esquema_ejemplo.sql
--   psql -d migrador_demo -c "\copy ciudades (id, nombre, codigo_dane) FROM 'data/ciudades.csv' WITH (FORMAT csv, HEADER true)"
--   psql -d migrador_demo -f migracion.sql      -- el SQL que genera la app
-- =====================================================================

CREATE TABLE ciudades (
    id          INTEGER PRIMARY KEY,
    nombre      TEXT NOT NULL,
    codigo_dane TEXT
);

CREATE TABLE tipos_servicio (
    id                SERIAL PRIMARY KEY,
    nombre            TEXT NOT NULL,
    codigo            TEXT NOT NULL UNIQUE,
    penalidad         NUMERIC,
    meses_permanencia INTEGER
);

CREATE TABLE categorias_producto (
    id          SERIAL PRIMARY KEY,
    nombre      TEXT NOT NULL UNIQUE,
    descripcion TEXT,
    activo      BOOLEAN NOT NULL DEFAULT TRUE
);

CREATE TABLE productos (
    id              SERIAL PRIMARY KEY,
    nombre          TEXT NOT NULL UNIQUE,
    referencia      TEXT NOT NULL,
    activo          BOOLEAN NOT NULL,
    maneja_seriales BOOLEAN NOT NULL,
    se_vende        BOOLEAN,
    se_compra       BOOLEAN,
    precio          NUMERIC,
    categoria_id    INTEGER REFERENCES categorias_producto (id),
    unidad          TEXT,
    velocidad_mbps  INTEGER,
    es_principal    BOOLEAN
);

CREATE TABLE consecutivos (
    id                SERIAL PRIMARY KEY,
    nombre            TEXT NOT NULL UNIQUE,
    prefijo           TEXT NOT NULL,
    sufijo            TEXT,
    numero_inicial    INTEGER,
    numero_actual     INTEGER,
    digitos           INTEGER,
    activo            BOOLEAN NOT NULL,
    tipo_documento_id INTEGER NOT NULL,
    organizacion_id   INTEGER
);

CREATE TABLE sedes (
    id             SERIAL PRIMARY KEY,
    nombre         TEXT NOT NULL UNIQUE,
    ciudad_id      INTEGER REFERENCES ciudades (id),
    consecutivo_id INTEGER REFERENCES consecutivos (id)
);

CREATE TABLE personas (
    id                     SERIAL PRIMARY KEY,
    tipo_identificacion_id INTEGER NOT NULL,
    documento              TEXT NOT NULL UNIQUE,
    nombres                TEXT NOT NULL,
    apellidos              TEXT,
    email                  TEXT
);

-- Cliente y empleado comparten el id de la persona.
CREATE TABLE clientes (
    id         INTEGER PRIMARY KEY REFERENCES personas (id),
    persona_id INTEGER NOT NULL UNIQUE REFERENCES personas (id)
);

CREATE TABLE almacenes (
    id              SERIAL PRIMARY KEY,
    nombre          TEXT NOT NULL UNIQUE,
    nombre_corto    TEXT,
    organizacion_id INTEGER
);

CREATE TABLE ubicaciones (
    id                SERIAL PRIMARY KEY,
    almacen_id        INTEGER NOT NULL REFERENCES almacenes (id),
    nombre            TEXT NOT NULL,
    tipo_ubicacion_id INTEGER,
    disponible_venta  BOOLEAN
);

CREATE TABLE empleados (
    id               INTEGER PRIMARY KEY REFERENCES personas (id),
    persona_id       INTEGER NOT NULL UNIQUE REFERENCES personas (id),
    activo           BOOLEAN NOT NULL,
    almacen_id       INTEGER REFERENCES almacenes (id),
    almacen_generado BOOLEAN
);

CREATE TABLE nodos (
    id     SERIAL PRIMARY KEY,
    nombre TEXT NOT NULL UNIQUE
);

INSERT INTO nodos (nombre) VALUES ('NODO-POBLADO'), ('NODO-ENVIGADO');

CREATE TABLE contratos (
    id                  SERIAL PRIMARY KEY,
    cliente_id          INTEGER NOT NULL REFERENCES clientes (id),
    sede_id             INTEGER NOT NULL REFERENCES sedes (id),
    tipo_servicio_id    INTEGER NOT NULL REFERENCES tipos_servicio (id),
    codigo_servicio     TEXT NOT NULL,
    numero_contrato     TEXT,
    factura_electronica BOOLEAN,
    latitud             NUMERIC,
    longitud            NUMERIC
);

CREATE TABLE contratos_instalacion (
    id          SERIAL PRIMARY KEY,
    contrato_id INTEGER NOT NULL UNIQUE REFERENCES contratos (id),
    ip_fija     TEXT,
    nodo_id     INTEGER REFERENCES nodos (id)
);

CREATE TABLE contratos_detalle (
    id          SERIAL PRIMARY KEY,
    contrato_id INTEGER NOT NULL REFERENCES contratos (id),
    producto_id INTEGER NOT NULL REFERENCES productos (id),
    precio      NUMERIC NOT NULL,
    cantidad    NUMERIC NOT NULL,
    cargo_unico BOOLEAN NOT NULL DEFAULT FALSE
);
