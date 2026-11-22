# Migrador — Validador de plantillas de migración

Aplicación web local (Streamlit) que recibe una **plantilla Excel de migración**, la valida hoja por hoja contra las reglas del sistema destino y, si no hay errores bloqueantes, genera el **script SQL de migración** para PostgreSQL.

El motor no conoce ningún sistema en particular: los nombres de hojas, columnas y tablas salen de un **archivo de esquema** (YAML). El repositorio trae un esquema de ejemplo, el de una empresa de servicios por suscripción, con su plantilla y su base de datos de prueba.

- **Validación fila por fila** con mensajes claros: campos obligatorios, duplicados, referencias entre hojas y ciudades contra un catálogo.
- **SQL autocontenido e idempotente**: las llaves foráneas se resuelven con subconsultas por llave de negocio, todo va en una transacción y cada `INSERT` comprueba que la fila no exista. Se puede correr dos veces sin duplicar nada.
- **Funciona sin conexión a la base de datos.** La conexión solo hace falta para la reconciliación opcional contra una base ya poblada, que usa *matching* tolerante para reutilizar productos y sedes existentes en vez de duplicarlos.
- **Nunca adivina**: si un dato es ambiguo, intenta resolverlo con evidencia de la misma fila y, si no puede, lo marca como error.

---

## 1. Probarlo en dos minutos

```powershell
git clone https://github.com/MiguelRamirezMejia/migrador.git
cd migrador
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
python ejemplos/generar_plantilla.py
streamlit run app.py
```

En el navegador (<http://localhost:8501>) sube `ejemplos/plantilla_ejemplo.xlsx` y presiona **Validar plantilla**. Verás 2 advertencias resueltas automáticamente (una sede ambigua y un contrato sin coordenadas) y podrás descargar el reporte y el `migracion.sql`.

Para ejecutar ese SQL de verdad, crea la base de ejemplo:

```bash
createdb migrador_demo
psql -d migrador_demo -f ejemplos/esquema_ejemplo.sql
psql -d migrador_demo -c "\copy ciudades (id, nombre, codigo_dane) FROM 'data/ciudades.csv' WITH (FORMAT csv, HEADER true)"
psql -d migrador_demo -f migracion.sql
```

---

## 2. Requisitos e instalación

| Requisito | Versión | Notas |
|---|---|---|
| Python | **3.10 o superior** (probado con 3.13) | El código usa la sintaxis `dict \| None`. |
| PostgreSQL | — | **No es necesario** para validar ni para generar el SQL. Solo hace falta para *ejecutar* el `.sql` o para la reconciliación opcional. |

Dependencias (en `requirements.txt`): `streamlit`, `pandas`, `openpyxl`, `psycopg2-binary`, `pyyaml`.

**Windows**: al instalar Python marca **"Add python.exe to PATH"**. Si PowerShell bloquea `Activate.ps1`, ejecuta una vez `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned`, o usa `venv\Scripts\activate.bat` desde `cmd`.

**Linux / macOS**:

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

> Si `pip install` falla detrás de un proxy corporativo, prueba con `--trusted-host pypi.org --trusted-host files.pythonhosted.org` o con `--proxy http://usuario:clave@proxy:puerto`.

---

## 3. Uso

1. En la barra lateral, comprueba qué **esquema** está cargado y sube la plantilla (`.xlsx`).
2. *(Opcional)* Activa **"Reconciliar contra una base ya poblada"** e ingresa los datos de conexión a PostgreSQL. Las consultas son de solo lectura.
3. Presiona **Validar plantilla** y revisa:
   - Tarjetas con el total de filas evaluadas, **OK**, **Advertencias** y **Errores bloqueantes**.
   - Tabla filtrable por hoja, severidad y texto libre.
4. Descarga:
   - **Reporte completo (Excel)** → `reporte_validacion_migracion.xlsx`.
   - **SQL de migración** → `migracion.sql` (solo se habilita si hay **0 errores**).
5. Ejecuta `migracion.sql` contra la base destino (con `psql`, DBeaver, etc.). Todo va dentro de `BEGIN; … COMMIT;`, así que si algo falla no queda a medias. Revisa al final del archivo la sección *ADVERTENCIAS DE GENERACIÓN*.

### Hojas de la plantilla (esquema de ejemplo)

| Hoja | Fila del encabezado | Qué crea |
|---|---|---|
| `TIPOS DE SERVICIO` | 1 | tipos de servicio (su `codigo` es el prefijo del número de contrato) |
| `CATEGORIAS` | 1 | categorías de producto |
| `PRODUCTOS` | 1 | productos |
| `SEDES` | 1 | un consecutivo de numeración + una sede por fila |
| `PERSONAS` | **2** (la fila 1 agrupa CLIENTE / EMPLEADO) | persona, y además cliente y/o empleado según las columnas llenas |
| `CONTRATOS` | 1 | contratos + información de instalación |
| `DETALLE CONTRATOS` | 1 | productos de cada contrato |

Las columnas se buscan **por nombre de encabezado**, no por posición. Las filas vacías y los encabezados repetidos a mitad de hoja se ignoran.

---

## 4. Adaptarlo a tu sistema

Copia el esquema de ejemplo a la carpeta `privado/` y ajústalo:

```powershell
mkdir privado
copy config\esquema.ejemplo.yaml privado\esquema.yaml
```

La app usa el primer esquema que encuentre en este orden:

1. La variable de entorno `MIGRADOR_ESQUEMA` (ruta a un `.yaml`).
2. `privado/esquema.yaml`.
3. `config/esquema.ejemplo.yaml`.

**La carpeta `privado/` está en `.gitignore`**: los nombres reales de tu base de datos, tus catálogos y tus notas nunca se suben al repositorio. Como no viaja con `git pull`, cópiala a mano a cada equipo donde uses la herramienta.

El esquema tiene tres partes:

| Sección | Qué define |
|---|---|
| `plantilla` | Para cada hoja lógica: nombre en el Excel, fila del encabezado y nombres de columna aceptados para cada campo (texto o lista de alternativas). También `marcadores_encabezado`, los textos que delatan un encabezado repetido a mitad de hoja. |
| `reglas` | Qué campos marcan a una persona como cliente y cuáles le son obligatorios, si se crea un almacén automático para empleados, campos obligatorios de los contratos, palabras clave para convertir texto libre en booleano y grupos de palabras para resolver sedes ambiguas. |
| `tablas` | Para cada tabla lógica: nombre real, columna real de cada campo, `extras` (columnas que pasan directo del Excel, de tipo `texto`, `numero`, `booleano` o `busqueda` por nombre en otra tabla) y `fijos` (columnas con un valor SQL constante). |

Hojas lógicas: `tipos_contrato`, `categorias`, `productos`, `productos_por_plan`, `consecutivos_proyectos`, `personas`, `contratos`, `productos_contratos`.
Tablas lógicas: `tipo_contrato`, `categoria`, `producto`, `consecutivo`, `proyecto`, `ciudad`, `persona`, `cliente`, `empleado`, `almacen`, `ubicacion`, `contrato`, `instalacion`, `detalle_contrato`.

Los nombres de tablas y columnas se interpolan en el SQL, así que el esquema solo acepta identificadores simples (`tabla` o `esquema.tabla`). Cualquier otra cosa se rechaza al cargarlo.

El catálogo de ciudades es un CSV con columnas `id` y `nombre` (la ruta va en `tablas.ciudad.catalogo_csv`). El incluido, `data/ciudades.csv`, trae los municipios de Colombia con su código DANE.

---

## 5. Estructura del proyecto

```
migrador/
├── app.py                     # Interfaz Streamlit (carga, resultados, descargas)
├── requirements.txt
├── config/
│   └── esquema.ejemplo.yaml   # Esquema de ejemplo (plantilla + tablas + reglas)
├── data/
│   └── ciudades.csv           # Catálogo de municipios (código DANE)
├── ejemplos/
│   ├── generar_plantilla.py   # Crea plantilla_ejemplo.xlsx con datos ficticios
│   └── esquema_ejemplo.sql    # Base PostgreSQL de ejemplo para ejecutar el SQL
├── privado/                   # (ignorada por git) tu esquema y catálogos reales
└── core/
    ├── esquema.py             # Carga y valida el esquema YAML
    ├── pipeline.py            # Orquesta la validación en orden de dependencias
    ├── validators.py          # Reglas de validación por hoja → Findings + índices
    ├── excel_loader.py        # Lectura del .xlsx por nombre de encabezado + normalización
    ├── catalogs.py            # Catálogo de ciudades local
    ├── matching.py            # Matching tolerante (exacto → tokens → contexto; nunca adivina)
    ├── db.py                  # Conexión PostgreSQL OPCIONAL para reconciliación
    ├── sql_generator.py       # Genera migracion.sql con FKs resueltas por subconsulta
    └── finding.py             # Modelo de resultado (OK / ADVERTENCIA / ERROR)
```

### Cómo funciona por dentro

1. **Validación** (`pipeline.ejecutar_validacion`): carga cada hoja y la valida en orden, primero los catálogos y luego lo que depende de ellos. Cada validador devuelve una lista de *findings* y un índice que usan las hojas siguientes; por ejemplo, CONTRATOS cruza contra PERSONAS, SEDES y TIPOS DE SERVICIO.
2. **Sedes ambiguas**: si un contrato nombra una ciudad en vez de una sede, y esa ciudad tiene varias sedes, se resuelve primero por el prefijo del número de contrato y luego por las palabras clave del tipo de servicio. Si nada lo distingue, es error.
3. **Relación detalle → contrato**: si `DETALLE CONTRATOS` trae su propia columna de código de servicio, se usa esa. Si no, la columna `contrato` puede traer el documento del cliente o el número de contrato, y la convención se detecta automáticamente probando contra `CONTRATOS`.
4. **Números de contrato**: si vienen en el Excel se respetan; si no, se generan con *código del tipo de servicio + prefijo + número con ceros + sufijo*, continuando desde el `numero_actual` de cada consecutivo.
5. **Reconciliación opcional**: con conexión a la base, productos y sedes se comparan contra los reales; los que coinciden se reutilizan en vez de duplicarse.
6. **Generación de SQL**: las FKs se resuelven con subconsultas por llave de negocio (nombre, documento, código de servicio…) al ejecutar el script, así que el `.sql` no depende de conocer los IDs del destino. La única excepción es la ciudad, que usa el ID del catálogo.

---

## 6. Problemas comunes

| Síntoma | Solución |
|---|---|
| `'streamlit' no se reconoce como un comando` | No activaste el entorno virtual. Ejecuta `.\venv\Scripts\Activate.ps1` o usa `python -m streamlit run app.py`. |
| `ModuleNotFoundError: No module named 'core'` | Estás ejecutando desde otra carpeta. Haz `cd` a la raíz del proyecto (donde está `app.py`). |
| `ModuleNotFoundError: No module named 'yaml'` | Ejecuta `pip install -r requirements.txt` otra vez. |
| "No se pudo cargar el esquema de migración" | El YAML tiene un error o un identificador SQL no válido; el mensaje indica cuál. |
| "Falta una hoja o columna esperada en el Excel" | Los nombres de las hojas no coinciden con `plantilla.hojas` del esquema cargado. |
| "No se pudo conectar a la base de datos" | Verifica host, puerto y usuario, que PostgreSQL acepte conexiones desde tu IP y que no lo bloquee un firewall o VPN. |
| El puerto 8501 está ocupado | `streamlit run app.py --server.port 8502` |
| `TypeError: unsupported operand type(s) for \|` | Python < 3.10. Instala una versión más reciente y recrea el `venv`. |
