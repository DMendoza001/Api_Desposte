# Especificación Técnica de Implementación: Sistema Dinámico de Diagramas para Stored Procedures (SQL Server + API + HTML)

Este documento sirve como **guía de arquitectura y contrato técnico** para que un agente de desarrollo implemente un módulo de documentación visual dinámica dentro de un contenedor `index.html`, conectado a SQL Server a través de una API existente.

---

## 1. Objetivo del Sistema

Construir un flujo automatizado donde:
1. El usuario visualiza en el `index.html` (menú contenedor) la lista en vivo de los Stored Procedures (`Sp_%`) disponibles en SQL Server.
2. Al seleccionar un SP, la API verifica si el procedimiento fue modificado en SQL Server (`modify_date`).
3. Si el SP está actualizado en caché, devuelve el JSON del diagrama inmediatamente; si es nuevo o cambió, la API extrae el código SQL vivo (`OBJECT_DEFINITION`), lo envía a la API de Gemini para parsearlo bajo un esquema JSON estricto, guarda el resultado en caché y lo entrega al frontend.
4. Una única plantilla (`visor_sp.html`) renderiza el diagrama vectorial de 4 columnas con conectores ortogonales y la paleta **Comité de Negocios (Pantone Blue & Grey)**.

---

## 2. Arquitectura del Flujo de Datos

```text
[index.html (Menú Contenedor)]
   │
   ├── 1. OnLoad ──► GET /api/sps/lista
   │                    └── SQL Server: sys.procedures + sys.schemas
   │                    └── Retorna: [{ nombre, esquema, modifyDate, tieneCache }]
   │
   └── 2. OnSelect ─► GET /api/sps/diagrama?nombre={NombreSP}&forzar={bool}
                        ├── SQL Server: OBJECT_DEFINITION(OBJECT_ID(@NombreSP)) + modify_date
                        ├── Validación Caché (Archivo .json o Tabla dbo.Doc_StoredProcedures):
                        │     ├── HIT (modify_date igual y forzar=false) ──► Retorna JSON guardado
                        │     └── MISS (nuevo, modificado o forzar=true):
                        │           ├── POST a Gemini API (System Prompt + CodigoSQL)
                        │           ├── Valida estructura JSON
                        │           └── Guarda en Caché + Retorna JSON al Frontend
                        │
                        └── [visor_sp.html] recibe JSON ──► Renderiza diagrama SVG/HTML
```

---

## 3. Capa de Base de Datos (SQL Server)

### 3.1. Consulta de Catálogo de SPs (`GET /api/sps/lista`)
Debe ejecutarse sobre la base de datos operativa (ej. `Desposte`) filtrando objetos de sistema:

```sql
SELECT 
    s.name AS Esquema,
    p.name AS NombreSP,
    p.create_date AS FechaCreacion,
    p.modify_date AS UltimaModificacion
FROM sys.procedures p
INNER JOIN sys.schemas s ON p.schema_id = s.schema_id
WHERE p.is_ms_shipped = 0
  AND p.name LIKE 'Sp_%'
ORDER BY p.name ASC;
```

### 3.2. Consulta de Definición Viva (`GET /api/sps/diagrama`)
Extrae el código fuente exacto compilado en el motor:

```sql
SELECT 
    p.name AS NombreSP,
    s.name AS Esquema,
    p.modify_date AS UltimaModificacion,
    OBJECT_DEFINITION(p.object_id) AS CodigoSQL
FROM sys.procedures p
INNER JOIN sys.schemas s ON p.schema_id = s.schema_id
WHERE p.name = @NombreSP
  AND p.is_ms_shipped = 0;
```

### 3.3. Tabla Opcional de Caché en SQL Server
Si se prefiere almacenar el caché en base de datos en lugar de archivos `.json` en disco:

```sql
CREATE TABLE dbo.Doc_StoredProcedures_Cache (
    NombreSP NVARCHAR(128) NOT NULL PRIMARY KEY,
    Esquema NVARCHAR(64) NOT NULL DEFAULT 'dbo',
    UltimaModificacionSQL DATETIME NOT NULL,
    DiagramaJson NVARCHAR(MAX) NOT NULL,
    FechaGeneracion DATETIME NOT NULL DEFAULT GETDATE()
);
```

---

## 4. Contrato JSON del Diagrama (Schema Estricto)

Tanto el prompt enviado a la API de Gemini como el renderizador `visor_sp.html` deben respetar **exactamente** esta estructura JSON:

```json
{
  "spName": "Sp_Informe_Prduccion_Cabezas",
  "esquema": "Desposte.dbo",
  "badge": "Desposte.dbo · SKU 040508029",
  "subtitulo": "Trazabilidad diaria de cabezas de cerdo: Corte Primario, Línea, Congelado, 3er Corte, Merma y Saldos",
  "ultimaModificacion": "2026-11-25T17:15:00",
  "etapa1_fuentes": {
    "titulo": "1. Fuentes y Tabla Temporal (Desposte.dbo)",
    "tarjetas": [
      {
        "titulo": "Tabla10 (Movimientos Desposte)",
        "estiloCabecera": "primary",
        "items": [
          "grupopuntoproceso = 1",
          "PuntoProceso, numero, plantilla (IS NOT NULL)",
          "ttra IN (13, 4) (13=Origen, 4=Resultado)",
          "Producto, cantidadunidades, cantidad"
        ]
      },
      {
        "titulo": "Fechas_OrdeProduccion(@Fecha1, @Fecha2, 1)",
        "estiloCabecera": "secondary",
        "joinLabel": "INNER JOIN: (grupopuntoproceso, PuntoProceso, numero)",
        "items": [
          "Obtiene FechaOrden oficial para agrupar por día"
        ]
      }
    ]
  },
  "etapa2_transformacion": {
    "titulo": "2. Clasificación por Nivel y Cálculo de Saldos (CTEs)",
    "tarjetas": [
      {
        "titulo": "CTE: LineaTercerCorte",
        "subbloques": [
          {
            "subtitulo": "LINEA DE CABEZA",
            "bullets": ["Punto 1003 · Plantilla 3016"]
          },
          {
            "subtitulo": "CONGELADO DIRECTO",
            "bullets": ["Punto 1011 · Plantilla 1027"]
          },
          {
            "subtitulo": "TERCER CORTE",
            "bullets": ["Punto 1003 · Plantillas 19, 1274, 1294, 1013"]
          }
        ]
      },
      {
        "titulo": "CTE: CortePrimario",
        "subbloques": [
          {
            "subtitulo": "Despiece Inicial",
            "bullets": ["Puntos 1016, 1017 · ttra = 4", "Producto fijo: '040508029'"]
          }
        ]
      }
    ],
    "nodoCentral": {
      "titulo": "UnionesBase & SaldoPendiente",
      "descripcion": "Filtra SKU 040508029 en ResumenFlujo",
      "formula": "Salida Línea (4) - Entrada 3er Corte (13) - Entrada Congelado (13) = SALDO PENDIENTE"
    }
  },
  "etapa3_catalogos": {
    "titulo": "3. DatasetFinal, Catálogo y Salida del Informe",
    "tituloTarjeta": "Consolidación y Catálogo",
    "pasos": [
      {
        "caja": "CTE: DatasetFinal",
        "detalle": "UnionesBase UNION ALL SaldoPendiente",
        "tipoJoin": "LEFT JOIN (Producto = CodigoArticulo)"
      },
      {
        "caja": "Articulo_Codigo_Corto()",
        "detalle": "Extrae CodigoCorto, Nombre y TipoDesposte",
        "tipoJoin": ""
      }
    ],
    "badges": [
      "1. CORTE PRIMARIO",
      "2. LINEA DE CABEZA",
      "3. CONGELADO DIRECTO",
      "4. TERCER CORTE",
      "5. SALDO PENDIENTE"
    ]
  },
  "etapa4_salida": {
    "tituloTarjeta": "Resultado Final",
    "columnas": [
      "FechaOrden",
      "ttra",
      "NTtra (ORIGEN / RESULTADO)",
      "Producto",
      "CodigoCorto",
      "Nombre",
      "TipoDesposte",
      "Unidades",
      "Kilos",
      "Nivel"
    ],
    "pieFiltro": "ORDER BY FechaOrden, Nivel, ttra"
  }
}
```

---

## 5. Prompt de Extracción para el Backend (Llamada a Gemini API)

Cuando la API detecte un *cache miss*, debe invocar al modelo generativo (ej. `gemini-2.5-flash` con `responseMimeType: "application/json"`) enviando el siguiente **System Instruction**:

> Eres un arquitecto de datos experto en SQL Server. Analiza el Stored Procedure proporcionado y extrae su flujo ETL completo en formato JSON estricto siguiendo el esquema definido.
> Reglas de extracción:
> 1. **Etapa 1 (Fuentes):** Identifica las tablas base (`Tabla10`, `#Temp`), funciones (`Fechas_OrdeProduccion`), filtros de `grupopuntoproceso`, rangos de fecha y llaves exactas de `JOIN`.
> 2. **Etapa 2 (Transformación):** Agrupa cada CTE o bloque lógico indicando explícitamente los códigos de `PuntoProceso`, códigos de `plantilla` incluidos o excluidos (`IN` / `NOT IN`), y valores de `ttra` (`13` = Origen/Insumo, `4` = Resultado/Producción). Resume cualquier cálculo de saldo, merma o cascada FIFO en `nodoCentral`.
> 3. **Etapa 3 (Catálogos):** Lista los `LEFT JOIN` o `INNER JOIN` con tablas maestras de artículos/productos y las llaves de cruce.
> 4. **Etapa 4 (Salida):** Lista exactamente las columnas del `SELECT` final y las cláusulas `WHERE` / `ORDER BY` de cierre.
> Devuelve únicamente JSON válido sin bloques markdown ni texto adicional.

---

## 6. Especificaciones del Frontend (`index.html` y `visor_sp.html`)

### 6.1. Integración en el `index.html` (Menú Contenedor)
* Agregar una sección **"Documentación de SPs"** en el menú lateral.
* Al iniciar, consumir `GET /api/sps/lista` y renderizar un buscador instantáneo + lista de botones con el nombre del SP y un indicador visual de fecha de última modificación.
* Al hacer clic en un SP, cargar `visor_sp.html?sp={NombreSP}` dentro del área principal (`<iframe>` o contenedor dinámico).

### 6.2. Plantilla Única `visor_sp.html`
Debe implementar la identidad visual **Comité de Negocios (Pantone Blue & Grey)**:
* **Variables CSS obligatorias:**
  * `--color-header-bg: #353A4C`
  * `--color-header-gradient: linear-gradient(135deg, #353A4C 0%, #3C3F4A 100%)`
  * `--color-primary: #5A77A8`
  * `--color-secondary: #748CAB`
  * `--color-bg: #F4F6F9`
  * `--color-surface: #FFFFFF`
  * `--color-ice: #C6D2E0`
* **Cálculo dinámico de conectores SVG:** Dado que la altura de las tarjetas varía según la cantidad de reglas de cada SP, el script `renderDiagrama(json)` debe:
  1. Inyectar el HTML de las 4 columnas dentro del tablero.
  2. Leer `getBoundingClientRect()` de cada tarjeta renderizada.
  3. Dibujar dinámicamente los `<path>` ortogonales de 90° (`stroke="#5A77A8"`, `stroke-width="2.5"`) con `<marker>` de flecha entre el borde derecho de cada origen y el borde izquierdo de cada destino.
* **Controles de barra superior:**
  * Botones de enfoque por etapa (`Vista Completa`, `1. Fuentes`, `2. Transformación`, `3. Catálogos y Salida`) que aplican `opacity: 0.22` al resto.
  * Controles de zoom escalable (`55%` a `150%`).
  * Botón **"↻ Regenerar desde SQL"** que llama a `GET /api/sps/diagrama?nombre={NombreSP}&forzar=true` para refrescar el caché bajo demanda.
  * Botón **"🖨️ Exportar PDF"** con `@media print` en orientación `landscape`.

---

## 7. Checklist de Tareas para el Agente Implementador

- [ ] **Backend 1:** Crear endpoint `GET /api/sps/lista` consultando `sys.procedures` y `sys.schemas`.
- [ ] **Backend 2:** Crear almacenamiento de caché (`Doc_StoredProcedures_Cache` en SQL o carpeta `/cache/sps/*.json`).
- [ ] **Backend 3:** Crear endpoint `GET /api/sps/diagrama?nombre={sp}&forzar={bool}` con validación de `modify_date` e integración con Gemini API (`application/json`).
- [ ] **Frontend 1:** Construir `visor_sp.html` con la paleta *Comité de Negocios*, renderizado dinámico desde JSON y trazado automático de rutas SVG por coordenadas `getBoundingClientRect()`.
- [ ] **Frontend 2:** Conectar el menú del `index.html` contenedor al endpoint de lista y pasar el parámetro `?sp=` al visor.