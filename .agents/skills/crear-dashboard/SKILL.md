---
name: crear-dashboard
description: >-
  Flujo completo paso a paso para crear un nuevo dashboard analítico en Api_Desposte
  (controlador .NET para Excel/OneDrive o SQL Server, archivo HTML responsive en Docu_Paginas/
  con filtros temporales adaptados a la granularidad real de los datos, registro en index.html
  y verificación de compilación). Usar cuando el usuario pida crear un nuevo dashboard, informe,
  reporte o módulo en el proyecto.
---

# Crear Dashboard en Api_Desposte

Procedimiento estándar de punta a punta para crear un nuevo módulo analítico cumpliendo todas las reglas de `AGENTS.md`, `CLAUDE.md` y `.agents/rules/`.

## Paso 1: Definir alcance y granularidad antes de codificar

Confirmar con el usuario (o inferir del contexto si ya lo especificó):
1. **Fuente de datos:**
   - **Excel (OneDrive):** ruta relativa dentro de `OneDrive - Corporación Rico SAC\...`, hoja(s) y nombre(s) de tabla(s).
   - **SQL Server:** nombre del SP o consulta SQL y parámetros (ej. `@Fecha1`, `@Fecha2`).
2. **Granularidad temporal real de los datos:**
   - Determinar cuál es el nivel más fino que realmente trae la fuente: **Diario** (fechas), **Semanal** (año/semana), **Mensual** (año/mes), **Anual** o **Sin serie temporal**.
3. **Datos del portal:** título, categoría (`OPERACIONES`, `CUMPLIMIENTOS`, `COSTOS`), descripción breve e ícono.

## Paso 2: Crear el controlador en .NET (`ApiDesposte/Controllers/`)

Seguir la **Regla 6** (nombre de clase alineado al archivo, único `[Route("api/[controller]")]` y único verbo HTTP por acción):

- **Si la fuente es Excel (OneDrive):**
  - Leer `.agents/rules/conexion-excel-onedrive.md`.
  - Crear `ApiDesposte/Controllers/Excel/<Nombre>Controller.cs`.
  - Resolver la ruta dinámicamente con `ObtenerRutaExcel()` usando `Environment.SpecialFolder.UserProfile`.
  - Delegar apertura, lectura y caché obligatoriamente a `LectorTablasExcel` y `CacheTablasExcel` (`ApiDesposte.Funciones.Excel`). Nunca instanciar `FileStream`, `MemoryStream` ni `SpreadsheetDocument` en el controlador.
- **Si la fuente es SQL Server:**
  - Crear `ApiDesposte/Controllers/Sql/<Nombre>Controller.cs`.
  - Usar `SqlConnection` (`SqlDesposte` de `appsettings.json`) + `Dapper` devolviendo las columnas originales de la base de datos.

## Paso 3: Crear el dashboard HTML (`Docu_Paginas/<Nombre>.html`)

Consultar las reglas aplicables en `.agents/rules/` antes de generar el archivo:
- **Conexión API** (`.agents/rules/api-connection.md`): incluir `<script src="Funciones/config.js"></script>` en `<head>`, declarar `APP_CONFIG` apuntando a la ruta canónica única, usar `generarBasesApi()`, `fetchConTimeout()` y `detectarServidorApi()`.
- **Diseño responsive y tema** (`.agents/rules/responsive-design.md`):
  - Enlazar `<link rel="stylesheet" href="Funciones/responsive.css?v=3">` y bloque `@media` autónomo (compacto desde `max-width: 1600px`, tablet `<= 980px`, móvil `<= 768px`).
  - Usar la paleta industrial oficial (`#38353B`, `#434B57`, `#BAA063`, `#7A828E`, `#F2F4F8`).
  - 4 tarjetas KPI en una sola fila en laptop (`repeat(4, minmax(0, 1fr))`), `min-width: 0` en hijos de grid/flex y contenedores de gráficos con clase (sin alturas inline fijas).
  - Si un gráfico usa `chartjs-plugin-datalabels`, ocultarlas en celular con `display: () => window.innerWidth > 768`.
- **Vistas temporales y filtros de periodo (solo según granularidad real):**
  - **Criterio obligatorio:** mostrar únicamente las vistas (**Diario / Semanal / Mensual / Anual**) para las que realmente existan datos en la fuente. No agregar vistas inexistentes o artificiales:
    - **Fuente con fechas diarias:** habilitar Diario / Semanal / Mensual (y Anual si abarca varios años).
    - **Fuente semanal** (ej. indicadores semanales): habilitar **solo Semanal** (y Mensual/Anual solo si el negocio lo requiere y es agregable; nunca Diario).
    - **Fuente mensual** (ej. índice de rotación mensual, costos mensuales): habilitar **solo Mensual** (y Anual si hay varios años; nunca Diario ni Semanal).
    - **Fuente anual o sin tiempo:** omitir selector de vistas temporales.
  - Para cada vista habilitada que coincida con `.agents/rules/filtros-periodo-vistas.md`, aplicar selección múltiple con su valor por defecto:
    - **Diario:** 2 últimas semanas.
    - **Semanal:** 12 últimas semanas.
    - **Mensual:** último año.
  - Todo cálculo de semanas debe seguir ISO-8601 iniciando en **lunes** (`.agents/rules/calculo-semanas.md`).
- **Matrices jerárquicas** (Regla 5): texto plano sin iconos ni emojis decorativos; único icono permitido: control de expandir/colapsar.
- **Exportación Data Plana SQL** (`.agents/rules/exportar-data-plana-sql.md`):
  - **Obligatorio solo si la fuente es SQL Server:** cargar SheetJS (`xlsx.full.min.js`), guardar el array crudo en `rawSqlData`, agregar botón `Data Plana SQL (.xlsx)` y función `exportarDataPlanaSQL()`.

## Paso 4: Registrar el dashboard en el portal

- Agregar la entrada correspondiente en el catálogo de `Docu_Paginas/index.html` con su `id`, `titulo`, `archivo`, `categoria`, `endpoint` y `descripcion`.

## Paso 5: Verificación obligatoria

1. Compilar la API .NET y confirmar 0 errores:
   ```powershell
   dotnet build ApiDesposte\ApiDesposte.csproj --no-restore /p:UseSharedCompilation=false
   ```
2. Verificar en los archivos creados/editados:
   - Cero URLs fijas en `fetch` y cero rutas de usuario hardcodeadas en C#.
   - Botón `Data Plana SQL (.xlsx)` presente si el controlador está en `Controllers/Sql/`.
   - Vistas temporales estrictamente acordes a la granularidad de los datos.
  
