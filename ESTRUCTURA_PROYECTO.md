# Estructura del proyecto y funciones compartidas

Referencia de **Api_Desposte** (API .NET + dashboards HTML): dónde va cada archivo y qué piezas reutilizables existen. La regla 10 de `CLAUDE.md` / `AGENTS.md` exige respetar esta organización.

- **Sección 1:** árbol de carpetas.
- **Secciones 2 a 5:** arranque de la API, funciones compartidas y endpoints.
- **Sección 6:** pendientes para acercarse a la estructura de referencia (`Guia_Proyecto_Nuevo/Referencia_Estructura_Informe_html.md`).

---

## 1. Árbol de carpetas

```
Api_Desposte/
├── CLAUDE.md                     ← reglas del proyecto para Claude Code
├── AGENTS.md                     ← mismas reglas para otros agentes (siempre sincronizado con CLAUDE.md)
├── ESTRUCTURA_PROYECTO.md        ← este documento
├── PaletaFinanza_Industrial.json ← paleta de colores de los dashboards
│
├── .agents/
│   ├── rules/                    ← detalle de cada regla
│   │   ├── api-connection.md          conexión frontend → API
│   │   ├── calculo-semanas.md         semanas ISO-8601, lunes a domingo
│   │   ├── conexion-excel-onedrive.md rutas dinámicas de OneDrive + LectorTablasExcel
│   │   ├── responsive-design.md       breakpoints 1600 / 980 / 768
│   │   ├── exportar-data-plana-sql.md botón "Data Plana SQL (.xlsx)"
│   │   └── filtros-periodo-vistas.md  filtros Diario / Semanal / Mensual
│   └── skills/crear-dashboard/   ← skill para crear dashboards nuevos
│
├── Documentacion/                ← documentos del proyecto (arquitectura del servidor, descripción de la API)
├── Guia_Proyecto_Nuevo/          ← cómo crear otro proyecto con esta misma base
│
├── ApiDesposte/                  ← API .NET 10 (namespace ApiDesposte)
│   ├── Program.cs                   CORS, protección de páginas en prueba, archivos estáticos, /api/health, puertos 8080/5000/5050/8085
│   ├── appsettings.json             sin secretos (se leen de variables de entorno)
│   ├── Controllers/                 un subdirectorio por fuente de datos
│   │   ├── Excel/                      tablas de Excel en OneDrive
│   │   ├── Sql/                        procedimientos de SQL Server (+ SpsController: documentación de SPs)
│   │   └── Portal/                     configuración del portal (PortalConfigController)
│   ├── Funciones/                   ← código reutilizable del backend (namespace ApiDesposte.Funciones.<Área>)
│   │   ├── Excel/LectorTablasExcel.cs   apertura segura de Excel + lectura de tablas + caché
│   │   ├── Excel/EditorTablasExcel.cs   CRUD sobre tablas de Excel (con respaldo previo a cada escritura)
│   │   ├── Sql/EjecutorSql.cs           ejecución de SP devolviendo filas tal cual
│   │   ├── Fechas/ParametrosFecha.cs    lectura de fechas de la query string
│   │   ├── Portal/EstadoDashboards.cs   ruta y lectura de Datos/dashboards_estado.json
│   │   └── IA/                          DeepSeek y Gemini con respaldo automático (GeneradorJsonIA)
│   └── Datos/                       ← datos de ejecución, fuera de la web (excluido de compilación y publicación)
│       ├── dashboards_estado.json      dashboards activos / en prueba (los "en prueba" solo los ve el administrador)
│       ├── sps/                        caché de diagramas de SPs generados con IA, una carpeta por base (se crea sola)
│       └── RespaldosExcel/             copia del Excel antes de cada escritura del CRUD (últimas 30; no se sube a git)
│
└── Docu_Paginas/                 ← carpeta web que sirve la API (http://RAEVSALL001:8080/)
    ├── Funciones/                   ← código reutilizable del frontend
    │   ├── config.js                   servidor y puertos de la API
    │   └── responsive.css              estilos responsive compartidos
    ├── index.html                   portal de dashboards (pestañas por categoría)
    ├── Documentacion_SPs.html       lista de SPs + visor_sp.html (diagrama de cada SP)
    └── <Dashboard>.html
```

### Convenciones

| Convención | Regla |
|---|---|
| **Código compartido** | En `Funciones/`, una carpeta por lado: `ApiDesposte/Funciones/` (C#) y `Docu_Paginas/Funciones/` (JS/CSS). |
| **Controladores por fuente** | `Controllers/Excel/`, `Controllers/Sql/`, `Controllers/Portal/`. Cada controlador es delgado: resuelve parámetros, llama a `Funciones/` y responde. |
| **Namespaces** | `ApiDesposte.Controllers.<Carpeta>` y `ApiDesposte.Funciones.<Área>`. |
| **Ruta canónica** | Un solo `[Route("api/[controller]")]` por controlador y un solo verbo por acción. El archivo se llama igual que la clase. |
| **Datos tal cual** | Los endpoints devuelven las filas del SP o de la tabla Excel sin columnas calculadas. Los cálculos se hacen en el dashboard. |
| **Datos de ejecución** | `ApiDesposte/Datos/`, nunca dentro de `Docu_Paginas/`. |
| **Secretos** | Variables de entorno de Windows (usuario): `ConnectionStrings__SqlDesposte`, `AdminConfig__Pin`, `DEEPSEEK_API_KEY`, `GEMINI_API_KEY`. Nunca en `appsettings.json` ni en el código. |

---

## 2. Arranque de la API (`ApiDesposte/Program.cs`)

```
AddControllers · AddHttpClient · AddCors("PermitirExcelWeb")
Aviso en consola si falta alguna variable de entorno
──────────────────────────────────────
UseCors("PermitirExcelWeb")
Protección de páginas: un .html marcado false en Datos/dashboards_estado.json
  redirige a /index.html?bloqueado=… salvo para el administrador (cookie desposte_admin o ?admin=PIN)
UseDefaultFiles(Docu_Paginas/)      "/" → "/index.html"
UseStaticFiles(Docu_Paginas/)
MapControllers()
MapGet("/api/health")               { status, message, servidor, timestamp }
Puertos: escucha en todos los libres de 8080, 5000, 5050, 8085
```

La carpeta web se busca en `../Docu_Paginas` y, si no existe, en `./Docu_Paginas`.

---

## 3. Funciones del backend (`ApiDesposte/Funciones/`)

### 3.1 `Excel/LectorTablasExcel.cs` — `ApiDesposte.Funciones.Excel`

| Miembro | Qué hace |
|---|---|
| `LectorTablasExcel.AbrirEnMemoria(ruta)` | Copia el archivo a memoria con `FileShare.ReadWrite \| FileShare.Delete`, elimina tablas dinámicas y devuelve el `XLWorkbook`. |
| `LectorTablasExcel.LeerTablaTalCual(wb, hoja, tabla)` | Lee una tabla estructurada con las columnas de su encabezado (fallback al rango usado). |
| `LectorTablasExcel.ObtenerValorCelda(celda)` | Convierte la celda a un valor JSON simple. |
| `new CacheTablasExcel(TimeSpan)` + `.Obtener(ruta, mapeo, forzarRecarga)` | Caché en memoria de las tablas de un libro. Se relee si cambia la fecha de modificación del archivo. |

### 3.1.1 `Excel/EditorTablasExcel.cs` — CRUD sobre tablas de Excel

| Miembro | Qué hace |
|---|---|
| `EditorTablasExcel.Leer(ruta, hoja, tabla)` | Columnas (con marca de columna calculada) y registros; cada registro lleva `_fila` (posición en la tabla). |
| `EditorTablasExcel.Insertar(ruta, respaldos, hoja, tabla, valores)` | Agrega una fila al final, copia formato y fórmulas de columnas calculadas. |
| `EditorTablasExcel.Actualizar(ruta, respaldos, hoja, tabla, fila, valores, original)` | Escribe sólo las celdas que cambiaron; ignora columnas calculadas. 409 si `original` ya no coincide. |
| `EditorTablasExcel.Eliminar(ruta, respaldos, hoja, tabla, fila, original)` | Elimina la fila (desplaza sólo las columnas de la tabla). 409 si `original` ya no coincide. |
| `EditorTablasExcel.Lote(ruta, respaldos, hoja, tabla, crear, actualizar, eliminar)` | Varias operaciones en una sola escritura, todo o nada (actualiza, elimina de abajo hacia arriba y luego crea). Es lo que usan los formularios. |

Cada escritura: respaldo en `Datos/RespaldosExcel/`, recálculo de fórmulas antes de guardar (ClosedXML no conserva los valores calculados) y restauración de las fórmulas de columna calculada que ClosedXML descarta. No usar `AbrirEnMemoria` para escribir: quita las tablas dinámicas.

La ruta se resuelve en cada controlador con `ObtenerRutaExcel()` a partir del perfil del usuario (regla 3).

### 3.2 `Sql/EjecutorSql.cs` — `ApiDesposte.Funciones.Sql`

| Miembro | Qué hace |
|---|---|
| `EjecutorSql.TimeoutPorDefecto` | `1200` segundos (20 min). |
| `EjecutorSql.EjecutarTalCualAsync(cadena, sql, parametros, commandTimeout)` | Ejecuta con Dapper (parámetros seguros) y devuelve cada fila como `Dictionary<string, object?>` en el orden de SQL Server. |

```csharp
var datos = await EjecutorSql.EjecutarTalCualAsync(
    _connectionString,
    "EXEC Desposte.dbo.Sp_Informe_Marinado_Congelado @Fecha1, @Fecha2",
    new { Fecha1 = f1, Fecha2 = f2 },
    commandTimeout: 2800);
```

Los controladores que mapean a un DTO tipado (`ProduccionCabezas`, `ProdChuletas`, `IndiceRotacionSemanal`) siguen usando `QueryAsync<T>`, porque el DTO define los nombres de las columnas en el JSON.

### 3.3 `Fechas/ParametrosFecha.cs` — `ApiDesposte.Funciones.Fechas`

| Miembro | Qué hace |
|---|---|
| `TryParsear(input, defecto, out fecha)` | Estricto: `yyyy-MM-dd`, `yyyy-MM-ddTHH:mm:ss`, `dd/MM/yyyy`. Vacío → `defecto`. `false` si es inválido (responder 400). |
| `FormatearParaSp(input, defecto, esFechaFinal)` | Flexible: devuelve el texto para el SP (`…T00:00:00` / `…T23:59:59`). Acepta varios formatos y cultura es-PE; nunca rechaza. |
| `Primero(a, b, …)` | Primer valor no vacío (alias de parámetros como `fechaInicio` / `fecha1`). |

### 3.4 `IA/` — `ApiDesposte.Funciones.IA`

| Miembro | Qué hace |
|---|---|
| `GeneradorJsonIA` (singleton) | `GenerarAsync(promptSistema, mensaje, procesar, tiempoMaximo, ct)`: prueba los proveedores configurados en orden y devuelve el primer JSON válido con `"Proveedor · modelo"`. `Activos()` lista los que tienen clave. |
| `ProveedorDeepSeek` | `deepseek-chat` en modo JSON (`DEEPSEEK_API_KEY`). Reintenta ante 429/5xx/timeout. 402 = sin saldo. |
| `ProveedorGemini` | Último Flash estable vía ListModels (`GEMINI_API_KEY`). Reintenta y cambia de modelo ante saturación. |

Orden por defecto: DeepSeek → Gemini. Se cambia con la configuración `IA:Orden` (variable de entorno `IA__Orden`, ej. `gemini,deepseek`). Modelo de DeepSeek opcional con `DeepSeek__Modelo`.

---

## 4. Endpoints de la API

| Ruta | Fuente | Dashboard |
|---|---|---|
| `GET /api/health` | — | todos (detección del servidor) |
| `GET /api/PortalConfig/config` · `POST /api/PortalConfig/verificar-pin` · `POST /api/PortalConfig/config` | `Datos/dashboards_estado.json` | `index.html`, `Funciones/config.js` |
| `GET /api/Sps/bases` · `GET /api/Sps/lista?base=` · `GET /api/Sps/diagrama?nombre=&base=&forzar=` | SQL Server (bases de `Sps__Bases`, por defecto Desposte y PlantasCore) + IA (DeepSeek / Gemini) | `Documentacion_SPs.html`, `visor_sp.html` |
| `GET /api/CostosProduccion` | SQL | `Informe_Costo_Produccion.html` |
| `GET /api/IndiceRotacionSemanal` | SQL | `Indice_Rotacion_Semanal.html` |
| `GET /api/KardexUnCerdos/PTC` · `/PBC` | SQL | `Kardex_Cerdos.html` |
| `GET /api/Marinado` | SQL | `InformeMarinado.html` |
| `GET /api/ProdCerdos` | SQL | — |
| `GET /api/ProdChuletas` | SQL | `RendimientosChuletas002.html` |
| `GET /api/ProduccionCabezas` | SQL | `InformeCabezas.html` |
| `GET /api/ComparativoPresupuesto/datos-comparativo` | Excel | `Presupuesto_Vs_ForeCast_Dark.html` |
| `GET /api/CumplimientoPtcDiario` · `/tablas/{tabla}` · `/subplantas` | Excel | `Comparativo_PTC_Diario.html`, `Brechas_SKU_PTC.html` |
| `GET /api/CumplimientoPtcTablas` · `/{tabla}` | Excel | `ResumenPl_PTC_Detalle.html` |
| `GET /api/CumplimientosEmbutidos` | Excel | `CumplimientosEmbutidos.html` |
| `GET /api/CumplimientosJorge` · `/{tabla}` | Excel | `Resumen_Pl_9Semanas*.html`, `Balance_Escenarios_MP.html` |
| `GET /api/IndicadoresN3` | Excel | `InformeN3.html` |
| `GET /api/InformeDiario` | Excel | `InformeDiario.html` |
| `GET /api/InformeDiarioCrud` · `GET /{tabla}` · `POST /{tabla}` · `PUT /{tabla}/{fila}` · `DELETE /{tabla}/{fila}` · `POST /{tabla}/lote` | Excel (lectura y escritura) | `CapturaInformeDiario.html` |
| `GET /api/NumeroCerdos` | Excel | `NumeroCerdos.html` |
| `GET /api/PersonalExcel/listar` · `POST crear` | Excel | `Personal_Prueba.html` |
| `GET /api/Presupuesto/dashboard-data` | Excel | `Presupuestos_Final.html` |
| `GET /api/ResumenPl` · `/grupos` | Excel | `ResumenPl_PTC.html` |

---

## 5. Funciones del frontend (`Docu_Paginas/Funciones/`)

| Archivo | Se enlaza así | Expone |
|---|---|---|
| `config.js` | `<script src="Funciones/config.js"></script>` | `GLOBAL_CONFIG { servidor, puertos }` y `generarBasesApiGlobal()`. |
| `responsive.css` | `<link rel="stylesheet" href="Funciones/responsive.css?v=3">` | Densidad laptop `≤ 1600px`, tablet `≤ 980px`, celular `≤ 768px`. |

Cada página mantiene además su bloque `APP_CONFIG` y su `@media` autónomo, para funcionar aunque se comparta sola (reglas 1 y 4).

---

## 6. Pendientes respecto a la estructura de referencia

| Pendiente | Detalle |
|---|---|
| **Login con usuarios y roles** | Hoy: PIN de administrador + `Datos/dashboards_estado.json`. La referencia usa `login.html`, `admin.html`, `sesion.js` y `Datos/seguridad.json`. Se dejó para una tarea aparte porque cambia el acceso de todos los usuarios. |
| **JS repetido en cada dashboard** | `APP_CONFIG` + `generarBasesApi()` + `fetchConTimeout()` + `detectarServidorApi()`, cálculo de semanas, formatos y `exportarDataPlanaSQL()` podrían ir a `Funciones/conexion.js`, `semanas.js`, `utilidades.js` y `exportar.js`. |
