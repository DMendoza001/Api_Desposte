# Estructura del proyecto y funciones compartidas

Guía de referencia de **Informe_html** (API PTC en .NET + dashboards HTML) para replicar su organización en otro proyecto o usarla como objetivo de una refactorización.

- **Sección 1:** árbol de carpetas y qué va en cada una.
- **Secciones 2 a 6:** las piezas reutilizables, con su firma y cómo usarlas.
- **Sección 7:** pasos para llevar la estructura a otro proyecto.
- **Sección 8:** código que hoy está duplicado y conviene mover a `Funciones/`.

---

## 1. Árbol de carpetas

```
Informe_html/
├── CLAUDE.md                     ← reglas del proyecto para Claude Code (resumen de cada regla)
├── AGENTS.md                     ← mismas reglas para otros agentes (siempre sincronizado con CLAUDE.md)
├── ESTRUCTURA_PROYECTO.md        ← este documento
├── PaletaDMendoza_Azules.json    ← paleta de colores oficial de los dashboards
│
├── .agents/rules/                ← detalle de cada regla (CLAUDE.md y AGENTS.md enlazan aquí)
│   ├── api-connection.md            conexión frontend → API (APP_CONFIG, detección de servidor)
│   ├── calculo-semanas.md           semanas ISO-8601, lunes a domingo, etiquetas "Sem XX"
│   ├── conexion-excel-onedrive.md   rutas dinámicas de OneDrive + LectorTablasExcel
│   ├── responsive-design.md         breakpoints 1600 / 980 / 768, min-width: 0
│   ├── exportar-data-plana-sql.md   botón "Data Plana SQL (.xlsx)" con SheetJS
│   ├── filtros-periodo-vistas.md    filtros por vista Diario / Semanal / Mensual
│   └── autenticacion-roles.md       login, usuarios, roles y permisos sin base de datos
│
├── .claude/skills/barralateral/  ← skill de Claude Code: barra lateral reutilizable (solo este proyecto)
│   ├── SKILL.md
│   └── assets/ (barra-lateral.css · .html · .js · ejemplo.html)
│
├── Api_PTC/                      ← API .NET 10 (proyecto Informe_html.csproj, namespace Informe_html)
│   ├── Program.cs                   arranque: puerto 80, CORS, sesión, archivos estáticos, /api/health
│   ├── appsettings.json             cadena de conexión SqlDesposte (sin PIN ni contraseñas de usuarios)
│   ├── Informe_html.http            peticiones de prueba
│   ├── Controllers/                 un subdirectorio por tipo de fuente de datos
│   │   ├── Excel/                      lectura de Excel de OneDrive   (BonosController)
│   │   ├── Sql/                        procedimientos de SQL Server    (HorasAnualController)
│   │   └── Seguridad/                  sesión y administración         (AuthController, SeguridadController)
│   ├── Funciones/                   ← código reutilizable del backend (namespace Informe_html.Funciones.<Área>)
│   │   ├── Excel/LectorTablasExcel.cs   apertura segura de Excel + lectura de tablas + caché
│   │   ├── Sql/EjecutorSql.cs           ejecución de SP devolviendo filas tal cual (timeout 20 min)
│   │   ├── Fechas/ParametrosFecha.cs    lectura de fechas de la query string
│   │   └── Seguridad/                   login, usuarios, roles y control de acceso
│   │       ├── AlmacenSeguridad.cs
│   │       ├── ClaveSegura.cs
│   │       └── ProteccionAcceso.cs
│   └── Datos/                       ← datos de ejecución, FUERA de la web y excluidos de la publicación
│       ├── seguridad.json              usuarios, roles y catálogo de dashboards
│       └── claves/                     claves de cifrado de la cookie de sesión
│
└── Dashboard/                    ← carpeta web que sirve la API (http://RAEVSALL001/)
    ├── Funciones/                   ← código reutilizable del frontend (enlazado desde cada HTML)
    │   ├── config.js                   servidor y puertos de la API
    │   ├── responsive.css              estilos responsive compartidos
    │   ├── sesion.js                   sesión del usuario, permisos en la interfaz
    │   ├── barra-lateral.css           barra lateral smart flotante (estilos)
    │   └── barra-lateral.js            barra lateral smart flotante (API BarraLateral)
    ├── login.html                   inicio de sesión, cambio de contraseña, solicitud de cuenta
    ├── index.html                   portal: dashboards permitidos al usuario según su rol
    ├── admin.html                   administración de usuarios, roles y catálogo
    ├── Informe_Horas_Anual.html     dashboard (fuente SQL)
    └── Pago_Bonos_Semanal.html      dashboard (fuente Excel)
```

### Convenciones de organización

| Convención | Regla |
|---|---|
| **Código compartido** | Va en una carpeta `Funciones/`, **una por lado**: `Api_PTC/Funciones/` (C#) y `Dashboard/Funciones/` (JS/CSS). C# y JavaScript no comparten archivos. |
| **Controladores por fuente** | `Controllers/Excel/`, `Controllers/Sql/`, `Controllers/Seguridad/`. Cada controlador es delgado: resuelve parámetros, llama a `Funciones/` y responde. |
| **Namespaces** | `Informe_html.Controllers.<Carpeta>` y `Informe_html.Funciones.<Área>`. |
| **Ruta canónica** | Un solo `[Route("api/[controller]")]` por controlador y un solo verbo por acción. La clase se llama igual que el archivo. |
| **Respuesta de la API** | `{ exito, ..., datos }` en éxito y `{ exito: false, error }` en fallo (400 / 401 / 403 / 404 / 500). |
| **Datos tal cual** | Los endpoints devuelven las filas del SP o de la tabla Excel **sin columnas calculadas**. Los cálculos se hacen en el dashboard. |
| **Datos de ejecución** | `Api_PTC/Datos/`: nunca dentro de `Dashboard/` (se publicaría en la web). |
| **Reglas** | Resumen en `CLAUDE.md` **y** `AGENTS.md` (sincronizados), detalle en `.agents/rules/<regla>.md`. |

---

## 2. Arranque de la API (`Api_PTC/Program.cs`)

Orden del *pipeline*. El orden importa: la protección debe ir **después** de `UseDefaultFiles` (para ver `/index.html`) y **antes** de `UseStaticFiles` (para bloquear páginas).

```
UseUrls("http://0.0.0.0:{PuertoApi}")   puerto 80 (configurable para pruebas: --PuertoApi=5099)
AddSingleton(AlmacenSeguridad)          Datos/seguridad.json
AddDataProtection → Datos/claves        las sesiones sobreviven a reinicios
AddAuthentication().AddCookie           cookie "ptc_sesion", HttpOnly, SameSite=Lax, 10 h deslizantes
──────────────────────────────────────
UseCors("PermitirExcelWeb")
UseAuthentication()
UseDefaultFiles(Dashboard/)             "/" → "/index.html"
UseProteccionAcceso()                   sesión y permisos por rol (páginas y /api)
UseAuthorization()
UseStaticFiles(Dashboard/)              no-store en .html
MapControllers()
MapGet("/api/health")                   público: { status, message, servidor = Environment.MachineName, timestamp }
```

- La carpeta web se busca en `../Dashboard` y, si no existe, en `./Dashboard` relativa al *content root*.
- `Informe_html.csproj` excluye `Datos\**` de la compilación y la publicación.

---

## 3. Funciones del backend (`Api_PTC/Funciones/`)

### 3.1 `Excel/LectorTablasExcel.cs` — namespace `Informe_html.Funciones.Excel`

Lectura de libros `.xlsx` / `.xlsm` sincronizados en OneDrive sin bloquear el archivo.

| Miembro | Firma | Qué hace |
|---|---|---|
| `LectorTablasExcel.AbrirEnMemoria` | `XLWorkbook AbrirEnMemoria(string rutaExcel)` | Copia el archivo a memoria con `FileShare.ReadWrite \| FileShare.Delete` (no bloquea a quien lo tenga abierto), elimina tablas dinámicas (evitan errores de ClosedXML) y devuelve el libro. |
| `LectorTablasExcel.LeerTablaTalCual` | `List<Dictionary<string, object?>> LeerTablaTalCual(XLWorkbook wb, string hoja, string tabla)` | Lee una tabla estructurada con las columnas de su encabezado. Si no la halla en la hoja, la busca en todo el libro; si no hay tabla formal, usa el rango usado de la hoja. Omite filas vacías. Lanza `InvalidOperationException` si no la encuentra. |
| `LectorTablasExcel.ObtenerValorCelda` | `object? ObtenerValorCelda(IXLCell celda)` | Convierte la celda a JSON simple: `bool`, `long` (enteros), `double` (4 decimales), fecha `yyyy-MM-dd` (o con hora), hora `hh:mm:ss`, texto o `null`. |
| `CacheTablasExcel` (constructor) | `new CacheTablasExcel(TimeSpan duracion)` | Caché en memoria de un conjunto de tablas de un libro. Se declara `static readonly` en el controlador. |
| `CacheTablasExcel.Obtener` | `Dictionary<string, List<...>> Obtener(string ruta, IReadOnlyDictionary<string,(string Hoja,string Tabla)> mapeo, bool forzarRecarga = false)` | Devuelve todas las tablas del mapeo, releyendo el archivo solo si venció la caché o `forzarRecarga = true`. Lanza `FileNotFoundException` si no existe el archivo. |

**Uso (controlador delgado):**
```csharp
private const string RutaRelativaExcel = @"OneDrive - Corporación Rico SAC\Desposte-03\Bonos\ExtareBonos.xlsm";
private static readonly Dictionary<string, (string Hoja, string Tabla)> MapeoTablas = new(StringComparer.OrdinalIgnoreCase)
{
    { "T_Bonos", ("Bonos", "T_Bonos") }
};
private static readonly CacheTablasExcel Cache = new(TimeSpan.FromSeconds(60));

private string ObtenerRutaExcel() =>
    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), RutaRelativaExcel);

[HttpGet]
public IActionResult ObtenerBonos([FromQuery] bool recargar = false)
{
    var datos = Cache.Obtener(ObtenerRutaExcel(), MapeoTablas, recargar)["T_Bonos"];
    return Ok(new { exito = true, totalRegistros = datos.Count, datos });
}
```
Nunca escribir el usuario de Windows en la ruta (`C:\Users\fulano\...`): se resuelve con el perfil del usuario que ejecuta la API.

### 3.2 `Sql/EjecutorSql.cs` — namespace `Informe_html.Funciones.Sql`

| Miembro | Firma | Qué hace |
|---|---|---|
| `EjecutorSql.TimeoutPorDefecto` | `const int = 1200` | 20 minutos de espera por SP. El dashboard debe usar `APP_CONFIG.timeoutMs = 1200000`. |
| `EjecutorSql.EjecutarTalCualAsync` | `Task<List<Dictionary<string, object?>>> EjecutarTalCualAsync(string connectionString, string sql, object? parametros = null, int commandTimeout = 1200)` | Ejecuta la consulta con Dapper (parámetros seguros, sin concatenar) y devuelve cada fila como diccionario columna → valor, en el orden de SQL Server. |

```csharp
var datos = await EjecutorSql.EjecutarTalCualAsync(
    _connectionString,
    "EXEC Desposte.dbo.Sp_Informe_Horas_Anual @Fecha1, @Fecha2",
    new { Fecha1 = $"{inicio:yyyy-MM-dd}T00:00:00", Fecha2 = $"{fin:yyyy-MM-dd}T23:59:59" });
```
La cadena de conexión se lee con `configuration.GetConnectionString("SqlDesposte")`.

### 3.3 `Fechas/ParametrosFecha.cs` — namespace `Informe_html.Funciones.Fechas`

| Miembro | Firma | Qué hace |
|---|---|---|
| `ParametrosFecha.TryParsear` | `bool TryParsear(string? input, DateTime valorPorDefecto, out DateTime fecha)` | Acepta `yyyy-MM-dd`, `yyyy-MM-ddTHH:mm:ss` y `dd/MM/yyyy`. Vacío → `valorPorDefecto`. Devuelve `false` si el formato es inválido (el controlador responde 400). |

### 3.4 `Seguridad/` — namespace `Informe_html.Funciones.Seguridad`

#### `ClaveSegura.cs`
| Miembro | Qué hace |
|---|---|
| `bool Coincide(string? escrita, string? guardada)` | Compara la contraseña escrita con la guardada (en texto plano, comparación de tiempo constante). |
| `bool VerificarHashAntiguo(string clave, string? almacenado)` | Verifica hashes PBKDF2 del formato anterior, solo para migrarlos a texto plano al iniciar sesión. |
| `string GenerarTemporal(int largo = 12)` | Contraseña aleatoria sin caracteres ambiguos (0/O, 1/l). |
| `string? ValidarPolitica(string? clave)` | `null` si es válida; si no, el motivo. Política: mínimo 8 caracteres, con letras y números. |

> Las contraseñas se guardan **en texto plano** por decisión del proyecto. Restringe quién puede leer `Api_PTC/Datos/` en el servidor y sus respaldos.

#### `AlmacenSeguridad.cs`
Almacén de usuarios, roles y catálogo en `Datos/seguridad.json` (sin base de datos). Se registra como *singleton*. Relee el archivo si cambió en disco (se puede editar a mano con la API encendida) y escribe de forma atómica (archivo temporal + reemplazo).

| Modelo | Campos |
|---|---|
| `RolSeguridad` | `Id`, `Descripcion`, `Dashboards` (archivos HTML; `"*"` = todos), `Acciones` (`exportar`, `administrar`) |
| `UsuarioSeguridad` | `Usuario`, `Nombre`, `Clave`, `Roles`, `Activo`, `DebeCambiarClave`, `Sello` (cambia al cambiar la contraseña → cierra sesiones), `SolicitadoEl` (solicitud desde el login pendiente) |
| `DashboardCatalogo` | `Archivo`, `Titulo`, `Categoria`, `Descripcion`, `Icono`, `Endpoints` (rutas de API que consume) |
| `PermisosUsuario` | `PuedeVer(archivo)`, `Tiene(accion)` |

| Método | Qué hace |
|---|---|
| `ObtenerUsuario(usuario)` | Usuario por nombre (sin distinguir mayúsculas). |
| `Permisos(usuario)` | Une los dashboards y acciones de todos sus roles. |
| `Catalogo()` / `Instantanea()` | Copia del catálogo / de toda la configuración. |
| `ValidarCredenciales(usuario, clave)` | Login. Bloquea 5 min tras 5 intentos fallidos. Informa "pendiente de aprobación" o "desactivado" solo si la contraseña es correcta. Migra hashes antiguos. |
| `RegistrarSolicitud(usuario, nombre, clave)` | Solicitud de cuenta desde el login: queda inactiva y sin roles (máx. 50 pendientes). |
| `CambiarClave(usuario, actual, nueva)` | Valida la actual y la política; renueva el sello. |
| `GuardarUsuario(entrada, esNuevo, editor)` | Crear/editar. Contraseña vacía en un usuario nuevo → temporal (se devuelve una vez). Activar una solicitud la aprueba. |
| `RestablecerClave(usuario)` | Genera una temporal y cierra sus sesiones. |
| `EliminarUsuario(usuario, editor)` | No permite eliminarse a sí mismo. |
| `GuardarRol(rol, esNuevo)` / `EliminarRol(id)` | No elimina roles en uso. |

Regla transversal: **ningún cambio puede dejar el sistema sin un usuario activo con `administrar`**.

Si `seguridad.json` no existe, se crea con los roles Administrador / Supervisor / Consulta y el usuario `admin` con una contraseña temporal que se muestra **una sola vez** en la consola.

#### `ProteccionAcceso.cs`
Middleware `app.UseProteccionAcceso()`:

| Petición | Sin sesión | Con sesión |
|---|---|---|
| `/login.html`, `/Funciones/*`, `/api/health`, `/api/Auth/login`, `/logout`, `/registro` | permitido | permitido |
| Página `.html` | redirige a `/login.html?volver=...` | `index.html`: todos · `admin.html`: solo `administrar` · dashboard del catálogo: según rol · HTML fuera del catálogo: solo `administrar` |
| `/api/Auth/*` | 401 | permitido |
| `/api/Seguridad*` | 401 | solo `administrar` (si no, 403) |
| Otro `/api/*` | 401 | permitido si algún dashboard que el usuario puede ver declara ese endpoint; si ninguno lo declara, solo `administrar` |
| Contraseña temporal | — | solo `/api/Auth/*`; las páginas redirigen a `/login.html?cambiar=1` |

Usuario eliminado, desactivado o con sello distinto → cierra la sesión.

---

## 4. Endpoints de la API

| Método y ruta | Acceso | Respuesta |
|---|---|---|
| `GET /api/health` | público | `{ status, message, servidor, timestamp }` |
| `POST /api/Auth/login` | público | `{ exito, debeCambiarClave }` + cookie |
| `POST /api/Auth/logout` | público | `{ exito }` |
| `POST /api/Auth/registro` | público | solicitud de cuenta pendiente |
| `GET /api/Auth/perfil` | sesión | `{ usuario, nombre, roles, acciones, debeCambiarClave, dashboards[] }` |
| `POST /api/Auth/cambiar-clave` | sesión | `{ exito }` |
| `GET /api/Seguridad` | administrar | `{ usuarios, roles, dashboards, acciones }` |
| `POST /api/Seguridad/usuarios` | administrar | `{ exito, claveTemporal }` |
| `POST /api/Seguridad/usuarios/{usuario}/restablecer-clave` | administrar | `{ exito, claveTemporal }` |
| `DELETE /api/Seguridad/usuarios/{usuario}` | administrar | `{ exito }` |
| `POST /api/Seguridad/roles` · `DELETE /api/Seguridad/roles/{id}` | administrar | `{ exito }` |
| `GET /api/HorasAnual?fechaInicio=&fechaFin=` | según catálogo | `{ exito, parametros, totalRegistros, datos }` (SP `Sp_Informe_Horas_Anual`) |
| `GET /api/Bonos?recargar=` | según catálogo | `{ exito, archivo, hoja, tabla, totalRegistros, datos }` (tabla `T_Bonos`) |

---

## 5. Funciones del frontend (`Dashboard/Funciones/`)

| Archivo | Se enlaza así | Expone |
|---|---|---|
| `config.js` | `<script src="Funciones/config.js"></script>` | `GLOBAL_CONFIG { servidor, puertos }` y `generarBasesApiGlobal()`: lista ordenada de URLs candidatas (el propio origen, localhost/127.0.0.1, el servidor configurado). |
| `responsive.css` | `<link rel="stylesheet" href="Funciones/responsive.css?v=3">` | Densidad laptop `≤ 1600px`, tablet `≤ 980px`, celular `≤ 768px`; `min-width: 0` en hijos de grid/flex. |
| `sesion.js` | `<script src="Funciones/sesion.js"></script>` | `window.Sesion` (ver abajo). |
| `barra-lateral.css` / `.js` | `<link … Funciones/barra-lateral.css>` + `<script … Funciones/barra-lateral.js>` (+ Font Awesome 6) | `window.BarraLateral` (ver abajo). |

### `sesion.js` → `window.Sesion`
| Miembro | Uso |
|---|---|
| `Sesion.listo` | Promesa con el perfil (`GET /api/Auth/perfil`); `null` sin sesión o en `file://`. |
| `Sesion.perfil` | `{ usuario, nombre, roles, acciones, dashboards }`. |
| `Sesion.tiene(accion)` | `true` si el usuario tiene `exportar` / `administrar`. |
| `Sesion.cerrar()` | Cierra la sesión y va a `login.html`. |
| `data-permiso="exportar"` | El elemento se elimina si el usuario no tiene la acción (solo interfaz). |
| `<div data-sesion></div>` | Se rellena con nombre, rol, "Portal" y "Salir". |
| Intercepta `fetch` | Ante un 401 de la API propia redirige a `login.html?volver=...`. |

### `barra-lateral.js` → `window.BarraLateral`
| Miembro | Uso |
|---|---|
| `iniciar({ vistaInicial, alCambiarVista, notificarInicial })` | Conecta los botones `data-vista` y el de pantalla completa (`data-accion="pantalla-completa"`). |
| `activar(vista)` | Cambia de vista desde código. |
| `setBadge(vista, n)` / `setBadges({ vista: n })` | Contadores (`data-badge`). |
| `setEstado('online' \| 'loading' \| 'error' \| '', texto?)` | Punto y texto de estado al pie. |
| `togglePantallaCompleta()` | Entra/sale de pantalla completa. |
| Variables CSS `--sb-*` | Paleta (`--sb-bg-start`, `--sb-bg-end`, `--sb-accent`, `--sb-accent-rgb`, `--sb-on-accent`, `--sb-subtext`, `--sb-subtext-rgb`, `--sb-shadow-rgb`). |

---

## 6. Páginas y plantilla de dashboard

| Página | Rol |
|---|---|
| `login.html` | Iniciar sesión, cambiar contraseña (con o sin sesión), solicitar cuenta. Rutas relativas `/api/Auth/*`. |
| `index.html` | Portal: dirección de la app con el nombre de la máquina, dashboards permitidos agrupados por categoría (del perfil), búsqueda y enlaces para copiar. |
| `admin.html` | Usuarios (ver/cambiar contraseña, aprobar/rechazar solicitudes), roles (CRUD) y catálogo (solo lectura). Los formularios se cierran solo con ✕ / Cancelar. |

### Esqueleto de un dashboard nuevo

```html
<head>
  <link rel="stylesheet" href="Funciones/responsive.css?v=3">
  <script src="Funciones/config.js"></script>
  <script src="Funciones/sesion.js"></script>
  <!-- opcional: barra lateral -->
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <link rel="stylesheet" href="Funciones/barra-lateral.css">
  <script src="Funciones/barra-lateral.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
  <!-- fuente SQL: SheetJS para la Data Plana -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <style> :root { /* paleta */ }  /* + bloques @media 1600 / 980 / 768 propios */ </style>
</head>
<body>
  <header class="app-header"> … <div data-sesion></div>
     <button id="btnExportarPlana" data-permiso="exportar" onclick="exportarDataPlanaSQL()">Data Plana SQL (.xlsx)</button>
  </header>
  …
  <script>
    const APP_CONFIG = { servidor: GLOBAL_CONFIG?.servidor ?? "RAEVSALL001", puertos: GLOBAL_CONFIG?.puertos ?? [80] };
    APP_CONFIG.controllerPath = "/api/MiControlador";
    APP_CONFIG.timeoutMs = 1200000;        // SP: 20 min · Excel: 60000
    // generarBasesApi() · fetchConTimeout() · detectarServidorApi() (usa /api/health)
    // calcularSemanaConInicioLunes() si agrupa por semana
    // rawSqlData = json.datos   ← data plana intacta para exportar
  </script>
</body>
```

Al terminar: registrar el dashboard en `Api_PTC/Datos/seguridad.json` (sección `dashboards`, con sus `endpoints`) y asignarlo a roles desde `admin.html`.

---

## 7. Cómo replicar la estructura en otro proyecto

1. **Crear la API** (`dotnet new webapi`) y las carpetas `Controllers/{Excel,Sql,Seguridad}`, `Funciones/{Excel,Sql,Fechas,Seguridad}` y `Datos/`.
2. **Copiar** `Api_PTC/Funciones/**` y `Controllers/Seguridad/**`. Cambiar el namespace raíz `Informe_html` por el del nuevo proyecto.
3. **Paquetes NuGet:** `ClosedXML`, `Dapper`, `Microsoft.Data.SqlClient` (ver `Informe_html.csproj`).
4. **`Program.cs`:** copiarlo y ajustar el puerto, el nombre de la cookie (`ptc_sesion`) y el nombre de aplicación de DataProtection, para que no choquen con otra app del mismo servidor.
5. **`.csproj`:** copiar el bloque que excluye `Datos\**`.
6. **`appsettings.json`:** solo la cadena de conexión.
7. **Web:** crear `Dashboard/Funciones/` con `config.js` (servidor y puertos del nuevo proyecto), `responsive.css`, `sesion.js` y, si se usa, `barra-lateral.*`. Copiar `login.html`, `index.html` y `admin.html` y cambiar textos y paleta.
8. **Reglas:** copiar `CLAUDE.md`, `AGENTS.md` y `.agents/rules/`. Reemplazar nombres de proyecto, carpetas y puerto.
9. **Skill:** copiar `.claude/skills/barralateral/` si se quiere la barra lateral.
10. **Primer arranque:** la API crea `Datos/seguridad.json` y muestra la contraseña temporal de `admin` en la consola.

> **Dos apps en el mismo servidor:** cada una necesita su propio puerto (80 se escribe sin número; cualquier otro, como `:8080`, hay que escribirlo) y su propio nombre de cookie, porque los navegadores comparten cookies entre puertos del mismo host.

---

## 8. Oportunidades de refactorización detectadas

Código que hoy está **repetido dentro de cada dashboard** y conviene mover a `Dashboard/Funciones/`:

| Hoy (copiado en cada HTML) | Propuesta | Beneficio |
|---|---|---|
| `APP_CONFIG`, `generarBasesApi()`, `fetchConTimeout()`, `detectarServidorApi()` | `Funciones/conexion.js` con `Conexion.obtenerJson(ruta, { timeoutMs })` | Una sola detección del servidor y un solo manejo de errores/timeouts. |
| `calcularSemanaConInicioLunes()`, `etiquetaEje`, `etiquetaCompleta`, `infoSemanaExcel` (AAAASS) | `Funciones/semanas.js` | Regla de semanas ISO en un solo lugar. |
| `mostrarToast()`, `escaparHtml()` / `esc()`, `fmtNum`, `fmtSoles` | `Funciones/utilidades.js` | Formato y mensajes uniformes. |
| `exportarDataPlanaSQL()` | `Funciones/exportar.js` con `exportarDataPlana(datos, modulo, rango)` | Mismo nombre de archivo y validaciones. |
| Opciones base de Chart.js (tooltip, ejes, colores de la paleta) y `etiquetaVisible()` de las etiquetas de datos | `Funciones/graficos.js` | Gráficos consistentes; etiquetas ocultas en celular en todos. |
| Matriz jerárquica (botón expandir, sangría, fila total) | `Funciones/matriz.css` + `matriz.js` | Regla 5 (etiquetas planas) aplicada igual en todas. |
| Variables de paleta `:root` | `Funciones/paleta.css` generado desde `PaletaDMendoza_Azules.json` | Cambiar la paleta en un solo archivo. |

Otros puntos a tener en cuenta:
- `detectarServidorApi()` usa `/api/health` en los dashboards, mientras la regla `api-connection.md` todavía describe el ping al endpoint de datos. Conviene actualizar la regla.
- `login.html`, `admin.html` y `sesion.js` usan rutas relativas `/api/...` (la cookie es del mismo origen); los dashboards usan `APP_CONFIG`. Con `conexion.js` ambos casos quedarían unificados.
- El catálogo de dashboards solo se edita a mano en `seguridad.json`: una pantalla en `admin.html` evitaría editar el JSON.
- `appsettings.json` contiene la contraseña de SQL en texto: en otro proyecto conviene `dotnet user-secrets` o variables de entorno.
