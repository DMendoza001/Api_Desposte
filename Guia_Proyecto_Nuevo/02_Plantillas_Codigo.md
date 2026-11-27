# 2. Plantillas de código

En los ejemplos, `ApiPlanta` es el nombre del proyecto .NET (`$Api` en el paso a paso). Reemplazarlo por el nombre real.

---

## 2.1 `Program.cs`

Hace cuatro cosas: habilita CORS, sirve `Docu_Paginas` como sitio web, expone `/api/health` y escucha en los primeros puertos libres (8080, 5000, 5050, 8085).

```csharp
using Microsoft.Extensions.FileProviders;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();

// CORS: permite que los dashboards (y Excel Web) consuman la API
builder.Services.AddCors(options =>
{
    options.AddPolicy("PermitirDashboards", policy =>
        policy.AllowAnyOrigin().AllowAnyHeader().AllowAnyMethod());
});

var app = builder.Build();
app.UseCors("PermitirDashboards");

// Sirve los dashboards de Docu_Paginas (http://servidor:puerto/ abre index.html)
string rutaWeb = Path.GetFullPath(Path.Combine(builder.Environment.ContentRootPath, "..", "Docu_Paginas"));
if (!Directory.Exists(rutaWeb))
{
    rutaWeb = Path.GetFullPath(Path.Combine(builder.Environment.ContentRootPath, "Docu_Paginas"));
}

if (Directory.Exists(rutaWeb))
{
    var fileProvider = new PhysicalFileProvider(rutaWeb);
    app.UseDefaultFiles(new DefaultFilesOptions { FileProvider = fileProvider, RequestPath = "" });
    app.UseStaticFiles(new StaticFileOptions { FileProvider = fileProvider, RequestPath = "" });
    Console.WriteLine($"🌐 Servidor web de dashboards habilitado desde: {rutaWeb}");
}
else
{
    Console.WriteLine($"⚠️ No se encontró la carpeta de páginas en: {rutaWeb}");
}

app.MapControllers();

// Ping que usan los dashboards para detectar la API (detectarServidorApi)
app.MapGet("/api/health", () => Results.Ok(new
{
    status = "online",
    servidor = Environment.MachineName,
    timestamp = DateTime.Now
}));

// Escucha en todos los puertos candidatos que estén libres.
// Se consulta la lista real de puertos en escucha del sistema: probar con un
// TcpListener no es confiable en Windows (puede "enlazar" un puerto ya ocupado
// por otra API y luego Kestrel falla al arrancar).
var enUso = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties()
    .GetActiveTcpListeners().Select(ep => ep.Port).ToHashSet();

int[] puertos = { 8080, 5000, 5050, 8085 };
foreach (int p in puertos)
{
    if (enUso.Contains(p))
    {
        Console.WriteLine($"⚠️ Puerto {p} ocupado.");
        continue;
    }
    app.Urls.Add($"http://0.0.0.0:{p}");
    Console.WriteLine($"✅ Escuchando en http://0.0.0.0:{p}");
}

if (app.Urls.Count == 0)
{
    Console.WriteLine("❌ Los puertos 8080, 5000, 5050 y 8085 están ocupados. ¿Ya hay otra instancia de la API corriendo?");
    return;
}

app.Run();
```

> Si se necesita un portal con dashboards "en prueba" y acceso de administrador (como `Api_Desposte`), agregar ese middleware después. El PIN debe leerse de `appsettings.json` y no escribirse en el código.

---

## 2.2 `appsettings.json`

Usar autenticación de Windows (`Integrated Security`) para no guardar contraseñas en el repositorio.

```json
{
  "ConnectionStrings": {
    "SqlPrincipal": "Server=<SERVIDOR_SQL>;Database=<BASE_DE_DATOS>;Integrated Security=True;TrustServerCertificate=True;"
  },
  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "Microsoft.AspNetCore": "Warning"
    }
  },
  "AllowedHosts": "*"
}
```

---

## 2.3 Controlador SQL Server (`Controllers/Sql/`)

Ejecuta un procedimiento almacenado con rango de fechas y devuelve los datos **tal cual** (el dashboard los guarda en `rawSqlData` para la exportación de data plana, regla 7).

```csharp
using Dapper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace ApiPlanta.Controllers.Sql
{
    [ApiController]
    [Route("api/[controller]")]          // regla 6: una sola ruta canónica
    public class EjemploSqlController : ControllerBase
    {
        private readonly string _connectionString;

        public EjemploSqlController(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("SqlPrincipal")
                ?? throw new InvalidOperationException("Falta la cadena de conexión 'SqlPrincipal' en appsettings.json");
        }

        [HttpGet]                         // regla 6: un solo atributo de verbo por acción
        public async Task<IActionResult> Obtener([FromQuery] string fechaInicio, [FromQuery] string fechaFin)
        {
            try
            {
                await using var db = new SqlConnection(_connectionString);
                var datos = await db.QueryAsync(
                    "EXEC dbo.Sp_Ejemplo @Fecha1, @Fecha2",
                    new { Fecha1 = $"{fechaInicio}T00:00:00", Fecha2 = $"{fechaFin}T23:59:59" },
                    commandTimeout: 600);

                return Ok(new { exito = true, datos });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }
    }
}
```

---

## 2.4 Controlador Excel en OneDrive (`Controllers/Excel/`)

Controlador delgado (regla 8): solo resuelve la ruta, declara qué tablas lee y responde. La apertura, la limpieza y la caché las hace `LectorTablasExcel` / `CacheTablasExcel`.

```csharp
using ApiPlanta.Funciones.Excel;
using Microsoft.AspNetCore.Mvc;

namespace ApiPlanta.Controllers.Excel
{
    [ApiController]
    [Route("api/[controller]")]
    public class EjemploExcelController : ControllerBase
    {
        private static readonly CacheTablasExcel Cache = new(TimeSpan.FromSeconds(60));

        // Clave de respuesta -> (Hoja, Tabla estructurada de Excel)
        private static readonly Dictionary<string, (string Hoja, string Tabla)> MapeoTablas = new()
        {
            ["registro"] = ("Registro", "tblRegistro")
        };

        // Regla 3: nunca escribir el usuario en la ruta; se resuelve con el perfil actual
        private static string ObtenerRutaExcel()
        {
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return Path.Combine(userProfile, @"OneDrive - <Empresa>\<Carpeta>\<Archivo>.xlsx");
        }

        [HttpGet]
        public IActionResult Obtener([FromQuery] bool recargar = false)
        {
            try
            {
                var tablas = Cache.Obtener(ObtenerRutaExcel(), MapeoTablas, recargar);
                return Ok(tablas["registro"]);
            }
            catch (FileNotFoundException ex)
            {
                return NotFound(new { error = ex.Message });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message });
            }
        }
    }
}
```

---

## 2.5 Dashboard HTML (`Docu_Paginas/`)

Esqueleto mínimo que ya cumple las reglas de conexión (1), responsive (4) y exportación de data plana (7). Usa la paleta `PaletaFinanza_Industrial.json`.

```html
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Ejemplo - Dashboard</title>

  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap" rel="stylesheet">
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
  <!-- Regla 7: SheetJS para exportar la data plana -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <!-- Regla 4: estilos responsive compartidos -->
  <link rel="stylesheet" href="Funciones/responsive.css?v=3">
  <!-- Regla 1: configuración compartida de la API -->
  <script src="Funciones/config.js"></script>

  <style>
    :root {
      --color-header-gradient: linear-gradient(135deg, #38353B 0%, #434B57 100%);
      --color-header-subtext: #E2D5B5;
      --color-heading: #38353B;
      --color-text: #4A474C;
      --color-text-muted: #787479;
      --color-primary: #BAA063;
      --color-bg: #F2F4F8;
      --color-surface: #FFFFFF;
      --color-border: #E2E8F0;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Inter', sans-serif; }
    body { background: var(--color-bg); color: var(--color-text); padding: 16px; }
    .header-banner { background: var(--color-header-gradient); color: #fff; padding: 18px 22px; border-radius: 12px;
                     border-bottom: 3px solid var(--color-primary); display: flex; justify-content: space-between;
                     align-items: center; gap: 12px; flex-wrap: wrap; margin-bottom: 16px; }
    .header-banner p { color: var(--color-header-subtext); font-size: .85rem; }
    .kpi-container { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px; margin-bottom: 16px; }
    .kpi-card, .chart-card { background: var(--color-surface); border: 1px solid var(--color-border); border-radius: 10px;
                             padding: 14px 16px; min-width: 0; }
    .chart-box { position: relative; width: 100%; min-width: 0; height: 320px; }
    .btn { background: var(--color-surface); border: 1px solid var(--color-border); border-radius: 8px;
           padding: 7px 12px; cursor: pointer; font-weight: 600; }

    /* Regla 4: bloque @media autónomo */
    @media screen and (max-width: 1600px) { body { padding: 12px; } .chart-box { height: 280px; } }
    @media screen and (max-width: 980px)  { .kpi-container { grid-template-columns: repeat(2, minmax(0, 1fr)); } }
    @media screen and (max-width: 768px)  { .kpi-container { grid-template-columns: 1fr; } .chart-box { height: 240px; } }
  </style>
</head>
<body>
  <header class="header-banner">
    <div>
      <h1>Ejemplo - Dashboard</h1>
      <p id="estado">Detectando API...</p>
    </div>
    <button class="btn" onclick="exportarDataPlanaSQL()">📄 Data Plana SQL (.xlsx)</button>
  </header>

  <section class="kpi-container">
    <div class="kpi-card"><div>Registros</div><strong id="kpiRegistros">0</strong></div>
    <div class="kpi-card"><div>KPI 2</div><strong>-</strong></div>
    <div class="kpi-card"><div>KPI 3</div><strong>-</strong></div>
    <div class="kpi-card"><div>KPI 4</div><strong>-</strong></div>
  </section>

  <section class="chart-card">
    <div class="chart-box"><canvas id="grafico"></canvas></div>
  </section>

  <script>
    // Regla 1: bloque estándar APP_CONFIG con fallback
    const APP_CONFIG = {
      servidor: (typeof GLOBAL_CONFIG !== 'undefined' && GLOBAL_CONFIG.servidor) ? GLOBAL_CONFIG.servidor : "<SERVIDOR>",
      puertos: (typeof GLOBAL_CONFIG !== 'undefined' && GLOBAL_CONFIG.puertos) ? GLOBAL_CONFIG.puertos : [8080, 5000],
      controllerPath: "/api/EjemploSql",
      timeoutMs: 120000
    };

    function generarBasesApi() {
      if (typeof generarBasesApiGlobal === 'function') return generarBasesApiGlobal();
      const bases = [];
      if (window.location.protocol.startsWith('http')) bases.push(window.location.origin);
      APP_CONFIG.puertos.forEach(p => bases.push(`http://${APP_CONFIG.servidor}:${p}`));
      return bases;
    }

    async function fetchConTimeout(url, timeoutMs) {
      const ctrl = new AbortController();
      const id = setTimeout(() => ctrl.abort(), timeoutMs);
      try { return await fetch(url, { signal: ctrl.signal }); }
      finally { clearTimeout(id); }
    }

    async function detectarServidorApi() {
      for (const base of generarBasesApi()) {
        try { if ((await fetchConTimeout(`${base}/api/health`, 3500)).ok) return base; } catch (e) { }
        try { if ((await fetchConTimeout(`${base}/`, 3500)).ok) return base; } catch (e) { }
      }
      return null;
    }

    // Regla 7: datos tal cual vienen de SQL Server, sin filtros
    let rawSqlData = [];
    let fechaIni = '', fechaFin = '';

    async function cargar() {
      const base = await detectarServidorApi();
      const estado = document.getElementById('estado');
      if (!base) { estado.textContent = `Sin conexión con ${APP_CONFIG.servidor}`; return; }

      const hoy = new Date();
      const ini = new Date(hoy); ini.setDate(hoy.getDate() - 30);
      fechaIni = ini.toISOString().slice(0, 10);
      fechaFin = hoy.toISOString().slice(0, 10);

      const res = await fetchConTimeout(`${base}${APP_CONFIG.controllerPath}?fechaInicio=${fechaIni}&fechaFin=${fechaFin}`, APP_CONFIG.timeoutMs);
      const json = await res.json();
      rawSqlData = Array.isArray(json) ? json : (json.datos || []);

      estado.textContent = `En línea · ${rawSqlData.length.toLocaleString('es-PE')} registros`;
      document.getElementById('kpiRegistros').textContent = rawSqlData.length.toLocaleString('es-PE');
      // Aquí: calcular KPIs, gráficos y matriz a partir de rawSqlData
    }

    function exportarDataPlanaSQL() {
      if (!rawSqlData.length) { alert('No hay datos cargados para exportar.'); return; }
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rawSqlData), 'Data');
      XLSX.writeFile(wb, `Data_Plana_SQL_Ejemplo_${fechaIni}_al_${fechaFin}.xlsx`);
    }

    window.addEventListener('DOMContentLoaded', cargar);
  </script>
</body>
</html>
```

Si el dashboard tiene vistas Diario / Semanal / Mensual, aplicar además la regla 9 (filtros de periodo). La implementación de referencia está en `Api_Desposte/Docu_Paginas/RendimientosChuletas002.html`.
