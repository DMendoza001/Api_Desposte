using ApiDesposte.Funciones.IA;
using ApiDesposte.Funciones.Portal;



using Microsoft.Extensions.FileProviders;

var builder = WebApplication.CreateBuilder(args);

// 1. Agregar soporte para Controladores API
builder.Services.AddControllers();

// Cliente HTTP para servicios externos + IA para la documentación de SPs (DeepSeek / Gemini)
builder.Services.AddHttpClient();
builder.Services.AddSingleton<GeneradorJsonIA>();

// 2. Configurar CORS (Para permitir solicitudes desde Excel Web y dashboards)
builder.Services.AddCors(options =>
{
    options.AddPolicy("PermitirExcelWeb", policy =>
    {
        policy.AllowAnyOrigin()
              .AllowAnyHeader()
              .AllowAnyMethod();
    });
});

var app = builder.Build();

// =========================================================================
// 🔐 SECRETOS: se leen de variables de entorno de Windows (no de appsettings.json)
// =========================================================================
// ConnectionStrings__SqlDesposte, AdminConfig__Pin, DEEPSEEK_API_KEY y GEMINI_API_KEY.
// Si falta alguno, la API arranca igual pero se avisa en consola qué variable configurar.
var secretosRequeridos = new (string Clave, string Variable, string Efecto)[]
{
    ("ConnectionStrings:SqlDesposte", "ConnectionStrings__SqlDesposte", "los endpoints SQL no podrán conectarse"),
    ("AdminConfig:Pin", "AdminConfig__Pin", "el acceso de administrador quedará deshabilitado")
};
foreach (var (clave, variable, efecto) in secretosRequeridos)
{
    if (string.IsNullOrWhiteSpace(builder.Configuration[clave]))
    {
        Console.WriteLine($"⚠️ Falta la variable de entorno '{variable}': {efecto}.");
    }
}
var proveedoresIA = app.Services.GetRequiredService<GeneradorJsonIA>().Activos().Select(p => p.Nombre).ToList();
Console.WriteLine(proveedoresIA.Count > 0
    ? $"🤖 IA para documentar SPs: {string.Join(" → ", proveedoresIA)}"
    : "ℹ️ Sin 'DEEPSEEK_API_KEY' ni 'GEMINI_API_KEY': la documentación de SPs solo mostrará diagramas ya guardados.");

// 3. Activar CORS
app.UseCors("PermitirExcelWeb");

// =========================================================================
// 🌐 SERVIR TABLERO Y REPORTES HTML (Docu_Paginas)
// =========================================================================
// Permite acceder por link directo (ej. http://RAEVSALL001:8080/ o http://localhost:8080/)
// y abrir index.html como portal principal de dashboards sin descargar archivos.
string rutaCarpetaWeb = Path.GetFullPath(Path.Combine(builder.Environment.ContentRootPath, "..", "Docu_Paginas"));
if (!Directory.Exists(rutaCarpetaWeb))
{
    rutaCarpetaWeb = Path.GetFullPath(Path.Combine(builder.Environment.ContentRootPath, "Docu_Paginas"));
}

if (Directory.Exists(rutaCarpetaWeb))
{
    var fileProvider = new PhysicalFileProvider(rutaCarpetaWeb);
    string rutaEstado = EstadoDashboards.RutaArchivo(builder.Environment.ContentRootPath);

    // 🔒 Middleware de protección: Intercepta peticiones directas a archivos HTML en prueba
    // Consulta dinámicamente Datos/dashboards_estado.json sin necesidad de listas fijas
    app.Use(async (context, next) =>
    {
        string path = context.Request.Path.Value ?? "";
        if (path.EndsWith(".html", StringComparison.OrdinalIgnoreCase) &&
            !path.Equals("/index.html", StringComparison.OrdinalIgnoreCase))
        {
            string fileName = Path.GetFileName(path);
            // PIN desde la variable de entorno AdminConfig__Pin; sin PIN configurado nadie es administrador
            string adminPin = (builder.Configuration["AdminConfig:Pin"] ?? "").Trim();

            // Verificar si el usuario cuenta con credencial de administrador (Cookie o Query string)
            bool esAdmin = false;
            if (adminPin.Length > 0)
            {
                if (context.Request.Cookies.TryGetValue("desposte_admin", out var cPin) && cPin == adminPin)
                {
                    esAdmin = true;
                }
                else if (context.Request.Query.TryGetValue("admin", out var qPin) && qPin == adminPin)
                {
                    esAdmin = true;
                }
            }

            if (!esAdmin)
            {
                try
                {
                    var estados = EstadoDashboards.Leer(rutaEstado);
                    // Verifica directamente si el archivo HTML está marcado como en prueba (false)
                    if (estados.TryGetValue(fileName, out bool activo) && !activo)
                    {
                        // Redirigir al portal principal con aviso de acceso restringido
                        context.Response.Redirect($"/index.html?bloqueado={Uri.EscapeDataString(fileName)}");
                        return;
                    }
                }
                catch
                {
                    // En caso de error leyendo el archivo json, continuar
                }
            }
        }

        await next();
    });

    // Permite que http://servidor:puerto/ abra automáticamente index.html
    app.UseDefaultFiles(new DefaultFilesOptions
    {
        FileProvider = fileProvider,
        RequestPath = ""
    });

    // Sirve todos los archivos .html, .js, .css de Docu_Paginas
    app.UseStaticFiles(new StaticFileOptions
    {
        FileProvider = fileProvider,
        RequestPath = ""
    });

    Console.WriteLine($"🌐 Servidor web de dashboards habilitado desde: {rutaCarpetaWeb}");
}
else
{
    Console.WriteLine($"⚠️ No se encontró la carpeta de páginas en: {rutaCarpetaWeb}");
    // Fallback básico si la carpeta no existe
    app.MapGet("/", () => Results.Content(
        "<html><body style='font-family: sans-serif; text-align: center; padding-top: 50px;'>" +
        "<h1>🚀 API Desposte está en ejecución</h1>" +
        "<p>Carpeta Docu_Paginas no encontrada.</p>" +
        "</body></html>",
        "text/html"
    ));
}

// 4. Activar el ruteo de controladores
app.MapControllers();

// 5. Endpoint de salud para verificación en vivo desde el Tablero
app.MapGet("/api/health", () => Results.Ok(new
{
    status = "online",
    message = "API Desposte operativa",
    servidor = Environment.MachineName,
    timestamp = DateTime.Now
}));

static bool IsPortAvailable(int port)
{
    try
    {
        using var tcpListener = new System.Net.Sockets.TcpListener(System.Net.IPAddress.Any, port);
        tcpListener.Start();
        tcpListener.Stop();
        return true;
    }
    catch
    {
        return false;
    }
}

int[] candidatePorts = { 8080, 5000, 5050, 8085 };
bool anyBound = false;

foreach (int port in candidatePorts)
{
    if (IsPortAvailable(port))
    {
        app.Urls.Add($"http://0.0.0.0:{port}");
        Console.WriteLine($"✅ Escuchando en http://0.0.0.0:{port}");
        anyBound = true;
    }
    else
    {
        Console.WriteLine($"⚠️ Puerto {port} ocupado.");
    }
}

if (!anyBound)
{
    Console.WriteLine("⚠️ No se encontraron puertos candidatos libres, intentando 8080 por defecto.");
    app.Urls.Add("http://0.0.0.0:8080");
}

app.Run();
