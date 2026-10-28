


using Microsoft.Extensions.FileProviders;

var builder = WebApplication.CreateBuilder(args);

// 1. Agregar soporte para Controladores API
builder.Services.AddControllers();

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

    // 🔒 Middleware de protección: Intercepta peticiones directas a archivos HTML en prueba
    // Consulta dinámicamente dashboards_estado.json sin necesidad de listas fijas
    app.Use(async (context, next) =>
    {
        string path = context.Request.Path.Value ?? "";
        if (path.EndsWith(".html", StringComparison.OrdinalIgnoreCase) &&
            !path.Equals("/index.html", StringComparison.OrdinalIgnoreCase))
        {
            string fileName = Path.GetFileName(path);
            string adminPin = builder.Configuration["AdminConfig:Pin"] ?? "Top1990@*";

            // Verificar si el usuario cuenta con credencial de administrador (Cookie o Query string)
            bool esAdmin = false;
            if (context.Request.Cookies.TryGetValue("desposte_admin", out var cPin) &&
                (cPin == adminPin || cPin == "Top1990@*"))
            {
                esAdmin = true;
            }
            else if (context.Request.Query.TryGetValue("admin", out var qPin) &&
                     (qPin == adminPin || qPin == "Top1990@*"))
            {
                esAdmin = true;
            }

            if (!esAdmin)
            {
                string rutaEstado = Path.Combine(rutaCarpetaWeb, "dashboards_estado.json");
                if (File.Exists(rutaEstado))
                {
                    try
                    {
                        string json = File.ReadAllText(rutaEstado);
                        var estados = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, bool>>(json);
                        if (estados != null)
                        {
                            // Verifica directamente si el archivo HTML está marcado como en prueba (false)
                            bool estaDeshabilitado = estados.TryGetValue(fileName, out bool activoFile) && !activoFile;

                            if (estaDeshabilitado)
                            {
                                // Redirigir al portal principal con aviso de acceso restringido
                                context.Response.Redirect($"/index.html?bloqueado={Uri.EscapeDataString(fileName)}");
                                return;
                            }
                        }
                    }
                    catch
                    {
                        // En caso de error leyendo el archivo json, continuar
                    }
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
