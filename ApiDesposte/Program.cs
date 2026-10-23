


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
