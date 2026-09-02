


var builder = WebApplication.CreateBuilder(args);

// 1. Agregar soporte para Controladores API
builder.Services.AddControllers();

// 2. Configurar CORS (Para permitir solicitudes desde Excel Web)
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

// 4. Activar el ruteo de controladores
app.MapControllers();

// 5. Ruta raíz amigable para confirmar funcionamiento
app.MapGet("/", () => Results.Content(
    "<html><body style='font-family: sans-serif; text-align: center; padding-top: 50px;'>" +
    "<h1>🚀 API Desposte está en ejecución</h1>" +
    "<p>Esta es una Web API y no tiene interfaz de usuario por defecto.</p>" +
    "<p>Puedes probar la conexión a la base de datos en: " +
    "<a href='/api/desposte/probar-conexion'>/api/desposte/probar-conexion</a></p>" +
    "</body></html>", 
    "text/html"
));

app.Run("http://0.0.0.0:8080");