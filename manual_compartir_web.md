# Manual: Compartir Páginas Web (`index.html`) en Red Local por Puerto

Este manual detalla paso a paso las dos mejores alternativas en el ecosistema **.NET** para publicar y compartir tu archivo `index.html` (y sus recursos asociados) con cualquier usuario de tu organización a través de un puerto de red.

---

## 🏆 Comparativa y Recomendación

| Criterio | Opción 1: En tu API .NET (`UseStaticFiles`) | Opción 2: Herramienta `dotnet-serve` |
| :--- | :--- | :--- |
| **Comando de inicio** | `dotnet run` (el mismo que ya usas) | `dotnet serve -a 0.0.0.0 -p 8090` |
| **Procesos en ejecución** | **1 solo proceso** (API + Web juntos) | **2 procesos** (1 terminal para la API, 1 para la web) |
| **Problemas de CORS** | **Cero**. Mismo origen y puerto | Puede requerir flags `--cors` si llama a otra API |
| **Ideal para...** | Páginas que consumen tu API (`Docu_Paginas`) | Proyectos web aislados que no tocan tu API |

> [!TIP]
> ### 💡 ¿Cuál recomiendo?
> **Recomiendo la Opción 1 (Integrada en la API .NET)** si las páginas HTML consumen endpoints de tu API (como `NumeroCerdos.html` o `HorasExtras.html`). Con un solo `dotnet run`, la API entrega tanto los datos como la interfaz gráfica, eliminando errores de red, restricciones de navegador (`file:///`) y bloqueos de CORS.
> 
> **Elige la Opción 2 (`dotnet-serve`)** si ese `index.html` pertenece a un proyecto totalmente independiente y solo quieres levantarlo temporalmente en 5 segundos sin modificar código C#.

---

## Opción 1: Servir la Web desde tu API .NET (Recomendada)

Permite que el mismo comando `dotnet run` que levanta tu API también publique el sitio web.

### Paso 1: Configurar [Program.cs](file:///c:/DENNIS_MENDOZA/Proyectos_GitHub/API_Desposte/apidesposte/Program.cs)

Abre el archivo [Program.cs](file:///c:/DENNIS_MENDOZA/Proyectos_GitHub/API_Desposte/apidesposte/Program.cs) y añade el soporte de archivos estáticos apuntando a la carpeta de tu proyecto web (por ejemplo, `Docu_Paginas`):

```csharp
using Microsoft.Extensions.FileProviders;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddControllers();

// Configuración CORS
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
// 🌐 CONFIGURACIÓN PARA SERVIR EL SITIO WEB (index.html, JS, CSS)
// =========================================================================
// Ajusta la ruta a la carpeta donde esté tu index.html:
string rutaCarpetaWeb = Path.Combine(builder.Environment.ContentRootPath, "..", "Docu_Paginas");

if (Directory.Exists(rutaCarpetaWeb))
{
    // Permite que http://RAEVSALL001:8080/ abra index.html por defecto
    app.UseDefaultFiles(new DefaultFilesOptions
    {
        FileProvider = new PhysicalFileProvider(rutaCarpetaWeb),
        RequestPath = ""
    });

    // Sirve todos los archivos .html, .js, .css, imágenes, etc.
    app.UseStaticFiles(new StaticFileOptions
    {
        FileProvider = new PhysicalFileProvider(rutaCarpetaWeb),
        RequestPath = ""
    });

    Console.WriteLine($"🌐 Servidor web estático habilitado desde: {rutaCarpetaWeb}");
}

app.UseCors("PermitirExcelWeb");
app.MapControllers();

// Configurar puertos de red accesibles externamente (0.0.0.0)
int[] candidatePorts = { 8080, 5000, 5050, 8085 };
foreach (int port in candidatePorts)
{
    app.Urls.Add($"http://0.0.0.0:{port}");
}

app.Run();
```

### Paso 2: Iniciar el servicio
En la terminal de tu proyecto:
```powershell
cd c:\DENNIS_MENDOZA\Proyectos_GitHub\API_Desposte\apidesposte
dotnet run
```

### Paso 3: Acceso para tus compañeros
Cualquier usuario conectado a la red local/corporativa puede ingresar desde Chrome, Edge o Firefox a:
* **Página principal (`index.html`):**  
  `http://RAEVSALL001:8080/`
* **Páginas específicas:**  
  `http://RAEVSALL001:8080/NumeroCerdos.html`  
  `http://RAEVSALL001:8080/HorasExtras.html`

---

## Opción 2: Usar la herramienta CLI `dotnet-serve`

Si tienes una carpeta en cualquier parte con un `index.html` y no deseas modificar ningún proyecto C#, usa `dotnet-serve`.

### Paso 1: Instalación (Solo se hace una vez en el equipo)
Abre PowerShell y ejecuta:
```powershell
dotnet tool install --global dotnet-serve
```

> [!NOTE]
> Para actualizarla en el futuro si lo necesitas: `dotnet tool update --global dotnet-serve`.

### Paso 2: Navegar a la carpeta del proyecto
Ubícate en el directorio donde reside tu `index.html`:
```powershell
cd "c:\Ruta\A\Tu\CarpetaWeb"
```

### Paso 3: Ejecutar el servidor web
Lanza el comando indicando que escuche en todas las IPs (`0.0.0.0`) y asignando un puerto disponible:
```powershell
dotnet serve -a 0.0.0.0 -p 8090 -o
```

**Parámetros clave explicados:**
* `-a 0.0.0.0`: **Obligatorio para compartir en red.** Si no lo pones, solo escuchará en `localhost` y tus compañeros no podrán entrar.
* `-p 8090`: Puerto donde se publicará la web (puedes cambiarlo a 8088, 8095, etc.).
* `-o`: Abre automáticamente la página en tu navegador predeterminado.
* `--default-file index.html`: *(Opcional)* Establece el archivo raíz si tuviera otro nombre.
* `-c`: *(Opcional)* Habilita encabezados CORS si este HTML realiza peticiones a otras APIs.

### Paso 4: Acceso para tus compañeros
Tus compañeros ingresan a:
```text
http://RAEVSALL001:8090/
```
*(Para detener el servidor, solo presionas `Ctrl + C` en esa ventana de terminal).*

---

## 🔒 Configuración de Red y Firewall de Windows

Si tus compañeros intentan ingresar y la página se queda cargando o dice *"No se puede acceder al sitio"*, significa que el Firewall de Windows en tu máquina bloquea las conexiones entrantes al puerto.

### 1. Obtener tu Dirección IP local
En PowerShell ejecuta:
```powershell
(Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notlike "*Loopback*" -and $_.IPAddress -like "172.*" -or $_.IPAddress -like "192.168.*" -or $_.IPAddress -like "10.*" }).IPAddress
```
*(Si no les resuelve por el nombre de equipo `RAEVSALL001`, pueden entrar directamente con tu IP, ej: `http://172.20.21.177:8090/`)*.

### 2. Abrir el puerto en el Firewall (Ejecutar en PowerShell como Administrador)
Para permitir el acceso a los puertos que elijas (ej. 8080 y 8090):
```powershell
New-NetFirewallRule -DisplayName "Acceso Servidor Web Desposte" -Direction Inbound -LocalPort 8080,8090 -Protocol TCP -Action Allow
```
