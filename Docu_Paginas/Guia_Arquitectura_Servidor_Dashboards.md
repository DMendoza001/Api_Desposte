# 📘 Informe Técnico y Guía Arquitectónica: Servidor Centralizado de Dashboards
**Proyecto:** `Api_Desposte`  
**Autor:** Antigravity AI Pair Programmer & Dennis Mendoza  
**Fecha:** Octubre 2026  
**Ubicación:** `Docu_Paginas/Guia_Arquitectura_Servidor_Dashboards.md`

---

## 🎯 1. Resumen Ejecutivo

Este documento explica de forma clara, didáctica y técnica cómo funciona el ecosistema de **ApiDesposte**: cómo un comando como `dotnet run` convierte una aplicación .NET en un servidor web dual capaz de hospedar páginas interactivas y servir datos desde SQL Server a cualquier computadora de la red de la planta mediante un simple enlace web, sin necesidad de compilar proyectos pesados ni enviar archivos `.html` sueltos.

---

## ⚙️ 2. ¿Qué ocurre exactamente cuando ejecutas `dotnet run`?

Cuando te ubicas en la carpeta de la API y ejecutas `dotnet run`, suceden varios pasos automáticos en cascada:

```
[dotnet run] 
     │
     ├── 1. Compilación de código C# (.NET 10 / ASP.NET Core)
     ├── 2. Lectura de configuración (appsettings.json)
     ├── 3. Inicialización del servidor web Kestrel
     ├── 4. Registro de Middlewares (CORS, Static Files, Controllers)
     └── 5. Apertura de sockets de red (0.0.0.0:8080, 5000)
```

### El motor interno: Servidor Kestrel
ASP.NET Core incluye un servidor web ultrarrápido y asíncrono llamado **Kestrel**. A diferencia de los servidores tradicionales pesados (como IIS tradicional), Kestrel vive dentro de tu propio proceso de consola.

### ¿Qué significa escuchar en `0.0.0.0` vs `localhost`?
En [`Program.cs`](file:///c:/DENNIS_MENDOZA/Proyectos_GitHub/Api_Desposte/ApiDesposte/Program.cs#L105), el código enlaza los puertos con:
```csharp
app.Urls.Add($"http://0.0.0.0:{port}");
```
* **Si usara `localhost` o `127.0.0.1`:** El sistema operativo solo aceptaría peticiones originadas dentro de tu propia computadora. Nadie más en la empresa podría entrar.
* **Al usar `0.0.0.0`:** Le indicas a Windows que escuche en **todas las tarjetas de red disponibles** (red cableada Ethernet, red Wi-Fi y Loopback local). Cualquier máquina en la red local de la empresa que escriba tu nombre de equipo (`http://RAEVFINL016:8080`) o tu dirección IP (`http://192.168.x.x:8080`) podrá conectarse directamente a tu servidor.

---

## 🌐 3. El Modelo Dual: Servidor de Páginas Web + Servidor de Datos API

En este proyecto, ASP.NET Core cumple **dos roles simultáneos e independientes**:

```
                       ┌─────────────────────────────────────────────────────────────┐
                       │               SERVIDOR WEB KESTREL (dotnet run)             │
                       └──────────────────────────────┬──────────────────────────────┘
                                                      │
                       ┌──────────────────────────────┴──────────────────────────────┐
                       ▼                                                             ▼
         [ ROL 1: SERVIDOR DE PÁGINAS ]                              [ ROL 2: SERVIDOR DE DATOS (API) ]
           Middleware: UseStaticFiles                                  Middleware: MapControllers
                       │                                                             │
         Busca archivos físicos en:                                    Recibe peticiones HTTP REST:
             /Docu_Paginas/*.html                                          /api/InformeDiario
             /Docu_Paginas/*.js                                            /api/Cumplimientos
             /Docu_Paginas/*.css                                           /api/CostosProduccion
                       │                                                             │
         Entrega el archivo tal cual                                   Ejecuta SQL con Dapper y
         al navegador del usuario.                                     devuelve datos puros en JSON.
```

### Configuración en Program.cs
En [`ApiDesposte/Program.cs`](file:///c:/DENNIS_MENDOZA/Proyectos_GitHub/Api_Desposte/ApiDesposte/Program.cs#L43-L54), se configuró este puente:

```csharp
var fileProvider = new PhysicalFileProvider(rutaCarpetaWeb);

// 1. Permite que http://servidor:8080/ abra automáticamente index.html
app.UseDefaultFiles(new DefaultFilesOptions
{
    FileProvider = fileProvider,
    RequestPath = ""
});

// 2. Sirve los archivos HTML, CSS y JS directamente desde Docu_Paginas
app.UseStaticFiles(new StaticFileOptions
{
    FileProvider = fileProvider,
    RequestPath = ""
});
```

---

## 🔄 4. El Ciclo de Vida: ¿Cómo ve un compañero tu reporte desde su laptop?

Imaginemos que compartes el enlace:  
`http://RAEVFINL016:8080/InformeDiario.html`

A continuación se detalla lo que ocurre milisegundo a milisegundo:

```
[ Laptop del Compañero (Navegador) ]                      [ Tu Computadora (dotnet run activo) ]
                 │                                                          │
   Paso 1:       │  GET /InformeDiario.html                                 │
   Petición Web  ├─────────────────────────────────────────────────────────►│
                 │                                                          │ Paso 2: Kestrel lee el archivo
                 │  Respuesta: Archivo HTML (con su CSS y JS)               │ físico desde Docu_Paginas/
   Paso 3:       │◄─────────────────────────────────────────────────────────┤
   Descarga      │                                                          │
                 ▼                                                          │
   El navegador del compañero lee el HTML,                                 │
   aplica el diseño Pantone Industrial y                                    │
   comienza a ejecutar el JavaScript interno.                               │
                 │                                                          │
   Paso 4:       │  GET /api/InformeDiario                                  │
   Petición Datos├─────────────────────────────────────────────────────────►│
                 │                                                          │ Paso 5: El Controller de .NET
                 │                                                          │ consulta SQL Server vía Dapper
                 │  Respuesta: [ { "fecha": "2026-10-23", "kg": 5420... } ] │ y empaqueta en JSON
   Paso 6:       │◄─────────────────────────────────────────────────────────┤
   Recepción     │                                                          │
                 ▼                                                          ▼
   El JavaScript del navegador recibe el JSON,
   hace los cálculos de rendimiento y dibuja
   las tablas y tarjetas al instante.
```

> [!NOTE]
> **Procesamiento distribuido:** La base de datos y la extracción de datos ocurren en tu servidor; pero el pintado, los filtros interactivos, las animaciones y los gráficos se renderizan en el procesador de la computadora de tu compañero. Tu servidor casi no sufre carga de CPU.

---

## ⚖️ 5. Comparativa Técnica: Esta Arquitectura vs Blazor

Como tienes experiencia previa con **Blazor** y arquitecturas por capas, es vital entender por qué para este tipo de dashboards ejecutivos esta solución es más ágil y liviana:

| Aspecto | Blazor Server | Blazor WebAssembly | Nuestra Arquitectura (API + HTML Estático) |
| :--- | :--- | :--- | :--- |
| **Tecnología del cliente** | DOM manipulado por SignalR (WebSocket). | Binarios C# `.dll` ejecutándose sobre WebAssembly mono. | Estándar W3C: HTML5, CSS y JavaScript moderno nativo. |
| **Peso inicial de descarga** | Ligero (~500 KB), pero requiere conexión constante. | Muy pesado (de 10 MB a 30 MB descargando librerías `.dll`). | **Ultraligero (50 KB a 200 KB)**. Carga en menos de medio segundo. |
| **Consumo de memoria en Servidor** | **Alto:** Cada usuario conectado mantiene un árbol de componentes en la RAM del servidor. | Nulo (solo sirve archivos). | **Prácticamente nulo:** Atiende la petición HTTP, entrega el archivo/datos y libera la memoria de inmediato. |
| **Sensibilidad a caídas de red** | Si el Wi-Fi parpadea, la pantalla se congela con error de reconexión (*"Reconnecting to server..."*). | Sigue funcionando sin conexión tras la descarga. | Si la red parpadea, la página sigue abierta visiblemente y reintenta con timeout limpio sin congelar la pantalla. |
| **Facilidad de edición / mantenimiento** | Requiere recompilar todo el proyecto (`dotnet build`) ante cada cambio visual. | Requiere recompilar y regenerar todo el paquete WebAssembly. | **Edición en vivo:** Modificas una línea en el `.html` o agregas un botón, guardas, y el usuario ve el cambio con solo pulsar `F5`. |
| **Compatibilidad** | Requiere navegadores con WebSocket estable. | Requiere navegadores con motor WebAssembly completo. | Funciona en cualquier navegador de PC, laptop, tablet o celular. |

---

## 🚫 6. Por qué el Enlace Centralizado Reemplaza al "Archivo HTML Suelto"

Antes, el flujo tradicional consistía en guardar el archivo HTML y enviarlo por correo, WhatsApp o Teams. Ese esquema generaba serios problemas operativos:

```
❌ ESQUEMA ANTERIOR (Archivos sueltos compartidos):
   [Tú] ──(Envía archivo v1.html por Teams)──► [Compañero 1] (Guarda en Descargas)
   [Tú] ──(Detectas un error, corriges y envías v2.html)──► [Compañero 2] (Usa v2, pero Compañero 1 sigue en v1)
   RESULTADO: Descoordinación, versiones desfasadas y reportes con datos falsos.

✅ ESQUEMA ACTUAL CENTRALIZADO (Acceso por Enlace en Red):
   [Tú] (Modificas Docu_Paginas/InformeDiario.html en tu PC)
   [Compañero 1] ───(Abre http://RAEVFINL016:8080/InformeDiario.html)───► [Tu Servidor]
   [Compañero 2] ───(Abre http://RAEVFINL016:8080/InformeDiario.html)───► [Tu Servidor]
   RESULTADO: Todo el mundo ve SIEMPRE la versión oficial más reciente con solo presionar F5.
```

### Ventajas adicionales de seguridad y mantenimiento:
1. **La base de datos queda protegida:** Nadie en la red conoce el usuario, la contraseña ni la cadena de conexión de SQL Server (`SqlDesposte`). Los compañeros solo acceden a las respuestas JSON públicas filtradas por los controladores de .NET.
2. **Cero instalaciones en las laptops de los clientes:** Tus usuarios no necesitan tener instalado .NET SDK, ni Visual Studio, ni SQL Server Management Studio. Solo requieren abrir Google Chrome, Microsoft Edge o Firefox.
3. **Control total de versiones:** Si decides retirar un reporte obsoleto o ajustar una lógica de negocio, lo haces en tu repositorio y el cambio impacta a toda la planta de inmediato.

---

## 🛡️ 7. El Mecanismo de Resiliencia en Red: `config.js` y `detectarServidorApi()`

Todos los dashboards del repositorio implementan la regla de resiliencia establecida en `.agents/rules/api-connection.md`:

1. **Detección inteligente de URL:**
   Cuando una persona abre el portal desde `http://RAEVFINL016:8080/`, el JavaScript analiza automáticamente:
   ```javascript
   if (window.location.protocol.startsWith('http')) {
       // Si lo abrieron por la web, el servidor activo es exactamente de donde vino la página
       return window.location.origin; 
   }
   ```
2. **Fallback híbrido:** Si alguien abre el archivo localmente con doble clic (`file:///...`), el sistema lee el archivo [`Docu_Paginas/config.js`](file:///c:/DENNIS_MENDOZA/Proyectos_GitHub/Api_Desposte/Docu_Paginas/config.js) e intenta conectarse a los puertos configurados (`8080` y `5000`) con un timeout controlado mediante `AbortController`, evitando que la página se quede congelada esperando indefinidamente.

---

## 📑 8. Catálogo del Tablero Central (`index.html`)

El portal [index.html](file:///c:/DENNIS_MENDOZA/Proyectos_GitHub/Api_Desposte/Docu_Paginas/index.html) fue diseñado siguiendo la **Paleta Pantone Industrial** (Amarillo Maquinaria `#F4D03F`, Carbón Oscuro `#2E2A2F`, Pizarra `#252C37`, Dorado `#BAA063` y Fondo Limpio `#F2F4F8`).

Para evitar sobrecargar la vista con más de 10 tarjetas simultáneas, los dashboards se organizaron en pestañas directas:

### 🥩 1. Operaciones (4 Dashboards - Vista Predeterminada al Iniciar)
- **Informe Diario de Operaciones (`InformeDiario.html`):** Control de turnos (Mañana, Tarde, Noche), asistencia de personal, kilos procesados, rendimiento y paradas de línea.
- **Planificación de Demanda y Cerdos (`NumeroCerdos.html`):** Estimación analítica de cabezas de cerdo requeridas por proyección de kilos.
- **Índice de Rotación Semanal (`Indice_Rotacion_Semanal.html`):** Cálculo de rotación, altas y bajas de personal siguiendo la norma ISO-8601 (semana de Lunes a Domingo).
- **Rendimiento Operativo de Chuletas (`RendimientosChuletas002.html`):** Control de mermas y rendimiento de línea en el trozado de chuletas.

### 🎯 2. Cumplimientos PTC (4 Dashboards)
- **Dashboard Macro PTC (`CumplimientosMacro.html`):** Visión ejecutiva de alto nivel de producción consolidada frente a pedidos comerciales.
- **Cumplimiento Real vs Proyección (`CumplimientosPTC.html`):** Detalle por línea de corte evaluando la producción física ejecutada contra la proyección.
- **Cumplimiento Programa vs Proyección (`CumplimientosPTC_Py.html`):** Comparativa entre el programa planificado y la proyección de demanda.
- **Resumen Planilla PTC (`ResumenPl_PTC.html`):** Horas hombre (HH) ejecutadas versus presupuesto asignado.

### 💰 3. Costos & Presupuesto (3 Dashboards)
- **Presupuesto vs ForeCast 2027 (`Presupuesto_Vs_ForeCast_Dark.html`):** Comparativo integral en formato oscuro ejecutivo.
- **Dashboard Analítico Presupuesto 2027 (`Presupuestos_Final.html`):** Desglose por centros de costo y partidas presupuestarias.
- **Informe de Costo de Producción (`Informe_Costo_Produccion.html`):** Costo unitario por kilogramo procesado y análisis de costos directos/indirectos.

### 👥 4. Personal (1 Dashboard)
- **Gestión de Personal (`Personal_Prueba.html`):** Consulta del padrón de colaboradores conectado al controlador de sincronización Excel.

---

## 🚀 9. Guía de Uso Diario: Pasos para Compartir Reportes

### Paso 1: Iniciar el servicio en tu computadora
Abre una terminal en la carpeta del proyecto y ejecuta:
```bash
cd c:\DENNIS_MENDOZA\Proyectos_GitHub\Api_Desposte\apidesposte
dotnet run
```
Verás la confirmación en consola:
```text
🌐 Servidor web de dashboards habilitado desde: C:\...\Docu_Paginas
✅ Escuchando en http://0.0.0.0:8080
✅ Escuchando en http://0.0.0.0:5000
```

### Paso 2: Obtener tu enlace para compartir
Para que cualquier compañero de la planta acceda, simplemente dale el enlace:
- **Al Tablero General:** `http://RAEVFINL016:8080/`
- **A un reporte específico:** `http://RAEVFINL016:8080/InformeDiario.html`

*(También puedes usar tu IP local en lugar del nombre del equipo, por ejemplo: `http://192.168.1.50:8080/`)*.

### Paso 3: Usar el botón "Copiar Link"
Dentro del portal `index.html`, cada tarjeta cuenta con un botón **"📋 Copiar Link"**. Al hacer clic, copia automáticamente la URL lista para pegar en Teams, correo o chat corporativo.

### Paso 4: ¿Qué hacer si alguien no puede entrar? (Firewall de Windows)
Si un compañero te indica que el enlace no le carga:
1. Abre en tu Windows el menú **Windows Defender Firewall** -> **Permitir que una aplicación se comunique a través de Firewall de Windows**.
2. Asegúrate de que `dotnet.exe` o el puerto `8080` tenga marcada la casilla de red **Privada / Dominio**.
