# Documentación: Entendiendo el Proyecto ApiDesposte

Este documento explica de forma sencilla cómo funciona el proyecto **ApiDesposte** y cómo se relaciona con los conceptos de **Arquitectura en Capas** y **Blazor** con los que ya venías trabajando.

---

## 1. ¿Qué es este proyecto y en qué se diferencia de Blazor?

En tu flujo de trabajo anterior con **Blazor** y **Arquitectura en Capas**, solías tener una estructura parecida a esta:
1. **Capa de Presentación (UI):** Componentes `.razor` (HTML, CSS y C# corriendo en Blazor WebAssembly o Blazor Server).
2. **Capa de Negocio (Servicios):** Clases con las reglas de negocio (ej. `DesposteService.cs`).
3. **Capa de Datos:** Clases que acceden a la base de datos mediante Entity Framework o Dapper (ej. `DesposteRepository.cs`).
4. **Base de Datos:** SQL Server.

### La diferencia principal con ApiDesposte:
El proyecto **ApiDesposte** es una **Web API pura** (ASP.NET Core Web API). 
* **No tiene interfaz de usuario (UI):** No hay archivos `.razor` ni vistas HTML en el código C#.
* **Su único propósito es servir datos:** Actúa como un intermediario o "puente" que recibe solicitudes desde el exterior, realiza la lógica y responde con datos puros en formato **JSON**.
* **¿Quién consume esta API?** La pueden consumir formularios HTML/JS tradicionales, aplicaciones móviles, scripts de Office (Office Scripts en Excel Web) o incluso tu misma aplicación Blazor.

---

## 2. Estructura de Archivos del Proyecto

El proyecto está diseñado de forma muy simple y directa (minimalista). Esta es su estructura:

```text
ApiDesposte/
│
├── Program.cs                  # Configuración de arranque del servidor
├── appsettings.json            # Cadenas de conexión y variables de configuración
├── ApiDesposte.csproj          # Configuración del proyecto de .NET (dependencias Nuget)
│
└── Controllers/                # Capa de Controladores (reciben las peticiones del exterior)
    ├── DesposteController.cs       # Acceso a SQL Server usando Dapper
    ├── PersonalExcelController.cs   # Lectura/Escritura directa sobre archivos Excel
    └── ExcelLocalController.cs      # (Vacío/Estructura por definir)
```

---

## 3. ¿Cómo funciona cada componente?

### A. Program.cs (El motor de arranque)
En [`Program.cs`](file:///g:/API_Desposte/ApiDesposte/ApiDesposte/Program.cs), se configura cómo el servidor web se comporta:
* **Habilita Controladores:** Registra las clases dentro de `Controllers/` para que respondan a URLs específicas.
* **Configura CORS (Cross-Origin Resource Sharing):** Permite que aplicaciones externas (como Excel en la web o páginas web en otro servidor/dominio) puedan solicitarle datos a esta API sin bloqueos de seguridad.
* **Define el puerto:** Configurado para ejecutarse en `http://0.0.0.0:5000` (el puerto 5000).

### B. appsettings.json (Configuración)
Aquí se guarda la cadena de conexión a tu base de datos SQL Server (`SqlDesposte`), que apunta a la instancia `DMENDOZA\SQLEXPRESS`.

---

## 4. Análisis de los Controladores

En este proyecto, en lugar de dividir en 3 proyectos físicos distintos (Presentación, Negocio, Datos), toda la lógica se maneja de forma compacta a través de los **Controladores** (`Controllers`):

### 📌 Controller 1: [DesposteController.cs](file:///g:/API_Desposte/ApiDesposte/ApiDesposte/Controllers/DesposteController.cs) (Acceso a SQL Server)
Este controlador utiliza **Dapper** para comunicarse con la base de datos SQL Server. 

* **Inyección de Dependencia:** Obtiene la cadena de conexión desde `appsettings.json` en su constructor.
* **Dapper (`using IDbConnection`):** Abre la conexión y ejecuta consultas directas o Procedimientos Almacenados (ej. `Sp_Informe_AtencionesTopico`).
* **Endpoints:**
  * `api/desposte/probar-conexion`: Para comprobar que la conexión a SQL Server funcione.
  * `api/desposte/atenciones-topico`: Ejecuta el procedimiento almacenado y te devuelve los datos formateados en JSON.
  * `api/desposte/dashboard-topico-pivot`: Obtiene datos consolidados por semanas/días haciendo una consulta PIVOT en SQL Server.

### 📌 Controller 2: [PersonalExcelController.cs](file:///g:/API_Desposte/ApiDesposte/ApiDesposte/Controllers/PersonalExcelController.cs) (Acceso a Excel)
Este controlador implementa una lógica donde la **"Base de datos" es un archivo Excel** (`AtencionesTopico001.xlsx`) alojado en tu OneDrive/SharePoint.

* **ClosedXML:** Carga el archivo `.xlsx` usando streams en memoria para leer o actualizar celdas de una tabla de Excel llamada `T_Personal`.
* **Lectura Compartida:** Abre el archivo con permisos compartidos para evitar fallos de lectura si el archivo está abierto en otra parte.
* **Endpoints:**
  * `api/PersonalExcel/listar`: Lee la tabla Excel y devuelve los nombres, DNIs y teléfonos.
  * `api/PersonalExcel/buscar/{dni}`: Busca una fila exacta por DNI.
  * `api/PersonalExcel/crear`, `actualizar`, `eliminar`: Realizan operaciones de escritura directa sobre el Excel físico y fuerzan a que OneDrive/SharePoint detecte el cambio modificando la fecha de última escritura.

---

## 5. Mapeo Mental: Blazor / Capas vs. ApiDesposte

Para ayudarte a asimilar este modelo comparándolo con lo que ya conoces:

| Concepto en Blazor / Capas | Equivalente en ApiDesposte (Web API) |
| :--- | :--- |
| **Componente `.razor`** | **No existe aquí.** La UI se desarrolla fuera (puede ser HTML independiente, Excel Web o incluso un proyecto Blazor WebAssembly aparte que consuma esta API). |
| **Petición del usuario (Clic en botón)** | **Llamada HTTP** (`GET`, `POST`, `PUT`, `DELETE`) hacia las URLs de la API (ej: `/api/desposte/atenciones-topico`). |
| **Capa de Servicios / Repositorios** | Los métodos dentro de los **Controladores** (`DesposteController.cs` y `PersonalExcelController.cs`) asumen temporalmente estas tareas para simplificar el flujo. |
| **Entity Framework Core (EF)** | Se reemplaza aquí por **Dapper** (más rápido y directo para ejecutar SPs de SQL Server) o por **ClosedXML** (para los archivos Excel). |

---

## 6. Guía paso a paso: Crear una Web API desde cero usando la Terminal

Si deseas recrear un proyecto como este de forma totalmente limpia utilizando la consola (`PowerShell` o `cmd`), sigue estos pasos:

### Paso 1: Crear la carpeta del proyecto
Abre tu terminal y crea un nuevo directorio para tu proyecto, luego accede a él:
```bash
mkdir MiApiProyecto
cd MiApiProyecto
```

### Paso 2: Crear el nuevo proyecto Web API de ASP.NET Core
Por defecto, las plantillas modernas de .NET generan APIs del tipo "Minimal APIs". Para crear una API basada en **Controladores** (como la de este proyecto), utiliza la bandera `--use-controllers`:
```bash
dotnet new webapi --use-controllers -o .
```
*(La bandera `-o .` indica que cree el proyecto directamente en la carpeta actual sin anidar más directorios).*

### Paso 3: Instalar los paquetes NuGet necesarios
Instala las dependencias que usa este proyecto para interactuar con SQL Server y Excel ejecutando los siguientes comandos:
```bash
# 1. Instalar Dapper (Micro-ORM rápido para consultas SQL)
dotnet add package Dapper

# 2. Instalar el proveedor de SQL Server
dotnet add package Microsoft.Data.SqlClient

# 3. Instalar ClosedXML (Para manipular archivos Excel)
dotnet add package ClosedXML
```

### Paso 4: Ejecutar y probar el proyecto
Para levantar el servidor web de desarrollo en tu máquina local:
```bash
dotnet run
```
La consola te mostrará las direcciones URL locales en las que el servidor está escuchando (por ejemplo, `http://localhost:5000` o `https://localhost:5001`). Puedes ingresar a esas rutas en tu navegador para validar que responde correctamente.

