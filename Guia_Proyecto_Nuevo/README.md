# Guía: crear un proyecto nuevo con la arquitectura de Api_Desposte

Esta carpeta explica cómo armar desde cero un proyecto con la misma lógica que `Api_Desposte`:

- una **API .NET** que lee datos de **SQL Server** y de **Excel en OneDrive**;
- la misma API **sirve los dashboards HTML** de `Docu_Paginas/` (no hace falta otro servidor web);
- **reglas para agentes** (`AGENTS.md`, `CLAUDE.md` y `.agents/rules/`) para que Claude Code / Codex generen código con el mismo estándar.

## Cómo usar esta guía

Seguir los archivos en orden:

| # | Archivo | Contenido |
|---|---|---|
| 1 | [01_Paso_a_Paso_Consola.md](01_Paso_a_Paso_Consola.md) | Comandos de PowerShell para crear el proyecto, de la carpeta vacía al primer commit |
| 2 | [02_Plantillas_Codigo.md](02_Plantillas_Codigo.md) | Código base: `Program.cs`, `appsettings.json`, controladores SQL y Excel, dashboard HTML |
| 3 | [03_Reglas_Agentes.md](03_Reglas_Agentes.md) | Cómo copiar y adaptar las reglas para agentes |
| 4 | [04_Checklist_Nuevo_Dashboard.md](04_Checklist_Nuevo_Dashboard.md) | Lista de control para cada endpoint + dashboard nuevo |

## Estructura final del proyecto

```
<Proyecto>/
├── AGENTS.md                     # reglas para agentes (Codex y otros)
├── CLAUDE.md                     # mismas reglas para Claude Code
├── .agents/rules/                # detalle de cada regla
├── .gitignore
├── <Api>/                        # API .NET
│   ├── <Api>.csproj
│   ├── Program.cs                # CORS + sirve Docu_Paginas + /api/health + puertos
│   ├── appsettings.json          # cadena de conexión SQL
│   ├── Controllers/
│   │   ├── Sql/                  # endpoints que ejecutan procedimientos almacenados
│   │   └── Excel/                # endpoints que leen tablas de Excel en OneDrive
│   └── Infraestructura/Excel/
│       └── LectorTablasExcel.cs  # lectura estándar de Excel (regla 8)
└── Docu_Paginas/                 # dashboards HTML servidos por la API
    ├── config.js                 # servidor y puertos de la API (regla 1)
    ├── responsive.css            # estilos responsive compartidos (regla 4)
    ├── index.html                # portal de dashboards
    └── <Dashboard>.html
```

## Requisitos

- Windows con PowerShell.
- [.NET SDK 10](https://dotnet.microsoft.com/download) (`dotnet --version` debe mostrar `10.x`).
- Git.
- Acceso a la copia de `Api_Desposte`, de donde se copian las reglas y los archivos base.
- Para SQL: acceso al servidor SQL Server con autenticación de Windows.
- Para Excel: OneDrive sincronizado en el equipo donde corre la API.
