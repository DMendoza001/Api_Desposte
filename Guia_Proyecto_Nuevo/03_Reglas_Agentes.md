# 3. Reglas para agentes (`AGENTS.md`, `CLAUDE.md`, `.agents/rules/`)

## 3.1 Cómo funcionan

| Archivo | Lo lee | Cuándo |
|---|---|---|
| `CLAUDE.md` | Claude Code | Automáticamente, al iniciar cada sesión en la carpeta del proyecto |
| `AGENTS.md` | Codex y otros agentes | Automáticamente, al iniciar |
| `.agents/rules/*.md` | El agente | Solo cuando `CLAUDE.md` / `AGENTS.md` lo remiten a ese archivo |

Por eso cada regla tiene dos partes:

1. **Resumen corto** en `CLAUDE.md` y `AGENTS.md` (iguales), que indica **cuándo aplica** y enlaza al detalle.
2. **Detalle** en `.agents/rules/<regla>.md`.

Una regla que solo existe en `.agents/rules/`, sin enlace en `CLAUDE.md` / `AGENTS.md`, el agente no la ve.

## 3.2 Reglas que se heredan de Api_Desposte

| # | Regla | Archivo de detalle |
|---|---|---|
| 1 | Conexión frontend → API .NET | `api-connection.md` |
| 2 | Semanas de lunes a domingo (ISO-8601) | `calculo-semanas.md` |
| 3 | Excel en OneDrive con ruta dinámica | `conexion-excel-onedrive.md` |
| 4 | Diseño responsive | `responsive-design.md` |
| 5 | Matrices jerárquicas sin iconos decorativos | (solo en `CLAUDE.md` / `AGENTS.md`) |
| 6 | Una ruta canónica por controlador | (solo en `CLAUDE.md` / `AGENTS.md`) |
| 7 | Exportar data plana de SQL Server | `exportar-data-plana-sql.md` |
| 8 | Uso de `LectorTablasExcel` / `CacheTablasExcel` | `conexion-excel-onedrive.md` |
| 9 | Filtros de periodo por vista temporal | `filtros-periodo-vistas.md` |

## 3.3 Adaptar los nombres al proyecto nuevo

Después de copiar los archivos (paso 6 de la consola), reemplazar los nombres propios de Api_Desposte. Ejecutar en la carpeta raíz del proyecto nuevo, con las variables del paso 0:

```powershell
$Servidor = $env:COMPUTERNAME     # equipo donde correrá la API

$archivos = @("AGENTS.md", "CLAUDE.md", "Docu_Paginas\Funciones\config.js") +
            (Get-ChildItem ".agents\rules\*.md").FullName

foreach ($f in $archivos) {
    $txt = Get-Content $f -Raw -Encoding UTF8
    $txt = $txt.Replace("Api_Desposte", $Proyecto).Replace("ApiDesposte", $Api).Replace("RAEVSALL001", $Servidor)
    Set-Content $f $txt -Encoding UTF8
}

# Revisar qué quedó con el nombre anterior
Select-String -Path $archivos -Pattern "Desposte|RAEVSALL001"
```

Revisar a mano:

- **Descripción del proyecto** en la primera sección de `CLAUDE.md` y `AGENTS.md` ("...contiene la API de Desposte..."): describir el proyecto nuevo.
- **Servidor:** si la API correrá en el mismo equipo que Api_Desposte, `RAEVSALL001` se mantiene (el reemplazo deja el mismo nombre).
- **Ruta de OneDrive** (`OneDrive - Corporación Rico SAC\...`) en `CLAUDE.md`, `AGENTS.md` y `conexion-excel-onedrive.md`: cambiarla si la empresa o la carpeta son otras.
- **Reglas que no apliquen:** si el proyecto no usa Excel, quitar las reglas 3 y 8 de los **dos** archivos (`CLAUDE.md` y `AGENTS.md`).

## 3.4 Agregar una regla nueva

1. Crear `.agents/rules/<nombre-regla>.md` con el detalle.
2. Agregar en `CLAUDE.md` **y** en `AGENTS.md` una sección numerada con el mismo texto:

```markdown
### 10. <Nombre de la regla> (Regla Obligatoria)
Todo <tipo de archivo> en `<carpeta>/` debe seguir lo definido en [.agents/rules/<nombre-regla>.md](.agents/rules/<nombre-regla>.md):
- <punto clave 1>
- <punto clave 2>
```

3. Para aplicarla en la sesión actual de Claude Code, pedirlo explícitamente ("aplica la regla X"). Las sesiones nuevas la cargan solas.
