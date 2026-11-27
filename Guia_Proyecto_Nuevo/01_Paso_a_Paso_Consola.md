# 1. Paso a paso en consola (PowerShell)

Todos los comandos son para **Windows PowerShell**. Ejecutarlos en orden, en la misma ventana, porque los pasos usan las variables del paso 0.

---

## Paso 0. Verificar herramientas y definir nombres

```powershell
dotnet --version   # debe mostrar 10.x
git --version

# Nombres del proyecto nuevo (cambiar a gusto)
$Proyecto = "Api_Planta"                       # carpeta raíz / repositorio
$Api      = "ApiPlanta"                        # proyecto .NET (sin espacios ni guiones)
$Base     = "C:\Proyectos"                     # dónde se creará el proyecto
$Origen   = "C:\ruta\a\Api_Desposte"           # copia local de Api_Desposte (de aquí se copian reglas y archivos base)
```

---

## Paso 1. Crear la carpeta raíz y el repositorio

```powershell
New-Item -ItemType Directory -Force "$Base\$Proyecto" | Out-Null
Set-Location "$Base\$Proyecto"
git init
dotnet new gitignore
```

---

## Paso 2. Crear la API .NET y sus paquetes

```powershell
dotnet new web -n $Api -o $Api -f net10.0

dotnet add $Api package ClosedXML                 # leer Excel
dotnet add $Api package Dapper                    # ejecutar SQL
dotnet add $Api package Microsoft.Data.SqlClient  # conexión a SQL Server
```

---

## Paso 3. Crear las carpetas

```powershell
New-Item -ItemType Directory -Force `
    "$Api\Controllers\Sql", `
    "$Api\Controllers\Excel", `
    "$Api\Funciones\Excel", `
    "$Api\Funciones\Sql", `
    "$Api\Funciones\Fechas", `
    "Docu_Paginas\Funciones", `
    ".agents\rules" | Out-Null
```

---

## Paso 4. Copiar las funciones compartidas del backend

`LectorTablasExcel.cs` (regla 8), `EjecutorSql.cs` y `ParametrosFecha.cs` se copian y se cambia el namespace:

```powershell
foreach ($f in "Excel\LectorTablasExcel.cs", "Sql\EjecutorSql.cs", "Fechas\ParametrosFecha.cs") {
    $destino = "$Api\Funciones\$f"
    Copy-Item "$Origen\ApiDesposte\Funciones\$f" $destino
    (Get-Content $destino -Raw -Encoding UTF8).Replace("namespace ApiDesposte.", "namespace $Api.") |
        Set-Content $destino -Encoding UTF8
}
```

---

## Paso 5. Copiar los archivos base del frontend

```powershell
Copy-Item "$Origen\Docu_Paginas\Funciones\config.js"      "Docu_Paginas\Funciones\"
Copy-Item "$Origen\Docu_Paginas\Funciones\responsive.css" "Docu_Paginas\Funciones\"
Copy-Item "$Origen\PaletaFinanza_Industrial.json" "."
```

Luego abrir `Docu_Paginas\Funciones\config.js` y cambiar `servidor` por el nombre del equipo donde correrá la API:

```powershell
$env:COMPUTERNAME   # nombre de este equipo, para usarlo en config.js
code Docu_Paginas\Funciones\config.js
```

> `index.html` (portal) se puede copiar de `Api_Desposte` y adaptar su lista `LISTA_DASHBOARDS`. No copiar PINs ni contraseñas escritos en el código: si se necesita acceso de administrador, leerlo de una variable de entorno (ej. `AdminConfig__Pin`).

---

## Paso 6. Copiar y adaptar las reglas para agentes

```powershell
Copy-Item "$Origen\AGENTS.md" "."
Copy-Item "$Origen\CLAUDE.md" "."
Copy-Item "$Origen\.agents\rules\*.md" ".agents\rules\"
```

Después ajustar los nombres propios según [03_Reglas_Agentes.md](03_Reglas_Agentes.md).

---

## Paso 7. Escribir `Program.cs` y `appsettings.json`

Reemplazar el contenido de `$Api\Program.cs` y `$Api\appsettings.json` con las plantillas de [02_Plantillas_Codigo.md](02_Plantillas_Codigo.md) (secciones 1 y 2).

```powershell
code "$Api\Program.cs"
code "$Api\appsettings.json"
```

---

## Paso 8. Compilar y ejecutar

```powershell
dotnet build $Api
dotnet run --project $Api
```

La consola debe mostrar algo como:

```
🌐 Servidor web de dashboards habilitado desde: ...\Docu_Paginas
✅ Escuchando en http://0.0.0.0:8080
```

> Si dice "Puerto 8080 ocupado", la API usa el siguiente libre (5000, 5050 u 8085). Si los cuatro están ocupados, ya hay otra API corriendo (por ejemplo Api_Desposte en el mismo equipo): detenerla o agregar otro puerto a la lista de `Program.cs`.

En **otra** ventana de PowerShell, verificar que responde (usar el puerto que mostró la consola):

```powershell
Invoke-RestMethod http://localhost:8080/api/health
```

Y abrir en el navegador un dashboard: `http://localhost:8080/<Dashboard>.html`. La raíz `http://localhost:8080/` abre el portal solo si se copió `index.html` en el paso 5.

---

## Paso 9. Permitir el acceso desde la red (opcional)

Para que otras PCs o celulares abran `http://<NOMBRE_EQUIPO>:8080/`, abrir los puertos en el firewall. **Requiere PowerShell como administrador:**

```powershell
New-NetFirewallRule -DisplayName "$Api API" -Direction Inbound -Protocol TCP -LocalPort 8080,5000 -Action Allow
```

---

## Paso 10. Primer commit

```powershell
git add .
git commit -m "Estructura inicial: API .NET + Docu_Paginas + reglas para agentes"
```

---

## Paso 11. Agregar endpoints y dashboards

Por cada informe nuevo:

1. Crear el controlador en `Controllers\Sql\` o `Controllers\Excel\` usando las plantillas de [02_Plantillas_Codigo.md](02_Plantillas_Codigo.md).
2. Crear el dashboard en `Docu_Paginas\` a partir de la plantilla HTML.
3. Revisar [04_Checklist_Nuevo_Dashboard.md](04_Checklist_Nuevo_Dashboard.md).

Con Claude Code abierto en la carpeta del proyecto, basta pedirlo ("crea un endpoint para el SP X y su dashboard"): las reglas de `CLAUDE.md` se aplican solas.
