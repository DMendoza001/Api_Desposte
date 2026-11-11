# Api_Desposte - Guía de Reglas del Proyecto

Este repositorio contiene la API de Desposte en .NET y el conjunto de dashboards/reportes analíticos en HTML interactivos ubicados en `Docu_Paginas/`.

## Reglas Generales de Desarrollo

### 1. Conexión Frontend -> API .NET (Regla Obligatoria)
Todos los archivos `.html` deben seguir el estándar híbrido y resiliente de conexión establecido en [.agents/rules/api-connection.md](file:///.agents/rules/api-connection.md):
- Importar `config.js` opcionalmente en el `<head>`: `<script src="config.js"></script>` (sin scripts inline en `<head>`).
- Declarar el bloque estándar `APP_CONFIG` con fallback a `RAEVSALL001`, puertos `[8080, 5000]`, `controllerPath` y `timeoutMs`.
- Utilizar `generarBasesApi()` para resolver URLs de la API.
- Implementar `fetchConTimeout(url, timeoutMs)` y `detectarServidorApi()` con timeout adecuado (mínimo 3.5s) y fallback de ping a raíz (`/`).
- Nunca hardcodear URLs fijas directamente en llamadas `fetch` sin pasar por `generarBasesApi()`.

### 2. Criterio de Cálculo de Semanas (Regla Obligatoria)
Todo cálculo, análisis, reporte o agregación semanal en el proyecto (tanto en C# .NET como en dashboards HTML / JavaScript) debe seguir estrictamente lo definido en [.agents/rules/calculo-semanas.md](file:///.agents/rules/calculo-semanas.md):
- **El primer día de la semana es obligatoriamente el LUNES** (y termina en Domingo).
- Utilizar el cálculo estándar ISO-8601 (`(d.getDay() + 6) % 7` en JS, `ISOWeek.GetWeekOfYear` en C#).
- En presentaciones visuales semanales (gráficos/tablas), incluir preferentemente el rango `Sem. XX (DD/MM al DD/MM)`.

### 3. Conexión Dinámica a Archivos Excel en OneDrive (Regla Obligatoria)
Todo controlador en C# .NET (`ApiDesposte/Controllers/Excel/`) que consuma archivos Excel locales sincronizados en OneDrive debe seguir estrictamente lo definido en [.agents/rules/conexion-excel-onedrive.md](file:///.agents/rules/conexion-excel-onedrive.md):
- **Nunca hardcodear el usuario en la ruta física** (ej. `C:\Users\dennis.mendoza\...`).
- Utilizar obligatoriamente `ObtenerRutaExcel()` resolviendo el perfil dinámicamente con `Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)` y `Path.Combine(userProfile, @"OneDrive - Corporación Rico SAC\...")`.
- Abrir siempre el archivo en modo compartido (`FileShare.ReadWrite | FileShare.Delete`) cargándolo en memoria (`MemoryStream`).

### 4. Diseño Responsive Híbrido para Laptops, Tablets y Celulares (Regla Obligatoria)
Todos los archivos `.html` en `Docu_Paginas/` deben seguir el estándar de diseño responsive definido en [.agents/rules/responsive-design.md](file:///.agents/rules/responsive-design.md):
- Enlazar `<link rel="stylesheet" href="responsive.css?v=3">` en el `<head>` junto con un bloque `@media` autónomo en `<style>`.
- **Prohibido usar alturas fijas inline sin clase en contenedores de gráficos** (ej. `<div style="height:430px;">`); usar siempre clases estándar (`.chart-box`, `.chart-container`) con `position: relative; width: 100%; min-width: 0;`.
- Aplicar `min-width: 0` tanto en hijos de **CSS Grid** como de **Flexbox** (`.dashboard-container > *`, `.kpi-card`, `.chart-card`, `.matrix-card`) para evitar que tablas anchas o gráficos desborden la pantalla.
- Mantener las 4 tarjetas KPI en una sola fila en laptops (`repeat(4, minmax(0, 1fr))` o `minmax(185px, 1fr)`, nunca `minmax(260px, 1fr)`).
- Activar la densidad compacta de laptop desde **`@media screen and (max-width: 1600px)`** (para cubrir laptops Full HD 1920x1080 con escala de Windows al 125% = `1536px`, además de 150% y 1366x768) y adaptar filtros, gráficos y columnas fijas en tablets (`<= 980px`) y celulares (`<= 768px`).
- **Etiquetas de datos en gráficos (`datalabels`) en celular:** En los gráficos que ya tengan etiquetas de datos (`chartjs-plugin-datalabels`), ocultarlas en pantallas móviles (`display: () => window.innerWidth > 768`) para evitar superposición de textos, conservándolas en el resto de pantallas (`> 768px`) únicamente en los gráficos que ya las tenían.

### 5. Matrices Jerárquicas sin Iconos Decorativos (Regla Obligatoria)
Toda tabla o matriz jerárquica de los dashboards en `Docu_Paginas/` debe mostrar sus etiquetas como **texto plano**:
- **Único icono permitido:** el control de expansión/colapso, porque comunica una acción. Cualquier icono, emoji o pictograma meramente decorativo se elimina.
- **Etiquetas planas:** los niveles se muestran sin fondo, borde, sombra ni decoración adicional; sólo tipografía y colores de la paleta.
- **Limpieza por CSS:** las decoraciones se anulan con reglas de hoja de estilos, sin modificar la lógica que genera las filas.
- **Jerarquía por estructura:** la profundidad se comunica con sangría, peso tipográfico y color, no con iconos repetidos en cada fila.
- **Reparto de anchos:** la columna de jerarquía absorbe el espacio sobrante y las columnas numéricas se ajustan a su contenido, para que el nombre de la fila se muestre completo.

### 6. Ruta Canónica Única por Controlador y Endpoint en la API .NET (Regla Obligatoria)
Todo controlador en `ApiDesposte/Controllers/` debe exponer **una única ruta canónica** por controlador y por método de acción (estándar REST / OpenAPI):
- **Nombre de clase alineado al archivo:** La clase `<Nombre>Controller : ControllerBase` debe coincidir con el nombre del archivo `<Nombre>Controller.cs`.
- **Un solo `[Route]` por controlador:** Usar únicamente `[Route("api/[controller]")]` en la clase (sin apilar alias legacy ni rutas duplicadas sobre el mismo controlador).
- **Un solo atributo de verbo HTTP por acción:** Cada método endpoint debe tener un único atributo de ruta/verbo (ej. un solo `[HttpGet]` o `[HttpGet("subruta")]`), evitando duplicar `[HttpGet]`, `[HttpGet("listar")]`, `[HttpGet("datos")]` sobre el mismo método.
- **Consumo directo desde el Frontend:** El dashboard HTML debe apuntar su `APP_CONFIG.controllerPath` directamente a esa ruta canónica única.

### 7. Botón de Exportación de Data Plana para Fuentes SQL Server (Regla Obligatoria)
Todo dashboard en `Docu_Paginas/` cuya fuente de datos provenga de SQL Server (controladores en `ApiDesposte/Controllers/Sql/`) debe seguir estrictamente lo definido en [.agents/rules/exportar-data-plana-sql.md](file:///.agents/rules/exportar-data-plana-sql.md):
- **Carga de SheetJS:** Incluir `<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>` en el `<head>`.
- **Preservación de data plana:** Almacenar en una variable de memoria (`rawSqlData` / `registrosCrudos`) el array de objetos original devuelto por la API .NET tal cual viene de SQL Server, sin filtros ni modificaciones.
- **Botón en barra de herramientas / cabecera:** Exponer un botón visible y accesible (`Data Plana SQL (.xlsx)` o `Exportar Data Plana SP`) con icono descriptivo y validación cuando no haya datos cargados.
- **Descarga estándar con SheetJS:** Implementar la función de exportación (`exportarDataPlanaSQL()`) convirtiendo los datos planos a hoja Excel con `XLSX.utils.json_to_sheet(rawSqlData)` y descargando el archivo `.xlsx` con nombre estandarizado `Data_Plana_SQL_<Modulo>_<RangoFechas>.xlsx`.
