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

### 5. Matrices Jerárquicas sin Iconos Decorativos (Regla Obligatoria)
Toda tabla o matriz jerárquica de los dashboards en `Docu_Paginas/` debe mostrar sus etiquetas como **texto plano**:
- **Único icono permitido:** el control de expansión/colapso, porque comunica una acción. Cualquier icono, emoji o pictograma meramente decorativo se elimina.
- **Etiquetas planas:** los niveles se muestran sin fondo, borde, sombra ni decoración adicional; sólo tipografía y colores de la paleta.
- **Limpieza por CSS:** las decoraciones se anulan con reglas de hoja de estilos, sin modificar la lógica que genera las filas.
- **Jerarquía por estructura:** la profundidad se comunica con sangría, peso tipográfico y color, no con iconos repetidos en cada fila.
- **Reparto de anchos:** la columna de jerarquía absorbe el espacio sobrante y las columnas numéricas se ajustan a su contenido, para que el nombre de la fila se muestre completo.
