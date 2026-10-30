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
- Enlazar `<link rel="stylesheet" href="responsive.css">` en el `<head>` junto con un bloque `@media` autónomo en `<style>`.
- Usar `min-width: 0` en hijos de CSS Grid y grillas fluidas (`auto-fit` / `minmax`) para evitar desbordes horizontales.
- Compactar espacios y alturas en pantallas tipo laptop (`<= 1440px`) y adaptar filtros, gráficos y columnas fijas (`.sticky-col`) en tablets y celulares (`<= 768px`).
