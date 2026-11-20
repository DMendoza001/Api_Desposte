# 4. Checklist: endpoint + dashboard nuevo

Usar esta lista cada vez que se agregue un informe.

## API .NET

- [ ] El controlador está en `Controllers/Sql/` o `Controllers/Excel/` según su fuente.
- [ ] El nombre de la clase coincide con el archivo: `<Nombre>Controller.cs` → `class <Nombre>Controller`.
- [ ] Un solo `[Route("api/[controller]")]` y un solo atributo de verbo por acción.
- [ ] **SQL:** usa la cadena de conexión de `appsettings.json`; no hay usuarios ni contraseñas en el código.
- [ ] **Excel:** la ruta se resuelve con `ObtenerRutaExcel()` y el perfil del usuario; la lectura usa `LectorTablasExcel` / `CacheTablasExcel`.
- [ ] `dotnet build` termina con 0 errores.
- [ ] El endpoint responde en el navegador: `http://localhost:8080/api/<Nombre>`.

## Dashboard HTML

- [ ] Está en `Docu_Paginas/` e incluye `config.js` y `responsive.css?v=3` en el `<head>`.
- [ ] Tiene el bloque `APP_CONFIG`, `generarBasesApi()`, `fetchConTimeout()` y `detectarServidorApi()`.
- [ ] `APP_CONFIG.controllerPath` apunta a la ruta canónica del controlador.
- [ ] Usa la paleta `PaletaFinanza_Industrial.json` (cabecera oscura con texto claro).
- [ ] Las 4 tarjetas KPI quedan en una fila en laptop; se ve bien en tablet (≤ 980px) y celular (≤ 768px).
- [ ] Los gráficos están dentro de `.chart-box` / `.chart-container`, sin alturas fijas inline.
- [ ] Las semanas van de lunes a domingo. En el eje X de los gráficos se muestran como `Sem XX`; el rango `Sem. XX (DD/MM al DD/MM)` va en el tooltip, tablas y filtros.
- [ ] **Fuente SQL:** guarda `rawSqlData` y tiene el botón `Data Plana SQL (.xlsx)`.
- [ ] **Matriz jerárquica:** etiquetas planas; el único icono es el de expandir/colapsar.
- [ ] **Vistas Diario / Semanal / Mensual:** filtros de periodo según la regla 9.

## Publicación

- [ ] Se agregó la entrada en `LISTA_DASHBOARDS` de `index.html`.
- [ ] Se agregó el archivo en `dashboards_estado.json` (`true` = visible, `false` = en prueba).
- [ ] Se probó desde otra PC de la red: `http://<SERVIDOR>:8080/<Dashboard>.html`.
- [ ] Commit con un mensaje que describa el informe.
