# Regla General: Resolución Dinámica de Rutas de Archivos Excel en OneDrive (Backend .NET)

Todo controlador o servicio en C# .NET (`ApiDesposte/Controllers/Excel/`) que se conecte o lea archivos Excel (`.xlsx`, `.xlsm`) sincronizados localmente mediante OneDrive / SharePoint **debe resolver la ruta del archivo de manera dinámica** en función del perfil del usuario del sistema operativo donde se ejecuta la API.

---

## 1. Estándar Obligatorio (`ObtenerRutaExcel`)

Queda **estrictamente prohibido** hardcodear rutas absolutas con nombres de usuario fijos (por ejemplo, `C:\Users\dennis.mendoza\...` o `C:\Users\Dennis\...`).

Siempre se debe implementar el método privado `ObtenerRutaExcel()` utilizando `Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)` y `Path.Combine`:

```csharp
private string ObtenerRutaExcel()
{
    string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
    return Path.Combine(userProfile, @"OneDrive - Corporación Rico SAC\Desposte-03\PlantasCore\Planificacion\Cumplimiento_PTC.xlsm");
}
```

*(Adaptando únicamente la subruta relativa dentro de `OneDrive - Corporación Rico SAC\...` según el archivo Excel que corresponda).*

---

## 2. Reglas de Lectura Segura de Archivos `.xlsx` / `.xlsm`

1. **Apertura sin bloqueo (`FileShare.ReadWrite | FileShare.Delete`):**
   Para permitir que los usuarios abran, editen o guarden el archivo en Excel o que OneDrive lo sincronice sin generar excepciones `IOException` por bloqueo de archivo, siempre copiar primero el archivo a un `MemoryStream`:
   ```csharp
   using var fileStream = new FileStream(rutaExcel, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
   ```
2. **Limpieza preventiva de `PivotTableParts` en memoria:**
   En libros `.xlsm` o `.xlsx` que contengan tablas dinámicas complejas, limpiar las partes de `PivotTableParts` y `PivotTableCacheDefinitionParts` sobre el `MemoryStream` con `SpreadsheetDocument.Open(msModificado, true)` antes de instanciar `new XLWorkbook(msModificado)` para evitar errores internos de ClosedXML.
3. **Endpoints de Tablas Planas (Consumo Crudo para JS):**
   Cuando se solicite exponer tablas de hojas de Excel para ser procesadas directamente en JavaScript, el endpoint debe leer las columnas dinámicamente a partir de `tabla.HeadersRow().Cells()` y devolver los registros **tal cual están en la tabla original**, sin inyectar ni calcular campos adicionales en el backend.

---

## 3. Uso Obligatorio de Funciones Compartidas (`LectorTablasExcel` y `CacheTablasExcel`)

Para evitar duplicar código de bajo nivel en los controladores, es **mandatorio** utilizar los componentes centralizados ubicados en `ApiDesposte/Funciones/Excel/`:

1. **Apertura segura en memoria:**
   - Usar `LectorTablasExcel.AbrirEnMemoria(rutaExcel)`.
   - Queda estrictamente prohibido instanciar `FileStream`, `MemoryStream` o `SpreadsheetDocument` manualmente dentro de los controladores.
2. **Lectura dinámica de tablas:**
   - Usar `LectorTablasExcel.LeerTablaTalCual(workbook, nombreHoja, nombreTabla)`.
   - Descubre dinámicamente las columnas desde `HeadersRow()`, soporta fallback a `RangeUsed()` en hojas sin tabla formal y mapea valores tipados (`Dictionary<string, object?>`).
3. **Caché en memoria concurrente:**
   - Para controladores que expongan colecciones de tablas completas, instanciar `private static readonly CacheTablasExcel Cache = new(TimeSpan.FromSeconds(60));` y llamar a `Cache.Obtener(rutaExcel, MapeoTablas, forzarRecarga)`.
4. **Responsabilidad del Controlador:**
   - Cada controlador debe ser delgado (*Thin Controller*): solo resuelve la ruta con `ObtenerRutaExcel()`, define el diccionario `MapeoTablas` y devuelve las respuestas HTTP.

