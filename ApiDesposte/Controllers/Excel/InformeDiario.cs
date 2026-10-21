using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System.IO;

namespace ApiDesposte.Controllers.Excel
{
    [ApiController]
    [Route("api/[controller]")]
    [Route("api/informe-diario")]
    public class InformeDiarioController : ControllerBase
    {
        // Mapeo de las 6 Hojas y Tablas del Excel
        private static readonly Dictionary<string, (string Hoja, string Tabla)> MapeoTablas = new(StringComparer.OrdinalIgnoreCase)
        {
            { "personal", ("BD_PERSONAL", "TablaPersonal") },
            { "sso", ("BD_SSO", "TablaSSO") },
            { "horarios", ("BD_HORARIOS", "TablaHorarios") },
            { "produccion", ("BD_PRODUCCION", "TablaProduccion") },
            { "detalle", ("BD_DETALLE", "TablaDetalle") },
            { "paradas", ("BD_PARADAS", "TablaParadas") }
        };

        // Cache en memoria para optimizar peticiones seguidas desde el dashboard HTML
        private static Dictionary<string, List<Dictionary<string, object?>>>? _cacheTablas = null;
        private static DateTime _cacheTimestamp = DateTime.MinValue;
        private static readonly object _cacheLock = new();
        private static readonly TimeSpan CacheDuracion = TimeSpan.FromSeconds(60);

        private string ObtenerRutaExcel()
        {
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return Path.Combine(userProfile, @"OneDrive - Corporación Rico SAC\Desposte-04\Informes Diario\Informe Diario 2026 - Base de Datos.xlsx");
        }

        private XLWorkbook CargarWorkbookEnMemoria(string rutaExcel)
        {
            byte[] fileBytes;
            using (var fileStream = new FileStream(rutaExcel, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var ms = new MemoryStream())
            {
                fileStream.CopyTo(ms);
                fileBytes = ms.ToArray();
            }

            var msModificado = new MemoryStream();
            msModificado.Write(fileBytes, 0, fileBytes.Length);
            msModificado.Position = 0;

            try
            {
                // Limpieza de PivotTables para evitar incompatibilidades en ClosedXML
                using (var doc = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(msModificado, true))
                {
                    if (doc.WorkbookPart != null)
                    {
                        foreach (var wsPart in doc.WorkbookPart.WorksheetParts)
                        {
                            var pParts = wsPart.PivotTableParts.ToList();
                            foreach (var p in pParts)
                            {
                                wsPart.DeletePart(p);
                            }
                        }

                        var pivotCaches = doc.WorkbookPart.PivotTableCacheDefinitionParts.ToList();
                        foreach (var cache in pivotCaches)
                        {
                            doc.WorkbookPart.DeletePart(cache);
                        }

                        if (doc.WorkbookPart.Workbook != null)
                        {
                            doc.WorkbookPart.Workbook.PivotCaches?.Remove();
                            doc.WorkbookPart.Workbook.Save();
                        }
                    }
                }
            }
            catch
            {
                msModificado = new MemoryStream(fileBytes);
            }

            msModificado.Position = 0;
            return new XLWorkbook(msModificado);
        }

        private static List<Dictionary<string, object?>> LeerTablaExcel(XLWorkbook workbook, string nombreHoja, string nombreTabla)
        {
            var registros = new List<Dictionary<string, object?>>();

            if (!workbook.Worksheets.TryGetWorksheet(nombreHoja, out var hoja))
            {
                hoja = workbook.Worksheets.FirstOrDefault(w => string.Equals(w.Name.Trim(), nombreHoja.Trim(), StringComparison.OrdinalIgnoreCase));
            }

            IXLTable? tabla = null;

            if (hoja != null)
            {
                tabla = hoja.Tables.FirstOrDefault(t => string.Equals(t.Name.Trim(), nombreTabla.Trim(), StringComparison.OrdinalIgnoreCase));
            }

            if (tabla == null)
            {
                tabla = workbook.Worksheets.SelectMany(w => w.Tables)
                    .FirstOrDefault(t => string.Equals(t.Name.Trim(), nombreTabla.Trim(), StringComparison.OrdinalIgnoreCase));
            }

            // Lectura de tabla oficial estructurada
            if (tabla != null)
            {
                var headers = tabla.HeadersRow().Cells().Select(c => c.Value.ToString().Trim()).ToList();

                foreach (var fila in tabla.DataRange.Rows())
                {
                    if (fila.IsEmpty()) continue;

                    var filaDict = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                    for (int i = 0; i < headers.Count; i++)
                    {
                        string columna = string.IsNullOrWhiteSpace(headers[i]) ? $"Columna_{i + 1}" : headers[i];
                        filaDict[columna] = ObtenerValorCelda(fila.Cell(i + 1));
                    }
                    registros.Add(filaDict);
                }

                return registros;
            }

            // Fallback si los datos no están en una tabla estructurada de Excel pero sí en la hoja
            if (hoja != null)
            {
                var rango = hoja.RangeUsed();
                if (rango != null && rango.RowCount() > 1)
                {
                    var headers = rango.FirstRow().Cells().Select(c => c.Value.ToString().Trim()).ToList();
                    foreach (var fila in rango.Rows().Skip(1))
                    {
                        if (fila.IsEmpty()) continue;

                        var filaDict = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                        for (int i = 0; i < headers.Count; i++)
                        {
                            string columna = string.IsNullOrWhiteSpace(headers[i]) ? $"Columna_{i + 1}" : headers[i];
                            filaDict[columna] = ObtenerValorCelda(fila.Cell(i + 1));
                        }
                        registros.Add(filaDict);
                    }
                }
            }

            return registros;
        }

        private static object? ObtenerValorCelda(IXLCell celda)
        {
            if (celda.IsEmpty()) return null;

            try
            {
                switch (celda.DataType)
                {
                    case XLDataType.Boolean:
                        return celda.GetBoolean();

                    case XLDataType.Number:
                        double num = celda.GetDouble();
                        if (Math.Abs(num % 1) < double.Epsilon && num >= long.MinValue && num <= long.MaxValue)
                            return (long)num;
                        return Math.Round(num, 4);

                    case XLDataType.DateTime:
                        var dt = celda.GetDateTime();
                        return dt.TimeOfDay.TotalSeconds == 0
                            ? dt.ToString("yyyy-MM-dd")
                            : dt.ToString("yyyy-MM-dd HH:mm:ss");

                    case XLDataType.TimeSpan:
                        var ts = celda.GetTimeSpan();
                        return ts.ToString(@"hh\:mm\:ss");

                    default:
                        var str = celda.Value.ToString().Trim();
                        return string.IsNullOrEmpty(str) ? null : str;
                }
            }
            catch
            {
                var fallback = celda.Value.ToString().Trim();
                return string.IsNullOrEmpty(fallback) ? null : fallback;
            }
        }

        /// <summary>
        /// ENDPOINT ÚNICO: Retorna las 6 tablas en un solo JSON para el dashboard HTML.
        /// GET /api/InformeDiario o GET /api/informe-diario
        /// Opcional: ?forzarRecarga=true para saltar la caché en memoria.
        /// </summary>
        [HttpGet]
        public IActionResult ObtenerInformeDiario([FromQuery] bool forzarRecarga = false)
        {
            try
            {
                lock (_cacheLock)
                {
                    if (!forzarRecarga && _cacheTablas != null && (DateTime.UtcNow - _cacheTimestamp) < CacheDuracion)
                    {
                        return Ok(new
                        {
                            exito = true,
                            tablas = _cacheTablas
                        });
                    }

                    string rutaExcel = ObtenerRutaExcel();
                    if (!System.IO.File.Exists(rutaExcel))
                    {
                        return NotFound(new { exito = false, error = $"No se encontró el archivo Excel en la ruta: {rutaExcel}" });
                    }

                    using var workbook = CargarWorkbookEnMemoria(rutaExcel);

                    var resultado = new Dictionary<string, List<Dictionary<string, object?>>>(StringComparer.OrdinalIgnoreCase);

                    foreach (var item in MapeoTablas)
                    {
                        resultado[item.Key] = LeerTablaExcel(workbook, item.Value.Hoja, item.Value.Tabla);
                    }

                    _cacheTablas = resultado;
                    _cacheTimestamp = DateTime.UtcNow;

                    return Ok(new
                    {
                        exito = true,
                        tablas = _cacheTablas
                    });
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = $"Error al leer el archivo Excel: {ex.Message}" });
            }
        }
    }
}
