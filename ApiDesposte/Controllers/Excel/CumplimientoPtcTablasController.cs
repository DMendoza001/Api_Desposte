using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System.IO;

namespace ApiDesposte.Controllers.Excel
{
    [ApiController]
    [Route("api/[controller]")]
    [Route("api/cumplimiento-ptc-tablas")]
    public class CumplimientoPtcTablasController : ControllerBase
    {
        // Mapeo oficial de Hojas y Tablas en Cumplimiento_PTC.xlsm
        private static readonly Dictionary<string, (string Hoja, string Tabla)> MapeoTablas = new(StringComparer.OrdinalIgnoreCase)
        {
            { "T_Produccion", ("Produccion", "T_Produccion") },
            { "T_DemandaPy", ("DemandaPy", "T_DemandaPy") },
            { "T_ResProd_Py_DB", ("ResProd_Py_DB", "T_ResProd_Py_DB") },
            { "T_Distribucion", ("Distribucion", "T_Distribucion") },
            { "T_CodigoRelacion", ("Auxiliares", "T_CodigoRelacion") }
        };

        // Alias por nombre de hoja o clave corta para facilitar consumo desde JS
        private static readonly Dictionary<string, string> AliasTablas = new(StringComparer.OrdinalIgnoreCase)
        {
            { "produccion", "T_Produccion" },
            { "t_produccion", "T_Produccion" },
            { "demandapy", "T_DemandaPy" },
            { "t_demandapy", "T_DemandaPy" },
            { "resprod_py_db", "T_ResProd_Py_DB" },
            { "resprod-py-db", "T_ResProd_Py_DB" },
            { "t_resprod_py_db", "T_ResProd_Py_DB" },
            { "distribucion", "T_Distribucion" },
            { "t_distribucion", "T_Distribucion" },
            { "auxiliares", "T_CodigoRelacion" },
            { "codigorelacion", "T_CodigoRelacion" },
            { "t_codigorelacion", "T_CodigoRelacion" }
        };

        // Cache en memoria para optimizar lecturas consecutivas desde el frontend JS
        private static Dictionary<string, List<Dictionary<string, object?>>>? _cacheTablas = null;
        private static DateTime _cacheTimestamp = DateTime.MinValue;
        private static readonly object _cacheLock = new();
        private static readonly TimeSpan CacheDuracion = TimeSpan.FromSeconds(60);

        private string ObtenerRutaExcel()
        {
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return Path.Combine(userProfile, @"OneDrive - Corporación Rico SAC\Desposte-03\PlantasCore\Planificacion\Cumplimiento_PTC.xlsm");
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
                // Limpieza de PivotTables para evitar conflictos al abrir con ClosedXML
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

        /// <summary>
        /// Lee una tabla estructurada de Excel tal cual está definida en la hoja,
        /// conservando exactamente sus encabezados y valores originales sin agregar campos calculados.
        /// </summary>
        private static List<Dictionary<string, object?>> LeerTablaExcelTalCual(XLWorkbook workbook, string nombreHoja, string nombreTabla)
        {
            var registros = new List<Dictionary<string, object?>>();

            if (!workbook.Worksheets.TryGetWorksheet(nombreHoja, out var hoja))
            {
                hoja = workbook.Worksheets.FirstOrDefault(w => string.Equals(w.Name.Trim(), nombreHoja.Trim(), StringComparison.OrdinalIgnoreCase));
            }

            IXLTable? tabla = null;

            if (hoja != null)
            {
                try
                {
                    tabla = hoja.Table(nombreTabla);
                }
                catch
                {
                    tabla = hoja.Tables.FirstOrDefault(t => string.Equals(t.Name.Trim(), nombreTabla.Trim(), StringComparison.OrdinalIgnoreCase));
                }
            }

            if (tabla == null)
            {
                tabla = workbook.Worksheets.SelectMany(w => w.Tables)
                    .FirstOrDefault(t => string.Equals(t.Name.Trim(), nombreTabla.Trim(), StringComparison.OrdinalIgnoreCase));
            }

            if (tabla == null)
            {
                throw new InvalidOperationException($"No se encontró la tabla '{nombreTabla}' en la hoja '{nombreHoja}'.");
            }

            var headers = tabla.HeadersRow().Cells().Select(c => c.Value.ToString().Trim()).ToList();

            foreach (var fila in tabla.DataRange.Rows())
            {
                if (fila.IsEmpty()) continue;

                var filaDict = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                bool tieneDatos = false;

                for (int i = 0; i < headers.Count; i++)
                {
                    string columna = string.IsNullOrWhiteSpace(headers[i]) ? $"Columna_{i + 1}" : headers[i];
                    object? valor = ObtenerValorCelda(fila.Cell(i + 1));
                    if (valor != null) tieneDatos = true;
                    filaDict[columna] = valor;
                }

                if (tieneDatos)
                {
                    registros.Add(filaDict);
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

        private Dictionary<string, List<Dictionary<string, object?>>> CargarTodasLasTablas(bool forzarRecarga = false)
        {
            lock (_cacheLock)
            {
                if (!forzarRecarga && _cacheTablas != null && (DateTime.UtcNow - _cacheTimestamp) < CacheDuracion)
                {
                    return _cacheTablas;
                }

                string rutaExcel = ObtenerRutaExcel();
                if (!System.IO.File.Exists(rutaExcel))
                {
                    throw new FileNotFoundException($"No se encontró el archivo Excel en la ruta: {rutaExcel}");
                }

                using var workbook = CargarWorkbookEnMemoria(rutaExcel);

                var resultado = new Dictionary<string, List<Dictionary<string, object?>>>(StringComparer.OrdinalIgnoreCase);

                foreach (var item in MapeoTablas)
                {
                    resultado[item.Key] = LeerTablaExcelTalCual(workbook, item.Value.Hoja, item.Value.Tabla);
                }

                _cacheTablas = resultado;
                _cacheTimestamp = DateTime.UtcNow;

                return _cacheTablas;
            }
        }

        /// <summary>
        /// Devuelve las 5 tablas (T_Produccion, T_DemandaPy, T_ResProd_Py_DB, T_Distribucion, T_CodigoRelacion) tal cual están en el Excel,
        /// sin agregar campos adicionales.
        /// GET /api/CumplimientoPtcTablas
        /// </summary>
        [HttpGet]
        public IActionResult ObtenerTodasLasTablas([FromQuery] bool recargar = false)
        {
            try
            {
                var tablas = CargarTodasLasTablas(recargar);
                return Ok(new
                {
                    exito = true,
                    tablas = tablas
                });
            }
            catch (FileNotFoundException fnfEx)
            {
                return NotFound(new { exito = false, error = fnfEx.Message });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }

        /// <summary>
        /// Devuelve una tabla específica tal cual está en el Excel (por nombre de tabla o de hoja:
        /// T_Produccion | Produccion | T_DemandaPy | DemandaPy | T_ResProd_Py_DB | ResProd_Py_DB | T_Distribucion | Distribucion | T_CodigoRelacion | Auxiliares).
        /// GET /api/CumplimientoPtcTablas/{nombreTabla}
        /// </summary>
        [HttpGet("{nombreTabla}")]
        public IActionResult ObtenerTablaPorNombre(string nombreTabla, [FromQuery] bool recargar = false)
        {
            try
            {
                if (!AliasTablas.TryGetValue(nombreTabla.Trim(), out var claveTablaOficial))
                {
                    return BadRequest(new
                    {
                        exito = false,
                        error = $"La tabla u hoja '{nombreTabla}' no está registrada. Valores permitidos: T_Produccion, T_DemandaPy, T_ResProd_Py_DB, T_Distribucion, T_CodigoRelacion."
                    });
                }

                var tablas = CargarTodasLasTablas(recargar);
                return Ok(tablas[claveTablaOficial]);
            }
            catch (FileNotFoundException fnfEx)
            {
                return NotFound(new { exito = false, error = fnfEx.Message });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }
    }
}

