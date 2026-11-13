using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System.IO;

namespace ApiDesposte.Controllers.Excel
{
    [ApiController]
    [Route("api/[controller]")]
    public class IndicadoresN3Controller : ControllerBase
    {
        private const string NombreHoja = "Registro";
        private const string NombreTabla = "tblRegistro";

        // Cache en memoria para optimizar lecturas consecutivas desde el frontend JS
        private static List<Dictionary<string, object?>>? _cacheRegistros = null;
        private static DateTime _cacheTimestamp = DateTime.MinValue;
        private static readonly object _cacheLock = new();
        private static readonly TimeSpan CacheDuracion = TimeSpan.FromSeconds(60);

        private string ObtenerRutaExcel()
        {
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return Path.Combine(userProfile, @"OneDrive - Corporación Rico SAC\Desposte-03\PlantasCore\Indicadores\BD Indicadores N3.xlsx");
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
                            foreach (var p in wsPart.PivotTableParts.ToList())
                            {
                                wsPart.DeletePart(p);
                            }
                        }

                        foreach (var cache in doc.WorkbookPart.PivotTableCacheDefinitionParts.ToList())
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
        /// Lee la tabla tal cual está definida en la hoja, con sus encabezados y valores originales,
        /// sin agregar campos calculados.
        /// </summary>
        private static List<Dictionary<string, object?>> LeerTablaExcelTalCual(XLWorkbook workbook, string nombreHoja, string nombreTabla)
        {
            var registros = new List<Dictionary<string, object?>>();

            if (!workbook.Worksheets.TryGetWorksheet(nombreHoja, out var hoja))
            {
                hoja = workbook.Worksheets.FirstOrDefault(w => string.Equals(w.Name.Trim(), nombreHoja.Trim(), StringComparison.OrdinalIgnoreCase));
            }

            IXLTable? tabla = hoja?.Tables.FirstOrDefault(t => string.Equals(t.Name.Trim(), nombreTabla.Trim(), StringComparison.OrdinalIgnoreCase));

            tabla ??= workbook.Worksheets.SelectMany(w => w.Tables)
                .FirstOrDefault(t => string.Equals(t.Name.Trim(), nombreTabla.Trim(), StringComparison.OrdinalIgnoreCase));

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
                        return celda.GetTimeSpan().ToString(@"hh\:mm\:ss");

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

        private List<Dictionary<string, object?>> CargarRegistros(bool forzarRecarga)
        {
            lock (_cacheLock)
            {
                if (!forzarRecarga && _cacheRegistros != null && (DateTime.UtcNow - _cacheTimestamp) < CacheDuracion)
                {
                    return _cacheRegistros;
                }

                string rutaExcel = ObtenerRutaExcel();
                if (!System.IO.File.Exists(rutaExcel))
                {
                    throw new FileNotFoundException($"No se encontró el archivo Excel en la ruta: {rutaExcel}");
                }

                using var workbook = CargarWorkbookEnMemoria(rutaExcel);

                _cacheRegistros = LeerTablaExcelTalCual(workbook, NombreHoja, NombreTabla);
                _cacheTimestamp = DateTime.UtcNow;

                return _cacheRegistros;
            }
        }

        /// <summary>
        /// Devuelve la tabla tblRegistro (hoja Registro) de BD Indicadores N3.xlsx tal cual está en el Excel.
        /// GET /api/IndicadoresN3?recargar=true
        /// </summary>
        [HttpGet]
        public IActionResult ObtenerRegistro([FromQuery] bool recargar = false)
        {
            try
            {
                return Ok(CargarRegistros(recargar));
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
