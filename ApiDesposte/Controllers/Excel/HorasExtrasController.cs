using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System.IO;

namespace ApiDesposte.Controllers.Excel
{
    [ApiController]
    [Route("api/[controller]")]
    public class HorasExtrasController : ControllerBase
    {
        private const string HojaDBHoras = "DBHoras";
        private const string TablaHoras = "T_Horas";

        // Cache en memoria para evitar relecturas continuas del archivo .xlsm
        private static List<HorasExtraItemDto>? _cacheItems = null;
        private static DateTime _cacheTimestamp = DateTime.MinValue;
        private static readonly object _cacheLock = new();
        private static readonly TimeSpan CacheDuracion = TimeSpan.FromMinutes(15);

        private string ObtenerRutaExcel()
        {
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return Path.Combine(userProfile, @"OneDrive - Corporación Rico SAC\Desposte-03\Compartido_PBC\Informe_Rendimientos.xlsm");
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

        private static IXLTable? ObtenerTabla(IXLWorksheet ws, string nombreTabla)
        {
            try
            {
                return ws.Table(nombreTabla);
            }
            catch
            {
                return ws.Tables.FirstOrDefault(t => string.Equals(t.Name, nombreTabla, StringComparison.OrdinalIgnoreCase));
            }
        }

        private static int BuscarIndiceColumna(IXLTable tabla, params string[] nombresPosibles)
        {
            var headers = tabla.HeadersRow().Cells().ToList();
            foreach (var nombre in nombresPosibles)
            {
                for (int i = 0; i < headers.Count; i++)
                {
                    string headerName = headers[i].Value.ToString().Trim();
                    if (string.Equals(headerName, nombre, StringComparison.OrdinalIgnoreCase))
                    {
                        return i + 1; // ClosedXML usa índices basados en 1
                    }
                }
            }
            return -1;
        }

        private static string ObtenerValorCeldaTexto(IXLRangeRow fila, int colIndex, string valorPorDefecto = "")
        {
            if (colIndex <= 0) return valorPorDefecto;
            var celda = fila.Cell(colIndex);
            if (celda.IsEmpty()) return valorPorDefecto;
            return celda.Value.ToString().Trim();
        }

        private static int? ObtenerValorCeldaIntNullable(IXLRangeRow fila, int colIndex)
        {
            if (colIndex <= 0) return null;
            var celda = fila.Cell(colIndex);
            if (celda.IsEmpty()) return null;

            if (celda.DataType == XLDataType.Number)
            {
                return (int)Math.Round(celda.GetValue<double>());
            }

            string txt = celda.Value.ToString().Trim();
            if (int.TryParse(txt, out int val))
            {
                return val;
            }

            return null;
        }

        private static double? ObtenerValorCeldaNumeroNullable(IXLRangeRow fila, int colIndex)
        {
            if (colIndex <= 0) return null;
            var celda = fila.Cell(colIndex);
            if (celda.IsEmpty()) return null;

            if (celda.DataType == XLDataType.Number)
            {
                return celda.GetValue<double>();
            }

            string txt = celda.Value.ToString().Trim().Replace(",", ".");
            if (double.TryParse(txt, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double val))
            {
                return val;
            }

            return null;
        }

        private static string? ObtenerValorCeldaFechaTexto(IXLRangeRow fila, int colIndex)
        {
            if (colIndex <= 0) return null;
            var celda = fila.Cell(colIndex);
            if (celda.IsEmpty()) return null;

            if (celda.DataType == XLDataType.DateTime)
            {
                return celda.GetDateTime().ToString("yyyy-MM-dd");
            }

            if (celda.DataType == XLDataType.Number)
            {
                double numVal = celda.GetValue<double>();
                if (numVal > 30000 && numVal < 70000)
                {
                    try
                    {
                        return DateTime.FromOADate(numVal).ToString("yyyy-MM-dd");
                    }
                    catch
                    {
                        // Fallback a texto si no es fecha válida OLE Automation
                    }
                }
            }

            string txt = celda.Value.ToString().Trim();
            if (DateTime.TryParse(txt, out var dt))
            {
                return dt.ToString("yyyy-MM-dd");
            }

            return string.IsNullOrWhiteSpace(txt) ? null : txt;
        }

        private List<HorasExtraItemDto> CargarDatosTablaHoras(bool forzarRecarga = false)
        {
            lock (_cacheLock)
            {
                if (!forzarRecarga && _cacheItems != null && (DateTime.Now - _cacheTimestamp) < CacheDuracion)
                {
                    return _cacheItems;
                }

                string rutaExcel = ObtenerRutaExcel();
                if (!System.IO.File.Exists(rutaExcel))
                {
                    throw new FileNotFoundException($"No se encontró el archivo Excel en la ruta: {rutaExcel}");
                }

                using var workbook = CargarWorkbookEnMemoria(rutaExcel);

                if (!workbook.Worksheets.TryGetWorksheet(HojaDBHoras, out var ws))
                {
                    throw new InvalidOperationException($"No se encontró la hoja '{HojaDBHoras}' en el libro.");
                }

                var tabla = ObtenerTabla(ws, TablaHoras);
                if (tabla == null)
                {
                    throw new InvalidOperationException($"No se encontró la tabla '{TablaHoras}' en la hoja '{HojaDBHoras}'.");
                }

                int colCecosCod = BuscarIndiceColumna(tabla, "CecosCod", "Ceco", "Cecos", "CodCeco");
                int colNCentroCosto = BuscarIndiceColumna(tabla, "NCentroCosto", "CentroCosto", "NombreCentroCosto");
                int colCodeEmp = BuscarIndiceColumna(tabla, "CodeEmp", "CodEmp", "CodigoEmpleado", "Dni");
                int colNombres = BuscarIndiceColumna(tabla, "Nombres", "Nombre", "Trabajador", "Colaborador");
                int colAno = BuscarIndiceColumna(tabla, "Año", "Ano", "Anio", "Year");
                int colMes = BuscarIndiceColumna(tabla, "Mes", "Month");
                int colFecha = BuscarIndiceColumna(tabla, "Fecha", "Date", "Fec");
                int colTrimestre = BuscarIndiceColumna(tabla, "Trimestre", "Trim", "Q");
                int colHoras = BuscarIndiceColumna(tabla, "Horas", "TotalHoras", "CantHoras");
                int colHNor = BuscarIndiceColumna(tabla, "HNor", "H_Nor", "HorasNormales");
                int colH25 = BuscarIndiceColumna(tabla, "H25", "H_25", "Horas25");
                int colH35 = BuscarIndiceColumna(tabla, "H35", "H_35", "Horas35");
                int colH100 = BuscarIndiceColumna(tabla, "H100", "H_100", "Horas100");

                var items = new List<HorasExtraItemDto>();

                foreach (var fila in tabla.DataRange.Rows())
                {
                    if (fila.IsEmpty()) continue;

                    string cecosCod = ObtenerValorCeldaTexto(fila, colCecosCod);
                    string nCentroCosto = ObtenerValorCeldaTexto(fila, colNCentroCosto);
                    string codeEmp = ObtenerValorCeldaTexto(fila, colCodeEmp);
                    string nombres = ObtenerValorCeldaTexto(fila, colNombres);
                    int? ano = ObtenerValorCeldaIntNullable(fila, colAno);
                    string mes = ObtenerValorCeldaTexto(fila, colMes);
                    string? fecha = ObtenerValorCeldaFechaTexto(fila, colFecha);
                    string? trimestre = ObtenerValorCeldaTexto(fila, colTrimestre);
                    double? horas = ObtenerValorCeldaNumeroNullable(fila, colHoras);
                    double? hNor = ObtenerValorCeldaNumeroNullable(fila, colHNor);
                    double? h25 = ObtenerValorCeldaNumeroNullable(fila, colH25);
                    double? h35 = ObtenerValorCeldaNumeroNullable(fila, colH35);
                    double? h100 = ObtenerValorCeldaNumeroNullable(fila, colH100);

                    // Omitir fila si todos los campos identificadores y numéricos están vacíos
                    if (string.IsNullOrWhiteSpace(cecosCod) &&
                        string.IsNullOrWhiteSpace(codeEmp) &&
                        string.IsNullOrWhiteSpace(nombres) &&
                        !horas.HasValue)
                    {
                        continue;
                    }

                    items.Add(new HorasExtraItemDto
                    {
                        CecosCod = cecosCod,
                        NCentroCosto = nCentroCosto,
                        CodeEmp = codeEmp,
                        Nombres = nombres,
                        Ano = ano,
                        Mes = mes,
                        Fecha = fecha,
                        Trimestre = string.IsNullOrWhiteSpace(trimestre) ? null : trimestre,
                        Horas = horas.HasValue ? Math.Round(horas.Value, 4) : null,
                        HNor = hNor.HasValue ? Math.Round(hNor.Value, 4) : null,
                        H25 = h25.HasValue ? Math.Round(h25.Value, 4) : null,
                        H35 = h35.HasValue ? Math.Round(h35.Value, 4) : null,
                        H100 = h100.HasValue ? Math.Round(h100.Value, 4) : null
                    });
                }

                _cacheItems = items;
                _cacheTimestamp = DateTime.Now;

                return _cacheItems;
            }
        }

        /// <summary>
        /// Endpoint único para consultar los datos de la tabla T_Horas (hoja DBHoras).
        /// </summary>
        [HttpGet]
        public IActionResult ObtenerHorasExtras([FromQuery] bool recargar = false)
        {
            try
            {
                var datos = CargarDatosTablaHoras(recargar);
                return Ok(datos);
            }
            catch (FileNotFoundException ex)
            {
                return NotFound(new { exito = false, error = ex.Message });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }
    }

    public class HorasExtraItemDto
    {
        public string CecosCod { get; set; } = string.Empty;
        public string NCentroCosto { get; set; } = string.Empty;
        public string CodeEmp { get; set; } = string.Empty;
        public string Nombres { get; set; } = string.Empty;
        public int? Ano { get; set; }
        public int? Año => Ano;
        public string Mes { get; set; } = string.Empty;
        public string? Fecha { get; set; }
        public string? Trimestre { get; set; }
        public double? Horas { get; set; }
        public double? HNor { get; set; }
        public double? H25 { get; set; }
        public double? H35 { get; set; }
        public double? H100 { get; set; }
    }
}
