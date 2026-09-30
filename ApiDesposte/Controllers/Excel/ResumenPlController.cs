using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System.IO;

namespace ApiDesposte.Controllers.Excel
{
    [ApiController]
    [Route("api/[controller]")]
    public class ResumenPlController : ControllerBase
    {
        private const string HojaResumenPL = "ResumenPL";
        private const string TablaResumenPL = "T_ResumenPL";

        // Cache en memoria para evitar relecturas continuas del archivo .xlsm
        private static List<ResumenPlItemDto>? _cacheItems = null;
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
                        return i + 1; // ClosedXML es 1-based index
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

        private List<ResumenPlItemDto> CargarDatosTablaResumenPL(bool forzarRecarga = false)
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

                if (!workbook.Worksheets.TryGetWorksheet(HojaResumenPL, out var ws))
                {
                    throw new InvalidOperationException($"No se encontró la hoja '{HojaResumenPL}' en el libro.");
                }

                var tabla = ObtenerTabla(ws, TablaResumenPL);
                if (tabla == null)
                {
                    throw new InvalidOperationException($"No se encontró la tabla '{TablaResumenPL}' en la hoja '{HojaResumenPL}'.");
                }

                int colSemana = BuscarIndiceColumna(tabla, "Semana");
                int colGrupo = BuscarIndiceColumna(tabla, "Grupo", "SubPlanta", "Area");
                int colTipo = BuscarIndiceColumna(tabla, "Tipo", "TipoRegistro");
                int colUnidad = BuscarIndiceColumna(tabla, "Unidad", "Und", "Uom");
                int colCantidad = BuscarIndiceColumna(tabla, "Cantidad", "Kilos", "Kgs");

                var items = new List<ResumenPlItemDto>();

                foreach (var fila in tabla.DataRange.Rows())
                {
                    if (fila.IsEmpty()) continue;

                    string semTxt = ObtenerValorCeldaTexto(fila, colSemana);
                    if (string.IsNullOrWhiteSpace(semTxt)) continue;

                    int.TryParse(semTxt, out int semNum);
                    string grupo = ObtenerValorCeldaTexto(fila, colGrupo);
                    string tipo = ObtenerValorCeldaTexto(fila, colTipo);
                    string unidad = ObtenerValorCeldaTexto(fila, colUnidad);
                    double? cantidad = ObtenerValorCeldaNumeroNullable(fila, colCantidad);

                    items.Add(new ResumenPlItemDto
                    {
                        Semana = semNum,
                        Grupo = grupo,
                        Tipo = tipo,
                        Unidad = string.IsNullOrWhiteSpace(unidad) ? "KG" : unidad.ToUpper(),
                        Cantidad = cantidad.HasValue ? Math.Round(cantidad.Value, 4) : null
                    });
                }

                _cacheItems = items;
                _cacheTimestamp = DateTime.Now;

                return _cacheItems;
            }
        }

        /// <summary>
        /// Endpoint principal: Devuelve todos los registros de T_ResumenPL con filtros opcionales.
        /// </summary>
        [HttpGet]
        public IActionResult ObtenerResumenPL(
            [FromQuery] int? semana = null,
            [FromQuery] string? grupo = null,
            [FromQuery] string? tipo = null,
            [FromQuery] bool recargar = false)
        {
            try
            {
                var datos = CargarDatosTablaResumenPL(recargar);

                if (semana.HasValue)
                {
                    datos = datos.Where(x => x.Semana == semana.Value).ToList();
                }

                if (!string.IsNullOrWhiteSpace(grupo))
                {
                    var gruposFiltro = grupo.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                            .Select(g => g.Trim().ToUpper())
                                            .ToHashSet();
                    datos = datos.Where(x => gruposFiltro.Contains(x.Grupo.ToUpper())).ToList();
                }

                if (!string.IsNullOrWhiteSpace(tipo))
                {
                    var tiposFiltro = tipo.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                          .Select(t => t.Trim().ToUpper())
                                          .ToHashSet();
                    datos = datos.Where(x => tiposFiltro.Contains(x.Tipo.ToUpper())).ToList();
                }

                return Ok(datos);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }

        /// <summary>
        /// Devuelve el catálogo de Grupos disponibles (DESPOSTE, EMBUTIDOS, CHULETAS, etc.).
        /// </summary>
        [HttpGet("grupos")]
        public IActionResult ObtenerGrupos()
        {
            try
            {
                var datos = CargarDatosTablaResumenPL();
                var grupos = datos.Select(x => x.Grupo)
                                  .Where(g => !string.IsNullOrWhiteSpace(g))
                                  .Distinct()
                                  .OrderBy(g => g)
                                  .ToList();

                return Ok(grupos);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }
    }

    public class ResumenPlItemDto
    {
        public int Semana { get; set; }
        public string Grupo { get; set; } = string.Empty;
        public string Tipo { get; set; } = string.Empty;
        public string Unidad { get; set; } = "KG";
        public double? Cantidad { get; set; }
    }
}