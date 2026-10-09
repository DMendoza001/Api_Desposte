using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System.Globalization;
using System.IO;

namespace ApiDesposte.Controllers.Excel
{
    [ApiController]
    [Route("api/[controller]")]
    public class NumeroCerdosController : ControllerBase
    {
        private const string HojaPy13 = "Py13";
        private const string TablaPyCerdos = "T_PyCerdos";

        private string ObtenerRutaExcel()
        {
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return Path.Combine(userProfile, @"OneDrive - Corporación Rico SAC\Desposte-03\PlantasCore\Planificacion\NumeroCerdos.xlsx");
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
            for (int i = 0; i < headers.Count; i++)
            {
                string headerName = headers[i].Value.ToString().Trim();
                foreach (var nombre in nombresPosibles)
                {
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

        private static double ObtenerValorCeldaNumero(IXLRangeRow fila, int colIndex, double valorPorDefecto = 0.0)
        {
            if (colIndex <= 0) return valorPorDefecto;
            var celda = fila.Cell(colIndex);
            if (celda.IsEmpty()) return valorPorDefecto;

            if (celda.DataType == XLDataType.Number)
            {
                return celda.GetValue<double>();
            }

            string txt = celda.Value.ToString().Trim().Replace(",", ".");
            if (double.TryParse(txt, NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
            {
                return val;
            }

            if (double.TryParse(celda.Value.ToString().Trim(), out double valLocal))
            {
                return valLocal;
            }

            return valorPorDefecto;
        }

        /// <summary>
        /// Obtiene los datos de la tabla T_PyCerdos de la hoja Py13 del archivo NumeroCerdos.xlsx
        /// </summary>
        [HttpGet]
        public IActionResult ObtenerDatos(
            [FromQuery] string? semana = null,
            [FromQuery] string? codigoCorto = null)
        {
            try
            {
                string rutaExcel = ObtenerRutaExcel();
                if (!System.IO.File.Exists(rutaExcel))
                {
                    return NotFound(new { exito = false, mensaje = $"No se encontró el archivo Excel en la ruta: {rutaExcel}" });
                }

                using var workbook = CargarWorkbookEnMemoria(rutaExcel);

                if (!workbook.Worksheets.TryGetWorksheet(HojaPy13, out var wsPy13))
                {
                    return NotFound(new { exito = false, mensaje = $"No se encontró la hoja '{HojaPy13}' en el archivo." });
                }

                var tablaCerdos = ObtenerTabla(wsPy13, TablaPyCerdos);
                if (tablaCerdos == null)
                {
                    return NotFound(new { exito = false, mensaje = $"No se encontró la tabla '{TablaPyCerdos}' en la hoja '{HojaPy13}'." });
                }

                // Identificar índices de columnas
                int colSemana = BuscarIndiceColumna(tablaCerdos, "Semana");
                int colCodigoCorto = BuscarIndiceColumna(tablaCerdos, "CodigoCorto", "Codigo_Corto", "CodCorto");
                int colNombre = BuscarIndiceColumna(tablaCerdos, "Nombre", "Descripcion", "Articulo");
                int colKilos = BuscarIndiceColumna(tablaCerdos, "Kilos", "Kg", "Peso");
                int colWprom = BuscarIndiceColumna(tablaCerdos, "Wprom", "W_prom", "PesoPromedio", "Promedio");
                int colPiezas = BuscarIndiceColumna(tablaCerdos, "Piezas", "Pzs", "CantPiezas", "Unidades");
                int colFactor = BuscarIndiceColumna(tablaCerdos, "Factor", "Fact");
                int colNCerdos = BuscarIndiceColumna(tablaCerdos, "NCerdos", "N_Cerdos", "NumCerdos", "Cerdos");

                var lista = new List<NumeroCerdosDto>();

                foreach (var fila in tablaCerdos.DataRange.Rows())
                {
                    if (fila.IsEmpty()) continue;

                    string valSemana = ObtenerValorCeldaTexto(fila, colSemana);
                    string valCodigoCorto = ObtenerValorCeldaTexto(fila, colCodigoCorto);
                    string valNombre = ObtenerValorCeldaTexto(fila, colNombre);
                    double valKilos = ObtenerValorCeldaNumero(fila, colKilos);
                    double valWprom = ObtenerValorCeldaNumero(fila, colWprom);
                    double valPiezas = ObtenerValorCeldaNumero(fila, colPiezas);
                    double valFactor = ObtenerValorCeldaNumero(fila, colFactor);
                    double valNCerdos = ObtenerValorCeldaNumero(fila, colNCerdos);

                    // Si toda la fila está vacía, omitir
                    if (string.IsNullOrWhiteSpace(valSemana) &&
                        string.IsNullOrWhiteSpace(valCodigoCorto) &&
                        string.IsNullOrWhiteSpace(valNombre) &&
                        valKilos == 0 && valNCerdos == 0)
                    {
                        continue;
                    }

                    // Filtros opcionales
                    if (!string.IsNullOrWhiteSpace(semana) && !string.Equals(valSemana, semana.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    if (!string.IsNullOrWhiteSpace(codigoCorto) && !string.Equals(valCodigoCorto, codigoCorto.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    // Fórmulas de cálculo de respaldo si no vinieron precalculadas
                    if (valPiezas == 0 && valWprom > 0 && valKilos > 0)
                    {
                        valPiezas = valKilos / valWprom;
                    }
                    if (valNCerdos == 0 && valFactor > 0 && valPiezas > 0)
                    {
                        valNCerdos = valPiezas / valFactor;
                    }

                    int anio = 0;
                    int nSemana = 0;
                    if (!string.IsNullOrWhiteSpace(valSemana))
                    {
                        if (valSemana.Length >= 5)
                        {
                            int.TryParse(valSemana.Substring(0, 4), out anio);
                            int.TryParse(valSemana.Substring(4), out nSemana);
                        }
                        else
                        {
                            int.TryParse(valSemana, out nSemana);
                        }
                    }

                    lista.Add(new NumeroCerdosDto
                    {
                        Semana = valSemana,
                        Anio = anio,
                        NSemana = nSemana,
                        CodigoCorto = valCodigoCorto,
                        Nombre = valNombre,
                        Kilos = Math.Round(valKilos, 2),
                        Wprom = Math.Round(valWprom, 4),
                        Piezas = Math.Round(valPiezas, 2),
                        Factor = Math.Round(valFactor, 2),
                        NCerdos = Math.Round(valNCerdos, 2)
                    });
                }

                return Ok(lista.OrderBy(x => x.Semana).ToList());
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message, stackTrace = ex.ToString() });
            }
                }
    }

    public class NumeroCerdosDto
    {
        public string Semana { get; set; } = string.Empty;
        public int Anio { get; set; }
        public int NSemana { get; set; }
        public string CodigoCorto { get; set; } = string.Empty;
        public string Nombre { get; set; } = string.Empty;
        public double Kilos { get; set; }
        public double Wprom { get; set; }
        public double Piezas { get; set; }
        public double Factor { get; set; }
        public double NCerdos { get; set; }
    }
}