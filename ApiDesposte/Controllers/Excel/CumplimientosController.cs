using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System.IO;

namespace ApiDesposte.Controllers.Excel
{
    [ApiController]
    [Route("api/[controller]")]
    public class CumplimientosController : ControllerBase
    {
        private const string HojaDatosPY = "DatosPY";
        private const string TablaProyectado = "T_Proyectado";

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
                        // 1. Eliminar PivotTableParts de todas las hojas para evitar conflictos con ClosedXML
                        foreach (var wsPart in doc.WorkbookPart.WorksheetParts)
                        {
                            var pParts = wsPart.PivotTableParts.ToList();
                            foreach (var p in pParts)
                            {
                                wsPart.DeletePart(p);
                            }
                        }

                        // 2. Eliminar PivotTableCacheDefinitionParts del Workbook
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
            if (double.TryParse(txt, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double val))
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
        /// Obtiene los datos de la tabla T_Proyectado de la hoja DatosPY del archivo Cumplimiento_PTC.xlsm
        /// </summary>
        [HttpGet("datos-proyectado")]
        public IActionResult ObtenerDatosProyectado(
            [FromQuery] string? semana = null,
            [FromQuery] string? tipoPy = null,
            [FromQuery] string? origenDestino = null,
            [FromQuery] string? tipo = null)
        {
            try
            {
                string rutaExcel = ObtenerRutaExcel();
                if (!System.IO.File.Exists(rutaExcel))
                {
                    return NotFound(new { exito = false, mensaje = $"No se encontró el archivo Excel en la ruta: {rutaExcel}" });
                }

                using var workbook = CargarWorkbookEnMemoria(rutaExcel);

                if (!workbook.Worksheets.TryGetWorksheet(HojaDatosPY, out var wsDatosPY))
                {
                    return NotFound(new { exito = false, mensaje = $"No se encontró la hoja '{HojaDatosPY}' en el archivo." });
                }

                IXLTable? tablaProyectado = null;
                try
                {
                    tablaProyectado = wsDatosPY.Table(TablaProyectado);
                }
                catch
                {
                    tablaProyectado = wsDatosPY.Tables.FirstOrDefault(t => string.Equals(t.Name, TablaProyectado, StringComparison.OrdinalIgnoreCase));
                }

                if (tablaProyectado == null)
                {
                    return NotFound(new { exito = false, mensaje = $"No se encontró la tabla '{TablaProyectado}' en la hoja '{HojaDatosPY}'." });
                }

                // Identificar índices de columnas
                int colSemana = BuscarIndiceColumna(tablaProyectado, "Semana");
                int colTipoPy = BuscarIndiceColumna(tablaProyectado, "TipoPy");
                int colOrigenDestino = BuscarIndiceColumna(tablaProyectado, "OrigenDestino");
                int colTipo = BuscarIndiceColumna(tablaProyectado, "Tipo");
                int colUnidades = BuscarIndiceColumna(tablaProyectado, "Unidades");
                int colProm = BuscarIndiceColumna(tablaProyectado, "Prom");
                int colKilos = BuscarIndiceColumna(tablaProyectado, "Kilos");

                var listaProyectado = new List<CumplimientoProyectadoDto>();

                foreach (var fila in tablaProyectado.DataRange.Rows())
                {
                    if (fila.IsEmpty()) continue;

                    string valSemana = ObtenerValorCeldaTexto(fila, colSemana);
                    string valTipoPy = ObtenerValorCeldaTexto(fila, colTipoPy);
                    string valOrigenDestino = ObtenerValorCeldaTexto(fila, colOrigenDestino);
                    string valTipo = ObtenerValorCeldaTexto(fila, colTipo);
                    double valUnidades = ObtenerValorCeldaNumero(fila, colUnidades);
                    double valProm = ObtenerValorCeldaNumero(fila, colProm);
                    double valKilos = ObtenerValorCeldaNumero(fila, colKilos);

                    // Si toda la fila está vacía, omitir
                    if (string.IsNullOrWhiteSpace(valSemana) && 
                        string.IsNullOrWhiteSpace(valTipoPy) && 
                        string.IsNullOrWhiteSpace(valOrigenDestino) &&
                        valUnidades == 0 && valKilos == 0)
                    {
                        continue;
                    }

                    // Filtros opcionales
                    if (!string.IsNullOrWhiteSpace(semana) && !string.Equals(valSemana, semana.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                    if (!string.IsNullOrWhiteSpace(tipoPy) && !string.Equals(valTipoPy, tipoPy.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                    if (!string.IsNullOrWhiteSpace(origenDestino) && !string.Equals(valOrigenDestino, origenDestino.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                    if (!string.IsNullOrWhiteSpace(tipo) && !string.Equals(valTipo, tipo.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    int anio = 0;
                    int nSemana = 0;
                    if (!string.IsNullOrWhiteSpace(valSemana) && valSemana.Length >= 5)
                    {
                        int.TryParse(valSemana.Substring(0, 4), out anio);
                        int.TryParse(valSemana.Substring(4), out nSemana);
                    }

                    listaProyectado.Add(new CumplimientoProyectadoDto
                    {
                        Semana = valSemana,
                        Anio = anio,
                        NSemana = nSemana,
                        TipoPy = valTipoPy,
                        OrigenDestino = valOrigenDestino,
                        Tipo = valTipo,
                        Unidades = Math.Round(valUnidades, 2),
                        Promedio = Math.Round(valProm, 4),
                        Kilos = Math.Round(valKilos, 2)
                    });
                }

                return Ok(listaProyectado);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message, stackTrace = ex.ToString() });
            }
        }

        /// <summary>
        /// Obtiene los datos dinámicos como diccionario clave-valor según los encabezados originales de la tabla
        /// </summary>
        [HttpGet("raw")]
        public IActionResult ObtenerDatosRaw()
        {
            try
            {
                string rutaExcel = ObtenerRutaExcel();
                if (!System.IO.File.Exists(rutaExcel))
                {
                    return NotFound(new { exito = false, mensaje = $"No se encontró el archivo Excel en la ruta: {rutaExcel}" });
                }

                using var workbook = CargarWorkbookEnMemoria(rutaExcel);
                if (!workbook.Worksheets.TryGetWorksheet(HojaDatosPY, out var wsDatosPY))
                {
                    return NotFound(new { exito = false, mensaje = $"No se encontró la hoja '{HojaDatosPY}'" });
                }

                IXLTable? tabla = null;
                try
                {
                    tabla = wsDatosPY.Table(TablaProyectado);
                }
                catch
                {
                    tabla = wsDatosPY.Tables.FirstOrDefault(t => string.Equals(t.Name, TablaProyectado, StringComparison.OrdinalIgnoreCase));
                }

                if (tabla == null)
                {
                    return NotFound(new { exito = false, mensaje = $"No se encontró la tabla '{TablaProyectado}' en la hoja '{HojaDatosPY}'." });
                }

                var encabezados = tabla.HeadersRow().Cells()
                    .Select(c => c.Value.ToString().Trim())
                    .ToList();

                var resultado = new List<Dictionary<string, object>>();

                foreach (var fila in tabla.DataRange.Rows())
                {
                    if (fila.IsEmpty()) continue;

                    var filaDiccionario = new Dictionary<string, object>();
                    for (int i = 0; i < encabezados.Count; i++)
                    {
                        var celda = fila.Cell(i + 1);
                        if (celda.IsEmpty())
                        {
                            filaDiccionario[encabezados[i]] = string.Empty;
                        }
                        else if (celda.DataType == XLDataType.Number)
                        {
                            filaDiccionario[encabezados[i]] = celda.GetValue<double>();
                        }
                        else if (celda.DataType == XLDataType.DateTime)
                        {
                            filaDiccionario[encabezados[i]] = celda.GetDateTime().ToString("yyyy-MM-dd");
                        }
                        else
                        {
                            filaDiccionario[encabezados[i]] = celda.Value.ToString().Trim();
                        }
                    }
                    resultado.Add(filaDiccionario);
                }

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }
    }

    public class CumplimientoProyectadoDto
    {
        public string Semana { get; set; } = string.Empty;
        public int Anio { get; set; }
        public int NSemana { get; set; }
        public string TipoPy { get; set; } = string.Empty;
        public string OrigenDestino { get; set; } = string.Empty;
        public string Tipo { get; set; } = string.Empty;
        public double Unidades { get; set; }
        public double Promedio { get; set; }
        public double Kilos { get; set; }
    }
}
