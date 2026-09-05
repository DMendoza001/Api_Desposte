using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System.IO;

namespace ApiDesposte.Controllers.Excel
{
    [ApiController]
    [Route("api/[controller]")]
    public class ComparativoPresupuestoController : ControllerBase
    {
        private const string HojaCabecera = "Cabecera";
        private const string TablaCabecera = "T_Cabecera";

        private const string HojaForeCast = "ForeCast";
        private const string TablaForeCast = "T_ForeCast";

        private string ObtenerRutaExcel()
        {
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return Path.Combine(userProfile, @"OneDrive - Corporación Rico SAC\Desposte-03\Presupuestos\2027\Presupuesto_2027.xlsm");
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
                        // 1. Eliminar PivotTableParts de todas las hojas para evitar errores de ClosedXML
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
                // Si falla la manipulación de OpenXml, intentamos cargar directamente
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

        private static double ObtenerValorCeldaNumero(IXLRangeRow fila, int colIndex)
        {
            if (colIndex <= 0) return 0.0;
            var celda = fila.Cell(colIndex);
            if (celda.IsEmpty()) return 0.0;

            if (celda.DataType == XLDataType.Number)
            {
                return celda.GetValue<double>();
            }

            if (double.TryParse(celda.Value.ToString().Trim(), out double val))
            {
                return val;
            }

            return 0.0;
        }

        // Endpoint principal: Retorna la unión completa agrupada (Full Outer Join)
        [HttpGet("datos-comparativo")]
        public IActionResult ObtenerDatosComparativo()
        {
            try
            {
                string rutaExcel = ObtenerRutaExcel();
                if (!System.IO.File.Exists(rutaExcel))
                {
                    return NotFound(new { error = $"No se encontró el archivo Excel en la ruta: {rutaExcel}" });
                }

                using var workbook = CargarWorkbookEnMemoria(rutaExcel);

                if (!workbook.Worksheets.TryGetWorksheet(HojaCabecera, out var wsCabecera))
                    return NotFound(new { error = $"No se encontró la hoja '{HojaCabecera}'" });

                if (!workbook.Worksheets.TryGetWorksheet(HojaForeCast, out var wsForeCast))
                    return NotFound(new { error = $"No se encontró la hoja '{HojaForeCast}'" });

                var tablaCabecera = wsCabecera.Table(TablaCabecera);
                var tablaForeCast = wsForeCast.Table(TablaForeCast);

                // Índices de columnas en T_Cabecera
                int cCecoCab = BuscarIndiceColumna(tablaCabecera, "CentroCostos", "Centro Costos", "CECO");
                int cMesCab = BuscarIndiceColumna(tablaCabecera, "Mes", "MES");
                int cGrupoCab = BuscarIndiceColumna(tablaCabecera, "Ngrupo", "GrupoNombre", "Grupo");
                int cClaseCab = BuscarIndiceColumna(tablaCabecera, "Nclase", "ClaseNombre", "Clase");
                int cPartidaCab = BuscarIndiceColumna(tablaCabecera, "PartidaPresupuestalNombre", "PartidaPresupuestal", "Partida");
                int cMontoCab = BuscarIndiceColumna(tablaCabecera, "Presupuesto", "Monto", "Importe");

                // Índices de columnas en T_ForeCast
                int cCecoFc = BuscarIndiceColumna(tablaForeCast, "CECO", "CentroCostos", "Centro Costos");
                int cMesFc = BuscarIndiceColumna(tablaForeCast, "Mes", "MES");
                int cGrupoFc = BuscarIndiceColumna(tablaForeCast, "GrupoNombre", "Ngrupo", "Grupo");
                int cClaseFc = BuscarIndiceColumna(tablaForeCast, "ClaseNombre", "Nclase", "Clase");
                int cPartidaFc = BuscarIndiceColumna(tablaForeCast, "PartidaPresupuestalNombre", "PartidaPresupuestal", "Partida");
                int cMontoFc = BuscarIndiceColumna(tablaForeCast, "Monto", "Presupuesto", "ForeCast", "Importe");

                // Diccionario para unir y consolidar por clave compuesta
                var consolidadoMap = new Dictionary<string, ComparativoItemDto>(StringComparer.OrdinalIgnoreCase);

                // 1. Procesar T_Cabecera (Presupuesto)
                foreach (var fila in tablaCabecera.DataRange.Rows())
                {
                    string ceco = ObtenerValorCeldaTexto(fila, cCecoCab, "SIN CECO");
                    string mes = ObtenerValorCeldaTexto(fila, cMesCab, "SIN MES");
                    string grupo = ObtenerValorCeldaTexto(fila, cGrupoCab, "SIN GRUPO");
                    string clase = ObtenerValorCeldaTexto(fila, cClaseCab, "SIN CLASE");
                    string partida = ObtenerValorCeldaTexto(fila, cPartidaCab, "SIN PARTIDA");
                    double presupuesto = ObtenerValorCeldaNumero(fila, cMontoCab);

                    string clave = $"{ceco}|{mes}|{grupo}|{clase}|{partida}";

                    if (!consolidadoMap.TryGetValue(clave, out var item))
                    {
                        item = new ComparativoItemDto
                        {
                            CentroCostos = ceco,
                            Mes = mes,
                            Grupo = grupo,
                            Clase = clase,
                            PartidaPresupuestal = partida
                        };
                        consolidadoMap[clave] = item;
                    }

                    item.Presupuesto += presupuesto;
                }

                // 2. Procesar T_ForeCast (ForeCast / Proyección)
                foreach (var fila in tablaForeCast.DataRange.Rows())
                {
                    string ceco = ObtenerValorCeldaTexto(fila, cCecoFc, "SIN CECO");
                    string mes = ObtenerValorCeldaTexto(fila, cMesFc, "SIN MES");
                    string grupo = ObtenerValorCeldaTexto(fila, cGrupoFc, "SIN GRUPO");
                    string clase = ObtenerValorCeldaTexto(fila, cClaseFc, "SIN CLASE");
                    string partida = ObtenerValorCeldaTexto(fila, cPartidaFc, "SIN PARTIDA");
                    double forecastMonto = ObtenerValorCeldaNumero(fila, cMontoFc);

                    string clave = $"{ceco}|{mes}|{grupo}|{clase}|{partida}";

                    if (!consolidadoMap.TryGetValue(clave, out var item))
                    {
                        item = new ComparativoItemDto
                        {
                            CentroCostos = ceco,
                            Mes = mes,
                            Grupo = grupo,
                            Clase = clase,
                            PartidaPresupuestal = partida
                        };
                        consolidadoMap[clave] = item;
                    }

                    item.ForeCast += forecastMonto;
                }

                // Redondear totales y devolver lista ordenada
                var resultado = consolidadoMap.Values
                    .Select(i => new
                    {
                        centroCostos = i.CentroCostos,
                        mes = i.Mes,
                        grupo = i.Grupo,
                        clase = i.Clase,
                        partidaPresupuestal = i.PartidaPresupuestal,
                        presupuesto = Math.Round(i.Presupuesto, 2),
                        forecast = Math.Round(i.ForeCast, 2),
                        variacion = Math.Round(i.ForeCast - i.Presupuesto, 2),
                        porcentajeCumplimiento = i.Presupuesto > 0 
                            ? Math.Round((i.ForeCast / i.Presupuesto) * 100, 2) 
                            : (i.ForeCast > 0 ? 100.0 : 0.0)
                    })
                    .OrderBy(x => x.centroCostos)
                    .ThenBy(x => x.mes)
                    .ThenBy(x => x.grupo)
                    .ThenBy(x => x.clase)
                    .ToList();

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message, stackTrace = ex.ToString() });
            }
        }

    }

    public class ComparativoItemDto
    {
        public string CentroCostos { get; set; } = string.Empty;
        public string Mes { get; set; } = string.Empty;
        public string Grupo { get; set; } = string.Empty;
        public string Clase { get; set; } = string.Empty;
        public string PartidaPresupuestal { get; set; } = string.Empty;
        public double Presupuesto { get; set; }
        public double ForeCast { get; set; }
    }
}
