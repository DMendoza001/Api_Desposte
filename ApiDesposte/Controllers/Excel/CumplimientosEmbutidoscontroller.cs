using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System.IO;

namespace ApiDesposte.Controllers.Excel
{
    [ApiController]
    [Route("api/[controller]")]
    public class CumplimientosEmbutidosController : ControllerBase
    {
        private const string HojaCumplimientosEmb = "CumplimientosEmb";
        private const string TablaCumplimientos = "T_Cumplientos";

        // Cache en memoria para optimizar lecturas consecutivas desde dashboards JS
        private static List<CumplimientoEmbutidoDto>? _cacheItems = null;
        private static DateTime _cacheTimestamp = DateTime.MinValue;
        private static readonly object _cacheLock = new();
        private static readonly TimeSpan CacheDuracion = TimeSpan.FromSeconds(45);

        private string ObtenerRutaExcel()
        {
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string rutaOneDrive = Path.Combine(userProfile, @"OneDrive - Corporación Rico SAC\Desposte-03\PlantasCore\Planificacion\Cumplimiento_PTC.xlsm");
            if (System.IO.File.Exists(rutaOneDrive))
            {
                return rutaOneDrive;
            }

            // Fallback dinámico buscando directorios OneDrive en el perfil del usuario actual
            try
            {
                var dirBase = new DirectoryInfo(userProfile);
                var oneDriveDirs = dirBase.GetDirectories("OneDrive*");
                foreach (var oneDriveDir in oneDriveDirs)
                {
                    string rutaAlt = Path.Combine(oneDriveDir.FullName, @"Desposte-03\PlantasCore\Planificacion\Cumplimiento_PTC.xlsm");
                    if (System.IO.File.Exists(rutaAlt))
                    {
                        return rutaAlt;
                    }
                }
            }
            catch
            {
                // Ignorar error al explorar directorios
            }

            return rutaOneDrive;
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

        private static IXLTable? ObtenerTabla(IXLWorksheet ws, string nombreTabla)
        {
            try
            {
                return ws.Table(nombreTabla);
            }
            catch
            {
                return ws.Tables.FirstOrDefault(t => string.Equals(t.Name, nombreTabla, StringComparison.OrdinalIgnoreCase))
                    ?? ws.Tables.FirstOrDefault(t => t.Name.Contains("Cumplien", StringComparison.OrdinalIgnoreCase))
                    ?? ws.Tables.FirstOrDefault();
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

            try
            {
                string fmt = celda.GetFormattedString();
                if (!string.IsNullOrWhiteSpace(fmt)) return fmt.Trim();
            }
            catch
            {
                // Fallback si GetFormattedString() falla
            }

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

        private List<CumplimientoEmbutidoDto> CargarDatosTablaCumplimientos(bool forzarRecarga = false)
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

                if (!workbook.Worksheets.TryGetWorksheet(HojaCumplimientosEmb, out var ws))
                {
                    throw new InvalidOperationException($"No se encontró la hoja '{HojaCumplimientosEmb}' en el libro.");
                }

                var tabla = ObtenerTabla(ws, TablaCumplimientos);
                if (tabla == null)
                {
                    throw new InvalidOperationException($"No se encontró la tabla '{TablaCumplimientos}' en la hoja '{HojaCumplimientosEmb}'.");
                }

                // Identificar columnas de T_Cumplientos: Semana | Codigo | Descripcion | Pedidos | Atendido
                int colSemana = BuscarIndiceColumna(tabla, "Semana", "Sem", "NroSemana");
                int colCodigo = BuscarIndiceColumna(tabla, "Codigo", "Cod", "CodigoProducto", "CodigoArticulo");
                int colDescripcion = BuscarIndiceColumna(tabla, "Descripcion", "Nombre", "Producto", "Articulo", "Desc");
                int colPedidos = BuscarIndiceColumna(tabla, "Pedidos", "Pedido", "KilosPedido", "CantPedido");
                int colAtendido = BuscarIndiceColumna(tabla, "Atendido", "Atendidos", "KilosAtendido", "CantAtendido");

                var items = new List<CumplimientoEmbutidoDto>();

                foreach (var fila in tabla.DataRange.Rows())
                {
                    if (fila.IsEmpty()) continue;

                    string semana = ObtenerValorCeldaTexto(fila, colSemana);
                    string codigo = ObtenerValorCeldaTexto(fila, colCodigo);
                    string descripcion = ObtenerValorCeldaTexto(fila, colDescripcion);
                    double pedidos = ObtenerValorCeldaNumero(fila, colPedidos);
                    double atendido = ObtenerValorCeldaNumero(fila, colAtendido);

                    // Si toda la fila está vacía o sin datos relevantes, ignorar
                    if (string.IsNullOrWhiteSpace(semana) &&
                        string.IsNullOrWhiteSpace(codigo) &&
                        string.IsNullOrWhiteSpace(descripcion) &&
                        pedidos == 0 && atendido == 0)
                    {
                        continue;
                    }

                    double diferencia = Math.Round(atendido - pedidos, 2);
                    double cumplimiento = pedidos > 0
                        ? Math.Round((atendido / pedidos) * 100.0, 2)
                        : (atendido > 0 ? 100.0 : 0.0);

                    items.Add(new CumplimientoEmbutidoDto
                    {
                        Semana = semana,
                        Codigo = codigo,
                        Descripcion = descripcion,
                        Pedidos = Math.Round(pedidos, 2),
                        Atendido = Math.Round(atendido, 2),
                        Diferencia = diferencia,
                        Cumplimiento = cumplimiento
                    });
                }

                _cacheItems = items;
                _cacheTimestamp = DateTime.Now;

                return _cacheItems;
            }
        }

        /// <summary>
        /// Obtiene los datos de la tabla T_Cumplientos de la hoja CumplimientosEmb del archivo Cumplimiento_PTC.xlsm tal cual está para ser consumido en JavaScript.
        /// </summary>
        /// <param name="semana">Filtro opcional por semana (ej. "262643")</param>
        /// <param name="codigo">Filtro opcional por código de producto (ej. "05813")</param>
        /// <param name="recargar">Fuerza la recarga del archivo Excel evitando el caché en memoria</param>
        /// <returns>Lista de registros de la tabla T_Cumplientos</returns>
        [HttpGet]
        [HttpGet("datos")]
        [HttpGet("listar")]
        public IActionResult ObtenerCumplimientosEmbutidos(
            [FromQuery] string? semana = null,
            [FromQuery] string? codigo = null,
            [FromQuery] bool recargar = false)
        {
            try
            {
                var datos = CargarDatosTablaCumplimientos(recargar);

                if (!string.IsNullOrWhiteSpace(semana))
                {
                    datos = datos.Where(x => string.Equals(x.Semana, semana.Trim(), StringComparison.OrdinalIgnoreCase)).ToList();
                }

                if (!string.IsNullOrWhiteSpace(codigo))
                {
                    datos = datos.Where(x => string.Equals(x.Codigo, codigo.Trim(), StringComparison.OrdinalIgnoreCase)).ToList();
                }

                return Ok(datos);
            }
            catch (FileNotFoundException fnfEx)
            {
                return NotFound(new { exito = false, mensaje = fnfEx.Message });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message, stackTrace = ex.ToString() });
            }
        }
    }

    /// <summary>
    /// DTO que representa un registro tal cual de la tabla T_Cumplientos en la hoja CumplimientosEmb
    /// </summary>
    public class CumplimientoEmbutidoDto
    {
        public string Semana { get; set; } = string.Empty;
        public string Codigo { get; set; } = string.Empty;
        public string Descripcion { get; set; } = string.Empty;
        public double Pedidos { get; set; }
        public double Atendido { get; set; }
        public double Diferencia { get; set; }
        public double Cumplimiento { get; set; }
    }
}
