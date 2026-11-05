using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Concurrent;

namespace ApiDesposte.Controllers.Excel
{
    /// <summary>
    /// Controlador para exponer en formato JSON las 3 tablas del archivo
    /// Cumplimiento_PTC___Jorge.xlsm (Últimas 9 Semanas_PL):
    ///   - Hoja ResProd_Py_DB -> Tabla T_ResProd_Py_DB  (Semana, CodigoCorto, Kilos, Escenario)
    ///   - Hoja DemandaPy     -> Tabla T_DemandaPy      (Semana, CodigoCorto, Kilos, Escenario)
    ///   - Hoja Auxiliares    -> Tabla T_CodigoRelacion (CodigoCorto, Nombre, SubPlanta, TipoMP)
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    [Route("api/Cumplimientos_Jorge")]
    [Route("api/cumplimientos-jorge")]
    public class CumplimientosJorgeController : ControllerBase
    {
        private readonly ILogger<CumplimientosJorgeController> _logger;

        // Caché en memoria para lectura instantánea
        private static readonly ConcurrentDictionary<string, List<Dictionary<string, object?>>> _cacheTablas = new(StringComparer.OrdinalIgnoreCase);
        private static DateTime _ultimaModificacionArchivo = DateTime.MinValue;
        private static readonly object _lockCarga = new();

        // Mapeo oficial de las 3 tablas según su Hoja y Nombre de Tabla en Excel
        private static readonly Dictionary<string, (string Hoja, string Tabla)> MapeoTablas = new(StringComparer.OrdinalIgnoreCase)
        {
            { "T_ResProd_Py_DB",  ("ResProd_Py_DB", "T_ResProd_Py_DB") },
            { "T_DemandaPy",      ("DemandaPy",     "T_DemandaPy") },
            { "T_CodigoRelacion", ("Auxiliares",    "T_CodigoRelacion") }
        };

        public CumplimientosJorgeController(ILogger<CumplimientosJorgeController> logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// Obtiene la ruta dinámica hacia el archivo Excel sincronizado en OneDrive para cualquier usuario.
        /// </summary>
        private string ObtenerRutaExcel()
        {
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return Path.Combine(
                userProfile,
                @"OneDrive - Corporación Rico SAC\Desposte-03\PlantasCore\Planificacion\Pedido\Ultimas 9 Semanas_PL\Cumplimiento_PTC___Jorge.xlsm"
            );
        }

        /// <summary>
        /// GET: api/CumplimientosJorge
        /// Devuelve un objeto JSON con las 3 tablas completas listas para consumir.
        /// Pasar ?recargar=true para forzar relectura del Excel.
        /// </summary>
        [HttpGet]
        public IActionResult ObtenerTodasLasTablas([FromQuery] bool recargar = false)
        {
            try
            {
                var tablas = AsegurarCargaTablas(recargar);

                var resumen = tablas.ToDictionary(
                    k => k.Key,
                    v => v.Value.Count
                );

                return Ok(new
                {
                    exito = true,
                    archivo = Path.GetFileName(ObtenerRutaExcel()),
                    ultimaActualizacionArchivo = _ultimaModificacionArchivo,
                    totalFilasPorTabla = resumen,
                    tablas = tablas
                });
            }
            catch (FileNotFoundException ex)
            {
                _logger.LogWarning(ex, "Archivo Excel no encontrado.");
                return NotFound(new { exito = false, mensaje = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al leer las tablas de Cumplimiento_PTC___Jorge.xlsm");
                return StatusCode(500, new
                {
                    exito = false,
                    mensaje = "Ocurrió un error al procesar el archivo Excel.",
                    detalle = ex.Message
                });
            }
        }

        /// <summary>
        /// GET: api/CumplimientosJorge/{nombreTabla}
        /// Devuelve únicamente la tabla solicitada (ej: T_ResProd_Py_DB, T_DemandaPy, T_CodigoRelacion).
        /// </summary>
        [HttpGet("{nombreTabla}")]
        public IActionResult ObtenerTablaIndividual(string nombreTabla, [FromQuery] bool recargar = false)
        {
            try
            {
                if (!MapeoTablas.ContainsKey(nombreTabla))
                {
                    return BadRequest(new
                    {
                        exito = false,
                        mensaje = $"La tabla '{nombreTabla}' no está registrada.",
                        tablasDisponibles = MapeoTablas.Keys
                    });
                }

                var tablas = AsegurarCargaTablas(recargar);

                if (tablas.TryGetValue(nombreTabla, out var filas))
                {
                    return Ok(filas);
                }

                return NotFound(new { exito = false, mensaje = $"No se encontró información para la tabla '{nombreTabla}'." });
            }
            catch (FileNotFoundException ex)
            {
                _logger.LogWarning(ex, "Archivo Excel no encontrado.");
                return NotFound(new { exito = false, mensaje = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener la tabla {NombreTabla}", nombreTabla);
                return StatusCode(500, new
                {
                    exito = false,
                    mensaje = $"Error al leer la tabla '{nombreTabla}'.",
                    detalle = ex.Message
                });
            }
        }

        /// <summary>
        /// Carga las 3 tablas desde el Excel en memoria (solo si el archivo cambió o si se fuerza recarga).
        /// </summary>
        private ConcurrentDictionary<string, List<Dictionary<string, object?>>> AsegurarCargaTablas(bool forzarRecarga)
        {
            string rutaExcel = ObtenerRutaExcel();

            if (!System.IO.File.Exists(rutaExcel))
            {
                throw new FileNotFoundException($"No se encontró el archivo Excel en la ruta esperada: {rutaExcel}");
            }

            DateTime fechaModificacionActual = System.IO.File.GetLastWriteTimeUtc(rutaExcel);

            if (!forzarRecarga && !_cacheTablas.IsEmpty && fechaModificacionActual == _ultimaModificacionArchivo)
            {
                return _cacheTablas;
            }

            lock (_lockCarga)
            {
                if (!forzarRecarga && !_cacheTablas.IsEmpty && fechaModificacionActual == _ultimaModificacionArchivo)
                {
                    return _cacheTablas;
                }

                _logger.LogInformation("Cargando tablas desde Excel: {Ruta}", rutaExcel);

                using var memoryStream = new MemoryStream();
                using (var fileStream = new FileStream(rutaExcel, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                {
                    fileStream.CopyTo(memoryStream);
                }
                memoryStream.Position = 0;

                using var workbook = new XLWorkbook(memoryStream);

                foreach (var kvp in MapeoTablas)
                {
                    string claveTabla = kvp.Key;
                    string nombreHoja = kvp.Value.Hoja;
                    string nombreTablaExcel = kvp.Value.Tabla;

                    var datosTabla = ExtraerDatosTabla(workbook, nombreHoja, nombreTablaExcel);
                    _cacheTablas[claveTabla] = datosTabla;
                }

                _ultimaModificacionArchivo = fechaModificacionActual;
                return _cacheTablas;
            }
        }

        /// <summary>
        /// Extrae todas las filas de un ListObject (Tabla de Excel) o, en su defecto, del rango usado de la hoja.
        /// </summary>
        private List<Dictionary<string, object?>> ExtraerDatosTabla(XLWorkbook workbook, string nombreHoja, string nombreTabla)
        {
            var resultado = new List<Dictionary<string, object?>>();

            // 1. Intentar buscar la tabla directamente en el libro o en la hoja indicada
            IXLTable? tabla = null;

            if (workbook.Worksheets.TryGetWorksheet(nombreHoja, out var hoja))
            {
                tabla = hoja.Tables.FirstOrDefault(t => t.Name.Equals(nombreTabla, StringComparison.OrdinalIgnoreCase));
            }

            if (tabla == null)
            {
                tabla = workbook.Worksheets
                    .SelectMany(ws => ws.Tables)
                    .FirstOrDefault(t => t.Name.Equals(nombreTabla, StringComparison.OrdinalIgnoreCase));
            }

            if (tabla != null)
            {
                var encabezados = tabla.Fields.Select(f => f.Name?.Trim() ?? string.Empty).ToList();
                var dataRange = tabla.DataRange;

                if (dataRange == null)
                    return resultado;

                foreach (var row in dataRange.Rows())
                {
                    if (row.IsEmpty())
                        continue;

                    var filaDict = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                    for (int i = 0; i < encabezados.Count; i++)
                    {
                        string nombreColumna = string.IsNullOrEmpty(encabezados[i]) ? $"Columna{i + 1}" : encabezados[i];
                        var cell = row.Cell(i + 1);
                        filaDict[nombreColumna] = ObtenerValorCelda(cell);
                    }
                    resultado.Add(filaDict);
                }

                return resultado;
            }

            // 2. Fallback: si no está definida como objeto Tabla pero existe la hoja, leer el rango utilizado
            if (hoja != null)
            {
                var rangeUsed = hoja.RangeUsed();
                if (rangeUsed == null)
                    return resultado;

                var rows = rangeUsed.RowsUsed().ToList();
                if (rows.Count <= 1)
                    return resultado;

                var headerRow = rows[0];
                int totalCols = rangeUsed.ColumnCount();
                var headers = new List<string>();

                for (int c = 1; c <= totalCols; c++)
                {
                    string h = headerRow.Cell(c).GetString().Trim();
                    headers.Add(string.IsNullOrEmpty(h) ? $"Columna{c}" : h);
                }

                for (int r = 1; r < rows.Count; r++)
                {
                    var currentRow = rows[r];
                    var filaDict = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);

                    for (int c = 1; c <= totalCols; c++)
                    {
                        filaDict[headers[c - 1]] = ObtenerValorCelda(currentRow.Cell(c));
                    }

                    resultado.Add(filaDict);
                }
            }

            return resultado;
        }

        /// <summary>
        /// Convierte el valor de una celda de ClosedXML a un tipo primitivo limpio para serialización JSON.
        /// </summary>
        private static object? ObtenerValorCelda(IXLCell cell)
        {
            if (cell == null || cell.IsEmpty())
                return null;

            if (cell.Value.IsError)
                return null;

            if (cell.Value.IsNumber)
            {
                double num = cell.Value.GetNumber();
                if (Math.Abs(num % 1) < 0.0000001 && num >= long.MinValue && num <= long.MaxValue)
                {
                    return (long)Math.Round(num);
                }
                return Math.Round(num, 4);
            }

            if (cell.Value.IsDateTime)
            {
                return cell.Value.GetDateTime().ToString("yyyy-MM-dd");
            }

            if (cell.Value.IsBoolean)
            {
                return cell.Value.GetBoolean();
            }

            string texto = cell.GetString().Trim();
            return string.IsNullOrEmpty(texto) ? null : texto;
        }
    }
}
