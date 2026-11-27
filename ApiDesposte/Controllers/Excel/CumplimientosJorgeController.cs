using ApiDesposte.Funciones.Excel;
using Microsoft.AspNetCore.Mvc;
using System.IO;

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
    public class CumplimientosJorgeController : ControllerBase
    {
        private readonly ILogger<CumplimientosJorgeController> _logger;

        // Mapeo oficial de las 3 tablas según su Hoja y Nombre de Tabla en Excel
        private static readonly Dictionary<string, (string Hoja, string Tabla)> MapeoTablas = new(StringComparer.OrdinalIgnoreCase)
        {
            { "T_ResProd_Py_DB",  ("ResProd_Py_DB", "T_ResProd_Py_DB") },
            { "T_DemandaPy",      ("DemandaPy",     "T_DemandaPy") },
            { "T_CodigoRelacion", ("Auxiliares",    "T_CodigoRelacion") }
        };

        // Cache en memoria para lectura instantánea
        private static readonly CacheTablasExcel Cache = new(TimeSpan.FromSeconds(60));

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
                string rutaExcel = ObtenerRutaExcel();
                var tablas = Cache.Obtener(rutaExcel, MapeoTablas, recargar);

                var resumen = tablas.ToDictionary(
                    k => k.Key,
                    v => v.Value.Count
                );

                DateTime ultimaModificacion = System.IO.File.Exists(rutaExcel)
                    ? System.IO.File.GetLastWriteTimeUtc(rutaExcel)
                    : DateTime.MinValue;

                return Ok(new
                {
                    exito = true,
                    archivo = Path.GetFileName(rutaExcel),
                    ultimaActualizacionArchivo = ultimaModificacion,
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

                string rutaExcel = ObtenerRutaExcel();
                var tablas = Cache.Obtener(rutaExcel, MapeoTablas, recargar);

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
    }
}
