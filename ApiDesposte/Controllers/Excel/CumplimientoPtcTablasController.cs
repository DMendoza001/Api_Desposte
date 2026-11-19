using ApiDesposte.Infraestructura.Excel;
using Microsoft.AspNetCore.Mvc;
using System.IO;

namespace ApiDesposte.Controllers.Excel
{
    [ApiController]
    [Route("api/[controller]")]
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

        // Cache en memoria para optimizar lecturas consecutivas desde el frontend JS
        private static readonly CacheTablasExcel Cache = new(TimeSpan.FromSeconds(60));

        private string ObtenerRutaExcel()
        {
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return Path.Combine(userProfile, @"OneDrive - Corporación Rico SAC\Desposte-03\PlantasCore\Planificacion\Cumplimiento_PTC.xlsm");
        }

        private Dictionary<string, List<Dictionary<string, object?>>> CargarTodasLasTablas(bool forzarRecarga = false) =>
            Cache.Obtener(ObtenerRutaExcel(), MapeoTablas, forzarRecarga);

        /// <summary>
        /// Devuelve las 5 tablas tal cual están en el Excel, sin agregar campos adicionales.
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
        /// Devuelve una tabla específica tal cual está en el Excel.
        /// Valores permitidos: T_Produccion | T_DemandaPy | T_ResProd_Py_DB | T_Distribucion | T_CodigoRelacion.
        /// GET /api/CumplimientoPtcTablas/{nombreTabla}
        /// </summary>
        [HttpGet("{nombreTabla}")]
        public IActionResult ObtenerTablaPorNombre(string nombreTabla, [FromQuery] bool recargar = false)
        {
            try
            {
                if (!MapeoTablas.TryGetValue(nombreTabla.Trim(), out var mapeo))
                {
                    return BadRequest(new
                    {
                        exito = false,
                        error = $"La tabla '{nombreTabla}' no está registrada. Valores permitidos: T_Produccion, T_DemandaPy, T_ResProd_Py_DB, T_Distribucion, T_CodigoRelacion."
                    });
                }

                var tablas = CargarTodasLasTablas(recargar);
                return Ok(tablas[mapeo.Tabla]);
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
