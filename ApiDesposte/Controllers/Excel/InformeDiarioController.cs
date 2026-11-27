using ApiDesposte.Funciones.Excel;
using Microsoft.AspNetCore.Mvc;
using System.IO;

namespace ApiDesposte.Controllers.Excel
{
    [ApiController]
    [Route("api/[controller]")]
    public class InformeDiarioController : ControllerBase
    {
        // Mapeo de las 6 Hojas y Tablas del Excel
        private static readonly Dictionary<string, (string Hoja, string Tabla)> MapeoTablas = new(StringComparer.OrdinalIgnoreCase)
        {
            { "personal", ("BD_PERSONAL", "TablaPersonal") },
            { "sso", ("BD_SSO", "TablaSSO") },
            { "horarios", ("BD_HORARIOS", "TablaHorarios") },
            { "produccion", ("BD_PRODUCCION", "TablaProduccion") },
            { "detalle", ("BD_DETALLE", "TablaDetalle") },
            { "paradas", ("BD_PARADAS", "TablaParadas") }
        };

        // Cache en memoria para optimizar peticiones seguidas desde el dashboard HTML
        private static readonly CacheTablasExcel Cache = new(TimeSpan.FromSeconds(60));

        private string ObtenerRutaExcel()
        {
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return Path.Combine(userProfile, @"OneDrive - Corporación Rico SAC\Desposte-04\Informes Diario\Informe Diario 2026 - Base de Datos.xlsx");
        }

        /// <summary>
        /// ENDPOINT ÚNICO: Retorna las 6 tablas en un solo JSON para el dashboard HTML.
        /// GET /api/InformeDiario
        /// Opcional: ?forzarRecarga=true para saltar la caché en memoria.
        /// </summary>
        [HttpGet]
        public IActionResult ObtenerInformeDiario([FromQuery] bool forzarRecarga = false)
        {
            try
            {
                string rutaExcel = ObtenerRutaExcel();
                var tablas = Cache.Obtener(rutaExcel, MapeoTablas, forzarRecarga);

                return Ok(new
                {
                    exito = true,
                    tablas
                });
            }
            catch (FileNotFoundException ex)
            {
                return NotFound(new { exito = false, error = ex.Message });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = $"Error al leer el archivo Excel: {ex.Message}" });
            }
        }
    }
}
