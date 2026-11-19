using ApiDesposte.Infraestructura.Excel;
using Microsoft.AspNetCore.Mvc;
using System.IO;

namespace ApiDesposte.Controllers.Excel
{
    [ApiController]
    [Route("api/[controller]")]
    public class IndicadoresN3Controller : ControllerBase
    {
        private const string NombreHoja = "Registro";
        private const string NombreTabla = "tblRegistro";

        private static readonly Dictionary<string, (string Hoja, string Tabla)> MapeoTablas = new(StringComparer.OrdinalIgnoreCase)
        {
            { NombreTabla, (NombreHoja, NombreTabla) }
        };

        // Cache en memoria para optimizar lecturas consecutivas desde el frontend JS
        private static readonly CacheTablasExcel Cache = new(TimeSpan.FromSeconds(60));

        private string ObtenerRutaExcel()
        {
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return Path.Combine(userProfile, @"OneDrive - Corporación Rico SAC\Desposte-03\PlantasCore\Indicadores\BD Indicadores N3.xlsx");
        }

        /// <summary>
        /// Devuelve la tabla tblRegistro (hoja Registro) de BD Indicadores N3.xlsx tal cual está en el Excel.
        /// GET /api/IndicadoresN3?recargar=true
        /// </summary>
        [HttpGet]
        public IActionResult ObtenerRegistro([FromQuery] bool recargar = false)
        {
            try
            {
                string rutaExcel = ObtenerRutaExcel();
                var tablas = Cache.Obtener(rutaExcel, MapeoTablas, recargar);
                return Ok(tablas[NombreTabla]);
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
