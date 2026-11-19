using ApiDesposte.Infraestructura.Excel;
using Microsoft.AspNetCore.Mvc;
using System.IO;

namespace ApiDesposte.Controllers.Excel
{
    [ApiController]
    [Route("api/[controller]")]
    public class PresupuestoController : ControllerBase
    {
        private const string NombreHoja = "Cabecera";
        private const string NombreTabla = "T_Cabecera";

        private static readonly Dictionary<string, (string Hoja, string Tabla)> MapeoTablas = new(StringComparer.OrdinalIgnoreCase)
        {
            { NombreTabla, (NombreHoja, NombreTabla) }
        };

        private static readonly CacheTablasExcel Cache = new(TimeSpan.FromSeconds(60));

        private string ObtenerRutaExcel()
        {
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return Path.Combine(userProfile, @"OneDrive - Corporación Rico SAC\Desposte-03\Presupuestos\2027\Presupuesto_2027.xlsm");
        }

        [HttpGet("dashboard-data")]
        public IActionResult ObtenerDatosDashboard([FromQuery] bool recargar = false)
        {
            try
            {
                string rutaExcel = ObtenerRutaExcel();
                if (!System.IO.File.Exists(rutaExcel))
                {
                    return NotFound(new { error = $"No se encontró el archivo en la ruta: {rutaExcel}" });
                }

                var tablas = Cache.Obtener(rutaExcel, MapeoTablas, recargar);
                return Ok(tablas[NombreTabla]);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message });
            }
        }
    }
}
