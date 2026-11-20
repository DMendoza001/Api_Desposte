using Dapper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Globalization;

namespace ApiDesposte.Controllers.Sql
{
    [ApiController]
    [Route("api/[controller]")]
    public class MarinadoController : ControllerBase
    {
        private readonly string _connectionString;

        public MarinadoController(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("SqlDesposte")
                ?? throw new InvalidOperationException("No se encontró la cadena de conexión 'SqlDesposte' en appsettings.json");
        }

        /// <summary>
        /// Informe de Marinado Congelado. Ejecuta:
        /// exec Desposte.dbo.Sp_Informe_Marinado_Congelado @Fecha1, @Fecha2
        /// Fecha1 = yyyy-MM-ddT00:00:00 | Fecha2 = yyyy-MM-ddT23:59:59
        /// Devuelve las filas tal cual las entrega SQL Server (mismas columnas,
        /// sin campos calculados).
        /// </summary>
        /// <param name="fechaInicio">Fecha inicial (yyyy-MM-dd). Por defecto, hace 7 días.</param>
        /// <param name="fechaFin">Fecha final (yyyy-MM-dd). Por defecto, hoy.</param>
        [HttpGet]
        public async Task<IActionResult> ObtenerInformeMarinado(
            [FromQuery] string? fechaInicio = null,
            [FromQuery] string? fechaFin = null)
        {
            DateTime hoy = DateTime.Today;

            if (!TryParsearFecha(fechaInicio, hoy.AddDays(-7), out DateTime inicio) ||
                !TryParsearFecha(fechaFin, hoy, out DateTime fin))
            {
                return BadRequest(new { exito = false, error = "Formato de fecha inválido. Use yyyy-MM-dd." });
            }

            if (inicio > fin)
            {
                return BadRequest(new { exito = false, error = "La fecha inicial no puede ser mayor que la fecha final." });
            }

            string fecha1 = $"{inicio:yyyy-MM-dd}T00:00:00";
            string fecha2 = $"{fin:yyyy-MM-dd}T23:59:59";

            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);

                var filas = await db.QueryAsync(
                    "EXEC Desposte.dbo.Sp_Informe_Marinado_Congelado @Fecha1, @Fecha2",
                    new { Fecha1 = fecha1, Fecha2 = fecha2 },
                    commandTimeout: 2800);

                // Cada fila como diccionario columna -> valor, tal cual viene del SP
                var datos = filas
                    .Select(f => new Dictionary<string, object?>((IDictionary<string, object?>)f))
                    .ToList();

                return Ok(new
                {
                    exito = true,
                    parametros = new { Fecha1 = fecha1, Fecha2 = fecha2 },
                    totalRegistros = datos.Count,
                    datos
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }

        private static bool TryParsearFecha(string? input, DateTime valorPorDefecto, out DateTime fecha)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                fecha = valorPorDefecto;
                return true;
            }

            string[] formatos = { "yyyy-MM-dd", "yyyy-MM-ddTHH:mm:ss", "dd/MM/yyyy" };
            return DateTime.TryParseExact(input.Trim(), formatos, CultureInfo.InvariantCulture, DateTimeStyles.None, out fecha);
        }
    }
}
