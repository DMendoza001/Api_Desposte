using Dapper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Globalization;

namespace ApiDesposte.Controllers.Sql
{
    /// <summary>
    /// Controlador para ejecutar el procedimiento almacenado:
    ///   EXEC Desposte.dbo.Sp_Informe_Prod_Cerdos @Fecha1, @Fecha2
    /// Devuelve únicamente las columnas y filas originales del SP sin agregar columnas adicionales.
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    [Route("api/InformeProdCerdos")]
    [Route("api/Sp_Informe_Prod_Cerdos")]
    public class ProdCerdosController : ControllerBase
    {
        private readonly string _connectionString;

        public ProdCerdosController(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("SqlDesposte")
                ?? throw new InvalidOperationException("No se encontró la cadena de conexión 'SqlDesposte' en appsettings.json");
        }

        /// <summary>
        /// Normaliza las fechas exactamente con el formato requerido por el SP:
        ///   Fecha1 = "yyyy-MM-ddT00:00:00"
        ///   Fecha2 = "yyyy-MM-ddT23:59:59"
        /// </summary>
        private static (string f1, string f2) FormatearFechasSp(string? fecha1Raw, string? fecha2Raw)
        {
            DateTime hoy = DateTime.Today;
            DateTime inicioDefault = new DateTime(hoy.Year, hoy.Month, 1);

            string f1 = FormatearFechaIndividual(fecha1Raw, inicioDefault, esFechaFinal: false);
            string f2 = FormatearFechaIndividual(fecha2Raw, hoy, esFechaFinal: true);

            return (f1, f2);
        }

        private static string FormatearFechaIndividual(string? input, DateTime valorPorDefecto, bool esFechaFinal)
        {
            string sufijoHora = esFechaFinal ? "T23:59:59" : "T00:00:00";

            if (string.IsNullOrWhiteSpace(input))
            {
                return $"{valorPorDefecto:yyyy-MM-dd}{sufijoHora}";
            }

            string limpio = input.Trim();

            // Si ya viene con formato completo yyyy-MM-ddTHH:mm:ss, respetarlo
            if (limpio.Contains('T') && limpio.Length >= 19)
            {
                return limpio;
            }

            // Intentar parsear formatos comunes (yyyy-MM-dd, dd/MM/yyyy, etc.)
            string[] formatos = { "yyyy-MM-dd", "dd/MM/yyyy", "dd-MM-yyyy", "yyyy/MM/dd", "M/d/yyyy", "d/M/yyyy" };
            if (DateTime.TryParseExact(limpio, formatos, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime fechaExacta) ||
                DateTime.TryParse(limpio, new CultureInfo("es-PE"), DateTimeStyles.None, out fechaExacta) ||
                DateTime.TryParse(limpio, CultureInfo.InvariantCulture, DateTimeStyles.None, out fechaExacta))
            {
                return $"{fechaExacta:yyyy-MM-dd}{sufijoHora}";
            }

            return $"{limpio}{sufijoHora}";
        }

        /// <summary>
        /// GET: api/ProdCerdos?fecha1=2026-10-01&fecha2=2026-10-31
        /// (También acepta ?fechaInicio=...&fechaFin=...)
        /// Ejecuta: EXEC Desposte.dbo.Sp_Informe_Prod_Cerdos @Fecha1, @Fecha2
        /// y retorna únicamente los datos del SP sin columnas adicionales.
        /// </summary>
        [HttpGet]
        [HttpGet("listar")]
        public async Task<IActionResult> ObtenerInformeProdCerdos(
            [FromQuery] string? fecha1 = null,
            [FromQuery] string? fecha2 = null,
            [FromQuery] string? fechaInicio = null,
            [FromQuery] string? fechaFin = null)
        {
            try
            {
                string? rawF1 = !string.IsNullOrWhiteSpace(fecha1) ? fecha1 : fechaInicio;
                string? rawF2 = !string.IsNullOrWhiteSpace(fecha2) ? fecha2 : fechaFin;

                var (f1, f2) = FormatearFechasSp(rawF1, rawF2);

                using IDbConnection db = new SqlConnection(_connectionString);

                const string sql = "EXEC Desposte.dbo.Sp_Informe_Prod_Cerdos @Fecha1, @Fecha2";

                var resultados = await db.QueryAsync(
                    sql,
                    new { Fecha1 = f1, Fecha2 = f2 },
                    commandTimeout: 2800
                );

                return Ok(resultados);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new
                {
                    exito = false,
                    mensaje = "Error al ejecutar Desposte.dbo.Sp_Informe_Prod_Cerdos",
                    detalle = ex.Message
                });
            }
        }
    }
}

