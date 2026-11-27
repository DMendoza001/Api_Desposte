using ApiDesposte.Funciones.Fechas;
using ApiDesposte.Funciones.Sql;
using Microsoft.AspNetCore.Mvc;

namespace ApiDesposte.Controllers.Sql
{
    /// <summary>
    /// Controlador para ejecutar el procedimiento almacenado:
    ///   EXEC Desposte.dbo.Sp_Informe_Prod_Cerdos @Fecha1, @Fecha2
    /// Devuelve únicamente las columnas y filas originales del SP sin agregar columnas adicionales.
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    public class ProdCerdosController : ControllerBase
    {
        private const string Sp = "Desposte.dbo.Sp_Informe_Prod_Cerdos";
        private readonly string _connectionString;

        public ProdCerdosController(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("SqlDesposte")
                ?? throw new InvalidOperationException("No se encontró la cadena de conexión 'SqlDesposte' (variable de entorno ConnectionStrings__SqlDesposte)");
        }

        /// <summary>
        /// GET: api/ProdCerdos?fecha1=2026-10-01&amp;fecha2=2026-10-31
        /// (También acepta ?fechaInicio=...&amp;fechaFin=...)
        /// Por defecto: del primer día del mes actual hasta hoy.
        /// </summary>
        [HttpGet]
        public async Task<IActionResult> ObtenerInformeProdCerdos(
            [FromQuery] string? fecha1 = null,
            [FromQuery] string? fecha2 = null,
            [FromQuery] string? fechaInicio = null,
            [FromQuery] string? fechaFin = null)
        {
            try
            {
                DateTime hoy = DateTime.Today;
                string f1 = ParametrosFecha.FormatearParaSp(ParametrosFecha.Primero(fecha1, fechaInicio), new DateTime(hoy.Year, hoy.Month, 1), esFechaFinal: false);
                string f2 = ParametrosFecha.FormatearParaSp(ParametrosFecha.Primero(fecha2, fechaFin), hoy, esFechaFinal: true);

                var datos = await EjecutorSql.EjecutarTalCualAsync(
                    _connectionString,
                    $"EXEC {Sp} @Fecha1, @Fecha2",
                    new { Fecha1 = f1, Fecha2 = f2 },
                    commandTimeout: 2800);

                return Ok(datos);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new
                {
                    exito = false,
                    mensaje = $"Error al ejecutar {Sp}",
                    detalle = ex.Message
                });
            }
        }
    }
}
