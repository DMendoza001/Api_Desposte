using Dapper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Globalization;

namespace ApiDesposte.Controllers.Sql
{
    /// <summary>
    /// Controlador para ejecutar los procedimientos almacenados del Kardex de Unidades de Cerdos:
    ///   EXEC Desposte.dbo.SP_Kardex_UnCerdos_PTC @Fecha1, @Fecha2
    ///   EXEC Desposte.dbo.SP_Kardex_UnCerdos_PBC @Fecha1, @Fecha2
    /// Cada endpoint devuelve únicamente las columnas y filas originales del SP,
    /// sin agregar columnas calculadas ni transformar los datos (data plana).
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]          // regla 6: una sola ruta canónica por controlador
    public class KardexUnCerdosController : ControllerBase
    {
        private const string SpPtc = "Desposte.dbo.SP_Kardex_UnCerdos_PTC";
        private const string SpPbc = "Desposte.dbo.SP_Kardex_UnCerdos_PBC";

        // Cache en memoria por (SP + rango de fechas) para evitar re-ejecutar el SP en cada consulta
        private static readonly Dictionary<string, (DateTime Timestamp, List<dynamic> Items)> _cachePorRango = new(StringComparer.OrdinalIgnoreCase);
        private static readonly object _cacheLock = new();
        private static readonly TimeSpan CacheDuracion = TimeSpan.FromMinutes(15);

        private readonly string _connectionString;

        public KardexUnCerdosController(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("SqlDesposte")
                ?? throw new InvalidOperationException("No se encontró la cadena de conexión 'SqlDesposte' en appsettings.json");
        }

        /// <summary>
        /// Normaliza las fechas exactamente con el formato requerido por los SP:
        ///   Fecha1 = "yyyy-MM-ddT00:00:00"
        ///   Fecha2 = "yyyy-MM-ddT23:59:59"
        /// Por defecto: del 01 de enero del año actual hasta hoy.
        /// </summary>
        private static (string f1, string f2) FormatearFechasSp(string? fecha1Raw, string? fecha2Raw)
        {
            DateTime hoy = DateTime.Today;
            DateTime inicioDefault = new DateTime(hoy.Year, 1, 1);

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
        /// Ejecuta el SP indicado con los dos parámetros de fecha y devuelve las filas tal cual
        /// (mismas columnas y tipos que entrega SQL Server).
        /// </summary>
        private async Task<IActionResult> EjecutarSpAsync(string nombreSp, string? fecha1Raw, string? fecha2Raw, bool forzarRecarga)
        {
            try
            {
                var (f1, f2) = FormatearFechasSp(fecha1Raw, fecha2Raw);
                string cacheKey = $"{nombreSp}_{f1}_{f2}";

                lock (_cacheLock)
                {
                    if (!forzarRecarga && _cachePorRango.TryGetValue(cacheKey, out var cacheEntry))
                    {
                        if ((DateTime.Now - cacheEntry.Timestamp) < CacheDuracion)
                        {
                            return Ok(cacheEntry.Items);
                        }
                    }
                }

                using IDbConnection db = new SqlConnection(_connectionString);

                string sql = $"EXEC {nombreSp} @Fecha1, @Fecha2";
                var resultados = (await db.QueryAsync(
                    sql,
                    new { Fecha1 = f1, Fecha2 = f2 },
                    commandTimeout: 2800
                )).ToList();

                lock (_cacheLock)
                {
                    _cachePorRango[cacheKey] = (DateTime.Now, resultados);
                }

                return Ok(resultados);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new
                {
                    exito = false,
                    mensaje = $"Error al ejecutar {nombreSp}",
                    detalle = ex.Message
                });
            }
        }

        /// <summary>
        /// GET: api/KardexUnCerdos/PTC?fechaInicio=2026-01-01&amp;fechaFin=2026-11-22
        /// Ejecuta: EXEC Desposte.dbo.SP_Kardex_UnCerdos_PTC @Fecha1, @Fecha2
        /// Columnas: Semana, TipoCerdo, SaldoInicial, IngresoPBC, IngresoDev, Producción,
        ///           SalidaDespacho, Ajuste, SaldoFinal
        /// </summary>
        [HttpGet("PTC")]                  // regla 6: un solo atributo de verbo por acción
        public Task<IActionResult> ObtenerKardexPtc(
            [FromQuery] string? fechaInicio = null,
            [FromQuery] string? fechaFin = null,
            [FromQuery] string? fecha1 = null,
            [FromQuery] string? fecha2 = null,
            [FromQuery] bool recargar = false)
        {
            string? rawF1 = !string.IsNullOrWhiteSpace(fechaInicio) ? fechaInicio : fecha1;
            string? rawF2 = !string.IsNullOrWhiteSpace(fechaFin) ? fechaFin : fecha2;

            return EjecutarSpAsync(SpPtc, rawF1, rawF2, recargar);
        }

        /// <summary>
        /// GET: api/KardexUnCerdos/PBC?fechaInicio=2026-01-01&amp;fechaFin=2026-11-22
        /// Ejecuta: EXEC Desposte.dbo.SP_Kardex_UnCerdos_PBC @Fecha1, @Fecha2
        /// Columnas: Semana, TipoCerdo, SaldoInicial, Beneficiado, IngresoDev, Carcasa,
        ///           EnviadoPTC, Harinas, Ajuste, SaldoFinal
        /// </summary>
        [HttpGet("PBC")]
        public Task<IActionResult> ObtenerKardexPbc(
            [FromQuery] string? fechaInicio = null,
            [FromQuery] string? fechaFin = null,
            [FromQuery] string? fecha1 = null,
            [FromQuery] string? fecha2 = null,
            [FromQuery] bool recargar = false)
        {
            string? rawF1 = !string.IsNullOrWhiteSpace(fechaInicio) ? fechaInicio : fecha1;
            string? rawF2 = !string.IsNullOrWhiteSpace(fechaFin) ? fechaFin : fecha2;

            return EjecutarSpAsync(SpPbc, rawF1, rawF2, recargar);
        }
    }
}
