using Dapper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Globalization;

namespace ApiDesposte.Controllers.Sql
{
    [ApiController]
    [Route("api/[controller]")]
    [Route("api/ProduccionCabezasControlles")]
    public class ProdCabezasController : ControllerBase
    {
        private readonly string _connectionString;

        // Cache en memoria para evitar re-ejecutar el SP de larga duración en cada consulta
        private static readonly Dictionary<string, (DateTime Timestamp, List<ProdCabezaItemDto> Items)> _cachePorRango = new(StringComparer.OrdinalIgnoreCase);
        private static readonly object _cacheLock = new();
        private static readonly TimeSpan CacheDuracion = TimeSpan.FromMinutes(15);

        public ProdCabezasController(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("SqlDesposte")
                ?? throw new InvalidOperationException("No se encontró la cadena de conexión 'SqlDesposte' en appsettings.json");
        }

        private static (string f1, string f2) NormalizarFechas(string? fechaInicio, string? fechaFin)
        {
            DateTime hoy = DateTime.Today;
            DateTime inicioDefault = hoy.AddDays(-7);

            string f1 = FormatearFechaIndividual(fechaInicio, inicioDefault, esFechaFinal: false);
            string f2 = FormatearFechaIndividual(fechaFin, hoy, esFechaFinal: true);

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

            if (limpio.Contains('T') && limpio.Length >= 19)
            {
                return limpio;
            }

            string[] formatos = { "yyyy-MM-dd", "dd/MM/yyyy", "dd-MM-yyyy", "yyyy/MM/dd", "M/d/yyyy", "d/M/yyyy" };
            if (DateTime.TryParseExact(limpio, formatos, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime fechaExacta) ||
                DateTime.TryParse(limpio, new CultureInfo("es-PE"), DateTimeStyles.None, out fechaExacta) ||
                DateTime.TryParse(limpio, CultureInfo.InvariantCulture, DateTimeStyles.None, out fechaExacta))
            {
                return $"{fechaExacta:yyyy-MM-dd}{sufijoHora}";
            }

            if (!limpio.Contains('T') && !limpio.Contains(' '))
            {
                return $"{limpio}{sufijoHora}";
            }

            return limpio;
        }

        private async Task<List<ProdCabezaItemDto>> ObtenerDatosSpAsync(string f1, string f2, bool forzarRecarga = false)
        {
            string cacheKey = $"{f1}_{f2}";

            lock (_cacheLock)
            {
                if (!forzarRecarga && _cachePorRango.TryGetValue(cacheKey, out var cacheEntry))
                {
                    if ((DateTime.Now - cacheEntry.Timestamp) < CacheDuracion)
                    {
                        return cacheEntry.Items;
                    }
                }
            }

            using IDbConnection db = new SqlConnection(_connectionString);

            string sql = "EXEC dbo.Sp_Informe_Prduccion_Cabezas @Fecha1, @Fecha2";
            var datos = (await db.QueryAsync<ProdCabezaItemDto>(
                sql,
                new { Fecha1 = f1, Fecha2 = f2 },
                commandTimeout: 2800
            )).ToList();

            lock (_cacheLock)
            {
                _cachePorRango[cacheKey] = (DateTime.Now, datos);
            }

            return datos;
        }

        /// <summary>
        /// Endpoint principal: Obtiene el informe ejecutando:
        /// exec dbo.Sp_Informe_Prduccion_Cabezas @Fecha1, @Fecha2
        /// </summary>
        [HttpGet]
        public async Task<IActionResult> ObtenerInformeProdCabezas(
            [FromQuery] string? fechaInicio = null,
            [FromQuery] string? fechaFin = null,
            [FromQuery] string? fecha1 = null,
            [FromQuery] string? fecha2 = null,
            [FromQuery] bool recargar = false)
        {
            try
            {
                string? rawF1 = !string.IsNullOrWhiteSpace(fechaInicio) ? fechaInicio : fecha1;
                string? rawF2 = !string.IsNullOrWhiteSpace(fechaFin) ? fechaFin : fecha2;

                var (f1, f2) = NormalizarFechas(rawF1, rawF2);

                var lista = await ObtenerDatosSpAsync(f1, f2, recargar);

                return Ok(new
                {
                    exito = true,
                    parametros = new { Fecha1 = f1, Fecha2 = f2 },
                    totalRegistros = lista.Count,
                    datos = lista
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message, stackTrace = ex.ToString() });
            }
        }
    }

    public class ProdCabezaItemDto
    {
        public DateTime? FechaOrden { get; set; }
        public int Ttra { get; set; }
        public string NTtra { get; set; } = string.Empty;
        public string Producto { get; set; } = string.Empty;
        public string CodigoCorto { get; set; } = string.Empty;
        public string Nombre { get; set; } = string.Empty;
        public string TipoDesposte { get; set; } = string.Empty;
        public decimal Unidades { get; set; }
        public decimal Kilos { get; set; }
        public string Nivel { get; set; } = string.Empty;
    }
}