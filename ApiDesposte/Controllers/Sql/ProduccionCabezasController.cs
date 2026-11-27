using ApiDesposte.Funciones.Fechas;
using Dapper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Data;

namespace ApiDesposte.Controllers.Sql
{
    [ApiController]
    [Route("api/[controller]")]
    public class ProduccionCabezasController : ControllerBase
    {
        private readonly string _connectionString;

        // Cache en memoria para evitar re-ejecutar el SP de larga duración en cada consulta
        private static readonly Dictionary<string, (DateTime Timestamp, List<ProdCabezaItemDto> Items)> _cachePorRango = new(StringComparer.OrdinalIgnoreCase);
        private static readonly object _cacheLock = new();
        private static readonly TimeSpan CacheDuracion = TimeSpan.FromMinutes(15);

        public ProduccionCabezasController(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("SqlDesposte")
                ?? throw new InvalidOperationException("No se encontró la cadena de conexión 'SqlDesposte' (variable de entorno ConnectionStrings__SqlDesposte)");
        }

        private static (string f1, string f2) NormalizarFechas(string? fechaInicio, string? fechaFin)
        {
            DateTime hoy = DateTime.Today;
            return (
                ParametrosFecha.FormatearParaSp(fechaInicio, hoy.AddDays(-7), esFechaFinal: false),
                ParametrosFecha.FormatearParaSp(fechaFin, hoy, esFechaFinal: true));
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
                var (f1, f2) = NormalizarFechas(ParametrosFecha.Primero(fechaInicio, fecha1), ParametrosFecha.Primero(fechaFin, fecha2));

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