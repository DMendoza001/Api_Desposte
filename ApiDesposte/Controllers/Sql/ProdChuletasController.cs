using Dapper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Data;

namespace ApiDesposte.Controllers.Sql
{
    [ApiController]
    [Route("api/[controller]")]
    public class ProdChuletasController : ControllerBase
    {
        private readonly string _connectionString;

        // Cache en memoria para evitar re-ejecutar el SP de larga duración en cada filtro o consulta
        private static readonly Dictionary<string, (DateTime Timestamp, List<ProdChuletaItemDto> Items)> _cachePorRango = new(StringComparer.OrdinalIgnoreCase);
        private static readonly object _cacheLock = new();
        private static readonly TimeSpan CacheDuracion = TimeSpan.FromMinutes(15);

        public ProdChuletasController(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("SqlDesposte")
                ?? throw new InvalidOperationException("No se encontró la cadena de conexión 'SqlDesposte' en appsettings.json");
        }

        private static (string f1, string f2) NormalizarFechas(string? fechaInicio, string? fechaFin)
        {
            string f1 = string.IsNullOrWhiteSpace(fechaInicio) ? "2026-01-01T00:00:00" : fechaInicio.Trim();
            if (!f1.Contains('T') && !f1.Contains(' '))
            {
                f1 = $"{f1}T00:00:00";
            }

            string f2 = string.IsNullOrWhiteSpace(fechaFin) ? "2026-08-31T23:59:59" : fechaFin.Trim();
            if (!f2.Contains('T') && !f2.Contains(' '))
            {
                f2 = $"{f2}T23:59:59";
            }

            return (f1, f2);
        }

        /// <summary>
        /// Ejecuta el procedimiento almacenado Desposte.dbo.Sp_Informe_Prod_Chuletas @Fecha1, @Fecha2
        /// utilizando almacenamiento en caché en memoria para alto rendimiento.
        /// </summary>
        private async Task<List<ProdChuletaItemDto>> ObtenerDatosSpAsync(string f1, string f2, bool forzarRecarga = false)
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

            // Timeout de 1800 segundos (30 min) para soportar rangos extensos de varios meses
            string sql = "EXEC Desposte.dbo.Sp_Informe_Prod_Chuletas @Fecha1, @Fecha2";
            var datos = (await db.QueryAsync<ProdChuletaItemDto>(
                sql,
                new { Fecha1 = f1, Fecha2 = f2 },
                commandTimeout: 1800
            )).ToList();

            lock (_cacheLock)
            {
                _cachePorRango[cacheKey] = (DateTime.Now, datos);
            }

            return datos;
        }

        /// <summary>
        /// Endpoint principal: Obtiene el informe de producción de chuletas ejecutando:
        /// exec Desposte.dbo.Sp_Informe_Prod_Chuletas @Fecha1, @Fecha2
        /// Por defecto usa: '2026-01-01T00:00:00' hasta '2026-08-31T23:59:59'
        /// </summary>
        [HttpGet]
        [HttpGet("listar")]
        [HttpGet("informe")]
        public async Task<IActionResult> ObtenerInformeProdChuletas(
            [FromQuery] string? fechaInicio = null,
            [FromQuery] string? fechaFin = null,
            [FromQuery] string? fecha1 = null,
            [FromQuery] string? fecha2 = null,
            [FromQuery] string? tipo = null,
            [FromQuery] string? tipoChuleta = null,
            [FromQuery] string? nTtra = null,
            [FromQuery] string? tipoInforme = null,
            [FromQuery] string? codigoCorto = null,
            [FromQuery] string? producto = null,
            [FromQuery] string? nombre = null,
            [FromQuery] string? search = null,
            [FromQuery] bool recargar = false)
        {
            try
            {
                // Soporta alias fechaInicio/fechaFin o fecha1/fecha2
                string? rawF1 = !string.IsNullOrWhiteSpace(fechaInicio) ? fechaInicio : fecha1;
                string? rawF2 = !string.IsNullOrWhiteSpace(fechaFin) ? fechaFin : fecha2;

                var (f1, f2) = NormalizarFechas(rawF1, rawF2);

                var lista = await ObtenerDatosSpAsync(f1, f2, recargar);

                // Filtros opcionales en memoria
                if (!string.IsNullOrWhiteSpace(tipo))
                {
                    lista = lista.Where(x => string.Equals(x.Tipo, tipo.Trim(), StringComparison.OrdinalIgnoreCase)).ToList();
                }

                if (!string.IsNullOrWhiteSpace(tipoChuleta))
                {
                    lista = lista.Where(x => string.Equals(x.TipoChuleta, tipoChuleta.Trim(), StringComparison.OrdinalIgnoreCase)).ToList();
                }

                if (!string.IsNullOrWhiteSpace(nTtra))
                {
                    lista = lista.Where(x => string.Equals(x.NTtra, nTtra.Trim(), StringComparison.OrdinalIgnoreCase)).ToList();
                }

                if (!string.IsNullOrWhiteSpace(tipoInforme))
                {
                    lista = lista.Where(x => string.Equals(x.TipoInforme, tipoInforme.Trim(), StringComparison.OrdinalIgnoreCase)).ToList();
                }

                if (!string.IsNullOrWhiteSpace(codigoCorto))
                {
                    lista = lista.Where(x => string.Equals(x.CodigoCorto, codigoCorto.Trim(), StringComparison.OrdinalIgnoreCase)).ToList();
                }

                if (!string.IsNullOrWhiteSpace(producto))
                {
                    lista = lista.Where(x => string.Equals(x.Producto, producto.Trim(), StringComparison.OrdinalIgnoreCase)).ToList();
                }

                if (!string.IsNullOrWhiteSpace(nombre))
                {
                    lista = lista.Where(x => x.Nombre.Contains(nombre.Trim(), StringComparison.OrdinalIgnoreCase)).ToList();
                }

                if (!string.IsNullOrWhiteSpace(search))
                {
                    string q = search.Trim();
                    lista = lista.Where(x =>
                        x.CodigoCorto.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                        x.Producto.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                        x.Nombre.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                        x.NPlantilla.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                        x.TipoChuleta.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                        x.TipoInforme.Contains(q, StringComparison.OrdinalIgnoreCase)
                    ).ToList();
                }

                return Ok(new
                {
                    exito = true,
                    parametros = new { Fecha1 = f1, Fecha2 = f2 },
                    totalRegistros = lista.Count,
                    totalKilos = Math.Round(lista.Sum(x => x.Kilos), 2),
                    totalUnidades = Math.Round(lista.Sum(x => x.Unidades), 2),
                    totalCosto = Math.Round(lista.Sum(x => x.CostoTotal), 2),
                    totalVenta = Math.Round(lista.Sum(x => x.VentaTotal), 2),
                    datos = lista
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message, stackTrace = ex.ToString() });
            }
        }
    }

    /// <summary>
    /// DTO que representa un registro devuelto por Desposte.dbo.Sp_Informe_Prod_Chuletas
    /// </summary>
    public class ProdChuletaItemDto
    {
        public DateTime? FechaOrden { get; set; }
        public string? FechaOrdenStr => FechaOrden?.ToString("yyyy-MM-dd");
        public string Tipo { get; set; } = string.Empty;
        public string TipoChuleta { get; set; } = string.Empty;
        public int? Plantilla { get; set; }
        public string NPlantilla { get; set; } = string.Empty;
        public int? Ttra { get; set; }
        public string NTtra { get; set; } = string.Empty;
        public string Producto { get; set; } = string.Empty;
        public string CodigoCorto { get; set; } = string.Empty;
        public string Nombre { get; set; } = string.Empty;
        public string TipoInforme { get; set; } = string.Empty;
        public decimal Unidades { get; set; }
        public decimal Kilos { get; set; }
        public decimal CostoTotal { get; set; }
        public decimal VentaTotal { get; set; }
        public decimal MargenBruto => Math.Round(VentaTotal - CostoTotal, 2);
        public decimal PctMargen => VentaTotal > 0 ? Math.Round((VentaTotal - CostoTotal) / VentaTotal * 100m, 2) : 0m;
        public decimal CostoPorKilo => Kilos > 0 ? Math.Round(CostoTotal / Kilos, 3) : 0m;
        public decimal VentaPorKilo => Kilos > 0 ? Math.Round(VentaTotal / Kilos, 3) : 0m;
    }
}
