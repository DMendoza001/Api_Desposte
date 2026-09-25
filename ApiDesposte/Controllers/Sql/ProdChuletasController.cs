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
        /// Prueba de conectividad con la base de datos SQL Server y verificación de la tabla y SP.
        /// </summary>
        [HttpGet("probar-conexion")]
        public async Task<IActionResult> ProbarConexion()
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string query = @"
                    SELECT 
                        @@SERVERNAME AS Servidor,
                        DB_NAME() AS BaseDatos,
                        OBJECT_ID('Desposte.dbo.Sp_Informe_Prod_Chuletas') AS IdSp";

                var info = await db.QueryFirstOrDefaultAsync(query);

                return Ok(new
                {
                    exito = true,
                    mensaje = "Conexión a SQL Server exitosa",
                    info
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
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

        /// <summary>
        /// Resumen ejecutivo consolidado:
        /// Totales generales, desglose por NTtra (ORIGEN vs RESULTADO), desglose por TipoChuleta,
        /// desglose por TipoInforme, y rendimiento/margen.
        /// </summary>
        [HttpGet("resumen")]
        public async Task<IActionResult> ObtenerResumenEjecutivo(
            [FromQuery] string? fechaInicio = null,
            [FromQuery] string? fechaFin = null,
            [FromQuery] bool recargar = false)
        {
            try
            {
                var (f1, f2) = NormalizarFechas(fechaInicio, fechaFin);
                var lista = await ObtenerDatosSpAsync(f1, f2, recargar);

                decimal totalKilos = Math.Round(lista.Sum(x => x.Kilos), 2);
                decimal totalUnidades = Math.Round(lista.Sum(x => x.Unidades), 2);
                decimal totalCosto = Math.Round(lista.Sum(x => x.CostoTotal), 2);
                decimal totalVenta = Math.Round(lista.Sum(x => x.VentaTotal), 2);
                decimal margenBruto = Math.Round(totalVenta - totalCosto, 2);
                decimal pctMargen = totalVenta > 0 ? Math.Round((margenBruto / totalVenta) * 100m, 2) : 0m;

                // Desglose por NTtra (ORIGEN vs RESULTADO)
                var porTtra = lista
                    .GroupBy(x => string.IsNullOrWhiteSpace(x.NTtra) ? "OTROS" : x.NTtra.ToUpper())
                    .Select(g => new
                    {
                        NTtra = g.Key,
                        Unidades = Math.Round(g.Sum(x => x.Unidades), 2),
                        Kilos = Math.Round(g.Sum(x => x.Kilos), 2),
                        CostoTotal = Math.Round(g.Sum(x => x.CostoTotal), 2),
                        VentaTotal = Math.Round(g.Sum(x => x.VentaTotal), 2),
                        PctKilos = totalKilos > 0 ? Math.Round((g.Sum(x => x.Kilos) / totalKilos) * 100m, 2) : 0m,
                        TotalRegistros = g.Count()
                    })
                    .OrderByDescending(x => x.Kilos)
                    .ToList();

                // Desglose por TipoChuleta (CONGELADO, FRESCO, etc.)
                var porTipoChuleta = lista
                    .GroupBy(x => string.IsNullOrWhiteSpace(x.TipoChuleta) ? "SIN TIPO" : x.TipoChuleta.ToUpper())
                    .Select(g => new
                    {
                        TipoChuleta = g.Key,
                        Unidades = Math.Round(g.Sum(x => x.Unidades), 2),
                        Kilos = Math.Round(g.Sum(x => x.Kilos), 2),
                        CostoTotal = Math.Round(g.Sum(x => x.CostoTotal), 2),
                        VentaTotal = Math.Round(g.Sum(x => x.VentaTotal), 2),
                        PctKilos = totalKilos > 0 ? Math.Round((g.Sum(x => x.Kilos) / totalKilos) * 100m, 2) : 0m,
                        TotalRegistros = g.Count()
                    })
                    .OrderByDescending(x => x.Kilos)
                    .ToList();

                // Desglose por TipoInforme
                var porTipoInforme = lista
                    .GroupBy(x => string.IsNullOrWhiteSpace(x.TipoInforme) ? "OTROS" : x.TipoInforme.ToUpper())
                    .Select(g => new
                    {
                        TipoInforme = g.Key,
                        Unidades = Math.Round(g.Sum(x => x.Unidades), 2),
                        Kilos = Math.Round(g.Sum(x => x.Kilos), 2),
                        CostoTotal = Math.Round(g.Sum(x => x.CostoTotal), 2),
                        VentaTotal = Math.Round(g.Sum(x => x.VentaTotal), 2),
                        TotalRegistros = g.Count()
                    })
                    .OrderByDescending(x => x.Kilos)
                    .ToList();

                return Ok(new
                {
                    exito = true,
                    parametros = new { Fecha1 = f1, Fecha2 = f2 },
                    totales = new
                    {
                        TotalKilos = totalKilos,
                        TotalUnidades = totalUnidades,
                        TotalCosto = totalCosto,
                        TotalVenta = totalVenta,
                        MargenBruto = margenBruto,
                        PctMargen = pctMargen,
                        TotalRegistros = lista.Count,
                        TotalSkus = lista.Select(x => x.CodigoCorto).Distinct().Count()
                    },
                    desglosePorTtra = porTtra,
                    desglosePorTipoChuleta = porTipoChuleta,
                    desglosePorTipoInforme = porTipoInforme
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }

        /// <summary>
        /// Agrupación por Producto (Código Corto / Artículo):
        /// Muestra el volumen en Kilos, Unidades, CostoTotal, VentaTotal y Margen por cada SKU.
        /// </summary>
        [HttpGet("por-producto")]
        public async Task<IActionResult> ObtenerPorProducto(
            [FromQuery] string? fechaInicio = null,
            [FromQuery] string? fechaFin = null,
            [FromQuery] string? nTtra = null,
            [FromQuery] bool recargar = false)
        {
            try
            {
                var (f1, f2) = NormalizarFechas(fechaInicio, fechaFin);
                var lista = await ObtenerDatosSpAsync(f1, f2, recargar);

                if (!string.IsNullOrWhiteSpace(nTtra))
                {
                    lista = lista.Where(x => string.Equals(x.NTtra, nTtra.Trim(), StringComparison.OrdinalIgnoreCase)).ToList();
                }

                var agrupado = lista
                    .GroupBy(x => new { x.CodigoCorto, x.Producto, x.Nombre, x.TipoChuleta, x.NTtra })
                    .Select(g =>
                    {
                        decimal k = Math.Round(g.Sum(x => x.Kilos), 2);
                        decimal u = Math.Round(g.Sum(x => x.Unidades), 2);
                        decimal c = Math.Round(g.Sum(x => x.CostoTotal), 2);
                        decimal v = Math.Round(g.Sum(x => x.VentaTotal), 2);
                        decimal margen = Math.Round(v - c, 2);
                        decimal pctM = v > 0 ? Math.Round((margen / v) * 100m, 2) : 0m;

                        return new
                        {
                            CodigoCorto = g.Key.CodigoCorto,
                            Producto = g.Key.Producto,
                            Nombre = g.Key.Nombre,
                            TipoChuleta = g.Key.TipoChuleta,
                            NTtra = g.Key.NTtra,
                            Unidades = u,
                            Kilos = k,
                            CostoTotal = c,
                            VentaTotal = v,
                            MargenBruto = margen,
                            PctMargen = pctM,
                            CostoPorKilo = k > 0 ? Math.Round(c / k, 3) : 0m,
                            VentaPorKilo = k > 0 ? Math.Round(v / k, 3) : 0m
                        };
                    })
                    .OrderByDescending(x => x.Kilos)
                    .ToList();

                return Ok(new
                {
                    exito = true,
                    parametros = new { Fecha1 = f1, Fecha2 = f2 },
                    totalProductos = agrupado.Count,
                    datos = agrupado
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }

        /// <summary>
        /// Agrupación por Plantilla de producción:
        /// </summary>
        [HttpGet("por-plantilla")]
        public async Task<IActionResult> ObtenerPorPlantilla(
            [FromQuery] string? fechaInicio = null,
            [FromQuery] string? fechaFin = null,
            [FromQuery] bool recargar = false)
        {
            try
            {
                var (f1, f2) = NormalizarFechas(fechaInicio, fechaFin);
                var lista = await ObtenerDatosSpAsync(f1, f2, recargar);

                var agrupado = lista
                    .GroupBy(x => new { x.Plantilla, x.NPlantilla, x.TipoChuleta })
                    .Select(g =>
                    {
                        decimal k = Math.Round(g.Sum(x => x.Kilos), 2);
                        decimal u = Math.Round(g.Sum(x => x.Unidades), 2);
                        decimal c = Math.Round(g.Sum(x => x.CostoTotal), 2);
                        decimal v = Math.Round(g.Sum(x => x.VentaTotal), 2);

                        return new
                        {
                            Plantilla = g.Key.Plantilla,
                            NPlantilla = g.Key.NPlantilla,
                            TipoChuleta = g.Key.TipoChuleta,
                            Unidades = u,
                            Kilos = k,
                            CostoTotal = c,
                            VentaTotal = v,
                            MargenBruto = Math.Round(v - c, 2),
                            TotalSkus = g.Select(x => x.CodigoCorto).Distinct().Count(),
                            TotalRegistros = g.Count()
                        };
                    })
                    .OrderByDescending(x => x.Kilos)
                    .ToList();

                return Ok(new
                {
                    exito = true,
                    parametros = new { Fecha1 = f1, Fecha2 = f2 },
                    totalPlantillas = agrupado.Count,
                    datos = agrupado
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
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
