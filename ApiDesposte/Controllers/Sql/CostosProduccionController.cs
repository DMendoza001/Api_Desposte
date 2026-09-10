using Dapper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Data;

namespace ApiDesposte.Controllers.Sql;

[ApiController]
[Route("api/[controller]")]
public class CostosProduccionController : ControllerBase
{
    private readonly string _connectionString;

    public CostosProduccionController(IConfiguration configuration)
    {
        // Obtiene la cadena de conexión configurada en el appsettings.json
        _connectionString = configuration.GetConnectionString("SqlDesposte")!;
    }

    // Endpoint de prueba de conexión
    [HttpGet("probar-conexion")]
    public async Task<IActionResult> ProbarConexion()
    {
        try
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string query = "SELECT COUNT(1) AS TotalRegistros FROM Desposte.dbo.CostoProduccion";
            var total = await db.QueryFirstOrDefaultAsync<int>(query);

            return Ok(new 
            { 
                exito = true, 
                mensaje = "Conexión a base de datos Desposte exitosa", 
                totalRegistrosCostoProduccion = total 
            });
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { exito = false, error = ex.Message });
        }
    }

    // Endpoint principal: Obtener todos los costos de producción con búsqueda y filtros opcionales
    [HttpGet("listar")]
    [HttpGet]
    public async Task<IActionResult> ObtenerCostosProduccion(
        [FromQuery] int? anio = null,
        [FromQuery] int? mes = null,
        [FromQuery] string? codigo = null,
        [FromQuery] string? codigoCorto = null,
        [FromQuery] string? nombre = null)
    {
        try
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            string sql = @"
                SELECT 
                    YEAR(a.Fecha) AS [Año],
                    YEAR(a.Fecha) AS [Anio],
                    MONTH(a.Fecha) AS [Mes],
                    a.Codigo,
                    b.CodigoCorto,
                    b.Nombre,
                    a.Costo
                FROM Desposte.dbo.CostoProduccion a
                LEFT JOIN Desposte.dbo.Articulo_Codigo_Corto() b ON a.Codigo = b.CodigoArticulo
                WHERE 1 = 1";

            var parametros = new DynamicParameters();

            if (anio.HasValue)
            {
                sql += " AND YEAR(a.Fecha) = @Anio";
                parametros.Add("Anio", anio.Value);
            }

            if (mes.HasValue)
            {
                sql += " AND MONTH(a.Fecha) = @Mes";
                parametros.Add("Mes", mes.Value);
            }

            if (!string.IsNullOrWhiteSpace(codigo))
            {
                sql += " AND a.Codigo = @Codigo";
                parametros.Add("Codigo", codigo.Trim());
            }

            if (!string.IsNullOrWhiteSpace(codigoCorto))
            {
                sql += " AND b.CodigoCorto = @CodigoCorto";
                parametros.Add("CodigoCorto", codigoCorto.Trim());
            }

            if (!string.IsNullOrWhiteSpace(nombre))
            {
                sql += " AND b.Nombre LIKE @Nombre";
                parametros.Add("Nombre", $"%{nombre.Trim()}%");
            }

            sql += " ORDER BY YEAR(a.Fecha) DESC, MONTH(a.Fecha) DESC, b.CodigoCorto ASC, a.Codigo ASC;";

            var resultados = await db.QueryAsync(sql, parametros, commandTimeout: 1200);

            return Ok(resultados);
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { exito = false, error = ex.Message });
        }
    }

    // Endpoint para obtener resumen comparativo pivotado (opcional para consumo rápido)
    [HttpGet("periodos-disponibles")]
    public async Task<IActionResult> ObtenerPeriodosDisponibles()
    {
        try
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            string sql = @"
                SELECT DISTINCT 
                    YEAR(Fecha) AS [Año],
                    MONTH(Fecha) AS [Mes],
                    CAST(YEAR(Fecha) AS VARCHAR(4)) + RIGHT('0' + CAST(MONTH(Fecha) AS VARCHAR(2)), 2) AS [PeriodoKey]
                FROM Desposte.dbo.CostoProduccion
                ORDER BY [Año] DESC, [Mes] DESC;";

            var periodos = await db.QueryAsync(sql);
            return Ok(periodos);
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { exito = false, error = ex.Message });
        }
    }
}


