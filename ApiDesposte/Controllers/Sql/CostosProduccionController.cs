using ApiDesposte.Funciones.Sql;
using Dapper;
using Microsoft.AspNetCore.Mvc;

namespace ApiDesposte.Controllers.Sql;

[ApiController]
[Route("api/[controller]")]
public class CostosProduccionController : ControllerBase
{
    private readonly string _connectionString;

    public CostosProduccionController(IConfiguration configuration)
    {
        // Cadena de conexión desde la variable de entorno ConnectionStrings__SqlDesposte
        _connectionString = configuration.GetConnectionString("SqlDesposte")!;
    }

    // Endpoint principal: Obtener todos los costos de producción con búsqueda y filtros opcionales
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

            var resultados = await EjecutorSql.EjecutarTalCualAsync(_connectionString, sql, parametros);

            return Ok(resultados);
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { exito = false, error = ex.Message });
        }
    }
}


