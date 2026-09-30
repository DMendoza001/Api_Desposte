using Dapper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Data;

namespace ApiDesposte.Controllers.Sql;

[ApiController]
[Route("api/[controller]")]
public class DesposteController : ControllerBase
{
    private readonly string _connectionString;

    public DesposteController(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("SqlDesposte")!;
    }

    // Endpoint para ejecutar el Stored Procedure con rango de fechas (usado en Office Scripts)
    [HttpGet("atenciones-topico")]
    public async Task<IActionResult> ObtenerAtencionesTopico([FromQuery] string fechaInicial, [FromQuery] string fechaFinal)
    {
        try
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            string f1 = $"{fechaInicial}T00:00:00";
            string f2 = $"{fechaFinal}T23:59:59";

            string sql = "EXEC PlantasCore.dbo.Sp_Informe_AtencionesTopico @Fecha1, @Fecha2";

            var resultados = await db.QueryAsync(sql, new { Fecha1 = f1, Fecha2 = f2 });

            return Ok(resultados);
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { exito = false, error = ex.Message });
        }
    }
}