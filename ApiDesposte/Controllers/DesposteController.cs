using Dapper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Data;

[ApiController]
[Route("api/[controller]")]
public class DesposteController : ControllerBase
{
    private readonly string _connectionString;

    public DesposteController(IConfiguration configuration)
    {
        // Obtiene la cadena de conexión configurada en el appsettings.json
        _connectionString = configuration.GetConnectionString("SqlDesposte")!;
    }

    // Endpoint de prueba para verificar la conexión con SQL Server
    [HttpGet("probar-conexion")]
    public async Task<IActionResult> ProbarConexion()
    {
        try
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            
            // Reemplazo de Set cn = New ADODB.Connection / cn.Open
            string query = "SELECT @@VERSION AS VersionSQL";
            var resultado = await db.QueryFirstOrDefaultAsync<string>(query);

            return Ok(new { exito = true, mensaje = "Conexión a DMENDOZA\\SQLEXPRESS exitosa", version = resultado });
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { exito = false, error = ex.Message });
        }
    }

    // Endpoint para ejecutar el Stored Procedure con rango de fechas
    [HttpGet("atenciones-topico")]
    public async Task<IActionResult> ObtenerAtencionesTopico([FromQuery] string fechaInicial, [FromQuery] string fechaFinal)
    {
    try
    {
        using IDbConnection db = new SqlConnection(_connectionString);

        // Formatear las fechas tal como lo hacías en VBA
        string f1 = $"{fechaInicial}T00:00:00";
        string f2 = $"{fechaFinal}T23:59:59";

        // Consulta usando el Stored Procedure
        string sql = "EXEC PlantasCore.dbo.Sp_Informe_AtencionesTopico @Fecha1, @Fecha2";

        // Dapper ejecuta el SP y mapea dinámicamente el resultado a objetos JSON
        var resultados = await db.QueryAsync(sql, new { Fecha1 = f1, Fecha2 = f2 });

        return Ok(resultados);
    }
    catch (Exception ex)
    {
        return StatusCode(500, new { exito = false, error = ex.Message });
    }
    }



 [HttpGet("dashboard-topico-pivot")]
public async Task<IActionResult> ObtenerDashboardPivot([FromQuery] string fechaInicial, [FromQuery] string fechaFinal)
{
    try
    {
        using IDbConnection db = new SqlConnection(_connectionString);

        // Fuerza el Lunes como día 1 de la semana
        string sql = @"
            SET DATEFIRST 1;

            WITH DatosProcesados AS (
                SELECT 
                    'Semana ' + CAST(DATEPART(WEEK, a.Fecha) AS VARCHAR) AS Semana,
                    ISNULL(CAST(a.Ceco AS VARCHAR), 'SIN CC') AS NCentroCosto,
                    ISNULL(b.Diagnostico, 'SIN DIAGNOSTICO') AS Diagnostico,
                    DATEPART(WEEKDAY, a.Fecha) AS NumDia -- 1=Lunes, 2=Martes, ..., 7=Domingo
                FROM PlantasCore.dbo.Personal_AtencionesTopico a
                INNER JOIN PlantasCore.dbo.Personal_Diagnosticos b 
                    ON a.IdDiagnostico = b.Id
                WHERE a.Fecha >= @f1 AND a.Fecha <= @f2
            )
            SELECT 
                Semana,
                NCentroCosto,
                Diagnostico,
                ISNULL([1], 0) AS Lunes,
                ISNULL([2], 0) AS Martes,
                ISNULL([3], 0) AS Miercoles,
                ISNULL([4], 0) AS Jueves,
                ISNULL([5], 0) AS Viernes,
                ISNULL([6], 0) AS Sabado,
                ISNULL([7], 0) AS Domingo,
                (ISNULL([1], 0) + ISNULL([2], 0) + ISNULL([3], 0) + 
                 ISNULL([4], 0) + ISNULL([5], 0) + ISNULL([6], 0) + ISNULL([7], 0)) AS TotalGeneral
            FROM DatosProcesados
            PIVOT (
                COUNT(NumDia)
                FOR NumDia IN ([1], [2], [3], [4], [5], [6], [7])
            ) AS PivotTable
            ORDER BY Semana, NCentroCosto, Diagnostico;";

        string f1 = $"{fechaInicial}T00:00:00";
        string f2 = $"{fechaFinal}T23:59:59";

        var resultados = await db.QueryAsync(sql, new { f1, f2 });
        return Ok(resultados);
    }
    catch (Exception ex)
    {
        return StatusCode(500, new { exito = false, error = ex.Message });
    }
}
    
}