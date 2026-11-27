using Dapper;
using Microsoft.Data.SqlClient;

namespace ApiDesposte.Funciones.Sql
{
    /// <summary>
    /// Ejecución de consultas y procedimientos almacenados de SQL Server devolviendo
    /// las filas tal cual (mismas columnas, en el orden de SQL Server, sin campos calculados).
    /// La cadena de conexión la resuelve cada controlador con configuration.GetConnectionString("SqlDesposte").
    /// </summary>
    public static class EjecutorSql
    {
        /// <summary>20 minutos de espera por SP. El dashboard debe usar APP_CONFIG.timeoutMs acorde.</summary>
        public const int TimeoutPorDefecto = 1200;

        /// <summary>
        /// Ejecuta la consulta con Dapper (parámetros seguros, sin concatenar valores)
        /// y devuelve cada fila como diccionario columna → valor.
        /// </summary>
        public static async Task<List<Dictionary<string, object?>>> EjecutarTalCualAsync(
            string connectionString,
            string sql,
            object? parametros = null,
            int commandTimeout = TimeoutPorDefecto)
        {
            await using var db = new SqlConnection(connectionString);
            var filas = await db.QueryAsync(sql, parametros, commandTimeout: commandTimeout);

            return filas
                .Select(f => new Dictionary<string, object?>((IDictionary<string, object?>)f))
                .ToList();
        }
    }
}
