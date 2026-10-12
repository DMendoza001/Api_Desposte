using Dapper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Globalization;
using System.Text;

namespace ApiDesposte.Controllers.Sql
{
    [ApiController]
    [Route("api/[controller]")]
    public class IndiceRotacionSemanalController : ControllerBase
    {
        private readonly string _connectionString;

        // Cache en memoria para evitar re-ejecutar el SP de larga duración en cada consulta
        private static readonly Dictionary<string, (DateTime Timestamp, List<IndiceRotacionSemanalItemDto> Items)> _cachePorRango = new(StringComparer.OrdinalIgnoreCase);
        private static readonly object _cacheLock = new();
        private static readonly TimeSpan CacheDuracion = TimeSpan.FromMinutes(15);

        /// <summary>
        /// El SP devuelve la columna "Año" (con eñe). El mapeo por defecto de Dapper no la
        /// relaciona con la propiedad "Anio", por lo que se registra un TypeMap explícito
        /// que ignora tildes al comparar los nombres de columna.
        /// </summary>
        static IndiceRotacionSemanalController()
        {
            SqlMapper.SetTypeMap(
                typeof(IndiceRotacionSemanalItemDto),
                new CustomPropertyTypeMap(
                    typeof(IndiceRotacionSemanalItemDto),
                    (tipo, nombreColumna) => tipo.GetProperties().FirstOrDefault(p =>
                        CoincideNombreColumna(p.Name, nombreColumna))!));
        }

        /// <summary>
        /// Compara un nombre de columna SQL contra una propiedad C#, ignorando mayúsculas y tildes
        /// (p. ej. "Año" y "Anio" coinciden). La base de datos es legacy y usa esa nomenclatura.
        /// </summary>
        private static bool CoincideNombreColumna(string nombrePropiedad, string nombreColumna)
        {
            if (string.Equals(nombrePropiedad, nombreColumna, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            static string Normalizar(string valor)
            {
                var descompuesto = valor.Trim().Replace("_", string.Empty).Normalize(NormalizationForm.FormD);
                var sb = new StringBuilder(descompuesto.Length);

                foreach (char c in descompuesto)
                {
                    if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                    {
                        sb.Append(c);
                    }
                }

                return sb.ToString().Normalize(NormalizationForm.FormC).ToUpperInvariant();
            }

            return Normalizar(nombrePropiedad) == Normalizar(nombreColumna);
        }

        public IndiceRotacionSemanalController(IConfiguration configuration)
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

            string f2 = string.IsNullOrWhiteSpace(fechaFin) ? "2026-10-11T00:00:00" : fechaFin.Trim();
            if (!f2.Contains('T') && !f2.Contains(' '))
            {
                f2 = $"{f2}T23:59:59";
            }

            return (f1, f2);
        }

        /// <summary>
        /// Ejecuta el procedimiento almacenado desposte.dbo.Sp_Indice_Rotacion_Semanal @Fecha1, @Fecha2
        /// utilizando almacenamiento en caché en memoria para alto rendimiento.
        /// </summary>
        private async Task<List<IndiceRotacionSemanalItemDto>> ObtenerDatosSpAsync(string f1, string f2, bool forzarRecarga = false)
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
            string sql = "EXEC desposte.dbo.Sp_Indice_Rotacion_Semanal @Fecha1, @Fecha2";
            var datos = (await db.QueryAsync<IndiceRotacionSemanalItemDto>(
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
        /// Devuelve el índice de rotación semanal de personal ejecutando:
        /// exec desposte.dbo.Sp_Indice_Rotacion_Semanal @Fecha1, @Fecha2
        /// Por defecto usa: '2026-01-01T00:00:00' hasta '2026-10-11T00:00:00'
        /// Responde con el arreglo plano de registros, listo para manipular desde JavaScript.
        /// </summary>
        [HttpGet]
        public async Task<IActionResult> ObtenerIndiceRotacionSemanal(
            [FromQuery] string? fechaInicio = null,
            [FromQuery] string? fechaFin = null,
            [FromQuery] bool recargar = false)
        {
            try
            {
                var (f1, f2) = NormalizarFechas(fechaInicio, fechaFin);
                var lista = await ObtenerDatosSpAsync(f1, f2, recargar);

                return Ok(lista);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }
    }

    /// <summary>
    /// DTO que representa un registro devuelto por desposte.dbo.Sp_Indice_Rotacion_Semanal
    /// </summary>
    public class IndiceRotacionSemanalItemDto
    {
        public int Anio { get; set; }

        public int Semana { get; set; }

        public string CecosCod { get; set; } = string.Empty;

        public string NCentroCosto { get; set; } = string.Empty;

        public int NumeroTrabajadores { get; set; }

        public int TrabajadoresInicio { get; set; }

        public int TrabajadoresFin { get; set; }

        public int Retirados { get; set; }

        public decimal IndiceRotacion { get; set; }
    }
}
