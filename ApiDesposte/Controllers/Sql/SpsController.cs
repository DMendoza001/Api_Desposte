using ApiDesposte.Funciones.IA;
using Dapper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Collections.Concurrent;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace ApiDesposte.Controllers.Sql
{
    /// <summary>
    /// Documentación visual dinámica de los Stored Procedures de una o varias bases de datos.
    /// - GET /api/Sps/bases     : bases de datos disponibles para documentar ("Sps:Bases", por defecto Desposte,PlantasCore).
    /// - GET /api/Sps/lista     : catálogo en vivo de sys.procedures de la base elegida + estado de caché.
    /// - GET /api/Sps/diagrama  : JSON del diagrama. Usa la caché si modify_date no cambió;
    ///                            si el SP es nuevo, cambió o forzar=true, lo regenera con IA.
    /// Caché en archivos: &lt;ContentRoot&gt;/Datos/sps/&lt;base&gt;/&lt;esquema&gt;.&lt;nombre&gt;.json
    /// IA: DeepSeek (DEEPSEEK_API_KEY) y, como respaldo, Gemini (GEMINI_API_KEY).
    /// El orden se cambia con "IA:Orden" (ver Funciones/IA/GeneradorJsonIA.cs).
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    public class SpsController : ControllerBase
    {
        private readonly string _connectionString;
        private readonly IConfiguration _configuration;
        private readonly GeneradorJsonIA _ia;
        private readonly string _carpetaCache;
        private readonly string _basePredeterminada;

        // Evita dos llamadas simultáneas a la IA para el mismo SP
        private static readonly ConcurrentDictionary<string, SemaphoreSlim> _bloqueos = new(StringComparer.OrdinalIgnoreCase);

        private static readonly Regex NombreValido = new(@"^[A-Za-z0-9_]{1,128}$", RegexOptions.Compiled);

        private static readonly JsonSerializerOptions JsonCache = new() { WriteIndented = true };

        public SpsController(IConfiguration configuration, GeneradorJsonIA ia, IWebHostEnvironment env)
        {
            _configuration = configuration;
            _ia = ia;
            _connectionString = configuration.GetConnectionString("SqlDesposte")
                ?? throw new InvalidOperationException("No se encontró la cadena de conexión 'SqlDesposte' (variable de entorno ConnectionStrings__SqlDesposte)");
            _carpetaCache = Path.Combine(env.ContentRootPath, "Datos", "sps");
            string catalogo = new SqlConnectionStringBuilder(_connectionString).InitialCatalog;
            _basePredeterminada = string.IsNullOrWhiteSpace(catalogo) ? "Desposte" : catalogo;
        }

        // =====================================================================
        // BASES DE DATOS PERMITIDAS
        // =====================================================================
        // Lista blanca en "Sps:Bases" (variable de entorno Sps__Bases, ej. "Desposte,PlantasCore").
        // Solo se ofrecen las que existen, están en línea y a las que la cuenta de la API tiene acceso.
        private List<string> BasesConfiguradas()
        {
            var lista = (_configuration["Sps:Bases"] ?? "Desposte,PlantasCore")
                .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .ToList();
            if (!lista.Contains(_basePredeterminada, StringComparer.OrdinalIgnoreCase)) lista.Insert(0, _basePredeterminada);
            return lista;
        }

        private async Task<List<string>> BasesDisponiblesAsync()
        {
            var configuradas = BasesConfiguradas();
            const string sql = @"
                SELECT name FROM sys.databases
                WHERE name IN @Bases AND state_desc = 'ONLINE' AND HAS_DBACCESS(name) = 1;";

            await using var db = new SqlConnection(_connectionString);
            var existentes = (await db.QueryAsync<string>(sql, new { Bases = configuradas }, commandTimeout: 30)).ToList();

            // Se respeta el orden configurado y el nombre exacto que devuelve SQL Server
            return configuradas
                .Select(b => existentes.FirstOrDefault(e => e.Equals(b, StringComparison.OrdinalIgnoreCase)))
                .Where(b => b != null)
                .Select(b => b!)
                .ToList();
        }

        /// <summary>Nombre exacto de la base si está permitida; null si no lo está. Vacío = base predeterminada.</summary>
        private async Task<string?> ResolverBaseAsync(string? baseDatos)
        {
            if (string.IsNullOrWhiteSpace(baseDatos)) return _basePredeterminada;
            var disponibles = await BasesDisponiblesAsync();
            return disponibles.FirstOrDefault(b => b.Equals(baseDatos.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        // Conexión abierta en la base elegida (ya validada contra la lista blanca)
        private async Task<SqlConnection> AbrirEnBaseAsync(string baseDatos)
        {
            var db = new SqlConnection(_connectionString);
            await db.OpenAsync();
            if (!db.Database.Equals(baseDatos, StringComparison.OrdinalIgnoreCase))
            {
                await db.ChangeDatabaseAsync(baseDatos);
            }
            return db;
        }

        // =====================================================================
        // GET /api/Sps/bases
        // =====================================================================
        [HttpGet("bases")]
        public async Task<IActionResult> Bases()
        {
            try
            {
                return Ok(new { exito = true, predeterminada = _basePredeterminada, bases = await BasesDisponiblesAsync() });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, mensaje = $"Error al listar las bases de datos: {ex.Message}" });
            }
        }

        // =====================================================================
        // GET /api/Sps/lista
        // =====================================================================
        [HttpGet("lista")]
        public async Task<IActionResult> Lista([FromQuery(Name = "base")] string? baseDatos = null)
        {
            const string sql = @"
                SELECT s.name AS Esquema, p.name AS NombreSP,
                       p.create_date AS FechaCreacion, p.modify_date AS UltimaModificacion
                FROM sys.procedures p
                INNER JOIN sys.schemas s ON p.schema_id = s.schema_id
                WHERE p.is_ms_shipped = 0
                  -- Todos los SPs del usuario (Sp_, Lista_, Inserta_, Actualiza_...). Excluye los SPs que SSMS crea para sus diagramas (sp_alterdiagram, sp_helpdiagrams, …)
                  AND p.name NOT LIKE 'sp[_]%diagram%'
                ORDER BY p.name ASC;";

            try
            {
                string? baseElegida = await ResolverBaseAsync(baseDatos);
                if (baseElegida == null)
                {
                    return BadRequest(new { exito = false, mensaje = $"La base de datos '{baseDatos}' no está habilitada para documentar." });
                }

                await using var db = await AbrirEnBaseAsync(baseElegida);
                var filas = (await db.QueryAsync<SpCatalogoDto>(sql, commandTimeout: 60)).ToList();

                var resultado = filas.Select(f =>
                {
                    var cache = LeerCache(baseElegida, f.Esquema, f.NombreSP);
                    return new
                    {
                        nombre = f.NombreSP,
                        baseDatos = baseElegida,
                        esquema = f.Esquema,
                        fechaCreacion = f.FechaCreacion,
                        modifyDate = f.UltimaModificacion,
                        tieneCache = cache != null,
                        cacheVigente = cache != null && MismaFecha(cache.UltimaModificacionSQL, f.UltimaModificacion),
                        fechaGeneracion = cache?.FechaGeneracion
                    };
                }).ToList();

                return Ok(new
                {
                    exito = true,
                    baseDatos = baseElegida,
                    total = resultado.Count,
                    iaConfigurada = _ia.Activos().Count > 0,
                    proveedoresIA = _ia.Activos().Select(p => p.Nombre),
                    registros = resultado
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, mensaje = $"Error al listar los SPs: {ex.Message}" });
            }
        }

        // =====================================================================
        // GET /api/Sps/diagrama?nombre=Sp_X&base=Desposte&forzar=false
        // =====================================================================
        [HttpGet("diagrama")]
        public async Task<IActionResult> Diagrama(
            [FromQuery] string? nombre,
            [FromQuery(Name = "base")] string? baseDatos = null,
            [FromQuery] bool forzar = false)
        {
            if (string.IsNullOrWhiteSpace(nombre) || !NombreValido.IsMatch(nombre.Trim()))
            {
                return BadRequest(new { exito = false, mensaje = "Parámetro 'nombre' inválido. Solo letras, números y guion bajo." });
            }
            nombre = nombre.Trim();

            string? baseElegida;
            SpDefinicionDto? definicion;
            try
            {
                baseElegida = await ResolverBaseAsync(baseDatos);
                if (baseElegida == null)
                {
                    return BadRequest(new { exito = false, mensaje = $"La base de datos '{baseDatos}' no está habilitada para documentar." });
                }
                definicion = await ObtenerDefinicionAsync(baseElegida, nombre);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, mensaje = $"Error al leer el SP en SQL Server: {ex.Message}" });
            }

            if (definicion == null)
            {
                return NotFound(new { exito = false, mensaje = $"No existe el procedimiento '{nombre}' en {baseElegida}." });
            }

            var cache = LeerCache(baseElegida, definicion.Esquema, definicion.NombreSP);
            if (!forzar && cache != null && MismaFecha(cache.UltimaModificacionSQL, definicion.UltimaModificacion))
            {
                return Ok(Respuesta(cache, "cache"));
            }

            if (string.IsNullOrWhiteSpace(definicion.CodigoSQL))
            {
                return StatusCode(422, new
                {
                    exito = false,
                    mensaje = "SQL Server no devolvió la definición del SP (puede estar cifrado WITH ENCRYPTION o el usuario no tiene permiso VIEW DEFINITION)."
                });
            }

            if (_ia.Activos().Count == 0)
            {
                // Sin IA: si hay una versión anterior se entrega marcada como desactualizada
                if (cache != null) return Ok(Respuesta(cache, "cache-desactualizada"));
                return StatusCode(503, new
                {
                    exito = false,
                    mensaje = "No hay diagrama en caché y no hay ninguna IA configurada. Define la variable de entorno DEEPSEEK_API_KEY (o GEMINI_API_KEY) en el servidor."
                });
            }

            var bloqueo = _bloqueos.GetOrAdd($"{baseElegida}.{definicion.Esquema}.{definicion.NombreSP}", _ => new SemaphoreSlim(1, 1));
            await bloqueo.WaitAsync(HttpContext.RequestAborted);
            try
            {
                // Otra petición pudo haber generado el diagrama mientras se esperaba el bloqueo
                cache = LeerCache(baseElegida, definicion.Esquema, definicion.NombreSP);
                if (!forzar && cache != null && MismaFecha(cache.UltimaModificacionSQL, definicion.UltimaModificacion))
                {
                    return Ok(Respuesta(cache, "cache"));
                }

                var (diagrama, modeloUsado) = await _ia.GenerarAsync(
                    PromptSistema,
                    $"Base de datos: {baseElegida}. Esquema: {definicion.Esquema}. Procedimiento: {definicion.NombreSP}.\n\n{definicion.CodigoSQL}",
                    d => ProcesarDiagrama(d, definicion, baseElegida),
                    TiempoMaximoIA,
                    HttpContext.RequestAborted);

                var nuevo = new SpCacheEntrada
                {
                    NombreSP = definicion.NombreSP,
                    BaseDatos = baseElegida,
                    Esquema = definicion.Esquema,
                    UltimaModificacionSQL = definicion.UltimaModificacion,
                    FechaGeneracion = DateTime.Now,
                    Modelo = modeloUsado,
                    Diagrama = diagrama
                };
                GuardarCache(nuevo);
                return Ok(Respuesta(nuevo, "ia"));
            }
            catch (OperationCanceledException)
            {
                return StatusCode(499, new { exito = false, mensaje = "La solicitud fue cancelada." });
            }
            catch (Exception ex)
            {
                if (cache != null)
                {
                    var respuesta = Respuesta(cache, "cache-desactualizada");
                    respuesta["advertencia"] = $"No se pudo regenerar el diagrama: {ex.Message}";
                    return Ok(respuesta);
                }
                // 503 = IA saturada (temporal) · 502 = otra falla de la IA
                return StatusCode(ex is IASaturadaException ? 503 : 502,
                    new { exito = false, mensaje = $"No se pudo generar el diagrama: {ex.Message}" });
            }
            finally
            {
                bloqueo.Release();
            }
        }

        // =====================================================================
        // SQL SERVER
        // =====================================================================
        private async Task<SpDefinicionDto?> ObtenerDefinicionAsync(string baseDatos, string nombre)
        {
            const string sql = @"
                SELECT TOP 1 p.name AS NombreSP, s.name AS Esquema,
                       p.modify_date AS UltimaModificacion,
                       OBJECT_DEFINITION(p.object_id) AS CodigoSQL
                FROM sys.procedures p
                INNER JOIN sys.schemas s ON p.schema_id = s.schema_id
                WHERE p.name = @NombreSP
                  AND p.is_ms_shipped = 0
                ORDER BY CASE WHEN s.name = 'dbo' THEN 0 ELSE 1 END;";

            await using var db = await AbrirEnBaseAsync(baseDatos);
            return await db.QueryFirstOrDefaultAsync<SpDefinicionDto>(sql, new { NombreSP = nombre }, commandTimeout: 60);
        }

        // =====================================================================
        // IA (DeepSeek / Gemini) → Funciones/IA/GeneradorJsonIA
        // =====================================================================
        // El visor espera hasta 240 s: se deja margen para leer SQL y responder
        private static readonly TimeSpan TiempoMaximoIA = TimeSpan.FromSeconds(200);

        private static JsonObject ProcesarDiagrama(JsonObject diagrama, SpDefinicionDto def, string baseDatos)
        {
            ValidarEstructura(diagrama);

            // Valores autoritativos desde SQL Server
            diagrama["spName"] = def.NombreSP;
            diagrama["esquema"] = $"{baseDatos}.{def.Esquema}";
            diagrama["ultimaModificacion"] = def.UltimaModificacion.ToString("yyyy-MM-ddTHH:mm:ss");
            return diagrama;
        }

        private static void ValidarEstructura(JsonObject d)
        {
            var faltantes = new List<string>();
            if (d["etapa1_fuentes"]?["tarjetas"] is not JsonArray) faltantes.Add("etapa1_fuentes.tarjetas");
            if (d["etapa2_transformacion"]?["tarjetas"] is not JsonArray) faltantes.Add("etapa2_transformacion.tarjetas");
            if (d["etapa3_catalogos"] is not JsonObject) faltantes.Add("etapa3_catalogos");
            if (d["etapa4_salida"]?["columnas"] is not JsonArray) faltantes.Add("etapa4_salida.columnas");

            if (faltantes.Count > 0)
            {
                throw new InvalidOperationException($"El JSON no cumple el esquema; faltan: {string.Join(", ", faltantes)}.");
            }
        }

        // =====================================================================
        // CACHÉ EN ARCHIVOS
        // =====================================================================
        // Cada base en su carpeta: Datos/sps/<base>/<esquema>.<nombre>.json
        private string RutaCache(string baseDatos, string esquema, string nombre) =>
            Path.Combine(_carpetaCache, baseDatos, $"{esquema}.{nombre}.json");

        // Ruta anterior (sin carpeta por base), usada solo para la base predeterminada
        private string RutaCacheAnterior(string esquema, string nombre) =>
            Path.Combine(_carpetaCache, $"{esquema}.{nombre}.json");

        private SpCacheEntrada? LeerCache(string baseDatos, string esquema, string nombre)
        {
            string ruta = RutaCache(baseDatos, esquema, nombre);
            if (!System.IO.File.Exists(ruta) && baseDatos.Equals(_basePredeterminada, StringComparison.OrdinalIgnoreCase))
            {
                ruta = RutaCacheAnterior(esquema, nombre);
            }
            if (!System.IO.File.Exists(ruta)) return null;
            try
            {
                return JsonSerializer.Deserialize<SpCacheEntrada>(System.IO.File.ReadAllText(ruta));
            }
            catch
            {
                return null; // Archivo dañado: se regenera
            }
        }

        private void GuardarCache(SpCacheEntrada entrada)
        {
            string ruta = RutaCache(entrada.BaseDatos, entrada.Esquema, entrada.NombreSP);
            Directory.CreateDirectory(Path.GetDirectoryName(ruta)!);
            string temporal = ruta + ".tmp";
            System.IO.File.WriteAllText(temporal, JsonSerializer.Serialize(entrada, JsonCache), Encoding.UTF8);
            System.IO.File.Move(temporal, ruta, overwrite: true);

            // Si existía la versión en la ruta anterior, se elimina para no dejar duplicados
            if (entrada.BaseDatos.Equals(_basePredeterminada, StringComparison.OrdinalIgnoreCase))
            {
                string anterior = RutaCacheAnterior(entrada.Esquema, entrada.NombreSP);
                if (System.IO.File.Exists(anterior)) System.IO.File.Delete(anterior);
            }
        }

        // modify_date tiene precisión de milisegundos; se compara al segundo para evitar falsos MISS
        private static bool MismaFecha(DateTime a, DateTime b) =>
            Math.Abs((a - b).TotalSeconds) < 1;

        private static JsonObject Respuesta(SpCacheEntrada c, string origen) => new()
        {
            ["exito"] = true,
            ["origen"] = origen,
            ["baseDatos"] = c.BaseDatos,
            ["fechaGeneracion"] = c.FechaGeneracion.ToString("yyyy-MM-ddTHH:mm:ss"),
            ["modelo"] = c.Modelo,
            ["diagrama"] = c.Diagrama?.DeepClone()
        };

        // =====================================================================
        // PROMPT DEL SISTEMA (sección 5 de la especificación + contrato JSON)
        // =====================================================================
        private const string PromptSistema = """
Eres un arquitecto de datos experto en SQL Server. Analiza el Stored Procedure proporcionado y extrae su flujo ETL completo en formato JSON estricto siguiendo el esquema definido.
Reglas de extracción:
1. Etapa 1 (Fuentes): Identifica las tablas base (Tabla10, #Temp), funciones (Fechas_OrdeProduccion), SPs previos ejecutados, filtros de grupopuntoproceso, rangos de fecha y llaves exactas de JOIN. Ordena las tarjetas en el orden en que fluye el dato.
2. Etapa 2 (Transformación): Agrupa cada CTE, tabla temporal o bloque lógico indicando explícitamente los códigos de PuntoProceso, códigos de plantilla incluidos o excluidos (IN / NOT IN) y valores de ttra (13 = Origen/Insumo, 4 = Resultado/Producción). Resume cualquier cálculo de saldo, merma o cascada FIFO en nodoCentral (si no existe, usa null).
3. Etapa 3 (Catálogos): Lista en "pasos" los LEFT JOIN o INNER JOIN con tablas maestras de artículos/productos y las llaves de cruce. Usa "badges" para los niveles, tipos o destinos que asigna el SP (vacío si no hay).
4. Etapa 4 (Salida): Lista exactamente las columnas del SELECT final y en pieFiltro las cláusulas WHERE / ORDER BY de cierre.
Escribe todos los textos en español, breves (máximo ~90 caracteres por item). estiloCabecera admite "primary", "secondary" o "dark".
Devuelve únicamente JSON válido sin bloques markdown ni texto adicional, con exactamente esta estructura:
{
  "spName": "string",
  "esquema": "Desposte.dbo",
  "badge": "string corto (ej. 'Desposte.dbo · SKU 040508029')",
  "subtitulo": "string (qué hace el SP en una línea)",
  "ultimaModificacion": "string",
  "etapa1_fuentes": {
    "titulo": "1. Fuentes ...",
    "tarjetas": [ { "titulo": "string", "estiloCabecera": "primary", "joinLabel": "string opcional", "items": ["string"] } ]
  },
  "etapa2_transformacion": {
    "titulo": "2. ...",
    "tarjetas": [ { "titulo": "string", "subbloques": [ { "subtitulo": "string", "bullets": ["string"] } ] } ],
    "nodoCentral": { "titulo": "string", "descripcion": "string", "formula": "string" }
  },
  "etapa3_catalogos": {
    "titulo": "3. ...",
    "tituloTarjeta": "string",
    "pasos": [ { "caja": "string", "detalle": "string", "tipoJoin": "string" } ],
    "badges": ["string"]
  },
  "etapa4_salida": {
    "tituloTarjeta": "Resultado Final",
    "columnas": ["string"],
    "pieFiltro": "string"
  }
}
""";
    }

    public class SpCatalogoDto
    {
        public string Esquema { get; set; } = "";
        public string NombreSP { get; set; } = "";
        public DateTime FechaCreacion { get; set; }
        public DateTime UltimaModificacion { get; set; }
    }

    public class SpDefinicionDto
    {
        public string NombreSP { get; set; } = "";
        public string Esquema { get; set; } = "";
        public DateTime UltimaModificacion { get; set; }
        public string? CodigoSQL { get; set; }
    }

    public class SpCacheEntrada
    {
        public string NombreSP { get; set; } = "";
        public string BaseDatos { get; set; } = "Desposte";
        public string Esquema { get; set; } = "dbo";
        public DateTime UltimaModificacionSQL { get; set; }
        public DateTime FechaGeneracion { get; set; }
        public string Modelo { get; set; } = "";
        public JsonObject? Diagrama { get; set; }
    }
}
