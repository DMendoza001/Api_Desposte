using System.Diagnostics;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace ApiDesposte.Funciones.IA
{
    /// <summary>
    /// Genera un objeto JSON con el primer proveedor de IA que responda bien.
    /// Orden: configuración "IA:Orden" (ej. "deepseek,gemini"); por defecto DeepSeek y luego Gemini.
    /// Solo participan los proveedores con clave configurada. Si uno falla (saturado, sin saldo,
    /// JSON inválido…) se pasa al siguiente dentro del tiempo máximo total.
    /// Se registra como singleton en Program.cs.
    /// </summary>
    public sealed class GeneradorJsonIA
    {
        private readonly IConfiguration _configuration;
        private readonly Dictionary<string, IProveedorIA> _proveedores;

        public GeneradorJsonIA(IConfiguration configuration, IHttpClientFactory httpClientFactory)
        {
            _configuration = configuration;
            _proveedores = new Dictionary<string, IProveedorIA>(StringComparer.OrdinalIgnoreCase)
            {
                ["deepseek"] = new ProveedorDeepSeek(configuration, httpClientFactory),
                ["gemini"] = new ProveedorGemini(configuration, httpClientFactory)
            };
        }

        /// <summary>Proveedores configurados, en el orden en que se prueban.</summary>
        public IReadOnlyList<IProveedorIA> Activos()
        {
            string orden = _configuration["IA:Orden"] ?? "deepseek,gemini";
            var lista = orden.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Where(_proveedores.ContainsKey)
                .Select(n => _proveedores[n])
                .ToList();

            // Proveedores no mencionados en IA:Orden van al final
            lista.AddRange(_proveedores.Values.Where(p => !lista.Contains(p)));
            return lista.Where(p => p.Configurado).ToList();
        }

        /// <summary>
        /// Prueba los proveedores en orden. <paramref name="procesar"/> convierte y valida el JSON;
        /// si lanza una excepción se considera una respuesta inválida y se prueba el siguiente proveedor.
        /// </summary>
        /// <returns>El objeto procesado y "Proveedor · modelo" que lo generó.</returns>
        public async Task<(JsonObject Resultado, string Modelo)> GenerarAsync(
            string promptSistema,
            string mensajeUsuario,
            Func<JsonObject, JsonObject> procesar,
            TimeSpan tiempoMaximo,
            CancellationToken ct)
        {
            var proveedores = Activos();
            if (proveedores.Count == 0)
            {
                throw new InvalidOperationException("No hay ninguna IA configurada (DEEPSEEK_API_KEY o GEMINI_API_KEY).");
            }

            var reloj = Stopwatch.StartNew();
            var errores = new List<string>();
            bool todosSaturados = true;

            foreach (var proveedor in proveedores)
            {
                TimeSpan restante = tiempoMaximo - reloj.Elapsed;
                if (restante < TimeSpan.FromSeconds(15))
                {
                    errores.Add($"{proveedor.Nombre}: no quedó tiempo para intentarlo");
                    break;
                }

                try
                {
                    var respuesta = await proveedor.GenerarJsonAsync(promptSistema, mensajeUsuario, restante, ct);
                    return (procesar(ConvertirAObjeto(respuesta.Texto, proveedor.Nombre)), $"{proveedor.Nombre} · {respuesta.Modelo}");
                }
                catch (OperationCanceledException) when (ct.IsCancellationRequested)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    if (ex is not IASaturadaException) todosSaturados = false;
                    errores.Add($"{proveedor.Nombre}: {ex.Message}");
                }
            }

            string detalle = string.Join(" | ", errores);
            if (todosSaturados)
            {
                throw new IASaturadaException($"Las IA están saturadas en este momento; intenta de nuevo en unos minutos. {detalle}");
            }
            throw new InvalidOperationException(detalle);
        }

        private static JsonObject ConvertirAObjeto(string texto, string proveedor)
        {
            // Por si el modelo envolvió la respuesta en un bloque markdown
            string json = Regex.Replace(texto.Trim(), @"^```(?:json)?\s*|\s*```$", "");
            return JsonNode.Parse(json) as JsonObject
                ?? throw new InvalidOperationException($"La respuesta de {proveedor} no es un objeto JSON.");
        }
    }
}