using System.Diagnostics;
using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace ApiDesposte.Funciones.IA
{
    /// <summary>
    /// Google Gemini: generateContent con responseMimeType = application/json.
    /// Clave: configuración "Gemini:ApiKey" o variable de entorno GEMINI_API_KEY.
    /// Modelo automático: último Flash estable según ListModels (opcional "Gemini:Familia" = flash/pro,
    /// o "Gemini:Modelo" para fijarlo a mano). La lista se guarda 24 h; si Google retira un modelo (404)
    /// se vuelve a consultar. Ante saturación (429/5xx/timeout) reintenta y pasa al siguiente modelo.
    /// </summary>
    public sealed class ProveedorGemini(IConfiguration configuration, IHttpClientFactory httpClientFactory) : IProveedorIA
    {
        private const string UrlBase = "https://generativelanguage.googleapis.com/v1beta";
        private const string ModeloRespaldo = "gemini-flash-latest";   // Alias que Google mantiene apuntando al último Flash
        private static readonly TimeSpan VigenciaModelos = TimeSpan.FromHours(24);
        private static readonly SemaphoreSlim _bloqueoModelos = new(1, 1);
        private static (List<string> Modelos, DateTime Fecha)? _modelosResueltos;

        private const int MaxModelosAProbar = 4;
        private const int MaxIntentosPorModelo = 3;
        private static readonly TimeSpan TimeoutPorLlamada = TimeSpan.FromSeconds(90);

        public string Nombre => "Gemini";

        private string ApiKey =>
            configuration["Gemini:ApiKey"] ?? Environment.GetEnvironmentVariable("GEMINI_API_KEY") ?? "";

        public bool Configurado => !string.IsNullOrWhiteSpace(ApiKey);

        private sealed class ModeloNoDisponibleException(string mensaje) : Exception(mensaje);

        public async Task<RespuestaIA> GenerarJsonAsync(string promptSistema, string mensajeUsuario, TimeSpan tiempoDisponible, CancellationToken ct)
        {
            var (modelos, automatico) = await ResolverModelosAsync(ct);
            var reloj = Stopwatch.StartNew();
            var saturados = new List<string>();
            var retirados = new List<string>();

            foreach (string modelo in modelos.Take(MaxModelosAProbar))
            {
                for (int intento = 1; intento <= MaxIntentosPorModelo; intento++)
                {
                    TimeSpan restante = tiempoDisponible - reloj.Elapsed;
                    if (restante < TimeSpan.FromSeconds(10)) break;

                    try
                    {
                        TimeSpan timeout = restante < TimeoutPorLlamada ? restante : TimeoutPorLlamada;
                        return new RespuestaIA(await LlamarAsync(promptSistema, mensajeUsuario, modelo, timeout, ct), modelo);
                    }
                    catch (ModeloNoDisponibleException)
                    {
                        retirados.Add(modelo);
                        break;
                    }
                    catch (IATransitoriaException ex)
                    {
                        TimeSpan espera = UtilIA.Espera(ex, intento);
                        bool hayTiempo = reloj.Elapsed + espera + TimeSpan.FromSeconds(5) < tiempoDisponible;
                        if (intento == MaxIntentosPorModelo || !hayTiempo)
                        {
                            saturados.Add(modelo);
                            break;
                        }
                        await Task.Delay(espera, ct);
                    }
                }

                if (reloj.Elapsed >= tiempoDisponible) break;
            }

            // Si algún modelo fue retirado, la próxima solicitud vuelve a consultar la lista de Google
            if (retirados.Count > 0 && automatico) _modelosResueltos = null;

            if (saturados.Count > 0 || retirados.Count == 0)
            {
                throw new IASaturadaException(
                    $"Gemini está saturado en este momento (se intentó con {string.Join(", ", saturados.DefaultIfEmpty(modelos.FirstOrDefault() ?? ModeloRespaldo))}).");
            }
            throw new InvalidOperationException(
                $"Ningún modelo de Gemini está disponible para esta clave (retirados: {string.Join(", ", retirados)}).");
        }

        // ---------------------------------------------------------------------
        // Selección dinámica del modelo
        // ---------------------------------------------------------------------
        private async Task<(List<string> Modelos, bool Automatico)> ResolverModelosAsync(CancellationToken ct)
        {
            string? fijo = configuration["Gemini:Modelo"];
            if (!string.IsNullOrWhiteSpace(fijo)) return (new List<string> { fijo.Trim() }, false);

            await _bloqueoModelos.WaitAsync(ct);
            try
            {
                if (_modelosResueltos is { } r && DateTime.Now - r.Fecha < VigenciaModelos)
                {
                    return (r.Modelos, true);
                }

                string familia = (configuration["Gemini:Familia"] ?? "flash").Trim().ToLowerInvariant();
                var modelos = await BuscarModelosAsync(familia, ct);
                if (modelos.Count == 0) modelos.Add(ModeloRespaldo);
                _modelosResueltos = (modelos, DateTime.Now);
                return (modelos, true);
            }
            finally
            {
                _bloqueoModelos.Release();
            }
        }

        // Modelos estables de la familia, del más nuevo al más antiguo (lista vacía si no se pudo consultar).
        // Se excluyen preview / exp / lite / tts / image / live y versiones fijadas (-001).
        private async Task<List<string>> BuscarModelosAsync(string familia, CancellationToken ct)
        {
            try
            {
                using var request = new HttpRequestMessage(HttpMethod.Get, $"{UrlBase}/models?pageSize=1000");
                request.Headers.Add("x-goog-api-key", ApiKey);

                var cliente = httpClientFactory.CreateClient();
                cliente.Timeout = TimeSpan.FromSeconds(20);
                using var respuesta = await cliente.SendAsync(request, ct);
                if (!respuesta.IsSuccessStatusCode) return new List<string>();

                if (JsonNode.Parse(await respuesta.Content.ReadAsStringAsync(ct))?["models"] is not JsonArray modelos)
                {
                    return new List<string>();
                }

                // Solo nombres estables: "gemini-<versión>-<familia>" (ej. gemini-3.8-flash, gemini-3-pro)
                var patron = new Regex($@"^models/gemini-(\d+(?:\.\d+)?)-{Regex.Escape(familia)}$", RegexOptions.IgnoreCase);

                return modelos
                    .Select(m => new
                    {
                        Nombre = m?["name"]?.GetValue<string>() ?? "",
                        Metodos = (m?["supportedGenerationMethods"] as JsonArray)?
                            .Select(x => x?.GetValue<string>()).ToList() ?? []
                    })
                    .Where(m => m.Metodos.Contains("generateContent"))
                    .Select(m => new { m.Nombre, Coincidencia = patron.Match(m.Nombre) })
                    .Where(m => m.Coincidencia.Success)
                    .Select(m =>
                    {
                        string numero = m.Coincidencia.Groups[1].Value;
                        return new
                        {
                            Modelo = m.Nombre["models/".Length..],
                            Version = Version.TryParse(numero.Contains('.') ? numero : numero + ".0", out var v) ? v : new Version(0, 0)
                        };
                    })
                    .OrderByDescending(m => m.Version)
                    .Select(m => m.Modelo)
                    .ToList();
            }
            catch (Exception ex) when (ex is not OperationCanceledException || !ct.IsCancellationRequested)
            {
                return new List<string>(); // Sin lista de modelos: se usa el alias de respaldo
            }
        }

        // ---------------------------------------------------------------------
        // Llamada a generateContent
        // ---------------------------------------------------------------------
        private async Task<string> LlamarAsync(string promptSistema, string mensajeUsuario, string modelo, TimeSpan timeout, CancellationToken ct)
        {
            string url = $"{UrlBase}/models/{Uri.EscapeDataString(modelo)}:generateContent";

            var cuerpo = new JsonObject
            {
                ["systemInstruction"] = new JsonObject
                {
                    ["parts"] = new JsonArray(new JsonObject { ["text"] = promptSistema })
                },
                ["contents"] = new JsonArray(new JsonObject
                {
                    ["role"] = "user",
                    ["parts"] = new JsonArray(new JsonObject { ["text"] = mensajeUsuario })
                }),
                ["generationConfig"] = new JsonObject
                {
                    ["responseMimeType"] = "application/json",
                    ["temperature"] = 0.1
                }
            };

            using var request = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new StringContent(cuerpo.ToJsonString(), Encoding.UTF8, "application/json")
            };
            request.Headers.Add("x-goog-api-key", ApiKey);

            var cliente = httpClientFactory.CreateClient();
            cliente.Timeout = timeout;

            HttpResponseMessage respuesta;
            try
            {
                respuesta = await cliente.SendAsync(request, ct);
            }
            catch (TaskCanceledException) when (!ct.IsCancellationRequested)
            {
                throw new IATransitoriaException($"El modelo '{modelo}' no respondió a tiempo.");
            }
            catch (HttpRequestException ex)
            {
                throw new IATransitoriaException($"Error de red al llamar a '{modelo}': {ex.Message}");
            }

            using (respuesta)
            {
                string texto = await respuesta.Content.ReadAsStringAsync(ct);
                int codigo = (int)respuesta.StatusCode;

                if (!respuesta.IsSuccessStatusCode)
                {
                    string detalle = UtilIA.Recortar(texto);
                    if (codigo == 404)
                    {
                        throw new ModeloNoDisponibleException($"El modelo '{modelo}' ya no está disponible en Gemini: {detalle}");
                    }
                    // 429 = límite de consultas · 500/502/503/504 = saturación o falla temporal de Google
                    if (codigo == 429 || codigo >= 500)
                    {
                        throw new IATransitoriaException($"Gemini respondió HTTP {codigo} ({modelo}): {detalle}", UtilIA.RetryAfter(respuesta));
                    }
                    throw new InvalidOperationException($"Gemini respondió HTTP {codigo} ({modelo}): {detalle}");
                }

                string? contenido = JsonNode.Parse(texto)?["candidates"]?[0]?["content"]?["parts"]?[0]?["text"]?.GetValue<string>();
                if (string.IsNullOrWhiteSpace(contenido))
                {
                    throw new InvalidOperationException("Gemini no devolvió contenido (respuesta vacía o bloqueada).");
                }
                return contenido;
            }
        }
    }
}
