using System.Diagnostics;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json.Nodes;

namespace ApiDesposte.Funciones.IA
{
    /// <summary>
    /// DeepSeek (API compatible con OpenAI): POST https://api.deepseek.com/chat/completions en modo JSON.
    /// Clave: configuración "DeepSeek:ApiKey" o variable de entorno DEEPSEEK_API_KEY.
    /// Modelo: "deepseek-chat" por defecto. Es un alias que DeepSeek mantiene apuntando a su última
    /// versión, así que no hay que actualizarlo a mano. Opcional "DeepSeek:Modelo" para fijar otro.
    /// </summary>
    public sealed class ProveedorDeepSeek(IConfiguration configuration, IHttpClientFactory httpClientFactory) : IProveedorIA
    {
        private const string Url = "https://api.deepseek.com/chat/completions";
        private const string ModeloPorDefecto = "deepseek-chat";
        private const int MaxIntentos = 3;
        private static readonly TimeSpan TimeoutPorLlamada = TimeSpan.FromSeconds(120);

        public string Nombre => "DeepSeek";

        private string ApiKey =>
            configuration["DeepSeek:ApiKey"] ?? Environment.GetEnvironmentVariable("DEEPSEEK_API_KEY") ?? "";

        private string Modelo =>
            string.IsNullOrWhiteSpace(configuration["DeepSeek:Modelo"]) ? ModeloPorDefecto : configuration["DeepSeek:Modelo"]!.Trim();

        public bool Configurado => !string.IsNullOrWhiteSpace(ApiKey);

        public async Task<RespuestaIA> GenerarJsonAsync(string promptSistema, string mensajeUsuario, TimeSpan tiempoDisponible, CancellationToken ct)
        {
            var reloj = Stopwatch.StartNew();
            string modelo = Modelo;
            string ultimoError = "";

            for (int intento = 1; intento <= MaxIntentos; intento++)
            {
                TimeSpan restante = tiempoDisponible - reloj.Elapsed;
                if (restante < TimeSpan.FromSeconds(10)) break;

                try
                {
                    return new RespuestaIA(await LlamarAsync(promptSistema, mensajeUsuario, modelo, Min(restante, TimeoutPorLlamada), ct), modelo);
                }
                catch (IATransitoriaException ex)
                {
                    ultimoError = ex.Message;
                    TimeSpan espera = UtilIA.Espera(ex, intento);
                    if (intento == MaxIntentos || reloj.Elapsed + espera + TimeSpan.FromSeconds(10) >= tiempoDisponible) break;
                    await Task.Delay(espera, ct);
                }
            }

            throw new IASaturadaException($"DeepSeek no respondió a tiempo o está saturado ({ultimoError}).");
        }

        private async Task<string> LlamarAsync(string promptSistema, string mensajeUsuario, string modelo, TimeSpan timeout, CancellationToken ct)
        {
            var cuerpo = new JsonObject
            {
                ["model"] = modelo,
                ["messages"] = new JsonArray(
                    new JsonObject { ["role"] = "system", ["content"] = promptSistema },
                    new JsonObject { ["role"] = "user", ["content"] = mensajeUsuario }),
                // Modo JSON: el prompt de sistema ya pide "JSON" explícitamente, como exige DeepSeek
                ["response_format"] = new JsonObject { ["type"] = "json_object" },
                ["temperature"] = 0.1,
                ["max_tokens"] = 8192,   // El valor por defecto (4K) puede cortar diagramas grandes
                ["stream"] = false
            };

            using var request = new HttpRequestMessage(HttpMethod.Post, Url)
            {
                Content = new StringContent(cuerpo.ToJsonString(), Encoding.UTF8, "application/json")
            };
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", ApiKey);

            var cliente = httpClientFactory.CreateClient();
            cliente.Timeout = timeout;

            HttpResponseMessage respuesta;
            try
            {
                respuesta = await cliente.SendAsync(request, ct);
            }
            catch (TaskCanceledException) when (!ct.IsCancellationRequested)
            {
                throw new IATransitoriaException($"'{modelo}' no respondió en {timeout.TotalSeconds:0} s");
            }
            catch (HttpRequestException ex)
            {
                throw new IATransitoriaException($"error de red: {ex.Message}");
            }

            using (respuesta)
            {
                string texto = await respuesta.Content.ReadAsStringAsync(ct);
                int codigo = (int)respuesta.StatusCode;

                if (!respuesta.IsSuccessStatusCode)
                {
                    string detalle = UtilIA.Recortar(texto);
                    if (codigo == 429 || codigo >= 500)
                    {
                        throw new IATransitoriaException($"HTTP {codigo}: {detalle}", UtilIA.RetryAfter(respuesta));
                    }
                    throw new InvalidOperationException(codigo switch
                    {
                        401 => "DeepSeek rechazó la clave de API (DEEPSEEK_API_KEY inválida).",
                        402 => "La cuenta de DeepSeek no tiene saldo suficiente.",
                        _ => $"DeepSeek respondió HTTP {codigo}: {detalle}"
                    });
                }

                var eleccion = JsonNode.Parse(texto)?["choices"]?[0];
                string? contenido = eleccion?["message"]?["content"]?.GetValue<string>();
                if (string.IsNullOrWhiteSpace(contenido))
                {
                    throw new InvalidOperationException("DeepSeek no devolvió contenido (respuesta vacía).");
                }
                if (eleccion?["finish_reason"]?.GetValue<string>() == "length")
                {
                    throw new InvalidOperationException("La respuesta de DeepSeek se cortó por longitud (SP demasiado grande).");
                }
                return contenido;
            }
        }

        private static TimeSpan Min(TimeSpan a, TimeSpan b) => a < b ? a : b;
    }
}
