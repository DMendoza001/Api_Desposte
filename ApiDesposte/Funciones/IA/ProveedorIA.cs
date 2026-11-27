namespace ApiDesposte.Funciones.IA
{
    /// <summary>
    /// Proveedor de IA que recibe un prompt de sistema + un mensaje y devuelve texto JSON.
    /// Implementaciones: <see cref="ProveedorDeepSeek"/> y <see cref="ProveedorGemini"/>.
    /// </summary>
    public interface IProveedorIA
    {
        /// <summary>Nombre corto para mensajes y para la caché (ej. "DeepSeek").</summary>
        string Nombre { get; }

        /// <summary>true si tiene clave de API configurada.</summary>
        bool Configurado { get; }

        /// <summary>
        /// Genera la respuesta JSON. Reintenta internamente ante saturación dentro de <paramref name="tiempoDisponible"/>.
        /// Lanza <see cref="IASaturadaException"/> si el servicio sigue saturado al agotar los reintentos.
        /// </summary>
        Task<RespuestaIA> GenerarJsonAsync(string promptSistema, string mensajeUsuario, TimeSpan tiempoDisponible, CancellationToken ct);
    }

    /// <param name="Texto">Contenido devuelto por el modelo (se espera un objeto JSON).</param>
    /// <param name="Modelo">Modelo que respondió (ej. deepseek-chat, gemini-3-flash).</param>
    public sealed record RespuestaIA(string Texto, string Modelo);

    /// <summary>Falla temporal (429, 5xx, timeout, red): se puede reintentar.</summary>
    public sealed class IATransitoriaException(string mensaje, TimeSpan? reintentarEn = null) : Exception(mensaje)
    {
        public TimeSpan? ReintentarEn { get; } = reintentarEn;
    }

    /// <summary>El servicio siguió saturado tras los reintentos.</summary>
    public sealed class IASaturadaException(string mensaje) : Exception(mensaje);

    /// <summary>Utilidades comunes de los proveedores.</summary>
    internal static class UtilIA
    {
        /// <summary>Espera entre reintentos: la indicada por el servicio (Retry-After, máx. 30 s) o 2 s, 8 s…</summary>
        public static TimeSpan Espera(IATransitoriaException ex, int intento) =>
            ex.ReintentarEn ?? TimeSpan.FromSeconds(2 * intento * intento);

        public static TimeSpan? RetryAfter(HttpResponseMessage respuesta)
        {
            TimeSpan? valor = respuesta.Headers.RetryAfter?.Delta;
            return valor > TimeSpan.FromSeconds(30) ? TimeSpan.FromSeconds(30) : valor;
        }

        public static string Recortar(string texto) => texto.Length > 400 ? texto[..400] : texto;
    }
}
