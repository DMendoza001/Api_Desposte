using System.Globalization;

namespace ApiDesposte.Funciones.Fechas
{
    /// <summary>
    /// Lectura de fechas recibidas por query string para los SP de SQL Server.
    /// Los SP reciben Fecha1 = "yyyy-MM-ddT00:00:00" y Fecha2 = "yyyy-MM-ddT23:59:59".
    /// </summary>
    public static class ParametrosFecha
    {
        private static readonly string[] FormatosEstrictos = { "yyyy-MM-dd", "yyyy-MM-ddTHH:mm:ss", "dd/MM/yyyy" };
        private static readonly string[] FormatosFlexibles = { "yyyy-MM-dd", "dd/MM/yyyy", "dd-MM-yyyy", "yyyy/MM/dd", "M/d/yyyy", "d/M/yyyy" };
        private static readonly CultureInfo CulturaPeru = new("es-PE");

        /// <summary>
        /// Validación estricta: acepta yyyy-MM-dd, yyyy-MM-ddTHH:mm:ss y dd/MM/yyyy.
        /// Vacío → valorPorDefecto. Devuelve false si el formato es inválido (el controlador responde 400).
        /// </summary>
        public static bool TryParsear(string? input, DateTime valorPorDefecto, out DateTime fecha)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                fecha = valorPorDefecto;
                return true;
            }

            return DateTime.TryParseExact(input.Trim(), FormatosEstrictos, CultureInfo.InvariantCulture, DateTimeStyles.None, out fecha);
        }

        /// <summary>
        /// Conversión flexible al texto que espera el SP, sin rechazar la petición:
        /// - Vacío → valorPorDefecto con la hora de inicio (T00:00:00) o de fin (T23:59:59).
        /// - Si ya viene completo (yyyy-MM-ddTHH:mm:ss) se respeta tal cual.
        /// - Acepta formatos comunes (yyyy-MM-dd, dd/MM/yyyy, dd-MM-yyyy, yyyy/MM/dd, M/d/yyyy) y la cultura es-PE.
        /// - Si no se puede interpretar, se envía el texto al SP para que SQL Server lo valide.
        /// </summary>
        public static string FormatearParaSp(string? input, DateTime valorPorDefecto, bool esFechaFinal)
        {
            string sufijoHora = esFechaFinal ? "T23:59:59" : "T00:00:00";

            if (string.IsNullOrWhiteSpace(input))
            {
                return $"{valorPorDefecto:yyyy-MM-dd}{sufijoHora}";
            }

            string limpio = input.Trim();

            if (limpio.Contains('T') && limpio.Length >= 19)
            {
                return limpio;
            }

            if (DateTime.TryParseExact(limpio, FormatosFlexibles, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime fecha) ||
                DateTime.TryParse(limpio, CulturaPeru, DateTimeStyles.None, out fecha) ||
                DateTime.TryParse(limpio, CultureInfo.InvariantCulture, DateTimeStyles.None, out fecha))
            {
                return $"{fecha:yyyy-MM-dd}{sufijoHora}";
            }

            return limpio.Contains('T') || limpio.Contains(' ') ? limpio : $"{limpio}{sufijoHora}";
        }

        /// <summary>
        /// Devuelve el primer valor no vacío: permite aceptar alias de parámetros
        /// (ej. fechaInicio / fecha1) sin repetir la misma comprobación en cada controlador.
        /// </summary>
        public static string? Primero(params string?[] valores) =>
            valores.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));
    }
}
