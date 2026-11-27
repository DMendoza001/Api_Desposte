using System.Text.Json;

namespace ApiDesposte.Funciones.Portal
{
    /// <summary>
    /// Estado de visibilidad de los dashboards (true = visible, false = en prueba: solo administrador).
    /// Se guarda en ApiDesposte/Datos/dashboards_estado.json, fuera de la carpeta web,
    /// y se consulta desde la API (GET /api/PortalConfig/config).
    /// </summary>
    public static class EstadoDashboards
    {
        public static string RutaArchivo(string contentRootPath) =>
            Path.Combine(contentRootPath, "Datos", "dashboards_estado.json");

        /// <summary>Lee el archivo; si no existe devuelve un diccionario vacío. Claves sin distinguir mayúsculas.</summary>
        public static Dictionary<string, bool> Leer(string rutaArchivo)
        {
            var estados = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            if (!File.Exists(rutaArchivo)) return estados;

            var datos = JsonSerializer.Deserialize<Dictionary<string, bool>>(File.ReadAllText(rutaArchivo));
            if (datos != null)
            {
                foreach (var kvp in datos) estados[kvp.Key] = kvp.Value;
            }
            return estados;
        }
    }
}
