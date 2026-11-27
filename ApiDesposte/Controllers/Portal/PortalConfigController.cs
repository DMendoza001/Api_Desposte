using ApiDesposte.Funciones.Portal;
using Microsoft.AspNetCore.Mvc;
using System.Text.Json;

namespace ApiDesposte.Controllers.Portal
{
    [ApiController]
    [Route("api/[controller]")]
    public class PortalConfigController : ControllerBase
    {
        private readonly string _adminPin;
        private readonly string _archivoConfig;
        private static readonly object _lock = new();

        public PortalConfigController(IConfiguration configuration, IWebHostEnvironment env)
        {
            // PIN desde la variable de entorno AdminConfig__Pin (sin valor por defecto en el código)
            _adminPin = (configuration["AdminConfig:Pin"] ?? "").Trim();

            // Datos de ejecución fuera de la carpeta web: ApiDesposte/Datos/dashboards_estado.json
            _archivoConfig = EstadoDashboards.RutaArchivo(env.ContentRootPath);
        }

        // Sin PIN configurado en el servidor no se concede acceso de administrador
        private bool PinValido(string? pin) =>
            _adminPin.Length > 0 && (pin ?? "").Trim() == _adminPin;

        /// <summary>
        /// Obtiene el estado actual de visibilidad de los dashboards (habilitado / en prueba)
        /// </summary>
        [HttpGet("config")]
        public IActionResult ObtenerConfiguracion()
        {
            lock (_lock)
            {
                try
                {
                    if (System.IO.File.Exists(_archivoConfig))
                    {
                        string json = System.IO.File.ReadAllText(_archivoConfig);
                        var datos = JsonSerializer.Deserialize<Dictionary<string, bool>>(json) ?? new Dictionary<string, bool>();
                        return Ok(new { exito = true, estados = datos });
                    }
                }
                catch (Exception ex)
                {
                    return StatusCode(500, new { exito = false, mensaje = "Error leyendo configuración", detalle = ex.Message });
                }

                return Ok(new { exito = true, estados = new Dictionary<string, bool>() });
            }
        }

        /// <summary>
        /// Valida el PIN de Administrador
        /// </summary>
        [HttpPost("verificar-pin")]
        public IActionResult VerificarPin([FromBody] PinRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Pin))
            {
                return BadRequest(new { exito = false, mensaje = "El PIN es requerido." });
            }

            if (PinValido(request.Pin))
            {
                return Ok(new { exito = true, mensaje = "PIN válido. Acceso concedido." });
            }

            return Unauthorized(new { exito = false, mensaje = "PIN incorrecto." });
        }

        /// <summary>
        /// Guarda el estado de visibilidad de los dashboards (requiere PIN)
        /// </summary>
        [HttpPost("config")]
        public IActionResult GuardarConfiguracion([FromBody] GuardarConfigRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Pin))
            {
                return BadRequest(new { exito = false, mensaje = "El PIN de administrador es requerido." });
            }

            if (!PinValido(request.Pin))
            {
                return Unauthorized(new { exito = false, mensaje = "PIN incorrecto. No autorizado." });
            }

            lock (_lock)
            {
                try
                {
                    Dictionary<string, bool> estadosActuales = new(StringComparer.OrdinalIgnoreCase);

                    if (System.IO.File.Exists(_archivoConfig))
                    {
                        try
                        {
                            string jsonExistente = System.IO.File.ReadAllText(_archivoConfig);
                            var deserializado = JsonSerializer.Deserialize<Dictionary<string, bool>>(jsonExistente);
                            if (deserializado != null)
                            {
                                foreach (var kvp in deserializado)
                                {
                                    estadosActuales[kvp.Key] = kvp.Value;
                                }
                            }
                        }
                        catch
                        {
                            estadosActuales = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
                        }
                    }

                    // Actualización unitaria si viene id y activo
                    if (!string.IsNullOrWhiteSpace(request.Id) && request.Activo.HasValue)
                    {
                        estadosActuales[request.Id.Trim()] = request.Activo.Value;
                    }

                    // Actualización masiva si viene diccionario de estados
                    if (request.Estados != null)
                    {
                        foreach (var kvp in request.Estados)
                        {
                            estadosActuales[kvp.Key.Trim()] = kvp.Value;
                        }
                    }

                    var opciones = new JsonSerializerOptions { WriteIndented = true };
                    string nuevoJson = JsonSerializer.Serialize(estadosActuales, opciones);
                    Directory.CreateDirectory(Path.GetDirectoryName(_archivoConfig)!);
                    System.IO.File.WriteAllText(_archivoConfig, nuevoJson);

                    return Ok(new { exito = true, mensaje = "Configuración guardada correctamente.", estados = estadosActuales });
                }
                catch (Exception ex)
                {
                    return StatusCode(500, new { exito = false, mensaje = "Error al guardar configuración.", detalle = ex.Message });
                }
            }
        }
    }

    public class PinRequest
    {
        public string Pin { get; set; } = string.Empty;
    }

    public class GuardarConfigRequest
    {
        public string Pin { get; set; } = string.Empty;
        public string? Id { get; set; }
        public bool? Activo { get; set; }
        public Dictionary<string, bool>? Estados { get; set; }
    }
}
