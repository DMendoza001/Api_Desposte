using System.Text.Json;
using ApiDesposte.Funciones.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.ModelBinding;

namespace ApiDesposte.Controllers.Excel
{
    /// <summary>
    /// CRUD de las tablas del libro "Informe Diario 2026 - Base de Datos.xlsx".
    ///   GET    /api/InformeDiarioCrud                  -> catálogo de tablas editables
    ///   GET    /api/InformeDiarioCrud/{tabla}          -> columnas + registros (cada uno con "_fila")
    ///   POST   /api/InformeDiarioCrud/{tabla}          -> crea: body { "FECHA": "2026-11-27", "AREA": "DESPOSTE", ... }
    ///   PUT    /api/InformeDiarioCrud/{tabla}/{fila}   -> actualiza: body { "valores": {...}, "original": {...} }
    ///   DELETE /api/InformeDiarioCrud/{tabla}/{fila}   -> elimina: body opcional { "original": {...} }
    ///   POST   /api/InformeDiarioCrud/{tabla}/lote     -> varias operaciones en una sola escritura (todo o nada):
    ///          { "crear": [{...}], "actualizar": [{ "fila", "valores", "original" }], "eliminar": [{ "fila", "original" }] }
    /// "original" son los valores que el cliente leyó; si la fila cambió entretanto se responde 409.
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    public class InformeDiarioCrudController : ControllerBase
    {
        private static readonly Dictionary<string, (string Hoja, string Tabla)> MapeoTablas = new(StringComparer.OrdinalIgnoreCase)
        {
            { "personal",   ("BD_PERSONAL",   "TablaPersonal") },
            { "sso",        ("BD_SSO",        "TablaSSO") },
            { "horarios",   ("BD_HORARIOS",   "TablaHorarios") },
            { "produccion", ("BD_PRODUCCION", "TablaProduccion") },
            { "detalle",    ("BD_DETALLE",    "TablaDetalle") },
            { "paradas",    ("BD_PARADAS",    "TablaParadas") },
            { "areas",      ("CATALOGOS",     "TablaAreas") },
            { "turnos",     ("CATALOGOS",     "TablaTurnos") },
            { "tipos",      ("CATALOGOS",     "TablaTipos") },
            { "secciones",  ("CATALOGOS",     "TablaSecciones") },
            { "categorias", ("CATALOGOS",     "TablaCategorias") },
            { "conceptos",  ("CATALOGOS",     "TablaConceptos") }
        };

        private readonly string _carpetaRespaldos;

        public InformeDiarioCrudController(IWebHostEnvironment env)
        {
            _carpetaRespaldos = Path.Combine(env.ContentRootPath, "Datos", "RespaldosExcel");
        }

        private string ObtenerRutaExcel()
        {
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return Path.Combine(userProfile, @"OneDrive - Corporación Rico SAC\Desposte-04\Informes Diario\Informe Diario 2026 - Base de Datos.xlsx");
        }

        [HttpGet]
        public IActionResult ListarTablas() =>
            Ok(new
            {
                exito = true,
                tablas = MapeoTablas.Select(m => new { clave = m.Key, hoja = m.Value.Hoja, tabla = m.Value.Tabla })
            });

        [HttpGet("{tabla}")]
        public IActionResult Listar(string tabla) => Ejecutar(tabla, t =>
        {
            var datos = EditorTablasExcel.Leer(ObtenerRutaExcel(), t.Hoja, t.Tabla);
            return Ok(new { exito = true, tabla = t.Tabla, columnas = datos.Columnas, totalRegistros = datos.Registros.Count, registros = datos.Registros });
        });

        [HttpPost("{tabla}")]
        public IActionResult Crear(string tabla, [FromBody] Dictionary<string, JsonElement> valores) => Ejecutar(tabla, t =>
        {
            int fila = EditorTablasExcel.Insertar(ObtenerRutaExcel(), _carpetaRespaldos, t.Hoja, t.Tabla, valores);
            return Ok(new { exito = true, mensaje = "Registro creado.", fila });
        });

        [HttpPost("{tabla}/lote")]
        public IActionResult GuardarLote(string tabla, [FromBody] SolicitudLote solicitud) => Ejecutar(tabla, t =>
        {
            var r = EditorTablasExcel.Lote(ObtenerRutaExcel(), _carpetaRespaldos, t.Hoja, t.Tabla,
                solicitud.Crear,
                solicitud.Actualizar.Select(a => new OperacionActualizar(a.Fila, a.Valores, a.Original)).ToList(),
                solicitud.Eliminar.Select(e => new OperacionEliminar(e.Fila, e.Original)).ToList());
            return Ok(new { exito = true, mensaje = "Cambios guardados.", actualizadas = r.Actualizadas, eliminadas = r.Eliminadas, creadas = r.Creadas });
        });

        [HttpPut("{tabla}/{fila:int}")]
        public IActionResult Actualizar(string tabla, int fila, [FromBody] SolicitudActualizar solicitud) => Ejecutar(tabla, t =>
        {
            var cambiadas = EditorTablasExcel.Actualizar(ObtenerRutaExcel(), _carpetaRespaldos, t.Hoja, t.Tabla, fila, solicitud.Valores, solicitud.Original);
            return Ok(new { exito = true, mensaje = cambiadas.Count == 0 ? "Sin cambios." : "Registro actualizado.", fila, columnasCambiadas = cambiadas });
        });

        [HttpDelete("{tabla}/{fila:int}")]
        public IActionResult Eliminar(string tabla, int fila,
            [FromBody(EmptyBodyBehavior = EmptyBodyBehavior.Allow)] SolicitudEliminar? solicitud) => Ejecutar(tabla, t =>
        {
            EditorTablasExcel.Eliminar(ObtenerRutaExcel(), _carpetaRespaldos, t.Hoja, t.Tabla, fila, solicitud?.Original);
            return Ok(new { exito = true, mensaje = "Registro eliminado.", fila });
        });

        private IActionResult Ejecutar(string clave, Func<(string Hoja, string Tabla), IActionResult> accion)
        {
            if (!MapeoTablas.TryGetValue(clave, out var t))
                return NotFound(new { exito = false, error = $"Tabla '{clave}' no reconocida. Válidas: {string.Join(", ", MapeoTablas.Keys)}." });
            try
            {
                return accion(t);
            }
            catch (ExcepcionTablaExcel ex)
            {
                return StatusCode(ex.Codigo, new { exito = false, error = ex.Message });
            }
            catch (FileNotFoundException ex)
            {
                return NotFound(new { exito = false, error = ex.Message });
            }
            catch (IOException)
            {
                return StatusCode(409, new { exito = false, error = "El archivo Excel está bloqueado (abierto en Excel de escritorio). Ciérralo e intenta de nuevo." });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }
    }

    public class SolicitudActualizar
    {
        public Dictionary<string, JsonElement> Valores { get; set; } = new();
        public Dictionary<string, JsonElement>? Original { get; set; }
    }

    public class SolicitudEliminar
    {
        public Dictionary<string, JsonElement>? Original { get; set; }
    }

    public class SolicitudLote
    {
        public List<Dictionary<string, JsonElement>> Crear { get; set; } = new();
        public List<SolicitudLoteFila> Actualizar { get; set; } = new();
        public List<SolicitudLoteFila> Eliminar { get; set; } = new();
    }

    public class SolicitudLoteFila
    {
        public int Fila { get; set; }
        public Dictionary<string, JsonElement> Valores { get; set; } = new();
        public Dictionary<string, JsonElement>? Original { get; set; }
    }
}
