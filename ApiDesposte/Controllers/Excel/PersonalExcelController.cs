using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System.IO;

namespace ApiDesposte.Controllers.Excel
{
    [ApiController]
    [Route("api/[controller]")]
    public class PersonalExcelController : ControllerBase
    {
        // Ruta al archivo sincronizado localmente por OneDrive / SharePoint
        private readonly string _rutaExcel = @"C:\Users\Dennis\OneDrive - Corporación Rico SAC\Desposte-03\Desarrollos\Informes\AtencionesTopico001.xlsx";
        private readonly string _nombreHoja = "Prueba";
        private readonly string _nombreTabla = "T_Personal";

        // 1. LECTURA: Listar todo el personal
        [HttpGet("listar")]
        public IActionResult ListarTodo()
        {
            try
            {
                var lista = new List<object>();

                // Abrir en modo lectura compartida para no chocar con bloqueos
                using (var stream = new FileStream(_rutaExcel, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var workbook = new XLWorkbook(stream))
                {
                    var hoja = workbook.Worksheet(_nombreHoja);
                    var tabla = hoja.Table(_nombreTabla);

                    foreach (var fila in tabla.DataRange.Rows())
                    {
                        lista.Add(new
                        {
                            Dni = fila.Cell(1).Value.ToString().Trim(),
                            NombresApellidos = fila.Cell(2).Value.ToString().Trim(),
                            Telefonos = fila.Cell(3).Value.ToString().Trim()
                        });
                    }
                }
                return Ok(lista);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }

        // 2. BUSCAR POR DNI
        [HttpGet("buscar/{dni}")]
        public IActionResult BuscarPorDni(string dni)
        {
            try
            {
                using (var stream = new FileStream(_rutaExcel, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var workbook = new XLWorkbook(stream))
                {
                    var hoja = workbook.Worksheet(_nombreHoja);
                    var tabla = hoja.Table(_nombreTabla);

                    var fila = tabla.DataRange.Rows()
                        .FirstOrDefault(r => r.Cell(1).Value.ToString().Trim() == dni.Trim());

                    if (fila == null)
                        return NotFound(new { exito = false, mensaje = "DNI no encontrado." });

                    return Ok(new
                    {
                        Dni = fila.Cell(1).Value.ToString().Trim(),
                        NombresApellidos = fila.Cell(2).Value.ToString().Trim(),
                        Telefonos = fila.Cell(3).Value.ToString().Trim()
                    });
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }

        // 3. CREAR REGISTRO Y FORZAR SINCRONIZACIÓN WEB
        [HttpPost("crear")]
        public IActionResult Crear([FromBody] PersonaDto persona)
        {
            try
            {
                using (var stream = new FileStream(_rutaExcel, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
                using (var workbook = new XLWorkbook(stream))
                {
                    var hoja = workbook.Worksheet(_nombreHoja);
                    var tabla = hoja.Table(_nombreTabla);

                    // Validar si el DNI ya existe
                    bool existe = tabla.DataRange.Rows()
                        .Any(r => r.Cell(1).Value.ToString().Trim() == persona.Dni.Trim());

                    if (existe)
                        return BadRequest(new { exito = false, mensaje = "El DNI ya se encuentra registrado." });

                    // Insertar nueva fila
                    tabla.InsertRowsBelow(1);
                    var nuevaFila = tabla.DataRange.LastRow();

                    nuevaFila.Cell(1).Value = persona.Dni.Trim();
                    nuevaFila.Cell(2).Value = persona.NombresApellidos.Trim();
                    nuevaFila.Cell(3).Value = persona.Telefonos.Trim();

                    // Guardar sobre el stream local
                    stream.Position = 0;
                    workbook.SaveAs(stream);
                    stream.SetLength(stream.Position);
                }

                // FORZAR A ONEDRIVE / SHAREPOINT A DETECTAR EL CAMBIO DE INMEDIATO
                System.IO.File.SetLastWriteTime(_rutaExcel, DateTime.Now);

                return Ok(new { exito = true, mensaje = "Persona registrada correctamente y sincronizando con SharePoint." });
            }
            catch (IOException)
            {
                return StatusCode(409, new { exito = false, mensaje = "El archivo Excel está bloqueado por la aplicación de escritorio. Por favor guarda/cierra el archivo e intenta nuevamente." });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }

        // 4. ACTUALIZAR REGISTRO Y FORZAR SINCRONIZACIÓN WEB
        [HttpPut("actualizar")]
        public IActionResult Actualizar([FromBody] PersonaDto persona)
        {
            try
            {
                using (var stream = new FileStream(_rutaExcel, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
                using (var workbook = new XLWorkbook(stream))
                {
                    var hoja = workbook.Worksheet(_nombreHoja);
                    var tabla = hoja.Table(_nombreTabla);

                    var fila = tabla.DataRange.Rows()
                        .FirstOrDefault(r => r.Cell(1).Value.ToString().Trim() == persona.Dni.Trim());

                    if (fila == null)
                        return NotFound(new { exito = false, mensaje = "DNI no encontrado para actualizar." });

                    fila.Cell(2).Value = persona.NombresApellidos.Trim();
                    fila.Cell(3).Value = persona.Telefonos.Trim();

                    stream.Position = 0;
                    workbook.SaveAs(stream);
                    stream.SetLength(stream.Position);
                }

                // FORZAR A ONEDRIVE / SHAREPOINT A DETECTAR EL CAMBIO DE INMEDIATO
                System.IO.File.SetLastWriteTime(_rutaExcel, DateTime.Now);

                return Ok(new { exito = true, mensaje = "Datos actualizados correctamente y sincronizando con SharePoint." });
            }
            catch (IOException)
            {
                return StatusCode(409, new { exito = false, mensaje = "El archivo Excel está bloqueado por la aplicación de escritorio. Por favor guarda/cierra el archivo e intenta nuevamente." });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }

        // 5. ELIMINAR REGISTRO Y FORZAR SINCRONIZACIÓN WEB
        [HttpDelete("eliminar/{dni}")]
        public IActionResult Eliminar(string dni)
        {
            try
            {
                using (var stream = new FileStream(_rutaExcel, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
                using (var workbook = new XLWorkbook(stream))
                {
                    var hoja = workbook.Worksheet(_nombreHoja);
                    var tabla = hoja.Table(_nombreTabla);

                    var fila = tabla.DataRange.Rows()
                        .FirstOrDefault(r => r.Cell(1).Value.ToString().Trim() == dni.Trim());

                    if (fila == null)
                        return NotFound(new { exito = false, mensaje = "DNI no encontrado." });

                    fila.Delete();

                    stream.Position = 0;
                    workbook.SaveAs(stream);
                    stream.SetLength(stream.Position);
                }

                // FORZAR A ONEDRIVE / SHAREPOINT A DETECTAR EL CAMBIO DE INMEDIATO
                System.IO.File.SetLastWriteTime(_rutaExcel, DateTime.Now);

                return Ok(new { exito = true, mensaje = "Registro eliminado correctamente y sincronizando con SharePoint." });
            }
            catch (IOException)
            {
                return StatusCode(409, new { exito = false, mensaje = "El archivo Excel está bloqueado por la aplicación de escritorio. Por favor guarda/cierra el archivo e intenta nuevamente." });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }
    }

    public class PersonaDto
    {
        public string Dni { get; set; } = string.Empty;
        public string NombresApellidos { get; set; } = string.Empty;
        public string Telefonos { get; set; } = string.Empty;
    }
}