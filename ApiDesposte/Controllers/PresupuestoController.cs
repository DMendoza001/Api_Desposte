using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System.IO;

namespace ApiDesposte.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class PresupuestoController : ControllerBase
    {
        private readonly string _nombreHoja = "Cabecera";
        private readonly string _nombreTabla = "T_Cabecera";

        private string ObtenerRutaExcel()
        {
            // Detecta automáticamente la carpeta del usuario actual (ej. C:\Users\dennis.mendoza o C:\Users\Dennis)
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return Path.Combine(userProfile, @"OneDrive - Corporación Rico SAC\Desposte-03\Presupuestos\2027\Presupuesto_2027.xlsm");
        }

        [HttpGet("dashboard-data")]
        public IActionResult ObtenerDatosDashboard()
        {
            try
            {
                string rutaExcel = ObtenerRutaExcel();

                if (!System.IO.File.Exists(rutaExcel))
                {
                    return NotFound(new { error = $"No se encontró el archivo en la ruta: {rutaExcel}" });
                }

                var resultado = new List<Dictionary<string, object>>();

                // Copiar el archivo a memoria primero para leerlo aunque esté abierto en Excel
                byte[] fileBytes;
                using (var fileStream = new FileStream(rutaExcel, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                using (var ms = new MemoryStream())
                {
                    fileStream.CopyTo(ms);
                    fileBytes = ms.ToArray();
                }

                // Cargar el Excel desde la copia en memoria
                using (var memoryStream = new MemoryStream(fileBytes))
                using (var workbook = new XLWorkbook(memoryStream))
                {
                    var hoja = workbook.Worksheet(_nombreHoja);
                    var tabla = hoja.Table(_nombreTabla);

                    // Extraer nombres de cabeceras dinámicamente
                    var encabezados = tabla.HeadersRow().Cells()
                        .Select(c => c.Value.ToString().Trim())
                        .ToList();

                    // Recorrer filas
                    foreach (var fila in tabla.DataRange.Rows())
                    {
                        var filaDiccionario = new Dictionary<string, object>();
                        for (int i = 0; i < encabezados.Count; i++)
                        {
                            var celda = fila.Cell(i + 1);
                            
                            if (celda.DataType == XLDataType.Number)
                            {
                                filaDiccionario[encabezados[i]] = celda.GetValue<double>();
                            }
                            else
                            {
                                filaDiccionario[encabezados[i]] = celda.Value.ToString().Trim();
                            }
                        }
                        resultado.Add(filaDiccionario);
                    }
                }

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message });
            }
        }
    }
}