using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System.Globalization;
using System.IO;

namespace ApiDesposte.Controllers.Excel
{
    [ApiController]
    [Route("api/[controller]")]
    public class CumplimientoPtcDiarioController : ControllerBase
    {
        // Ruta relativa del archivo dentro de OneDrive (el perfil del usuario se resuelve dinámicamente)
        private const string RutaRelativaExcel = @"OneDrive - Corporación Rico SAC\Desposte-03\PlantasCore\Planificacion\Cumplimiento_PTC_Diario.xlsm";

        // Mapeo oficial de Hojas y Tablas en Cumplimiento_PTC_Diario.xlsm
        private static readonly Dictionary<string, (string Hoja, string Tabla)> MapeoTablas = new(StringComparer.OrdinalIgnoreCase)
        {
            { "T_Produccion", ("Produccion", "T_Produccion") },
            { "T_ResProd_Py_DB", ("ResProd_Py_DB", "T_ResProd_Py_DB") },
            { "T_Distribucion", ("Distribucion", "T_Distribucion") },
            { "T_CodigoRelacion", ("Auxiliares", "T_CodigoRelacion") }
        };

        // Nombres oficiales de campo por concepto (Llave: Fecha + CodigoCorto | Cálculo: Kilos | Auxiliares: SubPlanta)
        private static readonly string[] CamposFecha = { "Fecha" };
        private static readonly string[] CamposCodigoCorto = { "CodigoCorto", "CódigoCorto", "Codigo", "Código" };
        private static readonly string[] CamposKilos = { "Kilos", "Kgs", "Peso", "KilosNetos" };
        private static readonly string[] CamposNombre = { "Nombre", "Descripcion", "Descripción" };
        private static readonly string[] CamposSubPlanta = { "SubPlanta", "Sub Planta", "Planta" };

        // Cache en memoria para optimizar lecturas consecutivas desde el frontend JS
        private static Dictionary<string, List<Dictionary<string, object?>>>? _cacheTablas = null;
        private static DateTime _cacheTimestamp = DateTime.MinValue;
        private static readonly object _cacheLock = new();
        private static readonly TimeSpan CacheDuracion = TimeSpan.FromSeconds(60);

        private string ObtenerRutaExcel()
        {
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return Path.Combine(userProfile, RutaRelativaExcel);
        }

        private XLWorkbook CargarWorkbookEnMemoria(string rutaExcel)
        {
            byte[] fileBytes;
            using (var fileStream = new FileStream(rutaExcel, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var ms = new MemoryStream())
            {
                fileStream.CopyTo(ms);
                fileBytes = ms.ToArray();
            }

            var msModificado = new MemoryStream();
            msModificado.Write(fileBytes, 0, fileBytes.Length);
            msModificado.Position = 0;

            try
            {
                // Limpieza de PivotTables para evitar conflictos al abrir con ClosedXML
                using (var doc = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(msModificado, true))
                {
                    if (doc.WorkbookPart != null)
                    {
                        foreach (var wsPart in doc.WorkbookPart.WorksheetParts)
                        {
                            var pParts = wsPart.PivotTableParts.ToList();
                            foreach (var p in pParts)
                            {
                                wsPart.DeletePart(p);
                            }
                        }

                        var pivotCaches = doc.WorkbookPart.PivotTableCacheDefinitionParts.ToList();
                        foreach (var cache in pivotCaches)
                        {
                            doc.WorkbookPart.DeletePart(cache);
                        }

                        if (doc.WorkbookPart.Workbook != null)
                        {
                            doc.WorkbookPart.Workbook.PivotCaches?.Remove();
                            doc.WorkbookPart.Workbook.Save();
                        }
                    }
                }
            }
            catch
            {
                msModificado = new MemoryStream(fileBytes);
            }

            msModificado.Position = 0;
            return new XLWorkbook(msModificado);
        }

        /// <summary>
        /// Lee una tabla estructurada de Excel tal cual está definida en la hoja,
        /// conservando exactamente sus encabezados y valores originales sin agregar campos calculados.
        /// </summary>
        private static List<Dictionary<string, object?>> LeerTablaExcelTalCual(XLWorkbook workbook, string nombreHoja, string nombreTabla)
        {
            var registros = new List<Dictionary<string, object?>>();

            if (!workbook.Worksheets.TryGetWorksheet(nombreHoja, out var hoja))
            {
                hoja = workbook.Worksheets.FirstOrDefault(w => string.Equals(w.Name.Trim(), nombreHoja.Trim(), StringComparison.OrdinalIgnoreCase));
            }

            IXLTable? tabla = null;

            if (hoja != null)
            {
                try
                {
                    tabla = hoja.Table(nombreTabla);
                }
                catch
                {
                    tabla = hoja.Tables.FirstOrDefault(t => string.Equals(t.Name.Trim(), nombreTabla.Trim(), StringComparison.OrdinalIgnoreCase));
                }
            }

            if (tabla == null)
            {
                tabla = workbook.Worksheets.SelectMany(w => w.Tables)
                    .FirstOrDefault(t => string.Equals(t.Name.Trim(), nombreTabla.Trim(), StringComparison.OrdinalIgnoreCase));
            }

            if (tabla == null)
            {
                throw new InvalidOperationException($"No se encontró la tabla '{nombreTabla}' en la hoja '{nombreHoja}'.");
            }

            var headers = tabla.HeadersRow().Cells().Select(c => c.Value.ToString().Trim()).ToList();

            foreach (var fila in tabla.DataRange.Rows())
            {
                if (fila.IsEmpty()) continue;

                var filaDict = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                bool tieneDatos = false;

                for (int i = 0; i < headers.Count; i++)
                {
                    string columna = string.IsNullOrWhiteSpace(headers[i]) ? $"Columna_{i + 1}" : headers[i];
                    object? valor = ObtenerValorCelda(fila.Cell(i + 1));
                    if (valor != null) tieneDatos = true;
                    filaDict[columna] = valor;
                }

                if (tieneDatos)
                {
                    registros.Add(filaDict);
                }
            }

            return registros;
        }

        private static object? ObtenerValorCelda(IXLCell celda)
        {
            if (celda.IsEmpty()) return null;

            try
            {
                switch (celda.DataType)
                {
                    case XLDataType.Boolean:
                        return celda.GetBoolean();

                    case XLDataType.Number:
                        double num = celda.GetDouble();
                        if (Math.Abs(num % 1) < double.Epsilon && num >= long.MinValue && num <= long.MaxValue)
                            return (long)num;
                        return Math.Round(num, 4);

                    case XLDataType.DateTime:
                        var dt = celda.GetDateTime();
                        return dt.TimeOfDay.TotalSeconds == 0
                            ? dt.ToString("yyyy-MM-dd")
                            : dt.ToString("yyyy-MM-dd HH:mm:ss");

                    case XLDataType.TimeSpan:
                        var ts = celda.GetTimeSpan();
                        return ts.ToString(@"hh\:mm\:ss");

                    default:
                        var str = celda.Value.ToString().Trim();
                        return string.IsNullOrEmpty(str) ? null : str;
                }
            }
            catch
            {
                var fallback = celda.Value.ToString().Trim();
                return string.IsNullOrEmpty(fallback) ? null : fallback;
            }
        }

        private Dictionary<string, List<Dictionary<string, object?>>> CargarTodasLasTablas(bool forzarRecarga = false)
        {
            lock (_cacheLock)
            {
                if (!forzarRecarga && _cacheTablas != null && (DateTime.UtcNow - _cacheTimestamp) < CacheDuracion)
                {
                    return _cacheTablas;
                }

                string rutaExcel = ObtenerRutaExcel();
                if (!System.IO.File.Exists(rutaExcel))
                {
                    throw new FileNotFoundException($"No se encontró el archivo Excel en la ruta: {rutaExcel}");
                }

                using var workbook = CargarWorkbookEnMemoria(rutaExcel);

                var resultado = new Dictionary<string, List<Dictionary<string, object?>>>(StringComparer.OrdinalIgnoreCase);

                foreach (var item in MapeoTablas)
                {
                    resultado[item.Key] = LeerTablaExcelTalCual(workbook, item.Value.Hoja, item.Value.Tabla);
                }

                _cacheTablas = resultado;
                _cacheTimestamp = DateTime.UtcNow;

                return _cacheTablas;
            }
        }

        // =====================================================================
        // 🔧 UTILIDADES DE NORMALIZACIÓN (Llave: Fecha + CodigoCorto | Cálculo: Kilos)
        // =====================================================================

        private static string ObtenerTexto(Dictionary<string, object?> fila, params string[] campos)
        {
            foreach (var campo in campos)
            {
                if (!fila.TryGetValue(campo, out var valor) || valor == null) continue;

                string texto = valor.ToString()?.Trim() ?? "";
                if (!string.IsNullOrWhiteSpace(texto)) return texto;
            }
            return "";
        }

        private static double ObtenerNumero(Dictionary<string, object?> fila, params string[] campos)
        {
            foreach (var campo in campos)
            {
                if (!fila.TryGetValue(campo, out var valor) || valor == null) continue;

                switch (valor)
                {
                    case double d: return d;
                    case long l: return l;
                    case int i: return i;
                    case decimal m: return (double)m;
                    case float f: return f;
                }

                string texto = (valor.ToString() ?? "").Trim().Replace(",", ".");
                if (double.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out double numero))
                {
                    return numero;
                }
            }
            return 0d;
        }

        private static DateTime? ObtenerFecha(Dictionary<string, object?> fila, params string[] campos)
        {
            foreach (var campo in campos)
            {
                if (!fila.TryGetValue(campo, out var valor) || valor == null) continue;

                DateTime? fecha = ConvertirFecha(valor);
                if (fecha.HasValue) return fecha;
            }
            return null;
        }

        private static DateTime? ConvertirFecha(object valor)
        {
            try
            {
                if (valor is DateTime dt) return dt.Date;

                if (valor is double d) return DateTime.FromOADate(d).Date;
                if (valor is long l) return DateTime.FromOADate(l).Date;
                if (valor is int i) return DateTime.FromOADate(i).Date;
            }
            catch
            {
                return null;
            }

            string texto = (valor.ToString() ?? "").Trim();
            if (string.IsNullOrWhiteSpace(texto)) return null;

            // Serial numérico de Excel (ej. 46000)
            if (double.TryParse(texto.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double serial)
                && serial > 20000 && serial < 90000)
            {
                try { return DateTime.FromOADate(serial).Date; }
                catch { /* continuar con los formatos de texto */ }
            }

            string[] formatos = { "yyyy-MM-dd", "yyyy-MM-dd HH:mm:ss", "yyyy/MM/dd", "dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "d-M-yyyy" };
            if (DateTime.TryParseExact(texto, formatos, CultureInfo.InvariantCulture, DateTimeStyles.None, out var exacta))
            {
                return exacta.Date;
            }

            if (DateTime.TryParse(texto, CultureInfo.GetCultureInfo("es-PE"), DateTimeStyles.None, out var fecha))
            {
                return fecha.Date;
            }

            if (DateTime.TryParse(texto, CultureInfo.InvariantCulture, DateTimeStyles.None, out fecha))
            {
                return fecha.Date;
            }

            return null;
        }

        /// <summary>
        /// Normaliza el código corto a texto comparable (sin espacios y sin el sufijo decimal que agrega Excel a los números).
        /// </summary>
        private static string NormalizarCodigo(string codigo)
        {
            string normalizado = (codigo ?? "").Trim();
            if (normalizado.EndsWith(".0", StringComparison.Ordinal) &&
                double.TryParse(normalizado, NumberStyles.Any, CultureInfo.InvariantCulture, out _))
            {
                normalizado = normalizado[..^2];
            }
            return normalizado;
        }

        /// <summary>
        /// Calcula el lunes de la semana ISO-8601 (la semana inicia obligatoriamente en LUNES).
        /// </summary>
        private static DateTime ObtenerLunes(DateTime fecha)
        {
            return fecha.Date.AddDays(-(((int)fecha.DayOfWeek + 6) % 7));
        }

        /// <summary>
        /// Construye el catálogo CodigoCorto -> SubPlanta a partir de la tabla Auxiliares (T_CodigoRelacion).
        /// </summary>
        private static Dictionary<string, string> ConstruirCatalogoSubPlantas(List<Dictionary<string, object?>> codigoRelacion)
        {
            var catalogo = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var fila in codigoRelacion)
            {
                string codigo = NormalizarCodigo(ObtenerTexto(fila, CamposCodigoCorto));
                string subPlanta = ObtenerTexto(fila, CamposSubPlanta);

                if (string.IsNullOrWhiteSpace(codigo) || string.IsNullOrWhiteSpace(subPlanta)) continue;
                if (!catalogo.ContainsKey(codigo)) catalogo[codigo] = subPlanta;
            }

            return catalogo;
        }

        // =====================================================================
        // 🧮 CONSOLIDADO POR LLAVE (Fecha + CodigoCorto) CON CAMPO CÁLCULO KILOS
        // =====================================================================

        private sealed class ConsolidadoAcumulado
        {
            public DateTime Fecha { get; set; }
            public string CodigoCorto { get; set; } = "";
            public string Nombre { get; set; } = "";
            public string SubPlanta { get; set; } = "";
            public double KilosProduccion { get; set; }
            public double KilosResProdPyDb { get; set; }
            public double KilosDistribucion { get; set; }
            public HashSet<string> Origenes { get; } = new(StringComparer.OrdinalIgnoreCase);

            public double KilosTotal => KilosProduccion + KilosResProdPyDb + KilosDistribucion;
        }

        private static void AcumularTabla(
            IEnumerable<Dictionary<string, object?>> filas,
            string origen,
            Dictionary<string, string> catalogoSubPlantas,
            Dictionary<(DateTime Fecha, string Codigo), ConsolidadoAcumulado> acumulados)
        {
            foreach (var fila in filas)
            {
                DateTime? fecha = ObtenerFecha(fila, CamposFecha);
                string codigo = NormalizarCodigo(ObtenerTexto(fila, CamposCodigoCorto));

                // La llave obligatoria de las tablas de movimiento es Fecha + CodigoCorto
                if (!fecha.HasValue || string.IsNullOrWhiteSpace(codigo)) continue;

                double kilos = ObtenerNumero(fila, CamposKilos);
                string nombre = ObtenerTexto(fila, CamposNombre);
                string subPlanta = ObtenerTexto(fila, CamposSubPlanta);

                if (string.IsNullOrWhiteSpace(subPlanta) && catalogoSubPlantas.TryGetValue(codigo, out var subPlantaCatalogo))
                {
                    subPlanta = subPlantaCatalogo;
                }

                var llave = (fecha.Value.Date, codigo.ToUpperInvariant());

                if (!acumulados.TryGetValue(llave, out var acumulado))
                {
                    acumulado = new ConsolidadoAcumulado
                    {
                        Fecha = fecha.Value.Date,
                        CodigoCorto = codigo
                    };
                    acumulados[llave] = acumulado;
                }

                if (string.IsNullOrWhiteSpace(acumulado.Nombre) && !string.IsNullOrWhiteSpace(nombre))
                {
                    acumulado.Nombre = nombre;
                }

                if (string.IsNullOrWhiteSpace(acumulado.SubPlanta) && !string.IsNullOrWhiteSpace(subPlanta))
                {
                    acumulado.SubPlanta = subPlanta;
                }

                switch (origen)
                {
                    case "Produccion":
                        acumulado.KilosProduccion += kilos;
                        break;
                    case "ResProd_Py_DB":
                        acumulado.KilosResProdPyDb += kilos;
                        break;
                    case "Distribucion":
                        acumulado.KilosDistribucion += kilos;
                        break;
                }

                acumulado.Origenes.Add(origen);
            }
        }

        /// <summary>
        /// Consolida T_Produccion, T_ResProd_Py_DB y T_Distribucion agrupando por Fecha + CodigoCorto
        /// y sumando el campo de cálculo Kilos. La SubPlanta se toma de la propia tabla o, en su defecto,
        /// del catálogo T_CodigoRelacion (hoja Auxiliares).
        /// </summary>
        private static List<CumplimientoPtcDiarioItemDto> ConstruirConsolidado(
            Dictionary<string, List<Dictionary<string, object?>>> tablas)
        {
            var catalogoSubPlantas = ConstruirCatalogoSubPlantas(tablas["T_CodigoRelacion"]);
            var acumulados = new Dictionary<(DateTime Fecha, string Codigo), ConsolidadoAcumulado>();

            AcumularTabla(tablas["T_Produccion"], "Produccion", catalogoSubPlantas, acumulados);
            AcumularTabla(tablas["T_ResProd_Py_DB"], "ResProd_Py_DB", catalogoSubPlantas, acumulados);
            AcumularTabla(tablas["T_Distribucion"], "Distribucion", catalogoSubPlantas, acumulados);

            return acumulados.Values
                .OrderBy(a => a.Fecha)
                .ThenBy(a => a.CodigoCorto, StringComparer.OrdinalIgnoreCase)
                .Select(a =>
                {
                    int semana = ISOWeek.GetWeekOfYear(a.Fecha);
                    int anioSemana = ISOWeek.GetYear(a.Fecha);
                    DateTime lunes = ObtenerLunes(a.Fecha);
                    DateTime domingo = lunes.AddDays(6);

                    return new CumplimientoPtcDiarioItemDto
                    {
                        Fecha = a.Fecha.ToString("yyyy-MM-dd"),
                        Semana = semana,
                        AnioSemana = anioSemana,
                        Lunes = lunes.ToString("yyyy-MM-dd"),
                        Domingo = domingo.ToString("yyyy-MM-dd"),
                        RangoSemana = $"{lunes:dd/MM} al {domingo:dd/MM}",
                        CodigoCorto = a.CodigoCorto,
                        Nombre = a.Nombre,
                        SubPlanta = a.SubPlanta,
                        KilosProduccion = Math.Round(a.KilosProduccion, 4),
                        KilosResProdPyDb = Math.Round(a.KilosResProdPyDb, 4),
                        KilosDistribucion = Math.Round(a.KilosDistribucion, 4),
                        KilosTotal = Math.Round(a.KilosTotal, 4),
                        Origen = string.Join("|", a.Origenes)
                    };
                })
                .ToList();
        }

        private static List<CumplimientoPtcDiarioItemDto> AplicarFiltros(
            List<CumplimientoPtcDiarioItemDto> registros,
            DateTime? desde,
            DateTime? hasta,
            int? semana,
            int? anioSemana,
            string? codigoCorto,
            string? subPlanta)
        {
            IEnumerable<CumplimientoPtcDiarioItemDto> consulta = registros;

            if (desde.HasValue)
            {
                string desdeTexto = desde.Value.ToString("yyyy-MM-dd");
                consulta = consulta.Where(r => string.CompareOrdinal(r.Fecha, desdeTexto) >= 0);
            }

            if (hasta.HasValue)
            {
                string hastaTexto = hasta.Value.ToString("yyyy-MM-dd");
                consulta = consulta.Where(r => string.CompareOrdinal(r.Fecha, hastaTexto) <= 0);
            }

            if (semana.HasValue)
            {
                consulta = consulta.Where(r => r.Semana == semana.Value);
            }

            if (anioSemana.HasValue)
            {
                consulta = consulta.Where(r => r.AnioSemana == anioSemana.Value);
            }

            if (!string.IsNullOrWhiteSpace(codigoCorto))
            {
                var codigosFiltro = codigoCorto.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                               .Select(c => NormalizarCodigo(c).ToUpperInvariant())
                                               .ToHashSet();
                consulta = consulta.Where(r => codigosFiltro.Contains(NormalizarCodigo(r.CodigoCorto).ToUpperInvariant()));
            }

            if (!string.IsNullOrWhiteSpace(subPlanta))
            {
                var subPlantasFiltro = subPlanta.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                               .Select(s => s.Trim().ToUpperInvariant())
                                               .ToHashSet();
                consulta = consulta.Where(r => subPlantasFiltro.Contains((r.SubPlanta ?? "").Trim().ToUpperInvariant()));
            }

            return consulta.ToList();
        }

        /// <summary>
        /// Resumen semanal con semana ISO-8601 (inicio obligatorio en LUNES y cierre en DOMINGO).
        /// </summary>
        private static List<CumplimientoPtcSemanaDto> ConstruirResumenPorSemana(List<CumplimientoPtcDiarioItemDto> registros)
        {
            return registros
                .GroupBy(r => new { r.AnioSemana, r.Semana })
                .Select(g =>
                {
                    var primeraFila = g.First();
                    return new CumplimientoPtcSemanaDto
                    {
                        Semana = g.Key.Semana,
                        AnioSemana = g.Key.AnioSemana,
                        EtiquetaSemana = $"Sem. {g.Key.Semana:00} ({primeraFila.RangoSemana})",
                        Lunes = primeraFila.Lunes,
                        Domingo = primeraFila.Domingo,
                        RangoSemana = primeraFila.RangoSemana,
                        KilosProduccion = Math.Round(g.Sum(r => r.KilosProduccion), 4),
                        KilosResProdPyDb = Math.Round(g.Sum(r => r.KilosResProdPyDb), 4),
                        KilosDistribucion = Math.Round(g.Sum(r => r.KilosDistribucion), 4),
                        KilosTotal = Math.Round(g.Sum(r => r.KilosTotal), 4)
                    };
                })
                .OrderBy(x => x.AnioSemana)
                .ThenBy(x => x.Semana)
                .ToList();
        }

        private static List<CumplimientoPtcSubPlantaDto> ConstruirResumenPorSubPlanta(List<CumplimientoPtcDiarioItemDto> registros)
        {
            return registros
                .GroupBy(r => string.IsNullOrWhiteSpace(r.SubPlanta) ? "SIN SUBPLANTA" : r.SubPlanta.Trim(), StringComparer.OrdinalIgnoreCase)
                .Select(g => new CumplimientoPtcSubPlantaDto
                {
                    SubPlanta = g.Key,
                    Codigos = g.Select(r => r.CodigoCorto).Distinct(StringComparer.OrdinalIgnoreCase).Count(),
                    KilosProduccion = Math.Round(g.Sum(r => r.KilosProduccion), 4),
                    KilosResProdPyDb = Math.Round(g.Sum(r => r.KilosResProdPyDb), 4),
                    KilosDistribucion = Math.Round(g.Sum(r => r.KilosDistribucion), 4),
                    KilosTotal = Math.Round(g.Sum(r => r.KilosTotal), 4)
                })
                .OrderByDescending(x => x.KilosTotal)
                .ToList();
        }

        // =====================================================================
        // 🚀 ENDPOINTS
        // =====================================================================

        /// <summary>
        /// Endpoint principal: consolida las tablas T_Produccion, T_ResProd_Py_DB y T_Distribucion
        /// agrupando por llave Fecha + CodigoCorto y sumando el campo de cálculo Kilos.
        /// La SubPlanta se resuelve con la tabla Auxiliares (T_CodigoRelacion).
        /// Incluye semana ISO-8601 con inicio en LUNES.
        /// GET /api/CumplimientoPtcDiario?desde=2026-11-01&amp;hasta=2026-11-16&amp;semana=46&amp;subPlanta=DESPOSTE&amp;codigoCorto=0123
        /// </summary>
        [HttpGet]
        public IActionResult ObtenerConsolidadoDiario(
            [FromQuery] DateTime? desde = null,
            [FromQuery] DateTime? hasta = null,
            [FromQuery] int? semana = null,
            [FromQuery] int? anioSemana = null,
            [FromQuery] string? codigoCorto = null,
            [FromQuery] string? subPlanta = null,
            [FromQuery] bool recargar = false)
        {
            try
            {
                if (desde.HasValue && hasta.HasValue && desde.Value.Date > hasta.Value.Date)
                {
                    return BadRequest(new
                    {
                        exito = false,
                        error = "El parámetro 'desde' no puede ser mayor que el parámetro 'hasta'."
                    });
                }

                if (semana.HasValue && (semana.Value < 1 || semana.Value > 53))
                {
                    return BadRequest(new
                    {
                        exito = false,
                        error = "El parámetro 'semana' debe estar entre 1 y 53."
                    });
                }

                var tablas = CargarTodasLasTablas(recargar);
                var registros = ConstruirConsolidado(tablas);
                registros = AplicarFiltros(registros, desde, hasta, semana, anioSemana, codigoCorto, subPlanta);

                var resumenPorSemana = ConstruirResumenPorSemana(registros);
                var resumenPorSubPlanta = ConstruirResumenPorSubPlanta(registros);

                return Ok(new
                {
                    exito = true,
                    archivo = Path.GetFileName(ObtenerRutaExcel()),
                    llave = new[] { "Fecha", "CodigoCorto" },
                    campoCalculo = "Kilos",
                    totalRegistros = registros.Count,
                    filtros = new
                    {
                        desde = desde?.ToString("yyyy-MM-dd"),
                        hasta = hasta?.ToString("yyyy-MM-dd"),
                        semana,
                        anioSemana,
                        codigoCorto,
                        subPlanta
                    },
                    resumen = new
                    {
                        totalKilosProduccion = Math.Round(registros.Sum(r => r.KilosProduccion), 4),
                        totalKilosResProdPyDb = Math.Round(registros.Sum(r => r.KilosResProdPyDb), 4),
                        totalKilosDistribucion = Math.Round(registros.Sum(r => r.KilosDistribucion), 4),
                        totalKilosGeneral = Math.Round(registros.Sum(r => r.KilosTotal), 4)
                    },
                    resumenPorSemana,
                    resumenPorSubPlanta,
                    registros
                });
            }
            catch (FileNotFoundException fnfEx)
            {
                return NotFound(new { exito = false, error = fnfEx.Message });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }

        /// <summary>
        /// Devuelve una tabla específica tal cual está en el Excel.
        /// Valores permitidos (nombre oficial de tabla): T_Produccion | T_ResProd_Py_DB | T_Distribucion | T_CodigoRelacion.
        /// GET /api/CumplimientoPtcDiario/tablas/{nombreTabla}?recargar=true
        /// </summary>
        [HttpGet("tablas/{nombreTabla}")]
        public IActionResult ObtenerTablaTalCual(string nombreTabla, [FromQuery] bool recargar = false)
        {
            try
            {
                if (!MapeoTablas.TryGetValue(nombreTabla.Trim(), out var mapeo))
                {
                    return BadRequest(new
                    {
                        exito = false,
                        error = $"La tabla '{nombreTabla}' no está registrada. Valores permitidos: T_Produccion, T_ResProd_Py_DB, T_Distribucion, T_CodigoRelacion."
                    });
                }

                var tablas = CargarTodasLasTablas(recargar);
                return Ok(tablas[mapeo.Tabla]);
            }
            catch (FileNotFoundException fnfEx)
            {
                return NotFound(new { exito = false, error = fnfEx.Message });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }

        /// <summary>
        /// Devuelve el catálogo CodigoCorto -> SubPlanta de la hoja Auxiliares (T_CodigoRelacion).
        /// GET /api/CumplimientoPtcDiario/subplantas?recargar=true
        /// </summary>
        [HttpGet("subplantas")]
        public IActionResult ObtenerCatalogoSubPlantas([FromQuery] bool recargar = false)
        {
            try
            {
                var tablas = CargarTodasLasTablas(recargar);
                var catalogo = ConstruirCatalogoSubPlantas(tablas["T_CodigoRelacion"]);

                var codigos = catalogo
                    .Select(kv => new CodigoRelacionDto
                    {
                        CodigoCorto = kv.Key,
                        SubPlanta = kv.Value
                    })
                    .OrderBy(c => c.SubPlanta, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(c => c.CodigoCorto, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                var subPlantas = codigos
                    .Select(c => c.SubPlanta)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                return Ok(new
                {
                    exito = true,
                    totalCodigos = codigos.Count,
                    subPlantas,
                    codigos
                });
            }
            catch (FileNotFoundException fnfEx)
            {
                return NotFound(new { exito = false, error = fnfEx.Message });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }
    }

    // =========================================================================
    // 📦 DTOs DE RESPUESTA
    // =========================================================================

    public class CumplimientoPtcDiarioItemDto
    {
        public string Fecha { get; set; } = "";
        public int Semana { get; set; }
        public int AnioSemana { get; set; }
        public string Lunes { get; set; } = "";
        public string Domingo { get; set; } = "";
        public string RangoSemana { get; set; } = "";
        public string CodigoCorto { get; set; } = "";
        public string Nombre { get; set; } = "";
        public string SubPlanta { get; set; } = "";
        public double KilosProduccion { get; set; }
        public double KilosResProdPyDb { get; set; }
        public double KilosDistribucion { get; set; }
        public double KilosTotal { get; set; }
        public string Origen { get; set; } = "";
    }

    public class CumplimientoPtcSemanaDto
    {
        public int Semana { get; set; }
        public int AnioSemana { get; set; }
        public string EtiquetaSemana { get; set; } = "";
        public string Lunes { get; set; } = "";
        public string Domingo { get; set; } = "";
        public string RangoSemana { get; set; } = "";
        public double KilosProduccion { get; set; }
        public double KilosResProdPyDb { get; set; }
        public double KilosDistribucion { get; set; }
        public double KilosTotal { get; set; }
    }

    public class CumplimientoPtcSubPlantaDto
    {
        public string SubPlanta { get; set; } = "";
        public int Codigos { get; set; }
        public double KilosProduccion { get; set; }
        public double KilosResProdPyDb { get; set; }
        public double KilosDistribucion { get; set; }
        public double KilosTotal { get; set; }
    }

    public class CodigoRelacionDto
    {
        public string CodigoCorto { get; set; } = "";
        public string SubPlanta { get; set; } = "";
    }
}
