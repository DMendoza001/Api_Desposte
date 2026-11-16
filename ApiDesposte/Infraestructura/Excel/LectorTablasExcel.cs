using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;

namespace ApiDesposte.Infraestructura.Excel
{
    /// <summary>
    /// Lectura estándar de tablas estructuradas de Excel sincronizadas en OneDrive
    /// (regla .agents/rules/conexion-excel-onedrive.md):
    /// - Apertura en modo compartido (FileShare.ReadWrite | FileShare.Delete) copiando el archivo a memoria.
    /// - Limpieza de PivotTableParts / PivotTableCacheDefinitionParts antes de abrir con ClosedXML.
    /// - Tablas devueltas tal cual: columnas tomadas de HeadersRow(), sin campos calculados.
    /// La ruta del archivo la sigue resolviendo cada controlador con su propio ObtenerRutaExcel().
    /// </summary>
    public static class LectorTablasExcel
    {
        public static XLWorkbook AbrirEnMemoria(string rutaExcel)
        {
            byte[] fileBytes;
            using (var fileStream = new FileStream(rutaExcel, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var ms = new MemoryStream())
            {
                fileStream.CopyTo(ms);
                fileBytes = ms.ToArray();
            }

            // Stream expandible: SpreadsheetDocument necesita poder reescribir el paquete
            var msModificado = new MemoryStream();
            msModificado.Write(fileBytes, 0, fileBytes.Length);
            msModificado.Position = 0;

            try
            {
                // Limpieza de PivotTables para evitar conflictos al abrir con ClosedXML
                using var doc = SpreadsheetDocument.Open(msModificado, true);
                if (doc.WorkbookPart != null)
                {
                    foreach (var wsPart in doc.WorkbookPart.WorksheetParts)
                    {
                        foreach (var p in wsPart.PivotTableParts.ToList())
                        {
                            wsPart.DeletePart(p);
                        }
                    }

                    foreach (var cache in doc.WorkbookPart.PivotTableCacheDefinitionParts.ToList())
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
            catch
            {
                msModificado = new MemoryStream(fileBytes);
            }

            msModificado.Position = 0;
            return new XLWorkbook(msModificado);
        }

        /// <summary>
        /// Lee una tabla estructurada tal cual está definida en la hoja, conservando exactamente
        /// sus encabezados y valores originales. Si la tabla no está en la hoja indicada, la busca
        /// en todo el libro.
        /// </summary>
        public static List<Dictionary<string, object?>> LeerTablaTalCual(XLWorkbook workbook, string nombreHoja, string nombreTabla)
        {
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
                    tabla = hoja.Tables.FirstOrDefault(t => MismoNombre(t.Name, nombreTabla));
                }
            }

            tabla ??= workbook.Worksheets.SelectMany(w => w.Tables).FirstOrDefault(t => MismoNombre(t.Name, nombreTabla))
                ?? throw new InvalidOperationException($"No se encontró la tabla '{nombreTabla}' en la hoja '{nombreHoja}'.");

            var headers = tabla.HeadersRow().Cells().Select(c => c.Value.ToString().Trim()).ToList();
            var registros = new List<Dictionary<string, object?>>();

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

        /// <summary>
        /// Convierte una celda a un valor JSON simple: bool, long (enteros), double (4 decimales),
        /// fecha "yyyy-MM-dd" (o con hora), hora "hh:mm:ss" o texto. Vacío = null.
        /// </summary>
        public static object? ObtenerValorCelda(IXLCell celda)
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
                        return celda.GetTimeSpan().ToString(@"hh\:mm\:ss");

                    default:
                        return TextoONulo(celda);
                }
            }
            catch
            {
                return TextoONulo(celda);
            }
        }

        private static string? TextoONulo(IXLCell celda)
        {
            var texto = celda.Value.ToString().Trim();
            return string.IsNullOrEmpty(texto) ? null : texto;
        }

        private static bool MismoNombre(string a, string b) =>
            string.Equals(a.Trim(), b.Trim(), StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Cache en memoria de un conjunto de tablas de un mismo libro. Se declara una instancia
    /// estática por controlador para que lecturas consecutivas desde el frontend no reabran el Excel.
    /// </summary>
    public sealed class CacheTablasExcel
    {
        private readonly object _lock = new();
        private readonly TimeSpan _duracion;
        private Dictionary<string, List<Dictionary<string, object?>>>? _tablas;
        private DateTime _timestamp = DateTime.MinValue;

        public CacheTablasExcel(TimeSpan duracion)
        {
            _duracion = duracion;
        }

        /// <param name="rutaExcel">Ruta resuelta por el ObtenerRutaExcel() del controlador.</param>
        /// <param name="mapeo">Clave de respuesta -> (Hoja, Tabla) en el libro.</param>
        /// <param name="forzarRecarga">true para ignorar el cache y releer el archivo.</param>
        public Dictionary<string, List<Dictionary<string, object?>>> Obtener(
            string rutaExcel,
            IReadOnlyDictionary<string, (string Hoja, string Tabla)> mapeo,
            bool forzarRecarga = false)
        {
            lock (_lock)
            {
                if (!forzarRecarga && _tablas != null && (DateTime.UtcNow - _timestamp) < _duracion)
                {
                    return _tablas;
                }

                if (!File.Exists(rutaExcel))
                {
                    throw new FileNotFoundException($"No se encontró el archivo Excel en la ruta: {rutaExcel}");
                }

                using var workbook = LectorTablasExcel.AbrirEnMemoria(rutaExcel);

                var resultado = new Dictionary<string, List<Dictionary<string, object?>>>(StringComparer.OrdinalIgnoreCase);
                foreach (var (clave, (hoja, tabla)) in mapeo)
                {
                    resultado[clave] = LectorTablasExcel.LeerTablaTalCual(workbook, hoja, tabla);
                }

                _tablas = resultado;
                _timestamp = DateTime.UtcNow;
                return _tablas;
            }
        }
    }
}
