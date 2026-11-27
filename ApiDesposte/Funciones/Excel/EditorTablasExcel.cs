using System.Collections.Concurrent;
using System.Globalization;
using System.Text.Json;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Table = DocumentFormat.OpenXml.Spreadsheet.Table;

namespace ApiDesposte.Funciones.Excel
{
    /// <summary>
    /// Error de negocio al editar una tabla de Excel. <see cref="Codigo"/> es el estado HTTP
    /// que debe devolver el controlador (400 dato inválido, 404 no existe, 409 conflicto).
    /// </summary>
    public sealed class ExcepcionTablaExcel : Exception
    {
        public int Codigo { get; }
        public ExcepcionTablaExcel(int codigo, string mensaje) : base(mensaje) => Codigo = codigo;
    }

    /// <summary>Columna de una tabla de Excel. Las columnas calculadas (fórmula de tabla) no se escriben.</summary>
    public sealed record ColumnaTablaExcel(string Nombre, bool Calculada);

    /// <summary>Contenido de una tabla para edición: columnas + registros con su número de fila.</summary>
    public sealed record TablaEditable(List<ColumnaTablaExcel> Columnas, List<Dictionary<string, object?>> Registros);

    public sealed record OperacionActualizar(int Fila, Dictionary<string, JsonElement> Valores, Dictionary<string, JsonElement>? Original);
    public sealed record OperacionEliminar(int Fila, Dictionary<string, JsonElement>? Original);

    /// <summary>Resultado de un lote: filas actualizadas (con algún cambio), eliminadas y números de fila creados.</summary>
    public sealed record ResultadoLote(int Actualizadas, int Eliminadas, List<int> Creadas);

    /// <summary>
    /// CRUD sobre tablas estructuradas de un libro Excel sincronizado en OneDrive
    /// (regla .agents/rules/conexion-excel-onedrive.md).
    ///
    /// - Cada registro se identifica con "_fila": su posición (1 = primera fila de datos) dentro de la tabla.
    ///   Como la posición cambia al eliminar, Actualizar/Eliminar aceptan "original" (los valores que el
    ///   cliente leyó) y responden 409 si la fila ya no coincide.
    /// - Las columnas calculadas de la tabla no se escriben: al insertar se copia la fórmula de la fila anterior.
    /// - Al actualizar sólo se escriben las celdas cuyo valor cambió, para no borrar fórmulas sueltas.
    /// - Cada escritura deja antes un respaldo del archivo en la carpeta indicada (se conservan los últimos 30).
    /// - El libro se guarda sin limpiar tablas dinámicas (a diferencia de la lectura de LectorTablasExcel).
    /// </summary>
    public static class EditorTablasExcel
    {
        public const string CampoFila = "_fila";
        private const int RespaldosAConservar = 30;
        private static readonly ConcurrentDictionary<string, object> Bloqueos = new(StringComparer.OrdinalIgnoreCase);

        // =====================================================================
        // LECTURA
        // =====================================================================

        public static TablaEditable Leer(string rutaExcel, string nombreHoja, string nombreTabla)
        {
            byte[] bytes = LectorTablasExcel.LeerBytesCompartido(rutaExcel);
            var calculadas = ColumnasCalculadas(FormulasCalculadas(bytes), nombreTabla);
            using var workbook = LectorTablasExcel.AbrirDesdeBytes(bytes);
            var tabla = BuscarTabla(workbook, nombreHoja, nombreTabla);

            var columnas = Encabezados(tabla)
                .Select(n => new ColumnaTablaExcel(n, calculadas.Contains(n)))
                .ToList();

            var registros = new List<Dictionary<string, object?>>();
            int numero = 0;
            foreach (var fila in tabla.DataRange.Rows())
            {
                numero++;
                var registro = new Dictionary<string, object?> { [CampoFila] = numero };
                bool tieneDatos = false;
                for (int i = 0; i < columnas.Count; i++)
                {
                    object? valor = LectorTablasExcel.ObtenerValorCelda(fila.Cell(i + 1));
                    if (valor != null) tieneDatos = true;
                    registro[columnas[i].Nombre] = valor;
                }
                if (tieneDatos) registros.Add(registro);
            }

            return new TablaEditable(columnas, registros);
        }

        // =====================================================================
        // ESCRITURA
        // =====================================================================

        /// <returns>Número de fila ("_fila") del registro creado.</returns>
        public static int Insertar(string rutaExcel, string carpetaRespaldos, string nombreHoja, string nombreTabla,
            Dictionary<string, JsonElement> valores) =>
            Escribir(rutaExcel, carpetaRespaldos, nombreTabla,
                (workbook, calculadas) => InsertarEn(BuscarTabla(workbook, nombreHoja, nombreTabla), calculadas, valores));

        /// <returns>Nombres de las columnas que cambiaron.</returns>
        public static List<string> Actualizar(string rutaExcel, string carpetaRespaldos, string nombreHoja, string nombreTabla,
            int fila, Dictionary<string, JsonElement> valores, Dictionary<string, JsonElement>? original) =>
            Escribir(rutaExcel, carpetaRespaldos, nombreTabla,
                (workbook, calculadas) => ActualizarEn(BuscarTabla(workbook, nombreHoja, nombreTabla), calculadas, fila, valores, original));

        public static void Eliminar(string rutaExcel, string carpetaRespaldos, string nombreHoja, string nombreTabla,
            int fila, Dictionary<string, JsonElement>? original) =>
            Escribir(rutaExcel, carpetaRespaldos, nombreTabla, (workbook, calculadas) =>
            {
                EliminarEn(BuscarTabla(workbook, nombreHoja, nombreTabla), calculadas, fila, original);
                return true;
            });

        /// <summary>
        /// Aplica varias operaciones sobre una tabla en UNA sola escritura del archivo (todo o nada):
        /// primero las actualizaciones, luego las eliminaciones (de la última fila a la primera, para que
        /// los números de fila recibidos sigan siendo válidos) y al final las altas.
        /// Si una operación falla no se guarda nada.
        /// </summary>
        public static ResultadoLote Lote(string rutaExcel, string carpetaRespaldos, string nombreHoja, string nombreTabla,
            List<Dictionary<string, JsonElement>> crear, List<OperacionActualizar> actualizar, List<OperacionEliminar> eliminar)
        {
            var repetidas = actualizar.Select(a => a.Fila).Concat(eliminar.Select(e => e.Fila))
                .GroupBy(f => f).Where(g => g.Count() > 1).Select(g => g.Key).ToList();
            if (repetidas.Count > 0)
                throw new ExcepcionTablaExcel(400, $"Filas repetidas en el lote: {string.Join(", ", repetidas)}.");
            if (crear.Count + actualizar.Count + eliminar.Count == 0)
                return new ResultadoLote(0, 0, []);

            return Escribir(rutaExcel, carpetaRespaldos, nombreTabla, (workbook, calculadas) =>
            {
                var tabla = BuscarTabla(workbook, nombreHoja, nombreTabla);
                int actualizadas = 0;
                foreach (var op in actualizar)
                    if (ActualizarEn(tabla, calculadas, op.Fila, op.Valores, op.Original).Count > 0) actualizadas++;
                foreach (var op in eliminar.OrderByDescending(e => e.Fila))
                    EliminarEn(tabla, calculadas, op.Fila, op.Original);
                var creadas = crear.Select(v => InsertarEn(tabla, calculadas, v)).ToList();
                return new ResultadoLote(actualizadas, eliminar.Count, creadas);
            });
        }

        // =====================================================================
        // OPERACIONES SOBRE EL LIBRO ABIERTO
        // =====================================================================

        private static int InsertarEn(IXLTable tabla, HashSet<string> calculadas, Dictionary<string, JsonElement> valores)
        {
            if (tabla.ShowTotalsRow)
                throw new ExcepcionTablaExcel(400, $"La tabla '{tabla.Name}' tiene fila de totales; no se admite insertar desde la API.");

            var encabezados = Encabezados(tabla);
            var indices = IndicesColumnas(encabezados, valores.Keys);
            var hoja = tabla.Worksheet;
            int colIni = tabla.RangeAddress.FirstAddress.ColumnNumber;
            int colFin = tabla.RangeAddress.LastAddress.ColumnNumber;
            int filasDatos = tabla.DataRange.RowCount();

            IXLRangeRow? anterior = filasDatos > 0 ? tabla.DataRange.Row(filasDatos) : null;
            int filaHoja;
            int numero;

            if (filasDatos == 1 && FilaVacia(anterior!))
            {
                // Tabla sin datos: Excel deja siempre una fila en blanco, se usa esa
                filaHoja = anterior!.RangeAddress.FirstAddress.RowNumber;
                numero = 1;
                anterior = null;
            }
            else
            {
                filaHoja = tabla.RangeAddress.LastAddress.RowNumber + 1;
                var debajo = hoja.Range(filaHoja, colIni, filaHoja, colFin);
                // Si debajo de la tabla hay datos, se desplazan sólo las celdas de sus columnas
                if (debajo.CellsUsed(XLCellsUsedOptions.Contents).Any())
                    debajo.InsertRowsAbove(1);
                tabla.Resize(hoja.Cell(tabla.RangeAddress.FirstAddress.RowNumber, colIni), hoja.Cell(filaHoja, colFin));
                numero = filasDatos + 1;
            }

            for (int i = 0; i < encabezados.Count; i++)
            {
                if (anterior == null) continue;
                var celda = hoja.Cell(filaHoja, colIni + i);
                celda.Style = anterior.Cell(i + 1).Style;
                if (!calculadas.Contains(encabezados[i])) continue;

                // Fórmula de la fila con fórmula más cercana hacia arriba (alguna fila pudo ser sobrescrita a mano)
                var modelo = tabla.DataRange.Column(i + 1).Cells()
                    .Where(c => c.Address.RowNumber < filaHoja && c.HasFormula)
                    .LastOrDefault();
                if (modelo == null) continue;
                // Referencias estructuradas ([#This Row]) no dependen de la posición: se copian en A1
                if (modelo.FormulaA1.Contains('[')) celda.FormulaA1 = modelo.FormulaA1;
                else celda.FormulaR1C1 = modelo.FormulaR1C1;
            }

            foreach (var (nombre, i) in indices)
            {
                if (calculadas.Contains(encabezados[i])) continue;
                var celda = hoja.Cell(filaHoja, colIni + i);
                var referencia = anterior?.Cell(i + 1) ?? celda;
                celda.Value = AValorCelda(valores[nombre], referencia, encabezados[i]);
            }

            return numero;
        }

        private static List<string> ActualizarEn(IXLTable tabla, HashSet<string> calculadas, int fila,
            Dictionary<string, JsonElement> valores, Dictionary<string, JsonElement>? original)
        {
            var encabezados = Encabezados(tabla);
            var filaRango = ObtenerFila(tabla, fila);
            VerificarOriginal(filaRango, encabezados, original, fila);

            var cambiadas = new List<string>();
            foreach (var (nombre, i) in IndicesColumnas(encabezados, valores.Keys))
            {
                if (calculadas.Contains(encabezados[i])) continue;
                var celda = filaRango.Cell(i + 1);
                if (Coincide(LectorTablasExcel.ObtenerValorCelda(celda), valores[nombre])) continue;
                celda.Value = AValorCelda(valores[nombre], celda, encabezados[i]);
                cambiadas.Add(encabezados[i]);
            }
            return cambiadas;
        }

        private static void EliminarEn(IXLTable tabla, HashSet<string> calculadas, int fila, Dictionary<string, JsonElement>? original)
        {
            var encabezados = Encabezados(tabla);
            var filaRango = ObtenerFila(tabla, fila);
            VerificarOriginal(filaRango, encabezados, original, fila);

            if (tabla.DataRange.RowCount() == 1)
            {
                // Una tabla de Excel no puede quedarse sin filas de datos: se vacían los valores
                for (int i = 0; i < encabezados.Count; i++)
                    if (!calculadas.Contains(encabezados[i])) filaRango.Cell(i + 1).Clear(XLClearOptions.Contents);
            }
            else
            {
                filaRango.Delete(XLShiftDeletedCells.ShiftCellsUp);
            }
        }

        // =====================================================================
        // INTERNOS
        // =====================================================================

        private static T Escribir<T>(string rutaExcel, string carpetaRespaldos, string nombreTabla,
            Func<XLWorkbook, HashSet<string>, T> accion)
        {
            if (!File.Exists(rutaExcel))
                throw new FileNotFoundException($"No se encontró el archivo Excel en la ruta: {rutaExcel}");

            lock (Bloqueos.GetOrAdd(Path.GetFullPath(rutaExcel), _ => new object()))
            {
                byte[] bytes = LectorTablasExcel.LeerBytesCompartido(rutaExcel);
                var formulas = FormulasCalculadas(bytes);
                var calculadas = ColumnasCalculadas(formulas, nombreTabla);

                byte[] nuevo;
                T resultado;
                using (var workbook = new XLWorkbook(new MemoryStream(bytes)))
                {
                    resultado = accion(workbook, calculadas);
                    // ClosedXML no conserva los valores calculados de las fórmulas: se evalúan al guardar para que
                    // otros lectores (dashboards) los vean, y Excel recalcula todo al abrir el archivo.
                    workbook.RecalculateAllFormulas();
                    workbook.FullCalculationOnLoad = true;
                    using var salida = new MemoryStream();
                    workbook.SaveAs(salida, new SaveOptions { EvaluateFormulasBeforeSaving = true });
                    RestaurarFormulasCalculadas(salida, formulas);
                    nuevo = salida.ToArray();
                }

                GuardarRespaldo(rutaExcel, carpetaRespaldos, bytes);

                // Se escribe todo de una vez; si Excel de escritorio tiene el archivo bloqueado lanza IOException
                using (var destino = new FileStream(rutaExcel, FileMode.Truncate, FileAccess.Write, FileShare.ReadWrite | FileShare.Delete))
                {
                    destino.Write(nuevo, 0, nuevo.Length);
                }
                File.SetLastWriteTime(rutaExcel, DateTime.Now);
                return resultado;
            }
        }

        private static void GuardarRespaldo(string rutaExcel, string carpetaRespaldos, byte[] bytes)
        {
            Directory.CreateDirectory(carpetaRespaldos);
            string nombre = Path.GetFileNameWithoutExtension(rutaExcel);
            string extension = Path.GetExtension(rutaExcel);
            File.WriteAllBytes(Path.Combine(carpetaRespaldos, $"{nombre}_{DateTime.Now:yyyyMMdd_HHmmss_fff}{extension}"), bytes);

            foreach (var viejo in new DirectoryInfo(carpetaRespaldos)
                         .GetFiles($"{nombre}_*{extension}")
                         .OrderByDescending(f => f.Name)
                         .Skip(RespaldosAConservar))
            {
                try { viejo.Delete(); } catch { /* respaldo en uso: se borrará en la próxima escritura */ }
            }
        }

        /// <summary>
        /// Fórmulas de columna calculada de todas las tablas del libro: tabla -> columna -> elemento XML.
        /// </summary>
        private static Dictionary<string, Dictionary<string, CalculatedColumnFormula>> FormulasCalculadas(byte[] bytes)
        {
            var resultado = new Dictionary<string, Dictionary<string, CalculatedColumnFormula>>(StringComparer.OrdinalIgnoreCase);
            using var doc = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
            foreach (var def in TablasDelLibro(doc))
            {
                var columnas = new Dictionary<string, CalculatedColumnFormula>(StringComparer.OrdinalIgnoreCase);
                foreach (var col in def.TableColumns?.Elements<TableColumn>() ?? [])
                {
                    if (col.CalculatedColumnFormula != null && col.Name?.Value != null)
                        columnas[col.Name.Value.Trim()] = (CalculatedColumnFormula)col.CalculatedColumnFormula.CloneNode(true);
                }
                resultado[def.Name?.Value ?? def.DisplayName?.Value ?? ""] = columnas;
            }
            return resultado;
        }

        private static HashSet<string> ColumnasCalculadas(Dictionary<string, Dictionary<string, CalculatedColumnFormula>> formulas, string nombreTabla)
        {
            var clave = formulas.Keys.FirstOrDefault(k => MismoNombre(k, nombreTabla));
            return clave == null
                ? new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                : new HashSet<string>(formulas[clave].Keys, StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// ClosedXML descarta al guardar las fórmulas de columna calculada con referencias estructuradas
        /// (ej. SUM(Tabla[[#This Row],[A]:[B]])): se vuelven a poner en la definición de la tabla.
        /// </summary>
        private static void RestaurarFormulasCalculadas(MemoryStream libro, Dictionary<string, Dictionary<string, CalculatedColumnFormula>> formulas)
        {
            libro.Position = 0;
            using (var doc = SpreadsheetDocument.Open(libro, true))
            {
                foreach (var def in TablasDelLibro(doc))
                {
                    var clave = formulas.Keys.FirstOrDefault(k => MismoNombre(k, def.Name?.Value ?? def.DisplayName?.Value ?? ""));
                    if (clave == null) continue;
                    bool cambio = false;
                    foreach (var col in def.TableColumns?.Elements<TableColumn>() ?? [])
                    {
                        string nombre = col.Name?.Value?.Trim() ?? "";
                        if (!formulas[clave].TryGetValue(nombre, out var original)) continue;
                        if (!string.IsNullOrWhiteSpace(col.CalculatedColumnFormula?.Text)) continue;
                        col.CalculatedColumnFormula?.Remove();
                        col.PrependChild((CalculatedColumnFormula)original.CloneNode(true));
                        cambio = true;
                    }
                    if (cambio) def.Save();
                }
            }
            libro.Position = 0;
        }

        private static IEnumerable<Table> TablasDelLibro(SpreadsheetDocument doc) =>
            (doc.WorkbookPart?.WorksheetParts.SelectMany(w => w.TableDefinitionParts) ?? [])
                .Select(p => p.Table)
                .Where(t => t != null)!;

        private static IXLTable BuscarTabla(XLWorkbook workbook, string nombreHoja, string nombreTabla)
        {
            var hoja = workbook.Worksheets.FirstOrDefault(w => MismoNombre(w.Name, nombreHoja));
            var tabla = hoja?.Tables.FirstOrDefault(t => MismoNombre(t.Name, nombreTabla))
                        ?? workbook.Worksheets.SelectMany(w => w.Tables).FirstOrDefault(t => MismoNombre(t.Name, nombreTabla));
            return tabla ?? throw new ExcepcionTablaExcel(404, $"No se encontró la tabla '{nombreTabla}' en la hoja '{nombreHoja}'.");
        }

        private static List<string> Encabezados(IXLTable tabla) =>
            tabla.HeadersRow().Cells()
                .Select((c, i) => { var n = c.Value.ToString().Trim(); return string.IsNullOrWhiteSpace(n) ? $"Columna_{i + 1}" : n; })
                .ToList();

        private static List<(string Nombre, int Indice)> IndicesColumnas(List<string> encabezados, IEnumerable<string> claves)
        {
            var resultado = new List<(string, int)>();
            var desconocidas = new List<string>();
            foreach (var clave in claves)
            {
                if (clave == CampoFila) continue;
                int i = encabezados.FindIndex(e => MismoNombre(e, clave));
                if (i < 0) desconocidas.Add(clave); else resultado.Add((clave, i));
            }
            if (desconocidas.Count > 0)
                throw new ExcepcionTablaExcel(400, $"Columnas que no existen en la tabla: {string.Join(", ", desconocidas)}. Columnas válidas: {string.Join(", ", encabezados)}.");
            return resultado;
        }

        private static IXLRangeRow ObtenerFila(IXLTable tabla, int fila)
        {
            int total = tabla.DataRange.RowCount();
            if (fila < 1 || fila > total)
                throw new ExcepcionTablaExcel(404, $"La fila {fila} no existe (la tabla tiene {total} filas de datos).");
            return tabla.DataRange.Row(fila);
        }

        private static bool FilaVacia(IXLRangeRow fila) => !fila.CellsUsed(XLCellsUsedOptions.Contents).Any();

        private static void VerificarOriginal(IXLRangeRow fila, List<string> encabezados, Dictionary<string, JsonElement>? original, int numero)
        {
            if (original == null) return;
            foreach (var (nombre, i) in IndicesColumnas(encabezados, original.Keys))
            {
                if (!Coincide(LectorTablasExcel.ObtenerValorCelda(fila.Cell(i + 1)), original[nombre]))
                    throw new ExcepcionTablaExcel(409, $"La fila {numero} cambió desde que se leyó (columna '{encabezados[i]}'). Recarga los datos e intenta de nuevo.");
            }
        }

        /// <summary>Compara el valor actual de la celda (como lo devuelve la lectura) con el valor JSON recibido.</summary>
        private static bool Coincide(object? actual, JsonElement esperado)
        {
            switch (esperado.ValueKind)
            {
                case JsonValueKind.Null:
                case JsonValueKind.Undefined:
                    return actual == null;
                case JsonValueKind.True:
                case JsonValueKind.False:
                    return actual is bool b && b == esperado.GetBoolean();
                case JsonValueKind.Number:
                    return actual switch
                    {
                        long l => Math.Abs(l - esperado.GetDouble()) < 1e-9,
                        double d => Math.Abs(d - esperado.GetDouble()) < 1e-6,
                        _ => false
                    };
                case JsonValueKind.String:
                    string texto = esperado.GetString()!.Trim();
                    if (actual == null) return texto.Length == 0;
                    if (actual is long or double)
                        return double.TryParse(texto, NumberStyles.Float, CultureInfo.InvariantCulture, out var n)
                               && Math.Abs(Convert.ToDouble(actual, CultureInfo.InvariantCulture) - n) < 1e-6;
                    return string.Equals(Convert.ToString(actual, CultureInfo.InvariantCulture), texto, StringComparison.Ordinal);
                default:
                    return false;
            }
        }

        private static readonly string[] FormatosFecha =
        [
            "yyyy-MM-dd", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd HH:mm", "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-ddTHH:mm",
            "yyyy-MM-ddTHH:mm:ss.fff", "yyyy-MM-ddTHH:mm:ss.fffZ", "yyyy-MM-ddTHH:mm:ssZ"
        ];
        private static readonly string[] FormatosHora = [@"h\:mm", @"hh\:mm", @"h\:mm\:ss", @"hh\:mm\:ss"];

        /// <summary>
        /// Convierte el valor JSON al tipo de celda: null/"" = vacío, número, booleano,
        /// fecha "yyyy-MM-dd[ HH:mm[:ss]]", hora "HH:mm[:ss]" o texto. Un texto numérico sólo se
        /// guarda como número si la celda de referencia (misma columna) ya es numérica.
        /// </summary>
        private static XLCellValue AValorCelda(JsonElement valor, IXLCell referencia, string columna)
        {
            switch (valor.ValueKind)
            {
                case JsonValueKind.Null:
                case JsonValueKind.Undefined:
                    return Blank.Value;
                case JsonValueKind.True:
                case JsonValueKind.False:
                    return valor.GetBoolean();
                case JsonValueKind.Number:
                    return valor.GetDouble();
                case JsonValueKind.String:
                    string texto = valor.GetString()!.Trim();
                    if (texto.Length == 0) return Blank.Value;
                    if (DateTime.TryParseExact(texto, FormatosFecha, CultureInfo.InvariantCulture, DateTimeStyles.None, out var fecha))
                        return fecha;
                    if (TimeSpan.TryParseExact(texto, FormatosHora, CultureInfo.InvariantCulture, out var hora))
                        return hora;
                    if (referencia.DataType == XLDataType.Number
                        && double.TryParse(texto, NumberStyles.Float, CultureInfo.InvariantCulture, out var numero))
                        return numero;
                    return texto;
                default:
                    throw new ExcepcionTablaExcel(400, $"Valor no admitido en la columna '{columna}': se esperaba texto, número, booleano, fecha o null.");
            }
        }

        private static bool MismoNombre(string a, string b) =>
            string.Equals(a.Trim(), b.Trim(), StringComparison.OrdinalIgnoreCase);
    }
}
