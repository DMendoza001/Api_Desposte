using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System.IO;

namespace ApiDesposte.Controllers.Excel
{
    [ApiController]
    [Route("api/[controller]")]
    [Route("api/CumplimientosDetallePy")]
    public class CumplimientosDetalleControllerPy : ControllerBase
    {
        private const string HojaResProdPy = "ResProd_Py";
        private const string TablaResProdPy = "T_ResProd_Py";

        private const string HojaDemandaPy = "DemandaPy";
        private const string TablaDemandaPy = "T_DemandaPy";

        private const string HojaAuxiliares = "Auxiliares";
        private const string TablaCodigoRelacion = "T_CodigoRelacion";
        private const string TablaArticuloVentas = "T_ArticuloVentas";

        // Cache en memoria para evitar relecturas continuas de disco al filtrar
        private static List<CumplimientoPyResumenDto>? _cacheResumen = null;
        private static List<ResProdPyItemDto>? _cacheProdItems = null;
        private static List<DemandaPyItemDto>? _cacheDemandaItems = null;
        private static DateTime _cacheTimestamp = DateTime.MinValue;
        private static readonly object _cacheLock = new();
        private static readonly TimeSpan CacheDuracion = TimeSpan.FromSeconds(45);

        private string ObtenerRutaExcel()
        {
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return Path.Combine(userProfile, @"OneDrive - Corporación Rico SAC\Desposte-03\PlantasCore\Planificacion\Cumplimiento_PTC.xlsm");
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
                using (var doc = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(msModificado, true))
                {
                    if (doc.WorkbookPart != null)
                    {
                        // 1. Eliminar PivotTableParts de todas las hojas para evitar conflictos con ClosedXML
                        foreach (var wsPart in doc.WorkbookPart.WorksheetParts)
                        {
                            var pParts = wsPart.PivotTableParts.ToList();
                            foreach (var p in pParts)
                            {
                                wsPart.DeletePart(p);
                            }
                        }

                        // 2. Eliminar PivotTableCacheDefinitionParts del Workbook
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

        private static int BuscarIndiceColumna(IXLTable tabla, params string[] nombresPosibles)
        {
            var headers = tabla.HeadersRow().Cells().ToList();
            foreach (var nombre in nombresPosibles)
            {
                for (int i = 0; i < headers.Count; i++)
                {
                    string headerName = headers[i].Value.ToString().Trim();
                    if (string.Equals(headerName, nombre, StringComparison.OrdinalIgnoreCase))
                    {
                        return i + 1; // ClosedXML es 1-based index
                    }
                }
            }
            return -1;
        }

        private static string ObtenerValorCeldaTexto(IXLRangeRow fila, int colIndex, string valorPorDefecto = "")
        {
            if (colIndex <= 0) return valorPorDefecto;
            var celda = fila.Cell(colIndex);
            if (celda.IsEmpty()) return valorPorDefecto;
            return celda.Value.ToString().Trim();
        }

        private static double ObtenerValorCeldaNumero(IXLRangeRow fila, int colIndex, double valorPorDefecto = 0.0)
        {
            if (colIndex <= 0) return valorPorDefecto;
            var celda = fila.Cell(colIndex);
            if (celda.IsEmpty()) return valorPorDefecto;

            if (celda.DataType == XLDataType.Number)
            {
                return celda.GetValue<double>();
            }

            string txt = celda.Value.ToString().Trim().Replace(",", ".");
            if (double.TryParse(txt, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double val))
            {
                return val;
            }

            if (double.TryParse(celda.Value.ToString().Trim(), out double valLocal))
            {
                return valLocal;
            }

            return valorPorDefecto;
        }

        private static IXLTable? ObtenerTabla(IXLWorksheet ws, string nombreTabla)
        {
            try
            {
                return ws.Table(nombreTabla);
            }
            catch
            {
                return ws.Tables.FirstOrDefault(t => string.Equals(t.Name, nombreTabla, StringComparison.OrdinalIgnoreCase));
            }
        }

        /// <summary>
        /// Genera la tabla resumen consolidada agrupando T_ResProd_Py y T_DemandaPy por (Semana, CodigoCorto)
        /// incorporando el campo SubPlanta y cruzando con T_CodigoRelacion y auxiliares.
        /// </summary>
        private List<CumplimientoPyResumenDto> GenerarResumenConsolidado(bool forzarRecarga = false)
        {
            lock (_cacheLock)
            {
                if (!forzarRecarga && _cacheResumen != null && (DateTime.Now - _cacheTimestamp) < CacheDuracion)
                {
                    return _cacheResumen;
                }

                string rutaExcel = ObtenerRutaExcel();
                if (!System.IO.File.Exists(rutaExcel))
                {
                    throw new FileNotFoundException($"No se encontró el archivo Excel en la ruta: {rutaExcel}");
                }

                using var workbook = CargarWorkbookEnMemoria(rutaExcel);

                // 1. CARGAR TABLAS AUXILIARES
                var dicCodigoRelacion = new Dictionary<string, CodigoRelacionPyAuxDto>(StringComparer.OrdinalIgnoreCase);
                var dicArticuloVentas = new Dictionary<string, ArticuloVentasPyAuxDto>(StringComparer.OrdinalIgnoreCase);

                if (workbook.Worksheets.TryGetWorksheet(HojaAuxiliares, out var wsAux))
                {
                    // T_CodigoRelacion: CodigoCorto -> Nombre, CodigoArticulo, TipoCong
                    var tablaRel = ObtenerTabla(wsAux, TablaCodigoRelacion);
                    if (tablaRel != null)
                    {
                        int colCodArt = BuscarIndiceColumna(tablaRel, "CodigoArticulo", "Producto", "Codigo_Articulo");
                        int colCodCorto = BuscarIndiceColumna(tablaRel, "CodigoCorto", "CodCorto", "Codigo_Corto");
                        int colNombre = BuscarIndiceColumna(tablaRel, "Nombre", "Descripcion");

                        foreach (var fila in tablaRel.DataRange.Rows())
                        {
                            if (fila.IsEmpty()) continue;
                            string codCorto = ObtenerValorCeldaTexto(fila, colCodCorto);
                            if (string.IsNullOrWhiteSpace(codCorto)) continue;

                            dicCodigoRelacion[codCorto] = new CodigoRelacionPyAuxDto
                            {
                                CodigoCorto = codCorto,
                                CodigoArticulo = ObtenerValorCeldaTexto(fila, colCodArt),
                                Nombre = ObtenerValorCeldaTexto(fila, colNombre)
                            };
                        }
                    }

                    // T_ArticuloVentas: CodigoVenta -> NLinea, NFamilia, NCategoria, Nombre
                    var tablaArt = ObtenerTabla(wsAux, TablaArticuloVentas);
                    if (tablaArt != null)
                    {
                        int colCodVenta = BuscarIndiceColumna(tablaArt, "CodigoVenta", "Codigoventa");
                        int colNombre = BuscarIndiceColumna(tablaArt, "Nombre", "Descripcion");
                        int colLinea = BuscarIndiceColumna(tablaArt, "NLinea", "Linea");
                        int colFamilia = BuscarIndiceColumna(tablaArt, "NFamilia", "Familia");
                        int colCategoria = BuscarIndiceColumna(tablaArt, "NCategoria", "Categoria");

                        foreach (var fila in tablaArt.DataRange.Rows())
                        {
                            if (fila.IsEmpty()) continue;
                            string codVenta = ObtenerValorCeldaTexto(fila, colCodVenta);
                            if (string.IsNullOrWhiteSpace(codVenta)) continue;

                            dicArticuloVentas[codVenta] = new ArticuloVentasPyAuxDto
                            {
                                CodigoVenta = codVenta,
                                Nombre = ObtenerValorCeldaTexto(fila, colNombre),
                                NLinea = ObtenerValorCeldaTexto(fila, colLinea),
                                NFamilia = ObtenerValorCeldaTexto(fila, colFamilia),
                                NCategoria = ObtenerValorCeldaTexto(fila, colCategoria)
                            };
                        }
                    }
                }

                // 2. CARGAR T_ResProd_Py (Hoja ResProd_Py)
                var prodItems = new List<ResProdPyItemDto>();
                var prodAgrupado = new Dictionary<string, AcumuladorProdPyDto>(StringComparer.OrdinalIgnoreCase);

                if (workbook.Worksheets.TryGetWorksheet(HojaResProdPy, out var wsProd))
                {
                    var tablaProd = ObtenerTabla(wsProd, TablaResProdPy);
                    if (tablaProd != null)
                    {
                        int colSemana = BuscarIndiceColumna(tablaProd, "Semana");
                        int colProducto = BuscarIndiceColumna(tablaProd, "Producto");
                        int colCodCorto = BuscarIndiceColumna(tablaProd, "CodigoCorto", "CodCorto");
                        int colNombre = BuscarIndiceColumna(tablaProd, "Nombre");
                        int colKilos = BuscarIndiceColumna(tablaProd, "Kilos", "Kgs");
                        int colLinea = BuscarIndiceColumna(tablaProd, "Linea", "NLinea");
                        int colFamilia = BuscarIndiceColumna(tablaProd, "Familia", "NFamilia");
                        int colCategoria = BuscarIndiceColumna(tablaProd, "Categoria", "NCategoria");
                        int colSubPlanta = BuscarIndiceColumna(tablaProd, "SubPlanta", "Sub_Planta", "Planta");

                        foreach (var fila in tablaProd.DataRange.Rows())
                        {
                            if (fila.IsEmpty()) continue;
                            string sem = ObtenerValorCeldaTexto(fila, colSemana);
                            string codCorto = ObtenerValorCeldaTexto(fila, colCodCorto);
                            if (string.IsNullOrWhiteSpace(sem) || string.IsNullOrWhiteSpace(codCorto)) continue;

                            double kilos = ObtenerValorCeldaNumero(fila, colKilos);
                            string prod = ObtenerValorCeldaTexto(fila, colProducto);
                            string nom = ObtenerValorCeldaTexto(fila, colNombre);
                            string lin = ObtenerValorCeldaTexto(fila, colLinea);
                            string fam = ObtenerValorCeldaTexto(fila, colFamilia);
                            string cat = ObtenerValorCeldaTexto(fila, colCategoria);
                            string subPlanta = ObtenerValorCeldaTexto(fila, colSubPlanta).ToUpper();

                            int anio = 0, nSemana = 0;
                            if (sem.Length >= 5)
                            {
                                int.TryParse(sem.Substring(0, 4), out anio);
                                int.TryParse(sem.Substring(4), out nSemana);
                            }

                            if (string.IsNullOrWhiteSpace(subPlanta)) subPlanta = "DESPOSTE";

                            dicCodigoRelacion.TryGetValue(codCorto, out var rel);
                            if (rel != null)
                            {
                                if (string.IsNullOrWhiteSpace(nom)) nom = rel.Nombre;
                                if (string.IsNullOrWhiteSpace(prod)) prod = rel.CodigoArticulo;
                            }

                            dicArticuloVentas.TryGetValue(codCorto, out var art);
                            if (art != null)
                            {
                                if (string.IsNullOrWhiteSpace(nom)) nom = art.Nombre;
                                if (string.IsNullOrWhiteSpace(lin)) lin = art.NLinea;
                                if (string.IsNullOrWhiteSpace(fam)) fam = art.NFamilia;
                                if (string.IsNullOrWhiteSpace(cat)) cat = art.NCategoria;
                            }

                            prodItems.Add(new ResProdPyItemDto
                            {
                                Semana = sem,
                                Anio = anio,
                                NSemana = nSemana,
                                SubPlanta = subPlanta,
                                CodigoCorto = codCorto,
                                Producto = prod,
                                Nombre = nom,
                                Kilos = Math.Round(kilos, 2),
                                Linea = lin,
                                Familia = string.IsNullOrWhiteSpace(fam) ? "SIN FAMILIA" : fam,
                                Categoria = string.IsNullOrWhiteSpace(cat) ? "SIN CATEGORIA" : cat
                            });

                            string key = $"{sem}|{codCorto}";
                            if (!prodAgrupado.TryGetValue(key, out var acum))
                            {
                                acum = new AcumuladorProdPyDto
                                {
                                    Semana = sem,
                                    CodigoCorto = codCorto,
                                    SubPlanta = subPlanta,
                                    Producto = prod,
                                    Nombre = nom,
                                    Linea = lin,
                                    Familia = fam,
                                    Categoria = cat,
                                    KilosProd = 0
                                };
                                prodAgrupado[key] = acum;
                            }
                            acum.KilosProd += kilos;
                            if (string.IsNullOrWhiteSpace(acum.SubPlanta) && !string.IsNullOrWhiteSpace(subPlanta)) acum.SubPlanta = subPlanta;
                            if (string.IsNullOrWhiteSpace(acum.Producto) && !string.IsNullOrWhiteSpace(prod)) acum.Producto = prod;
                            if (string.IsNullOrWhiteSpace(acum.Nombre) && !string.IsNullOrWhiteSpace(nom)) acum.Nombre = nom;
                            if (string.IsNullOrWhiteSpace(acum.Linea) && !string.IsNullOrWhiteSpace(lin)) acum.Linea = lin;
                            if (string.IsNullOrWhiteSpace(acum.Familia) && !string.IsNullOrWhiteSpace(fam)) acum.Familia = fam;
                            if (string.IsNullOrWhiteSpace(acum.Categoria) && !string.IsNullOrWhiteSpace(cat)) acum.Categoria = cat;
                        }
                    }
                }

                // 3. CARGAR T_DemandaPy (Hoja DemandaPy)
                var demItems = new List<DemandaPyItemDto>();
                var demAgrupado = new Dictionary<string, AcumuladorDemPyDto>(StringComparer.OrdinalIgnoreCase);

                if (workbook.Worksheets.TryGetWorksheet(HojaDemandaPy, out var wsDem))
                {
                    var tablaDem = ObtenerTabla(wsDem, TablaDemandaPy);
                    if (tablaDem != null)
                    {
                        int colSemana = BuscarIndiceColumna(tablaDem, "Semana");
                        int colCodCorto = BuscarIndiceColumna(tablaDem, "CodigoCorto", "CodCorto");
                        int colProducto = BuscarIndiceColumna(tablaDem, "Producto");
                        int colNombre = BuscarIndiceColumna(tablaDem, "Nombre");
                        int colKilos = BuscarIndiceColumna(tablaDem, "Kilos", "Kgs");
                        int colOrigen = BuscarIndiceColumna(tablaDem, "Origen", "Canal", "Destino");
                        int colLinea = BuscarIndiceColumna(tablaDem, "Linea", "NLinea");
                        int colFamilia = BuscarIndiceColumna(tablaDem, "Familia", "NFamilia");
                        int colCategoria = BuscarIndiceColumna(tablaDem, "Categoria", "NCategoria");
                        int colSubPlanta = BuscarIndiceColumna(tablaDem, "SubPlanta", "Sub_Planta", "Planta");

                        foreach (var fila in tablaDem.DataRange.Rows())
                        {
                            if (fila.IsEmpty()) continue;
                            string sem = ObtenerValorCeldaTexto(fila, colSemana);
                            string codCorto = ObtenerValorCeldaTexto(fila, colCodCorto);
                            if (string.IsNullOrWhiteSpace(sem) || string.IsNullOrWhiteSpace(codCorto)) continue;

                            double kilos = ObtenerValorCeldaNumero(fila, colKilos);
                            string prod = ObtenerValorCeldaTexto(fila, colProducto);
                            string nom = ObtenerValorCeldaTexto(fila, colNombre);
                            string origen = ObtenerValorCeldaTexto(fila, colOrigen);
                            string lin = ObtenerValorCeldaTexto(fila, colLinea);
                            string fam = ObtenerValorCeldaTexto(fila, colFamilia);
                            string cat = ObtenerValorCeldaTexto(fila, colCategoria);
                            string subPlanta = ObtenerValorCeldaTexto(fila, colSubPlanta).ToUpper();

                            int anio = 0, nSemana = 0;
                            if (sem.Length >= 5)
                            {
                                int.TryParse(sem.Substring(0, 4), out anio);
                                int.TryParse(sem.Substring(4), out nSemana);
                            }

                            if (string.IsNullOrWhiteSpace(subPlanta)) subPlanta = "DESPOSTE";

                            dicCodigoRelacion.TryGetValue(codCorto, out var rel);
                            if (rel != null)
                            {
                                if (string.IsNullOrWhiteSpace(nom)) nom = rel.Nombre;
                                if (string.IsNullOrWhiteSpace(prod)) prod = rel.CodigoArticulo;
                            }

                            dicArticuloVentas.TryGetValue(codCorto, out var art);
                            if (art != null)
                            {
                                if (string.IsNullOrWhiteSpace(nom)) nom = art.Nombre;
                                if (string.IsNullOrWhiteSpace(lin)) lin = art.NLinea;
                                if (string.IsNullOrWhiteSpace(fam)) fam = art.NFamilia;
                                if (string.IsNullOrWhiteSpace(cat)) cat = art.NCategoria;
                            }

                            demItems.Add(new DemandaPyItemDto
                            {
                                Semana = sem,
                                Anio = anio,
                                NSemana = nSemana,
                                SubPlanta = subPlanta,
                                CodigoCorto = codCorto,
                                Producto = prod,
                                Nombre = nom,
                                Kilos = Math.Round(kilos, 2),
                                Origen = string.IsNullOrWhiteSpace(origen) ? "OTROS" : origen.ToUpper(),
                                Linea = lin,
                                Familia = string.IsNullOrWhiteSpace(fam) ? "SIN FAMILIA" : fam,
                                Categoria = string.IsNullOrWhiteSpace(cat) ? "SIN CATEGORIA" : cat
                            });

                            string key = $"{sem}|{codCorto}";
                            if (!demAgrupado.TryGetValue(key, out var acum))
                            {
                                acum = new AcumuladorDemPyDto
                                {
                                    Semana = sem,
                                    CodigoCorto = codCorto,
                                    SubPlanta = subPlanta,
                                    Producto = prod,
                                    Nombre = nom,
                                    Linea = lin,
                                    Familia = fam,
                                    Categoria = cat,
                                    KilosPy = 0
                                };
                                demAgrupado[key] = acum;
                            }
                            acum.KilosPy += kilos;
                            if (string.IsNullOrWhiteSpace(acum.SubPlanta) && !string.IsNullOrWhiteSpace(subPlanta)) acum.SubPlanta = subPlanta;
                            if (string.IsNullOrWhiteSpace(acum.Producto) && !string.IsNullOrWhiteSpace(prod)) acum.Producto = prod;
                            if (string.IsNullOrWhiteSpace(acum.Nombre) && !string.IsNullOrWhiteSpace(nom)) acum.Nombre = nom;
                            if (string.IsNullOrWhiteSpace(acum.Linea) && !string.IsNullOrWhiteSpace(lin)) acum.Linea = lin;
                            if (string.IsNullOrWhiteSpace(acum.Familia) && !string.IsNullOrWhiteSpace(fam)) acum.Familia = fam;
                            if (string.IsNullOrWhiteSpace(acum.Categoria) && !string.IsNullOrWhiteSpace(cat)) acum.Categoria = cat;
                        }
                    }
                }

                // 4. FULL OUTER JOIN DE TODAS LAS LLAVES (Semana, CodigoCorto)
                var todasLasLlaves = new HashSet<string>(prodAgrupado.Keys, StringComparer.OrdinalIgnoreCase);
                todasLasLlaves.UnionWith(demAgrupado.Keys);

                var listaResumen = new List<CumplimientoPyResumenDto>(todasLasLlaves.Count);

                foreach (var key in todasLasLlaves)
                {
                    prodAgrupado.TryGetValue(key, out var pItem);
                    demAgrupado.TryGetValue(key, out var dItem);

                    string sem = pItem?.Semana ?? dItem?.Semana ?? "";
                    string codCorto = pItem?.CodigoCorto ?? dItem?.CodigoCorto ?? "";
                    double kilosProd = Math.Round(pItem?.KilosProd ?? 0.0, 2);
                    double kilosPy = Math.Round(dItem?.KilosPy ?? 0.0, 2);
                    double diferencia = Math.Round(kilosProd - kilosPy, 2);

                    // Resolver SubPlanta
                    string subPlanta = !string.IsNullOrWhiteSpace(pItem?.SubPlanta) ? pItem!.SubPlanta
                                     : !string.IsNullOrWhiteSpace(dItem?.SubPlanta) ? dItem!.SubPlanta : "";
                    if (string.IsNullOrWhiteSpace(subPlanta)) subPlanta = "DESPOSTE";

                    // Buscar en auxiliares: T_CodigoRelacion para Nombre y CodigoArticulo
                    dicCodigoRelacion.TryGetValue(codCorto, out var codRel);
                    dicArticuloVentas.TryGetValue(codCorto, out var artVenta);

                    // Resolver Producto
                    string prod = pItem?.Producto ?? dItem?.Producto ?? codRel?.CodigoArticulo ?? "";
                    if (artVenta == null && !string.IsNullOrWhiteSpace(prod))
                    {
                        dicArticuloVentas.TryGetValue(prod, out artVenta);
                    }

                    // Resolver Nombre
                    string nombre = !string.IsNullOrWhiteSpace(codRel?.Nombre) ? codRel!.Nombre
                                  : !string.IsNullOrWhiteSpace(artVenta?.Nombre) ? artVenta!.Nombre
                                  : !string.IsNullOrWhiteSpace(pItem?.Nombre) ? pItem!.Nombre
                                  : !string.IsNullOrWhiteSpace(dItem?.Nombre) ? dItem!.Nombre : "";

                    // Resolver Linea, Familia, Categoria
                    string linea = !string.IsNullOrWhiteSpace(artVenta?.NLinea) ? artVenta!.NLinea
                                 : !string.IsNullOrWhiteSpace(pItem?.Linea) ? pItem!.Linea
                                 : !string.IsNullOrWhiteSpace(dItem?.Linea) ? dItem!.Linea : "";

                    string familia = !string.IsNullOrWhiteSpace(artVenta?.NFamilia) ? artVenta!.NFamilia
                                   : !string.IsNullOrWhiteSpace(pItem?.Familia) ? pItem!.Familia
                                   : !string.IsNullOrWhiteSpace(dItem?.Familia) ? dItem!.Familia : "";

                    string categoria = !string.IsNullOrWhiteSpace(artVenta?.NCategoria) ? artVenta!.NCategoria
                                     : !string.IsNullOrWhiteSpace(pItem?.Categoria) ? pItem!.Categoria
                                     : !string.IsNullOrWhiteSpace(dItem?.Categoria) ? dItem!.Categoria : "";

                    int anio = 0, nSemana = 0;
                    if (!string.IsNullOrWhiteSpace(sem) && sem.Length >= 5)
                    {
                        int.TryParse(sem.Substring(0, 4), out anio);
                        int.TryParse(sem.Substring(4), out nSemana);
                    }

                    double pctCumplimiento = kilosPy > 0 ? Math.Round((kilosProd / kilosPy) * 100.0, 2) : (kilosProd > 0 ? 100.0 : 0.0);
                    double pctVariacion = kilosPy > 0 ? Math.Round(((kilosProd - kilosPy) / kilosPy) * 100.0, 2) : 0.0;

                    listaResumen.Add(new CumplimientoPyResumenDto
                    {
                        Anio = anio,
                        NSemana = nSemana,
                        Semana = sem,
                        SubPlanta = subPlanta.ToUpper(),
                        CodigoCorto = codCorto,
                        Producto = prod,
                        KilosProd = kilosProd,
                        KilosPy = kilosPy,
                        Diferencia = diferencia,
                        PctVariacion = pctVariacion,
                        PctCumplimiento = pctCumplimiento,
                        Linea = linea,
                        Familia = string.IsNullOrWhiteSpace(familia) ? "SIN FAMILIA" : familia,
                        Categoria = string.IsNullOrWhiteSpace(categoria) ? "SIN CATEGORIA" : categoria,
                        Nombre = nombre
                    });
                }

                listaResumen = listaResumen
                    .OrderBy(r => r.Semana)
                    .ThenBy(r => r.SubPlanta)
                    .ThenBy(r => r.Familia)
                    .ThenBy(r => r.Categoria)
                    .ThenBy(r => r.CodigoCorto)
                    .ToList();

                _cacheResumen = listaResumen;
                _cacheProdItems = prodItems;
                _cacheDemandaItems = demItems;
                _cacheTimestamp = DateTime.Now;

                return _cacheResumen;
            }
        }

        /// <summary>
        /// Obtiene la tabla resumen plana calculada con soporte para filtro por SubPlanta, Semana, Año, Familia, etc.
        /// </summary>
        [HttpGet("resumen")]
        public IActionResult ObtenerTablaResumen(
            [FromQuery] string? semana = null,
            [FromQuery] string? anio = null,
            [FromQuery] string? subPlanta = null,
            [FromQuery] string? familia = null,
            [FromQuery] string? categoria = null,
            [FromQuery] string? search = null,
            [FromQuery] bool recargar = false)
        {
            try
            {
                var resumen = GenerarResumenConsolidado(recargar);

                if (!string.IsNullOrWhiteSpace(semana))
                {
                    var semanasFiltro = semana.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                              .Select(s => s.Trim().ToUpper())
                                              .ToHashSet();
                    resumen = resumen.Where(r => semanasFiltro.Contains(r.Semana.ToUpper())).ToList();
                }

                if (!string.IsNullOrWhiteSpace(anio) && int.TryParse(anio, out int aFiltro))
                {
                    resumen = resumen.Where(r => r.Anio == aFiltro).ToList();
                }

                if (!string.IsNullOrWhiteSpace(subPlanta))
                {
                    var spFiltro = subPlanta.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                            .Select(s => s.Trim().ToUpper())
                                            .ToHashSet();
                    resumen = resumen.Where(r => spFiltro.Contains(r.SubPlanta.ToUpper())).ToList();
                }

                if (!string.IsNullOrWhiteSpace(familia))
                {
                    resumen = resumen.Where(r => string.Equals(r.Familia, familia.Trim(), StringComparison.OrdinalIgnoreCase)).ToList();
                }

                if (!string.IsNullOrWhiteSpace(categoria))
                {
                    resumen = resumen.Where(r => string.Equals(r.Categoria, categoria.Trim(), StringComparison.OrdinalIgnoreCase)).ToList();
                }

                if (!string.IsNullOrWhiteSpace(search))
                {
                    string q = search.Trim();
                    resumen = resumen.Where(r =>
                        r.CodigoCorto.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                        r.Producto.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                        r.Nombre.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                        r.SubPlanta.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                        r.Familia.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                        r.Categoria.Contains(q, StringComparison.OrdinalIgnoreCase)
                    ).ToList();
                }

                return Ok(resumen);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message, stackTrace = ex.ToString() });
            }
        }

        /// <summary>
        /// Devuelve las semanas disponibles con sus totales para poblar el selector de filtros
        /// </summary>
        [HttpGet("semanas")]
        public IActionResult ObtenerSemanas([FromQuery] string? subPlanta = null)
        {
            try
            {
                var resumen = GenerarResumenConsolidado();

                if (!string.IsNullOrWhiteSpace(subPlanta))
                {
                    var spFiltro = subPlanta.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                            .Select(s => s.Trim().ToUpper())
                                            .ToHashSet();
                    resumen = resumen.Where(r => spFiltro.Contains(r.SubPlanta.ToUpper())).ToList();
                }

                var semanas = resumen
                    .GroupBy(r => r.Semana)
                    .Select(g => new
                    {
                        Semana = g.Key,
                        Anio = g.First().Anio,
                        NSemana = g.First().NSemana,
                        TotalProdKilos = Math.Round(g.Sum(x => x.KilosProd), 2),
                        TotalPyKilos = Math.Round(g.Sum(x => x.KilosPy), 2),
                        DiferenciaKilos = Math.Round(g.Sum(x => x.Diferencia), 2),
                        PctCumplimiento = g.Sum(x => x.KilosPy) > 0 ? Math.Round((g.Sum(x => x.KilosProd) / g.Sum(x => x.KilosPy)) * 100.0, 2) : 0,
                        TotalSkus = g.Count(),
                        TieneDemanda = g.Any(x => x.KilosPy > 0),
                        TieneProduccion = g.Any(x => x.KilosProd > 0)
                    })
                    .OrderBy(s => s.Semana)
                    .ToList();

                return Ok(semanas);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }

        /// <summary>
        /// Devuelve el catálogo de SubPlantas con sus totales de producción y demanda proyectada
        /// </summary>
        [HttpGet("subplantas")]
        public IActionResult ObtenerSubPlantas([FromQuery] string? semana = null)
        {
            try
            {
                var resumen = GenerarResumenConsolidado();

                if (!string.IsNullOrWhiteSpace(semana))
                {
                    var semFilter = semana.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                          .Select(s => s.Trim().ToUpper())
                                          .ToHashSet();
                    resumen = resumen.Where(r => semFilter.Contains(r.Semana.ToUpper())).ToList();
                }

                var subPlantas = resumen
                    .GroupBy(r => string.IsNullOrWhiteSpace(r.SubPlanta) ? "DESPOSTE" : r.SubPlanta.ToUpper())
                    .Select(g => new
                    {
                        SubPlanta = g.Key,
                        TotalProdKilos = Math.Round(g.Sum(x => x.KilosProd), 2),
                        TotalPyKilos = Math.Round(g.Sum(x => x.KilosPy), 2),
                        DiferenciaKilos = Math.Round(g.Sum(x => x.Diferencia), 2),
                        PctCumplimiento = g.Sum(x => x.KilosPy) > 0 ? Math.Round((g.Sum(x => x.KilosProd) / g.Sum(x => x.KilosPy)) * 100.0, 2) : 0,
                        TotalSkus = g.Count(),
                        TieneDemanda = g.Any(x => x.KilosPy > 0),
                        TieneProduccion = g.Any(x => x.KilosProd > 0)
                    })
                    .OrderByDescending(sp => sp.TotalProdKilos)
                    .ToList();

                return Ok(subPlantas);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }

        /// <summary>
        /// Matriz 1: Agrupada jerárquicamente por Semana -> SubPlanta -> Familia -> Categoria
        /// </summary>
        [HttpGet("matriz-semana")]
        public IActionResult ObtenerMatrizSemana([FromQuery] string? semana = null, [FromQuery] string? subPlanta = null)
        {
            try
            {
                var resumen = GenerarResumenConsolidado();

                if (!string.IsNullOrWhiteSpace(semana))
                {
                    var semFilter = semana.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                          .Select(s => s.Trim().ToUpper())
                                          .ToHashSet();
                    resumen = resumen.Where(r => semFilter.Contains(r.Semana.ToUpper())).ToList();
                }

                if (!string.IsNullOrWhiteSpace(subPlanta))
                {
                    var spFilter = subPlanta.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                            .Select(s => s.Trim().ToUpper())
                                            .ToHashSet();
                    resumen = resumen.Where(r => spFilter.Contains(r.SubPlanta.ToUpper())).ToList();
                }

                var matriz = resumen
                    .GroupBy(r => new { r.Semana, r.Anio, r.NSemana })
                    .OrderByDescending(gSem => gSem.Key.Semana)
                    .Select(gSem => new
                    {
                        Semana = gSem.Key.Semana,
                        Anio = gSem.Key.Anio,
                        NSemana = gSem.Key.NSemana,
                        KilosProd = Math.Round(gSem.Sum(x => x.KilosProd), 2),
                        KilosPy = Math.Round(gSem.Sum(x => x.KilosPy), 2),
                        Diferencia = Math.Round(gSem.Sum(x => x.Diferencia), 2),
                        PctVariacion = gSem.Sum(x => x.KilosPy) > 0 ? Math.Round(((gSem.Sum(x => x.KilosProd) - gSem.Sum(x => x.KilosPy)) / gSem.Sum(x => x.KilosPy)) * 100.0, 2) : 0,
                        PctCumplimiento = gSem.Sum(x => x.KilosPy) > 0 ? Math.Round((gSem.Sum(x => x.KilosProd) / gSem.Sum(x => x.KilosPy)) * 100.0, 2) : 0,
                        SubPlantas = gSem
                            .GroupBy(x => string.IsNullOrWhiteSpace(x.SubPlanta) ? "DESPOSTE" : x.SubPlanta.ToUpper())
                            .OrderBy(gSp => gSp.Key)
                            .Select(gSp => new
                            {
                                SubPlanta = gSp.Key,
                                KilosProd = Math.Round(gSp.Sum(x => x.KilosProd), 2),
                                KilosPy = Math.Round(gSp.Sum(x => x.KilosPy), 2),
                                Diferencia = Math.Round(gSp.Sum(x => x.Diferencia), 2),
                                PctVariacion = gSp.Sum(x => x.KilosPy) > 0 ? Math.Round(((gSp.Sum(x => x.KilosProd) - gSp.Sum(x => x.KilosPy)) / gSp.Sum(x => x.KilosPy)) * 100.0, 2) : 0,
                                PctCumplimiento = gSp.Sum(x => x.KilosPy) > 0 ? Math.Round((gSp.Sum(x => x.KilosProd) / gSp.Sum(x => x.KilosPy)) * 100.0, 2) : 0,
                                Familias = gSp
                                    .GroupBy(x => string.IsNullOrWhiteSpace(x.Familia) ? "SIN FAMILIA" : x.Familia)
                                    .OrderBy(gFam => gFam.Key)
                                    .Select(gFam => new
                                    {
                                        Familia = gFam.Key,
                                        KilosProd = Math.Round(gFam.Sum(x => x.KilosProd), 2),
                                        KilosPy = Math.Round(gFam.Sum(x => x.KilosPy), 2),
                                        Diferencia = Math.Round(gFam.Sum(x => x.Diferencia), 2),
                                        PctVariacion = gFam.Sum(x => x.KilosPy) > 0 ? Math.Round(((gFam.Sum(x => x.KilosProd) - gFam.Sum(x => x.KilosPy)) / gFam.Sum(x => x.KilosPy)) * 100.0, 2) : 0,
                                        PctCumplimiento = gFam.Sum(x => x.KilosPy) > 0 ? Math.Round((gFam.Sum(x => x.KilosProd) / gFam.Sum(x => x.KilosPy)) * 100.0, 2) : 0,
                                        Categorias = gFam
                                            .GroupBy(x => string.IsNullOrWhiteSpace(x.Categoria) ? "SIN CATEGORIA" : x.Categoria)
                                            .OrderBy(gCat => gCat.Key)
                                            .Select(gCat => new
                                            {
                                                Categoria = gCat.Key,
                                                KilosProd = Math.Round(gCat.Sum(x => x.KilosProd), 2),
                                                KilosPy = Math.Round(gCat.Sum(x => x.KilosPy), 2),
                                                Diferencia = Math.Round(gCat.Sum(x => x.Diferencia), 2),
                                                PctVariacion = gCat.Sum(x => x.KilosPy) > 0 ? Math.Round(((gCat.Sum(x => x.KilosProd) - gCat.Sum(x => x.KilosPy)) / gCat.Sum(x => x.KilosPy)) * 100.0, 2) : 0,
                                                PctCumplimiento = gCat.Sum(x => x.KilosPy) > 0 ? Math.Round((gCat.Sum(x => x.KilosProd) / gCat.Sum(x => x.KilosPy)) * 100.0, 2) : 0
                                            }).ToList()
                                    }).ToList()
                            }).ToList()
                    }).ToList();

                return Ok(matriz);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }

        /// <summary>
        /// Matriz 2: Agrupada jerárquicamente por SubPlanta -> Familia -> Categoria -> CodigoCorto -> Nombre
        /// </summary>
        [HttpGet("matriz-producto")]
        public IActionResult ObtenerMatrizProducto([FromQuery] string? semana = null, [FromQuery] string? subPlanta = null)
        {
            try
            {
                var resumen = GenerarResumenConsolidado();

                if (!string.IsNullOrWhiteSpace(semana))
                {
                    var semFilter = semana.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                          .Select(s => s.Trim().ToUpper())
                                          .ToHashSet();
                    resumen = resumen.Where(r => semFilter.Contains(r.Semana.ToUpper())).ToList();
                }

                if (!string.IsNullOrWhiteSpace(subPlanta))
                {
                    var spFilter = subPlanta.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                            .Select(s => s.Trim().ToUpper())
                                            .ToHashSet();
                    resumen = resumen.Where(r => spFilter.Contains(r.SubPlanta.ToUpper())).ToList();
                }

                var matriz = resumen
                    .GroupBy(r => string.IsNullOrWhiteSpace(r.SubPlanta) ? "DESPOSTE" : r.SubPlanta.ToUpper())
                    .OrderBy(gSp => gSp.Key)
                    .Select(gSp => new
                    {
                        SubPlanta = gSp.Key,
                        KilosProd = Math.Round(gSp.Sum(x => x.KilosProd), 2),
                        KilosPy = Math.Round(gSp.Sum(x => x.KilosPy), 2),
                        Diferencia = Math.Round(gSp.Sum(x => x.Diferencia), 2),
                        PctVariacion = gSp.Sum(x => x.KilosPy) > 0 ? Math.Round(((gSp.Sum(x => x.KilosProd) - gSp.Sum(x => x.KilosPy)) / gSp.Sum(x => x.KilosPy)) * 100.0, 2) : 0,
                        PctCumplimiento = gSp.Sum(x => x.KilosPy) > 0 ? Math.Round((gSp.Sum(x => x.KilosProd) / gSp.Sum(x => x.KilosPy)) * 100.0, 2) : 0,
                        Familias = gSp
                            .GroupBy(x => string.IsNullOrWhiteSpace(x.Familia) ? "SIN FAMILIA" : x.Familia)
                            .OrderBy(gFam => gFam.Key)
                            .Select(gFam => new
                            {
                                Familia = gFam.Key,
                                KilosProd = Math.Round(gFam.Sum(x => x.KilosProd), 2),
                                KilosPy = Math.Round(gFam.Sum(x => x.KilosPy), 2),
                                Diferencia = Math.Round(gFam.Sum(x => x.Diferencia), 2),
                                PctVariacion = gFam.Sum(x => x.KilosPy) > 0 ? Math.Round(((gFam.Sum(x => x.KilosProd) - gFam.Sum(x => x.KilosPy)) / gFam.Sum(x => x.KilosPy)) * 100.0, 2) : 0,
                                PctCumplimiento = gFam.Sum(x => x.KilosPy) > 0 ? Math.Round((gFam.Sum(x => x.KilosProd) / gFam.Sum(x => x.KilosPy)) * 100.0, 2) : 0,
                                Categorias = gFam
                                    .GroupBy(x => string.IsNullOrWhiteSpace(x.Categoria) ? "SIN CATEGORIA" : x.Categoria)
                                    .OrderBy(gCat => gCat.Key)
                                    .Select(gCat => new
                                    {
                                        Categoria = gCat.Key,
                                        KilosProd = Math.Round(gCat.Sum(x => x.KilosProd), 2),
                                        KilosPy = Math.Round(gCat.Sum(x => x.KilosPy), 2),
                                        Diferencia = Math.Round(gCat.Sum(x => x.Diferencia), 2),
                                        PctVariacion = gCat.Sum(x => x.KilosPy) > 0 ? Math.Round(((gCat.Sum(x => x.KilosProd) - gCat.Sum(x => x.KilosPy)) / gCat.Sum(x => x.KilosPy)) * 100.0, 2) : 0,
                                        PctCumplimiento = gCat.Sum(x => x.KilosPy) > 0 ? Math.Round((gCat.Sum(x => x.KilosProd) / gCat.Sum(x => x.KilosPy)) * 100.0, 2) : 0,
                                        Articulos = gCat
                                            .GroupBy(x => new { x.CodigoCorto, x.Nombre, x.Producto })
                                            .OrderBy(gArt => gArt.Key.CodigoCorto)
                                            .Select(gArt => new
                                            {
                                                CodigoCorto = gArt.Key.CodigoCorto,
                                                Nombre = gArt.Key.Nombre,
                                                Producto = gArt.Key.Producto,
                                                KilosProd = Math.Round(gArt.Sum(x => x.KilosProd), 2),
                                                KilosPy = Math.Round(gArt.Sum(x => x.KilosPy), 2),
                                                Diferencia = Math.Round(gArt.Sum(x => x.Diferencia), 2),
                                                PctVariacion = gArt.Sum(x => x.KilosPy) > 0 ? Math.Round(((gArt.Sum(x => x.KilosProd) - gArt.Sum(x => x.KilosPy)) / gArt.Sum(x => x.KilosPy)) * 100.0, 2) : 0,
                                                PctCumplimiento = gArt.Sum(x => x.KilosPy) > 0 ? Math.Round((gArt.Sum(x => x.KilosProd) / gArt.Sum(x => x.KilosPy)) * 100.0, 2) : 0
                                            }).ToList()
                                    }).ToList()
                            }).ToList()
                    }).ToList();

                return Ok(matriz);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }

        /// <summary>
        /// Análisis comparativo consolidado general entre T_ResProd_Py y T_DemandaPy
        /// </summary>
        [HttpGet("analisis")]
        public IActionResult ObtenerAnalisisComparativo([FromQuery] string? semana = null, [FromQuery] string? subPlanta = null)
        {
            try
            {
                var resumen = GenerarResumenConsolidado();

                if (!string.IsNullOrWhiteSpace(semana))
                {
                    var semFilter = semana.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                          .Select(s => s.Trim().ToUpper())
                                          .ToHashSet();
                    resumen = resumen.Where(r => semFilter.Contains(r.Semana.ToUpper())).ToList();
                }

                if (!string.IsNullOrWhiteSpace(subPlanta))
                {
                    var spFilter = subPlanta.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                            .Select(s => s.Trim().ToUpper())
                                            .ToHashSet();
                    resumen = resumen.Where(r => spFilter.Contains(r.SubPlanta.ToUpper())).ToList();
                }

                double totalProd = resumen.Sum(r => r.KilosProd);
                double totalPy = resumen.Sum(r => r.KilosPy);
                double diferenciaNeta = totalProd - totalPy;
                double pctCumplimientoGlobal = totalPy > 0 ? Math.Round((totalProd / totalPy) * 100.0, 2) : 0.0;

                int skusTotales = resumen.Count;
                int skusAmbos = resumen.Count(r => r.KilosProd > 0 && r.KilosPy > 0);
                int skusSoloProd = resumen.Count(r => r.KilosProd > 0 && r.KilosPy == 0);
                int skusSoloDem = resumen.Count(r => r.KilosProd == 0 && r.KilosPy > 0);

                var topSobrecumplimiento = resumen
                    .Where(r => r.Diferencia > 0)
                    .OrderByDescending(r => r.Diferencia)
                    .Take(10)
                    .Select(r => new
                    {
                        r.Anio,
                        r.NSemana,
                        r.Semana,
                        r.SubPlanta,
                        r.CodigoCorto,
                        r.Producto,
                        r.Nombre,
                        r.Familia,
                        r.Categoria,
                        r.KilosProd,
                        r.KilosPy,
                        ExcedenteKilos = r.Diferencia,
                        PctCumplimiento = r.PctCumplimiento
                    }).ToList();

                var topSubcumplimiento = resumen
                    .Where(r => r.Diferencia < 0)
                    .OrderBy(r => r.Diferencia)
                    .Take(10)
                    .Select(r => new
                    {
                        r.Anio,
                        r.NSemana,
                        r.Semana,
                        r.SubPlanta,
                        r.CodigoCorto,
                        r.Producto,
                        r.Nombre,
                        r.Familia,
                        r.Categoria,
                        r.KilosProd,
                        r.KilosPy,
                        DeficitKilos = Math.Abs(r.Diferencia),
                        PctCumplimiento = r.PctCumplimiento
                    }).ToList();

                var resumenSubPlantas = resumen
                    .GroupBy(r => string.IsNullOrWhiteSpace(r.SubPlanta) ? "DESPOSTE" : r.SubPlanta.ToUpper())
                    .Select(g => new
                    {
                        SubPlanta = g.Key,
                        KilosProd = Math.Round(g.Sum(x => x.KilosProd), 2),
                        KilosPy = Math.Round(g.Sum(x => x.KilosPy), 2),
                        Diferencia = Math.Round(g.Sum(x => x.Diferencia), 2),
                        PctCumplimiento = g.Sum(x => x.KilosPy) > 0 ? Math.Round((g.Sum(x => x.KilosProd) / g.Sum(x => x.KilosPy)) * 100.0, 2) : 0,
                        TotalSkus = g.Count()
                    })
                    .OrderByDescending(sp => sp.KilosProd)
                    .ToList();

                var resumenFamilias = resumen
                    .GroupBy(r => string.IsNullOrWhiteSpace(r.Familia) ? "SIN FAMILIA" : r.Familia)
                    .Select(g => new
                    {
                        Familia = g.Key,
                        KilosProd = Math.Round(g.Sum(x => x.KilosProd), 2),
                        KilosPy = Math.Round(g.Sum(x => x.KilosPy), 2),
                        Diferencia = Math.Round(g.Sum(x => x.Diferencia), 2),
                        PctCumplimiento = g.Sum(x => x.KilosPy) > 0 ? Math.Round((g.Sum(x => x.KilosProd) / g.Sum(x => x.KilosPy)) * 100.0, 2) : 0,
                        TotalSkus = g.Count()
                    })
                    .OrderByDescending(f => f.KilosProd)
                    .ToList();

                return Ok(new
                {
                    TotalProdKilos = Math.Round(totalProd, 2),
                    TotalPyKilos = Math.Round(totalPy, 2),
                    DiferenciaNetaKilos = Math.Round(diferenciaNeta, 2),
                    PctCumplimientoGlobal = pctCumplimientoGlobal,
                    SkusTotales = skusTotales,
                    SkusCoincidentes = skusAmbos,
                    SkusSoloProduccion = skusSoloProd,
                    SkusSoloDemanda = skusSoloDem,
                    TopSobrecumplimiento = topSobrecumplimiento,
                    TopSubcumplimiento = topSubcumplimiento,
                    ResumenSubPlantas = resumenSubPlantas,
                    ResumenFamilias = resumenFamilias
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }

        /// <summary>
        /// Análisis a profundidad exclusivo para T_ResProd_Py:
        /// Kilos producidos, desglose por SubPlanta, desglose por Familia, y Top 10 SKUs en Kilos.
        /// </summary>
        [HttpGet("analisis-produccion")]
        public IActionResult ObtenerAnalisisProduccion([FromQuery] string? semana = null, [FromQuery] string? subPlanta = null)
        {
            try
            {
                GenerarResumenConsolidado();
                var items = _cacheProdItems ?? new List<ResProdPyItemDto>();

                if (!string.IsNullOrWhiteSpace(semana))
                {
                    var semFilter = semana.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                          .Select(s => s.Trim().ToUpper())
                                          .ToHashSet();
                    items = items.Where(r => semFilter.Contains(r.Semana.ToUpper())).ToList();
                }

                if (!string.IsNullOrWhiteSpace(subPlanta))
                {
                    var spFilter = subPlanta.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                            .Select(s => s.Trim().ToUpper())
                                            .ToHashSet();
                    items = items.Where(r => spFilter.Contains(r.SubPlanta.ToUpper())).ToList();
                }

                double totalKilos = Math.Round(items.Sum(x => x.Kilos), 2);
                int totalRegistros = items.Count;
                int totalSkus = items.Select(x => x.CodigoCorto).Distinct().Count();

                // 1. Desglose por SubPlanta
                var porSubPlanta = items
                    .GroupBy(x => string.IsNullOrWhiteSpace(x.SubPlanta) ? "DESPOSTE" : x.SubPlanta.ToUpper())
                    .Select(g =>
                    {
                        double k = g.Sum(x => x.Kilos);
                        return new
                        {
                            SubPlanta = g.Key,
                            Kilos = Math.Round(k, 2),
                            PctKilos = totalKilos > 0 ? Math.Round((k / totalKilos) * 100, 2) : 0,
                            TotalRegistros = g.Count(),
                            TotalSkus = g.Select(x => x.CodigoCorto).Distinct().Count()
                        };
                    })
                    .OrderByDescending(x => x.Kilos)
                    .ToList();

                // 2. Desglose por Familia
                var porFamilia = items
                    .GroupBy(x => x.Familia)
                    .Select(g =>
                    {
                        double k = g.Sum(x => x.Kilos);
                        return new
                        {
                            Familia = g.Key,
                            Kilos = Math.Round(k, 2),
                            PctKilos = totalKilos > 0 ? Math.Round((k / totalKilos) * 100, 2) : 0,
                            TotalSkus = g.Select(x => x.CodigoCorto).Distinct().Count()
                        };
                    })
                    .OrderByDescending(x => x.Kilos)
                    .ToList();

                // 3. Top 10 SKUs en Kilos
                var topSkusKilos = items
                    .GroupBy(x => new { x.CodigoCorto, x.Nombre, x.SubPlanta, x.Familia, x.Categoria })
                    .Select(g => new
                    {
                        CodigoCorto = g.Key.CodigoCorto,
                        Nombre = g.Key.Nombre,
                        SubPlanta = g.Key.SubPlanta,
                        Familia = g.Key.Familia,
                        Categoria = g.Key.Categoria,
                        Kilos = Math.Round(g.Sum(x => x.Kilos), 2)
                    })
                    .OrderByDescending(x => x.Kilos)
                    .Take(10)
                    .ToList();

                // 4. Evolución Semanal
                var evolucionSemanal = items
                    .GroupBy(x => new { x.Semana, x.Anio, x.NSemana })
                    .OrderBy(g => g.Key.Semana)
                    .Select(g => new
                    {
                        Semana = g.Key.Semana,
                        Anio = g.Key.Anio,
                        NSemana = g.Key.NSemana,
                        Kilos = Math.Round(g.Sum(x => x.Kilos), 2)
                    })
                    .ToList();

                return Ok(new
                {
                    TotalKilos = totalKilos,
                    TotalRegistros = totalRegistros,
                    TotalSkus = totalSkus,
                    PorSubPlanta = porSubPlanta,
                    PorFamilia = porFamilia,
                    TopSkusKilos = topSkusKilos,
                    EvolucionSemanal = evolucionSemanal
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }

        /// <summary>
        /// Análisis a profundidad exclusivo para T_DemandaPy:
        /// Kilos demandados, desglose por canal u Origen (COMERCIAL, EMBUTIDOS, WESTPHALIA),
        /// desglose por SubPlanta, desglose por Familia y Top SKUs por cada canal de origen.
        /// </summary>
        [HttpGet("analisis-demanda")]
        public IActionResult ObtenerAnalisisDemanda([FromQuery] string? semana = null, [FromQuery] string? subPlanta = null)
        {
            try
            {
                GenerarResumenConsolidado();
                var items = _cacheDemandaItems ?? new List<DemandaPyItemDto>();

                if (!string.IsNullOrWhiteSpace(semana))
                {
                    var semFilter = semana.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                          .Select(s => s.Trim().ToUpper())
                                          .ToHashSet();
                    items = items.Where(r => semFilter.Contains(r.Semana.ToUpper())).ToList();
                }

                if (!string.IsNullOrWhiteSpace(subPlanta))
                {
                    var spFilter = subPlanta.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                            .Select(s => s.Trim().ToUpper())
                                            .ToHashSet();
                    items = items.Where(r => spFilter.Contains(r.SubPlanta.ToUpper())).ToList();
                }

                double totalKilos = Math.Round(items.Sum(x => x.Kilos), 2);
                int totalRegistros = items.Count;
                int totalSkus = items.Select(x => x.CodigoCorto).Distinct().Count();
                var origenesDistintos = items.Select(x => x.Origen).Distinct().ToList();

                // 1. Desglose por Origen (COMERCIAL, EMBUTIDOS, WESTPHALIA)
                var porOrigen = items
                    .GroupBy(x => x.Origen)
                    .Select(g =>
                    {
                        double k = g.Sum(x => x.Kilos);
                        return new
                        {
                            Origen = g.Key,
                            Kilos = Math.Round(k, 2),
                            PctParticipacion = totalKilos > 0 ? Math.Round((k / totalKilos) * 100, 2) : 0,
                            TotalRegistros = g.Count(),
                            TotalSkus = g.Select(x => x.CodigoCorto).Distinct().Count()
                        };
                    })
                    .OrderByDescending(x => x.Kilos)
                    .ToList();

                // 2. Desglose por SubPlanta
                var porSubPlanta = items
                    .GroupBy(x => string.IsNullOrWhiteSpace(x.SubPlanta) ? "DESPOSTE" : x.SubPlanta.ToUpper())
                    .Select(g =>
                    {
                        double k = g.Sum(x => x.Kilos);
                        return new
                        {
                            SubPlanta = g.Key,
                            Kilos = Math.Round(k, 2),
                            PctParticipacion = totalKilos > 0 ? Math.Round((k / totalKilos) * 100, 2) : 0,
                            TotalSkus = g.Select(x => x.CodigoCorto).Distinct().Count()
                        };
                    })
                    .OrderByDescending(x => x.Kilos)
                    .ToList();

                // 3. Desglose por Familia
                var porFamilia = items
                    .GroupBy(x => x.Familia)
                    .Select(g =>
                    {
                        double k = g.Sum(x => x.Kilos);
                        return new
                        {
                            Familia = g.Key,
                            Kilos = Math.Round(k, 2),
                            PctParticipacion = totalKilos > 0 ? Math.Round((k / totalKilos) * 100, 2) : 0,
                            TotalSkus = g.Select(x => x.CodigoCorto).Distinct().Count()
                        };
                    })
                    .OrderByDescending(x => x.Kilos)
                    .ToList();

                // 4. Top 10 SKUs en Kilos Global Demanda
                var topSkusGlobal = items
                    .GroupBy(x => new { x.CodigoCorto, x.Nombre, x.SubPlanta, x.Familia, x.Categoria })
                    .Select(g => new
                    {
                        CodigoCorto = g.Key.CodigoCorto,
                        Nombre = g.Key.Nombre,
                        SubPlanta = g.Key.SubPlanta,
                        Familia = g.Key.Familia,
                        Categoria = g.Key.Categoria,
                        Kilos = Math.Round(g.Sum(x => x.Kilos), 2)
                    })
                    .OrderByDescending(x => x.Kilos)
                    .Take(10)
                    .ToList();

                // 5. Top 5 SKUs por cada Origen
                var topPorOrigen = new Dictionary<string, object>();
                foreach (var orig in origenesDistintos)
                {
                    var topO = items
                        .Where(x => x.Origen == orig)
                        .GroupBy(x => new { x.CodigoCorto, x.Nombre, x.SubPlanta, x.Familia })
                        .Select(g => new
                        {
                            CodigoCorto = g.Key.CodigoCorto,
                            Nombre = g.Key.Nombre,
                            SubPlanta = g.Key.SubPlanta,
                            Familia = g.Key.Familia,
                            Kilos = Math.Round(g.Sum(x => x.Kilos), 2)
                        })
                        .OrderByDescending(x => x.Kilos)
                        .Take(5)
                        .ToList();

                    topPorOrigen[orig] = topO;
                }

                // 6. Evolución Semanal de Demanda
                var evolucionSemanal = items
                    .GroupBy(x => new { x.Semana, x.Anio, x.NSemana })
                    .OrderBy(g => g.Key.Semana)
                    .Select(g => new
                    {
                        Semana = g.Key.Semana,
                        Anio = g.Key.Anio,
                        NSemana = g.Key.NSemana,
                        Kilos = Math.Round(g.Sum(x => x.Kilos), 2),
                        TotalSkus = g.Select(x => x.CodigoCorto).Distinct().Count()
                    })
                    .ToList();

                return Ok(new
                {
                    TotalKilos = totalKilos,
                    TotalRegistros = totalRegistros,
                    TotalSkus = totalSkus,
                    TotalOrigenes = origenesDistintos.Count,
                    PorOrigen = porOrigen,
                    PorSubPlanta = porSubPlanta,
                    PorFamilia = porFamilia,
                    TopSkusGlobal = topSkusGlobal,
                    TopPorOrigen = topPorOrigen,
                    EvolucionSemanal = evolucionSemanal
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }

        /// <summary>
        /// Obtiene los registros individuales de T_ResProd_Py
        /// </summary>
        [HttpGet("produccion-items")]
        public IActionResult ObtenerProduccionItems([FromQuery] string? semana = null, [FromQuery] string? subPlanta = null)
        {
            try
            {
                GenerarResumenConsolidado();
                var items = _cacheProdItems ?? new List<ResProdPyItemDto>();

                if (!string.IsNullOrWhiteSpace(semana))
                {
                    var semFilter = semana.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                          .Select(s => s.Trim().ToUpper())
                                          .ToHashSet();
                    items = items.Where(r => semFilter.Contains(r.Semana.ToUpper())).ToList();
                }

                if (!string.IsNullOrWhiteSpace(subPlanta))
                {
                    var spFilter = subPlanta.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                            .Select(s => s.Trim().ToUpper())
                                            .ToHashSet();
                    items = items.Where(r => spFilter.Contains(r.SubPlanta.ToUpper())).ToList();
                }

                return Ok(items);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }

        /// <summary>
        /// Obtiene los registros individuales de T_DemandaPy
        /// </summary>
        [HttpGet("demanda-items")]
        public IActionResult ObtenerDemandaItems([FromQuery] string? semana = null, [FromQuery] string? subPlanta = null)
        {
            try
            {
                GenerarResumenConsolidado();
                var items = _cacheDemandaItems ?? new List<DemandaPyItemDto>();

                if (!string.IsNullOrWhiteSpace(semana))
                {
                    var semFilter = semana.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                          .Select(s => s.Trim().ToUpper())
                                          .ToHashSet();
                    items = items.Where(r => semFilter.Contains(r.Semana.ToUpper())).ToList();
                }

                if (!string.IsNullOrWhiteSpace(subPlanta))
                {
                    var spFilter = subPlanta.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                            .Select(s => s.Trim().ToUpper())
                                            .ToHashSet();
                    items = items.Where(r => spFilter.Contains(r.SubPlanta.ToUpper())).ToList();
                }

                return Ok(items);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }

        /// <summary>
        /// Matriz T_ResProd_Py: Gráfico de barras por SubPlanta (Kilos) y Matriz jerárquica SubPlanta -> Familia -> CodigoCorto -> Nombre (Kilos)
        /// </summary>
        [HttpGet("matriz-produccion-tipo")]
        [HttpGet("matriz-produccion-subplanta")]
        public IActionResult ObtenerMatrizProduccionSubPlanta([FromQuery] string? semana = null, [FromQuery] string? subPlanta = null)
        {
            try
            {
                GenerarResumenConsolidado();
                var items = _cacheProdItems ?? new List<ResProdPyItemDto>();

                if (!string.IsNullOrWhiteSpace(semana))
                {
                    var semFilter = semana.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                          .Select(s => s.Trim().ToUpper())
                                          .ToHashSet();
                    items = items.Where(r => semFilter.Contains(r.Semana.ToUpper())).ToList();
                }

                if (!string.IsNullOrWhiteSpace(subPlanta))
                {
                    var spFilter = subPlanta.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                            .Select(s => s.Trim().ToUpper())
                                            .ToHashSet();
                    items = items.Where(r => spFilter.Contains(r.SubPlanta.ToUpper())).ToList();
                }

                double totalKilosGlobal = items.Sum(x => x.Kilos);

                // Gráfico de barras por SubPlanta
                var barras = items
                    .GroupBy(x => string.IsNullOrWhiteSpace(x.SubPlanta) ? "DESPOSTE" : x.SubPlanta.ToUpper())
                    .Select(g =>
                    {
                        double k = Math.Round(g.Sum(x => x.Kilos), 2);
                        return new
                        {
                            SubPlanta = g.Key,
                            Kilos = k,
                            PctKilos = totalKilosGlobal > 0 ? Math.Round((k / totalKilosGlobal) * 100.0, 2) : 0.0,
                            TotalSkus = g.Select(x => x.CodigoCorto).Distinct().Count()
                        };
                    })
                    .OrderByDescending(x => x.Kilos)
                    .ToList();

                // Matriz jerárquica: SubPlanta -> Familia -> Categoria -> CodigoCorto -> Nombre
                var matriz = items
                    .GroupBy(x => string.IsNullOrWhiteSpace(x.SubPlanta) ? "DESPOSTE" : x.SubPlanta.ToUpper())
                    .OrderByDescending(gSp => gSp.Sum(x => x.Kilos))
                    .Select(gSp => new
                    {
                        SubPlanta = gSp.Key,
                        Kilos = Math.Round(gSp.Sum(x => x.Kilos), 2),
                        Familias = gSp
                            .GroupBy(x => string.IsNullOrWhiteSpace(x.Familia) ? "SIN FAMILIA" : x.Familia)
                            .OrderBy(gFam => gFam.Key)
                            .Select(gFam => new
                            {
                                Familia = gFam.Key,
                                Kilos = Math.Round(gFam.Sum(x => x.Kilos), 2),
                                Categorias = gFam
                                    .GroupBy(x => string.IsNullOrWhiteSpace(x.Categoria) ? "SIN CATEGORIA" : x.Categoria)
                                    .OrderBy(gCat => gCat.Key)
                                    .Select(gCat => new
                                    {
                                        Categoria = gCat.Key,
                                        Kilos = Math.Round(gCat.Sum(x => x.Kilos), 2),
                                        Articulos = gCat
                                            .GroupBy(x => new { x.CodigoCorto, x.Nombre, x.Producto })
                                            .OrderBy(gArt => gArt.Key.CodigoCorto)
                                            .Select(gArt => new
                                            {
                                                CodigoCorto = gArt.Key.CodigoCorto,
                                                Nombre = string.IsNullOrWhiteSpace(gArt.Key.Nombre) ? (gArt.Key.Producto ?? "SIN NOMBRE") : gArt.Key.Nombre,
                                                Producto = gArt.Key.Producto,
                                                Kilos = Math.Round(gArt.Sum(x => x.Kilos), 2)
                                            }).ToList()
                                    }).ToList()
                            }).ToList()
                    }).ToList();

                return Ok(new
                {
                    TotalKilos = Math.Round(totalKilosGlobal, 2),
                    Barras = barras,
                    Matriz = matriz
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }

        /// <summary>
        /// Matriz T_DemandaPy: Gráfico de barras por Origen (Kilos) y Matriz jerárquica Origen -> SubPlanta -> Familia -> Categoria -> CodigoCorto -> Nombre (Kilos)
        /// </summary>
        [HttpGet("matriz-demanda-origen")]
        public IActionResult ObtenerMatrizDemandaOrigen([FromQuery] string? semana = null, [FromQuery] string? subPlanta = null)
        {
            try
            {
                GenerarResumenConsolidado();
                var items = _cacheDemandaItems ?? new List<DemandaPyItemDto>();

                if (!string.IsNullOrWhiteSpace(semana))
                {
                    var semFilter = semana.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                          .Select(s => s.Trim().ToUpper())
                                          .ToHashSet();
                    items = items.Where(r => semFilter.Contains(r.Semana.ToUpper())).ToList();
                }

                if (!string.IsNullOrWhiteSpace(subPlanta))
                {
                    var spFilter = subPlanta.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                            .Select(s => s.Trim().ToUpper())
                                            .ToHashSet();
                    items = items.Where(r => spFilter.Contains(r.SubPlanta.ToUpper())).ToList();
                }

                double totalKilosGlobal = items.Sum(x => x.Kilos);

                // Gráfico de barras por Origen
                var barras = items
                    .GroupBy(x => string.IsNullOrWhiteSpace(x.Origen) ? "OTROS" : x.Origen.ToUpper())
                    .Select(g =>
                    {
                        double k = Math.Round(g.Sum(x => x.Kilos), 2);
                        return new
                        {
                            Origen = g.Key,
                            Kilos = k,
                            PctKilos = totalKilosGlobal > 0 ? Math.Round((k / totalKilosGlobal) * 100.0, 2) : 0.0,
                            TotalSkus = g.Select(x => x.CodigoCorto).Distinct().Count()
                        };
                    })
                    .OrderByDescending(x => x.Kilos)
                    .ToList();

                // Matriz jerárquica: Origen -> SubPlanta -> Familia -> Categoria -> CodigoCorto -> Nombre
                var matriz = items
                    .GroupBy(x => string.IsNullOrWhiteSpace(x.Origen) ? "OTROS" : x.Origen.ToUpper())
                    .OrderByDescending(gOrig => gOrig.Sum(x => x.Kilos))
                    .Select(gOrig => new
                    {
                        Origen = gOrig.Key,
                        Kilos = Math.Round(gOrig.Sum(x => x.Kilos), 2),
                        SubPlantas = gOrig
                            .GroupBy(x => string.IsNullOrWhiteSpace(x.SubPlanta) ? "DESPOSTE" : x.SubPlanta.ToUpper())
                            .OrderBy(gSp => gSp.Key)
                            .Select(gSp => new
                            {
                                SubPlanta = gSp.Key,
                                Kilos = Math.Round(gSp.Sum(x => x.Kilos), 2),
                                Familias = gSp
                                    .GroupBy(x => string.IsNullOrWhiteSpace(x.Familia) ? "SIN FAMILIA" : x.Familia)
                                    .OrderBy(gFam => gFam.Key)
                                    .Select(gFam => new
                                    {
                                        Familia = gFam.Key,
                                        Kilos = Math.Round(gFam.Sum(x => x.Kilos), 2),
                                        Categorias = gFam
                                            .GroupBy(x => string.IsNullOrWhiteSpace(x.Categoria) ? "SIN CATEGORIA" : x.Categoria)
                                            .OrderBy(gCat => gCat.Key)
                                            .Select(gCat => new
                                            {
                                                Categoria = gCat.Key,
                                                Kilos = Math.Round(gCat.Sum(x => x.Kilos), 2),
                                                Articulos = gCat
                                                    .GroupBy(x => new { x.CodigoCorto, x.Nombre, x.Producto })
                                                    .OrderBy(gArt => gArt.Key.CodigoCorto)
                                                    .Select(gArt => new
                                                    {
                                                        CodigoCorto = gArt.Key.CodigoCorto,
                                                        Nombre = string.IsNullOrWhiteSpace(gArt.Key.Nombre) ? (gArt.Key.Producto ?? "SIN NOMBRE") : gArt.Key.Nombre,
                                                        Producto = gArt.Key.Producto,
                                                        Kilos = Math.Round(gArt.Sum(x => x.Kilos), 2)
                                                    }).ToList()
                                            }).ToList()
                                    }).ToList()
                            }).ToList()
                    }).ToList();

                return Ok(new
                {
                    TotalKilos = Math.Round(totalKilosGlobal, 2),
                    Barras = barras,
                    Matriz = matriz
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { exito = false, error = ex.Message });
            }
        }
    }

    public class CumplimientoPyResumenDto
    {
        public int Anio { get; set; }
        public int NSemana { get; set; }
        public string Semana { get; set; } = string.Empty;
        public string SubPlanta { get; set; } = string.Empty;
        public string CodigoCorto { get; set; } = string.Empty;
        public string Producto { get; set; } = string.Empty;
        public double KilosProd { get; set; }
        public double KilosPy { get; set; }
        public double Diferencia { get; set; }
        public double PctVariacion { get; set; }
        public double PctCumplimiento { get; set; }
        public string Linea { get; set; } = string.Empty;
        public string Familia { get; set; } = string.Empty;
        public string Categoria { get; set; } = string.Empty;
        public string Nombre { get; set; } = string.Empty;
    }

    public class ResProdPyItemDto
    {
        public string Semana { get; set; } = string.Empty;
        public int Anio { get; set; }
        public int NSemana { get; set; }
        public string SubPlanta { get; set; } = string.Empty;
        public string Producto { get; set; } = string.Empty;
        public string CodigoCorto { get; set; } = string.Empty;
        public string Nombre { get; set; } = string.Empty;
        public double Kilos { get; set; }
        public string Linea { get; set; } = string.Empty;
        public string Familia { get; set; } = string.Empty;
        public string Categoria { get; set; } = string.Empty;
    }

    public class DemandaPyItemDto
    {
        public string Semana { get; set; } = string.Empty;
        public int Anio { get; set; }
        public int NSemana { get; set; }
        public string SubPlanta { get; set; } = string.Empty;
        public string CodigoCorto { get; set; } = string.Empty;
        public string Producto { get; set; } = string.Empty;
        public string Nombre { get; set; } = string.Empty;
        public double Kilos { get; set; }
        public string Origen { get; set; } = string.Empty;
        public string Linea { get; set; } = string.Empty;
        public string Familia { get; set; } = string.Empty;
        public string Categoria { get; set; } = string.Empty;
    }

    internal class CodigoRelacionPyAuxDto
    {
        public string CodigoCorto { get; set; } = string.Empty;
        public string CodigoArticulo { get; set; } = string.Empty;
        public string Nombre { get; set; } = string.Empty;
    }

    internal class ArticuloVentasPyAuxDto
    {
        public string CodigoVenta { get; set; } = string.Empty;
        public string Nombre { get; set; } = string.Empty;
        public string NLinea { get; set; } = string.Empty;
        public string NFamilia { get; set; } = string.Empty;
        public string NCategoria { get; set; } = string.Empty;
    }

    internal class AcumuladorProdPyDto
    {
        public string Semana { get; set; } = string.Empty;
        public string CodigoCorto { get; set; } = string.Empty;
        public string SubPlanta { get; set; } = string.Empty;
        public string Producto { get; set; } = string.Empty;
        public string Nombre { get; set; } = string.Empty;
        public string Linea { get; set; } = string.Empty;
        public string Familia { get; set; } = string.Empty;
        public string Categoria { get; set; } = string.Empty;
        public double KilosProd { get; set; }
    }

    internal class AcumuladorDemPyDto
    {
        public string Semana { get; set; } = string.Empty;
        public string CodigoCorto { get; set; } = string.Empty;
        public string SubPlanta { get; set; } = string.Empty;
        public string Producto { get; set; } = string.Empty;
        public string Nombre { get; set; } = string.Empty;
        public string Linea { get; set; } = string.Empty;
        public string Familia { get; set; } = string.Empty;
        public string Categoria { get; set; } = string.Empty;
        public double KilosPy { get; set; }
    }
}
