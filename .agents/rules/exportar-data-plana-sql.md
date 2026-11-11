# Regla: Botón de Exportación de Data Plana en Excel para Dashboards con Fuente SQL Server

Esta regla es de cumplimiento obligatorio para todos los dashboards HTML en `Docu_Paginas/` que se conecten a la API .NET y consuman datos provenientes de **SQL Server** (directamente o mediante procedimientos almacenados / vistas).

---

## 1. Dashboards Aplicables (Fuentes SQL Server)

Actualmente identificados en el proyecto:
1. `Docu_Paginas/InformeCabezas.html` (Controlador: `/api/ProduccionCabezas`, SP: `Sp_Informe_Prduccion_Cabezas`)
2. `Docu_Paginas/RendimientosChuletas002.html` (Controlador: `/api/ProdChuletas`, SP: `Sp_Informe_Prod_Chuletas`)
3. `Docu_Paginas/Indice_Rotacion_Semanal.html` (Controlador: `/api/IndiceRotacionSemanal`, SP: `Sp_Indice_Rotacion_Semanal`)
4. `Docu_Paginas/Informe_Costo_Produccion.html` (Controlador: `/api/CostosProduccion`, Tabla: `CostoProduccion`)
*(Cualquier nuevo dashboard que consuma un controlador en `ApiDesposte/Controllers/Sql/` deberá incluir obligatoriamente esta funcionalidad).*

---

## 2. Requerimientos Técnicos Obligatorios

### A. Carga de Librería SheetJS
Todo dashboard con fuente SQL debe incluir en su `<head>` la librería SheetJS para la generación de archivos Excel `.xlsx`:
```html
<!-- SheetJS para exportar a Excel (.xlsx) -->
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
```

### B. Preservar la Data Plana Original en Memoria
Al recibir la respuesta exitosa del `fetch` desde el controlador de SQL Server, se debe almacenar el arreglo crudo tal cual llega de la base de datos antes de cualquier mapeo, agrupación o transformación local:
```javascript
const json = await res.json();
const items = Array.isArray(json) ? json : (json.datos || []);
rawSqlData = items; // Almacenamiento íntegro de la data plana tal cual viene de SQL Server
```

### C. Botón de Exportación en la Barra de Herramientas o Cabecera
El dashboard debe exponer un botón visible, accesible y estilizado con la paleta oficial del proyecto, con icono descriptivo:
```html
<button class="btn-tool btn-excel-flat" id="btnExportarPlana" onclick="exportarDataPlanaSQL()">
  <span>📄</span> Data Plana SQL (.xlsx)
</button>
```
- Debe permanecer deshabilitado o mostrar validación preventiva cuando la consulta aún no se haya realizado o retorne 0 registros.

### D. Función Estándar de Descarga (`exportarDataPlanaSQL`)
La exportación debe generar un libro de trabajo Excel nativo con extensión `.xlsx` conteniendo todas las columnas y filas originales devueltas por la consulta SQL:
```javascript
function exportarDataPlanaSQL() {
  if (!rawSqlData || rawSqlData.length === 0) {
    mostrarToast("⚠️ No hay datos de SQL Server cargados para exportar", 3000);
    return;
  }
  
  mostrarToast("⏳ Generando Excel de Data Plana...", 2000);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rawSqlData);
  XLSX.utils.book_append_sheet(wb, ws, "Data_Plana_SQL");

  const fIni = document.getElementById('fechaInicio')?.value || '';
  const fFin = document.getElementById('fechaFin')?.value || '';
  const rangoStr = (fIni && fFin) 
    ? `_${fIni.replace(/-/g, '')}_al_${fFin.replace(/-/g, '')}` 
    : `_${new Date().toISOString().slice(0, 10).replace(/-/g, '')}`;

  const nombreArchivo = `Data_Plana_SQL_<Modulo>${rangoStr}.xlsx`;
  XLSX.writeFile(wb, nombreArchivo);
  mostrarToast(`📄 Data Plana SQL exportada (${rawSqlData.length.toLocaleString('es-PE')} filas)`, 3500);
}
```

---

## 3. Beneficios Operativos
- Permite a los usuarios y auditores descargar los registros exactos de SQL Server sin los filtros visuales o jerárquicos de la interfaz.
- Garantiza la reproducibilidad de los cálculos al contrastar la vista gráfica o agregada contra la fuente plana de origen.

