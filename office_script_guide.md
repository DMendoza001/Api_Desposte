# Guía para Usar la API de Desposte en Excel con Office Scripts

Esta guía te guiará paso a paso para configurar un **Office Script** en Microsoft Excel (de Escritorio o Web). Al ejecutar este script, se abrirá un formulario lateral pidiendo las fechas y luego se consultará la API para insertar los datos formateados automáticamente en la hoja activa.

---

## Código del Script (TypeScript)

Copia todo el código que se muestra a continuación:

```typescript
/**
 * Script de Excel para obtener atenciones de tópico desde la API de Desposte
 * utilizando un formulario lateral de parámetros.
 * 
 * @param fechaInicial Fecha de inicio en formato AAAA-MM-DD (ejemplo: 2026-08-01)
 * @param fechaFinal Fecha de fin en formato AAAA-MM-DD (ejemplo: 2026-08-26)
 */
async function main(
  workbook: ExcelScript.Workbook,
  fechaInicial: string,
  fechaFinal: string
) {
  const fIni = fechaInicial.trim();
  const fFin = fechaFinal.trim();

  if (!fIni || !fFin) {
    console.log("Error: Debes ingresar tanto la fecha inicial como la fecha final.");
    return;
  }

  // ==========================================
  // CONFIGURACIÓN DE LA URL DE LA API
  // ==========================================
  // IMPORTANTE PARA EXCEL WEB:
  // Excel en la Web requiere una conexión HTTPS segura debido a políticas del navegador.
  // - Si estás en Excel Web, debes exponer tu API local con HTTPS (ej. usando ngrok o devtunnel)
  //   o subir la API a un servidor en la nube con HTTPS.
  // - Si estás en Excel de Escritorio, puedes usar "http://localhost:5000" directamente.
  const API_BASE_URL = "http://localhost:5000/api/Desposte/atenciones-topico";
  
  const url = `${API_BASE_URL}?fechaInicial=${encodeURIComponent(fIni)}&fechaFinal=${encodeURIComponent(fFin)}`;

  console.log(`Realizando petición a: ${url}`);

  try {
    const response = await fetch(url);
    
    if (!response.ok) {
      throw new Error(`La API respondió con error: ${response.status} (${response.statusText})`);
    }

    const data: Atencion[] = await response.json();

    if (!data || data.length === 0) {
      console.log("No se encontraron registros para el rango de fechas seleccionado.");
      return;
    }

    // 1. Obtener la hoja activa
    let sheet = workbook.getActiveWorksheet();
    
    // 2. Limpiar los datos viejos de la hoja para evitar solapamientos
    let usedRange = sheet.getUsedRange();
    if (usedRange) {
      usedRange.clear(ExcelScript.ClearApplyTo.all);
    }

    // 3. Extraer las columnas basadas en el primer objeto devuelto por la API
    const headers = Object.keys(data[0]);

    // 4. Preparar la matriz de datos para insertar en Excel
    const matrix: (string | number | boolean)[][] = [];
    
    // Agregar cabecera como primera fila
    matrix.push(headers);

    // Agregar filas de datos
    for (const item of data) {
      const row: (string | number | boolean)[] = [];
      for (const key of headers) {
        let val = item[key];
        if (val === null || val === undefined) {
          row.push("");
        } else if (typeof val === "object") {
          row.push(JSON.stringify(val));
        } else {
          row.push(val);
        }
      }
      matrix.push(row);
    }

    // 5. Escribir los datos en la hoja de Excel
    const rowsCount = matrix.length;
    const colsCount = headers.length;
    const targetRange = sheet.getRangeByIndexes(0, 0, rowsCount, colsCount);
    targetRange.setValues(matrix);

    // 6. Formatear la tabla con un aspecto premium
    // - Encabezados en negrita con fondo azul oscuro y texto blanco
    const headerRange = sheet.getRangeByIndexes(0, 0, 1, colsCount);
    headerRange.getFormat().getFont().setBold(true);
    headerRange.getFormat().getFont().setColor("white");
    headerRange.getFormat().getFill().setColor("#1F4E78"); // Azul oscuro premium

    // - Ajustar automáticamente el ancho de las columnas
    sheet.getUsedRange().getFormat().autofitColumns();

    console.log(`Éxito: Se cargaron ${data.length} filas en la hoja.`);
  } catch (error) {
    console.error("Error al obtener o procesar los datos de la API:");
    console.error(error);
  }
}

interface Atencion {
  [key: string]: any;
}
```

---

## Paso a Paso para Instalar en Excel

### 1. Crear el Script
1. Abre tu archivo de Excel (ya sea en la versión de **Escritorio** o en **Excel Web**).
2. Ve a la pestaña **Automatizar** (Automate) en la cinta superior.
3. Haz clic en **Nuevo script** (New Script). Se abrirá un panel lateral derecho con el Editor de código.
4. Borra todo el código que viene por defecto en el editor.
5. Pega el código de arriba.
6. En la parte superior del panel del editor, haz clic sobre el nombre del script (por defecto se llama *Script 1*) y cámbialo a algo descriptivo, como `ObtenerAtencionesTopico`.
7. Haz clic en **Guardar script** (Save Script).

### 2. Ejecutar y Probar el Script
1. Tras guardar, aparecerá un botón **Ejecutar** (Run) en el mismo panel lateral.
2. Dado que el código define parámetros adicionales en la función `main` (`fechaInicial` y `fechaFinal`), verás que Excel dibuja automáticamente dos campos de texto titulados:
   - **fechaInicial**
   - **fechaFinal**
3. Introduce los valores con formato `AAAA-MM-DD` (por ejemplo, `2026-08-01` y `2026-08-26`).
4. Presiona el botón **Ejecutar** (Run).
5. Los datos se insertarán y formatearán automáticamente en tu hoja de Excel actual.

---

## ⚠️ Consideración Importante: HTTP vs HTTPS (Mixed Content)

Debido a las políticas de seguridad de los navegadores web modernos, cuando usas **Excel Web** (que funciona bajo `https://...` de Microsoft), no puedes realizar solicitudes HTTP normales a un servidor `http://localhost:5000` porque el navegador bloqueará la solicitud por considerarla **contenido mixto inseguro** (*Mixed Content*).

### ¿Cómo solucionarlo si usas Excel Web?

#### Opción A: Exponer tu API local de forma segura (Recomendado para pruebas rápidas)
Puedes usar una herramienta de túnel seguro como **ngrok** o **devtunnel** (que viene integrado con Visual Studio/dotnet cli) para obtener una URL pública HTTPS temporal que redirija a tu puerto local `5000`.

- **Con devtunnel (si tienes el SDK de .NET):**
  Abre una consola/PowerShell y ejecuta:
  ```powershell
  devtunnel loopback 5000 -p http
  ```
  Esto te dará una URL HTTPS como `https://xxxxx-5000.use.devtunnels.ms`. Copia esa dirección y actualiza el valor de `API_BASE_URL` en tu Office Script:
  ```typescript
  const API_BASE_URL = "https://xxxxx-5000.use.devtunnels.ms/api/Desposte/atenciones-topico";
  ```

- **Con ngrok:**
  Descarga ngrok y ejecuta en consola:
  ```bash
  ngrok http 5000
  ```
  Copia la URL HTTPS generada (ej. `https://xxxx.ngrok-free.app`) y colócala en tu script:
  ```typescript
  const API_BASE_URL = "https://xxxx.ngrok-free.app/api/Desposte/atenciones-topico";
  ```

#### Opción B: Usar Excel de Escritorio
En Excel de Escritorio para Windows, el motor que ejecuta Office Scripts tiene menos restricciones de contenido mixto para dominios de loopback local. En la mayoría de entornos, podrás usar `http://localhost:5000` directamente sin necesidad de crear un túnel HTTPS.
