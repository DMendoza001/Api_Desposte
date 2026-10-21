# Regla General: Estándar de Conexión a la API (Frontend HTML)

Todos los archivos `.html` ubicados en `Docu_Paginas/` (o cualquier nuevo reporte/tablero web que se cree en el proyecto) **deben implementar obligatoriamente el patrón de conexión híbrida estándar y resiliente** detallado en este documento.

---

## 1. Principio del Patrón Híbrido

El sistema debe operar de forma dual sin requerir cambios manuales ni romper dependencias:
1. **Desarrollo y Portabilidad Local:** Si el archivo compartido `config.js` existe en la misma carpeta, la página toma dinámicamente el host y puertos configurados en `GLOBAL_CONFIG`.
2. **Archivos Compartidos Autónomos / Red Local:** Si se envía un archivo `.html` individual a un compañero o se abre desde otra máquina de la red (incluso sin `config.js`), el archivo debe funcionar de forma 100% autónoma usando los valores por defecto del servidor central (`RAEVSALL001`, puertos `[8080, 5000]`).

---

## 2. Ubicación en el Documento HTML

- **En `<head>`:** Incluir únicamente la importación opcional de `config.js` sin bloques `<script>` inline con lógica:
  ```html
  <!-- Configuración compartida de conexión a la API -->
  <script src="config.js"></script>
  ```
- **En `<body>` (Inicio del script principal):** Declarar el bloque de configuración `APP_CONFIG`, `generarBasesApi()`, `fetchConTimeout()` y `detectarServidorApi()`.

---

## 3. Bloque Estándar Obligatorio

Todo dashboard o reporte HTML debe incluir la siguiente estructura exacta (ejemplo con `APP_CONFIG.controllerPath` adaptado a la página):

```javascript
// =========================================================================
// ⚙️ CONFIGURACIÓN DE CONEXIÓN HÍBRIDA A LA API .NET
// =========================================================================
const APP_CONFIG = {
    servidor: (typeof GLOBAL_CONFIG !== 'undefined' && GLOBAL_CONFIG.servidor)
        ? GLOBAL_CONFIG.servidor
        : "RAEVSALL001",
    puertos: (typeof GLOBAL_CONFIG !== 'undefined' && GLOBAL_CONFIG.puertos)
        ? GLOBAL_CONFIG.puertos
        : [8080, 5000]
};

function generarBasesApi() {
    if (typeof generarBasesApiGlobal === 'function') {
        return generarBasesApiGlobal();
    }
    const bases = [];
    for (const puerto of APP_CONFIG.puertos) {
        bases.push(`http://${APP_CONFIG.servidor}:${puerto}`);
    }
    return bases;
}

APP_CONFIG.controllerPath = "/api/NombreDelControlador";
APP_CONFIG.timeoutMs = 25000;

// Estado global de conexión
let apiBaseActiva = null;

// =========================================================================
// 🌐 UTILIDADES DE RED Y DETECCIÓN RESILIENTE
// =========================================================================

// Fetch con AbortController y timeout configurable
async function fetchConTimeout(url, timeoutMs = 10000) {
    const controller = new AbortController();
    const id = setTimeout(() => controller.abort(), timeoutMs);
    try {
        const res = await fetch(url, { signal: controller.signal });
        clearTimeout(id);
        return res;
    } catch (err) {
        clearTimeout(id);
        throw err;
    }
}

// Detección automática del host API activo (con fallback a ping raíz)
async function detectarServidorApi() {
    const bases = generarBasesApi();
    for (const base of bases) {
        try {
            // Intento 1: Consultar el endpoint del controlador con timeout de 3.5s
            const res = await fetchConTimeout(`${base}${APP_CONFIG.controllerPath}`, 3500);
            if (res.ok) {
                return base;
            }
        } catch (e) {
            // Intento 2 (Fallback): Ping rápido a la raíz '/' por si el endpoint demora leyendo datos pesados (Excel/BD)
            try {
                const pingRes = await fetchConTimeout(`${base}/`, 2000);
                if (pingRes.ok) return base;
            } catch (errPing) {
                // Siguiente candidato
            }
        }
    }
    return null;
}
```

---

## 4. Reglas Críticas para Evitar Fallos en Red / Otras Máquinas

1. **Nunca usar timeouts menores a 3 segundos en red:** En máquinas remotas o en red local corporativa, la resolución NetBIOS/DNS (`RAEVSALL001`) y el handshake TCP pueden demorar 500ms - 1s. Un timeout muy bajo (ej. 2s) aborta prematuramente conexiones legítimas.
2. **Fallback a ping raíz (`/`):** Si un endpoint tarda en su primer llamado debido a lectura de archivos Excel (ClosedXML) o consultas SQL pesadas, el ping a `${base}/` responde inmediatamente (HTTP 200), evitando descartar incorrectamente el servidor activo.
3. **Petición de datos con timeout suficiente:** La carga final de datos (`cargarDatosDesdeApi`) debe usar `fetchConTimeout(url, APP_CONFIG.timeoutMs)` con al menos 15-30 segundos según el volumen de datos.
4. **Mensaje de error explícito:** Si `apiBaseActiva` es `null`, lanzar un error claro indicando el servidor y puertos:
   ```javascript
   throw new Error(`No se pudo conectar a ${APP_CONFIG.servidor} en puertos ${APP_CONFIG.puertos.join('/')}`);
   ```
5. **No hardcodear URLs fijas:** Jamás invocar directamente `fetch("http://localhost:8080/...")` ni `fetch("http://RAEVSALL001:8080/...")` fuera del flujo `apiBaseActiva` / `generarBasesApi()`.

---

## 5. Archivo Central `config.js`

El archivo `Docu_Paginas/config.js` define:
- `servidor`: Nombre de equipo en red (e.g. `"RAEVSALL001"`) o `"localhost"`.
- `puertos`: Lista priorizada de puertos `[8080, 5000]`.
- Función `generarBasesApiGlobal()` con detección y fallbacks automáticos a `localhost` y `127.0.0.1`.
