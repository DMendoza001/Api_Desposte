# Reglas de Conexión a la API (Frontend / Informes HTML)

## 1. Detección Dinámica de Hosts
- Ningún archivo HTML o script frontend debe hardcodear nombres fijos de equipos (e.g. `DMENDOZA`, `RAEVSALL001`) para permitir la portabilidad total del proyecto entre diferentes máquinas.
- Se debe obtener el host actual dinámicamente mediante:
  1. `window.location.hostname` (si se accede por red LAN o servidor web).
  2. `localhost` y `127.0.0.1` como hosts locales estándar.
  3. Rutas relativas (`/api/...`) para cuando la página sea servida directamente por el backend.

## 2. Orden de Prioridad de Puertos
- **Prioridad 1:** Puerto `8080` (puerto preferente del servidor de desarrollo y producción).
- **Prioridad 2:** Puerto `5000` (puerto secundario / fallback estándar de Kestrel/ASP.NET Core).
- **Prioridad 3:** Ruta relativa directa (`""` o `/api/...`).

## 3. Manejo de Timeouts y Conexión Ágil
- Los intentos de detección y sondeo de puertos deben tener un timeout corto (entre 1.5 y 3 segundos con `AbortController`) para que, en caso de que el puerto 8080 esté ocupado o apagado, la conexión pase inmediatamente al puerto 5000 sin demorar la experiencia del usuario.

## 4. Patrón Estándar de Generación de Bases / Endpoints
```javascript
function generarBasesApi() {
    const hosts = [];
    if (typeof window !== "undefined" && window.location && window.location.hostname && window.location.hostname !== "") {
        hosts.push(window.location.hostname);
    }
    hosts.push("localhost", "127.0.0.1");

    const puertos = [8080, 5000];
    const bases = [];
    for (const puerto of puertos) {
        for (const host of hosts) {
            bases.push(`http://${host}:${puerto}`);
        }
    }
    bases.push(""); // Relativo
    return Array.from(new Set(bases));
}
```
