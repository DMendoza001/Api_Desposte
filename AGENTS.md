# Api_Desposte - Guía de Reglas del Proyecto

Este repositorio contiene la API de Desposte en .NET y el conjunto de dashboards/reportes analíticos en HTML interactivos ubicados en `Docu_Paginas/`.

## Reglas Generales de Desarrollo

### 1. Conexión Frontend -> API .NET (Regla Obligatoria)
Todos los archivos `.html` deben seguir el estándar híbrido de conexión establecido en [.agents/rules/api-connection.md](file:///.agents/rules/api-connection.md):
- Importar `config.js` opcionalmente: `<script src="config.js"></script>`.
- Declarar el bloque estándar `APP_CONFIG` con fallback a `RAEVSALL001` y puertos `[8080, 5000]`.
- Utilizar `generarBasesApi()` para resolver URLs de la API.
- Nunca hardcodear URLs fijas directamente en llamadas `fetch` sin pasar por `generarBasesApi()`.
