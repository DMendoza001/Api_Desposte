# Regla General: Estándar de Conexión a la API (Frontend HTML)

Todos los archivos `.html` ubicados en `Docu_Paginas/` (o cualquier nuevo reporte/tablero web que se cree en el proyecto) **deben implementar obligatoriamente el patrón de conexión híbrida estándar** detallado en este documento.

---

## 1. Principio del Patrón Híbrido

El sistema debe operar de forma dual sin requerir cambios manuales ni romper dependencias:
1. **Desarrollo y Portabilidad Local:** Si el archivo compartido `config.js` existe en la misma carpeta, la página toma dinámicamente el host y puertos configurados en `GLOBAL_CONFIG`.
2. **Archivos Compartidos Autónomos:** Si se envía un archivo `.html` individual a un compañero o usuario final (sin `config.js`), el archivo debe funcionar de forma 100% autónoma usando los valores por defecto (`RAEVSALL001`, puertos `[8080, 5000]`).

---

## 2. Bloque Estándar Obligatorio

Todo archivo HTML que consuma datos de la API .NET debe incluir exactamente este bloque antes de su lógica de negocio:

```html
<!-- Configuración compartida de conexión a la API -->
<script src="config.js"></script>

<script>
    // =========================================================================
    // ⚙️ CONFIGURACIÓN DE CONEXIÓN A LA API (Híbrida: usa config.js si existe)
    // =========================================================================
    const APP_CONFIG = {
        servidor: (typeof GLOBAL_CONFIG !== 'undefined' && GLOBAL_CONFIG.servidor) 
                  ? GLOBAL_CONFIG.servidor 
                  : "RAEVSALL001",
        puertos: (typeof GLOBAL_CONFIG !== 'undefined' && GLOBAL_CONFIG.puertos) 
                 ? GLOBAL_CONFIG.puertos 
                 : [8080, 5000]
    };

    // Generación de servidores API apuntando al nombre del equipo
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
```

---

## 3. Extensiones Específicas por Página

Cualquier propiedad o endpoint adicional que requiera la página debe derivarse a partir del bloque base:

### Caso A: Detección por Lista de Hosts (`hosts`)
Para páginas que iteran buscando el host activo (e.g. `CumplimientosPTC.html`, `CumplimientosPTC_Py.html`):
```javascript
APP_CONFIG.hosts = generarBasesApi();
APP_CONFIG.controllerPath = "/api/NombreDelControlador";
```

### Caso B: Generador de Endpoints Relativos (`generarEndpointsApi`)
Para páginas que consultan directamente rutas relativas (e.g. `Informe_Costo_Produccion.html`, `RendimientosChuletas.html`):
```javascript
APP_CONFIG.rutaListar = "/api/Controlador/listar";
APP_CONFIG.timeoutMs = 20000;

function generarEndpointsApi(rutaRelativa) {
    return generarBasesApi().map(base => `${base}${rutaRelativa}`);
}
```

---

## 4. Archivo Central `config.js`

El archivo `Docu_Paginas/config.js` define:
- `servidor`: Nombre de equipo en red (e.g. `"RAEVSALL001"`) o `"localhost"`.
- `puertos`: Lista priorizada de puertos `[8080, 5000]`.
- Función `generarBasesApiGlobal()` con detección y fallbacks a `localhost` y `127.0.0.1`.
