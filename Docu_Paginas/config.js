// =========================================================================
// ⚙️ CONFIGURACIÓN GLOBAL DE CONEXIÓN A LA API
// =========================================================================
// Este archivo permite cambiar la dirección del servidor para todos los reportes HTML.
// Si trabajas en la oficina, en casa o en local, solo cambia el valor de 'servidor'.

var GLOBAL_CONFIG = {
    // Nombre del equipo o IP donde corre tu API .NET
    // 🏢 Oficina: "RAEVSALL001"
    // 🏠 Casa / Local: "localhost" o el nombre de tu máquina
    //RAEVFINL016
    servidor: "RAEVFINL016",

    // Puertos a los que intentará conectarse en orden de prioridad
    puertos: [8080, 5000]
};

function generarBasesApiGlobal() {
    var s = (GLOBAL_CONFIG && GLOBAL_CONFIG.servidor) ? GLOBAL_CONFIG.servidor : "RAEVSALL001";
    var puertos = (GLOBAL_CONFIG && GLOBAL_CONFIG.puertos) ? GLOBAL_CONFIG.puertos : [8080, 5000];
    var bases = [];

    // Si la página se sirve por HTTP/HTTPS (ej. http://localhost:8080 o http://RAEVFINL016:8080), ese origen tiene máxima prioridad
    if (typeof window !== 'undefined' && window.location && window.location.origin && window.location.protocol.startsWith('http')) {
        bases.push(window.location.origin);
    }

    for (var i = 0; i < puertos.length; i++) {
        var baseServidor = "http://" + s + ":" + puertos[i];
        if (bases.indexOf(baseServidor) === -1) bases.push(baseServidor);
    }

    // Fallbacks si el servidor principal no es localhost
    if (s && s.toLowerCase() !== "localhost" && s !== "127.0.0.1") {
        for (var j = 0; j < puertos.length; j++) {
            var fallback = "http://localhost:" + puertos[j];
            if (bases.indexOf(fallback) === -1) bases.push(fallback);
        }
        for (var k = 0; k < puertos.length; k++) {
            var fallbackIp = "http://127.0.0.1:" + puertos[k];
            if (bases.indexOf(fallbackIp) === -1) bases.push(fallbackIp);
        }
    }
    return bases;
}

// =========================================================================
// 🛡️ GUARDIÁN DE ACCESO: BLOQUEO AUTOMÁTICO DE DASHBOARDS EN PRUEBA
// =========================================================================
(function verificarAccesoDashboard() {
    if (typeof window === 'undefined' || !window.location) return;

    var path = window.location.pathname || "";
    var fileName = path.substring(path.lastIndexOf('/') + 1) || "";
    if (!fileName || fileName.toLowerCase() === "index.html") return;

    // Verificar si el usuario actual cuenta con sesión activa de Administrador
    var esAdmin = false;
    try {
        if (sessionStorage.getItem('desposte_admin_pin')) esAdmin = true;
        if (document.cookie && document.cookie.indexOf('desposte_admin=') !== -1) esAdmin = true;
    } catch (e) { }

    if (esAdmin) return;

    function ejecutarBloqueo() {
        // Redirigir de inmediato al portal principal
        var rutaIndex = window.location.protocol.startsWith('http')
            ? "index.html?bloqueado=" + encodeURIComponent(fileName)
            : "index.html?bloqueado=" + encodeURIComponent(fileName);
        window.location.replace(rutaIndex);
    }

    try {
        // 1. Revisar si está deshabilitado en LocalStorage
        var local = localStorage.getItem('portal_dashboards_estado');
        if (local) {
            var estadosLocal = JSON.parse(local);
            if (estadosLocal && estadosLocal[fileName] === false) {
                ejecutarBloqueo();
                return;
            }
        }

        // 2. Revisar dashboards_estado.json
        fetch("dashboards_estado.json?_t=" + Date.now())
            .then(function (res) { return res.json(); })
            .then(function (estados) {
                if (estados && estados[fileName] === false) {
                    ejecutarBloqueo();
                }
            })
            .catch(function () { });
    } catch (err) { }
})();
