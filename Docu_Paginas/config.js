// =========================================================================
// ⚙️ CONFIGURACIÓN GLOBAL DE CONEXIÓN A LA API
// =========================================================================
// Este archivo permite cambiar la dirección del servidor para todos los reportes HTML.
// Si trabajas en la oficina, en casa o en local, solo cambia el valor de 'servidor'.

var GLOBAL_CONFIG = {
    // Nombre del equipo o IP donde corre tu API .NET
    // 🏢 Oficina: "RAEVSALL001"
    // 🏠 Casa / Local: "localhost" o el nombre de tu máquina
    servidor: "RAEVSALL001",

    // Puertos a los que intentará conectarse en orden de prioridad
    puertos: [8080, 5000]
};

// Función auxiliar global para generar las bases URL de la API
function generarBasesApiGlobal() {
    var s = (GLOBAL_CONFIG && GLOBAL_CONFIG.servidor) ? GLOBAL_CONFIG.servidor : "RAEVSALL001";
    var puertos = (GLOBAL_CONFIG && GLOBAL_CONFIG.puertos) ? GLOBAL_CONFIG.puertos : [8080, 5000];
    var bases = [];

    for (var i = 0; i < puertos.length; i++) {
        bases.push("http://" + s + ":" + puertos[i]);
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
