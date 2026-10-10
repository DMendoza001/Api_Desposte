# Regla General: Criterio de Cálculo de Semanas (Inicio en Lunes)

En todo el proyecto (`Api_Desposte`), tanto en los controladores de la API .NET (C#) como en los dashboards y reportes analíticos (`Docu_Paginas/*.html` en JavaScript), **el primer día de la semana debe ser obligatoriamente el día LUNES** (y terminar en **DOMINGO**).

---

## 1. Principio Fundamental
- **Inicio de Semana:** Lunes.
- **Fin de Semana:** Domingo.
- **Estándar:** ISO-8601 / Regla empresarial corporativa (semanas completas de Lunes a Domingo).
- Nunca se debe calcular semanas iniciando en Domingo.

---

## 2. Implementación Estándar en JavaScript (Frontend HTML)

Todo frontend o dashboard que calcule o agrupe por semanas debe implementar este algoritmo estándar:

```javascript
/**
 * Calcula la semana del año y el rango de fechas con inicio estricto en LUNES.
 * @param {Date} d - Objeto Date a evaluar
 * @returns {{ semana: number, anoSemana: number, lunesFecha: string, rangoTexto: string }}
 */
function calcularSemanaConInicioLunes(d) {
    const copia = new Date(d.getFullYear(), d.getMonth(), d.getDate());

    // Con LUNES como día 0 (offset) y DOMINGO como día 6:
    const diaOffsetLunes = (copia.getDay() + 6) % 7;

    // Lunes de la semana (inicio)
    const lunes = new Date(copia);
    lunes.setDate(copia.getDate() - diaOffsetLunes);

    // Domingo de la semana (cierre)
    const domingo = new Date(lunes);
    domingo.setDate(lunes.getDate() + 6);

    // Cálculo ISO-8601 estándar (Semana inicia en LUNES)
    const dateUtc = new Date(Date.UTC(copia.getFullYear(), copia.getMonth(), copia.getDate()));
    const diaUtc = dateUtc.getUTCDay() === 0 ? 7 : dateUtc.getUTCDay(); // Lunes=1, Domingo=7
    dateUtc.setUTCDate(dateUtc.getUTCDate() + 4 - diaUtc); // Ajustar al jueves de la semana
    const anoSemana = dateUtc.getUTCFullYear();
    const inicioAno = new Date(Date.UTC(anoSemana, 0, 1));
    const numSemana = Math.ceil((((dateUtc - inicioAno) / 86400000) + 1) / 7);

    const pad = n => String(n).padStart(2, '0');
    const lunesStr = `${lunes.getFullYear()}-${pad(lunes.getMonth() + 1)}-${pad(lunes.getDate())}`;
    const rangoTexto = `${pad(lunes.getDate())}/${pad(lunes.getMonth() + 1)} al ${pad(domingo.getDate())}/${pad(domingo.getMonth() + 1)}`;

    return {
        semana: numSemana,
        anoSemana: anoSemana,
        lunesFecha: lunesStr,
        rangoTexto: rangoTexto
    };
}
```

### Formato de Presentación Visual:
En gráficos y tablas semanales se debe mostrar el rango para evitar ambigüedades:
```javascript
`Sem. ${String(semana).padStart(2, '0')} (${rangoTexto})`
// Ejemplo: "Sem. 41 (05/10 al 11/10)"
```

---

## 3. Implementación Estándar en C# / .NET (Backend)

En el backend, al procesar fechas en C# o LINQ:

```csharp
using System.Globalization;

// Opción recomendada (.NET Core / .NET 6+):
int numSemana = ISOWeek.GetWeekOfYear(fecha);
int anoSemana = ISOWeek.GetYear(fecha);

// O mediante CultureInfo / Calendar:
CultureInfo ci = CultureInfo.CurrentCulture;
int numSemana = ci.Calendar.GetWeekOfYear(
    fecha, 
    CalendarWeekRule.FirstFourDayWeek, 
    DayOfWeek.Monday
);
```
