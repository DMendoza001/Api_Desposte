# Regla General: Filtros de Periodo por Vista Temporal

Todo dashboard en `Docu_Paginas/` que tenga vistas temporales **Diario / Semanal / Mensual** debe ofrecer un filtro de periodo con **selección múltiple** y un valor **por defecto** según la vista.

Implementación de referencia: `Docu_Paginas/RendimientosChuletas002.html`.

---

## 1. Reglas por vista

| Vista | Qué se filtra | Selección | Por defecto |
|---|---|---|---|
| **Diario** | Semanas | Múltiple | Las **2 últimas semanas** |
| **Semanal** | Semanas | Múltiple | Las **12 últimas semanas** |
| **Mensual** | Años | Múltiple | El **último año** |

---

## 2. Criterios

- Las semanas se calculan con ISO-8601: empiezan el **lunes** y terminan el domingo (ver `calculo-semanas.md`).
- "Últimas N" se cuenta desde el periodo más reciente que tenga datos en la consulta.
- Solo se pueden elegir periodos que existan en el rango de fechas consultado.
- Cada vista guarda su propia selección: cambiar de vista no altera la selección de las demás.
- Una nueva consulta o "Reestablecer" devuelve las tres vistas a su selección por defecto.
- Si se desmarcan todos los periodos, se muestran todos los disponibles.
