# Regla General: Diseño Responsive Híbrido (Laptops, Tablets y Celulares)

Todos los archivos `.html` ubicados en `Docu_Paginas/` (o cualquier nuevo reporte/dashboard web que se cree en el proyecto) **deben implementar obligatoriamente el estándar de diseño responsive híbrido** para garantizar su correcta visualización tanto en monitores grandes como en **laptops (13"–15" con escala de Windows 125%–150%)**, **tablets** y **teléfonos celulares**.

---

## 1. Principio del Patrón Responsive Híbrido

Al igual que `config.js`, el diseño responsive opera de forma dual:
1. **Hoja de Estilos Compartida (`responsive.css`):** Todo archivo `.html` en `Docu_Paginas/` debe enlazar en el `<head>` el archivo compartido `<link rel="stylesheet" href="responsive.css">`.
2. **Autonomía Individual (Inline `@media` en `<style>`):** Todo archivo `.html` debe contar también dentro de su propio `<style>` con las reglas esenciales de adaptación (`min-width: 0` en hijos de CSS Grid y `@media` queries para `1440px`, `1150px`, `768px` y `480px`) por si el archivo `.html` se comparte individualmente sin `responsive.css`.

---

## 2. Ubicación en el `<head>` del Documento HTML

```html
<!-- Configuración compartida de conexión a la API y diseño responsive -->
<link rel="stylesheet" href="responsive.css">
<script src="config.js"></script>
```

---

## 3. Reglas Técnicas Obligatorias de Maquetación

1. **Prevención de desbordamiento en CSS Grid (`min-width: 0`):**
   - Todo contenedor tarjeta (`.card`, `.kpi-card`, `.chart-card`, `.matrix-card`) dentro de un CSS Grid debe tener `min-width: 0;` para evitar que `<canvas>` (Chart.js) o tablas anchas expandan la columna más allá del ancho de la pantalla.
2. **Grillas fluidas en lugar de columnas rígidas:**
   - Para KPIs usar `grid-template-columns: repeat(auto-fit, minmax(200px, 1fr))` en vez de `repeat(4, 1fr)` fijo.
   - Si se usa `.full-width { grid-column: span 2; }`, en pantallas `<= 1150px` cambiar obligatoriamente a `.full-width { grid-column: 1 / -1; }` cuando la grilla pase a `1fr`.
3. **Densidad compacta automática en Laptops (`@media (max-width: 1440px)`):**
   - Reducir paddings de `body` (`14px 18px`), tarjetas (`14px 16px`), barras de filtros (`12px 16px`) y altura de gráficos (`280px`–`300px`) para que en laptops con escala de Windows al 125% o 150% se visualice el dashboard completo sin cortes.
4. **Adaptación Móvil (`@media (max-width: 768px)`):**
   - Filtros (`select`, `input`, `.week-combobox-wrapper`, botones) al `100%` de ancho (`min-width: 0`).
   - En tablas con primera columna fija (`.sticky-col` o `th:first-child / td:first-child`), limitar en móvil su ancho a `min-width: 145px; max-width: 170px; white-space: normal;` para que la columna fija no ocupe toda la pantalla del celular y permita deslizar las columnas numéricas cómodamente.

