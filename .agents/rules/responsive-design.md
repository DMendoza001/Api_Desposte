# Regla General: Diseño Responsive Híbrido (Laptops, Tablets y Celulares)

Todos los archivos `.html` ubicados en `Docu_Paginas/` (o cualquier nuevo reporte/dashboard web que se cree en el proyecto) **deben implementar obligatoriamente el estándar de diseño responsive híbrido** para garantizar su correcta visualización tanto en monitores grandes como en **laptops (13"–15.6" con escala de Windows 100%, 125% y 150%)**, **tablets** y **teléfonos celulares**.

---

## 1. Principio del Patrón Responsive Híbrido

Al igual que `config.js`, el diseño responsive opera de forma dual:
1. **Hoja de Estilos Compartida (`responsive.css`):** Todo archivo `.html` en `Docu_Paginas/` debe enlazar en el `<head>` el archivo compartido `<link rel="stylesheet" href="responsive.css?v=3">`.
2. **Autonomía Individual (Inline `@media` en `<style>`):** Todo archivo `.html` debe contar también dentro de su propio `<style>` con las reglas esenciales de adaptación (`min-width: 0` en hijos de CSS Grid/Flexbox y `@media` queries para `1600px`, `1024px`/`980px` y `768px`) por si el archivo `.html` se comparte individualmente sin `responsive.css`.

---

## 2. Ubicación en el `<head>` del Documento HTML

```html
<!-- Configuración compartida de conexión a la API y diseño responsive -->
<link rel="stylesheet" href="responsive.css?v=3">
<script src="config.js"></script>
```

---

## 3. Reglas Técnicas Obligatorias de Maquetación (Prevención de Errores en Laptop y Móvil)

1. **Prohibido usar alturas fijas inline sin clase en contenedores de gráficos (`<canvas>`):**
   - **Nunca** envolver un `<canvas>` de Chart.js en un `<div style="height:430px;">` sin clase CSS.
   - Todo `<canvas>` debe ir dentro de un contenedor con clase estándar (`.chart-box`, `.chart-container` o `.chart-wrapper`) que tenga obligatoriamente `position: relative; width: 100%; min-width: 0;`, permitiendo que `responsive.css` y las reglas `@media` reduzcan su altura en laptop (`280px`–`290px`) y celular (`220px`–`240px`) y que Chart.js redimensione el lienzo sin desbordar.
2. **Prevención de desbordamiento en CSS Grid y Flexbox (`min-width: 0`):**
   - Todo contenedor principal (`.dashboard-container > *`, `.kpi-grid > *`, `.card`, `.kpi-card`, `.chart-card`, `.matrix-card`) tanto en **CSS Grid** como en **Flexbox (`flex-direction: column`)** debe tener `min-width: 0;` para evitar que un `<canvas>` o una tabla ancha de múltiples columnas expanda el ancho de la página más allá de la pantalla.
3. **Alineación de Tarjetas KPI en Laptops (Prohibido `minmax` mayores a `200px`):**
   - **Nunca** usar `minmax(260px, 1fr)` en `.kpi-grid`, ya que en laptops con escala de Windows al 125% o 150% provoca que la 4ta tarjeta KPI caiga sola a una segunda fila.
   - En laptop/escritorio (`> 980px`), mantener las 4 tarjetas KPI alineadas en una sola fila usando `grid-template-columns: repeat(4, minmax(0, 1fr))` o `repeat(auto-fit, minmax(185px, 1fr))`. Pasar a `repeat(2, 1fr)` recién en tablets (`<= 980px`) y a `1fr` en celulares (`<= 768px`).
4. **Densidad compacta automática en Laptops (`@media screen and (max-width: 1600px)`):**
   - El breakpoint para laptops debe ser **`<= 1600px`** (y no solo `1440px`), porque una laptop Full HD (`1920x1080`) con escala de Windows al **125%** tiene un viewport CSS de **`1536px`**.
   - En `<= 1600px` se deben compactar paddings de `body` (`10px 14px`), tarjetas (`12px 16px`), separadores de filtros (`margin: 8px 0`), altura de gráficos (`280px`–`290px`) y limitar la altura de matrices largas (`.table-responsive { max-height: 460px; overflow-y: auto; }`) para mantener visibles los encabezados `sticky` y evitar scroll excesivo.
5. **Adaptación Móvil (`@media screen and (max-width: 768px)`):**
   - Filtros (`select`, `input`, `.week-combobox-wrapper`, botones) al `100%` de ancho (`min-width: 0`).
   - En tablas con primera columna fija (`.sticky-col` o `th:first-child / td:first-child`), desactivar o limitar en móvil su ancho (`min-width: 130px; max-width: 200px; white-space: normal;`) para permitir deslizar las columnas numéricas cómodamente.
