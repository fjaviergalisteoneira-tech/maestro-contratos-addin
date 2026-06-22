# Maestro de Contratos — complemento de Excel

Panel lateral (Office Add-in) que explora los proyectos del **maestro de contratos**.
Lee los datos **en local** de la hoja `Maestro Contratos` del libro abierto usando Office.js;
ningún dato se envía a internet. GitHub Pages solo aloja la interfaz (HTML/JS/CSS).

## Estructura
- `taskpane.html` / `taskpane.css` / `taskpane.js` — el panel lateral.
- `commands.html` / `commands.js` — archivo de funciones (botón de la cinta).
- `manifest.xml` — manifiesto del complemento (apunta a las URLs de GitHub Pages).
- `assets/` — iconos.

## Instalación en Excel (una vez)
1. Abrir `maestro_contratos.xlsx` en Excel de escritorio.
2. **Insertar → Complementos → Mis complementos → Cargar mi complemento**.
3. Elegir `manifest.xml`.
4. Pestaña **Inicio → Explorador de proyectos**.

## Datos
El panel busca la hoja `Maestro Contratos` y mapea las columnas por su cabecera
(robusto a reordenaciones). Botón ↻ para recargar tras editar el libro.
