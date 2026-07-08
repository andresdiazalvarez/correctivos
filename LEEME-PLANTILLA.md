# Base Extintor

Esta carpeta es una base limpia para crear otra app independiente basada en la app Extintor.

## Qué incluye

- App web/PWA para GitHub Pages.
- Guardado local en el móvil/navegador mediante IndexedDB.
- Importar Excel.
- Importar registros repetidos sin descartarlos.
- Tabla con filtros y ordenación.
- Formulario para meter datos nuevos y corregir registros.
- Fotos 1 y 2 con eliminar/cambiar.
- Defectos seleccionables.
- Campo Cliente.
- Descargar Excel con datos, defectos y fotos.

## Datos iniciales

La app viene sin datos iniciales:

```js
window.INITIAL_EXTINTORES_LISTADOS = [];
```

Los datos se cargan importando un Excel desde la app.

## Para crear otra app

1. Crear un repositorio nuevo en GitHub.
2. Subir todos los archivos de esta carpeta.
3. Activar GitHub Pages.
4. Cambiar el nombre visible de la app si hace falta:
   - `index.html`
   - `manifest.webmanifest`
   - `README.md`

## Archivos principales

- `index.html`: pantallas y formulario.
- `app.js`: lógica, importación, guardado y exportación.
- `styles.css`: diseño.
- `sw.js`: caché de la app instalada.
- `manifest.webmanifest`: nombre/icono de instalación.
- `exceljs.min.js`: generación y lectura de Excel.
