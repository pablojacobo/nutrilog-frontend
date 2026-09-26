# Flujo de desarrollo

Este repositorio tiene **dos líneas de desarrollo independientes** que **nunca se mezclan**.

---

## 1. Rama `main` — AppScript tradicional (producción actual)

- **Backend**: Google Apps Script (`Code.gs`).
- **Frontend**: `Index.html` embebido en el proyecto de Apps Script.
- **Despliegue**: Aplicación web de Apps Script (Implementar > Nueva implementación).
- **Estado**: Producción. Tus usuarios actuales usan esta versión.

### Flujo de trabajo

```bash
# Modificar Code.gs o Index.html
git add Code.gs Index.html
git commit -m "Nuevo cambio en la versión AppScript"
git push origin main
```

Después de cada cambio, recordá de crear una nueva versión en el editor de Apps Script
(**Implementar > Administrar implementaciones > Editar > Nueva versión**).

---

## 2. Rama `pages-frontend` — Frontend en GitHub Pages + Apps Script como backend

- **Backend**: Igual que `main` (`Code.gs`), pero expuesto vía `doPost` en lugar de `google.script.run`.
- **Frontend**: `Index.html` modificado para usar `fetch` a la URL de Apps Script.
- **Despliegue**: GitHub Pages (Settings > Pages) + AppScript desplegado como aplicación web.
- **Estado**: En desarrollo. Sirve para probar la nueva estrategia sin afectar a `main`.

### Flujo de trabajo

```bash
git checkout pages-frontend
# Modificar Index.html y Code.gs
git add Index.html Code.gs
git commit -m "Nuevo cambio en la versión Pages"
git push origin pages-frontend
```

GitHub Pages se actualiza automáticamente desde esta rama. Para que el frontend funcione, completá
`APPSCRIPT_WEB_URL` en `Index.html` con la URL de tu aplicación web de Apps Script.

---

## Reglas

- **Nunca merges entre `main` y `pages-frontend`**. Cada rama es independiente.
- `main` es la línea de producción. Los cambios allí son para usuarios reales.
- `pages-frontend` es una línea de prueba. Los cambios allí son para probar la nueva arquitectura.
- Si un cambio debe ir a ambos lados, hacelo en cada rama por separado.