# Cotizaciones de Valorización

Herramienta web para registrar cotizaciones de valorización/eliminación de
residuos (Apartados 1–5 según la hoja INSTRUCCIONES de
`BBDD_cotizaciones_v1.xlsx`), con actualización automática de la base de
destinatarios y un panel de historial.

## 1. Requisitos

- Node.js 18 o superior
- Una cuenta de GitHub

## 2. Probar en tu computador

```bash
npm install
npm run dev
```

Abre la URL que muestra la terminal (normalmente `http://localhost:5173`).

## 3. Subir a GitHub

```bash
git init
git add .
git commit -m "Primera versión de la plataforma de cotizaciones"
git branch -M main
git remote add origin https://github.com/TU-USUARIO/TU-REPO.git
git push -u origin main
```

## 4. Publicar en GitHub Pages (automático)

Este repo ya trae un workflow (`.github/workflows/deploy.yml`) que compila y
publica la app cada vez que haces push a `main`. Solo falta:

1. En `vite.config.js`, cambia `base: '/NOMBRE-DEL-REPO/'` por el nombre real
   de tu repositorio (tiene que calzar exacto, con las mayúsculas/minúsculas).
2. En GitHub → **Settings → Pages**, en "Build and deployment" selecciona
   **GitHub Actions** como fuente.
3. Haz push. En la pestaña **Actions** puedes ver el progreso; cuando termine,
   tu app queda en `https://TU-USUARIO.github.io/TU-REPO/`.

Si prefieres desplegar a mano en vez de con Actions:

```bash
npm run build
npx gh-pages -d dist
```

## 5. Actualizar el catálogo de residuos o destinatarios

Los datos de las hojas `RESIDUOS`, `LISTAS` y `DESTINATARIOS` de tu Excel
están precompilados en `src/data.js` para que la app cargue rápido y no
dependa de tener el Excel disponible en el navegador. Si esas hojas cambian
(nuevos materiales, nuevas comunas, etc.), regenera ese archivo:

```bash
pip install openpyxl
python scripts/extract_data.py ruta/a/tu/nuevo/BBDD_cotizaciones.xlsx
git add src/data.js
git commit -m "Actualiza catálogo de residuos y destinatarios"
git push
```

Los **destinatarios nuevos que se agregan usando el formulario** (no desde
el Excel) NO tocan este archivo — quedan guardados aparte, según se explica
abajo.

## 6. Dónde se guardan los datos

Los datos viven en un **Google Sheet**, al que la app accede mediante un
**Google Apps Script** publicado como aplicación web (el código está en
`apps-script/Codigo.gs`).

Eso significa que la base es compartida: si alguien registra una cotización
o un destinatario nuevo, el resto del equipo lo ve de inmediato.

El Sheet tiene tres hojas:

| Hoja | Para qué sirve |
|---|---|
| `COTIZACIONES` | Recibe cada cotización nueva. Parte vacía: el histórico vive en el Excel |
| `DESTINATARIOS` | Fuente viva de establecimientos; crece desde el formulario |
| `USUARIOS` | Usuario, PIN y estado ACTIVO de cada persona |

### Si cambia la URL del Apps Script

Cada vez que crees una **implementación nueva** del script, la URL cambia.
Cuando pase, actualiza `API_URL` en `src/api.js` y vuelve a publicar.

Para cambios en el código del script que NO deban cambiar la URL, usa
**Implementar → Gestionar implementaciones → editar (lápiz) → Versión: Nueva
versión**. Así la URL se mantiene.

### Numeración de cotizaciones

El correlativo lo asigna el servidor, no el navegador, y usa un bloqueo para
que dos personas guardando al mismo tiempo no reciban el mismo ID. Continúa
desde `COT-01667`, que es la siguiente al último registro del Excel histórico
(`COT-01666`). Ese piso está en la constante `ULTIMO_ID_HISTORICO` del script.

## 7. Usuarios y acceso

El login valida usuario y PIN **en el servidor**, contra la hoja `USUARIOS`.
Los PIN nunca se envían al navegador ni aparecen en ninguna respuesta.

Para administrar el acceso, edita esa hoja directamente:

- **Dar de alta**: agrega una fila con usuario, PIN y ACTIVO = `SI`
- **Revocar**: cambia ACTIVO a `NO` (mejor que borrar la fila: conserva el historial)
- **Cambiar un PIN**: edítalo en la celda; surte efecto de inmediato

Alcance real de esta protección: la URL del servicio es pública, así que el
PIN es la única barrera. Sirve para dejar registro de quién ingresó cada
cotización y para impedir que un tercero cualquiera escriba en la base, pero
no equivale a una autenticación corporativa. Por eso conviene usar PIN de 6
o más dígitos sin patrones obvios, y compartir el Sheet solo con quien deba
administrarlo — los usuarios de la app no necesitan acceso al Sheet.

## 8. Alimentar el Excel desde el Sheet (Power Query)

El Excel sigue siendo el maestro del histórico, y el Sheet aporta lo nuevo.
Para unirlos:

1. En el Sheet: **Archivo → Compartir → Publicar en la web**, elige la hoja
   `COTIZACIONES` y formato **CSV**. Copia la URL.
2. En Excel: **Datos → Obtener datos → Desde otras fuentes → Desde la web**,
   pega esa URL.
3. Haz *Anexar consultas* entre la hoja `COTIZACIONES` histórica y la nueva
   consulta del Sheet.

Como el Sheet parte vacío y el Excel conserva hasta `COT-01666`, no hay
riesgo de duplicados.

## Estructura del proyecto

```
├── src/
│   ├── App.jsx          # Lógica y pantallas (login, formulario, historial)
│   ├── data.js          # Catálogo de residuos y listas (generado, no editar a mano)
│   ├── api.js           # Conexión con el Google Sheet — acá va la URL del Apps Script
│   └── main.jsx
├── apps-script/
│   └── Codigo.gs        # Código que se pega en Google Apps Script
├── scripts/
│   └── extract_data.py  # Regenera src/data.js desde un Excel actualizado
└── .github/workflows/deploy.yml  # Publica la app en GitHub Pages en cada push a main
```
