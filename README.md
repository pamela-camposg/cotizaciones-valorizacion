# Cotizaciones de Valorización

Herramienta web para registrar cotizaciones de valorización y eliminación de
residuos. El equipo cotiza desde la app, todo queda en un Google Sheet, y el
Excel refleja ese contenido para analizarlo.

---

## Cómo está armado

```
   app (GitHub Pages)  ──escribe y lee──►  Google Sheet  ──lee──►  Excel
```

**El Google Sheet es el único maestro.** Ahí viven los datos y la
configuración: el catálogo de residuos, las listas de selección, los
destinatarios, los usuarios y las cotizaciones.

**La app no lleva catálogo propio.** Lo pide al Sheet cada vez que alguien
inicia sesión, así que editar una celda allá cambia de inmediato lo que
ofrece el formulario, sin tocar código ni volver a publicar.

**El Excel solo lee.** Son consultas de Power Query que traen tres hojas del
Sheet. Lo que se escriba en esas hojas del Excel se pierde al actualizar.

La app llega al Sheet a través de un **Google Apps Script** publicado como
aplicación web. Una página estática no puede tocar Drive por su cuenta: esa
pieza es el puente, y su código está en `apps-script/Codigo.gs`.

---

## Las hojas del Google Sheet

| Hoja | Qué contiene | Quién la edita |
|---|---|---|
| `COTIZACIONES` | Todas: el histórico y lo que entra por la app | La app escribe; la revisión marca VALIDACION |
| `RESIDUOS` | Catálogo de residuos y sus códigos | A mano, en el Sheet |
| `LISTAS` | Comunas, regiones, unidades, tipo de servicio, canal, tipo de negocio | A mano, en el Sheet |
| `DESTINATARIOS` | Clientes y establecimientos | A mano, y la app agrega los nuevos |
| `USUARIOS` | Usuario, PIN, nombre y si está activo | A mano |
| `NOTAS` | Apuntes personales | — |

---

## Tareas frecuentes

### Agregar un residuo, una comuna o una unidad

Se edita la hoja correspondiente del Sheet. **No hay que tocar el código.**
El cambio se ve la próxima vez que alguien entre a la app.

En `RESIDUOS`, las columnas de código arman el `COD_RES` con sus seis
segmentos, en este orden: peligrosidad, familia, material, tipo, estado y
formato. `PE` es peligroso y `NP` no peligroso.

### Dar de alta o revocar un usuario

En la hoja `USUARIOS`:

- **Alta**: agregar una fila con usuario, PIN, nombre y `ACTIVO = Si`
- **Revocar**: cambiar `ACTIVO` a `No`. Es mejor que borrar la fila, porque
  conserva el historial de quién registró cada cotización
- **Cambiar un PIN**: se edita la celda y surte efecto de inmediato

Solo los que tienen `ACTIVO = Si` aparecen en el desplegable del login.

### Revisar cotizaciones

La columna `VALIDACION` de `COTIZACIONES` acepta `Si` o `No`. La app graba
`Si` al guardar; en la revisión se cambia a `No` lo que se rechace.

Las marcadas `No` **desaparecen del Excel** pero siguen en el Sheet y en el
historial de la app, donde se ven atenuadas y con una marca. Se ocultan del
Excel para que no entren en los análisis, y se dejan visibles en la app para
que nadie concluya que una cotización nunca se registró.

---

## Mantener el código

### Cambiar la app

Los archivos de `src/` se suben a GitHub y Actions publica solo. En
`vite.config.js`, `base` tiene que calzar con el nombre del repositorio.

Para probar en el computador hace falta Node.js 18 o superior:

```bash
npm install
npm run dev
```

### Cambiar el Apps Script

Se pega el código en [script.google.com](https://script.google.com), se
guarda, y después:

**Implementar → Gestionar implementaciones → lápiz → Versión: Nueva versión**

Esto es importante: **guardar no basta**, la URL publicada sigue corriendo el
código anterior hasta que se implementa una versión nueva.

Y conviene usar siempre *Nueva versión*, no *Nueva implementación*. La
segunda genera una URL distinta, y habría que actualizarla en `src/api.js` y
en las tres consultas de Power Query.

### Las consultas de Power Query

Cada una trae una hoja del Sheet a través del script. La dirección va
completa en una sola línea, incluyendo la clave de lectura:

```
.../exec?accion=csv&clave=LA-CLAVE&hoja=COTIZACIONES
```

Armarla con el registro `Query` de Power Query **no funciona**: codifica los
parámetros a su manera y la clave llega distinta al servidor.

La de `COTIZACIONES` además filtra las rechazadas. Para revisarlas, se cambia
ese `<> "NO"` por `= "NO"`.

El refresco es manual: **Datos → Actualizar todo**, con el Excel abierto en
el escritorio. Excel Online no puede actualizar consultas web sin un gateway
de datos.

---

## Numeración de cotizaciones

El correlativo lo asigna el servidor, nunca el navegador, y usa un bloqueo
para que dos personas guardando al mismo tiempo no reciban el mismo ID.

Sale del mayor `COT-` que haya en la hoja `COTIZACIONES`. Por eso **borrar
filas del final es riesgoso**: el número queda libre y se vuelve a usar. Para
descartar una cotización, se marca `VALIDACION = No` en vez de borrarla.

Cuando una cotización tiene varios residuos, cada línea recibe su propio
código, y los códigos quedan consecutivos.

---

## Límites conocidos

**El PIN no es seguridad real.** La URL del servicio es pública, así que el
PIN es la única barrera. Sirve para dejar registro de quién ingresó cada
cotización y para impedir que un tercero cualquiera escriba en la base, pero
no equivale a una autenticación corporativa. Conviene usar PIN de 6 o más
dígitos sin patrones, y compartir el Sheet solo con quien lo administre: los
usuarios de la app no necesitan acceso.

**Sin conexión al Sheet la app no funciona**, ni para consultar. El catálogo
vive allá, así que no hay nada que mostrar si no se puede llegar.

**El historial de la app trae las 300 más recientes.** El resto se consulta
en el Excel, que es donde conviene analizar. Ese número está en
`COTIZACIONES_RECIENTES`, dentro del script.

**Un destinatario con dos establecimientos en la misma comuna y el mismo
sistema de declaración no se puede distinguir** en el formulario: la app
elige el primero. Hoy pasa con tres casos.

---

## Estructura del proyecto

```
├── src/
│   ├── App.jsx          # Pantallas y lógica: login, formulario, historial
│   ├── api.js           # Conexión con el Sheet — acá va la URL del Apps Script
│   ├── styles.css       # Estilos, incluida la versión para celular
│   └── main.jsx
├── apps-script/
│   └── Codigo.gs        # Se pega en Google Apps Script
├── index.html
├── vite.config.js       # 'base' debe calzar con el nombre del repositorio
└── .github/workflows/deploy.yml   # Publica en GitHub Pages con cada push a main
```
