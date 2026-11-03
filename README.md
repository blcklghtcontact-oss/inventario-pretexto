# Inventario Café Pretexto

App de inventario de la barra de Café Pretexto, hecha para el celular. Tiene la misma libreta del recetario.

- **Libreta**: existencias, precios, presentación (galón, bote…) y mejores días para pedir cada cosa.
- **Del día**: lo que se va usando (− 250 ml, − 1 L, “se acabó”), la venta de caja del día y lo que llegó.
- **Conteo**: contar envases sin abrir y cuánto queda en el que está en uso (¼, ½, ¾ o ml/g exactos).
- **Pedir**: lo que falta por proveedor, cantidades por unidad, total del pedido completo y envío por WhatsApp.
- **Gráficas**: vendido, compras, insumos usados, consumo por día de la semana y por producto.
- **Avisos**: semana de pedidos (martes a domingo; el lunes está cerrado) y recordatorios con alarma en Google Calendar.

## Dónde se guardan los datos

Todo se guarda en una **hoja de Google** de la cuenta del café. Cada cosa que se anota queda como una fila con fecha, hora, día de la semana, cantidad, costo y quién lo anotó. Así queda el historial día por día para ver cuánto se gasta de cada cosa y planear mejor las compras.

Pestañas que crea la hoja:

| Pestaña | Qué tiene |
|---|---|
| Productos | El inventario actual |
| Movimientos | Cada entrada, uso, merma y conteo (una fila por registro) |
| Ventas | La venta de caja de cada día |
| Reportes | Lista de reportes creados en Google Docs |
| Proveedores, Recordatorios, Config | Lo demás de la app |
| Consumo diario, Consumo por producto, Compras por día, Consumo por día de la semana | Resúmenes que se calculan solos |

## Reportes y Obsidian

- En **Gráficas → Reportes en Google Docs** se crea un documento con el dinero, el consumo por producto, lo de cada día y lo que llegó. Los reportes quedan listados ahí para abrirlos después, y el botón **Ver la hoja de datos** abre la hoja completa.
- Cada noche se copia el día a la bóveda de Obsidian (Segundo Cerebro), en `wiki/cafeteria-pretexto/inventario/`.

## Instalación (una sola vez, unos 10 minutos)

1. **Crea la hoja.** En Google Drive: Nuevo → Hojas de cálculo de Google. Ponle de nombre “Inventario Café Pretexto”.
2. **Pega el código.** En la hoja: Extensiones → Apps Script. Borra lo que aparece y pega todo el contenido de [`apps-script/Codigo.gs`](apps-script/Codigo.gs). Guarda (ícono del disquete).
3. **Configura.** Vuelve a la hoja y recárgala. Aparece el menú **Inventario** → **1. Configurar (primera vez)**. Google pedirá permisos (es tu propio código). Escribe la **clave del equipo**: la que todos van a usar para entrar a la app. Se crean las pestañas y se cargan los 25 productos de la barra.
4. **Publica la base de datos.** En Apps Script: **Implementar → Nueva implementación** → tipo **Aplicación web**.
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier usuario**

   Copia la dirección que termina en `/exec`.
5. **Conecta la app.** Abre la app en el teléfono y pega la dirección donde dice “Conectar la base de datos”. Para que ya no la pida en ningún teléfono, pégala también en `config.js` (`SCRIPT_URL`).

Listo. Abre la app, escribe tu nombre y la clave del equipo.

> La clave evita que alguien que encuentre la dirección pueda escribir en la hoja. Se cambia desde la hoja: Inventario → Cambiar clave del equipo.

Si después cambias algo en `Codigo.gs`, publica de nuevo con **Implementar → Administrar implementaciones → Editar → Nueva versión** para que la dirección siga siendo la misma.

## Ponerla en el teléfono como app

- iPhone: abrir en Safari → Compartir → **Agregar a inicio**.
- Android: abrir en Chrome → menú ⋮ → **Agregar a pantalla principal**.

## Archivos

- `index.html` — la app completa.
- `config.js` — dirección de la base de datos.
- `apps-script/Codigo.gs` — código de la hoja de Google.
- `manifest.webmanifest`, `icono-*.png`, `logo-pretexto.png` — ícono y datos para instalarla en el teléfono.
