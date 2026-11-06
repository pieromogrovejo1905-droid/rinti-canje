> ## 👉 Ingreso único: **[rinti.holafree.com](https://rinti.holafree.com/)** (Promotores · Back Office · Equipo Rinti)
> Dashboard del cliente: **[rinti.holafree.com/dashboard.html](https://rinti.holafree.com/dashboard.html)**
> App de impulsadoras: **[rinti.holafree.com/app.html](https://rinti.holafree.com/app.html)**
> Dashboard interno Free: **[rinti.holafree.com/interno.html](https://rinti.holafree.com/interno.html)** (usuarios de supervisor o coordinador)
>
> Esta página es solo el código. El dashboard se abre en el link de arriba, con usuario y contraseña.

# Canje Rinti · Registro de campo y dashboard

Sistema para la campaña **Canje Rinti (Ricocan / Ricocat)**: 9 impulsadoras en 8 mercados + Mercado Productores.
Mecánica: el cliente compra **N sacos (Ricocan 22 kg y/o Ricocat 15 kg, se pueden sumar)** y recibe **1 vale de pavo Redondos de 7 kg**. **Máximo 1 vale por DNI al día**, aunque compre varias veces. Meta: **6 canjes por impulsadora al día**.

| Mercado | Sacos para 1 vale |
|---|---|
| Mercado Productores | 25 |
| Fiori | 17 |
| Unicachi Sur | 12 |
| Parada | 12 |
| Unicachi Norte Mayorista | 10 |
| Unicachi Norte Minorista, Huamantanga, Ciudad de Dios, Caquetá | 5 |

Las reglas están en `SACOS_POR_MERCADO` (apps-script/Code.gs, que manda) y en `config.js` (copia para usar sin señal).

| Página | Para quién | Qué hace |
|---|---|---|
| `index.html` | Todos | Inicio con 3 botones: Promotores, Back Office (interno Free) y Equipo Rinti (cliente) |
| `app.html` | Impulsadoras y supervisor (celular) | Asistencia con selfie y GPS, pedidos y canjes (código de vale, boleta, firma, foto), incidencias y cierre de pavos |
| `dashboard.html` | Cliente (Rinti), entra con usuario y contraseña | Avance en vivo contra la meta, por mercado e impulsadora, asistencia, cuadre de pavos, incidencias, galería de evidencias, CSV y PDF |
| `apps-script/Code.gs` | Backend | Guarda todo en Google Sheets y las fotos en Google Drive |

Sin `API_URL` configurada, ambas páginas funcionan en **modo demo** (la app no envía nada y el dashboard muestra datos ficticios).

## Cobertura del TDR

| Requisito del TDR | Dónde se registra | Dónde se ve |
|---|---|---|
| Mercado | Mercado del día (app) / puesto en Productores | Filtro y tabla por mercado |
| Impulsador | Login con nombre + PIN | Tabla "Equipo en campo" |
| Fecha de ejecución | Automática (hora de Lima) | Filtro de periodo |
| Pedidos generados | Cada registro es un pedido (con o sin canje) | KPI "Pedidos generados" |
| Cantidad de sacos | Ricocan / Ricocat por separado | KPI "Sacos comprados" y tabla por mercado |
| Cumplimiento de la mecánica | Boleta + n.°, registro, firma, foto, compra en Lima (obligatorios para canjear) + validación del supervisor | KPI "Cumplimiento de mecánica" |
| Número de canjes | Pedido con 8 sacos o más | KPI, gráficos y tablas |
| Número de vales entregados | 1 vale cada 8 sacos, **cada uno con su código** (no se repiten) | KPI "Vales entregados" y detalle |
| Evidencia fotográfica | Selfie de asistencia, foto 1 de la boleta, foto 2 del cliente con el vale y firma | Galería con detalle por canje |
| Incidencias y observaciones | Formulario de incidencias + observaciones del canje | Panel de incidencias |
| Seguimiento periódico y consolidado final | — | Filtros Hoy / Ayer / 7 días / Campaña / Rango + **Descargar Excel (CSV)** e **Imprimir / PDF** |

## Despliegue (una sola vez, ~20 minutos)

### 1. Base de datos (Google Sheet + Apps Script)
1. Crea un Google Sheet vacío llamado **Canje Rinti · Base de datos**.
2. En el Sheet: **Extensiones → Apps Script**. Borra lo que haya y pega todo `apps-script/Code.gs`. Guarda.
3. Elige la función **`setup`** y pulsa **Ejecutar**. Acepta los permisos (Sheets y Drive).
   Esto crea las pestañas, la carpeta de fotos en Drive, 9 impulsadoras + 1 supervisor con **PIN aleatorio** y el token del dashboard.
4. Ve a **Ver → Registros de ejecución** y copia el **token del dashboard**.

### 2. Publicar el backend
1. **Implementar → Nueva implementación → Tipo: Aplicación web**.
2. Ejecutar como: **Yo**. Quién tiene acceso: **Cualquier usuario**.
3. Copia la URL que termina en `/exec`.

> Cada vez que cambies `Code.gs`: **Implementar → Gestionar implementaciones → editar → Nueva versión** (así la URL no cambia).

### 3. Configurar y subir a GitHub Pages
1. En `config.js` pega la URL en `API_URL`.
2. Sube esta carpeta a un repositorio de GitHub y activa **Settings → Pages → Deploy from branch → main / root**.
3. Links finales:
   - Impulsadoras: `https://TU-USUARIO.github.io/rinti-canje/`
   - Cliente: `https://TU-USUARIO.github.io/rinti-canje/dashboard.html?k=TOKEN`

### 4. Antes de salir a campo
- En la pestaña **Equipo** cambia "Impulsadora 1…9" por los nombres reales y revisa el mercado de cada una.
- **Accesos de impulsadoras**: uno por mercado, en la pestaña **Equipo** (columnas `usuario` y `contrasena`). Cada impulsadora solo registra en su mercado.
- **Supervisores y coordinador** (`supervisor`, `supervisor2`, `coordinador`): eligen mercado, hacen el cierre de pavos y tienen el **Panel del equipo** (quién llegó, quién llegó tarde, quién no marcó entrada, canjes de cada una contra la cuota de 6, último registro, incidencias). Pon el celular de cada impulsadora en la columna `celular` de **Equipo** para que el panel muestre los botones de llamar y WhatsApp.
- **Regla del vale**: el código son **exactamente 7 números** y cada vale se usa una vez. Si cargas los códigos impresos en la pestaña **Talonario** (columna `codigo`, y `mercado` si cada mercado tiene su talonario), solo se aceptan esos códigos.
- **Acceso del cliente**: pestaña **Config**, filas `cliente_usuario` y `cliente_clave`.
- **Dos dashboards**:
  - **Cliente** (`dashboard.html`): avance, canjes, sacos, asistencia (cada impulsadora solo como Activa / Inactiva), cuadre de pavos y evidencias. No muestra incidencias, observaciones, tardanzas, celulares ni el DNI completo.
  - **Interno Free** (`interno.html`): todo lo anterior más incidencias, observaciones, horas de entrada y salida, DNI completo y la **validación de cada canje** (Válido / Observado con motivo) desde la foto del vale. Se entra con el usuario de supervisor o coordinador.
- Las contraseñas se pueden cambiar escribiendo en la hoja; aplica al instante.
- En **Config** pon `fecha_inicio`, `fecha_fin` y `hora_ingreso`.
- Haz una prueba completa y **borra las filas de prueba** (no las cabeceras) en Asistencia, Pedidos, Vales, Incidencias y Cierres.
- En la capacitación, cada chica abre el link y lo agrega a la pantalla de inicio ("Agregar a inicio").

## Operación diaria
- **Validación (supervisor o coordinador):** en el dashboard interno abre la foto de cada canje y marca **Válido** u **Observado** (con motivo). Queda guardado quién y cuándo. Todo canje entra como Pendiente; el cliente lo ve como "En revisión por Free" hasta que se valida. Lo observado baja el % de cumplimiento.
- **Sin señal:** la app guarda los registros en el teléfono y los envía sola al volver la señal. No se duplican.
- **Vale repetido:** el sistema rechaza un código de vale ya usado.
- **Link del cliente filtrado:** ejecuta `renovarTokenDashboard()` en Apps Script y envía el link nuevo.

## Pruebas del backend

```bash
osascript -l JavaScript tests/test_backend.js
```

## Privacidad
- El DNI y el celular del comprador solo se guardan con su consentimiento (Ley 29733). El dashboard del cliente no muestra el celular y muestra el DNI enmascarado (solo los 3 últimos dígitos).
- Las fotos quedan en tu Drive con acceso "cualquiera con el enlace" (necesario para mostrarlas en el dashboard). Los enlaces no son públicos ni se pueden buscar.
