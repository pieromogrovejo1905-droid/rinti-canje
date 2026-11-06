# Canje Rinti · Registro de campo y dashboard

Sistema para la campaña **Canje Rinti (Ricocan / Ricocat)**: 9 impulsadoras en 8 mercados + Mercado Productores.
Mecánica: **8 sacos (Ricocan 22 kg o Ricocat 15 kg) = 1 vale de pavo Redondos de 7 kg**. Meta: **6 canjes por impulsadora al día**.

| Página | Para quién | Qué hace |
|---|---|---|
| `index.html` | Impulsadoras y supervisor (celular) | Asistencia con selfie y GPS, pedidos y canjes (código de vale, boleta, firma, foto), incidencias y cierre de pavos |
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
| Evidencia fotográfica | Selfie de asistencia, boleta, entrega del vale y firma | Galería con detalle por canje |
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
- **Accesos** (pestaña **Config**): `app_usuario` / `app_clave` es el acceso general de las promotoras (luego eligen su nombre); `cliente_usuario` / `cliente_clave` es el acceso del dashboard para Rinti. Puedes cambiar las contraseñas ahí mismo; aplica al instante.
- El PIN de la pestaña Equipo queda como alternativa por si luego se quiere acceso individual.
- En **Config** pon `fecha_inicio`, `fecha_fin` y `hora_ingreso`.
- Haz una prueba completa y **borra las filas de prueba** (no las cabeceras) en Asistencia, Pedidos, Vales, Incidencias y Cierres.
- En la capacitación, cada chica abre el link y lo agrega a la pantalla de inicio ("Agregar a inicio").

## Operación diaria
- **Supervisor (agencia):** en la pestaña Pedidos cambia la columna `validacion` a **Válido** u **Observado** después de revisar las fotos. Lo observado baja el % de cumplimiento en el dashboard.
- **Sin señal:** la app guarda los registros en el teléfono y los envía sola al volver la señal. No se duplican.
- **Vale repetido:** el sistema rechaza un código de vale ya usado.
- **Link del cliente filtrado:** ejecuta `renovarTokenDashboard()` en Apps Script y envía el link nuevo.

## Privacidad
- El celular del comprador solo se guarda con su consentimiento (Ley 29733) y **no aparece en el dashboard del cliente**.
- Las fotos quedan en tu Drive con acceso "cualquiera con el enlace" (necesario para mostrarlas en el dashboard). Los enlaces no son públicos ni se pueden buscar.
