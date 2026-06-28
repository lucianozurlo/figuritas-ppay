# Runbook de Implementación — Mi Figurita Personal Pay IA
**Versión 3.0 — Formulario HTML custom · Captura de cámara · Conversión JPEG en cliente**

> Cambio principal respecto a v2: Google Form fue eliminado completamente.
> El formulario es ahora una página HTML hosteada en GitHub Pages o Netlify.
> La foto se captura o selecciona en el dispositivo del participante,
> se convierte a JPEG y se corrige la orientación EXIF en el browser,
> **antes** de subir. El backend es un Apps Script Web App anónimo (doPost).

---

## Índice

1. [Arquitectura del sistema](#1-arquitectura-del-sistema)
2. [Estructura de archivos del proyecto](#2-estructura-de-archivos-del-proyecto)
3. [Setup de Google Drive y Google Sheets](#3-setup-de-google-drive-y-google-sheets)
4. [Setup del backend en Apps Script](#4-setup-del-backend-en-apps-script)
5. [Publicación del Web App](#5-publicación-del-web-app)
6. [Setup del frontend](#6-setup-del-frontend)
7. [Deploy en GitHub Pages o Netlify](#7-deploy-en-github-pages-o-netlify)
8. [Pruebas end-to-end antes del evento](#8-pruebas-end-to-end-antes-del-evento)
9. [Configuración del mural digital](#9-configuración-del-mural-digital)
10. [Día del evento: operación y monitoreo](#10-día-del-evento-operación-y-monitoreo)
11. [Checklist previo al lanzamiento](#11-checklist-previo-al-lanzamiento)
12. [Plan B operativo](#12-plan-b-operativo)
13. [Cierre de operación](#13-cierre-de-operación)

---

## 1. Arquitectura del sistema

```
┌─────────────────────────────────────────────────────────────┐
│  DISPOSITIVO DEL PARTICIPANTE (browser)                     │
│                                                             │
│  index.html (GitHub Pages / Netlify)                        │
│  ├── Captura foto: cámara o galería                         │
│  ├── Lee orientación EXIF del archivo                       │
│  ├── Dibuja imagen corregida en <canvas>                    │
│  ├── Exporta como JPEG (sin HEIC, sin rotación)             │
│  └── POST JSON + foto base64 → Apps Script Web App          │
└────────────────────┬────────────────────────────────────────┘
                     │ HTTPS (fetch)
                     ▼
┌─────────────────────────────────────────────────────────────┐
│  APPS SCRIPT WEB APP (doPost, ejecuta como propietario)     │
│                                                             │
│  ├── Valida payload                                         │
│  ├── Detecta duplicados por mail                            │
│  ├── Guarda JPEG en Drive (carpeta 00_Fotos_Originales)     │
│  └── Escribe fila en Google Sheets con estado PENDIENTE     │
└────────────────────┬────────────────────────────────────────┘
                     │ trigger cada 10 min
                     ▼
┌─────────────────────────────────────────────────────────────┐
│  PROCESAMIENTO POR LOTES (Apps Script, trigger de tiempo)   │
│                                                             │
│  procesarLoteFiguritas()                                    │
│  ├── Reset zombies (PROCESANDO > 15 min → PENDIENTE)        │
│  ├── Por cada fila PENDIENTE:                               │
│  │   ├── Copia plantilla de Google Slides                   │
│  │   ├── Reemplaza marcadores de texto                      │
│  │   ├── Reemplaza forma con foto (DriveApp.getBlob)        │
│  │   ├── Exporta diapositiva como PNG                       │
│  │   └── Guarda PNG en Drive, actualiza estado              │
│                                                             │
│  enviarLoteCorreos()                                        │
│  └── Por cada fila FIGURITA_CREADA: envía email con adjunto │
└─────────────────────────────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────────┐
│  MURAL DIGITAL (pestaña Mural en el Google Sheet)           │
│  =QUERY() filtra por estado=EMAIL_ENVIADO + consentimiento  │
│  =IMAGE() renderiza las figuritas en celdas                 │
└─────────────────────────────────────────────────────────────┘
```

**Por qué no hay HEIC ni problemas de orientación:**
El browser del participante convierte cualquier formato (HEIC, HEIF, PNG, WebP, etc.)
a JPEG antes de enviar. Esto se hace con `createImageBitmap()` + `<canvas>.toBlob()`,
que el browser nativo maneja independientemente del formato de origen.
La orientación EXIF se lee del archivo original y se aplica como transformación
de canvas antes de exportar. El servidor recibe siempre un JPEG derecho.

---

## 2. Estructura de archivos del proyecto

```
figurita-personal-pay/
│
├── frontend/
│   └── index.html          ← formulario completo (HTML + CSS + JS en un archivo)
│
└── backend/                ← estos archivos van en el editor de Apps Script
    ├── Config.gs            ← constantes globales (IDs, nombres de columnas)
    ├── Utils.gs             ← funciones utilitarias compartidas
    ├── WebApp.gs            ← doPost: recibe el form, guarda foto, escribe Sheet
    └── Processing.gs        ← generación por lotes, envío, menú admin, triggers
```

Los 4 archivos `.gs` van todos en el **mismo proyecto de Apps Script**,
vinculado al Google Sheet `Registro de Figuritas`.

---

## 3. Setup de Google Drive y Google Sheets

### 3.1. Estructura de carpetas en Drive

Crear la siguiente estructura bajo la cuenta del operador técnico
(la misma que va a correr el Web App):

```
[EVENTO] Mi Figurita Personal Pay IA/
├── 00_Fotos_Originales/        ← donde doPost guarda los JPEGs subidos
├── 01_Plantillas/              ← plantilla de Slides
├── 03_Figuritas_Generadas/     ← donde el script guarda los PNGs finales
└── 04_Mural_Assets/            ← reservado para assets del mural
```

Anotar el ID de cada carpeta (aparece en la URL de Drive al abrirla):
```
https://drive.google.com/drive/folders/AQUI_ESTA_EL_ID
```

| Carpeta | ID a copiar |
|---------|-------------|
| `00_Fotos_Originales` | → `CONFIG.FOLDER_FOTOS_ID` |
| `03_Figuritas_Generadas` | → `CONFIG.FOLDER_FIGURITAS_ID` |

### 3.2. Google Sheet: Registro de Figuritas

1. Ir a Google Sheets → Nueva hoja de cálculo.
2. Nombre: `Registro de Figuritas`.
3. Moverla a la carpeta raíz del evento.
4. En la fila 1, crear los siguientes encabezados **en este orden exacto**,
   empezando en la celda A1:

```
Nombre | Mail | Área | Superpoder | Actitud | Consentimiento mural |
id_archivo_drive | url_archivo_drive |
id_figurita | id_figurita_generada | url_figurita_generada |
estado | detalle_error | timestamp_procesando
```

> Estos encabezados los escribe el script directamente (ya no es Google Forms).
> Deben coincidir exactamente con los valores en `CONFIG.COLUMNS`.

5. Formatear la columna `estado` con colores condicionales para facilitar el monitoreo:
   - `PROCESANDO` → fondo amarillo
   - `FIGURITA_CREADA` → fondo azul claro
   - `EMAIL_ENVIADO` → fondo verde
   - `ERROR` → fondo rojo
   - `ERROR_EMAIL` → fondo naranja

### 3.3. Plantilla de Google Slides

1. Crear una nueva presentación en `01_Plantillas/`.
2. Nombre: `Plantilla Figurita`.
3. Configurar el tamaño de la diapositiva:
   Archivo → Configuración de página → Personalizado → usar las proporciones de la figurita.
4. Diseñar la figurita con los assets de marca.
5. Insertar marcadores de texto (cada uno en su propia caja independiente):

```
{{nombre}}
{{area}}
{{superpoder}}
{{actitud}}
```

6. Insertar una forma rectangular o circular donde irá la foto.
7. Clic derecho sobre esa forma → **Texto alternativo** → Título:

```
FOTO_PERFIL_REEMPLAZAR
```

8. Anotar el ID de la presentación (en la URL, entre `/d/` y `/edit`):
   → va en `CONFIG.SLIDE_TEMPLATE_ID`.

---

## 4. Setup del backend en Apps Script

### 4.1. Crear el proyecto de Apps Script

1. Abrir el Sheet `Registro de Figuritas`.
2. `Extensiones → Apps Script`.
3. Cambiar el nombre del proyecto a: `Figuritas Admin`.
4. El proyecto tiene por defecto un archivo `Code.gs`. Renombrarlo a `Config.gs`
   haciendo clic sobre el nombre del archivo → Renombrar.
5. Agregar los tres archivos restantes con el botón `+` junto a "Archivos":
   - `Utils.gs`
   - `WebApp.gs`
   - `Processing.gs`
6. Pegar el contenido de cada archivo `.gs` en el archivo correspondiente.

### 4.2. Habilitar la API de Google Slides

1. En el panel izquierdo del editor, hacer clic en `Servicios +`.
2. Buscar `Google Slides API`.
3. Hacer clic en `Agregar`.
4. Verificar que aparece en la lista.

### 4.3. Completar CONFIG.gs con los IDs reales

Abrir `Config.gs` y reemplazar los cuatro placeholders:

```javascript
FOLDER_FIGURITAS_ID: 'REEMPLAZAR_CON_ID_CARPETA_03_FIGURITAS',
FOLDER_FOTOS_ID:     'REEMPLAZAR_CON_ID_CARPETA_FOTOS',
SLIDE_TEMPLATE_ID:   'REEMPLAZAR_CON_ID_PLANTILLA_SLIDES',
```

También verificar que los nombres en `COLUMNS` coinciden exactamente con
los encabezados del Sheet creados en el paso 3.2.

### 4.4. Ejecutar validarConfiguracion()

1. En el editor, seleccionar la función `validarConfiguracion` en el desplegable superior.
2. Hacer clic en `Ejecutar`.
3. La primera vez pedirá autorización — aceptar todos los permisos solicitados
   (Drive, Sheets, Gmail, UrlFetch).
4. Verificar que el resultado sea "✅ Todo en orden". Si hay errores, corregirlos
   antes de continuar.

> Los permisos que se autorizan aquí son los que usa el Web App cuando
> corre como el propietario del script. Sin esta autorización previa,
> el doPost fallará silenciosamente.

---

## 5. Publicación del Web App

Esta es la parte más crítica del setup. El Web App es la URL pública a la que
el formulario HTML envía los datos. Debe estar publicado correctamente antes
de configurar el frontend.

### 5.1. Publicar por primera vez

1. En el editor de Apps Script, hacer clic en el botón azul **Implementar** (arriba a la derecha).
2. Seleccionar **Nueva implementación**.
3. Hacer clic en el engranaje junto a "Seleccionar tipo" → **Aplicación web**.
4. Configurar:
   - **Descripción:** `Figuritas v1`
   - **Ejecutar como:** `Yo (tu cuenta)` — esto es lo que permite acceder a Drive y Sheets sin que el usuario esté autenticado.
   - **Quién tiene acceso:** `Cualquier usuario` — necesario para que los participantes anónimos puedan enviar el formulario.
5. Hacer clic en **Implementar**.
6. Copiar la **URL del Web App**. Tiene este formato:

```
https://script.google.com/macros/s/AKfycby.../exec
```

> Guardar esta URL. Se necesita en el paso 6.

### 5.2. IMPORTANTE: republicar después de cada cambio de código

Cada vez que se modifique el código del backend (cualquiera de los `.gs`),
el Web App **no se actualiza automáticamente**. Se debe:

1. Ir a **Implementar → Administrar implementaciones**.
2. Hacer clic en el lápiz (editar) de la implementación activa.
3. Cambiar la versión a **Nueva versión**.
4. Guardar.

Si no se republica, el formulario sigue usando el código viejo.

### 5.3. Verificar que el Web App responde

Abrir la URL del Web App en el navegador (GET simple). Debe responder con el
HTML del formulario (si `doGet()` está implementado) o un error controlado.
Si devuelve "Error 404" o pantalla en blanco, revisar que el tipo sea
"Aplicación web" y que el acceso sea "Cualquier usuario".

---

## 6. Setup del frontend

### 6.1. Abrir index.html

Abrir el archivo `frontend/index.html` en un editor de texto.

### 6.2. Reemplazar la URL del backend

Buscar esta línea al inicio del bloque `<script>`:

```javascript
const BACKEND_URL = 'https://script.google.com/macros/s/REEMPLAZAR_CON_WEB_APP_ID/exec';
```

Reemplazar con la URL real obtenida en el paso 5.1.

### 6.3. Actualizar las opciones del campo Área

Buscar el bloque `<select id="inputArea">` en el HTML y reemplazar las opciones
con las áreas reales de la empresa:

```html
<select id="inputArea">
  <option value="" disabled selected>Seleccioná tu área</option>
  <option>Producto</option>
  <option>Tecnología</option>
  <!-- agregar o modificar según lista oficial -->
</select>
```

### 6.4. Personalizar textos y branding

Los colores principales están definidos como variables CSS al inicio del `<style>`:

```css
--c-brand:  #0062FF;   /* azul principal */
--c-accent: #00C6A2;   /* verde acento */
```

Reemplazar con los colores oficiales de Personal Pay si difieren.

El título y subtítulo están en el bloque `.header`:

```html
<h1 class="header__title">Tu <span>Figurita</span><br>Personal Pay</h1>
```

Modificar según el copy final del evento.

---

## 7. Deploy en GitHub Pages o Netlify

### Opción A: GitHub Pages

1. Crear un repositorio en GitHub (puede ser privado; GitHub Pages puede
   publicar desde rama `main` en repos privados en planes de pago, o público gratis).
2. Subir el archivo `index.html` a la raíz del repositorio (o a una carpeta `/docs`).
3. Ir a Settings → Pages → Source → seleccionar la rama y carpeta.
4. GitHub genera una URL del tipo `https://usuario.github.io/nombre-repo/`.
5. Esperar 1-2 minutos y verificar que la página carga.

### Opción B: Netlify (recomendada para más control)

1. Crear una cuenta en [netlify.com](https://netlify.com) si no existe.
2. En el dashboard, arrastrar la carpeta `frontend/` directamente a la zona
   de deploy (Netlify acepta drag & drop de carpetas).
3. Netlify genera una URL aleatoria del tipo `https://random-name.netlify.app`.
4. Opcionalmente, configurar un dominio personalizado o un subdominio limpio
   desde Site Settings → Domain management.

### Verificar CORS

El formulario hace un `fetch()` desde el dominio de GitHub Pages o Netlify
hacia el Web App de Apps Script. Apps Script, cuando está publicado como
"Cualquier usuario", **no requiere configuración de CORS adicional** porque
el request llega como `text/plain` (configurado así en el frontend para
evitar el preflight OPTIONS). Verificar que el header del fetch en `index.html`
sea exactamente:

```javascript
headers: { 'Content-Type': 'text/plain' },
```

Si se cambia a `application/json`, el browser lanzará un preflight OPTIONS
que Apps Script no responde, y el envío fallará.

---

## 8. Pruebas end-to-end antes del evento

### 8.1. Prueba básica de conectividad

1. Abrir el formulario desde un celular real (no desktop, no simulador).
2. Tocar el botón **Cámara** → sacar una selfie.
3. Verificar que aparece la vista previa en el recuadro.
4. Completar todos los campos.
5. Marcar el consentimiento obligatorio.
6. Tocar **Crear mi figurita**.
7. Verificar que aparece la pantalla de éxito.
8. Verificar en el Sheet que apareció una nueva fila con estado `PENDIENTE`.
9. Verificar en Drive (carpeta `00_Fotos_Originales`) que se guardó el JPEG.

### 8.2. Prueba con galería y con HEIC

1. Repetir la prueba anterior usando el botón **Galería** en lugar de Cámara.
2. Si hay un iPhone disponible, intentar subir una foto HEIC directamente.
   El formulario debe convertirla sin error. Verificar en Drive que el archivo
   guardado es `image/jpeg`, no `image/heic`.

### 8.3. Prueba de generación

1. Con al menos 3-5 registros en el Sheet con estado `PENDIENTE`:
2. Ir al Sheet → menú `🎴 Figuritas Admin → 1. Procesar lote de figuritas`.
3. Esperar a que termine (puede tardar 30-60 segundos por figurita).
4. Verificar en el Sheet que las filas pasaron a `FIGURITA_CREADA`.
5. Verificar en Drive (carpeta `03_Figuritas_Generadas`) que se crearon los PNGs.
6. Abrir al menos un PNG y verificar que la foto está correctamente orientada,
   que los textos se ven bien y que la imagen no está en blanco.

### 8.4. Prueba de envío

1. Con registros en `FIGURITA_CREADA`:
2. Ir al menú → `2. Enviar lote de correos`.
3. Verificar en el Sheet que pasaron a `EMAIL_ENVIADO`.
4. Verificar en el correo del participante de prueba que llegó el mail
   con el PNG adjunto y el link de Drive funcional.

### 8.5. Prueba de error y reprocesamiento

1. Crear un registro de prueba con un ID de foto inválido en `id_archivo_drive`
   (borrar el valor o poner un texto cualquiera).
2. Ejecutar `Procesar lote de figuritas`.
3. Verificar que la fila queda en `ERROR` con un mensaje en `detalle_error`.
4. Verificar que las otras filas del lote siguen procesándose.
5. Ejecutar `Reprocesar errores`.
6. Verificar que la fila volvió a `PENDIENTE`.

### 8.6. Prueba de zombie

1. Cambiar manualmente el estado de una fila a `PROCESANDO` y dejar el
   `timestamp_procesando` vacío (o poner una fecha de hace 30 minutos).
2. Ejecutar `Procesar lote de figuritas`.
3. Verificar que al inicio del log aparece el reset del zombie y la fila
   vuelve a `PENDIENTE` antes de procesar el lote.

### 8.7. Prueba de duplicado

1. Enviar el formulario dos veces con el mismo mail.
2. Verificar que en el Sheet solo aparece una fila (la segunda envío responde
   `ok` pero no genera fila nueva).

### 8.8. Prueba de triggers automáticos

1. Ejecutar `Crear triggers de tiempo` desde el menú admin.
2. Agregar un registro de prueba al formulario.
3. Esperar máximo 15 minutos sin ejecutar nada manualmente.
4. Verificar que el registro fue procesado automáticamente.

---

## 9. Configuración del mural digital

El mural es una pestaña del Sheet con `=QUERY()` y `=IMAGE()`.
No usa Looker Studio (no funciona con URLs de Drive como imágenes).

### 9.1. Crear la pestaña Mural

1. En el Sheet `Registro de Figuritas`, hacer clic en `+` para agregar hoja.
2. Nombre: `Mural`.

### 9.2. Encabezados

En la fila 1 escribir:

```
A1: Nombre   B1: Área   C1: URL   D1: Vista previa
```

### 9.3. Fórmula de datos

En la celda `A2`, pegar la siguiente fórmula (ajustar las letras de columna
según la posición real de cada campo en el Sheet):

```
=IFERROR(
  QUERY(
    'Registro de Figuritas'!A:N,
    "SELECT A, C, K WHERE L = 'EMAIL_ENVIADO' AND F = 'Sí'",
    0
  ),
  "Sin figuritas en el mural aún."
)
```

> Columnas del SELECT:
> - `A` = Nombre (columna A del Sheet)
> - `C` = Área (columna C)
> - `K` = url_figurita_generada (columna K, verificar posición real)
> - `L` = estado (columna L, verificar)
> - `F` = Consentimiento mural (columna F, verificar)
>
> **Importante:** verificar las letras de columna contra el Sheet real antes
> de usar esta fórmula. Pueden diferir si se agregaron columnas en otro orden.

### 9.4. Imágenes con `=IMAGE()`

1. En la celda `D2` pegar: `=IF(C2<>"", IMAGE(C2, 4, 150, 120), "")`
2. Seleccionar D2 y extender hacia abajo arrastrando hasta la fila 500
   (o el número máximo esperado de participantes con consentimiento de mural).
3. Seleccionar toda la columna D → Formato → Alto de fila → 150px.
4. Seleccionar toda la columna D → Formato → Ancho de columna → 130px.

> `IMAGE(url, 4, alto, ancho)` con modo 4 usa dimensiones en píxeles.
> Las URLs en columna C deben ser del formato
> `https://drive.google.com/uc?export=view&id=FILE_ID`
> que es exactamente lo que guarda el script de generación.

### 9.5. Proteger la pestaña

Clic derecho sobre la pestaña `Mural` → Proteger la hoja.
Solo el operador puede editar; el resto solo lee.

### 9.6. Proyección durante el evento

Abrir la pestaña `Mural` en pantalla completa (`F11`) en la pantalla del evento.
El Sheet se actualiza automáticamente cada vez que se ejecuta el script de generación.

---

## 10. Día del evento: operación y monitoreo

### 10.1. Puesto de control

Tener abierto el Sheet en una pantalla dedicada, con la hoja `Registro de Figuritas`
visible, columnas `estado` y `detalle_error` en vista. Monitorear el resumen
ejecutando `🎴 Figuritas Admin → Ver resumen de estados` cada 15-20 minutos.

### 10.2. Flujo normal esperado

```
Participante envía formulario
        ↓
Fila aparece en Sheet con estado PENDIENTE
        ↓
Trigger (o ejecución manual) procesa el lote
        ↓
Estado pasa a PROCESANDO (max 15 min)
        ↓
Estado pasa a FIGURITA_CREADA
        ↓
Trigger (o ejecución manual) envía correos
        ↓
Estado pasa a EMAIL_ENVIADO ✓
```

### 10.3. Gestión de errores

#### ERROR (fallo en generación)

1. Leer `detalle_error` en el Sheet.
2. Causas comunes:

| Mensaje | Causa | Acción |
|---------|-------|--------|
| `No hay ID de foto para este registro` | El doPost no guardó la foto | Revisar logs del Web App en Apps Script |
| `Blob de foto inválido o vacío` | Archivo corrupto en Drive | Pedir al participante que reenvíe desde el formulario |
| `Forma con alt text no encontrada` | Plantilla modificada | Restaurar la plantilla y verificar el alt text |
| `Exportación PNG falló con HTTP 4xx` | Token expirado o cuota | Esperar y reprocesar |
| `PNG exportado está vacío` | Error transitorio de Google | Reprocesar |

3. Usar `Figuritas Admin → Reprocesar errores`.
4. Ejecutar `Procesar lote de figuritas` manualmente si no puede esperar el trigger.

#### ERROR_EMAIL (fallo en envío)

1. Revisar `detalle_error`.
2. Verificar cuota con `Ver resumen de estados`.
3. Si `id_figurita_generada` no está vacío, el PNG existe. Usar `Reprocesar errores`.
4. Ejecutar `Enviar lote de correos` manualmente.

#### PROCESANDO zombie (fila atascada)

El script los resuelve automáticamente al inicio de cada ciclo.
Si se necesita resolución inmediata: cambiar el estado a `PENDIENTE` manualmente
en el Sheet y ejecutar el lote.

### 10.4. Cuándo ejecutar manualmente

Los triggers corren cada ~10 min (puede ser hasta 25 min con carga alta de Google).
Ejecutar manualmente cuando:
- Hay muchos participantes esperando y el trigger tardó más de 20 minutos.
- Se quiere confirmar que un lote específico fue procesado.
- Después de reprocesar errores.

Secuencia de ejecución manual:
1. `Procesar lote de figuritas`
2. `Enviar lote de correos`
3. `Ver resumen de estados` para confirmar

---

## 11. Checklist previo al lanzamiento

### Drive y Sheets

- [ ] Estructura de carpetas creada en Drive.
- [ ] Sheet `Registro de Figuritas` creado con los 14 encabezados correctos.
- [ ] Plantilla de Slides con marcadores y alt text configurados.
- [ ] IDs de carpetas y plantilla anotados.

### Backend (Apps Script)

- [ ] Los 4 archivos `.gs` pegados en el proyecto.
- [ ] API de Google Slides habilitada.
- [ ] IDs reales cargados en `Config.gs` (sin placeholders).
- [ ] Nombres de columnas en `Config.gs` verificados contra el Sheet.
- [ ] `validarConfiguracion()` ejecutada sin errores.
- [ ] Permisos autorizados (Drive, Sheets, Gmail, UrlFetch).

### Web App

- [ ] Web App publicado con "Ejecutar como: Yo" y "Acceso: Cualquier usuario".
- [ ] URL del Web App copiada.
- [ ] Web App responde a un GET sin error.

### Frontend

- [ ] `BACKEND_URL` reemplazada con la URL real del Web App.
- [ ] Opciones del campo Área actualizadas con las áreas reales.
- [ ] Colores y textos de branding ajustados.
- [ ] Formulario probado desde celular iOS real.
- [ ] Formulario probado desde celular Android real.
- [ ] Foto sacada con cámara: llega como JPEG en el Sheet.
- [ ] Foto desde galería HEIC: llega como JPEG en el Sheet.
- [ ] Fotos rotadas: llegan correctamente orientadas.

### Deploy

- [ ] Frontend deployado en GitHub Pages o Netlify.
- [ ] URL final probada desde celular externo (no el del desarrollador).
- [ ] QR generado apuntando a la URL final.
- [ ] QR probado desde celular real.
- [ ] Link directo disponible como respaldo del QR.

### Procesamiento

- [ ] Al menos 10 figuritas generadas en prueba end-to-end.
- [ ] Al menos 5 mails recibidos con PNG adjunto correcto.
- [ ] Prueba de error y reprocesamiento exitosa.
- [ ] Prueba de zombie exitosa.
- [ ] Prueba de duplicado exitosa.
- [ ] Triggers creados con `createTimeTriggers()`.
- [ ] Triggers verificados en el panel de Apps Script.

### Mural

- [ ] Pestaña `Mural` creada con fórmulas correctas.
- [ ] `=IMAGE()` renderiza imágenes de prueba.
- [ ] Vista de proyección probada en pantalla grande.

### Operación

- [ ] Responsable de monitoreo asignado y entrenado.
- [ ] Runbook impreso o disponible offline para el operador.

---

## 12. Plan B operativo

| Componente | Fallo posible | Plan B |
|------------|---------------|--------|
| Frontend caído | GitHub Pages / Netlify no responde | Tener el HTML guardado localmente y servir desde un servidor local con `npx serve .` o `python -m http.server`. |
| Web App no acepta envíos | Error 500 o timeout en doPost | Revisar logs en Apps Script (Ejecutar → Registros de ejecución). Republicar el Web App. |
| QR no funciona | Error de lectura o link roto | Tener la URL del formulario en un slide del evento y dictarla si hace falta. |
| Trigger no ejecuta | Demorado por carga de Google | Ejecutar manualmente desde el menú Admin. |
| Foto HEIC no convierte | Browser muy viejo que no soporta `createImageBitmap` para HEIC | Pedir al participante que elija una foto JPG de su galería (las fotos viejas generalmente están en JPG). |
| Export PNG falla masivamente | Problema con la API de Slides | Reducir `BATCH_SIZE_GENERACION` a 5 en `Config.gs` y republicar el Web App. |
| Cuota de mail agotada | Más de ~1400 envíos en el día | Enviar desde cuenta alternativa autorizada, o diferir envíos al día siguiente. |
| Mural no actualiza | Fórmula QUERY no refresca | Ctrl+Shift+F5 en el Sheet para forzar recalculo. |
| Participante no recibe el correo | Mail corporativo bloqueó el adjunto | El email incluye un link directo a Drive como respaldo — el participante puede acceder desde ahí. |

---

## 13. Cierre de operación

1. Ejecutar `Ver resumen de estados` y anotar los totales.
2. Ejecutar `Reprocesar errores` para reintentar cualquier fallo pendiente.
3. Ejecutar un último ciclo manual: generar + enviar.
4. Confirmar estado final en el Sheet.
5. Registrar métricas de cierre:

| Métrica | Valor |
|---------|-------|
| Total de envíos del formulario | |
| Total con EMAIL_ENVIADO | |
| Total con ERROR sin resolver | |
| Total en el Mural | |
| Cuota de mail restante | |

6. Ejecutar `Eliminar todos los triggers` desde el menú Admin.
7. Revocar el acceso público al Web App:
   Implementar → Administrar implementaciones → Editar → Acceso → "Solo yo" → Guardar.
8. Resguardar la carpeta raíz del evento en una carpeta de archivo en Drive.
9. Exportar el Sheet como Excel o PDF para registro permanente.

---

*Runbook v3.0 — Formulario HTML custom con captura de cámara y conversión JPEG en cliente.*
*Elimina Google Forms, HEIC y errores de orientación EXIF como problemas operativos.*
