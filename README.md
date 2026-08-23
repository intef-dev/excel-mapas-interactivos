# 🗺️ Convierte tu Tabla de Excel en un Mapa Interactivo

¡Bienvenido a este repositorio! Aquí encontrarás el código completo y los snippets listos para usar de las **3 aplicaciones empresariales de mapas interactivos dentro de Microsoft Excel**, utilizando la tecnología oficial y gratuita de **Script Lab** y **Leaflet.js**.

Este proyecto complementa el video tutorial paso a paso de mi canal de YouTube **Inteligencia Eficiente**. Aprenderás cómo conectar tablas de clientes, telemetría GPS en vivo y análisis inmobiliario con resaltado condicional de celdas en tiempo real.

---

## 🎥 1. Video Tutorial Paso a Paso

Mira el desarrollo y la explicación técnica completa en YouTube:  
👉 **[Ver Video en YouTube - HAZ CLIC AQUÍ](https://www.youtube.com/InteligenciaEficiente?sub_confirmation=1)**

¡Si te sirve este contenido, no olvides dejar tu **Like**, **comentar** qué caso de uso aplicarás en tu trabajo y **suscribirte al canal** para no perderte futuros tutoriales!

---

## 🌐 2. Demo Interactiva en Vivo

Puedes probar la simulación interactiva completa de Microsoft Excel Web con el panel lateral de Script Lab directamente en tu navegador (sin necesidad de abrir Excel):  
👉 **[Ver Demo Interactiva en GitHub Pages](https://inteligenciaeficiente.github.io/AGO202603VL02-Convierte-tu-tabla-de-Excel-en-un-mapa-interactivo/demo/)**

---

## 📥 3. Descarga Rápida del Proyecto

Si deseas descargar todos los archivos fuente, hojas de datos CSV y plantillas YAML en un solo archivo comprimido ZIP:

[![Descargar ZIP](https://img.shields.io/badge/Descargar-Proyecto_Completo_ZIP-blue?style=for-the-badge&logo=github)](https://github.com/InteligenciaEficiente/AGO202603VL02-Convierte-tu-tabla-de-Excel-en-un-mapa-interactivo/archive/refs/heads/main.zip)

---

## 🚀 4. Las 3 Aplicaciones Incluidas

| App | Caso de Uso | Características Principales |
| :---: | :--- | :--- |
| **App 1** | **Gestión Comercial & Rutas** | Pines de clientes categorizados, selector interactivo de radio de zona y cálculo de ruta con descarga automática de la tabla `Ruta_Generada` a Excel. |
| **App 2** | **Rastreo GPS & Telemetría** | Simulación de ruta con radar animado, tacómetro, odómetro y combustible en vivo, con **resaltado en amarillo de la fila activa en Excel** en tiempo real. |
| **App 3** | **Explorador Inmobiliario** | Slicer deslizante de presupuesto (€150k a €1M), **resaltado bicolor en Excel** (Verde = Mejor Oportunidad €/m², Azul = En presupuesto) y descarga de **Resumen Ejecutivo con KPIs**. |

---

## ⚡ 5. Método Rápido: Importar en 1 Clic (Gists Oficiales o Archivos YAML)

La forma más rápida de ejecutar cualquiera de las 3 aplicaciones en tu Excel es importando los **GitHub Gists oficiales** directamente en Script Lab sin tener que copiar archivo por archivo:

### 🔗 Enlaces directos a los Gists de cada App:

| Aplicación | Caso de Uso | Enlace del GitHub Gist (1 Clic) |
| :--- | :--- | :--- |
| 📍 **App 1** | **Planificador de Rutas y Zonas de Clientes** | [`https://gist.github.com/intef-dev/beaa3c1d368562f0a47de9939986a243`](https://gist.github.com/intef-dev/beaa3c1d368562f0a47de9939986a243) |
| 🚚 **App 2** | **Rastreo GPS de Flotas y Telemetría en Vivo** | [`https://gist.github.com/intef-dev/3e6eb2e3bf0ce05f5a68fdbcfe411559`](https://gist.github.com/intef-dev/3e6eb2e3bf0ce05f5a68fdbcfe411559) |
| 🏢 **App 3** | **Explorador Inmobiliario y Slicer de Presupuesto** | [`https://gist.github.com/intef-dev/0cc96faa4ad891599678a131955cb0bc`](https://gist.github.com/intef-dev/0cc96faa4ad891599678a131955cb0bc) |

---

### 📋 Paso a Paso para Importar en Excel con Script Lab:

#### Opción A: Mediante la URL del Gist (Recomendado)
1. **Copia la URL** del Gist de la app que deseas probar (de la tabla superior).
2. En **Microsoft Excel**, ve a la pestaña **Script Lab** y haz clic en **Code** para abrir el panel lateral.
3. En el menú superior izquierdo de Script Lab (icono de hamburguesa **☰** o menú desplegable de snippets), haz clic en **Import** (Importar).
4. En el cuadro superior que dice **"Import from a GitHub Gist URL or ID"**, pega la URL del Gist.
5. Haz clic en **Import**.
6. ¡Listo! Se cargará el código completo con todas las librerías configuradas. Haz clic en **Run** (botón de reproducción) para iniciar el mapa interactivo.

#### Opción B: Mediante archivo YAML local (Sin conexión a Gist)
1. Abre cualquiera de los archivos `.yaml` de la carpeta [`snippets/`](./snippets/):
   - 📍 [`snippets/app1_clientes_rutas.yaml`](./snippets/app1_clientes_rutas.yaml)
   - 🚚 [`snippets/app2_rastreo_gps.yaml`](./snippets/app2_rastreo_gps.yaml)
   - 🏢 [`snippets/app3_inmobiliario_presupuesto.yaml`](./snippets/app3_inmobiliario_presupuesto.yaml)
2. Copia todo su contenido de texto (`Ctrl + A` y luego `Ctrl + C`).
3. En el menú **Import** de Script Lab, pega el texto en el cuadro inferior **"Paste snippet YAML or JSON"** y pulsa **Import**.

---

## 🛠️ 6. Método Manual: Pestaña por Pestaña

Si prefieres copiar y pegar el código en las 4 pestañas individuales de Script Lab:

### Requisitos Previos:
Tener instalado el complemento gratuito **Script Lab** en Excel (pestaña _Insertar_ ➔ _Obtener complementos_ ➔ Buscar *"Script Lab"* ➔ _Agregar_).

### Pestañas de Configuración:
1. **Pestaña `Libraries`**: Pega las siguientes dependencias CDN:
   ```text
   https://appsforoffice.microsoft.com/lib/1/hosted/office.js
   @types/office-js
   https://unpkg.com/leaflet@1.9.4/dist/leaflet.js
   https://unpkg.com/leaflet@1.9.4/dist/leaflet.css
   @types/leaflet
   https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css
   ```
2. **Pestaña `HTML`**: Copia el archivo `index.html` de la carpeta de la app deseada ([`app1`](./app1/index.html), [`app2`](./app2/index.html) o [`app3`](./app3/index.html)).
3. **Pestaña `CSS`**: Copia el archivo `index.css` ([`app1`](./app1/index.css), [`app2`](./app2/index.css) o [`app3`](./app3/index.css)).
4. **Pestaña `Script`**: Copia el código TypeScript de `index.ts` o JavaScript de `index.js`.
5. Haz clic en **Run**.

---

## 📂 7. Estructura del Repositorio

```text
├── app1/                             # App 1: Clientes y Rutas
│   ├── index.html                    # Estructura del panel lateral
│   ├── index.css                     # Estilos minimalistas
│   ├── index.ts / index.js           # Lógica TypeScript / JavaScript
│   ├── clientes_demo.csv             # Tabla de 18 clientes
│   └── app1_snippet.yaml             # Snippet YAML para importar
├── app2/                             # App 2: Rastreo GPS
│   ├── index.html / index.css        # UI de telemetría y controles
│   ├── index.ts / index.js           # Lógica de playback y resaltado Excel
│   ├── vehiculos_gps.csv             # 20 puntos GPS con telemetría
│   └── app2_snippet.yaml             # Snippet YAML para importar
├── app3/                             # App 3: Explorador Inmobiliario
│   ├── index.html / index.css        # Slicer de presupuesto y tarjetas
│   ├── index.ts / index.js           # Sincronización bicolor y KPIs
│   ├── propiedades_demo.csv          # 20 inmuebles en Madrid
│   └── app3_snippet.yaml             # Snippet YAML para importar
├── demo/                             # Simulador interactivo en GitHub Pages
│   ├── index.html                    # Simulador de interfaz Excel Web
│   ├── demo.css                      # Estilos de la cuadrícula y ribbon
│   └── demo.js                       # Sincronización reactiva del grid
├── snippets/                         # Archivos YAML para importación directa
├── LICENSE                           # Licencia MIT
└── README.md                         # Documentación oficial
```

---

## 📊 8. Datasets de Ejemplo (CSV)

Cada carpeta incluye un dataset de ejemplo listo para pegar en Excel o inicializar automáticamente:
* 🛒 [`clientes_demo.csv`](./app1/clientes_demo.csv): 18 clientes con coordenadas en Madrid, zonas de venta y prioridades.
* 🚛 [`vehiculos_gps.csv`](./app2/vehiculos_gps.csv): 20 registros cronológicos de ruta con velocidades y combustible.
* 🏡 [`propiedades_demo.csv`](./app3/propiedades_demo.csv): 20 inmuebles con precios escalonados para la demostración reactiva del slicer.

---

## 💼 9. Consultoría y Desarrollo a Medida (B2B)

¿Necesitas implementar una solución personalizada para tu empresa conectada a tus propias bases de datos, APIs de mapas satelitales, ERPs o CRMs?

* ✉️ **Contacto Profesional:** [Escríbeme a través de mi sitio web oficial](https://inteficiente.com/)
* 🚀 **Soluciones Especializadas:** Add-ins a medida para Microsoft Excel, automatizaciones avanzadas y dashboards ejecutivos en Office.js.

---

## 📱 10. ¡Únete a la Comunidad! Sígueme en mis Redes

Para enterarte de nuevos tutoriales, macros, trucos de productividad y automatizaciones gratuitas:

* 🎥 **YouTube:** [Suscríbete a mi Canal de YouTube](https://www.youtube.com/InteligenciaEficiente?sub_confirmation=1)
* 📸 **Instagram:** [@inteligenciaeficiente](https://www.instagram.com/inteligenciaeficiente/)
* 👤 **Facebook:** [Página Oficial de Facebook](https://www.facebook.com/InteligenciaEficiente/)
* 🌐 **Sitio Web:** [inteficiente.com](https://inteficiente.com/)

---

## ☕ 11. Apoya este Proyecto

Si este proyecto te ha sido de utilidad y deseas apoyar la creación continua de plantillas de código abierto y tutoriales gratuitos:

* 💳 **PayPal:** [Donar por PayPal.Me](https://www.paypal.com/paypalme/INTEFDonativo)
* ☕ **Buy Me a Coffee:** [Invítame a un café en Buy Me a Coffee](https://buymeacoffee.com/inteligenciaeficiente)

---

## 📄 12. Licencia

Este proyecto está bajo la Licencia **MIT**. Eres libre de usar, modificar y distribuir este código tanto para fines personales como comerciales. Consulta el archivo [`LICENSE`](./LICENSE) para más detalles.
