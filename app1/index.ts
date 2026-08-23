/**
 * ============================================================================
 * MAPA INTERACTIVO DE CLIENTES (EJEMPLO 1)
 * Microsoft Script Lab + Excel + Leaflet.js
 * ============================================================================
 * 
 * 📋 CONFIGURACIÓN PARA SCRIPT LAB:
 * En la pestaña 'Libraries' de Script Lab, puedes incluir:
 *   https://unpkg.com/leaflet@1.9.4/dist/leaflet.js
 *   https://unpkg.com/leaflet@1.9.4/dist/leaflet.css
 *   https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css
 *   https://appsforoffice.microsoft.com/lib/1/hosted/office.js
 *   @types/office-js
 * ============================================================================
 */

// Declaraciones de tipos para librerías globales
declare const L: any;
declare const Office: any;
declare const Excel: any;

// Interface para el modelo de datos de clientes
interface Cliente {
  id: string;
  nombre: string;
  contacto: string;
  ciudad: string;
  lat: number;
  lng: number;
  zona: string;
  venta: number;
  prioridad: string;
}

// Configuración de colores de zonas
const ZONE_COLORS: Record<string, string> = {
  "Norte": "#2563eb",
  "Centro": "#0284c7",
  "Sur": "#059669",
  "Este": "#d97706",
  "Oeste": "#7c3aed",
  "VIP Express": "#e11d48"
};

// Nombres de tablas en Excel
const NOMBRE_TABLA_EXCEL = "Datos_Clientes";
const NOMBRE_TABLA_RUTA = "Ruta_Generada";

// Dataset inicial (Espejo de clientes_demo.csv para inicialización automática)
const DATASET_INICIAL: Cliente[] = [
  { id: "CLI-001", nombre: "Supermercado La Sierra", contacto: "Roberto Méndez", ciudad: "Madrid", lat: 40.4532, lng: -3.6883, zona: "Norte", venta: 12500, prioridad: "Alta" },
  { id: "CLI-002", nombre: "Farmacias San Miguel", contacto: "Lucía Gómez", ciudad: "Madrid", lat: 40.4611, lng: -3.6912, zona: "Norte", venta: 8900, prioridad: "Media" },
  { id: "CLI-003", nombre: "Distribuidora Del Valle", contacto: "Carlos Ortiz", ciudad: "Madrid", lat: 40.4485, lng: -3.6720, zona: "Norte", venta: 15400, prioridad: "Alta" },
  { id: "CLI-004", nombre: "Restaurante El Mirador", contacto: "Elena Ramos", ciudad: "Madrid", lat: 40.4705, lng: -3.7021, zona: "Norte", venta: 6200, prioridad: "Baja" },
  { id: "CLI-005", nombre: "Tecnología & Redes Global", contacto: "Marcos Díaz", ciudad: "Madrid", lat: 40.4182, lng: -3.7035, zona: "Centro", venta: 21000, prioridad: "Alta" },
  { id: "CLI-006", nombre: "Boutique Castellana", contacto: "Sofía Morales", ciudad: "Madrid", lat: 40.4250, lng: -3.6910, zona: "Centro", venta: 9800, prioridad: "Media" },
  { id: "CLI-007", nombre: "Cafetería Plaza Mayor", contacto: "Javier Santos", ciudad: "Madrid", lat: 40.4155, lng: -3.7074, zona: "Centro", venta: 5400, prioridad: "Baja" },
  { id: "CLI-008", nombre: "Librería Central Sol", contacto: "Ana Beltrán", ciudad: "Madrid", lat: 40.4168, lng: -3.7038, zona: "Centro", venta: 7300, prioridad: "Media" },
  { id: "CLI-009", nombre: "Automotriz del Sur", contacto: "Pedro Castillo", ciudad: "Madrid", lat: 40.3882, lng: -3.7125, zona: "Sur", venta: 18200, prioridad: "Alta" },
  { id: "CLI-010", nombre: "Almacenes Del Río", contacto: "Isabel Núñez", ciudad: "Madrid", lat: 40.3754, lng: -3.6980, zona: "Sur", venta: 11300, prioridad: "Media" },
  { id: "CLI-011", nombre: "Centro Médico Usera", contacto: "David Romero", ciudad: "Madrid", lat: 40.3820, lng: -3.7050, zona: "Sur", venta: 14100, prioridad: "Alta" },
  { id: "CLI-012", nombre: "Panificadora Imperial", contacto: "Carmen Vega", ciudad: "Madrid", lat: 40.3690, lng: -3.7180, zona: "Sur", venta: 4900, prioridad: "Baja" },
  { id: "CLI-013", nombre: "Muebles & Diseño Este", contacto: "Fernando Torres", ciudad: "Madrid", lat: 40.4280, lng: -3.6520, zona: "Este", venta: 13700, prioridad: "Media" },
  { id: "CLI-014", nombre: "Logística San Blas", contacto: "Raquel Pardo", ciudad: "Madrid", lat: 40.4350, lng: -3.6300, zona: "Este", venta: 22500, prioridad: "Alta" },
  { id: "CLI-015", nombre: "Construcciones Vicálvaro", contacto: "Hugo Lozano", ciudad: "Madrid", lat: 40.4050, lng: -3.6100, zona: "Este", venta: 16800, prioridad: "Alta" },
  { id: "CLI-016", nombre: "Clínica Dental Poniente", contacto: "Patricia Gil", ciudad: "Madrid", lat: 40.4390, lng: -3.7350, zona: "Oeste", venta: 9200, prioridad: "Media" },
  { id: "CLI-017", nombre: "Óptica Moncloa", contacto: "Andrés Serrano", ciudad: "Madrid", lat: 40.4355, lng: -3.7195, zona: "Oeste", venta: 7800, prioridad: "Baja" },
  { id: "CLI-018", nombre: "Electrónica Puerta de Hierro", contacto: "Marta Flores", ciudad: "Madrid", lat: 40.4620, lng: -3.7410, zona: "Oeste", venta: 19400, prioridad: "Alta" }
];

// Estado Global de la Aplicación
let clientes: Cliente[] = [...DATASET_INICIAL];
let map: any = null;
let markersLayerGroup: any = null;
let savedZonesLayerGroup: any = null;
let clientMarkerMap = new Map<string, any>();
let activeZoneFilter: string = "all";

// Almacén de geometrías de zonas guardadas por el usuario: { zona: { center: [lat, lng], radius: metros } }
// Inicialmente vacío: no se dibuja ningún círculo hasta que el usuario guarde una zona.
let savedZoneGeometries = new Map<string, { center: [number, number]; radius: number }>();

// Estado de Asignación de Zonas (Efecto WOW interactivo)
let isZoneModeActive: boolean = false;
let zoneCenterMarker: any = null;
let zoneCircle: any = null;
let clientsInRadius: Cliente[] = [];

// Estado de Rutas
let isRouteModeActive: boolean = false;
let routePolyline: any = null;

// ============================================================================
// CARGADOR AUTOMÁTICO DE DEPENDENCIAS EXTERNAS
// ============================================================================
async function asegurarDependencias(): Promise<void> {
  const fuentes = "https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap";
  const leafletCss = "https://unpkg.com/leaflet@1.9.4/dist/leaflet.css";
  const fontAwesomeCss = "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css";

  [fuentes, leafletCss, fontAwesomeCss].forEach(href => {
    if (!document.querySelector(`link[href="${href}"]`)) {
      const link = document.createElement("link");
      link.rel = "stylesheet";
      link.href = href;
      document.head.appendChild(link);
    }
  });

  if (typeof (window as any).L === "undefined") {
    await new Promise<void>((resolve, reject) => {
      const script = document.createElement("script");
      script.src = "https://unpkg.com/leaflet@1.9.4/dist/leaflet.js";
      script.crossOrigin = "";
      script.onload = () => resolve();
      script.onerror = () => reject(new Error("No se pudo cargar Leaflet desde CDN"));
      document.head.appendChild(script);
    });
  }
}

// ============================================================================
// INICIALIZACIÓN PRINCIPAL
// ============================================================================
async function iniciarAplicacion(): Promise<void> {
  try {
    await asegurarDependencias();
    inicializarMapa();
    renderizarMarcadores();
    renderizarZonasGuardadas();
    actualizarMetricas();
    configurarEventosUI();

    // Inicialización de Office.js para Script Lab
    if (typeof Office !== "undefined" && Office.onReady) {
      Office.onReady(async (info: any) => {
        if (info.host === Office.HostType.Excel) {
          await inicializarDatosDesdeExcel();
        }
      });
    }
  } catch (error) {
    console.error("Error al inicializar la aplicación:", error);
    mostrarToast("Error al cargar librerías de mapa", true);
  }
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", () => iniciarAplicacion());
} else {
  iniciarAplicacion();
}

/**
 * Inicializa el mapa Leaflet con CartoDB Positron (Tema Claro Minimalista)
 */
function inicializarMapa(): void {
  const mapContainer = document.getElementById("map");
  if (!mapContainer) return;

  if (map) {
    map.remove();
  }

  map = L.map("map", {
    center: [40.425, -3.690],
    zoom: 12,
    zoomControl: false
  });

  L.control.zoom({ position: "topleft" }).addTo(map);

  L.tileLayer("https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png", {
    attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors &copy; <a href="https://carto.com/attributions">CARTO</a>',
    subdomains: "abcd",
    maxZoom: 19
  }).addTo(map);

  savedZonesLayerGroup = L.layerGroup().addTo(map);
  markersLayerGroup = L.layerGroup().addTo(map);
}

/**
 * Renderiza ÚNICAMENTE los círculos de zonas que el usuario haya guardado explícitamente.
 */
function renderizarZonasGuardadas(): void {
  if (!savedZonesLayerGroup) return;

  savedZonesLayerGroup.clearLayers();

  if (savedZoneGeometries.size === 0) {
    return;
  }

  savedZoneGeometries.forEach((geo, nombreZona) => {
    if (activeZoneFilter !== "all" && activeZoneFilter.toLowerCase() !== nombreZona.toLowerCase()) {
      return;
    }

    const clientesZona = clientes.filter(c => c.zona.toLowerCase() === nombreZona.toLowerCase());
    const color = ZONE_COLORS[nombreZona] || "#2563eb";

    const circle = L.circle(geo.center, {
      radius: geo.radius,
      color: color,
      fillColor: color,
      fillOpacity: 0.10,
      weight: 1.5,
      dashArray: "4, 4"
    });

    circle.bindTooltip(`Zona ${nombreZona} (${clientesZona.length} clientes)`, {
      permanent: false,
      direction: "top",
      className: "zone-circle-tooltip"
    });

    savedZonesLayerGroup.addLayer(circle);
  });
}

/**
 * Renderiza los marcadores de clientes según el filtro activo
 */
function renderizarMarcadores(): void {
  if (!markersLayerGroup) return;

  markersLayerGroup.clearLayers();
  clientMarkerMap.clear();

  const clientesVisibles = activeZoneFilter === "all"
    ? clientes
    : clientes.filter(c => c.zona.toLowerCase() === activeZoneFilter.toLowerCase());

  clientesVisibles.forEach(cliente => {
    const color = ZONE_COLORS[cliente.zona] || "#2563eb";

    const customIcon = L.divIcon({
      className: "custom-client-icon-wrapper",
      html: `
        <div class="custom-client-marker" id="marker-${cliente.id}" style="border-color: ${color}; color: ${color};">
          <i class="fa-solid fa-shop"></i>
        </div>
      `,
      iconSize: [28, 28],
      iconAnchor: [14, 14],
      popupAnchor: [0, -14]
    });

    const marker = L.marker([cliente.lat, cliente.lng], { icon: customIcon });

    const popupContent = `
      <div class="client-popup-card">
        <div class="popup-header">
          <span class="popup-title">${cliente.nombre}</span>
          <span class="popup-zone-badge" style="background-color: ${color}">${cliente.zona}</span>
        </div>
        <div class="popup-detail-row">
          <span>Contacto:</span>
          <strong>${cliente.contacto}</strong>
        </div>
        <div class="popup-detail-row">
          <span>Prioridad:</span>
          <strong>${cliente.prioridad}</strong>
        </div>
        <div class="popup-detail-row">
          <span>Venta Mensual:</span>
          <span class="popup-sales">$${cliente.venta.toLocaleString()} USD</span>
        </div>
      </div>
    `;

    marker.bindPopup(popupContent);
    markersLayerGroup.addLayer(marker);
    clientMarkerMap.set(cliente.id, marker);
  });
}

/**
 * Actualiza contadores y métricas en pantalla
 */
function actualizarMetricas(): void {
  const visibles = activeZoneFilter === "all"
    ? clientes
    : clientes.filter(c => c.zona.toLowerCase() === activeZoneFilter.toLowerCase());

  const totalVentas = visibles.reduce((acc, c) => acc + c.venta, 0);
  const elemVisibles = document.getElementById("stat-visible-clients");
  const elemVentas = document.getElementById("stat-total-sales");
  if (elemVisibles) elemVisibles.textContent = visibles.length.toString();
  if (elemVentas) elemVentas.textContent = `$${totalVentas.toLocaleString()}`;

  const countAll = document.getElementById("count-all");
  const countNorte = document.getElementById("count-norte");
  const countCentro = document.getElementById("count-centro");
  const countSur = document.getElementById("count-sur");
  const countEste = document.getElementById("count-este");
  const countOeste = document.getElementById("count-oeste");

  if (countAll) countAll.textContent = clientes.length.toString();
  if (countNorte) countNorte.textContent = clientes.filter(c => c.zona === "Norte").length.toString();
  if (countCentro) countCentro.textContent = clientes.filter(c => c.zona === "Centro").length.toString();
  if (countSur) countSur.textContent = clientes.filter(c => c.zona === "Sur").length.toString();
  if (countEste) countEste.textContent = clientes.filter(c => c.zona === "Este").length.toString();
  if (countOeste) countOeste.textContent = clientes.filter(c => c.zona === "Oeste").length.toString();
}

// ============================================================================
// LECTURA Y CREACIÓN AUTOMÁTICA EN EXCEL (Office.js)
// ============================================================================

/**
 * Busca automáticamente la tabla 'Datos_Clientes'. Si no existe, la crea con el dataset inicial.
 */
async function inicializarDatosDesdeExcel(): Promise<void> {
  if (typeof Excel === "undefined" || !Excel.run) return;

  try {
    await Excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const tables = sheet.tables;
      tables.load("items/name");
      await context.sync();

      let targetTable = tables.items.find((t: any) => t.name === NOMBRE_TABLA_EXCEL);

      if (!targetTable) {
        const headers = [["ID", "Cliente", "Contacto", "Ciudad", "Latitud", "Longitud", "Zona", "Venta_USD", "Prioridad"]];
        const dataRows = DATASET_INICIAL.map(c => [
          c.id,
          c.nombre,
          c.contacto,
          c.ciudad,
          c.lat,
          c.lng,
          c.zona,
          c.venta,
          c.prioridad
        ]);

        const endRow = 1 + dataRows.length;
        const targetRangeAddress = `A1:I${endRow}`;
        const range = sheet.getRange(targetRangeAddress);
        range.values = [...headers, ...dataRows];

        targetTable = sheet.tables.add(targetRangeAddress, true);
        targetTable.name = NOMBRE_TABLA_EXCEL;
        targetTable.style = "TableStyleLight9";

        range.format.autofitColumns();
        await context.sync();

        clientes = [...DATASET_INICIAL];
        renderizarMarcadores();
        renderizarZonasGuardadas();
        actualizarMetricas();
        mostrarToast(`Tabla '${NOMBRE_TABLA_EXCEL}' creada y cargada en Excel`);
      } else {
        await leerTablaClientes(targetTable, context);
      }
    });
  } catch (error) {
    console.error("Error al inicializar tabla en Excel:", error);
  }
}

/**
 * Lee los datos de una tabla de Excel y actualiza el mapa
 */
async function leerTablaClientes(table: any, context: any): Promise<void> {
  const range = table.getDataBodyRange().load("values");
  const headers = table.getHeaderRowRange().load("values");
  await context.sync();

  const headerRow = headers.values[0].map((h: string) => h.toString().toLowerCase().trim());
  const rows = range.values;

  const idIdx = headerRow.findIndex((h: string) => h === "id");
  const nameIdx = headerRow.findIndex((h: string) => h.includes("client") || h.includes("nombre"));
  const contactIdx = headerRow.findIndex((h: string) => h.includes("contacto"));
  const cityIdx = headerRow.findIndex((h: string) => h.includes("ciudad"));
  const latIdx = headerRow.findIndex((h: string) => h.includes("lat"));
  const lngIdx = headerRow.findIndex((h: string) => h.includes("long") || h.includes("lng"));
  const zoneIdx = headerRow.findIndex((h: string) => h.includes("zona") || h.includes("region"));
  const salesIdx = headerRow.findIndex((h: string) => h.includes("venta") || h.includes("usd") || h.includes("total"));
  const priorityIdx = headerRow.findIndex((h: string) => h.includes("prioridad"));

  if (latIdx !== -1 && lngIdx !== -1) {
    const nuevosClientes: Cliente[] = [];
    rows.forEach((row: any[], i: number) => {
      const lat = parseFloat(row[latIdx]);
      const lng = parseFloat(row[lngIdx]);
      if (!isNaN(lat) && !isNaN(lng)) {
        nuevosClientes.push({
          id: idIdx !== -1 ? String(row[idIdx]) : `CLI-${String(i + 1).padStart(3, "0")}`,
          nombre: nameIdx !== -1 ? String(row[nameIdx]) : `Cliente ${i + 1}`,
          contacto: contactIdx !== -1 ? String(row[contactIdx]) : "Contacto",
          ciudad: cityIdx !== -1 ? String(row[cityIdx]) : "Madrid",
          lat: lat,
          lng: lng,
          zona: zoneIdx !== -1 ? String(row[zoneIdx]) : "Centro",
          venta: salesIdx !== -1 ? Number(row[salesIdx]) || 5000 : 5000,
          prioridad: priorityIdx !== -1 ? String(row[priorityIdx]) : "Media"
        });
      }
    });

    if (nuevosClientes.length > 0) {
      clientes = nuevosClientes;
      savedZoneGeometries.clear();
      renderizarMarcadores();
      renderizarZonasGuardadas();
      actualizarMetricas();
      mostrarToast(`Sincronizados ${nuevosClientes.length} clientes desde '${table.name || NOMBRE_TABLA_EXCEL}'`);
    }
  }
}

/**
 * Lee datos desde la tabla 'Datos_Clientes' o la selección activa en Excel
 */
async function sincronizarDesdeExcel(): Promise<void> {
  if (typeof Excel !== "undefined" && Excel.run) {
    try {
      await Excel.run(async (context: any) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const tables = sheet.tables;
        tables.load("items/name");
        await context.sync();

        let targetTable = tables.items.find((t: any) => t.name === NOMBRE_TABLA_EXCEL);
        if (targetTable) {
          await leerTablaClientes(targetTable, context);
        } else if (tables.items.length > 0) {
          await leerTablaClientes(tables.items[0], context);
        } else {
          await inicializarDatosDesdeExcel();
        }
      });
    } catch (error) {
      console.error("Error al sincronizar con Excel:", error);
      mostrarToast("Datos sincronizados en modo demo");
    }
  } else {
    mostrarToast("Sincronización simulada en navegador");
    renderizarMarcadores();
    renderizarZonasGuardadas();
    actualizarMetricas();
  }
}

/**
 * Actualiza la columna Zona en la tabla 'Datos_Clientes' de Excel
 */
async function sincronizarCambiosHaciaExcel(idsActualizados: Set<string>, nuevaZona: string): Promise<void> {
  await Excel.run(async (context: any) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const tables = sheet.tables;
    tables.load("items/name");
    await context.sync();

    let table = tables.items.find((t: any) => t.name === NOMBRE_TABLA_EXCEL) || tables.items[0];
    if (table) {
      const idColumn = table.columns.getItemOrNullObject("ID");
      const zonaColumn = table.columns.getItemOrNullObject("Zona");
      await context.sync();

      if (!idColumn.isNullObject && !zonaColumn.isNullObject) {
        const idRange = idColumn.getDataBodyRange().load("values");
        const zonaRange = zonaColumn.getDataBodyRange().load("values");
        await context.sync();

        const idValues = idRange.values;
        const zonaValues = zonaRange.values;

        for (let i = 0; i < idValues.length; i++) {
          const rowId = String(idValues[i][0]);
          if (idsActualizados.has(rowId)) {
            zonaValues[i][0] = nuevaZona;
          }
        }
        zonaRange.values = zonaValues;
        await context.sync();
      }
    }
  });
}

// ============================================================================
// ASIGNACIÓN DINÁMICA DE ZONAS (EFECTO WOW POR RADIO)
// ============================================================================

/**
 * Activa el modo interactivo para crear o asignar una zona con radio dinámico
 */
function activarModoZona(): void {
  isZoneModeActive = true;
  const zonePanel = document.getElementById("zone-panel");
  const btnZoneMode = document.getElementById("btn-zone-mode");
  if (zonePanel) zonePanel.classList.remove("hidden");
  if (btnZoneMode) btnZoneMode.classList.add("active");

  desactivarModoRuta();

  const center = map.getCenter();
  const radiusKm = parseFloat((document.getElementById("radius-slider") as HTMLInputElement)?.value || "3.5");
  const radiusMeters = radiusKm * 1000;

  const centerIcon = L.divIcon({
    className: "center-zone-pin-wrapper",
    html: `<div class="center-zone-pin"><i class="fa-solid fa-crosshairs"></i></div>`,
    iconSize: [34, 34],
    iconAnchor: [17, 34]
  });

  zoneCenterMarker = L.marker(center, {
    draggable: true,
    icon: centerIcon,
    zIndexOffset: 2000
  }).addTo(map);

  const zoneSelect = document.getElementById("zone-select") as HTMLSelectElement;
  const color = zoneSelect?.options[zoneSelect.selectedIndex]?.getAttribute("data-color") || "#2563eb";

  zoneCircle = L.circle(center, {
    radius: radiusMeters,
    color: color,
    fillColor: color,
    fillOpacity: 0.18,
    weight: 2,
    dashArray: "6, 6"
  }).addTo(map);

  zoneCenterMarker.on("drag", (e: any) => {
    const newPos = e.target.getLatLng();
    zoneCircle.setLatLng(newPos);
    evaluarClientesEnRadio();
  });

  evaluarClientesEnRadio();
}

/**
 * Desactiva y limpia el modo de asignación de zona
 */
function desactivarModoZona(): void {
  isZoneModeActive = false;
  const zonePanel = document.getElementById("zone-panel");
  const btnZoneMode = document.getElementById("btn-zone-mode");
  if (zonePanel) zonePanel.classList.add("hidden");
  if (btnZoneMode) btnZoneMode.classList.remove("active");

  if (zoneCenterMarker) {
    map.removeLayer(zoneCenterMarker);
    zoneCenterMarker = null;
  }
  if (zoneCircle) {
    map.removeLayer(zoneCircle);
    zoneCircle = null;
  }

  clientes.forEach(c => {
    const el = document.getElementById(`marker-${c.id}`);
    if (el) el.classList.remove("in-radius-preview");
  });
}

/**
 * Evalúa qué clientes caen dentro del radio del círculo y los resalta en vivo
 */
function evaluarClientesEnRadio(): void {
  if (!zoneCenterMarker || !zoneCircle) return;

  const centerLatLng = zoneCenterMarker.getLatLng();
  const radiusMeters = zoneCircle.getRadius();
  clientsInRadius = [];

  clientes.forEach(cliente => {
    const clientLatLng = L.latLng(cliente.lat, cliente.lng);
    const distMeters = map.distance(centerLatLng, clientLatLng);
    const el = document.getElementById(`marker-${cliente.id}`);

    if (distMeters <= radiusMeters) {
      clientsInRadius.push(cliente);
      if (el) el.classList.add("in-radius-preview");
    } else {
      if (el) el.classList.remove("in-radius-preview");
    }
  });

  const previewCount = document.getElementById("preview-count");
  if (previewCount) {
    previewCount.textContent = clientsInRadius.length.toString();
  }
}

/**
 * Guarda y confirma la asignación de la zona para los clientes dentro del radio.
 */
async function guardarAsignacionZona(): Promise<void> {
  if (clientsInRadius.length === 0) {
    mostrarToast("No hay clientes dentro del radio seleccionado", true);
    return;
  }

  const selectZone = document.getElementById("zone-select") as HTMLSelectElement;
  const nuevaZona = selectZone ? selectZone.value : "Norte";

  if (zoneCenterMarker && zoneCircle) {
    const c = zoneCenterMarker.getLatLng();
    savedZoneGeometries.set(nuevaZona, {
      center: [c.lat, c.lng],
      radius: zoneCircle.getRadius()
    });
  }

  const idsActualizados = new Set(clientsInRadius.map(c => c.id));
  clientes.forEach(c => {
    if (idsActualizados.has(c.id)) {
      c.zona = nuevaZona;
    }
  });

  if (typeof Excel !== "undefined" && Excel.run) {
    try {
      await sincronizarCambiosHaciaExcel(idsActualizados, nuevaZona);
    } catch (err) {
      console.warn("No se pudo escribir en Excel directamente:", err);
    }
  }

  mostrarToast(`¡${clientsInRadius.length} clientes asignados a Zona ${nuevaZona}!`);
  desactivarModoZona();
  renderizarMarcadores();
  renderizarZonasGuardadas();
  actualizarMetricas();
}

// ============================================================================
// TRAZADO Y DESCARGA DE RUTAS
// ============================================================================

/**
 * Obtiene la lista de clientes ordenados para la ruta actual
 */
function obtenerRutaOrdenada(): Cliente[] {
  const clientesVisibles = activeZoneFilter === "all"
    ? clientes
    : clientes.filter(c => c.zona.toLowerCase() === activeZoneFilter.toLowerCase());

  return [...clientesVisibles].sort((a, b) => b.lat - a.lat);
}

/**
 * Alterna el trazado de la ruta de visitas entre los clientes visibles
 */
function alternarModoRuta(): void {
  if (isRouteModeActive) {
    desactivarModoRuta();
  } else {
    activarModoRuta();
  }
}

/**
 * Traza la polilínea de la ruta conectando los clientes visibles
 */
function activarModoRuta(): void {
  desactivarModoZona();
  isRouteModeActive = true;

  const routePanel = document.getElementById("route-panel");
  const btnRouteMode = document.getElementById("btn-route-mode");
  if (routePanel) routePanel.classList.remove("hidden");
  if (btnRouteMode) btnRouteMode.classList.add("active");

  const rutaOrdenada = obtenerRutaOrdenada();

  if (rutaOrdenada.length < 2) {
    mostrarToast("Se requieren al menos 2 clientes para trazar una ruta", true);
    return;
  }

  const latLngs = rutaOrdenada.map(c => [c.lat, c.lng]);

  if (routePolyline) {
    map.removeLayer(routePolyline);
  }

  routePolyline = L.polyline(latLngs, {
    color: "#2563eb",
    weight: 4,
    opacity: 0.85,
    dashArray: "8, 8",
    lineJoin: "round"
  }).addTo(map);

  let distanciaMetros = 0;
  for (let i = 0; i < rutaOrdenada.length - 1; i++) {
    distanciaMetros += map.distance(
      [rutaOrdenada[i].lat, rutaOrdenada[i].lng],
      [rutaOrdenada[i + 1].lat, rutaOrdenada[i + 1].lng]
    );
  }

  const distanciaKm = (distanciaMetros / 1000).toFixed(1);
  const tiempoMin = Math.round(distanciaMetros / 1000 * 3.5);

  const elemStops = document.getElementById("route-stops");
  const elemDist = document.getElementById("route-distance");
  const elemTime = document.getElementById("route-time");

  if (elemStops) elemStops.textContent = rutaOrdenada.length.toString();
  if (elemDist) elemDist.textContent = `${distanciaKm} km`;
  if (elemTime) elemTime.textContent = `${tiempoMin} min`;

  map.fitBounds(routePolyline.getBounds(), { padding: [40, 40] });
  mostrarToast(`Ruta generada: ${distanciaKm} km (${rutaOrdenada.length} paradas)`);
}

/**
 * Desactiva y remueve la polilínea de ruta
 */
function desactivarModoRuta(): void {
  isRouteModeActive = false;
  const routePanel = document.getElementById("route-panel");
  const btnRouteMode = document.getElementById("btn-route-mode");
  if (routePanel) routePanel.classList.add("hidden");
  if (btnRouteMode) btnRouteMode.classList.remove("active");

  if (routePolyline) {
    map.removeLayer(routePolyline);
    routePolyline = null;
  }
}

/**
 * Exporta la ruta a una nueva tabla 'Ruta_Generada' en Excel,
 * colocada exactamente 2 columnas a la derecha de la tabla 'Datos_Clientes'.
 */
async function descargarRutaAExcel(): Promise<void> {
  const rutaOrdenada = obtenerRutaOrdenada();

  if (rutaOrdenada.length < 2) {
    mostrarToast("Genera una ruta antes de descargar", true);
    return;
  }

  const headers = [
    "Orden",
    "ID_Cliente",
    "Cliente",
    "Ciudad",
    "Zona",
    "Latitud",
    "Longitud",
    "Distancia_Tramo_km",
    "Distancia_Acumulada_km",
    "Tiempo_Estimado_min"
  ];

  let distAcumulada = 0;
  let tiempoAcumulado = 0;

  const dataRows = rutaOrdenada.map((c, i) => {
    let distTramo = 0;
    if (i > 0) {
      const prev = rutaOrdenada[i - 1];
      distTramo = map.distance([prev.lat, prev.lng], [c.lat, c.lng]) / 1000;
    }
    distAcumulada += distTramo;
    tiempoAcumulado += Math.round(distTramo * 3.5);

    return [
      i + 1,
      c.id,
      c.nombre,
      c.ciudad,
      c.zona,
      c.lat,
      c.lng,
      parseFloat(distTramo.toFixed(2)),
      parseFloat(distAcumulada.toFixed(2)),
      tiempoAcumulado
    ];
  });

  // Notificar al simulador de Excel (GitHub Pages) si está en iframe
  if (typeof window !== "undefined" && window.parent && window.parent !== window) {
    window.parent.postMessage({ type: "APP1_EXPORT_ROUTE", route: dataRows }, "*");
  }

  // 1. Si se ejecuta dentro de Microsoft Excel (Office.js)
  if (typeof Excel !== "undefined" && Excel.run) {
    try {
      await Excel.run(async (context: any) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const tables = sheet.tables;
        tables.load("items/name");
        await context.sync();

        let tableClientes = tables.items.find((t: any) => t.name === NOMBRE_TABLA_EXCEL) || tables.items[0];

        let startCol = 11; // Columna L por defecto
        let startRow = 0;  // Fila 1 por defecto

        if (tableClientes) {
          const rangeClientes = tableClientes.getRange().load(["columnIndex", "columnCount", "rowIndex", "rowCount"]);
          await context.sync();
          startCol = rangeClientes.columnIndex + rangeClientes.columnCount + 2;
          startRow = rangeClientes.rowIndex;
        }

        let existingRutaTable = tables.items.find((t: any) => t.name === NOMBRE_TABLA_RUTA);
        if (existingRutaTable) {
          existingRutaTable.delete();
          await context.sync();
        }

        const totalRows = dataRows.length + 1; // +1 para headers
        const totalCols = headers.length;

        const startCell = sheet.getCell(startRow, startCol);
        const targetRange = startCell.getResizedRange(totalRows - 1, totalCols - 1);
        targetRange.values = [headers, ...dataRows];

        const newRutaTable = sheet.tables.add(targetRange, true /* hasHeaders */);
        newRutaTable.name = NOMBRE_TABLA_RUTA;
        newRutaTable.style = "TableStyleLight10"; // Estilo limpio en tonos azul suave

        targetRange.format.autofitColumns();
        await context.sync();

        mostrarToast(`¡Ruta guardada en la tabla '${NOMBRE_TABLA_RUTA}' en Excel!`);
      });
      return;
    } catch (error: any) {
      console.error("Error al exportar tabla a Excel:", error);
      mostrarToast(`Error en Excel: ${error?.message || error}`, true);
    }
  } else {
    // 2. Fallback para navegador independiente: Descarga CSV directa
    try {
      const csvContent = [headers.join(","), ...dataRows.map(r => r.map(v => `"${v}"`).join(","))].join("\n");
      const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.setAttribute("href", url);
      link.setAttribute("download", `${NOMBRE_TABLA_RUTA}.csv`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      mostrarToast(`Ruta descargada como '${NOMBRE_TABLA_RUTA}.csv'`);
    } catch (err) {
      console.error("Error al descargar CSV:", err);
      mostrarToast("Error al exportar CSV", true);
    }
  }
}

// ============================================================================
// CONFIGURACIÓN DE EVENTOS DE INTERFAZ
// ============================================================================
function configurarEventosUI(): void {
  // Botones principales
  document.getElementById("btn-sync")?.addEventListener("click", sincronizarDesdeExcel);
  document.getElementById("btn-zone-mode")?.addEventListener("click", () => {
    if (isZoneModeActive) {
      desactivarModoZona();
    } else {
      activarModoZona();
    }
  });
  document.getElementById("btn-route-mode")?.addEventListener("click", alternarModoRuta);

  // Panel de Zonas
  document.getElementById("btn-close-zone")?.addEventListener("click", desactivarModoZona);
  document.getElementById("btn-cancel-zone")?.addEventListener("click", desactivarModoZona);
  document.getElementById("btn-save-zone")?.addEventListener("click", guardarAsignacionZona);

  // Slider de Radio
  const slider = document.getElementById("radius-slider") as HTMLInputElement;
  slider?.addEventListener("input", (e: any) => {
    const val = parseFloat(e.target.value);
    const radiusDisplay = document.getElementById("radius-value");
    if (radiusDisplay) radiusDisplay.textContent = `${val.toFixed(1)} km`;

    if (zoneCircle) {
      zoneCircle.setRadius(val * 1000);
      evaluarClientesEnRadio();
    }
  });

  // Selector de Color / Zona
  const zoneSelect = document.getElementById("zone-select") as HTMLSelectElement;
  zoneSelect?.addEventListener("change", () => {
    const selectedOption = zoneSelect.options[zoneSelect.selectedIndex];
    const color = selectedOption.getAttribute("data-color") || "#2563eb";
    if (zoneCircle) {
      zoneCircle.setStyle({ color: color, fillColor: color });
    }
  });

  // Panel de Rutas y Botón de Descarga
  document.getElementById("btn-close-route")?.addEventListener("click", desactivarModoRuta);
  document.getElementById("btn-export-route")?.addEventListener("click", descargarRutaAExcel);

  // Filtros por Zona
  const filterChips = document.querySelectorAll(".filter-chip");
  filterChips.forEach(chip => {
    chip.addEventListener("click", (e: any) => {
      filterChips.forEach(c => c.classList.remove("active"));
      const target = e.currentTarget as HTMLElement;
      target.classList.add("active");
      activeZoneFilter = target.getAttribute("data-zone") || "all";

      renderizarMarcadores();
      renderizarZonasGuardadas();
      actualizarMetricas();

      if (isRouteModeActive) {
        activarModoRuta();
      }
    });
  });
}

/**
 * Muestra un mensaje toast flotante en la UI
 */
function mostrarToast(mensaje: string, isError: boolean = false): void {
  const toast = document.getElementById("toast");
  const toastMsg = document.getElementById("toast-msg");
  if (!toast || !toastMsg) return;

  toastMsg.textContent = mensaje;
  toast.style.backgroundColor = isError ? "#b91c1c" : "#0f172a";
  toast.classList.remove("hidden");

  setTimeout(() => {
    toast.classList.add("hidden");
  }, 3200);
}
