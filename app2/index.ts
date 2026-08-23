/**
 * ============================================================================
 * RASTREO GPS DE VEHÍCULOS (EJEMPLO 2)
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

// Interface para el modelo de datos de telemetría GPS
interface RegistroGPS {
  paso: number;
  hora: string;
  vehiculoId: string;
  conductor: string;
  lat: number;
  lng: number;
  velocidad: number;
  combustible: number;
  estado: string;
  ubicacion: string;
}

// Nombre de la tabla objetivo en Excel
const NOMBRE_TABLA_EXCEL = "Registro_GPS";

// Dataset inicial (Espejo de vehiculos_gps.csv para inicialización automática)
const DATASET_INICIAL: RegistroGPS[] = [
  { paso: 1, hora: "08:00", vehiculoId: "UNIDAD-01", conductor: "Carlos Mendoza", lat: 40.4850, lng: -3.6920, velocidad: 0, combustible: 98, estado: "Detenido", ubicacion: "Centro Logístico Norte" },
  { paso: 2, hora: "08:08", vehiculoId: "UNIDAD-01", conductor: "Carlos Mendoza", lat: 40.4780, lng: -3.6890, velocidad: 38, combustible: 97, estado: "En Movimiento", ubicacion: "Av. de Burgos" },
  { paso: 3, hora: "08:15", vehiculoId: "UNIDAD-01", conductor: "Carlos Mendoza", lat: 40.4670, lng: -3.6860, velocidad: 54, combustible: 95, estado: "En Movimiento", ubicacion: "Paseo de la Castellana Norte" },
  { paso: 4, hora: "08:22", vehiculoId: "UNIDAD-01", conductor: "Carlos Mendoza", lat: 40.4530, lng: -3.6880, velocidad: 42, combustible: 93, estado: "En Movimiento", ubicacion: "Plaza de Castilla" },
  { paso: 5, hora: "08:30", vehiculoId: "UNIDAD-01", conductor: "Carlos Mendoza", lat: 40.4510, lng: -3.6890, velocidad: 0, combustible: 92, estado: "En Entrega", ubicacion: "Cliente 1 - Supermercado Sierra" },
  { paso: 6, hora: "08:45", vehiculoId: "UNIDAD-01", conductor: "Carlos Mendoza", lat: 40.4420, lng: -3.6900, velocidad: 32, combustible: 90, estado: "En Movimiento", ubicacion: "Castellana / Nuevos Ministerios" },
  { paso: 7, hora: "08:52", vehiculoId: "UNIDAD-01", conductor: "Carlos Mendoza", lat: 40.4350, lng: -3.6890, velocidad: 28, combustible: 88, estado: "En Movimiento", ubicacion: "Paseo de Eduardo Dato" },
  { paso: 8, hora: "09:00", vehiculoId: "UNIDAD-01", conductor: "Carlos Mendoza", lat: 40.4280, lng: -3.6880, velocidad: 0, combustible: 87, estado: "En Entrega", ubicacion: "Cliente 2 - Boutique Castellana" },
  { paso: 9, hora: "09:15", vehiculoId: "UNIDAD-01", conductor: "Carlos Mendoza", lat: 40.4200, lng: -3.6950, velocidad: 22, combustible: 85, estado: "En Movimiento", ubicacion: "Paseo del Prado" },
  { paso: 10, hora: "09:22", vehiculoId: "UNIDAD-01", conductor: "Carlos Mendoza", lat: 40.4150, lng: -3.7050, velocidad: 15, combustible: 83, estado: "En Movimiento", ubicacion: "Plaza Mayor / Sol" },
  { paso: 11, hora: "09:30", vehiculoId: "UNIDAD-01", conductor: "Carlos Mendoza", lat: 40.4120, lng: -3.7080, velocidad: 0, combustible: 82, estado: "En Entrega", ubicacion: "Cliente 3 - Cafetería Central" },
  { paso: 12, hora: "09:45", vehiculoId: "UNIDAD-01", conductor: "Carlos Mendoza", lat: 40.4020, lng: -3.7060, velocidad: 45, combustible: 80, estado: "En Movimiento", ubicacion: "Ronda de Atocha" },
  { paso: 13, hora: "09:52", vehiculoId: "UNIDAD-01", conductor: "Carlos Mendoza", lat: 40.3920, lng: -3.7080, velocidad: 58, combustible: 78, estado: "En Movimiento", ubicacion: "Av. del Manzanares" },
  { paso: 14, hora: "10:00", vehiculoId: "UNIDAD-01", conductor: "Carlos Mendoza", lat: 40.3820, lng: -3.7120, velocidad: 0, combustible: 76, estado: "En Entrega", ubicacion: "Cliente 4 - Automotriz del Sur" },
  { paso: 15, hora: "10:18", vehiculoId: "UNIDAD-01", conductor: "Carlos Mendoza", lat: 40.3750, lng: -3.7020, velocidad: 62, combustible: 74, estado: "En Movimiento", ubicacion: "Autovía M-30 Sur" },
  { paso: 16, hora: "10:28", vehiculoId: "UNIDAD-01", conductor: "Carlos Mendoza", lat: 40.3880, lng: -3.6700, velocidad: 75, combustible: 71, estado: "En Movimiento", ubicacion: "M-30 Este" },
  { paso: 17, hora: "10:38", vehiculoId: "UNIDAD-01", conductor: "Carlos Mendoza", lat: 40.4150, lng: -3.6550, velocidad: 68, combustible: 68, estado: "En Movimiento", ubicacion: "M-30 / O'Donnell" },
  { paso: 18, hora: "10:48", vehiculoId: "UNIDAD-01", conductor: "Carlos Mendoza", lat: 40.4320, lng: -3.6350, velocidad: 48, combustible: 65, estado: "En Movimiento", ubicacion: "Av. de San Blas" },
  { paso: 19, hora: "10:55", vehiculoId: "UNIDAD-01", conductor: "Carlos Mendoza", lat: 40.4360, lng: -3.6300, velocidad: 0, combustible: 64, estado: "En Entrega", ubicacion: "Cliente 5 - Logística Este" },
  { paso: 20, hora: "11:15", vehiculoId: "UNIDAD-01", conductor: "Carlos Mendoza", lat: 40.4850, lng: -3.6920, velocidad: 0, combustible: 58, estado: "Detenido", ubicacion: "Retorno a Centro Logístico" }
];

// Estado Global de la Aplicación
let registros: RegistroGPS[] = [...DATASET_INICIAL];
let map: any = null;
let vehicleMarker: any = null;
let stopsLayerGroup: any = null;
let completedPolyline: any = null;
let remainingPolyline: any = null;

// Estado de Reproducción
let currentStepIndex: number = 0;
let isPlaying: boolean = false;
let playbackTimer: any = null;
let playbackSpeed: number = 1; // 1x, 2x, 4x
let showFullTrail: boolean = true;

// Estado de Resaltado en Excel
let lastHighlightedRowIndex: number = -1;
let isHighlightingInProgress: boolean = false;
let pendingHighlightIndex: number | null = null;

// ============================================================================
// CARGADOR AUTOMÁTICO DE DEPENDENCIAS EXTERNAS
// ============================================================================
async function asegurarDependencias(): Promise<void> {
  const fuentes = "https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap";
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
// CONVERSIÓN Y FORMATEO DE HORA (EXCEL SERIAL A HH:MM)
// ============================================================================
/**
 * Convierte un valor de hora proveniente de Excel (número decimal serial o string) a formato "HH:mm"
 */
function formatearHoraExcel(val: any, textVal?: string): string {
  if (textVal && typeof textVal === "string" && textVal.includes(":")) {
    const partes = textVal.trim().split(":");
    if (partes.length >= 2) {
      return `${partes[0].padStart(2, "0")}:${partes[1].padStart(2, "0")}`;
    }
  }

  if (typeof val === "number") {
    const totalMinutes = Math.round(val * 24 * 60);
    const hours = Math.floor(totalMinutes / 60) % 24;
    const mins = totalMinutes % 60;
    return `${String(hours).padStart(2, "0")}:${String(mins).padStart(2, "0")}`;
  }

  if (typeof val === "string") {
    const num = parseFloat(val);
    if (!isNaN(num) && num >= 0 && num <= 1) {
      const totalMinutes = Math.round(num * 24 * 60);
      const hours = Math.floor(totalMinutes / 60) % 24;
      const mins = totalMinutes % 60;
      return `${String(hours).padStart(2, "0")}:${String(mins).padStart(2, "0")}`;
    }
    return val.trim();
  }

  return "08:00";
}

// ============================================================================
// INICIALIZACIÓN PRINCIPAL
// ============================================================================
async function iniciarAplicacion(): Promise<void> {
  try {
    await asegurarDependencias();
    inicializarMapa();
    actualizarPaso(0);
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
 * Inicializa el mapa Leaflet con CartoDB Positron
 */
function inicializarMapa(): void {
  const mapContainer = document.getElementById("map");
  if (!mapContainer) return;

  if (map) {
    map.remove();
  }

  const primerPunto = registros[0] || { lat: 40.435, lng: -3.685 };

  map = L.map("map", {
    center: [primerPunto.lat, primerPunto.lng],
    zoom: 13,
    zoomControl: false
  });

  L.control.zoom({ position: "topleft" }).addTo(map);

  L.tileLayer("https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png", {
    attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors &copy; <a href="https://carto.com/attributions">CARTO</a>',
    subdomains: "abcd",
    maxZoom: 19
  }).addTo(map);

  stopsLayerGroup = L.layerGroup().addTo(map);

  // Polilíneas de recorrido
  completedPolyline = L.polyline([], {
    color: "#2563eb",
    weight: 4,
    opacity: 0.9,
    lineJoin: "round"
  }).addTo(map);

  remainingPolyline = L.polyline([], {
    color: "#93c5fd",
    weight: 2,
    dashArray: "4, 6",
    opacity: 0.6,
    lineJoin: "round"
  }).addTo(map);

  // Marcador dinámico del vehículo con onda de radar
  const vehicleIcon = L.divIcon({
    className: "custom-vehicle-marker-wrapper",
    html: `
      <div class="vehicle-pulse-marker">
        <div class="vehicle-radar-wave"></div>
        <div class="vehicle-marker-core">
          <i class="fa-solid fa-truck"></i>
        </div>
      </div>
    `,
    iconSize: [36, 36],
    iconAnchor: [18, 18]
  });

  vehicleMarker = L.marker([primerPunto.lat, primerPunto.lng], {
    icon: vehicleIcon,
    zIndexOffset: 1000
  }).addTo(map);

  renderizarPuntosDeParada();
}

/**
 * Renderiza los pequeños hitos de parada en el mapa
 */
function renderizarPuntosDeParada(): void {
  if (!stopsLayerGroup) return;
  stopsLayerGroup.clearLayers();

  registros.forEach((reg, i) => {
    const esEntrega = reg.estado === "En Entrega";

    const stopIcon = L.divIcon({
      className: "custom-stop-point",
      html: `<div class="stop-point-marker ${esEntrega ? "stop-delivery" : ""}"></div>`,
      iconSize: [10, 10],
      iconAnchor: [5, 5]
    });

    const marker = L.marker([reg.lat, reg.lng], { icon: stopIcon });
    marker.bindTooltip(`<b>${reg.hora}</b> - ${reg.ubicacion}<br>Velocidad: ${reg.velocidad} km/h`, {
      direction: "top",
      offset: [0, -6]
    });

    stopsLayerGroup.addLayer(marker);
  });
}

/**
 * Resalta en amarillo suave la fila actual en la tabla de Excel
 * y remueve el formato de la fila anterior.
 */
async function resaltarFilaEnExcel(rowIndex: number): Promise<void> {
  if (typeof Excel === "undefined" || !Excel.run) return;

  if (isHighlightingInProgress) {
    pendingHighlightIndex = rowIndex;
    return;
  }

  isHighlightingInProgress = true;

  try {
    await Excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const tables = sheet.tables;
      tables.load("items/name");
      await context.sync();

      let targetTable = tables.items.find((t: any) => t.name === NOMBRE_TABLA_EXCEL) || tables.items[0];
      if (targetTable) {
        const bodyRange = targetTable.getDataBodyRange();
        bodyRange.load(["rowCount"]);
        await context.sync();

        if (rowIndex >= 0 && rowIndex < bodyRange.rowCount) {
          // Limpiar resaltado de la fila previa
          if (lastHighlightedRowIndex >= 0 && lastHighlightedRowIndex < bodyRange.rowCount && lastHighlightedRowIndex !== rowIndex) {
            const prevRow = bodyRange.getRow(lastHighlightedRowIndex);
            prevRow.format.fill.clear();
          }

          // Resaltar la fila actual en amarillo claro (#fef08a)
          const currentRow = bodyRange.getRow(rowIndex);
          currentRow.format.fill.color = "#fef08a";
          lastHighlightedRowIndex = rowIndex;

          await context.sync();
        }
      }
    });
  } catch (error) {
    // Silencioso para evitar trabar la fluidez de la animación
  } finally {
    isHighlightingInProgress = false;
    if (pendingHighlightIndex !== null && pendingHighlightIndex !== rowIndex) {
      const nextIdx = pendingHighlightIndex;
      pendingHighlightIndex = null;
      resaltarFilaEnExcel(nextIdx);
    }
  }
}

/**
 * Actualiza la posición del vehículo y la telemetría para un paso dado
 */
function actualizarPaso(index: number, panCamera: boolean = false): void {
  if (registros.length === 0) return;

  currentStepIndex = Math.max(0, Math.min(index, registros.length - 1));
  const registroActual = registros[currentStepIndex];

  // 1. Actualizar posición del marcador del vehículo
  const pos: [number, number] = [registroActual.lat, registroActual.lng];
  if (vehicleMarker) {
    vehicleMarker.setLatLng(pos);
  }

  if (panCamera && map) {
    map.panTo(pos, { animate: true, duration: 0.3 });
  }

  // 2. Actualizar polilíneas (Recorrido completado vs pendiente)
  const completedCoords = registros.slice(0, currentStepIndex + 1).map(r => [r.lat, r.lng]);
  const remainingCoords = registros.slice(currentStepIndex).map(r => [r.lat, r.lng]);

  if (completedPolyline) completedPolyline.setLatLngs(completedCoords);
  if (remainingPolyline) {
    remainingPolyline.setLatLngs(showFullTrail ? remainingCoords : []);
  }

  // 3. Calcular distancia acumulada en km
  let distAcumuladaKm = 0;
  for (let i = 0; i < currentStepIndex; i++) {
    const p1 = registros[i];
    const p2 = registros[i + 1];
    distAcumuladaKm += map.distance([p1.lat, p1.lng], [p2.lat, p2.lng]) / 1000;
  }

  // 4. Actualizar Panel de Telemetría
  const telSpeed = document.getElementById("tel-speed");
  const telDistance = document.getElementById("tel-distance");
  const telFuelText = document.getElementById("tel-fuel-text");
  const telFuelBar = document.getElementById("tel-fuel-bar");
  const telLocation = document.getElementById("tel-location");
  const telDriver = document.getElementById("tel-driver");
  const telStatus = document.getElementById("telemetry-status");

  if (telSpeed) telSpeed.textContent = registroActual.velocidad.toString();
  if (telDistance) telDistance.textContent = distAcumuladaKm.toFixed(1);
  if (telFuelText) telFuelText.textContent = `${registroActual.combustible}%`;
  if (telFuelBar) telFuelBar.style.width = `${registroActual.combustible}%`;
  if (telLocation) telLocation.textContent = registroActual.ubicacion;
  if (telDriver) telDriver.textContent = registroActual.conductor;

  if (telStatus) {
    telStatus.textContent = registroActual.estado;
    telStatus.className = "status-pill";
    if (registroActual.estado === "En Movimiento") telStatus.classList.add("status-moving");
    else if (registroActual.estado === "En Entrega") telStatus.classList.add("status-delivering");
    else telStatus.classList.add("status-stopped");
  }

  // 5. Actualizar Barra de Reproducción
  const slider = document.getElementById("timeline-slider") as HTMLInputElement;
  const timeText = document.getElementById("current-time-text");
  const stepCounter = document.getElementById("step-counter");

  if (slider && parseInt(slider.value) !== currentStepIndex) {
    slider.value = currentStepIndex.toString();
  }
  if (timeText) timeText.textContent = registroActual.hora;
  if (stepCounter) stepCounter.textContent = `Paso ${currentStepIndex + 1} de ${registros.length}`;

  // 6. Resaltar la fila correspondiente en Excel
  resaltarFilaEnExcel(currentStepIndex);

  // Notificar al simulador de Excel (GitHub Pages) si está en iframe
  if (typeof window !== "undefined" && window.parent && window.parent !== window) {
    window.parent.postMessage({ type: "APP2_GPS_STEP", step: currentStepIndex }, "*");
  }
}

// ============================================================================
// MOTOR DE REPRODUCCIÓN AUTOMÁTICA
// ============================================================================

/**
 * Alterna entre reproducir y pausar la simulación
 */
function alternarReproduccion(): void {
  if (isPlaying) {
    pausarReproduccion();
  } else {
    iniciarReproduccion();
  }
}

/**
 * Inicia la animación continua paso a paso
 */
function iniciarReproduccion(): void {
  if (currentStepIndex >= registros.length - 1) {
    currentStepIndex = 0;
  }

  isPlaying = true;
  actualizarBotonPlay();

  const intervalMs = Math.round(1100 / playbackSpeed);

  if (playbackTimer) clearInterval(playbackTimer);

  playbackTimer = setInterval(() => {
    if (currentStepIndex < registros.length - 1) {
      actualizarPaso(currentStepIndex + 1, true);
    } else {
      pausarReproduccion();
    }
  }, intervalMs);
}

/**
 * Pausa la animación continua
 */
function pausarReproduccion(): void {
  isPlaying = false;
  actualizarBotonPlay();
  if (playbackTimer) {
    clearInterval(playbackTimer);
    playbackTimer = null;
  }
}

/**
 * Actualiza el icono del botón de reproducción
 */
function actualizarBotonPlay(): void {
  const icon = document.getElementById("play-icon");
  if (!icon) return;

  if (isPlaying) {
    icon.className = "fa-solid fa-pause";
  } else {
    icon.className = "fa-solid fa-play";
  }
}

/**
 * Reinicia la simulación al paso 0
 */
function reiniciarRuta(): void {
  pausarReproduccion();
  actualizarPaso(0, true);
}

/**
 * Avanza un paso manual
 */
function avanzarPaso(): void {
  pausarReproduccion();
  if (currentStepIndex < registros.length - 1) {
    actualizarPaso(currentStepIndex + 1, true);
  }
}

// ============================================================================
// LECTURA Y CREACIÓN AUTOMÁTICA EN EXCEL (Office.js)
// ============================================================================

/**
 * Busca automáticamente la tabla 'Registro_GPS'. Si no existe, la crea con el dataset inicial.
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
        const headers = [["Paso", "Hora", "Vehiculo_ID", "Conductor", "Latitud", "Longitud", "Velocidad_kmh", "Combustible_pct", "Estado", "Ubicacion"]];
        const dataRows = DATASET_INICIAL.map(r => [
          r.paso,
          r.hora,
          r.vehiculoId,
          r.conductor,
          r.lat,
          r.lng,
          r.velocidad,
          r.combustible,
          r.estado,
          r.ubicacion
        ]);

        const totalRows = dataRows.length + 1;
        const totalCols = headers[0].length;

        const startCell = sheet.getCell(0, 0); // A1
        const targetRange = startCell.getResizedRange(totalRows - 1, totalCols - 1);
        targetRange.values = [...headers, ...dataRows];

        targetTable = sheet.tables.add(targetRange, true);
        targetTable.name = NOMBRE_TABLA_EXCEL;
        targetTable.style = "TableStyleLight9";

        targetRange.format.autofitColumns();
        await context.sync();

        registros = [...DATASET_INICIAL];
        lastHighlightedRowIndex = -1;
        actualizarSliderMax();
        renderizarPuntosDeParada();
        actualizarPaso(0, true);
        mostrarToast(`Tabla '${NOMBRE_TABLA_EXCEL}' creada y cargada en Excel`);
      } else {
        await leerTablaGPS(targetTable, context);
      }
    });
  } catch (error) {
    console.error("Error al inicializar tabla en Excel:", error);
  }
}

/**
 * Lee los datos de la tabla 'Registro_GPS' desde Excel
 */
async function leerTablaGPS(table: any, context: any): Promise<void> {
  const range = table.getDataBodyRange().load(["values", "text"]);
  const headers = table.getHeaderRowRange().load("values");
  await context.sync();

  const headerRow = headers.values[0].map((h: string) => h.toString().toLowerCase().trim());
  const rows = range.values;
  const texts = range.text || [];

  const pasoIdx = headerRow.findIndex((h: string) => h.includes("paso") || h === "id");
  const horaIdx = headerRow.findIndex((h: string) => h.includes("hora") || h.includes("time"));
  const vehiculoIdx = headerRow.findIndex((h: string) => h.includes("vehiculo") || h.includes("unidad"));
  const conductorIdx = headerRow.findIndex((h: string) => h.includes("conductor") || h.includes("chofer"));
  const latIdx = headerRow.findIndex((h: string) => h.includes("lat"));
  const lngIdx = headerRow.findIndex((h: string) => h.includes("long") || h.includes("lng"));
  const speedIdx = headerRow.findIndex((h: string) => h.includes("velocidad") || h.includes("kmh"));
  const fuelIdx = headerRow.findIndex((h: string) => h.includes("combustible") || h.includes("fuel") || h.includes("gas"));
  const estadoIdx = headerRow.findIndex((h: string) => h.includes("estado") || h.includes("status"));
  const ubicacionIdx = headerRow.findIndex((h: string) => h.includes("ubicacion") || h.includes("direccion") || h.includes("poi"));

  if (latIdx !== -1 && lngIdx !== -1) {
    const nuevosRegistros: RegistroGPS[] = [];
    rows.forEach((row: any[], i: number) => {
      const lat = parseFloat(row[latIdx]);
      const lng = parseFloat(row[lngIdx]);
      if (!isNaN(lat) && !isNaN(lng)) {
        const rawHora = horaIdx !== -1 ? row[horaIdx] : `08:${String(i * 5).padStart(2, "0")}`;
        const textHora = (texts[i] && horaIdx !== -1) ? texts[i][horaIdx] : undefined;
        const horaFormateada = formatearHoraExcel(rawHora, textHora);

        nuevosRegistros.push({
          paso: pasoIdx !== -1 ? Number(row[pasoIdx]) || (i + 1) : (i + 1),
          hora: horaFormateada,
          vehiculoId: vehiculoIdx !== -1 ? String(row[vehiculoIdx]) : "UNIDAD-01",
          conductor: conductorIdx !== -1 ? String(row[conductorIdx]) : "Conductor",
          lat: lat,
          lng: lng,
          velocidad: speedIdx !== -1 ? Number(row[speedIdx]) || 0 : 0,
          combustible: fuelIdx !== -1 ? Number(row[fuelIdx]) || 80 : 80,
          estado: estadoIdx !== -1 ? String(row[estadoIdx]) : "En Movimiento",
          ubicacion: ubicacionIdx !== -1 ? String(row[ubicacionIdx]) : `Punto ${i + 1}`
        });
      }
    });

    if (nuevosRegistros.length > 0) {
      registros = nuevosRegistros;
      lastHighlightedRowIndex = -1;
      actualizarSliderMax();
      renderizarPuntosDeParada();
      actualizarPaso(0, true);
      mostrarToast(`Sincronizados ${nuevosRegistros.length} puntos GPS desde Excel`);
    }
  }
}

/**
 * Sincroniza manualmente desde Excel
 */
async function sincronizarDesdeExcel(): Promise<void> {
  if (typeof Excel !== "undefined" && Excel.run) {
    try {
      await Excel.run(async (context: any) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const tables = sheet.tables;
        tables.load("items/name");
        await context.sync();

        let targetTable = tables.items.find((t: any) => t.name === NOMBRE_TABLA_EXCEL) || tables.items[0];
        if (targetTable) {
          await leerTablaGPS(targetTable, context);
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
    actualizarPaso(0, true);
  }
}

/**
 * Actualiza los rangos y etiquetas del slider según la cantidad de registros
 */
function actualizarSliderMax(): void {
  const slider = document.getElementById("timeline-slider") as HTMLInputElement;
  const startTime = document.getElementById("start-time-label");
  const endTime = document.getElementById("end-time-label");

  if (slider) {
    slider.max = (registros.length - 1).toString();
    slider.value = "0";
  }
  if (startTime && registros.length > 0) {
    startTime.textContent = registros[0].hora;
  }
  if (endTime && registros.length > 0) {
    endTime.textContent = registros[registros.length - 1].hora;
  }
}

// ============================================================================
// CONFIGURACIÓN DE EVENTOS DE INTERFAZ
// ============================================================================
function configurarEventosUI(): void {
  // Botones de cabecera
  document.getElementById("btn-sync")?.addEventListener("click", sincronizarDesdeExcel);

  document.getElementById("btn-center")?.addEventListener("click", () => {
    if (registros.length > 0 && map) {
      const actual = registros[currentStepIndex];
      map.setView([actual.lat, actual.lng], 15, { animate: true });
    }
  });

  document.getElementById("btn-toggle-trail")?.addEventListener("click", (e: any) => {
    showFullTrail = !showFullTrail;
    const btn = e.currentTarget as HTMLElement;
    btn.classList.toggle("active", showFullTrail);
    actualizarPaso(currentStepIndex);
  });

  // Controles de reproducción
  document.getElementById("btn-play")?.addEventListener("click", alternarReproduccion);
  document.getElementById("btn-reset")?.addEventListener("click", reiniciarRuta);
  document.getElementById("btn-step-fwd")?.addEventListener("click", avanzarPaso);

  // Slider de línea de tiempo
  const slider = document.getElementById("timeline-slider") as HTMLInputElement;
  slider?.addEventListener("input", (e: any) => {
    pausarReproduccion();
    const val = parseInt(e.target.value, 10);
    actualizarPaso(val, true);
  });

  // Selector de velocidad
  const speedChips = document.querySelectorAll(".speed-chip");
  speedChips.forEach(chip => {
    chip.addEventListener("click", (e: any) => {
      speedChips.forEach(c => c.classList.remove("active"));
      const target = e.currentTarget as HTMLElement;
      target.classList.add("active");
      playbackSpeed = parseFloat(target.getAttribute("data-speed") || "1");

      if (isPlaying) {
        iniciarReproduccion(); // Reiniciar intervalo con la nueva velocidad
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
