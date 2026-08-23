/**
 * ============================================================================
 * MAPA DE PRECIOS INMOBILIARIOS (EJEMPLO 3)
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

// Interface para el modelo de datos inmobiliarios
interface Propiedad {
  id: string;
  titulo: string;
  tipo: string;
  barrio: string;
  precio: number;
  superficie: number;
  habitaciones: number;
  banos: number;
  precioM2: number;
  lat: number;
  lng: number;
  calificacion: string;
}

// Nombres de tablas en Excel
const NOMBRE_TABLA_EXCEL = "Propiedades_Inmobiliarias";
const NOMBRE_TABLA_RESUMEN = "Propiedades_Filtradas";

// Dataset inicial escalonado estratégicamente para demostración dinámica
const DATASET_INICIAL: Propiedad[] = [
  { id: "PROP-001", titulo: "Piso Luminoso con Terraza", tipo: "Piso", barrio: "Salamanca", precio: 680000, superficie: 120, habitaciones: 3, banos: 2, precioM2: 5667, lat: 40.4290, lng: -3.6820, calificacion: "B" },
  { id: "PROP-002", titulo: "Ático Dúplex Vistas Panorámicas", tipo: "Ático", barrio: "Chamberí", precio: 890000, superficie: 165, habitaciones: 4, banos: 3, precioM2: 5394, lat: 40.4380, lng: -3.7020, calificacion: "A" },
  { id: "PROP-003", titulo: "Apartamento Reformado Centro", tipo: "Piso", barrio: "Centro / Sol", precio: 295000, superficie: 68, habitaciones: 2, banos: 1, precioM2: 4338, lat: 40.4170, lng: -3.7040, calificacion: "C" },
  { id: "PROP-004", titulo: "Chalet Adosado con Jardín", tipo: "Chalet", barrio: "Hortaleza", precio: 740000, superficie: 350, habitaciones: 5, banos: 3, precioM2: 2114, lat: 40.4680, lng: -3.6550, calificacion: "B" },
  { id: "PROP-005", titulo: "Estudio Moderno Inversión", tipo: "Estudio", barrio: "Malasaña", precio: 215000, superficie: 45, habitaciones: 1, banos: 1, precioM2: 4778, lat: 40.4260, lng: -3.7060, calificacion: "D" },
  { id: "PROP-006", titulo: "Piso Familiar Exterior", tipo: "Piso", barrio: "Retiro", precio: 540000, superficie: 110, habitaciones: 3, banos: 2, precioM2: 4909, lat: 40.4110, lng: -3.6780, calificacion: "B" },
  { id: "PROP-007", titulo: "Piso Acogedor Tetuán", tipo: "Piso", barrio: "Tetuán", precio: 245000, superficie: 75, habitaciones: 2, banos: 1, precioM2: 3267, lat: 40.4560, lng: -3.7010, calificacion: "C" },
  { id: "PROP-008", titulo: "Ático Exclusivo con Solárium", tipo: "Ático", barrio: "Chamartín", precio: 790000, superficie: 140, habitaciones: 3, banos: 2, precioM2: 5643, lat: 40.4620, lng: -3.6790, calificacion: "A" },
  { id: "PROP-009", titulo: "Chalet Independiente Piscina", tipo: "Chalet", barrio: "Arturo Soria", precio: 920000, superficie: 480, habitaciones: 6, banos: 4, precioM2: 1916, lat: 40.4510, lng: -3.6480, calificacion: "A" },
  { id: "PROP-010", titulo: "Piso Reformado Moncloa", tipo: "Piso", barrio: "Moncloa", precio: 380000, superficie: 85, habitaciones: 2, banos: 2, precioM2: 4471, lat: 40.4350, lng: -3.7220, calificacion: "C" },
  { id: "PROP-011", titulo: "Bajo con Patio Privado", tipo: "Piso", barrio: "Arganzuela", precio: 310000, superficie: 75, habitaciones: 2, banos: 1, precioM2: 4133, lat: 40.3980, lng: -3.6960, calificacion: "C" },
  { id: "PROP-012", titulo: "Dúplex Vanguardista", tipo: "Dúplex", barrio: "Salamanca", precio: 850000, superficie: 150, habitaciones: 3, banos: 3, precioM2: 5667, lat: 40.4310, lng: -3.6760, calificacion: "A" },
  { id: "PROP-013", titulo: "Piso Económico Oportunidad", tipo: "Piso", barrio: "Usera", precio: 185000, superficie: 64, habitaciones: 2, banos: 1, precioM2: 2890, lat: 40.3840, lng: -3.7070, calificacion: "E" },
  { id: "PROP-014", titulo: "Piso con Garaje y Trastero", tipo: "Piso", barrio: "Ciudad Lineal", precio: 365000, superficie: 140, habitaciones: 4, banos: 2, precioM2: 2607, lat: 40.4390, lng: -3.6420, calificacion: "C" },
  { id: "PROP-015", titulo: "Chalet Pareado Familiar", tipo: "Chalet", barrio: "San Blas", precio: 580000, superficie: 250, habitaciones: 5, banos: 3, precioM2: 2320, lat: 40.4320, lng: -3.6210, calificacion: "B" },
  { id: "PROP-016", titulo: "Ático Terraza 40m2", tipo: "Ático", barrio: "Carabanchel", precio: 280000, superficie: 78, habitaciones: 2, banos: 1, precioM2: 3590, lat: 40.3910, lng: -3.7250, calificacion: "D" },
  { id: "PROP-017", titulo: "Piso Alto con Ascensor", tipo: "Piso", barrio: "Puente de Vallecas", precio: 198000, superficie: 68, habitaciones: 2, banos: 1, precioM2: 2912, lat: 40.3960, lng: -3.6680, calificacion: "D" },
  { id: "PROP-018", titulo: "Piso Señorial Castellana", tipo: "Piso", barrio: "Chamartín", precio: 720000, superficie: 135, habitaciones: 3, banos: 2, precioM2: 5333, lat: 40.4490, lng: -3.6890, calificacion: "B" },
  { id: "PROP-019", titulo: "Dúplex con Terraza Ático", tipo: "Dúplex", barrio: "Chamberí", precio: 620000, superficie: 115, habitaciones: 3, banos: 2, precioM2: 5391, lat: 40.4410, lng: -3.7080, calificacion: "B" },
  { id: "PROP-020", titulo: "Chalet Urbano Exclusivo", tipo: "Chalet", barrio: "Mirasierra", precio: 880000, superficie: 240, habitaciones: 5, banos: 4, precioM2: 3667, lat: 40.4890, lng: -3.7150, calificacion: "A" }
];

// Estado Global de la Aplicación
let propiedades: Propiedad[] = [...DATASET_INICIAL];
let map: any = null;
let markersLayerGroup: any = null;
let propertyMarkerMap = new Map<string, any>();

// Estado de Filtros
let currentMaxBudget: number = 600000;
let currentTypeFilter: string = "all";

// Estado de Resaltado en Excel
let isSyncingExcelHighlight: boolean = false;
let pendingFilterHighlight: { idsFiltrados: Set<string>; idMejorOportunidad: string | null } | null = null;

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
// INICIALIZACIÓN PRINCIPAL
// ============================================================================
async function iniciarAplicacion(): Promise<void> {
  try {
    await asegurarDependencias();
    inicializarMapa();
    renderizarMarcadores();
    actualizarAnalitica();
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

  map = L.map("map", {
    center: [40.435, -3.688],
    zoom: 12,
    zoomControl: false
  });

  L.control.zoom({ position: "topleft" }).addTo(map);

  L.tileLayer("https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png", {
    attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors &copy; <a href="https://carto.com/attributions">CARTO</a>',
    subdomains: "abcd",
    maxZoom: 19
  }).addTo(map);

  markersLayerGroup = L.layerGroup().addTo(map);
}

/**
 * Renderiza los marcadores tipo Pin de Precio en el mapa
 */
function renderizarMarcadores(): void {
  if (!markersLayerGroup) return;

  markersLayerGroup.clearLayers();
  propertyMarkerMap.clear();

  propiedades.forEach(prop => {
    const isWithinBudget = prop.precio <= currentMaxBudget;
    const isMatchingType = currentTypeFilter === "all" || prop.tipo.toLowerCase() === currentTypeFilter.toLowerCase();
    const isActive = isWithinBudget && isMatchingType;

    const formattedPrice = `€ ${Math.round(prop.precio / 1000)}k`;

    const customIcon = L.divIcon({
      className: "custom-price-pin-wrapper",
      html: `
        <div id="pin-${prop.id}" class="price-tag-pin ${isActive ? "within-budget" : "over-budget"}">
          <span>${formattedPrice}</span>
        </div>
      `,
      iconSize: [60, 26],
      iconAnchor: [30, 13],
      popupAnchor: [0, -15]
    });

    const marker = L.marker([prop.lat, prop.lng], { icon: customIcon });

    const popupContent = `
      <div class="property-popup-card">
        <h4 class="popup-title">${prop.titulo}</h4>
        <div class="popup-badge-row">
          <span class="popup-type">${prop.tipo} en ${prop.barrio}</span>
          <span class="popup-energy">Eficiencia ${prop.calificacion}</span>
        </div>
        <div class="popup-price-row">
          <span class="popup-price">€ ${prop.precio.toLocaleString()}</span>
          <span class="popup-sqm-price">€ ${prop.precioM2.toLocaleString()} /m²</span>
        </div>
        <div class="popup-specs-row">
          <span class="popup-spec-item"><i class="fa-solid fa-ruler-combined"></i> ${prop.superficie} m²</span>
          <span class="popup-spec-item"><i class="fa-solid fa-bed"></i> ${prop.habitaciones} hab</span>
          <span class="popup-spec-item"><i class="fa-solid fa-bath"></i> ${prop.banos} baños</span>
        </div>
      </div>
    `;

    marker.bindPopup(popupContent);
    markersLayerGroup.addLayer(marker);
    propertyMarkerMap.set(prop.id, marker);
  });
}

/**
 * Actualiza el estado visual de los marcadores existentes sin reconstruir toda la capa
 */
function actualizarEstadoMarcadores(): void {
  propiedades.forEach(prop => {
    const isWithinBudget = prop.precio <= currentMaxBudget;
    const isMatchingType = currentTypeFilter === "all" || prop.tipo.toLowerCase() === currentTypeFilter.toLowerCase();
    const isActive = isWithinBudget && isMatchingType;

    const el = document.getElementById(`pin-${prop.id}`);
    if (el) {
      if (isActive) {
        el.className = "price-tag-pin within-budget";
      } else {
        el.className = "price-tag-pin over-budget";
      }
    }
  });
}

/**
 * Obtiene las propiedades que cumplen con los filtros actuales
 */
function obtenerPropiedadesFiltradas(): Propiedad[] {
  return propiedades.filter(prop => {
    const matchesBudget = prop.precio <= currentMaxBudget;
    const matchesType = currentTypeFilter === "all" || prop.tipo.toLowerCase() === currentTypeFilter.toLowerCase();
    return matchesBudget && matchesType;
  });
}

/**
 * Actualiza el panel flotante de analítica y la mejor oportunidad
 */
function actualizarAnalitica(): void {
  const filtradas = obtenerPropiedadesFiltradas();

  const countBadge = document.getElementById("filtered-count-badge");
  const statAvgPrice = document.getElementById("stat-avg-price");
  const statAvgSqm = document.getElementById("stat-avg-sqm");
  const oppTitle = document.getElementById("opp-title");
  const oppPrice = document.getElementById("opp-price");
  const oppSqmPrice = document.getElementById("opp-sqm-price");

  if (countBadge) {
    countBadge.textContent = `${filtradas.length} de ${propiedades.length}`;
  }

  if (filtradas.length === 0) {
    if (statAvgPrice) statAvgPrice.textContent = "€ 0";
    if (statAvgSqm) statAvgSqm.textContent = "€ 0 /m²";
    if (oppTitle) oppTitle.textContent = "Sin coincidencias";
    if (oppPrice) oppPrice.textContent = "-";
    if (oppSqmPrice) oppSqmPrice.textContent = "-";

    sincronizarResaltadoFiltradoEnExcel(filtradas, null);
    return;
  }

  const sumaPrecios = filtradas.reduce((acc, p) => acc + p.precio, 0);
  const sumaSqm = filtradas.reduce((acc, p) => acc + p.precioM2, 0);

  const precioMedio = Math.round(sumaPrecios / filtradas.length);
  const sqmMedio = Math.round(sumaSqm / filtradas.length);

  if (statAvgPrice) statAvgPrice.textContent = `€ ${precioMedio.toLocaleString()}`;
  if (statAvgSqm) statAvgSqm.textContent = `€ ${sqmMedio.toLocaleString()} /m²`;

  // Mejor oportunidad: Propiedad con el menor precio por m² dentro del rango
  const mejorOportunidad = [...filtradas].sort((a, b) => a.precioM2 - b.precioM2)[0];

  if (mejorOportunidad) {
    if (oppTitle) oppTitle.textContent = mejorOportunidad.titulo;
    if (oppPrice) oppPrice.textContent = `€ ${mejorOportunidad.precio.toLocaleString()}`;
    if (oppSqmPrice) oppSqmPrice.textContent = `€ ${mejorOportunidad.precioM2.toLocaleString()} /m²`;
  }

  // Sincronizar resaltado dinámico en Excel (Verde para la mejor opción, azul suave para las demás en presupuesto)
  sincronizarResaltadoFiltradoEnExcel(filtradas, mejorOportunidad ? mejorOportunidad.id : null);

  // Notificar al simulador de Excel (GitHub Pages) si está en iframe
  if (typeof window !== "undefined" && window.parent && window.parent !== window) {
    window.parent.postMessage({ type: "APP3_BUDGET_CHANGE", budget: currentMaxBudget }, "*");
  }
}

// ============================================================================
// RESALTADO DINÁMICO EN EXCEL (Verde para la mejor opción, Azul para el resto)
// ============================================================================
async function sincronizarResaltadoFiltradoEnExcel(filtradas: Propiedad[], idMejorOpcion: string | null): Promise<void> {
  if (typeof Excel === "undefined" || !Excel.run) return;

  const idsFiltrados = new Set(filtradas.map(p => p.id));

  if (isSyncingExcelHighlight) {
    pendingFilterHighlight = { idsFiltrados, idMejorOportunidad: idMejorOpcion };
    return;
  }

  isSyncingExcelHighlight = true;

  try {
    await Excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const tables = sheet.tables;
      tables.load("items/name");
      await context.sync();

      let table = tables.items.find((t: any) => t.name === NOMBRE_TABLA_EXCEL) || tables.items[0];
      if (table) {
        const bodyRange = table.getDataBodyRange();
        const headers = table.getHeaderRowRange().load("values");
        bodyRange.load(["values", "rowCount"]);
        await context.sync();

        const headerRow = headers.values[0].map((h: any) => h.toString().toLowerCase().trim());
        const idIdx = headerRow.findIndex((h: string) => h === "id" || h.includes("id"));

        const rows = bodyRange.values;
        for (let i = 0; i < rows.length; i++) {
          const rowId = idIdx !== -1 ? String(rows[i][idIdx]) : (propiedades[i] ? propiedades[i].id : "");
          const rowRange = bodyRange.getRow(i);

          if (rowId === idMejorOpcion) {
            // 🌟 LA MEJOR OPCIÓN (€/m² más bajo dentro del rango): Verde suave (#86efac)
            rowRange.format.fill.color = "#86efac";
          } else if (idsFiltrados.has(rowId)) {
            // 🏠 EN PRESUPUESTO: Azul sutil (#dbeafe)
            rowRange.format.fill.color = "#dbeafe";
          } else {
            // ❌ FUERA DE PRESUPUESTO: Limpiar formato de fondo
            rowRange.format.fill.clear();
          }
        }
        await context.sync();
      }
    });
  } catch (err) {
    console.error("Error al resaltar en Excel:", err);
  } finally {
    isSyncingExcelHighlight = false;
    if (pendingFilterHighlight) {
      const { idsFiltrados: nextIds, idMejorOportunidad: nextBest } = pendingFilterHighlight;
      pendingFilterHighlight = null;
      sincronizarResaltadoFiltradoEnExcel(propiedades.filter(p => nextIds.has(p.id)), nextBest);
    }
  }
}

// ============================================================================
// LECTURA Y CREACIÓN AUTOMÁTICA EN EXCEL (Office.js)
// ============================================================================

/**
 * Busca automáticamente la tabla 'Propiedades_Inmobiliarias'. Si no existe, la crea con el dataset inicial.
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
        const headers = [["ID", "Titulo", "Tipo", "Barrio", "Precio_EUR", "Superficie_m2", "Habitaciones", "Banos", "Precio_m2", "Latitud", "Longitud", "Calificacion"]];
        const dataRows = DATASET_INICIAL.map(p => [
          p.id,
          p.titulo,
          p.tipo,
          p.barrio,
          p.precio,
          p.superficie,
          p.habitaciones,
          p.banos,
          p.precioM2,
          p.lat,
          p.lng,
          p.calificacion
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

        propiedades = [...DATASET_INICIAL];
        renderizarMarcadores();
        actualizarAnalitica();
        mostrarToast(`Tabla '${NOMBRE_TABLA_EXCEL}' creada y cargada en Excel`);
      } else {
        await leerTablaInmobiliaria(targetTable, context);
      }
    });
  } catch (error) {
    console.error("Error al inicializar tabla en Excel:", error);
  }
}

/**
 * Lee los datos de la tabla inmobiliaria desde Excel
 */
async function leerTablaInmobiliaria(table: any, context: any): Promise<void> {
  const range = table.getDataBodyRange().load("values");
  const headers = table.getHeaderRowRange().load("values");
  await context.sync();

  const headerRow = headers.values[0].map((h: string) => h.toString().toLowerCase().trim());
  const rows = range.values;

  const idIdx = headerRow.findIndex((h: string) => h === "id");
  const titleIdx = headerRow.findIndex((h: string) => h.includes("titulo") || h.includes("propiedad") || h.includes("nombre"));
  const typeIdx = headerRow.findIndex((h: string) => h.includes("tipo"));
  const neighborhoodIdx = headerRow.findIndex((h: string) => h.includes("barrio") || h.includes("zona") || h.includes("ciudad"));
  const priceIdx = headerRow.findIndex((h: string) => h.includes("precio") || h.includes("eur") || h.includes("valor"));
  const sqIdx = headerRow.findIndex((h: string) => h.includes("superficie") || h.includes("m2") || h.includes("metros"));
  const roomsIdx = headerRow.findIndex((h: string) => h.includes("hab"));
  const bathsIdx = headerRow.findIndex((h: string) => h.includes("baño") || h.includes("bano"));
  const sqPriceIdx = headerRow.findIndex((h: string) => h.includes("precio_m2") || h.includes("m2_precio"));
  const latIdx = headerRow.findIndex((h: string) => h.includes("lat"));
  const lngIdx = headerRow.findIndex((h: string) => h.includes("long") || h.includes("lng"));
  const energyIdx = headerRow.findIndex((h: string) => h.includes("calific") || h.includes("energ"));

  if (latIdx !== -1 && lngIdx !== -1) {
    const nuevasPropiedades: Propiedad[] = [];
    rows.forEach((row: any[], i: number) => {
      const lat = parseFloat(row[latIdx]);
      const lng = parseFloat(row[lngIdx]);
      const precio = priceIdx !== -1 ? Number(row[priceIdx]) || 300000 : 300000;
      const superficie = sqIdx !== -1 ? Number(row[sqIdx]) || 80 : 80;
      const precioM2 = sqPriceIdx !== -1 && Number(row[sqPriceIdx]) ? Number(row[sqPriceIdx]) : Math.round(precio / superficie);

      if (!isNaN(lat) && !isNaN(lng)) {
        nuevasPropiedades.push({
          id: idIdx !== -1 ? String(row[idIdx]) : `PROP-${String(i + 1).padStart(3, "0")}`,
          titulo: titleIdx !== -1 ? String(row[titleIdx]) : `Inmueble ${i + 1}`,
          tipo: typeIdx !== -1 ? String(row[typeIdx]) : "Piso",
          barrio: neighborhoodIdx !== -1 ? String(row[neighborhoodIdx]) : "Madrid",
          precio: precio,
          superficie: superficie,
          habitaciones: roomsIdx !== -1 ? Number(row[roomsIdx]) || 2 : 2,
          banos: bathsIdx !== -1 ? Number(row[bathsIdx]) || 1 : 1,
          precioM2: precioM2,
          lat: lat,
          lng: lng,
          calificacion: energyIdx !== -1 ? String(row[energyIdx]) : "C"
        });
      }
    });

    if (nuevasPropiedades.length > 0) {
      propiedades = nuevasPropiedades;
      renderizarMarcadores();
      actualizarAnalitica();
      mostrarToast(`Sincronizadas ${nuevasPropiedades.length} propiedades desde Excel`);
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
          await leerTablaInmobiliaria(targetTable, context);
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
    actualizarAnalitica();
  }
}

// ============================================================================
// EXPORTACIÓN DE RESUMEN EJECUTIVO Y TABLA FILTRADA EN EXCEL
// ============================================================================

/**
 * Exporta tanto las métricas clave como la tabla de propiedades filtradas
 * en la zona situada exactamente 2 columnas a la derecha de la tabla original.
 */
async function descargarResumenExcel(): Promise<void> {
  const filtradas = obtenerPropiedadesFiltradas();

  if (filtradas.length === 0) {
    mostrarToast("No hay propiedades dentro del presupuesto seleccionado", true);
    return;
  }

  const sumaPrecios = filtradas.reduce((acc, p) => acc + p.precio, 0);
  const sumaSqm = filtradas.reduce((acc, p) => acc + p.precioM2, 0);
  const precioMedio = Math.round(sumaPrecios / filtradas.length);
  const sqmMedio = Math.round(sumaSqm / filtradas.length);
  const mejorOportunidad = [...filtradas].sort((a, b) => a.precioM2 - b.precioM2)[0];

  // Notificar al simulador de Excel (GitHub Pages) si está en iframe
  if (typeof window !== "undefined" && window.parent && window.parent !== window) {
    window.parent.postMessage({
      type: "APP3_EXPORT_SUMMARY",
      kpis: {
        presupuesto: currentMaxBudget,
        conteo: filtradas.length,
        precioMedio: precioMedio,
        mejorOpcion: mejorOportunidad ? mejorOportunidad.titulo : "N/A"
      }
    }, "*");
  }

  // 1. Si se ejecuta dentro de Microsoft Excel (Office.js)
  if (typeof Excel !== "undefined" && Excel.run) {
    try {
      await Excel.run(async (context: any) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const tables = sheet.tables;
        tables.load("items/name");
        await context.sync();

        let tablePrincipal = tables.items.find((t: any) => t.name === NOMBRE_TABLA_EXCEL) || tables.items[0];

        let startCol = 14; // Columna O por defecto
        let startRow = 0;  // Fila 1 por defecto

        if (tablePrincipal) {
          const rangePrincipal = tablePrincipal.getRange().load(["columnIndex", "columnCount", "rowIndex", "rowCount"]);
          await context.sync();
          startCol = rangePrincipal.columnIndex + rangePrincipal.columnCount + 2;
          startRow = rangePrincipal.rowIndex;
        }

        // Si ya existe la tabla 'Propiedades_Filtradas', eliminarla para regenerar limpiamente
        let existingSummaryTable = tables.items.find((t: any) => t.name === NOMBRE_TABLA_RESUMEN);
        if (existingSummaryTable) {
          existingSummaryTable.delete();
          await context.sync();
        }

        // -------------------------------------------------------------
        // A. Escribir el Bloque de Métricas Clave (KPIs Ejecutivos)
        // -------------------------------------------------------------
        const kpiHeaders = [
          ["RESUMEN EJECUTIVO DE MERCADO", ""],
          ["Presupuesto Máximo Filtrado:", `€ ${currentMaxBudget.toLocaleString()}`],
          ["Propiedades Seleccionadas:", `${filtradas.length} de ${propiedades.length}`],
          ["Precio Medio Inmuebles:", `€ ${precioMedio.toLocaleString()}`],
          ["Precio Medio / m²:", `€ ${sqmMedio.toLocaleString()} / m²`],
          ["Mejor Oportunidad (€/m²):", `${mejorOportunidad.titulo} (€ ${mejorOportunidad.precioM2.toLocaleString()}/m²)`]
        ];

        const kpiStartCell = sheet.getCell(startRow, startCol);
        const kpiRange = kpiStartCell.getResizedRange(kpiHeaders.length - 1, 1);
        kpiRange.values = kpiHeaders;

        // Estilo visual del bloque KPI
        kpiRange.format.fill.color = "#f8fafc";
        const headerKpiCell = kpiStartCell.getResizedRange(0, 1);
        headerKpiCell.format.font.bold = true;
        headerKpiCell.format.font.color = "#1e40af";
        headerKpiCell.format.fill.color = "#dbeafe";

        // -------------------------------------------------------------
        // B. Escribir y Crear la Tabla Oficial 'Propiedades_Filtradas'
        // -------------------------------------------------------------
        const tableHeaders = [
          "ID",
          "Titulo",
          "Tipo",
          "Barrio",
          "Precio_EUR",
          "Superficie_m2",
          "Habitaciones",
          "Banos",
          "Precio_m2",
          "Calificacion"
        ];

        const tableDataRows = filtradas.map(p => [
          p.id,
          p.titulo,
          p.tipo,
          p.barrio,
          p.precio,
          p.superficie,
          p.habitaciones,
          p.banos,
          p.precioM2,
          p.calificacion
        ]);

        const tableStartRow = startRow + kpiHeaders.length + 2; // 2 filas de separación
        const tableStartCell = sheet.getCell(tableStartRow, startCol);
        const targetTableRange = tableStartCell.getResizedRange(tableDataRows.length, tableHeaders.length - 1);
        targetTableRange.values = [tableHeaders, ...tableDataRows];

        const newSummaryTable = sheet.tables.add(targetTableRange, true /* hasHeaders */);
        newSummaryTable.name = NOMBRE_TABLA_RESUMEN;
        newSummaryTable.style = "TableStyleLight10"; // Estilo limpio azul sutil

        // Autoajustar columnas del reporte
        const fullReportRange = kpiStartCell.getResizedRange(tableDataRows.length + kpiHeaders.length + 3, tableHeaders.length - 1);
        fullReportRange.format.autofitColumns();
        await context.sync();

        mostrarToast(`¡Resumen y tabla '${NOMBRE_TABLA_RESUMEN}' guardados en Excel!`);
      });
      return;
    } catch (error: any) {
      console.error("Error al exportar resumen a Excel:", error);
      mostrarToast(`Error en Excel: ${error?.message || error}`, true);
    }
  } else {
    // 2. Fallback para navegador: Descarga de archivo de reporte
    try {
      const kpiText = [
        "=== RESUMEN EJECUTIVO DE MERCADO ===",
        `Presupuesto Máximo: € ${currentMaxBudget.toLocaleString()}`,
        `Propiedades Seleccionadas: ${filtradas.length} de ${propiedades.length}`,
        `Precio Medio: € ${precioMedio.toLocaleString()}`,
        `Precio Medio / m2: € ${sqmMedio.toLocaleString()}`,
        `Mejor Oportunidad: ${mejorOportunidad.titulo} (€ ${mejorOportunidad.precioM2.toLocaleString()}/m2)`,
        "",
        "=== PROPIEDADES FILTRADAS ===",
        "ID,Titulo,Tipo,Barrio,Precio_EUR,Superficie_m2,Habitaciones,Banos,Precio_m2,Calificacion",
        ...filtradas.map(p => `"${p.id}","${p.titulo}","${p.tipo}","${p.barrio}",${p.precio},${p.superficie},${p.habitaciones},${p.banos},${p.precioM2},"${p.calificacion}"`)
      ].join("\n");

      const blob = new Blob([kpiText], { type: "text/csv;charset=utf-8;" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.setAttribute("href", url);
      link.setAttribute("download", `${NOMBRE_TABLA_RESUMEN}.csv`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      mostrarToast(`Resumen descargado como '${NOMBRE_TABLA_RESUMEN}.csv'`);
    } catch (err) {
      console.error("Error al descargar archivo:", err);
      mostrarToast("Error al exportar archivo", true);
    }
  }
}

// ============================================================================
// CONFIGURACIÓN DE EVENTOS DE INTERFAZ
// ============================================================================
function configurarEventosUI(): void {
  // Botones de cabecera
  document.getElementById("btn-sync")?.addEventListener("click", sincronizarDesdeExcel);
  document.getElementById("btn-export-summary")?.addEventListener("click", descargarResumenExcel);

  // Slider de Presupuesto
  const slider = document.getElementById("budget-slider") as HTMLInputElement;
  const budgetDisplay = document.getElementById("budget-display");

  slider?.addEventListener("input", (e: any) => {
    const val = parseInt(e.target.value, 10);
    currentMaxBudget = val;

    if (budgetDisplay) {
      budgetDisplay.textContent = `€ ${val.toLocaleString()}`;
    }

    // Actualizar botones de preset rápido
    const budgetChips = document.querySelectorAll(".budget-chip");
    budgetChips.forEach(chip => {
      const chipVal = parseInt(chip.getAttribute("data-budget") || "0", 10);
      chip.classList.toggle("active", chipVal === val);
    });

    actualizarEstadoMarcadores();
    actualizarAnalitica();
  });

  // Chips de Presupuesto Rápido
  const budgetChips = document.querySelectorAll(".budget-chip");
  budgetChips.forEach(chip => {
    chip.addEventListener("click", (e: any) => {
      budgetChips.forEach(c => c.classList.remove("active"));
      const target = e.currentTarget as HTMLElement;
      target.classList.add("active");

      const budgetVal = parseInt(target.getAttribute("data-budget") || "600000", 10);
      currentMaxBudget = budgetVal;

      if (slider) slider.value = budgetVal.toString();
      if (budgetDisplay) budgetDisplay.textContent = `€ ${budgetVal.toLocaleString()}`;

      actualizarEstadoMarcadores();
      actualizarAnalitica();
    });
  });

  // Chips de Filtro por Tipo
  const typeChips = document.querySelectorAll(".type-chip");
  typeChips.forEach(chip => {
    chip.addEventListener("click", (e: any) => {
      typeChips.forEach(c => c.classList.remove("active"));
      const target = e.currentTarget as HTMLElement;
      target.classList.add("active");

      currentTypeFilter = target.getAttribute("data-type") || "all";

      actualizarEstadoMarcadores();
      actualizarAnalitica();
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
