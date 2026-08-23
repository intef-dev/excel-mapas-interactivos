// ============================================================================
// SIMULADOR DE EXCEL WEB & COMUNICACIÓN CON APPS DE SCRIPT LAB
// ============================================================================

// 1. DATASETS DE DEMOSTRACIÓN PARA LAS 3 HOJAS DE EXCEL

const DATASET_HOJA_1 = [
  ["CLI-001", "Supermercado La Sierra", "Roberto Méndez", "Madrid", 40.4532, -3.6883, "Norte", 12500, "Alta"],
  ["CLI-002", "Farmacias San Miguel", "Lucía Gómez", "Madrid", 40.4611, -3.6912, "Norte", 8900, "Media"],
  ["CLI-003", "Distribuidora Del Valle", "Carlos Ortiz", "Madrid", 40.4485, -3.6720, "Norte", 15400, "Alta"],
  ["CLI-004", "Restaurante El Mirador", "Elena Ramos", "Madrid", 40.4705, -3.7021, "Norte", 6200, "Baja"],
  ["CLI-005", "Tecnología & Redes Global", "Marcos Díaz", "Madrid", 40.4182, -3.7035, "Centro", 21000, "Alta"],
  ["CLI-006", "Boutique Castellana", "Sofía Morales", "Madrid", 40.4250, -3.6910, "Centro", 9800, "Media"],
  ["CLI-007", "Cafetería Plaza Mayor", "Javier Santos", "Madrid", 40.4155, -3.7074, "Centro", 5400, "Baja"],
  ["CLI-008", "Librería Central Sol", "Ana Beltrán", "Madrid", 40.4168, -3.7038, "Centro", 7300, "Media"],
  ["CLI-009", "Automotriz del Sur", "Pedro Castillo", "Madrid", 40.3882, -3.7125, "Sur", 18200, "Alta"],
  ["CLI-010", "Almacenes Del Río", "Isabel Núñez", "Madrid", 40.3754, -3.6980, "Sur", 11300, "Media"],
  ["CLI-011", "Centro Médico Usera", "David Romero", "Madrid", 40.3820, -3.7050, "Sur", 14100, "Alta"],
  ["CLI-012", "Panificadora Imperial", "Carmen Vega", "Madrid", 40.3690, -3.7180, "Sur", 4900, "Baja"],
  ["CLI-013", "Muebles & Diseño Este", "Fernando Torres", "Madrid", 40.4280, -3.6520, "Este", 13700, "Media"],
  ["CLI-014", "Logística San Blas", "Raquel Pardo", "Madrid", 40.4350, -3.6300, "Este", 22500, "Alta"],
  ["CLI-015", "Construcciones Vicálvaro", "Hugo Lozano", "Madrid", 40.4050, -3.6100, "Este", 16800, "Alta"],
  ["CLI-016", "Clínica Dental Poniente", "Patricia Gil", "Madrid", 40.4390, -3.7350, "Oeste", 9200, "Media"],
  ["CLI-017", "Óptica Moncloa", "Andrés Serrano", "Madrid", 40.4355, -3.7195, "Oeste", 7800, "Baja"],
  ["CLI-018", "Electrónica Puerta de Hierro", "Marta Flores", "Madrid", 40.4620, -3.7410, "Oeste", 19400, "Alta"]
];

const DATASET_HOJA_2 = [
  ["GPS-001", "UNIDAD-01", "08:00", 40.4532, -3.6883, 0, 98, "Inicio de Ruta - Centro Logístico Norte"],
  ["GPS-002", "UNIDAD-01", "08:10", 40.4580, -3.6850, 42, 96, "Av. de Burgos - Tráfico fluido"],
  ["GPS-003", "UNIDAD-01", "08:20", 40.4650, -3.6780, 58, 93, "Incorporación a M-30 Norte"],
  ["GPS-004", "UNIDAD-01", "08:30", 40.4720, -3.6690, 72, 90, "M-30 Km 3 - Velocidad crucero"],
  ["GPS-005", "UNIDAD-01", "08:40", 40.4620, -3.6550, 65, 87, "M-30 Este - Tramo Arturo Soria"],
  ["GPS-006", "UNIDAD-01", "08:50", 40.4480, -3.6480, 48, 84, "Salida hacia San Blas / Canillejas"],
  ["GPS-007", "UNIDAD-01", "09:00", 40.4380, -3.6420, 35, 82, "Calle Alcalá - Primera parada de reparto"],
  ["GPS-008", "UNIDAD-01", "09:10", 40.4380, -3.6420, 0, 82, "Parada: Descarga en Supermercado Este"],
  ["GPS-009", "UNIDAD-01", "09:20", 40.4310, -3.6500, 38, 79, "Avenida de Daroca"],
  ["GPS-010", "UNIDAD-01", "09:30", 40.4200, -3.6620, 52, 76, "Acceso a M-30 Sur"],
  ["GPS-011", "UNIDAD-01", "09:40", 40.4050, -3.6750, 68, 73, "M-30 Sur - Puente de Vallecas"],
  ["GPS-012", "UNIDAD-01", "09:50", 40.3920, -3.6890, 64, 70, "M-30 Sur - Nudo Sur / Méndez Álvaro"],
  ["GPS-013", "UNIDAD-01", "10:00", 40.3810, -3.7020, 30, 67, "Salida a Usera / Marcelo Usera"],
  ["GPS-014", "UNIDAD-01", "10:10", 40.3810, -3.7020, 0, 67, "Parada: Centro Médico Usera"],
  ["GPS-015", "UNIDAD-01", "10:25", 40.3890, -3.7150, 40, 64, "Calle Antonio López hacia Marqués de Vadillo"],
  ["GPS-016", "UNIDAD-01", "10:35", 40.4050, -3.7180, 28, 61, "Paseo de Santa María de la Cabeza"],
  ["GPS-017", "UNIDAD-01", "10:45", 40.4155, -3.7074, 18, 59, "Zona Centro - Acceso a Plaza Mayor"],
  ["GPS-018", "UNIDAD-01", "10:55", 40.4155, -3.7074, 0, 59, "Parada: Cafetería Plaza Mayor"],
  ["GPS-019", "UNIDAD-01", "11:05", 40.4250, -3.7038, 25, 57, "Calle Gran Vía hacia Plaza de España"],
  ["GPS-020", "UNIDAD-01", "11:15", 40.4350, -3.6950, 35, 55, "Llegada al Centro Logístico Castellana"]
];

const DATASET_HOJA_3 = [
  ["PROP-001", "Piso Reformado Chamberí", "Piso", 295000, 90, 3278, "Calle Fuencarral 102", 40.4310, -3.7020],
  ["PROP-002", "Ático con Terraza Salamanca", "Ático", 480000, 110, 4364, "Calle Serrano 45", 40.4285, -3.6870],
  ["PROP-003", "Piso Luminoso Retiro", "Piso", 340000, 85, 4000, "Avenida Menéndez Pelayo 20", 40.4180, -3.6790],
  ["PROP-004", "Chalet Adosado Hortaleza", "Chalet", 740000, 350, 2114, "Calle Gran Vía de Hortaleza 12", 40.4680, -3.6550],
  ["PROP-005", "Dúplex Moderno Malasaña", "Dúplex", 385000, 95, 4053, "Calle San Vicente Ferrer 14", 40.4260, -3.7045],
  ["PROP-006", "Piso Acogedor Arganzuela", "Piso", 245000, 72, 3403, "Paseo de las Delicias 30", 40.4010, -3.6940],
  ["PROP-007", "Ático Dúplex Moncloa", "Ático", 520000, 130, 4000, "Calle Princesa 60", 40.4320, -3.7170],
  ["PROP-008", "Piso Familiar Chamartín", "Piso", 410000, 105, 3905, "Calle Príncipe de Vergara 210", 40.4490, -3.6780],
  ["PROP-009", "Chalet Arturo Soria con Piscina", "Chalet", 920000, 480, 1916, "Calle Arturo Soria 180", 40.4550, -3.6520],
  ["PROP-010", "Dúplex de Diseño Centro", "Dúplex", 450000, 115, 3913, "Calle Mayor 35", 40.4160, -3.7090],
  ["PROP-011", "Piso Exterior Tetuán", "Piso", 215000, 68, 3162, "Calle Bravo Murillo 140", 40.4540, -3.7025],
  ["PROP-012", "Ático Vistas Palacio Real", "Ático", 610000, 140, 4357, "Calle Bailén 18", 40.4140, -3.7130],
  ["PROP-013", "Piso Económico Usera", "Piso", 185000, 64, 2890, "Calle Marcelo Usera 50", 40.3830, -3.7060],
  ["PROP-014", "Piso con Garaje Ciudad Lineal", "Piso", 365000, 140, 2607, "Calle Alcalá 320", 40.4340, -3.6450],
  ["PROP-015", "Chalet Pareado San Blas", "Chalet", 580000, 250, 2320, "Calle Pobladura del Valle 15", 40.4250, -3.6200],
  ["PROP-016", "Dúplex Terraza Pacífico", "Dúplex", 310000, 88, 3523, "Calle Doctor Esquerdo 150", 40.4060, -3.6710],
  ["PROP-017", "Piso Reformado Vallecas", "Piso", 195000, 70, 2786, "Avenida de la Albufera 40", 40.3950, -3.6650],
  ["PROP-018", "Piso Amplio La Latina", "Piso", 275000, 80, 3438, "Calle Toledo 70", 40.4110, -3.7090],
  ["PROP-019", "Chalet Exclusivo Pozuelo", "Chalet", 890000, 420, 2119, "Avenida de Europa 25", 40.4420, -3.7850],
  ["PROP-020", "Piso Inversión Lavapiés", "Piso", 230000, 65, 3538, "Calle Argumosa 12", 40.4085, -3.6985]
];

// Estado de la Simulación
let hojaActiva = 1;
let activeGpsStep = 0;
let currentBudget = 500000;
let generatedRouteRows = null;
let summaryKpiBlock = null;

// ============================================================================
// INICIALIZACIÓN DE LA UI
// ============================================================================
document.addEventListener("DOMContentLoaded", () => {
  renderizarHojaExcel(1);
  configurarEventosSimulador();
  iniciarComunicacionIframe();
});

function configurarEventosSimulador() {
  // Pestañas de Hojas
  document.querySelectorAll(".sheet-tab").forEach(tab => {
    tab.addEventListener("click", () => {
      document.querySelectorAll(".sheet-tab").forEach(t => t.classList.remove("active"));
      tab.classList.add("active");
      const hoja = parseInt(tab.getAttribute("data-sheet"));
      hojaActiva = hoja;
      renderizarHojaExcel(hoja);
      sincronizarAppConHoja(hoja);
    });
  });

  // Selector de Apps en el Taskpane
  document.querySelectorAll(".app-switch-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".app-switch-btn").forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      const appId = btn.getAttribute("data-app");
      cambiarAppTaskpane(appId);
    });
  });

  // Botón Script Lab Run en el Ribbon
  const btnRun = document.getElementById("ribbon-btn-run");
  if (btnRun) {
    btnRun.addEventListener("click", () => {
      const iframe = document.getElementById("taskpane-iframe");
      if (iframe) iframe.src = iframe.src;
    });
  }
}

function cambiarAppTaskpane(appId) {
  const iframe = document.getElementById("taskpane-iframe");
  if (!iframe) return;

  if (appId === "app1") {
    iframe.src = "runner1.html";
    seleccionarHoja(1);
  } else if (appId === "app2") {
    iframe.src = "runner2.html";
    seleccionarHoja(2);
  } else if (appId === "app3") {
    iframe.src = "runner3.html";
    seleccionarHoja(3);
  }
}

function seleccionarHoja(num) {
  hojaActiva = num;
  document.querySelectorAll(".sheet-tab").forEach(t => {
    t.classList.toggle("active", parseInt(t.getAttribute("data-sheet")) === num);
  });
  renderizarHojaExcel(num);
}

function sincronizarAppConHoja(hojaNum) {
  document.querySelectorAll(".app-switch-btn").forEach(b => {
    b.classList.toggle("active", b.getAttribute("data-app") === `app${hojaNum}`);
  });
  const iframe = document.getElementById("taskpane-iframe");
  if (iframe) {
    iframe.src = `runner${hojaNum}.html`;
  }
}

// ============================================================================
// RENDERIZADO DE LAS TABLAS EN EL GRID DE EXCEL
// ============================================================================
function renderizarHojaExcel(numHoja) {
  const gridContainer = document.getElementById("grid-table-body");
  if (!gridContainer) return;

  let html = "";
  const cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R"];

  // 1. Fila de Cabeceras de Columna (A, B, C...)
  html += `<tr><th class="col-header corner-header"></th>`;
  cols.forEach(c => {
    html += `<th class="col-header">${c}</th>`;
  });
  html += `</tr>`;

  if (numHoja === 1) {
    // HOJA 1: Datos_Clientes + (Ruta_Generada a la derecha)
    html += `<tr><th class="row-header">1</th>`;
    const headers1 = ["ID", "Cliente", "Contacto", "Ciudad", "Latitud", "Longitud", "Zona", "Venta_USD", "Prioridad"];
    headers1.forEach(h => html += `<th class="th-table-header">${h}</th>`);
    html += `<td></td><td></td>`; // Col J y K vacías (separación de 2 columnas)
    
    if (generatedRouteRows) {
      const headersRuta = ["Orden", "ID_Cliente", "Cliente", "Ciudad", "Zona", "Dist_km", "Tiempo_min"];
      headersRuta.forEach(h => html += `<th class="th-table-green">${h}</th>`);
    }
    html += `</tr>`;

    DATASET_HOJA_1.forEach((row, idx) => {
      const rowNum = idx + 2;
      html += `<tr><th class="row-header">${rowNum}</th>`;
      row.forEach(val => {
        html += `<td>${val}</td>`;
      });
      html += `<td></td><td></td>`; // Espaciador 2 cols

      if (generatedRouteRows && generatedRouteRows[idx]) {
        const r = generatedRouteRows[idx];
        html += `<td style="background:#f0fdf4; font-weight:600;">${r[0]}</td>`;
        html += `<td>${r[1]}</td><td>${r[2]}</td><td>${r[3]}</td><td>${r[4]}</td><td>${r[7]}</td><td>${r[9]}</td>`;
      }
      html += `</tr>`;
    });
  } 
  else if (numHoja === 2) {
    // HOJA 2: Registro_GPS (Con resaltado reactivo en amarillo)
    html += `<tr><th class="row-header">1</th>`;
    const headers2 = ["ID_Registro", "Vehiculo", "Hora", "Latitud", "Longitud", "Velocidad_kmh", "Combustible_pct", "Estado_Evento"];
    headers2.forEach(h => html += `<th class="th-table-header">${h}</th>`);
    html += `</tr>`;

    DATASET_HOJA_2.forEach((row, idx) => {
      const rowNum = idx + 2;
      const isGpsActive = (idx === activeGpsStep);
      const cellClass = isGpsActive ? "cell-highlight-gps" : "";

      html += `<tr><th class="row-header">${rowNum}</th>`;
      row.forEach((val, cIdx) => {
        let displayVal = val;
        if (cIdx === 5) displayVal = `${val} km/h`;
        else if (cIdx === 6) displayVal = `${val}%`;
        html += `<td class="${cellClass}">${displayVal}</td>`;
      });
      html += `</tr>`;
    });
  }
  else if (numHoja === 3) {
    // HOJA 3: Propiedades_Inmobiliarias + Slicer Resaltado Bicolor + Resumen
    html += `<tr><th class="row-header">1</th>`;
    const headers3 = ["ID", "Titulo", "Tipo", "Precio_EUR", "Superficie_m2", "Coste_m2", "Direccion", "Latitud", "Longitud"];
    headers3.forEach(h => html += `<th class="th-table-header">${h}</th>`);
    html += `<td></td><td></td>`; // Col J y K separadoras limpias

    if (summaryKpiBlock) {
      html += `<th class="th-table-green" colspan="4">RESUMEN EJECUTIVO INMOBILIARIO</th>`;
    }
    html += `</tr>`;

    // Encontrar la mejor oportunidad en el presupuesto actual
    let bestDealId = "";
    let minCost = Infinity;
    DATASET_HOJA_3.forEach(p => {
      const precio = p[3];
      const costeM2 = p[5];
      if (precio <= currentBudget && costeM2 < minCost) {
        minCost = costeM2;
        bestDealId = p[0];
      }
    });

    DATASET_HOJA_3.forEach((row, idx) => {
      const rowNum = idx + 2;
      const id = row[0];
      const precio = row[3];
      const inBudget = precio <= currentBudget;
      const isBest = (id === bestDealId);

      let cellClass = "";
      if (isBest) {
        cellClass = "cell-best-deal";
      } else if (inBudget) {
        cellClass = "cell-in-budget";
      } else {
        cellClass = "cell-over-budget";
      }

      html += `<tr><th class="row-header">${rowNum}</th>`;
      row.forEach((val, cIdx) => {
        let displayVal = val;
        if (cIdx === 3) displayVal = `€ ${val.toLocaleString()}`;
        else if (cIdx === 4) displayVal = `${val} m²`;
        else if (cIdx === 5) displayVal = `${val.toLocaleString()} €/m²`;
        
        html += `<td class="${cellClass}">${displayVal}</td>`;
      });

      html += `<td></td><td></td>`; // Columnas J y K limpias (sin fondos ni colores)

      // Renderizar Bloque KPI a partir de Columna L
      if (summaryKpiBlock) {
        if (idx === 0) {
          html += `<td colspan="2" class="kpi-excel-block"><strong>Presupuesto Máx:</strong></td><td colspan="2" class="kpi-excel-block">€ ${summaryKpiBlock.presupuesto.toLocaleString()}</td>`;
        } else if (idx === 1) {
          html += `<td colspan="2" class="kpi-excel-block"><strong>Total Propiedades:</strong></td><td colspan="2" class="kpi-excel-block">${summaryKpiBlock.conteo} disponibles</td>`;
        } else if (idx === 2) {
          html += `<td colspan="2" class="kpi-excel-block"><strong>Precio Medio:</strong></td><td colspan="2" class="kpi-excel-block">€ ${summaryKpiBlock.precioMedio.toLocaleString()}</td>`;
        } else if (idx === 3) {
          html += `<td colspan="2" class="kpi-excel-block"><strong>Mejor Oportunidad:</strong></td><td colspan="2" class="kpi-excel-block" style="background:#86efac; font-weight:700;">${summaryKpiBlock.mejorOpcion}</td>`;
        }
      }

      html += `</tr>`;
    });
  }

  gridContainer.innerHTML = html;
}

// ============================================================================
// SIMULACIÓN DE LA API DE EXCEL (Office.js) PARA LOS IFRAMES DE SCRIPT LAB
// ============================================================================
function iniciarComunicacionIframe() {
  window.addEventListener("message", (event) => {
    const data = event.data;
    if (!data || !data.type) return;

    if (data.type === "APP2_GPS_STEP") {
      activeGpsStep = data.step;
      if (hojaActiva === 2) {
        renderizarHojaExcel(2);
      }
    } else if (data.type === "APP3_BUDGET_CHANGE") {
      currentBudget = data.budget;
      if (hojaActiva === 3) {
        renderizarHojaExcel(3);
      }
    } else if (data.type === "APP1_EXPORT_ROUTE") {
      generatedRouteRows = data.route;
      if (hojaActiva === 1) {
        renderizarHojaExcel(1);
      }
    } else if (data.type === "APP3_EXPORT_SUMMARY") {
      summaryKpiBlock = data.kpis;
      if (hojaActiva === 3) {
        renderizarHojaExcel(3);
      }
    }
  });
}
