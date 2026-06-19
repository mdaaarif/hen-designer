/* ==========================================================================
   CORE JAVASCRIPT FOR PINCHHEN DESIGNER
   ========================================================================== */

// --- Global Application State ---
const state = {
  deltaTmin: 10,
  streams: [],
  matches: [],
  utilities: [],
  selectedMatchId: null,
  selectedUtilityId: null,
  draggedMatchId: null,
  
  // Interaction states
  interactionMode: 'normal', // 'normal', 'add-exchanger-step1', 'add-exchanger-step2', 'add-utility'
  pendingHotStreamId: null,
  pendingColdStreamId: null,

  // Calculation outputs
  calculatedTargets: {
    QHmin: 0,
    QCmin: 0,
    pinchShifted: 0,
    pinchHot: 0,
    pinchCold: 0,
    tempList: [],
    Rcas: [],
    nMin: 0
  },
  
  // Sim simulation outputs
  simulation: {
    streamTemps: {}, // streamId -> array of temperatures at slot boundaries [0...8]
    actualQH: 0,
    actualQC: 0,
    diagnostics: [],
    streamSatisfaction: {} // streamId -> % satisfied
  }
};

// --- Preloaded Examples Data ---
const EXAMPLES = {
  example4: {
    deltaTmin: 10,
    streams: [
      { id: 'H1', name: 'H1 (Hot 1)', type: 'hot', Tin: 150, Tout: 60, MCp: 20 },
      { id: 'H2', name: 'H2 (Hot 2)', type: 'hot', Tin: 150, Tout: 30, MCp: 80 },
      { id: 'C1', name: 'C1 (Cold 1)', type: 'cold', Tin: 20, Tout: 135, MCp: 80 },
      { id: 'C2', name: 'C2 (Cold 2)', type: 'cold', Tin: 80, Tout: 140, MCp: 40 }
    ],
    matches: [],
    utilities: []
  },
  example5: {
    deltaTmin: 15,
    streams: [
      { id: 'H1', name: 'H1 (Feed Preheat)', type: 'hot', Tin: 250, Tout: 120, MCp: 10 },
      { id: 'H2', name: 'H2 (Kero Reflux)', type: 'hot', Tin: 200, Tout: 80, MCp: 20 },
      { id: 'H3', name: 'H3 (Diesel Product)', type: 'hot', Tin: 180, Tout: 50, MCp: 5 },
      { id: 'C1', name: 'C1 (Crude Stream A)', type: 'cold', Tin: 60, Tout: 180, MCp: 15 },
      { id: 'C2', name: 'C2 (Crude Stream B)', type: 'cold', Tin: 90, Tout: 220, MCp: 25 }
    ],
    matches: [],
    utilities: []
  }
};

// --- Page Initialization ---
document.addEventListener('DOMContentLoaded', () => {
  loadExample('example4');
  setupEventListeners();
  initExcelUpload();
});

// --- Setup Event Listeners ---
function setupEventListeners() {
  // Example Select
  document.getElementById('example-select').addEventListener('change', (e) => {
    if (e.target.value !== 'custom') {
      loadExample(e.target.value);
    }
  });

  // Slider
  const slider = document.getElementById('tmin-slider');
  const valueDisplay = document.getElementById('tmin-value');
  slider.addEventListener('input', (e) => {
    state.deltaTmin = parseInt(e.target.value);
    valueDisplay.textContent = `${state.deltaTmin} °C`;
    runPinchAnalysis();
    simulateNetwork();
    renderAll();
  });

  // Upload Header Button
  document.getElementById('upload-header-btn').addEventListener('click', () => {
    document.getElementById('upload-modal').classList.remove('hidden');
    document.getElementById('file-name-display').textContent = '';
  });

  // Reset Button
  document.getElementById('reset-btn').addEventListener('click', () => {
    loadExample('example4');
  });

  // Add Stream Button
  document.getElementById('add-stream-btn').addEventListener('click', () => {
    addNewStreamRow();
  });

  // Tab buttons
  const tabButtons = document.querySelectorAll('.tab-nav .tab-btn');
  tabButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      tabButtons.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      
      const tabId = btn.getAttribute('data-tab');
      document.querySelectorAll('.tab-pane').forEach(pane => {
        pane.classList.remove('active');
      });
      document.getElementById(tabId).classList.add('active');
      
      // Force SVG re-render to handle sizing/dimensions correctly
      renderAll();
    });
  });

  // Exchanger / Match controls
  document.getElementById('add-exchanger-btn').addEventListener('click', () => {
    enterInteractionMode('add-exchanger-step1');
  });

  document.getElementById('add-utility-btn').addEventListener('click', () => {
    enterInteractionMode('add-utility');
  });

  document.getElementById('clear-design-btn').addEventListener('click', () => {
    state.matches = [];
    state.utilities = [];
    state.selectedMatchId = null;
    state.selectedUtilityId = null;
    simulateNetwork();
    renderAll();
  });

  document.getElementById('auto-design-btn').addEventListener('click', () => {
    autoDesignNetwork();
  });

  // Editor Cancel/Delete/Save
  document.getElementById('cancel-edit-btn').addEventListener('click', () => {
    closeEditor();
  });

  document.getElementById('delete-match-btn').addEventListener('click', () => {
    deleteSelectedDevice();
  });

  document.getElementById('save-edit-btn').addEventListener('click', () => {
    saveExchangerEdit();
  });

  // Dragging event handlers for Grid SVG
  const svg = document.getElementById('hen-svg');
  svg.addEventListener('mousemove', handleSvgMouseMove);
  svg.addEventListener('mouseup', handleSvgMouseUp);
  svg.addEventListener('mouseleave', handleSvgMouseUp);
}

// --- Load Selected Example ---
function loadExample(key) {
  const ex = EXAMPLES[key];
  if (!ex) return;

  state.deltaTmin = ex.deltaTmin;
  state.streams = JSON.parse(JSON.stringify(ex.streams));
  state.matches = JSON.parse(JSON.stringify(ex.matches));
  state.utilities = JSON.parse(JSON.stringify(ex.utilities));
  state.selectedMatchId = null;
  state.selectedUtilityId = null;

  // Sync UI controls
  document.getElementById('example-select').value = key;
  document.getElementById('tmin-slider').value = state.deltaTmin;
  document.getElementById('tmin-value').textContent = `${state.deltaTmin} °C`;

  runPinchAnalysis();
  simulateNetwork();
  renderAll();
}

// ==========================================================================
// EXCEL / SPREADSHEET LOADER (CLIENT-SIDE SHEETJS INTEGRATION)
// ==========================================================================
function initExcelUpload() {
  const modal = document.getElementById('upload-modal');
  const skipBtn = document.getElementById('skip-modal-btn');
  const dropZone = document.getElementById('drop-zone');
  const fileInput = document.getElementById('file-input');

  // Skip Modal
  skipBtn.addEventListener('click', () => {
    modal.classList.add('hidden');
  });

  // Click drop zone triggers file selector
  dropZone.addEventListener('click', () => {
    fileInput.click();
  });

  // File selector change
  fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) handleExcelUpload(file);
  });

  // Drag over / drag leave effects
  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
  });

  dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('dragover');
  });

  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file) handleExcelUpload(file);
  });
}

function handleExcelUpload(file) {
  if (!file) return;

  const display = document.getElementById('file-name-display');
  display.textContent = `Processing: ${file.name}...`;
  display.style.color = 'var(--text-highlight)';

  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      // 1. Parse Settings (Tmin)
      let tmin = 10; // default approach
      const settingsSheetName = workbook.SheetNames.find(name => name.toLowerCase() === 'settings');
      if (settingsSheetName) {
        const sheet = workbook.Sheets[settingsSheetName];
        const rows = XLSX.utils.sheet_to_json(sheet);
        // Find row where Parameter (first column) is 'tmin'
        const tminRow = rows.find(r => {
          const keys = Object.keys(r);
          if (keys.length === 0) return false;
          const keyVal = String(r[keys[0]]).trim().toLowerCase();
          return keyVal === 'tmin';
        });
        if (tminRow) {
          const keys = Object.keys(tminRow);
          if (keys.length > 1) tmin = parseFloat(tminRow[keys[1]]) || 10;
        }
      }

      // 2. Parse Streams
      const streamsSheetName = workbook.SheetNames.find(name => name.toLowerCase() === 'streams');
      if (!streamsSheetName) {
        throw new Error("Could not find a sheet named 'Streams' in the Excel workbook.");
      }

      const sheet = workbook.Sheets[streamsSheetName];
      const rows = XLSX.utils.sheet_to_json(sheet);
      if (rows.length === 0) {
        throw new Error("The 'Streams' sheet is empty.");
      }

      const parsedStreams = [];
      rows.forEach((r, idx) => {
        const keys = Object.keys(r);
        const getVal = (possibleHeaders) => {
          const matchKey = keys.find(k => possibleHeaders.includes(k.trim().toLowerCase()));
          return matchKey ? r[matchKey] : null;
        };

        const id = getVal(['stream', 'name', 'id', 'stream name']);
        const type = String(getVal(['type']) || '').trim().toLowerCase();
        const tin = parseFloat(getVal(['tin', 'tin (°c)', 't_in', 'supply']));
        const tout = parseFloat(getVal(['tout', 'tout (°c)', 't_out', 'target']));
        const mcp = parseFloat(getVal(['mcp', 'mcp (mw/°c)', 'fc_p', 'fcp', 'mcp']));

        if (!id) {
          throw new Error(`Row ${idx + 2}: Missing Stream ID.`);
        }
        if (type !== 'hot' && type !== 'cold') {
          throw new Error(`Row ${idx + 2} (${id}): Type must be 'hot' or 'cold' (found: '${type}')`);
        }
        if (isNaN(tin) || isNaN(tout) || isNaN(mcp)) {
          throw new Error(`Row ${idx + 2} (${id}): Temp (Tin, Tout) or MCp is invalid or missing.`);
        }

        parsedStreams.push({
          id: String(id).toUpperCase().trim(),
          name: String(id).toUpperCase().trim(),
          type: type,
          Tin: tin,
          Tout: tout,
          MCp: mcp
        });
      });

      // Update State
      state.deltaTmin = tmin;
      state.streams = parsedStreams;
      state.matches = [];
      state.utilities = [];

      // Update Controls
      document.getElementById('tmin-slider').value = tmin;
      document.getElementById('tmin-value').textContent = `${tmin} °C`;
      document.getElementById('example-select').value = 'custom';

      // Re-run
      runPinchAnalysis();
      simulateNetwork();
      renderAll();

      // Close Modal
      document.getElementById('upload-modal').classList.add('hidden');
      display.textContent = '';

    } catch (err) {
      console.error(err);
      display.textContent = `Error: ${err.message}`;
      display.style.color = 'var(--color-danger)';
    }
  };

  reader.onerror = function() {
    display.textContent = "Error reading file buffer.";
    display.style.color = 'var(--color-danger)';
  };

  reader.readAsArrayBuffer(file);
}

// --- Dynamic Stream Table ---
function populateStreamTable() {
  const tbody = document.getElementById('streams-body');
  tbody.innerHTML = '';

  state.streams.forEach(s => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><span class="sat-name ${s.type}">${s.id}</span></td>
      <td>
        <button class="stream-type-btn ${s.type}" onclick="toggleStreamType('${s.id}')">
          ${s.type}
        </button>
      </td>
      <td><input type="number" value="${s.Tin}" onchange="updateStreamField('${s.id}', 'Tin', this.value)"></td>
      <td><input type="number" value="${s.Tout}" onchange="updateStreamField('${s.id}', 'Tout', this.value)"></td>
      <td><input type="number" step="0.1" value="${s.MCp}" onchange="updateStreamField('${s.id}', 'MCp', this.value)"></td>
      <td><button class="delete-row-btn" onclick="deleteStream('${s.id}')">×</button></td>
    `;
    tbody.appendChild(tr);
  });
}

function updateStreamField(id, field, value) {
  const stream = state.streams.find(s => s.id === id);
  if (stream) {
    stream[field] = parseFloat(value);
    // Clear network elements if stream properties change drastically to avoid crashes
    state.matches = [];
    state.utilities = [];
    runPinchAnalysis();
    simulateNetwork();
    renderAll();
  }
}

function toggleStreamType(id) {
  const stream = state.streams.find(s => s.id === id);
  if (stream) {
    stream.type = stream.type === 'hot' ? 'cold' : 'hot';
    state.matches = [];
    state.utilities = [];
    runPinchAnalysis();
    simulateNetwork();
    renderAll();
  }
}

function deleteStream(id) {
  state.streams = state.streams.filter(s => s.id !== id);
  state.matches = [];
  state.utilities = [];
  runPinchAnalysis();
  simulateNetwork();
  renderAll();
}

function addNewStreamRow() {
  const id = prompt("Enter Unique Stream ID (e.g. H3, C3):");
  if (!id) return;
  if (state.streams.some(s => s.id.toLowerCase() === id.toLowerCase())) {
    alert("Stream ID already exists!");
    return;
  }

  const type = id.toUpperCase().startsWith('C') ? 'cold' : 'hot';
  state.streams.push({
    id: id.toUpperCase(),
    name: `${id.toUpperCase()} (Stream)`,
    type: type,
    Tin: type === 'hot' ? 120 : 30,
    Tout: type === 'hot' ? 40 : 110,
    MCp: 10
  });

  state.matches = [];
  state.utilities = [];
  runPinchAnalysis();
  simulateNetwork();
  renderAll();
}

// ==========================================================================
// STEP 1-7 ALGORITHMS — PORTED FROM THE USER'S PYTHON CODE
// ==========================================================================
function runPinchAnalysis() {
  if (state.streams.length === 0) return;

  const shift = state.deltaTmin / 2;
  
  // 1. Shift Temperatures
  const shiftedStreams = state.streams.map(s => {
    return {
      ...s,
      Tin_s: s.type === 'hot' ? s.Tin - shift : s.Tin + shift,
      Tout_s: s.type === 'hot' ? s.Tout - shift : s.Tout + shift
    };
  });

  // 2. Temperature Intervals
  const allTemps = new Set();
  shiftedStreams.forEach(s => {
    allTemps.add(s.Tin_s);
    allTemps.add(s.Tout_s);
  });
  const tempList = Array.from(allTemps).sort((a, b) => b - a);

  // 3. Fk (flowrate boundaries)
  const Fk = tempList.map(T => {
    let fk = 0;
    shiftedStreams.forEach(s => {
      if (Math.abs(T - s.Tin_s) < 1e-5) fk += s.MCp;
      if (Math.abs(T - s.Tout_s) < 1e-5) fk -= s.MCp;
    });
    return Number(fk.toFixed(6));
  });

  // 4. CumFk
  const CumFk = [];
  let cumulative = 0;
  Fk.forEach(fk => {
    cumulative += fk;
    CumFk.push(Number(cumulative.toFixed(6)));
  });

  // 5. Qk
  const Qk = [0];
  for (let i = 1; i < tempList.length; i++) {
    const q = CumFk[i - 1] * (tempList[i - 1] - tempList[i]);
    Qk.push(Number(q.toFixed(6)));
  }

  // 6. Qcas
  const Qcas = [];
  let cumulativeQ = 0;
  Qk.forEach(q => {
    cumulativeQ += q;
    Qcas.push(Number(cumulativeQ.toFixed(6)));
  });

  // 7. Rcas
  const minQcas = Math.min(...Qcas);
  const Rcas = Qcas.map(q => Number((q - minQcas).toFixed(6)));

  // Identify targets
  const QHmin = Rcas[0];
  const QCmin = Rcas[Rcas.length - 1];
  
  // Find pinch shifted index
  const pinchIdx = Rcas.findIndex(val => Math.abs(val) < 1e-5);
  const pinchShifted = tempList[pinchIdx] || 0;
  const pinchHot = pinchShifted + shift;
  const pinchCold = pinchShifted - shift;

  // Minimum Exchangers Calculation
  const hotStreams = state.streams.filter(s => s.type === 'hot');
  const coldStreams = state.streams.filter(s => s.type === 'cold');
  
  const aboveHot = hotStreams.filter(s => s.Tin > pinchHot);
  const aboveCold = coldStreams.filter(s => s.Tout > pinchCold);
  const belowHot = hotStreams.filter(s => s.Tout < pinchHot);
  const belowCold = coldStreams.filter(s => s.Tin < pinchCold);

  const N_above = aboveHot.length + aboveCold.length + 1; // + hot utility
  const N_below = belowHot.length + belowCold.length + 1; // + cold utility
  const nMin = (N_above - 1) + (N_below - 1);

  // Store outputs
  state.calculatedTargets = {
    QHmin,
    QCmin,
    pinchShifted,
    pinchHot,
    pinchCold,
    tempList,
    Rcas,
    nMin
  };

  // Update HUD elements
  document.getElementById('target-qh').innerHTML = `${QHmin.toFixed(1)} <span class="unit">MW</span>`;
  document.getElementById('target-qc').innerHTML = `${QCmin.toFixed(1)} <span class="unit">MW</span>`;
  document.getElementById('target-pinch-shifted').innerHTML = `${pinchShifted.toFixed(1)} <span class="unit">°C</span>`;
  document.getElementById('target-pinch-real').innerHTML = `${pinchHot.toFixed(1)} / ${pinchCold.toFixed(1)} <span class="unit">°C</span>`;
  document.getElementById('target-nmin').innerHTML = `${nMin} <span class="unit">units</span>`;
}

// ==========================================================================
// DYNAMIC HEN GRID NETWORK SIMULATOR & VALIDATOR
// ==========================================================================
function simulateNetwork() {
  const streamTemps = {};
  const hotStreams = state.streams.filter(s => s.type === 'hot');
  const coldStreams = state.streams.filter(s => s.type === 'cold');
  
  // Initialize stream boundary temperatures (we have 8 slots, so 9 boundaries: 0 to 8)
  // Indices: 0 is leftmost point (hot inlet), 8 is rightmost point (hot outlet)
  state.streams.forEach(s => {
    streamTemps[s.id] = new Array(9).fill(null);
  });

  // Calculate hot streams Left-to-Right
  hotStreams.forEach(h => {
    const temps = streamTemps[h.id];
    temps[0] = h.Tin; // Starts at supply temp

    for (let slot = 1; slot <= 8; slot++) {
      // Find devices in this slot
      const match = state.matches.find(m => m.hotStreamId === h.id && m.slot === slot);
      const cooler = state.utilities.find(u => u.streamId === h.id && u.slot === slot && u.type === 'cooler');
      
      let load = 0;
      if (match) load = match.load;
      if (cooler) load = cooler.load;

      temps[slot] = temps[slot - 1] - load / h.MCp;
    }
  });

  // Calculate cold streams Right-to-Left (since flow is opposite)
  coldStreams.forEach(c => {
    const temps = streamTemps[c.id];
    temps[8] = c.Tin; // Starts on the right at supply temp

    for (let slot = 8; slot >= 1; slot--) {
      const match = state.matches.find(m => m.coldStreamId === c.id && m.slot === slot);
      const heater = state.utilities.find(u => u.streamId === c.id && u.slot === slot && u.type === 'heater');
      
      let load = 0;
      if (match) load = match.load;
      if (heater) load = heater.load;

      temps[slot - 1] = temps[slot] + load / c.MCp;
    }
  });

  state.simulation.streamTemps = streamTemps;

  // Calculate actual utilities consumed
  let actualQH = 0;
  state.utilities.forEach(u => {
    if (u.type === 'heater') actualQH += u.load;
  });

  let actualQC = 0;
  state.utilities.forEach(u => {
    if (u.type === 'cooler') actualQC += u.load;
  });

  state.simulation.actualQH = actualQH;
  state.simulation.actualQC = actualQC;

  // Diagnostics and validation check
  runDiagnostics();
}

function runDiagnostics() {
  const diagnostics = [];
  let isFeasible = true;

  // 1. Utility Targets alignment
  const targetQH = state.calculatedTargets.QHmin;
  const targetQC = state.calculatedTargets.QCmin;
  const actualQH = state.simulation.actualQH;
  const actualQC = state.simulation.actualQC;

  if (actualQH < targetQH - 1e-2) {
    diagnostics.push({
      type: 'warning',
      text: `Hot utility is below thermodynamic target (${actualQH.toFixed(1)} / ${targetQH.toFixed(1)} MW). You need more heat.`
    });
  } else if (actualQH > targetQH + 1e-2) {
    diagnostics.push({
      type: 'warning',
      text: `Hot utility exceeds target (${actualQH.toFixed(1)} / ${targetQH.toFixed(1)} MW). Opportunity for heat recovery!`
    });
  }

  // 2. Stream target temperatures satisfaction
  const satisfaction = {};
  state.streams.forEach(s => {
    const temps = state.simulation.streamTemps[s.id];
    const finalTemp = s.type === 'hot' ? temps[8] : temps[0];
    const totalRequired = Math.abs(s.Tout - s.Tin) * s.MCp;
    
    let actualTransferred = 0;
    if (s.type === 'hot') {
      actualTransferred = (s.Tin - finalTemp) * s.MCp;
    } else {
      actualTransferred = (finalTemp - s.Tin) * s.MCp;
    }

    const percentage = Math.min(100, Math.max(0, (actualTransferred / totalRequired) * 100));
    satisfaction[s.id] = {
      percentage,
      finalTemp,
      isSatisfied: Math.abs(finalTemp - s.Tout) < 1e-2
    };

    if (!satisfaction[s.id].isSatisfied) {
      isFeasible = false;
      diagnostics.push({
        type: 'error',
        text: `Stream ${s.id} not satisfied: Outlet is ${finalTemp.toFixed(1)}°C (Target: ${s.Tout}°C).`
      });
    }
  });
  state.simulation.streamSatisfaction = satisfaction;

  // 3. Exchanger checks: Temperature Approach & Pinch violations
  state.matches.forEach(m => {
    const hTemps = state.simulation.streamTemps[m.hotStreamId];
    const cTemps = state.simulation.streamTemps[m.coldStreamId];
    
    // Temperatures at slot boundary
    // Match at slot S:
    // Hot flows left-to-right: enters at S-1, exits at S
    const Th_in = hTemps[m.slot - 1];
    const Th_out = hTemps[m.slot];
    // Cold flows right-to-left: enters at S, exits at S-1
    const Tc_in = cTemps[m.slot];
    const Tc_out = cTemps[m.slot - 1];

    // Left approach = Th_in - Tc_out
    // Right approach = Th_out - Tc_in
    const dt_left = Th_in - Tc_out;
    const dt_right = Th_out - Tc_in;

    m.hasCrossover = false;

    if (dt_left < 0 || dt_right < 0) {
      isFeasible = false;
      m.hasCrossover = true;
      diagnostics.push({
        type: 'error',
        text: `Exchanger Match ${m.id} has Temperature Crossover! (Left: ${dt_left.toFixed(1)}°C, Right: ${dt_right.toFixed(1)}°C)`
      });
    } else if (dt_left < state.deltaTmin || dt_right < state.deltaTmin) {
      m.hasCrossover = true; // Highlight
      diagnostics.push({
        type: 'warning',
        text: `Exchanger Match ${m.id} violates approach criteria ΔTmin (${Math.min(dt_left, dt_right).toFixed(1)}°C < ${state.deltaTmin}°C).`
      });
    }

    // Pinch violation checks
    // Pinch is between slot 4 and 5
    const pinchHot = state.calculatedTargets.pinchHot;
    const pinchCold = state.calculatedTargets.pinchCold;

    if (m.slot <= 4) { // Above pinch matches
      if (Th_out < pinchHot - 1e-2 || Tc_in < pinchCold - 1e-2) {
        isFeasible = false;
        diagnostics.push({
          type: 'error',
          text: `Exchanger Match ${m.id} crosses the Pinch! Transfers heat to below-pinch region.`
        });
      }
    } else { // Below pinch matches
      if (Th_in > pinchHot + 1e-2 || Tc_out > pinchCold + 1e-2) {
        isFeasible = false;
        diagnostics.push({
          type: 'error',
          text: `Exchanger Match ${m.id} crosses the Pinch! Transfers heat from above-pinch region.`
        });
      }
    }
  });

  // 4. Utility placements checks
  state.utilities.forEach(u => {
    if (u.type === 'heater' && u.slot >= 5) {
      isFeasible = false;
      diagnostics.push({
        type: 'error',
        text: `Heater ${u.id} is placed below the Pinch! (Violates Pinch heuristic: no hot utility below pinch)`
      });
    }
    if (u.type === 'cooler' && u.slot <= 4) {
      isFeasible = false;
      diagnostics.push({
        type: 'error',
        text: `Cooler ${u.id} is placed above the Pinch! (Violates Pinch heuristic: no cold utility above pinch)`
      });
    }
  });

  if (diagnostics.length === 0 && isFeasible) {
    diagnostics.push({
      type: 'success',
      text: 'Feasible and optimal heat exchanger network achieved!'
    });
  }

  state.simulation.diagnostics = diagnostics;
  
  // Update UI badge
  const badge = document.getElementById('network-status-badge');
  if (isFeasible && diagnostics.every(d => d.type !== 'warning')) {
    badge.className = 'badge badge-success';
    badge.textContent = 'Feasible';
  } else if (diagnostics.some(d => d.type === 'error')) {
    badge.className = 'badge badge-danger';
    badge.textContent = 'Violations';
  } else {
    badge.className = 'badge badge-warning';
    badge.textContent = 'Suboptimal';
  }
}

// ==========================================================================
// RENDERERS (COMPOSITE CURVES SVG, GCC SVG, HEN GRID SVG)
// ==========================================================================

function renderAll() {
  populateStreamTable();
  renderDiagnostics();
  renderSatisfaction();

  // Redraw canvases
  drawCompositeCurves();
  drawGrandCompositeCurve();
  drawHenGrid();
}

function renderDiagnostics() {
  const container = document.getElementById('diagnostic-list');
  container.innerHTML = '';

  state.simulation.diagnostics.forEach(d => {
    const item = document.createElement('div');
    item.className = `diagnostic-item ${d.type}`;
    
    // Icon selection
    let icon = '';
    if (d.type === 'success') {
      icon = `<svg class="diagnostic-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>`;
    } else if (d.type === 'warning') {
      icon = `<svg class="diagnostic-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>`;
    } else {
      icon = `<svg class="diagnostic-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>`;
    }

    item.innerHTML = `${icon} <span>${d.text}</span>`;
    container.appendChild(item);
  });

  // Update Progress bars
  const targetQH = state.calculatedTargets.QHmin;
  const targetQC = state.calculatedTargets.QCmin;
  const actualQH = state.simulation.actualQH;
  const actualQC = state.simulation.actualQC;

  document.getElementById('diag-qh-actual').textContent = actualQH.toFixed(1);
  document.getElementById('diag-qh-target').textContent = targetQH.toFixed(1);
  document.getElementById('diag-qc-actual').textContent = actualQC.toFixed(1);
  document.getElementById('diag-qc-target').textContent = targetQC.toFixed(1);

  const fillQH = document.getElementById('fill-qh');
  const fillQC = document.getElementById('fill-qc');

  fillQH.style.width = `${Math.min(100, (actualQH / (targetQH || 1)) * 100)}%`;
  fillQC.style.width = `${Math.min(100, (actualQC / (targetQC || 1)) * 100)}%`;
}

function renderSatisfaction() {
  const container = document.getElementById('satisfaction-list');
  container.innerHTML = '';

  state.streams.forEach(s => {
    const sat = state.simulation.streamSatisfaction[s.id] || { percentage: 0, finalTemp: s.Tin, isSatisfied: false };
    const item = document.createElement('div');
    item.className = 'satisfaction-item';
    
    item.innerHTML = `
      <div class="sat-header">
        <span class="sat-name ${s.type}">${s.id} (${s.type === 'hot' ? 'cooling' : 'heating'})</span>
        <span class="sat-temps">${s.Tin}°C → ${sat.finalTemp.toFixed(1)} / ${s.Tout}°C</span>
      </div>
      <div class="sat-progress-row">
        <div class="sat-progress-bar">
          <div class="sat-progress-fill ${s.type}" style="width: ${sat.percentage}%"></div>
        </div>
        <span class="sat-status ${sat.isSatisfied ? 'complete' : ''}">
          ${sat.percentage.toFixed(0)}%
        </span>
      </div>
    `;
    container.appendChild(item);
  });
}

// --- Helper: Build Composite Curves Coordinates (same logic as Python) ---
function buildComposite(streamList) {
  if (streamList.length === 0) return { T: [], H: [] };
  
  // Sort temperatures
  const temps = Array.from(new Set(
    streamList.map(s => s.Tin).concat(streamList.map(s => s.Tout))
  )).sort((a, b) => a - b); // ascending

  const coords_T = [temps[0]];
  const coords_HD = [0.0];
  let cum_HD = 0.0;

  for (let i = 0; i < temps.length - 1; i++) {
    const T_low = temps[i];
    const T_high = temps[i + 1];
    
    let mcp_total = 0;
    streamList.forEach(s => {
      const T_min_s = Math.min(s.Tin, s.Tout);
      const T_max_s = Math.max(s.Tin, s.Tout);
      if (T_min_s <= T_low && T_max_s >= T_high) {
        mcp_total += s.MCp;
      }
    });

    cum_HD += mcp_total * (T_high - T_low);
    coords_T.push(T_high);
    coords_HD.push(cum_HD);
  }

  return { T: coords_T, H: coords_HD };
}

function findXAtTemperature(coords_T, coords_HD, target_T) {
  for (let i = 0; i < coords_T.length - 1; i++) {
    const T1 = coords_T[i];
    const T2 = coords_T[i + 1];
    if (target_T >= T1 && target_T <= T2) {
      if (Math.abs(T2 - T1) < 1e-5) return coords_HD[i];
      const fraction = (target_T - T1) / (T2 - T1);
      const HD = coords_HD[i] + fraction * (coords_HD[i + 1] - coords_HD[i]);
      return HD;
    }
  }
  return 0;
}

// --- Render SVG Composite Curves ---
function drawCompositeCurves() {
  const svg = document.getElementById('composite-svg');
  svg.innerHTML = '';

  const hotStreams = state.streams.filter(s => s.type === 'hot');
  const coldStreams = state.streams.filter(s => s.type === 'cold');
  if (hotStreams.length === 0 || coldStreams.length === 0) return;

  const hotCC = buildComposite(hotStreams);
  const coldCC = buildComposite(coldStreams);

  // Shift cold curve by QC_min
  const QCmin = state.calculatedTargets.QCmin;
  const coldHShifted = coldCC.H.map(h => h + QCmin);

  // Math sizing
  const hMax = Math.max(Math.max(...hotCC.H), Math.max(...coldHShifted));
  const tMin = Math.min(Math.min(...hotCC.T), Math.min(...coldCC.T));
  const tMax = Math.max(Math.max(...hotCC.T), Math.max(...coldCC.T));
  
  // Coordinate transformations
  const padding = 50;
  const w = 800;
  const h = 500;
  
  const scaleX = (val) => padding + (val / (hMax || 1)) * (w - 2 * padding);
  const scaleY = (val) => h - padding - ((val - tMin) / ((tMax - tMin) || 1)) * (h - 2 * padding);

  // Draw Grid lines
  const gridGroup = document.createElementNS('http://www.w3.org/2000/svg', 'g');
  for (let t = Math.floor(tMin / 20) * 20; t <= tMax; t += 20) {
    const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
    line.setAttribute('x1', padding);
    line.setAttribute('y1', scaleY(t));
    line.setAttribute('x2', w - padding);
    line.setAttribute('y2', scaleY(t));
    line.setAttribute('class', 'chart-grid-line');
    gridGroup.appendChild(line);

    const text = document.createElementNS('http://www.w3.org/2000/svg', 'text');
    text.setAttribute('x', padding - 10);
    text.setAttribute('y', scaleY(t) + 4);
    text.setAttribute('text-anchor', 'end');
    text.setAttribute('class', 'chart-axis-text');
    text.textContent = t;
    gridGroup.appendChild(text);
  }

  for (let x = 0; x <= hMax; x += Math.ceil(hMax / 5 / 100) * 100) {
    const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
    line.setAttribute('x1', scaleX(x));
    line.setAttribute('y1', padding);
    line.setAttribute('x2', scaleX(x));
    line.setAttribute('y2', h - padding);
    line.setAttribute('class', 'chart-grid-line');
    gridGroup.appendChild(line);

    const text = document.createElementNS('http://www.w3.org/2000/svg', 'text');
    text.setAttribute('x', scaleX(x));
    text.setAttribute('y', h - padding + 15);
    text.setAttribute('text-anchor', 'middle');
    text.setAttribute('class', 'chart-axis-text');
    text.textContent = x;
    gridGroup.appendChild(text);
  }
  svg.appendChild(gridGroup);

  // Draw Axes
  const axes = document.createElementNS('http://www.w3.org/2000/svg', 'g');
  const ax = document.createElementNS('http://www.w3.org/2000/svg', 'line');
  ax.setAttribute('x1', padding); ax.setAttribute('y1', h - padding); ax.setAttribute('x2', w - padding); ax.setAttribute('y2', h - padding);
  ax.setAttribute('class', 'chart-axis-line');
  const ay = document.createElementNS('http://www.w3.org/2000/svg', 'line');
  ay.setAttribute('x1', padding); ay.setAttribute('y1', padding); ay.setAttribute('x2', padding); ay.setAttribute('y2', h - padding);
  ay.setAttribute('class', 'chart-axis-line');
  axes.appendChild(ax); axes.appendChild(ay);
  
  // Axes labels
  const xl = document.createElementNS('http://www.w3.org/2000/svg', 'text');
  xl.setAttribute('x', w / 2); xl.setAttribute('y', h - 10); xl.setAttribute('text-anchor', 'middle'); xl.setAttribute('class', 'chart-axis-text');
  xl.setAttribute('style', 'font-size:12px; fill:var(--text-main);');
  xl.textContent = 'Enthalpy Heat Duty (MW)';
  const yl = document.createElementNS('http://www.w3.org/2000/svg', 'text');
  yl.setAttribute('x', 15); yl.setAttribute('y', h / 2); yl.setAttribute('text-anchor', 'middle'); yl.setAttribute('class', 'chart-axis-text');
  yl.setAttribute('transform', `rotate(-90, 15, ${h / 2})`);
  yl.setAttribute('style', 'font-size:12px; fill:var(--text-main);');
  yl.textContent = 'Temperature (°C)';
  axes.appendChild(xl); axes.appendChild(yl);
  svg.appendChild(axes);

  // Draw Hot Composite Curve (Red)
  const drawCurve = (ccX, ccY, className, markerClass) => {
    const pathGroup = document.createElementNS('http://www.w3.org/2000/svg', 'g');
    let pathD = '';
    ccX.forEach((x, idx) => {
      const sx = scaleX(x);
      const sy = scaleY(ccY[idx]);
      pathD += `${idx === 0 ? 'M' : 'L'} ${sx} ${sy}`;
      
      // Marker dots
      const circle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
      circle.setAttribute('cx', sx);
      circle.setAttribute('cy', sy);
      circle.setAttribute('r', 4);
      circle.setAttribute('class', `chart-marker ${markerClass}`);
      pathGroup.appendChild(circle);
    });

    const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    path.setAttribute('d', pathD);
    path.setAttribute('class', `chart-curve ${className}`);
    pathGroup.insertBefore(path, pathGroup.firstChild);
    svg.appendChild(pathGroup);
  };

  drawCurve(hotCC.H, hotCC.T, 'hot', 'hot');
  drawCurve(coldHShifted, coldCC.T, 'cold', 'cold');

  // Draw Pinch point marker and connector dashed line
  const pinchHot = state.calculatedTargets.pinchHot;
  const pinchCold = state.calculatedTargets.pinchCold;
  
  const pinchXHot = findXAtTemperature(hotCC.T, hotCC.H, pinchHot);
  const pinchXCold = findXAtTemperature(coldCC.T, coldHShifted, pinchCold);

  const pinchGroup = document.createElementNS('http://www.w3.org/2000/svg', 'g');
  
  // Dashed vertical connector line
  const dashLine = document.createElementNS('http://www.w3.org/2000/svg', 'line');
  dashLine.setAttribute('x1', scaleX(pinchXHot));
  dashLine.setAttribute('y1', scaleY(pinchCold));
  dashLine.setAttribute('x2', scaleX(pinchXHot));
  dashLine.setAttribute('y2', scaleY(pinchHot));
  dashLine.setAttribute('stroke', 'var(--svg-match-line)');
  dashLine.setAttribute('stroke-dasharray', '4 4');
  dashLine.setAttribute('stroke-width', '1.5');
  pinchGroup.appendChild(dashLine);

  // Markers on curves
  const m1 = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
  m1.setAttribute('cx', scaleX(pinchXHot)); m1.setAttribute('cy', scaleY(pinchHot));
  m1.setAttribute('r', 6); m1.setAttribute('class', 'chart-pinch-marker');
  const m2 = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
  m2.setAttribute('cx', scaleX(pinchXCold)); m2.setAttribute('cy', scaleY(pinchCold));
  m2.setAttribute('r', 6); m2.setAttribute('class', 'chart-pinch-marker');
  pinchGroup.appendChild(m1); pinchGroup.appendChild(m2);

  // Pinch label
  const pText = document.createElementNS('http://www.w3.org/2000/svg', 'text');
  pText.setAttribute('x', scaleX(pinchXHot) + 12);
  pText.setAttribute('y', scaleY((pinchHot + pinchCold) / 2) + 4);
  pText.setAttribute('fill', 'var(--color-pinch)');
  pText.setAttribute('style', 'font-size: 11px; font-weight: 600;');
  pText.textContent = `Pinch (${pinchHot}°C / ${pinchCold}°C)`;
  pinchGroup.appendChild(pText);

  svg.appendChild(pinchGroup);
}

// --- Draw SVG Grand Composite Curve (GCC) ---
function drawGrandCompositeCurve() {
  const svg = document.getElementById('gcc-svg');
  svg.innerHTML = '';

  const { tempList, Rcas, pinchShifted } = state.calculatedTargets;
  if (tempList.length === 0) return;

  const padding = 50;
  const w = 600;
  const h = 500;

  const rMax = Math.max(...Rcas);
  const tMin = Math.min(...tempList);
  const tMax = Math.max(...tempList);

  const scaleX = (val) => padding + (val / (rMax || 1)) * (w - 2 * padding);
  const scaleY = (val) => h - padding - ((val - tMin) / ((tMax - tMin) || 1)) * (h - 2 * padding);

  // Ticks & Grid
  const gridGroup = document.createElementNS('http://www.w3.org/2000/svg', 'g');
  for (let t = Math.floor(tMin / 20) * 20; t <= tMax; t += 20) {
    const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
    line.setAttribute('x1', padding); line.setAttribute('y1', scaleY(t)); line.setAttribute('x2', w - padding); line.setAttribute('y2', scaleY(t));
    line.setAttribute('class', 'chart-grid-line');
    gridGroup.appendChild(line);

    const text = document.createElementNS('http://www.w3.org/2000/svg', 'text');
    text.setAttribute('x', padding - 10); text.setAttribute('y', scaleY(t) + 4); text.setAttribute('text-anchor', 'end');
    text.setAttribute('class', 'chart-axis-text');
    text.textContent = t;
    gridGroup.appendChild(text);
  }

  for (let x = 0; x <= rMax; x += Math.ceil(rMax / 4 / 50) * 50) {
    const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
    line.setAttribute('x1', scaleX(x)); line.setAttribute('y1', padding); line.setAttribute('x2', scaleX(x)); line.setAttribute('y2', h - padding);
    line.setAttribute('class', 'chart-grid-line');
    gridGroup.appendChild(line);

    const text = document.createElementNS('http://www.w3.org/2000/svg', 'text');
    text.setAttribute('x', scaleX(x)); text.setAttribute('y', h - padding + 15); text.setAttribute('text-anchor', 'middle');
    text.setAttribute('class', 'chart-axis-text');
    text.textContent = x;
    gridGroup.appendChild(text);
  }
  svg.appendChild(gridGroup);

  // Axes
  const axes = document.createElementNS('http://www.w3.org/2000/svg', 'g');
  const ax = document.createElementNS('http://www.w3.org/2000/svg', 'line');
  ax.setAttribute('x1', padding); ax.setAttribute('y1', h - padding); ax.setAttribute('x2', w - padding); ax.setAttribute('y2', h - padding);
  ax.setAttribute('class', 'chart-axis-line');
  const ay = document.createElementNS('http://www.w3.org/2000/svg', 'line');
  ay.setAttribute('x1', padding); ay.setAttribute('y1', padding); ay.setAttribute('x2', padding); ay.setAttribute('y2', h - padding);
  ay.setAttribute('class', 'chart-axis-line');
  axes.appendChild(ax); axes.appendChild(ay);

  // Labels
  const xl = document.createElementNS('http://www.w3.org/2000/svg', 'text');
  xl.setAttribute('x', w / 2); xl.setAttribute('y', h - 10); xl.setAttribute('text-anchor', 'middle'); xl.setAttribute('class', 'chart-axis-text');
  xl.setAttribute('style', 'font-size:12px; fill:var(--text-main);');
  xl.textContent = 'Revised Cascade Heat Flow (MW)';
  const yl = document.createElementNS('http://www.w3.org/2000/svg', 'text');
  yl.setAttribute('x', 15); yl.setAttribute('y', h / 2); yl.setAttribute('text-anchor', 'middle'); yl.setAttribute('class', 'chart-axis-text');
  yl.setAttribute('transform', `rotate(-90, 15, ${h / 2})`);
  yl.setAttribute('style', 'font-size:12px; fill:var(--text-main);');
  yl.textContent = 'Shifted Temperature (°C)';
  axes.appendChild(xl); axes.appendChild(yl);
  svg.appendChild(axes);

  // Draw Curve (GCC - black/white line)
  const pathGroup = document.createElementNS('http://www.w3.org/2000/svg', 'g');
  let pathD = '';
  Rcas.forEach((x, idx) => {
    const sx = scaleX(x);
    const sy = scaleY(tempList[idx]);
    pathD += `${idx === 0 ? 'M' : 'L'} ${sx} ${sy}`;

    const circle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
    circle.setAttribute('cx', sx);
    circle.setAttribute('cy', sy);
    circle.setAttribute('r', 4);
    circle.setAttribute('class', 'chart-marker gcc');
    pathGroup.appendChild(circle);
  });

  const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
  path.setAttribute('d', pathD);
  path.setAttribute('class', 'chart-curve gcc');
  pathGroup.insertBefore(path, pathGroup.firstChild);
  svg.appendChild(pathGroup);

  // Draw Pinch Point indicator horizontal line
  const pinchGroup = document.createElementNS('http://www.w3.org/2000/svg', 'g');
  const horizLine = document.createElementNS('http://www.w3.org/2000/svg', 'line');
  horizLine.setAttribute('x1', scaleX(0)); horizLine.setAttribute('y1', scaleY(pinchShifted));
  horizLine.setAttribute('x2', w - padding); horizLine.setAttribute('y2', scaleY(pinchShifted));
  horizLine.setAttribute('stroke', 'var(--color-pinch)');
  horizLine.setAttribute('stroke-dasharray', '5 3');
  horizLine.setAttribute('stroke-width', '1');
  pinchGroup.appendChild(horizLine);

  const pinchMarker = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
  pinchMarker.setAttribute('cx', scaleX(0)); pinchMarker.setAttribute('cy', scaleY(pinchShifted));
  pinchMarker.setAttribute('r', 6); pinchMarker.setAttribute('class', 'chart-pinch-marker');
  pinchGroup.appendChild(pinchMarker);

  const pText = document.createElementNS('http://www.w3.org/2000/svg', 'text');
  pText.setAttribute('x', scaleX(0) + 12); pText.setAttribute('y', scaleY(pinchShifted) + 4);
  pText.setAttribute('fill', 'var(--color-pinch)');
  pText.setAttribute('style', 'font-size: 11px; font-weight: 600;');
  pText.textContent = `Pinch (${pinchShifted.toFixed(1)}°C shifted)`;
  pinchGroup.appendChild(pText);

  svg.appendChild(pinchGroup);
}

// ==========================================================================
// INTERACTIVE HEN GRID DIAGRAM RENDERER & INTERACTION HANDLERS
// ==========================================================================

function drawHenGrid() {
  const svg = document.getElementById('hen-svg');
  svg.innerHTML = '';

  const hotStreams = state.streams.filter(s => s.type === 'hot');
  const coldStreams = state.streams.filter(s => s.type === 'cold');
  if (hotStreams.length === 0 && coldStreams.length === 0) return;

  // Render SVG gradients
  const defs = document.createElementNS('http://www.w3.org/2000/svg', 'defs');
  defs.innerHTML = `
    <linearGradient id="hot-gradient" x1="0%" y1="0%" x2="100%" y2="0%">
      <stop offset="0%" stop-color="var(--color-hot-start)" />
      <stop offset="100%" stop-color="var(--color-hot-end)" />
    </linearGradient>
    <linearGradient id="cold-gradient" x1="0%" y1="0%" x2="100%" y2="0%">
      <stop offset="0%" stop-color="var(--color-cold-end)" />
      <stop offset="100%" stop-color="var(--color-cold-start)" />
    </linearGradient>
  `;
  svg.appendChild(defs);

  // Dimensions & scaling
  const w = 1000;
  const paddingLeft = 100;
  const paddingRight = 100;
  const activeW = w - paddingLeft - paddingRight;
  const colSpacing = activeW / 9; // 8 slots = 9 spacings
  
  const getSlotX = (slotNum) => paddingLeft + slotNum * colSpacing;
  
  // Calculate stream heights dynamically
  const streamY = {};
  let currentY = 50;
  hotStreams.forEach(h => {
    streamY[h.id] = currentY;
    currentY += 60;
  });
  
  currentY += 40; // spacing between hot and cold streams
  
  coldStreams.forEach(c => {
    streamY[c.id] = currentY;
    currentY += 60;
  });

  // Adjust SVG height based on streams number
  svg.setAttribute('height', currentY + 30);
  svg.setAttribute('width', w);

  // 1. Draw Vertical Pinch Line between Slot 4 and Slot 5
  const pinchX = (getSlotX(4) + getSlotX(5)) / 2;
  const pLine = document.createElementNS('http://www.w3.org/2000/svg', 'line');
  pLine.setAttribute('x1', pinchX);
  pLine.setAttribute('y1', 20);
  pLine.setAttribute('x2', pinchX);
  pLine.setAttribute('y2', currentY);
  pLine.setAttribute('class', 'svg-pinch-line');
  svg.appendChild(pLine);

  // 2. Draw Horizontal Stream Lines
  state.streams.forEach(s => {
    const y = streamY[s.id];
    const isHot = s.type === 'hot';
    
    const lineGroup = document.createElementNS('http://www.w3.org/2000/svg', 'g');
    lineGroup.setAttribute('cursor', 'pointer');
    
    // Main Stream Line
    const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
    line.setAttribute('x1', paddingLeft);
    line.setAttribute('y1', y);
    line.setAttribute('x2', w - paddingRight);
    line.setAttribute('y2', y);
    line.setAttribute('class', `svg-stream-line ${s.type}`);
    
    // Clicking the line selects it in placement step
    line.addEventListener('click', () => handleStreamLineClick(s.id));
    lineGroup.appendChild(line);

    // Flow Arrows
    const arrow = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    const ax = isHot ? w - paddingRight + 10 : paddingLeft - 10;
    const arrowD = isHot 
      ? `M ${ax-10} ${y-6} L ${ax} ${y} L ${ax-10} ${y+6} Z`
      : `M ${ax+10} ${y-6} L ${ax} ${y} L ${ax+10} ${y+6} Z`;
    arrow.setAttribute('d', arrowD);
    arrow.setAttribute('fill', isHot ? 'var(--color-hot-end)' : 'var(--color-cold-end)');
    lineGroup.appendChild(arrow);

    // Stream Name Label on Left
    const label = document.createElementNS('http://www.w3.org/2000/svg', 'text');
    label.setAttribute('x', paddingLeft - 20);
    label.setAttribute('y', y + 4);
    label.setAttribute('text-anchor', 'end');
    label.setAttribute('class', 'svg-stream-label');
    label.textContent = `${s.id} [CP=${s.MCp}]`;
    lineGroup.appendChild(label);

    // Target Temperature Label on Right
    const tempLabel = document.createElementNS('http://www.w3.org/2000/svg', 'text');
    tempLabel.setAttribute('x', w - paddingRight + 20);
    tempLabel.setAttribute('y', y + 4);
    tempLabel.setAttribute('text-anchor', 'start');
    tempLabel.setAttribute('class', 'svg-stream-temp');
    tempLabel.textContent = `${s.Tout}°C`;
    lineGroup.appendChild(tempLabel);

    // Draw boundary temperatures along the stream (0 to 8 boundaries)
    const temps = state.simulation.streamTemps[s.id];
    if (temps) {
      for (let slot = 0; slot <= 8; slot++) {
        const tempX = slot === 0 ? paddingLeft + 5 : getSlotX(slot);
        
        // Display temperature above stream for hot, below for cold
        const tempY = isHot ? y - 10 : y + 18;
        const tempVal = temps[slot];
        
        if (tempVal !== null) {
          const tText = document.createElementNS('http://www.w3.org/2000/svg', 'text');
          tText.setAttribute('x', tempX);
          tText.setAttribute('y', tempY);
          tText.setAttribute('text-anchor', 'middle');
          tText.setAttribute('class', 'svg-stream-temp');
          tText.setAttribute('style', 'font-size: 9px; opacity: 0.85;');
          tText.textContent = `${tempVal.toFixed(0)}°`;
          lineGroup.appendChild(tText);
        }
      }
    }

    svg.appendChild(lineGroup);
  });

  // 3. Draw Heat Exchanger Matches (Connecting Circles)
  state.matches.forEach(m => {
    const x = getSlotX(m.slot);
    const yHot = streamY[m.hotStreamId];
    const yCold = streamY[m.coldStreamId];

    const matchGroup = document.createElementNS('http://www.w3.org/2000/svg', 'g');
    
    // Connecting dashed vertical line
    const link = document.createElementNS('http://www.w3.org/2000/svg', 'line');
    link.setAttribute('x1', x);
    link.setAttribute('y1', yHot);
    link.setAttribute('x2', x);
    link.setAttribute('y2', yCold);
    link.setAttribute('class', 'svg-match-link');
    matchGroup.appendChild(link);

    // Hot side Node
    const c1 = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
    c1.setAttribute('cx', x); c1.setAttribute('cy', yHot); c1.setAttribute('r', 10);
    c1.setAttribute('class', `svg-exchanger-circle hot-node ${m.hasCrossover ? 'crossover' : ''} ${state.selectedMatchId === m.id ? 'selected' : ''}`);
    c1.addEventListener('mousedown', (e) => startMatchDrag(e, m.id));
    c1.addEventListener('click', (e) => selectMatch(e, m.id));
    matchGroup.appendChild(c1);

    // Cold side Node
    const c2 = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
    c2.setAttribute('cx', x); c2.setAttribute('cy', yCold); c2.setAttribute('r', 10);
    c2.setAttribute('class', `svg-exchanger-circle cold-node ${m.hasCrossover ? 'crossover' : ''} ${state.selectedMatchId === m.id ? 'selected' : ''}`);
    c2.addEventListener('mousedown', (e) => startMatchDrag(e, m.id));
    c2.addEventListener('click', (e) => selectMatch(e, m.id));
    matchGroup.appendChild(c2);

    // Text label on connecting line showing heat load
    const loadText = document.createElementNS('http://www.w3.org/2000/svg', 'text');
    loadText.setAttribute('x', x + 6);
    loadText.setAttribute('y', (yHot + yCold) / 2 + 3);
    loadText.setAttribute('class', 'svg-exchanger-label');
    loadText.textContent = `${m.load.toFixed(1)} MW`;
    matchGroup.appendChild(loadText);

    svg.appendChild(matchGroup);
  });

  // 4. Draw Utilities (Heaters / Coolers)
  state.utilities.forEach(u => {
    const x = getSlotX(u.slot);
    const y = streamY[u.streamId];
    const isHeater = u.type === 'heater';

    const utGroup = document.createElementNS('http://www.w3.org/2000/svg', 'g');

    const circle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
    circle.setAttribute('cx', x);
    circle.setAttribute('cy', y);
    circle.setAttribute('r', 14);
    circle.setAttribute('class', `svg-utility-circle ${u.type} ${state.selectedUtilityId === u.id ? 'selected' : ''}`);
    circle.addEventListener('click', (e) => selectUtility(e, u.id));
    utGroup.appendChild(circle);

    const text = document.createElementNS('http://www.w3.org/2000/svg', 'text');
    text.setAttribute('x', x);
    text.setAttribute('y', y + 3);
    text.setAttribute('text-anchor', 'middle');
    text.setAttribute('class', 'svg-utility-text');
    text.textContent = isHeater ? 'H' : 'C';
    utGroup.appendChild(text);

    // Label showing load
    const label = document.createElementNS('http://www.w3.org/2000/svg', 'text');
    label.setAttribute('x', x);
    label.setAttribute('y', isHeater ? y - 20 : y + 28);
    label.setAttribute('text-anchor', 'middle');
    label.setAttribute('class', 'svg-exchanger-label');
    label.textContent = `${u.load.toFixed(1)} MW`;
    utGroup.appendChild(label);

    svg.appendChild(utGroup);
  });
}

// --- Placement / Interaction Workflows ---

function enterInteractionMode(mode) {
  state.interactionMode = mode;
  state.selectedMatchId = null;
  state.selectedUtilityId = null;
  closeEditor();

  const autoBtn = document.getElementById('auto-design-btn');
  const cancelBtn = document.getElementById('clear-design-btn');
  
  if (mode === 'add-exchanger-step1') {
    document.getElementById('add-exchanger-btn').classList.add('btn-primary');
    document.getElementById('add-exchanger-btn').textContent = 'Click Hot Stream...';
  } else if (mode === 'add-utility') {
    document.getElementById('add-utility-btn').classList.add('btn-primary');
    document.getElementById('add-utility-btn').textContent = 'Click Stream...';
  }
}

function handleStreamLineClick(streamId) {
  const stream = state.streams.find(s => s.id === streamId);
  if (!stream) return;

  if (state.interactionMode === 'add-exchanger-step1') {
    if (stream.type !== 'hot') {
      alert("Please select a HOT stream first!");
      return;
    }
    state.pendingHotStreamId = streamId;
    state.interactionMode = 'add-exchanger-step2';
    document.getElementById('add-exchanger-btn').textContent = 'Click Cold Stream...';
  } 
  else if (state.interactionMode === 'add-exchanger-step2') {
    if (stream.type !== 'cold') {
      alert("Please select a COLD stream!");
      return;
    }
    state.pendingColdStreamId = streamId;
    
    // Create new Exchanger Match
    const matchId = `M${state.matches.length + 1}`;
    // Determine default slot (slot 4 if above pinch, slot 5 if below pinch)
    const slot = 4; 
    
    // Calculate default tick-off load
    const hTemps = state.simulation.streamTemps[state.pendingHotStreamId];
    const cTemps = state.simulation.streamTemps[state.pendingColdStreamId];
    const hotRemaining = Math.max(0, (hTemps[0] - hTemps[8]) * state.streams.find(s => s.id === state.pendingHotStreamId).MCp);
    const coldRemaining = Math.max(0, (cTemps[0] - cTemps[8]) * state.streams.find(s => s.id === state.pendingColdStreamId).MCp);
    const defaultLoad = Math.min(100, Math.min(hotRemaining || 100, coldRemaining || 100));

    state.matches.push({
      id: matchId,
      hotStreamId: state.pendingHotStreamId,
      coldStreamId: state.pendingColdStreamId,
      load: Number(defaultLoad.toFixed(1)),
      slot: slot
    });

    resetInteractionMode();
    simulateNetwork();
    renderAll();
    
    // Open editor for this match
    state.selectedMatchId = matchId;
    openEditor(matchId, 'match');
  }
  else if (state.interactionMode === 'add-utility') {
    const isHot = stream.type === 'hot';
    const utId = `U${state.utilities.length + 1}`;
    
    state.utilities.push({
      id: utId,
      streamId: streamId,
      type: isHot ? 'cooler' : 'heater',
      load: 100,
      slot: isHot ? 8 : 1 // Coolers default to slot 8 (cold end), heaters to slot 1 (hot end)
    });

    resetInteractionMode();
    simulateNetwork();
    renderAll();
    
    state.selectedUtilityId = utId;
    openEditor(utId, 'utility');
  }
}

function resetInteractionMode() {
  state.interactionMode = 'normal';
  state.pendingHotStreamId = null;
  state.pendingColdStreamId = null;
  
  document.getElementById('add-exchanger-btn').className = 'btn btn-secondary';
  document.getElementById('add-exchanger-btn').textContent = '+ Add Exchanger Match';
  document.getElementById('add-utility-btn').className = 'btn btn-secondary';
  document.getElementById('add-utility-btn').textContent = '+ Add Utility (Heater/Cooler)';
}

// --- Drag and Drop Exchangers ---
function startMatchDrag(e, matchId) {
  e.stopPropagation();
  e.preventDefault();
  state.draggedMatchId = matchId;
}

function handleSvgMouseMove(e) {
  if (!state.draggedMatchId) return;

  const svg = document.getElementById('hen-svg');
  const rect = svg.getBoundingClientRect();
  const mouseX = e.clientX - rect.left;

  // Map mouseX to closest slot index (1 to 8)
  const w = 1000;
  const paddingLeft = 100;
  const paddingRight = 100;
  const activeW = w - paddingLeft - paddingRight;
  const colSpacing = activeW / 9;

  // slotX = paddingLeft + slot * colSpacing
  // slot = (mouseX - paddingLeft) / colSpacing
  let slot = Math.round((mouseX - paddingLeft) / colSpacing);
  slot = Math.max(1, Math.min(8, slot)); // bound between slot 1 and 8

  const match = state.matches.find(m => m.id === state.draggedMatchId);
  if (match && match.slot !== slot) {
    match.slot = slot;
    simulateNetwork();
    renderAll();
  }
}

function handleSvgMouseUp() {
  state.draggedMatchId = null;
}

// --- Editor Functions ---

function selectMatch(e, matchId) {
  e.stopPropagation();
  state.selectedMatchId = matchId;
  state.selectedUtilityId = null;
  renderAll();
  openEditor(matchId, 'match');
}

function selectUtility(e, utId) {
  e.stopPropagation();
  state.selectedUtilityId = utId;
  state.selectedMatchId = null;
  renderAll();
  openEditor(utId, 'utility');
}

function openEditor(id, type) {
  const panel = document.getElementById('editor-panel');
  const loadInput = document.getElementById('edit-load');
  
  let item = null;
  if (type === 'match') {
    item = state.matches.find(m => m.id === id);
  } else {
    item = state.utilities.find(u => u.id === id);
  }

  if (item) {
    loadInput.value = item.load;
    panel.classList.remove('hidden');
  }
}

function closeEditor() {
  document.getElementById('editor-panel').classList.add('hidden');
  state.selectedMatchId = null;
  state.selectedUtilityId = null;
}

function saveExchangerEdit() {
  const loadVal = parseFloat(document.getElementById('edit-load').value);
  if (isNaN(loadVal) || loadVal <= 0) return;

  if (state.selectedMatchId) {
    const match = state.matches.find(m => m.id === state.selectedMatchId);
    if (match) match.load = loadVal;
  } else if (state.selectedUtilityId) {
    const ut = state.utilities.find(u => u.id === state.selectedUtilityId);
    if (ut) ut.load = loadVal;
  }

  simulateNetwork();
  renderAll();
  closeEditor();
}

function deleteSelectedDevice() {
  if (state.selectedMatchId) {
    state.matches = state.matches.filter(m => m.id !== state.selectedMatchId);
  } else if (state.selectedUtilityId) {
    state.utilities = state.utilities.filter(u => u.id !== state.selectedUtilityId);
  }

  simulateNetwork();
  renderAll();
  closeEditor();
}

// ==========================================================================
// HEURISTIC NETWORK AUTO-DESIGN (PINCH DESIGN METHOD)
// ==========================================================================

function autoDesignNetwork() {
  // Clear any existing design first
  state.matches = [];
  state.utilities = [];
  
  const hotStreams = state.streams.filter(s => s.type === 'hot');
  const coldStreams = state.streams.filter(s => s.type === 'cold');
  if (hotStreams.length === 0 || coldStreams.length === 0) return;

  const pinchHot = state.calculatedTargets.pinchHot;
  const pinchCold = state.calculatedTargets.pinchCold;

  // Enthalpy capacity representations above and below pinch
  const aboveHotLoads = {};
  const aboveColdLoads = {};
  const belowHotLoads = {};
  const belowColdLoads = {};

  hotStreams.forEach(h => {
    aboveHotLoads[h.id] = Math.max(0, (h.Tin - Math.max(h.Tout, pinchHot)) * h.MCp);
    belowHotLoads[h.id] = Math.max(0, (Math.min(h.Tin, pinchHot) - h.Tout) * h.MCp);
  });

  coldStreams.forEach(c => {
    aboveColdLoads[c.id] = Math.max(0, (Math.max(c.Tout, pinchCold) - Math.max(c.Tin, pinchCold)) * c.MCp);
    belowColdLoads[c.id] = Math.max(0, (Math.min(c.Tout, pinchCold) - c.Tin) * c.MCp);
  });

  let matchIndex = 1;
  let utilityIndex = 1;

  // --- ABOVE PINCH: Match streams adjacent to pinch ---
  // Heuristic: CP_hot <= CP_cold for matches close to pinch
  // We sort hot streams ascending by CP, cold streams ascending by CP
  const activeAboveHot = hotStreams.filter(h => aboveHotLoads[h.id] > 0).sort((a,b) => a.MCp - b.MCp);
  const activeAboveCold = coldStreams.filter(c => aboveColdLoads[c.id] > 0).sort((a,b) => a.MCp - b.MCp);

  // Simple heuristic matcher above pinch:
  // Match the highest CP hot stream with the highest CP cold stream in slot 4 (adjacent to pinch)
  // Match lower CP ones in slot 3, 2, 1
  let slotAbove = 4;
  for (let i = activeAboveHot.length - 1; i >= 0; i--) {
    const h = activeAboveHot[i];
    if (aboveHotLoads[h.id] <= 0) continue;

    for (let j = activeAboveCold.length - 1; j >= 0; j--) {
      const c = activeAboveCold[j];
      if (aboveColdLoads[c.id] <= 0) continue;

      // Ensure approach validation heuristic or match
      const load = Math.min(aboveHotLoads[h.id], aboveColdLoads[c.id]);
      if (load > 0.1) {
        state.matches.push({
          id: `M${matchIndex++}`,
          hotStreamId: h.id,
          coldStreamId: c.id,
          load: Number(load.toFixed(1)),
          slot: slotAbove
        });
        aboveHotLoads[h.id] -= load;
        aboveColdLoads[c.id] -= load;
        
        slotAbove = Math.max(1, slotAbove - 1);
      }
    }
  }

  // Any remaining cold loads above pinch must be met by heaters (placed at slot 1 - hot end)
  coldStreams.forEach(c => {
    if (aboveColdLoads[c.id] > 0.1) {
      state.utilities.push({
        id: `U${utilityIndex++}`,
        streamId: c.id,
        type: 'heater',
        load: Number(aboveColdLoads[c.id].toFixed(1)),
        slot: 1
      });
    }
  });

  // --- BELOW PINCH: Match streams adjacent to pinch ---
  // Heuristic: CP_hot >= CP_cold
  // Sort hot streams descending by CP, cold streams descending by CP
  const activeBelowHot = hotStreams.filter(h => belowHotLoads[h.id] > 0).sort((a,b) => b.MCp - a.MCp);
  const activeBelowCold = coldStreams.filter(c => belowColdLoads[c.id] > 0).sort((a,b) => b.MCp - a.MCp);

  let slotBelow = 5; // adjacent to pinch
  for (let i = 0; i < activeBelowHot.length; i++) {
    const h = activeBelowHot[i];
    if (belowHotLoads[h.id] <= 0) continue;

    for (let j = 0; j < activeBelowCold.length; j++) {
      const c = activeBelowCold[j];
      if (belowColdLoads[c.id] <= 0) continue;

      const load = Math.min(belowHotLoads[h.id], belowColdLoads[c.id]);
      if (load > 0.1) {
        state.matches.push({
          id: `M${matchIndex++}`,
          hotStreamId: h.id,
          coldStreamId: c.id,
          load: Number(load.toFixed(1)),
          slot: slotBelow
        });
        belowHotLoads[h.id] -= load;
        belowColdLoads[c.id] -= load;

        slotBelow = Math.min(8, slotBelow + 1);
      }
    }
  }

  // Any remaining hot loads below pinch must be met by coolers (placed at slot 8 - cold end)
  hotStreams.forEach(h => {
    if (belowHotLoads[h.id] > 0.1) {
      state.utilities.push({
        id: `U${utilityIndex++}`,
        streamId: h.id,
        type: 'cooler',
        load: Number(belowHotLoads[h.id].toFixed(1)),
        slot: 8
      });
    }
  });

  simulateNetwork();
  renderAll();
}
