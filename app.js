/* app.js - Validado con los últimos cambios solicitados:
   - card-header: Caso (título, fijo) y Tema (subtítulo).
   - card-meta: muestra en ESTE orden exacto: Tipo de Tarea, Motivo, Subavaloración, Estado.
   - card-quick: UNA fila con tres columnas: Caso | Verificaciones | Observaciones.
   - Se remueven los botones "Ver detalle" y "Abrir detalle".
   - Hacer click en la tarjeta abre el detalle global (como antes).
   - Columna "color" se usa solo como acento visual y no se muestra en texto.
   - Manejo robusto de ausencia de headers/columnas/filas vacías.
   - Añadido envío silencioso a Google Forms cuando se hace click en cualquier tarjeta:
     envia ejecutivo (entry.775437783) y caso/título (entry.315589411) sin mostrar el form.
*/

const DATA_FILE = 'data.xlsx';

// DOM
const caseSearch = document.getElementById('caseSearch');
const clearBtn = document.getElementById('clearBtn');
const fuzzyCheckbox = document.getElementById('fuzzyCheckbox');
const cardsContainer = document.getElementById('cardsContainer');
const detailSection = document.getElementById('detailSection');
const detailContainer = document.getElementById('detailContainer');
const closeDetail = document.getElementById('closeDetail');

let headers = [];
let dataRows = [];
let filtered = [];
let fuse = null;
let colorColIndex = -1;

// Preferred detail order (kept as before)
const preferredDetailOrder = [
  'Tema',
  'Caso', // mapped from Col B
  'Verificaciones',
  'Validación de Datos',
  'Al dia en Pagos',
  'Tipo de Tarea',
  'Motivo',
  'Subvaloración',
  'Estado',
  'Observaciones'
];

/* IndexedDB helper para guardar/leer/borrar el nombre del Ejecutivo.
   Uso: saveEjecutivo(name), getEjecutivo(), deleteEjecutivo()
   Se expone en window para que initEjecutivoUI lo detecte (ya hace typeof check).
*/
(function(){
  const DB_NAME = 'visor-db';
  const DB_VERSION = 1;
  const STORE = 'config';
  const KEY = 'ejecutivo';

  function openDb() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = (ev) => {
        const db = ev.target.result;
        if (!db.objectStoreNames.contains(STORE)) {
          db.createObjectStore(STORE);
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  async function idbGet(key) {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readonly');
      const store = tx.objectStore(STORE);
      const r = store.get(key);
      r.onsuccess = () => {
        resolve(r.result);
        db.close();
      };
      r.onerror = () => {
        reject(r.error);
        db.close();
      };
    });
  }

  async function idbPut(key, value) {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      const store = tx.objectStore(STORE);
      const r = store.put(value, key);
      r.onsuccess = () => {
        resolve(r.result);
        db.close();
      };
      r.onerror = () => {
        reject(r.error);
        db.close();
      };
    });
  }

  async function idbDelete(key) {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      const store = tx.objectStore(STORE);
      const r = store.delete(key);
      r.onsuccess = () => {
        resolve();
        db.close();
      };
      r.onerror = () => {
        reject(r.error);
        db.close();
      };
    });
  }

  // Exportar funciones globalmente (las usa initEjecutivoUI)
  window.getEjecutivo = async function getEjecutivo() {
    try {
      const v = await idbGet(KEY);
      return v === undefined ? '' : v;
    } catch (e) {
      console.error('getEjecutivo error', e);
      return '';
    }
  };

  window.saveEjecutivo = async function saveEjecutivo(name) {
    try {
      if (name === undefined || name === null) name = '';
      await idbPut(KEY, String(name));
      return true;
    } catch (e) {
      console.error('saveEjecutivo error', e);
      throw e;
    }
  };

  window.deleteEjecutivo = async function deleteEjecutivo() {
    try {
      await idbDelete(KEY);
      return true;
    } catch (e) {
      console.error('deleteEjecutivo error', e);
      throw e;
    }
  };
})();

/* ------------ Nuevo módulo: envío silencioso a Google Forms ------------
   Requisitos: al hacer click en cualquier tarjeta enviar ejecutivo + título del caso
   - endpoint formResponse del form (POST, mode: 'no-cors')
   - entry.775437783 -> nombre del ejecutivo
   - entry.315589411 -> nombre del cliente / título del caso
   Implementación: fire-and-forget, no notificaciones al usuario.
*/
async function sendToGoogleForm(ejecutivoName, clienteName) {
  const formBase = 'https://docs.google.com/forms/d/e/1FAIpQLSe9z-6L2GE-JSc-EuKdm_DF-yg2FZ_MUP43Q6oEnBfOM2G56w/formResponse';
  try {
    const fd = new FormData();
    fd.append('entry.775437783', ejecutivoName || '');
    fd.append('entry.315589411', clienteName || '');
    fd.append('timestamp', new Date().toISOString());
    // Fire-and-forget; mode:no-cors to minimise CORS issues. Response is opaque.
    await fetch(formBase, {
      method: 'POST',
      mode: 'no-cors',
      body: fd
    });
    return true;
  } catch (err) {
    // No interrumpimos la UX; solo logueamos.
    console.warn('sendToGoogleForm failed:', err);
    return false;
  }
}

// Events
document.addEventListener('DOMContentLoaded', () => {
  loadData();
  initEjecutivoUI(); // initialize Ejecutivo UI when DOM is ready
  caseSearch.addEventListener('keydown', (e) => { if (e.key === 'Enter') applySearch(); });
  caseSearch.addEventListener('input', () => applySearch());
});
clearBtn.addEventListener('click', () => { caseSearch.value = ''; applySearch(); });

// "Ir arriba" button behavior
closeDetail.addEventListener('click', () => {
  window.scrollTo({ top: 0, behavior: 'smooth' });
  detailSection.classList.add('hidden');
  detailSection.setAttribute('aria-hidden','true');
});

// Load data.xlsx
async function loadData(){
  try {
    const resp = await fetch(DATA_FILE, {cache: "no-store"});
    if (!resp.ok) throw new Error(`HTTP ${resp.status} ${resp.statusText}`);
    const ab = await resp.arrayBuffer();
    const workbook = XLSX.read(ab, {type: 'array'});
    const firstSheet = workbook.SheetNames[0];
    const ws = workbook.Sheets[firstSheet];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval: ''});

    if (!rows || rows.length === 0) {
      console.error('El archivo está vacío o no tiene filas.');
      return;
    }

    // Headers
    const rawHeaders = rows[0];
    headers = rawHeaders.map((h,i) => {
      const t = (h || '').toString().trim();
      return t ? t : `Columna ${String.fromCharCode(65 + i)}`;
    });

    // Detect color column (case-insensitive)
    colorColIndex = headers.findIndex(h => /color/i.test(h));

    // Build dataRows
    dataRows = [];
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || !row.some(cell => cell !== '')) continue; // skip empty rows
      const obj = {};
      for (let c = 0; c < headers.length; c++) {
        obj[headers[c]] = row[c] !== undefined ? String(row[c]) : '';
      }
      obj.__colB = (row[1] !== undefined && row[1] !== null) ? String(row[1]) : '';
      obj.__color = (colorColIndex !== -1 && row[colorColIndex] !== undefined) ? String(row[colorColIndex]) : '';
      dataRows.push(obj);
    }

    setupFuse();
    filtered = dataRows.slice();
    renderCards(filtered);
  } catch (err) {
    console.error('Error cargando data.xlsx desde la raíz:', err);
  }
}

function setupFuse(){
  const options = {
    keys: ['__colB'],
    threshold: 0.45,
    ignoreLocation: true,
    minMatchCharLength: 2,
    includeScore: false
  };
  fuse = new Fuse(dataRows, options);
}

// Search
function applySearch(){
  const term = (caseSearch.value || '').trim();
  const useFuzzy = fuzzyCheckbox && fuzzyCheckbox.checked;

  if (!term) {
    filtered = dataRows.slice();
    renderCards(filtered);
    return;
  }

  if (useFuzzy && fuse) {
    const res = fuse.search(term).map(r => r.item);
    const foundSet = new Set(res);
    filtered = dataRows.filter(r => foundSet.has(r));
  } else {
    const t = term.toLowerCase();
    filtered = dataRows.filter(row => {
      const v = (row.__colB || '').toLowerCase();
      return v.indexOf(t) !== -1;
    });
  }

  renderCards(filtered);
}

// Color helpers
function looksLikeColor(s){
  if (!s) return false;
  const v = s.trim();
  if (/^#([0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$/i.test(v)) return true;
  if (/^(rgb|rgba|hsl|hsla)\s*\(/i.test(v)) return true;
  if (/^[a-z ]{3,}$/i.test(v)) return true;
  return false;
}
function cssColorToRgb(colorStr){
  try {
    const d = document.createElement('div');
    d.style.display = 'none';
    d.style.color = colorStr;
    document.body.appendChild(d);
    const cs = getComputedStyle(d).color;
    document.body.removeChild(d);
    const m = cs.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
    if (!m) return null;
    return { r: +m[1], g: +m[2], b: +m[3] };
  } catch (e) {
    return null;
  }
}

// Render cards with requested structure
function renderCards(rows){
  cardsContainer.innerHTML = '';
  if (!headers || headers.length === 0) return;

  // Prepare display headers for quick row (we won't show color column)
  const maxCols = Math.min(headers.length, 8);
  const displayHeaders = [];
  for (let i = 0; i < maxCols; i++) {
    if (i === colorColIndex) continue;
    displayHeaders.push(headers[i]);
  }

  // Header labels row (non-interactive) - using the quick single layout
  const headerRow = document.createElement('div');
  headerRow.className = 'card-row';
  headerRow.style.pointerEvents = 'none';
  const headerGrid = document.createElement('div');
  headerGrid.className = 'card-quick-single';
  // Columns: Caso | Verificaciones | Observaciones - show labels if those headers exist, otherwise fallback to first 3 non-color headers
  const labelCaso = headers[1] || 'Caso';
  const labelVer = (headers.find(h => /verific/i.test(h)) || 'Verificaciones');
  const labelObs = (headers.find(h => /observ/i.test(h)) || 'Observaciones');
  [labelCaso, labelVer, labelObs].forEach(lbl => {
    const cell = document.createElement('div');
    cell.className = 'card-field';
    const label = document.createElement('div'); label.className = 'label'; label.textContent = lbl;
    const val = document.createElement('div'); val.className = 'value'; val.textContent = '';
    cell.appendChild(label); cell.appendChild(val);
    headerGrid.appendChild(cell);
  });
  headerRow.appendChild(headerGrid);
  cardsContainer.appendChild(headerRow);

  rows.forEach((r, idx) => {
    const card = document.createElement('div');
    card.className = 'card-row card-accent';
    card.setAttribute('data-row-index', idx);

    // Accent color handling: faint tint + left stripe
    const colorVal = (r.__color || '').trim();
    let accentColor = '';
    if (colorVal && looksLikeColor(colorVal)) accentColor = colorVal;

    if (accentColor) {
      const rgb = cssColorToRgb(accentColor);
      const unique = `c-${Math.random().toString(36).slice(2,9)}`;
      card.dataset.accentId = unique;
      if (rgb) {
        card.style.background = `linear-gradient(90deg, rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, 0.06), var(--panel))`;
        const css = `[data-accent-id="${unique}"]::before { background: ${accentColor}; }`;
        const styleTag = document.createElement('style');
        styleTag.textContent = css;
        card.appendChild(styleTag);
        card.setAttribute('data-accent-id', unique);
      } else {
        card.style.borderLeft = `8px solid ${accentColor}`;
      }
    } else {
      card.style.background = 'var(--panel)';
    }

    // CARD HEADER: Title (Caso from Col B) + Subtitle (Tema)
    const header = document.createElement('div'); header.className = 'card-header';
    const titleArea = document.createElement('div'); titleArea.style.flex = '1';
    const title = document.createElement('div'); title.className = 'card-title';
    title.textContent = r.__colB || (headers[1] ? (r[headers[1]] || '') : '');
    const subtitle = document.createElement('div'); subtitle.className = 'card-subtitle';
    // get tema: header 'Tema' or any header containing 'tema'
    let temaText = '';
    if (typeof r['Tema'] !== 'undefined' && r['Tema'] !== '') {
      temaText = r['Tema'];
    } else {
      const tIdx = headers.findIndex(h => /tema/i.test(h));
      if (tIdx !== -1) temaText = r[headers[tIdx]] || '';
    }
    subtitle.textContent = temaText;
    titleArea.appendChild(title);
    titleArea.appendChild(subtitle);
    header.appendChild(titleArea);

    const right = document.createElement('div');
    right.style.fontSize = '13px';
    right.style.color = 'var(--muted)';
    right.textContent = `#${idx + 1}`;
    header.appendChild(right);

    card.appendChild(header);

    // CARD META: exact order: Tipo de Tarea, Motivo, Subavaloración, Estado
    const metaRow = document.createElement('div'); metaRow.className = 'card-meta';
    const metaOrder = ['Tipo de Tarea', 'Motivo', 'Subvaloración', 'Estado'];
    metaOrder.forEach(key => {
      // find header matching exactly (case-insensitive) or containing the key
      let hdr = headers.find(h => h.toLowerCase().trim() === key.toLowerCase().trim());
      if (!hdr) hdr = headers.find(h => h.toLowerCase().includes(key.toLowerCase()));
      if (hdr && r[hdr]) {
        const mi = document.createElement('div'); mi.className = 'meta-item';
        mi.textContent = `${key}: ${r[hdr]}`;
        metaRow.appendChild(mi);
      }
    });
    // only append if something was added
    if (metaRow.childElementCount > 0) card.appendChild(metaRow);

    // QUICK SINGLE ROW: Caso | Verificaciones | Observaciones
    const quick = document.createElement('div'); quick.className = 'card-quick-single';

    // Caso (always from Col B)
    const campoCaso = document.createElement('div'); campoCaso.className = 'card-field';
    const lblCaso = document.createElement('div'); lblCaso.className = 'label'; lblCaso.textContent = headers[1] || 'Caso';
    const valCaso = document.createElement('div'); valCaso.className = 'value'; valCaso.textContent = r.__colB || '';
    campoCaso.appendChild(lblCaso); campoCaso.appendChild(valCaso);
    quick.appendChild(campoCaso);

    // Verificaciones: find header containing 'verific' or a header named 'Verificaciones'
    const verHdr = headers.find(h => /verific/i.test(h)) || headers.find(h => h.toLowerCase().trim() === 'verificaciones');
    const campoVer = document.createElement('div'); campoVer.className = 'card-field';
    const lblVer = document.createElement('div'); lblVer.className = 'label'; lblVer.textContent = verHdr || 'Verificaciones';
    const valVer = document.createElement('div'); valVer.className = 'value'; valVer.textContent = verHdr ? (r[verHdr] || '') : '';
    campoVer.appendChild(lblVer); campoVer.appendChild(valVer);
    quick.appendChild(campoVer);

    // Observaciones: find header containing 'observ' or 'Observaciones'
    const obsHdr = headers.find(h => /observ/i.test(h)) || headers.find(h => h.toLowerCase().trim() === 'observaciones');
    const campoObs = document.createElement('div'); campoObs.className = 'card-field';
    const lblObs = document.createElement('div'); lblObs.className = 'label'; lblObs.textContent = obsHdr || 'Observaciones';
    const valObs = document.createElement('div'); valObs.className = 'value'; valObs.textContent = obsHdr ? (r[obsHdr] || '') : '';
    campoObs.appendChild(lblObs); campoObs.appendChild(valObs);
    quick.appendChild(campoObs);

    if (quick.childElementCount > 0) card.appendChild(quick);

    // Clicking the whole card opens the global detail (previous behavior)
    card.addEventListener('click', () => {
      // 1) Show detail as before
      showDetail(r);

      // 2) Envío silencioso a Google Forms: enviar ejecutivo y título/caso.
      try {
        const ejecutivoToSend = (typeof window !== 'undefined' && window.Ejecutivo) ? String(window.Ejecutivo) : '';
        // Usamos el título visible de la tarjeta: r.__colB (columna B) como "cliente/titular"
        const clienteToSend = r.__colB || '';
        // Fire-and-forget; no bloquear la UI ni la generación de documentos.
        sendToGoogleForm(ejecutivoToSend, clienteToSend).catch(err => {
          // No informar al usuario; solo log para depuración.
          console.warn('Error enviando a Google Forms (silencioso):', err);
        });
      } catch (err) {
        console.warn('Error preparando envío a Google Forms:', err);
      }

      // 3) Scroll to bottom so detail is visible (kept as before)
      setTimeout(() => {
        window.scrollTo({ top: document.body.scrollHeight, behavior: 'smooth' });
      }, 80);
    });

    cardsContainer.appendChild(card);
  });

  // if zero rows, nothing else (headerRow already added)
}

// Show detail below the cards (global detail)
function showDetail(row){
  detailContainer.innerHTML = '';

  const dl = document.createElement('dl');
  dl.style.margin = '0';

  // Build header set for checks
  const headerSet = new Set(headers.map(h => (h || '').toString().trim()));

  preferredDetailOrder.forEach(key => {
    if (key === 'Caso') {
      const dt = document.createElement('dt'); dt.textContent = 'Caso';
      const dd = document.createElement('dd'); dd.textContent = row.__colB || '';
      dl.appendChild(dt); dl.appendChild(dd);
      return;
    }
    // find exact or case-insensitive header match
    let matchHeader = null;
    for (const h of headers) {
      if ((h || '').toString().trim().toLowerCase() === key.toLowerCase()) {
        matchHeader = h; break;
      }
    }
    if (!matchHeader) {
      // fallback: header containing the key
      matchHeader = headers.find(h => (h || '').toString().toLowerCase().includes(key.toLowerCase()));
    }
    if (matchHeader) {
      const dt = document.createElement('dt'); dt.textContent = key;
      const dd = document.createElement('dd'); dd.textContent = row[matchHeader] !== undefined ? row[matchHeader] : '';
      dl.appendChild(dt); dl.appendChild(dd);
    }
  });

  // Do NOT include color column text per request

  detailContainer.appendChild(dl);

  detailSection.classList.remove('hidden');
  detailSection.setAttribute('aria-hidden','false');

  // Smooth scroll to bottom so detail is visible
  setTimeout(() => {
    window.scrollTo({ top: document.body.scrollHeight, behavior: 'smooth' });
  }, 50);
}

/* ------------ Ejecutivo UI (sin cambios funcionales respecto a tu original) ------------
*/
function initEjecutivoUI() {
  const gearBtn = document.getElementById('gearBtn');
  const modal = document.getElementById('ejecutivoModal');
  const ejecutivoInput = document.getElementById('ejecutivoInput');
  const editBtn = document.getElementById('editEjecutivoBtn');
  const deleteBtn = document.getElementById('deleteEjecutivoBtn');
  const cancelBtn = document.getElementById('cancelEjecutivoBtn');
  const acceptBtn = document.getElementById('acceptEjecutivoBtn');
  const nameSpan = document.getElementById('ejecutivoName');

  window.Ejecutivo = '';

  (async function loadAndRender() {
    try {
      const stored = (typeof getEjecutivo === 'function') ? await getEjecutivo() : '';
      window.Ejecutivo = stored || '';
      renderEjecutivoName();
    } catch (err) {
      console.error('No se pudo leer Ejecutivo desde IndexedDB', err);
    }
  })();

  function renderEjecutivoName() {
    if (!nameSpan) return;
    if (window.Ejecutivo && String(window.Ejecutivo).trim() !== '') {
      nameSpan.textContent = String(window.Ejecutivo);
      nameSpan.title = `Ejecutivo: ${window.Ejecutivo}`;
    } else {
      nameSpan.textContent = '';
      nameSpan.title = '';
    }
  }

  function openModal() {
    if (!modal) return;
    ejecutivoInput.value = window.Ejecutivo || '';
    ejecutivoInput.setAttribute('readonly', 'readonly');
    if (!window.Ejecutivo) {
      ejecutivoInput.removeAttribute('readonly');
      ejecutivoInput.focus();
    }
    modal.hidden = false;
    acceptBtn.focus();
  }

  function closeModal() {
    if (!modal) return;
    modal.hidden = true;
  }

  gearBtn && gearBtn.addEventListener('click', (e) => {
    openModal();
  });

  editBtn && editBtn.addEventListener('click', () => {
    ejecutivoInput.removeAttribute('readonly');
    ejecutivoInput.focus();
    const val = ejecutivoInput.value;
    ejecutivoInput.value = '';
    ejecutivoInput.value = val;
  });

  deleteBtn && deleteBtn.addEventListener('click', async () => {
    ejecutivoInput.removeAttribute('readonly');
    ejecutivoInput.value = '';
    ejecutivoInput.focus();
  });

  cancelBtn && cancelBtn.addEventListener('click', (e) => {
    e.preventDefault();
    ejecutivoInput.value = window.Ejecutivo || '';
    ejecutivoInput.setAttribute('readonly', 'readonly');
    closeModal();
  });

  acceptBtn && acceptBtn.addEventListener('click', async (e) => {
    e.preventDefault();
    const newName = (ejecutivoInput.value || '').trim();
    try {
      if (!newName) {
        if (typeof deleteEjecutivo === 'function') await deleteEjecutivo();
        window.Ejecutivo = '';
      } else {
        if (typeof saveEjecutivo === 'function') await saveEjecutivo(newName);
        window.Ejecutivo = newName;
      }
      renderEjecutivoName();
      ejecutivoInput.setAttribute('readonly', 'readonly');
      closeModal();
    } catch (err) {
      console.error('Error guardando/eliminando Ejecutivo', err);
      // if showMessage exists use it, otherwise simple alert fallback
      if (typeof showMessage === 'function') {
        showMessage('No se pudo guardar el nombre del Ejecutivo. Revisa la consola.', true, 6000);
      } else {
        alert('No se pudo guardar el nombre del Ejecutivo. Revisa la consola.');
      }
    }
  });

  document.addEventListener('keydown', (ev) => {
    if (ev.key === 'Escape') {
      const modalVisible = modal && !modal.hidden;
      if (modalVisible) {
        cancelBtn && cancelBtn.click();
      }
    }
  });

  modal && modal.addEventListener('click', (ev) => {
    if (ev.target === modal) {
      cancelBtn && cancelBtn.click();
    }
  });
}