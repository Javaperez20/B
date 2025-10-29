/* app.js - Lee data.xlsx desde la raíz y permite buscar en Columna B.
   - Un único input: "Buscar por caso" (busca en Col B, índice 1).
   - Toggle fuzzy (Fuse.js) para búsqueda difusa.
   - Resultados mostrados como filas-encapsuladas (cards), cada fila en su contenedor.
   - No hay scrollbars internos en la lista; al seleccionar una fila se desplaza al final y muestra detalle.
   - Mantén data.xlsx en la raíz.
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

// Events
document.addEventListener('DOMContentLoaded', () => {
  loadData();
  caseSearch.addEventListener('keydown', (e) => { if (e.key === 'Enter') applySearch(); });
  caseSearch.addEventListener('input', () => applySearch()); // live search
});
clearBtn.addEventListener('click', () => { caseSearch.value = ''; applySearch(); });
closeDetail.addEventListener('click', () => { detailSection.classList.add('hidden'); detailSection.setAttribute('aria-hidden','true'); });

// Load data.xlsx silently from root
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

    // Headers: row 0
    const rawHeaders = rows[0];
    headers = rawHeaders.map((h,i) => {
      const t = (h || '').toString().trim();
      return t ? t : `Columna ${String.fromCharCode(65 + i)}`;
    });

    // Data rows from row index 1 onward
    dataRows = [];
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row.some(cell => cell !== '')) continue; // skip empty rows
      const obj = {};
      for (let c = 0; c < headers.length; c++) {
        obj[headers[c]] = row[c] !== undefined ? String(row[c]) : '';
      }
      obj.__colB = row[1] !== undefined ? String(row[1]) : '';
      dataRows.push(obj);
    }

    // Setup Fuse (for fuzzy) on demand
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

// Apply search: substring (case-insensitive) by default, fuzzy when checked
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
    // keep order as in original dataRows while preserving only matches
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

// Render results as card list (each row enclosed in a container)
function renderCards(rows){
  cardsContainer.innerHTML = '';
  // If no headers or no data, render nothing (silent)
  if (!headers || headers.length === 0) return;

  // Limit columns displayed to first 8 (A..H)
  const maxCols = Math.min(headers.length, 8);
  const displayHeaders = headers.slice(0, maxCols);

  // Create a header card (labels)
  const headerRow = document.createElement('div');
  headerRow.className = 'card-row';
  headerRow.style.pointerEvents = 'none';
  const headerGrid = document.createElement('div');
  headerGrid.className = 'row-grid';
  displayHeaders.forEach(h => {
    const cell = document.createElement('div');
    cell.className = 'row-field';
    const label = document.createElement('div'); label.className = 'label'; label.textContent = h;
    cell.appendChild(label);
    headerGrid.appendChild(cell);
  });
  headerRow.appendChild(headerGrid);
  cardsContainer.appendChild(headerRow);

  // Rows as cards
  rows.forEach((r, idx) => {
    const card = document.createElement('div');
    card.className = 'card-row';
    card.setAttribute('data-row-index', idx);

    const grid = document.createElement('div');
    grid.className = 'row-grid';

    displayHeaders.forEach(h => {
      const field = document.createElement('div');
      field.className = 'row-field';
      const label = document.createElement('div'); label.className = 'label'; label.textContent = h;
      const value = document.createElement('div'); value.className = 'value'; value.textContent = r[h] !== undefined ? r[h] : '';
      field.appendChild(label);
      field.appendChild(value);
      grid.appendChild(field);
    });

    card.appendChild(grid);
    card.addEventListener('click', () => {
      showDetail(r);
      // scroll to bottom so the detail (rendered below) is visible
      // small timeout to allow DOM update
      setTimeout(() => {
        window.scrollTo({ top: document.body.scrollHeight, behavior: 'smooth' });
      }, 80);
    });

    cardsContainer.appendChild(card);
  });

  // If there are zero data rows, still render header and nothing else
  if (rows.length === 0) {
    // nothing else
  }
}

// Show detail below the cards (at the end of page)
function showDetail(row){
  detailContainer.innerHTML = '';

  const dl = document.createElement('dl');
  dl.style.margin = '0';

  headers.forEach(h => {
    const dt = document.createElement('dt'); dt.textContent = h;
    const dd = document.createElement('dd'); dd.textContent = row[h] !== undefined ? row[h] : '';
    dl.appendChild(dt);
    dl.appendChild(dd);
  });

  // reference to Col B
  const dtB = document.createElement('dt'); dtB.textContent = `Columna B (${headers[1] || 'Columna B'})`;
  const ddB = document.createElement('dd'); ddB.textContent = row.__colB || '';
  dl.appendChild(dtB);
  dl.appendChild(ddB);

  detailContainer.appendChild(dl);

  detailSection.classList.remove('hidden');
  detailSection.setAttribute('aria-hidden','false');

  // Smooth scroll to bottom (ensure visible)
  setTimeout(() => {
    window.scrollTo({ top: document.body.scrollHeight, behavior: 'smooth' });
  }, 50);
}