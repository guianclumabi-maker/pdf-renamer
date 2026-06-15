// ── PDF Page Renamer — app.js ─────────────────────────────────────────────
// Box 1: multi-page PDF  →  each page split into its own PDF file
// Box 2: Excel renaming table  →  rows (in order) give the output filename
//        filename = selected columns concatenated with separator

let pdfBytes   = null;   // raw ArrayBuffer of the uploaded PDF
let pageCount  = 0;      // number of pages in the PDF
let workbook2  = null;   // SheetJS workbook from the renaming table
let mappings   = [];     // [{ pageNum (1-based), outputName }]

// ── Download sample template ──────────────────────────────────────────────
document.getElementById('downloadTemplateBtn').addEventListener('click', () => {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    // Header row
    ['Type of Document', 'Courier name', 'Date'],
    // Sample rows
    ['DL1 COD', 'RAMOS Estephen Pecpec', '2026 06 01'],
    ['DL1 COD', 'SANTOS Maria Cruz',     '2026 06 01'],
    ['DL2 COD', 'DELA CRUZ Juan Pablo',  '2026 06 02'],
  ]);

  // Column widths
  ws['!cols'] = [{ wch: 20 }, { wch: 28 }, { wch: 14 }];

  XLSX.utils.book_append_sheet(wb, ws, 'Renaming Table');
  XLSX.writeFile(wb, 'renaming_table_template.xlsx');
});



// ── DOM refs ──────────────────────────────────────────────────────────────
const dropzone      = document.getElementById('dropzone');
const fileInput     = document.getElementById('fileInput');
const pageInfoWrap  = document.getElementById('page-info-wrap');
const pageChips     = document.getElementById('page-chips');

const dropzone2     = document.getElementById('dropzone2');
const fileInput2    = document.getElementById('fileInput2');
const configPanel   = document.getElementById('config-panel');
const nameSheetSel  = document.getElementById('nameSheetSel');
const headerRow2Inp = document.getElementById('headerRow2');
const colChecklist  = document.getElementById('col-checklist');
const separatorInp  = document.getElementById('separatorInp');
const filenamePrev  = document.getElementById('filename-preview');

const previewSec    = document.getElementById('preview-section');
const previewBody   = document.getElementById('previewBody');
const previewCount  = document.getElementById('previewCount');
const generateBtn   = document.getElementById('generateBtn');
const resetBtn      = document.getElementById('resetBtn');
const progressWrap  = document.getElementById('progress-wrap');
const progressFill  = document.getElementById('progress-fill');
const progressLbl   = document.getElementById('progress-label');
const toast         = document.getElementById('toast');
const toastMsg      = document.getElementById('toastMsg');

// ── Dropzone setup ────────────────────────────────────────────────────────
setupDropzone(dropzone,  fileInput,  handlePdf);
setupDropzone(dropzone2, fileInput2, handleExcel);

function setupDropzone(zone, input, handler) {
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    if (e.dataTransfer.files[0]) {
      input.value = ''; // Reset input value so future dialog selections trigger change
      handler(e.dataTransfer.files[0]);
    }
  });
  // Clear input value before file chooser opens so selecting the same file triggers change
  input.addEventListener('click', () => {
    input.value = '';
  });
  input.addEventListener('change', () => {
    if (input.files[0]) handler(input.files[0]);
  });
}

// ── Config listeners ──────────────────────────────────────────────────────
nameSheetSel.addEventListener('change',  () => { buildColChecklist(); buildPreview(); });
headerRow2Inp.addEventListener('input',  () => { buildColChecklist(); buildPreview(); });
separatorInp.addEventListener('input',   buildPreview);

// ── Reset ─────────────────────────────────────────────────────────────────
resetBtn.addEventListener('click', resetAll);

function resetAll() {
  pdfBytes = null; pageCount = 0; workbook2 = null; mappings = [];
  fileInput.value = ''; fileInput2.value = '';

  resetZone(dropzone,  '📄', 'Drop your <strong>.pdf</strong> file here or <strong>click to browse</strong>');
  resetZone(dropzone2, '🏷️', 'Drop your <strong>renaming table</strong> here or <strong>click to browse</strong>');

  pageInfoWrap.style.display = 'none';
  pageChips.innerHTML = '';
  configPanel.classList.remove('visible');
  colChecklist.innerHTML = '';
  filenamePrev.textContent = '—';
  previewSec.classList.remove('visible');
  generateBtn.disabled = true;
}

function resetZone(zone, icon, labelHtml) {
  zone.classList.remove('has-file');
  zone.querySelector('.icon').textContent = icon;
  zone.querySelector('.label').innerHTML = labelHtml;
}

// ── Box 1: PDF upload ─────────────────────────────────────────────────────
function handlePdf(file) {
  if (!file.name.match(/\.pdf$/i)) {
    showToast('Please upload a PDF file.', 'error');
    return;
  }
  const reader = new FileReader();
  reader.onload = async e => {
    try {
      pdfBytes  = e.target.result;
      const doc = await PDFLib.PDFDocument.load(pdfBytes, { ignoreEncryption: true });
      pageCount = doc.getPageCount();

      markDropzone(dropzone, '📄', `✓  ${file.name}  (${pageCount} page${pageCount !== 1 ? 's' : ''})`);

      // Show page chips (cap display at 50 for UI sanity)
      pageChips.innerHTML = '';
      const shown = Math.min(pageCount, 50);
      for (let i = 1; i <= shown; i++) {
        const chip = document.createElement('span');
        chip.className = 'sheet-chip';
        chip.textContent = `Page ${i}`;
        pageChips.appendChild(chip);
      }
      if (pageCount > 50) {
        const chip = document.createElement('span');
        chip.className = 'sheet-chip';
        chip.textContent = `+${pageCount - 50} more…`;
        pageChips.appendChild(chip);
      }
      pageInfoWrap.style.display = 'block';
      buildPreview();
    } catch (err) {
      console.error(err);
      showToast('Could not read the PDF. Make sure it is a valid (non-encrypted) PDF file.', 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

// ── Box 2: Excel upload ───────────────────────────────────────────────────
function handleExcel(file) {
  if (!file.name.match(/\.(xlsx|xls)$/i)) {
    showToast('Please upload an Excel (.xlsx / .xls) file.', 'error');
    return;
  }
  const reader = new FileReader();
  reader.onload = e => {
    try {
      workbook2 = XLSX.read(e.target.result, { type: 'array' });
      markDropzone(dropzone2, '🏷️', `✓  ${file.name}  (${workbook2.SheetNames.length} sheet${workbook2.SheetNames.length > 1 ? 's' : ''})`);
      populateNameSheets();
      buildColChecklist();
      configPanel.classList.add('visible');
      buildPreview();
    } catch {
      showToast('Could not read the Excel file.', 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

function populateNameSheets() {
  nameSheetSel.innerHTML = '';
  workbook2.SheetNames.forEach((name, i) => {
    const opt = document.createElement('option');
    opt.value = i;
    opt.textContent = name;
    nameSheetSel.appendChild(opt);
  });
}

// ── Renaming table helpers ────────────────────────────────────────────────
function getNameRows() {
  if (!workbook2) return { headers: [], rows: [] };
  const sheetIdx  = parseInt(nameSheetSel.value) || 0;
  const headerRow = Math.max(1, parseInt(headerRow2Inp.value) || 1) - 1;
  const ws   = workbook2.Sheets[workbook2.SheetNames[sheetIdx]];
  const all  = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const headers = (all[headerRow] || []).map(h => h.toString().trim());
  
  const rows = all.slice(headerRow + 1);
  // Trim trailing empty rows to prevent clutter, but preserve intermediate blank rows to keep mapping alignment
  while (rows.length > 0 && !rows[rows.length - 1].some(c => c !== '')) {
    rows.pop();
  }
  return { headers, rows };
}

function buildColChecklist() {
  const { headers, rows } = getNameRows();
  colChecklist.innerHTML = '';
  if (!headers.length) return;

  headers.forEach((h, i) => {
    const id   = `col-check-${i}`;
    const item = document.createElement('label');
    item.className = 'col-check-item';
    item.setAttribute('for', id);

    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.id   = id;
    cb.dataset.colIndex = i;
    cb.checked = true;
    cb.addEventListener('change', buildPreview);

    const span   = document.createElement('span');
    span.textContent = h || `Column ${i + 1}`;

    const sample = rows[0] ? (rows[0][i] || '').toString() : '';
    if (sample) {
      const em = document.createElement('em');
      em.className = 'col-sample';
      em.textContent = sample;
      span.appendChild(em);
    }

    item.appendChild(cb); item.appendChild(span);
    colChecklist.appendChild(item);
  });

  buildPreview();
}

function buildFilename(row) {
  const sep = separatorInp.value;
  return Array.from(colChecklist.querySelectorAll('input[type=checkbox]:checked'))
    .map(cb => (row[parseInt(cb.dataset.colIndex)] || '').toString().trim())
    .filter(v => v)
    .join(sep);
}

// ── Mapping preview ───────────────────────────────────────────────────────
function buildPreview() {
  if (!pageCount) return;

  const { rows } = workbook2 ? getNameRows() : { rows: [] };

  mappings = [];
  const usedNames = new Set();
  for (let i = 0; i < pageCount; i++) {
    const row        = rows[i] || [];
    const name       = workbook2 ? buildFilename(row) : '';
    const baseName   = name || `Page ${i + 1}`;
    let finalName    = sanitize(baseName);

    // Append unique suffix if the name is a duplicate (Windows filesystems are case-insensitive)
    if (usedNames.has(finalName.toLowerCase())) {
      let counter = 1;
      let altName = `${finalName}_${counter}`;
      while (usedNames.has(altName.toLowerCase())) {
        counter++;
        altName = `${finalName}_${counter}`;
      }
      finalName = altName;
    }
    usedNames.add(finalName.toLowerCase());
    mappings.push({ pageNum: i + 1, outputName: finalName, hasName: !!name });
  }

  // Live filename preview (first data row)
  if (workbook2 && rows.length > 0) {
    filenamePrev.textContent = (buildFilename(rows[0]) || '—') + '.pdf';
  }

  // Render table (cap at 200 rows to keep DOM light)
  previewBody.innerHTML = '';
  const shown = Math.min(mappings.length, 200);
  for (let i = 0; i < shown; i++) {
    const m  = mappings[i];
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="color:var(--muted);text-align:center;">Page ${m.pageNum}</td>
      <td class="file-name">${esc(m.outputName)}.pdf</td>
      <td class="${m.hasName ? 'status-ok' : 'status-warn'}">${m.hasName ? '✓ mapped' : '⚠ using page number'}</td>
    `;
    previewBody.appendChild(tr);
  }
  if (mappings.length > 200) {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td colspan="3" style="color:var(--muted);text-align:center;font-size:.78rem;">…and ${mappings.length - 200} more rows (not shown)</td>`;
    previewBody.appendChild(tr);
  }

  previewCount.textContent = `${mappings.length} page${mappings.length !== 1 ? 's' : ''}`;
  previewSec.classList.add('visible');
  generateBtn.disabled = mappings.length === 0;
}

// ── Generate: split PDF + ZIP ─────────────────────────────────────────────
generateBtn.addEventListener('click', async () => {
  if (!pdfBytes || !mappings.length) return;
  generateBtn.disabled = true;
  progressWrap.classList.add('visible');

  const zip     = new JSZip();
  const srcDoc  = await PDFLib.PDFDocument.load(pdfBytes, { ignoreEncryption: true });
  const total   = mappings.length;
  const baseDate = new Date(); // Use sequential modification times to preserve order in file explorers

  for (let idx = 0; idx < total; idx++) {
    const { pageNum, outputName } = mappings[idx];
    setProgress(idx / total, `Extracting page ${pageNum} of ${total}…`);
    await sleep(5);   // yield to repaint

    try {
      // Create a new single-page document
      const newDoc = await PDFLib.PDFDocument.create();
      const [copiedPage] = await newDoc.copyPages(srcDoc, [pageNum - 1]);
      newDoc.addPage(copiedPage);
      const pageBytes = await newDoc.save();
      const fileDate = new Date(baseDate.getTime() + idx * 1000); // 1 second difference per page
      zip.file(`${outputName}.pdf`, pageBytes, { date: fileDate });
    } catch (err) {
      console.warn(`Skipped page ${pageNum}:`, err);
    }
  }

  setProgress(0.98, 'Building ZIP…');
  await sleep(50);

  const blob = await zip.generateAsync({ type: 'blob' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url; a.download = 'renamed_pages.zip';
  a.click();
  URL.revokeObjectURL(url);

  setProgress(1, 'Done!');
  await sleep(400);
  progressWrap.classList.remove('visible');
  generateBtn.disabled = false;
  showToast(`✓  ${total} PDF page${total !== 1 ? 's' : ''} saved to renamed_pages.zip`, 'success');
});

// ── Helpers ───────────────────────────────────────────────────────────────
function markDropzone(zone, icon, text) {
  zone.classList.add('has-file');
  zone.querySelector('.icon').textContent = icon;
  zone.querySelector('.label').textContent = text;
}

function setProgress(frac, label) {
  progressFill.style.width = `${Math.round(frac * 100)}%`;
  progressLbl.textContent  = label;
}

function showToast(msg, type = 'success') {
  toastMsg.textContent = msg;
  toast.className = `visible ${type}`;
  setTimeout(() => { toast.className = ''; }, 5000);
}

function esc(str) {
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function sanitize(name) {
  return name.replace(/[\\/:*?"<>|]/g, '_').trim() || 'page';
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
