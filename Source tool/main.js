import * as XLSX from 'xlsx';

// Global Error Handler
window.onerror = function(msg, url, lineNo, columnNo, error) {
  const statusEl = document.getElementById('status-message');
  if (statusEl) {
    statusEl.textContent = `System Error: ${msg}`;
    statusEl.style.color = 'red';
  }
  return false; // let default handler run
};

// DOM Elements
const dropZone = document.getElementById("drop-zone");
const fileInput = document.getElementById("file-input");
const fileList = document.getElementById('file-list');
const processBtn = document.getElementById('process-btn');
const condensedBtn = document.getElementById('condensed-btn');
const headersBtn = document.getElementById('headers-btn');
const statusMessage = document.getElementById('status-message');

// State
let files = [];
let isProcessing = false;
let detectedHeaders = new Set();

// Event Listeners
dropZone.addEventListener('click', () => fileInput.click());

fileInput.addEventListener('change', (e) => handleFiles(e.target.files));

dropZone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropZone.classList.add('drag-over');
});

dropZone.addEventListener('dragleave', () => {
  dropZone.classList.remove('drag-over');
});

dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  handleFiles(e.dataTransfer.files);
});

processBtn.addEventListener('click', () => processFiles(false));
condensedBtn.addEventListener('click', () => processFiles(true));
headersBtn.addEventListener('click', exportAllHeaders);

function handleFiles(newFiles) {
  if (isProcessing) return;
  
  const validFiles = Array.from(newFiles).filter(file => 
    file.name.match(/\.(xlsx|xls|csv)$/i)
  );

  if (validFiles.length === 0) {
    statusMessage.textContent = 'Please upload a valid Excel or CSV file.';
    statusMessage.style.color = '#ef4444';
    return;
  }

  files = [...files, ...validFiles];
  updateFileList();
  updateStatus();
}

function updateFileList() {
  fileList.innerHTML = '';
  files.forEach((file, index) => {
    const div = document.createElement('div');
    div.className = 'file-item';
    div.innerHTML = `
      <span class="name">${file.name}</span>
      <span class="size">${(file.size / 1024).toFixed(1)} KB</span>
    `;
    fileList.appendChild(div);
  });

  processBtn.disabled = files.length === 0;
  condensedBtn.disabled = files.length === 0;
  headersBtn.disabled = files.length === 0;
}

function updateStatus(msg = 'Ready', color = 'var(--text-muted)') {
  statusMessage.textContent = msg;
  statusMessage.style.color = color;
}

async function processFiles(condensed = false) {
  if (files.length === 0) return;

  isProcessing = true;
  processBtn.disabled = true;
  condensedBtn.disabled = true;
  headersBtn.disabled = true;
  updateStatus(condensed ? 'Generating Condensed List...' : 'Processing All Data...', 'var(--primary)');

  try {
    const allData = [];
    detectedHeaders.clear();

    // Read all files
    for (const file of files) {
      const { data, headers } = await readFile(file);
      allData.push(...data);
      headers.forEach(h => detectedHeaders.add(h));
    }

    // Normalize
    const { normalized, mpnHeader, stats } = normalizeData(allData);
    
    // Export
    if (normalized.length === 0) {
      throw new Error('No data found to process.');
    }

    if (condensed) {
      const condensedData = generateCondensedData(normalized);
      exportFile(condensedData, 'condensed_list');
      updateStatus(`Success! Exported ${condensedData.length} items with important columns only.`, '#22c55e');
    } else {
      exportFile(normalized, 'master_list');
      updateStatus(`Done! Key: "${mpnHeader}". Saved ${stats.total} items (Merged ${stats.merged}).`, '#22c55e');
    }

  } catch (err) {
    console.error(err);
    updateStatus(`Error: ${err.message}`, '#ef4444');
  } finally {
    isProcessing = false;
    processBtn.disabled = false;
    condensedBtn.disabled = false;
    headersBtn.disabled = false;
  }
}

function generateCondensedData(data) {
  // Important columns to keep and their display names
  const importantMap = {
    'Manufacturer': 'Manufacturer',
    'MPN': 'MPN',
    'Min/Mult (MOQ)': 'Min / Mult (MOQ)',
    'Unit Price': 'Unit Price',
    'Stock Status': 'Stock Status',
    'Quantity Avail.': 'Quantity Avail.',
    'Description': 'Description',
    'Datasheet': 'Datasheet',
    'Product Link': 'Product Link'
  };

  return data.map(row => {
    const newRow = {};
    Object.entries(importantMap).forEach(([masterKey, displayName]) => {
      newRow[displayName] = row[masterKey] || "";
    });
    return newRow;
  });
}

async function exportAllHeaders() {
  if (files.length === 0) return;

  isProcessing = true;
  headersBtn.disabled = true;
  updateStatus('Extracting headers...', 'var(--primary)');

  try {
    const allHeaders = [];
    for (const file of files) {
      const { headers } = await readFile(file);
      headers.forEach(h => allHeaders.push(h));
    }

    if (allHeaders.length === 0) {
      throw new Error('No headers found in selected files.');
    }

    const headerList = allHeaders.join('\n');
    const blob = new Blob([headerList], { type: 'text/plain' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);

    a.href = url;
    a.download = `exported_headers_${timestamp}.txt`;
    document.body.appendChild(a);
    a.click();

    setTimeout(() => {
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    }, 100);

    updateStatus(`Exported ${allHeaders.length} total headers to TXT.`, '#22c55e');
  } catch (err) {
    console.error(err);
    updateStatus(`Header Error: ${err.message}`, '#ef4444');
  } finally {
    isProcessing = false;
    headersBtn.disabled = false;
  }
}

function readFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: "binary" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Convert to JSON
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        
        // Get headers correctly using SheetJS
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const headers = [];
        for(let C = range.s.c; C <= range.e.c; ++C) {
          const address = XLSX.utils.encode_col(C) + "1";
          const cell = worksheet[address];
          if(cell && cell.t) headers.push(XLSX.utils.format_cell(cell));
        }

        resolve({ data: jsonData, headers });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsBinaryString(file);
  });
}

function normalizeData(data) {
  if (data.length === 0) return { normalized: [], mpnHeader: null };

  const mapping = {
    'MPN': ['Mfr Part Number', 'BOM: Matched MPN', 'Mrf#', 'MPN', 'Part Number', 'LCSC#'],
    'Manufacturer': ['Manufacturer Name', 'BOM: Manufacturer Name', 'Mfr.'],
    'Description': ['Description', 'BOM: Description'],
    'Required Qty': ['Quantity 1', 'BOM: Requested Qty', 'Quantity'],
    'Unit Price': ['Unit Price 1 (EUR)', 'BOM: Unit Price($)', 'Unit Price(USD)'],
    'Total Price': ['Order Unit Price (EUR)', 'BOM: Total Line Price($)', 'Ext.Price(USD)'],
    'Stock Status': ['Availability', 'BOM: Stock Status', 'Stock Status'],
    'Quantity Avail.': ['BOM: Stock Availability', 'Availability'],
    'Lead Time': ['Lead Time in Days', 'BOM: Mfg Lead Time (weeks)', 'Target Lead Time'],
    'Min/Mult (MOQ)': ['Min./Mult.', 'Min / Mult'],
    'Datasheet': ['Datasheet URL'],
    'Product Link': ['Product Link']
  };

  // Helper to find the master key for a given raw header
  function getMasterKey(rawHeader) {
    const cleanHeader = rawHeader.toLowerCase().trim();
    for (const [master, candidates] of Object.entries(mapping)) {
      if (candidates.some(c => c.toLowerCase().trim() === cleanHeader)) {
        return master;
      }
    }
    return rawHeader; // Keep original if no map
  }

  // Detect which master key represents MPN based on the input data headers
  const firstRowHeaders = Object.keys(data[0]);
  let mpnMasterKey = 'MPN'; // Default
  
  // Find which raw header in this file points to MPN master key
  const actualMpnHeader = firstRowHeaders.find(h => getMasterKey(h) === 'MPN');

  const uniqueMap = new Map();
  const nonMpnRows = [];

  data.forEach((row) => {
    // Create a NEW row using Master Keys
    const mappedRow = {};
    Object.keys(row).forEach(header => {
      const masterKey = getMasterKey(header);
      mappedRow[masterKey] = row[header];
    });

    const mpnValue = mappedRow['MPN'];
    
    if (!mpnValue) {
        nonMpnRows.push(mappedRow);
        return;
    }

    const key = String(mpnValue).trim().toUpperCase();

    if (!uniqueMap.has(key)) {
      uniqueMap.set(key, mappedRow);
    } else {
      const existing = uniqueMap.get(key);
      Object.keys(mappedRow).forEach(k => {
          if ((existing[k] === undefined || existing[k] === "") && (mappedRow[k] !== undefined && mappedRow[k] !== "")) {
              existing[k] = mappedRow[k];
          }
      });
    }
  });
  
  const normalized = [...Array.from(uniqueMap.values()), ...nonMpnRows];

  return { 
    normalized, 
    mpnHeader: 'MPN',
    stats: {
        merged: data.length - normalized.length,
        total: normalized.length,
        original: data.length
    }
  };
}

function exportFile(data, prefix = 'normalized_mpn') {
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Normalized");
  
  // Write to binary string
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  
  // Create Blob
  const blob = new Blob([wbout], { type: "application/octet-stream" });
  
  // Create download link
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement("a");
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  
  a.href = url;
  a.download = `${prefix}_${timestamp}.xlsx`;
  document.body.appendChild(a);
  a.click();
  
  // Cleanup
  setTimeout(() => {
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
  }, 100);
}
