const csvInput = document.querySelector('#csvInput');
const dropzone = document.querySelector('#dropzone');
const message = document.querySelector('#message');
const fileMeta = document.querySelector('#fileMeta');
const rowCount = document.querySelector('#rowCount');
const columnCount = document.querySelector('#columnCount');
const delimiterLabel = document.querySelector('#delimiterLabel');
const previewCount = document.querySelector('#previewCount');
const tableHead = document.querySelector('#tableHead');
const tableBody = document.querySelector('#tableBody');
const previewNote = document.querySelector('#previewNote');
const clearButton = document.querySelector('#clearButton');

const previewLimit = 200;
let activeFile = null;

csvInput.addEventListener('change', handleSelection);
dropzone.addEventListener('dragover', handleDragOver);
dropzone.addEventListener('dragleave', handleDragLeave);
dropzone.addEventListener('drop', handleDrop);
clearButton.addEventListener('click', resetView);

function handleDragOver(event) {
  event.preventDefault();
  dropzone.classList.add('is-dragover');
}

function handleDragLeave() {
  dropzone.classList.remove('is-dragover');
}

function handleDrop(event) {
  event.preventDefault();
  dropzone.classList.remove('is-dragover');
  const [file] = event.dataTransfer.files;

  if (file) {
    processFile(file);
  }
}

async function handleSelection(event) {
  const [file] = event.target.files;
  if (file) {
    await processFile(file);
  }
}

async function processFile(file) {
  activeFile = file;
  clearButton.disabled = false;
  setMessage(`Reading ${file.name}...`);
  fileMeta.textContent = `${formatSize(file.size)} · ${formatDate(file.lastModified)}`;

  try {
    const text = await file.text();
    const delimiter = detectDelimiter(text);
    const rows = parseCsv(text, delimiter);

    if (rows.length === 0) {
      throw new Error('The file does not contain any readable rows.');
    }

    const headers = rows[0].map((value, index) => value?.trim() || `Column ${index + 1}`);
    const dataRows = rows.slice(1);
    const previewRows = dataRows.slice(0, previewLimit);
    const normalizedRows = previewRows.map((row) => padRow(row, headers.length));

    rowCount.textContent = String(dataRows.length);
    columnCount.textContent = String(headers.length);
    delimiterLabel.textContent = describeDelimiter(delimiter);
    previewCount.textContent = String(normalizedRows.length);

    renderTable(headers, normalizedRows);

    if (dataRows.length > previewLimit) {
      previewNote.textContent = `Showing the first ${previewLimit} rows out of ${dataRows.length}.`;
    } else {
      previewNote.textContent = '';
    }

    setMessage(`Loaded ${file.name} successfully.`);
  } catch (error) {
    resetTable();
    rowCount.textContent = '0';
    columnCount.textContent = '0';
    delimiterLabel.textContent = '-';
    previewCount.textContent = '0';
    previewNote.textContent = '';
    setMessage(error.message || 'Unable to parse that file.', true);
  }
}

function resetView() {
  activeFile = null;
  csvInput.value = '';
  clearButton.disabled = true;
  fileMeta.textContent = 'Waiting for upload';
  rowCount.textContent = '0';
  columnCount.textContent = '0';
  delimiterLabel.textContent = '-';
  previewCount.textContent = '0';
  previewNote.textContent = '';
  setMessage('No file selected yet.');
  resetTable();
}

function setMessage(text, isError = false) {
  message.textContent = text;
  message.classList.toggle('error', isError);
}

function resetTable() {
  tableHead.innerHTML = '';
  tableBody.innerHTML = '<tr><td class="empty-state">Upload a CSV to see the parsed data here.</td></tr>';
}

function renderTable(headers, rows) {
  tableHead.innerHTML = `<tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join('')}</tr>`;
  tableBody.innerHTML = rows.length
    ? rows
        .map(
          (row) =>
            `<tr>${row.map((cell) => `<td>${escapeHtml(cell ?? '')}</td>`).join('')}</tr>`,
        )
        .join('')
    : '<tr><td class="empty-state">The file only contains a header row.</td></tr>';
}

function detectDelimiter(text) {
  const sampleLines = text
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .split('\n')
    .filter(Boolean)
    .slice(0, 5);

  const candidates = [',', '\t', ';', '|'];

  return candidates
    .map((delimiter) => ({
      delimiter,
      score: sampleLines.reduce((total, line) => total + countOccurrences(line, delimiter), 0),
    }))
    .sort((left, right) => right.score - left.score)[0]?.delimiter || ',';
}

function describeDelimiter(delimiter) {
  switch (delimiter) {
    case '\t':
      return 'Tab';
    case ';':
      return 'Semicolon';
    case '|':
      return 'Pipe';
    default:
      return 'Comma';
  }
}

function parseCsv(text, delimiter) {
  const rows = [];
  let row = [];
  let cell = '';
  let inQuotes = false;

  const input = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

  for (let index = 0; index < input.length; index += 1) {
    const character = input[index];
    const nextCharacter = input[index + 1];

    if (character === '"') {
      if (inQuotes && nextCharacter === '"') {
        cell += '"';
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (!inQuotes && character === delimiter) {
      row.push(cell);
      cell = '';
      continue;
    }

    if (!inQuotes && character === '\n') {
      row.push(cell);
      if (row.some((value) => value.trim() !== '')) {
        rows.push(row);
      }
      row = [];
      cell = '';
      continue;
    }

    cell += character;
  }

  row.push(cell);
  if (row.some((value) => value.trim() !== '')) {
    rows.push(row);
  }

  if (inQuotes) {
    throw new Error('The CSV file has an unmatched quote.');
  }

  return rows;
}

function padRow(row, targetLength) {
  return Array.from({ length: targetLength }, (_, index) => row[index] ?? '');
}

function formatSize(bytes) {
  if (bytes < 1024) {
    return `${bytes} B`;
  }

  if (bytes < 1024 * 1024) {
    return `${(bytes / 1024).toFixed(1)} KB`;
  }

  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function formatDate(timestamp) {
  if (!timestamp) {
    return 'Unknown date';
  }

  return new Intl.DateTimeFormat(undefined, {
    dateStyle: 'medium',
    timeStyle: 'short',
  }).format(new Date(timestamp));
}

function countOccurrences(text, character) {
  return (text.match(new RegExp(escapeRegex(character), 'g')) || []).length;
}

function escapeRegex(character) {
  return character.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function escapeHtml(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

resetTable();
