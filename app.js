import * as XLSX from 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm';

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

// Validation constants
const MAX_FILE_SIZE = 50 * 1024 * 1024; // 50MB
const MAX_ROWS = 100000;
const DATE_HEADERS = new Set(['DOB', 'Onboarding Deadline', 'Date of Birth']);
const REQUIRED_HEADERS = [
  'Id',
  'Type',
  'First name',
  'Last name',
  'DOB',
  'Employer organisation',
  'Job title',
];
const EXPECTED_HEADERS = new Set([
  'Id',
  'Type',
  'Onboarding status',
  'The apprenticeship we have been advised you are applying for is:',
  'Onboarding Deadline',
  'Employer organisation',
  'First name',
  'Last name',
  'Current programme',
  'Group level 4',
  'DOB',
  'Have you used a previous name?',
  'Job title',
  'What will be your contracted weekly working hours?',
  'Weekly contracted hours',
  'In employment, including self-employment',
  'UK/EEA National',
  'Will you undertake more than 50% of your apprenticeship role within England?',
  'Country of residence',
  'Nationality',
  'Country of birth',
  'Resident in the UK/EEA for 3 years',
  'Requires a Work Permit',
  'In the last 12 months, have you undertaken, or are you planning to undertake, any other government-funded training (excluding apprenticeships)',
  'Details of evidence presented',
  'Will you be contracted for the full duration of your apprenticeship, including the End-Point Assessment?',
  'Will you be paid (at least) the apprenticeship minimum wage for the duration of the apprenticeship?',
  'Please list the full titles of the qualification(s) you will be using to meet the entry requirements of the apprenticeship, including their level and grade (e.g., Level 3 qualifications that hold UCAS points, a degree certificate).',
  'If you hold any additional professional qualifications please list them here (e.g. role specific training, CPD).',
  'Do you have a permanent contract?',
  'Apart from the apprenticeship you are currently applying for right now, are you enrolled on any another apprenticeship?',
  'Have you previously applied or studied with Leeds Beckett University?',
  'If yes, please give your Leeds Beckett Student Number if known.',
  'National insurance number',
]);

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

// Validation functions
function validateFileSize(file) {
  if (file.size > MAX_FILE_SIZE) {
    throw new Error(`File size exceeds maximum limit of ${(MAX_FILE_SIZE / (1024 * 1024)).toFixed(0)}MB. Current size: ${(file.size / (1024 * 1024)).toFixed(2)}MB`);
  }
}

function validateHeaders(headers) {
  const missingHeaders = REQUIRED_HEADERS.filter(
    (header) => !headers.some((h) => h.toLowerCase() === header.toLowerCase())
  );

  if (missingHeaders.length > 0) {
    throw new Error(`Missing required columns: ${missingHeaders.join(', ')}`);
  }

  const unexpectedHeaders = headers.filter(
    (header) => !Array.from(EXPECTED_HEADERS).some((h) => h.toLowerCase() === header.toLowerCase())
  );

  if (unexpectedHeaders.length > 0) {
    console.warn(`Warning: Unexpected columns found: ${unexpectedHeaders.join(', ')}`);
  }
}

function validateRowCount(rowCount) {
  if (rowCount > MAX_ROWS) {
    throw new Error(`File contains ${rowCount} rows, which exceeds the maximum limit of ${MAX_ROWS} rows.`);
  }
}

function validateContent(headers, dataRows) {
  const errors = [];
  const yesNoFields = [
    'Have you used a previous name?',
    'In employment, including self-employment',
    'UK/EEA National',
    'Will you undertake more than 50% of your apprenticeship role within England?',
    'Resident in the UK/EEA for 3 years',
    'Requires a Work Permit',
    'In the last 12 months, have you undertaken, or are you planning to undertake, any other government-funded training (excluding apprenticeships)',
    'Will you be contracted for the full duration of your apprenticeship, including the End-Point Assessment?',
    'Will you be paid (at least) the apprenticeship minimum wage for the duration of the apprenticeship?',
    'Do you have a permanent contract?',
    'Apart from the apprenticeship you are currently applying for right now, are you enrolled on any another apprenticeship?',
    'Have you previously applied or studied with Leeds Beckett University?',
    'Have you spent any time in care?',
  ];
  const getHeaderIndex = (headerName) => headers.findIndex((header) => header.toLowerCase() === headerName.toLowerCase());
  const normalize = (value) => String(value ?? '').trim();

  // Sample validation - check first 100 rows
  const sampleSize = Math.min(100, dataRows.length);
  let invalidCount = 0;

  for (let i = 0; i < sampleSize; i++) {
    const row = dataRows[i];
    const rowNum = i + 2; // +2 because row 1 is headers, row 2 is first data row

    const dobIndex = getHeaderIndex('DOB');
    if (dobIndex !== -1 && normalize(row[dobIndex])) {
      if (!isValidDate(row[dobIndex])) {
        errors.push(`Row ${rowNum}: DOB "${row[dobIndex]}" is not in a valid date format (expected DD/MM/YYYY, DD/MM/YY, MM/DD/YYYY, MM/DD/YY, or YYYY-MM-DD)`);
        invalidCount++;
      }
    }

    const typeIndex = getHeaderIndex('Type');
    if (typeIndex !== -1 && normalize(row[typeIndex])) {
      if (normalize(row[typeIndex]).toLowerCase() !== 'user') {
        errors.push(`Row ${rowNum}: Type "${row[typeIndex]}" must be "User"`);
        invalidCount++;
      }
    }

    const onboardingStatusIndex = getHeaderIndex('Onboarding status');
    if (onboardingStatusIndex !== -1 && normalize(row[onboardingStatusIndex])) {
      const status = normalize(row[onboardingStatusIndex]).toLowerCase();
      if (!['completed', 'in progress'].includes(status)) {
        errors.push(`Row ${rowNum}: Onboarding status "${row[onboardingStatusIndex]}" must be "Completed" or "In Progress"`);
        invalidCount++;
      }
    }

    // Validate Yes/No fields
    yesNoFields.forEach((field) => {
      const fieldIndex = getHeaderIndex(field);
      if (fieldIndex !== -1 && normalize(row[fieldIndex])) {
        const value = normalize(row[fieldIndex]).toLowerCase();
        if (!['yes', 'no', 'y', 'n'].includes(value)) {
          errors.push(`Row ${rowNum}: "${field}" value "${row[fieldIndex]}" must be "Yes" or "No"`);
          invalidCount++;
        }
      }
    });

    // This field is often populated with mixed free text such as "37.5 full time" or "37.5 - TBC".
    // It is therefore treated as descriptive text rather than a strict numeric field.

    const weeklyHoursIndex = getHeaderIndex('Weekly contracted hours');
    if (weeklyHoursIndex !== -1 && normalize(row[weeklyHoursIndex])) {
      if (!/^\d{1,2}:\d{2}$/.test(normalize(row[weeklyHoursIndex]))) {
        errors.push(`Row ${rowNum}: Weekly contracted hours "${row[weeklyHoursIndex]}" must be in HH:MM format`);
        invalidCount++;
      }
    }

    if (invalidCount >= 5) {
      break;
    }
  }

  if (errors.length > 0) {
    const displayErrors = errors.slice(0, 5);
    const message = `Data validation errors found:\n${displayErrors.join('\n')}${errors.length > 5 ? `\n... and ${errors.length - 5} more errors` : ''}`;
    throw new Error(message);
  }
}

function isValidDate(dateString) {
  if (!dateString || typeof dateString !== 'string') {
    return false;
  }

  const trimmed = dateString.trim();

  // Check YYYY-MM-DD format
  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
    const date = new Date(trimmed);
    return !isNaN(date.getTime());
  }

  // Check DD/MM/YYYY, DD/MM/YY, MM/DD/YYYY, or MM/DD/YY format
  if (/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(trimmed)) {
    const parts = trimmed.split('/');
    const day = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10);
    let year = parseInt(parts[2], 10);

    if (parts[2].length === 2) {
      year += year < 50 ? 2000 : 1900;
    }

    const date = new Date(year, month - 1, day);
    const dayMonthMatches = date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day;

    const alternativeDate = new Date(year, day - 1, month);
    const monthDayMatches = alternativeDate.getFullYear() === year && alternativeDate.getMonth() === day - 1 && alternativeDate.getDate() === month;

    return dayMonthMatches || monthDayMatches;
  }

  return false;
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
    // Validate file size
    validateFileSize(file);

    const parsed = await parseFile(file);

    if (parsed.rows.length === 0) {
      throw new Error('The file does not contain any readable rows.');
    }

    const headers = parsed.rows[0].map((value, index) => value?.trim() || `Column ${index + 1}`);
    
    // Validate headers
    validateHeaders(headers);

    const dataRows = parsed.rows.slice(1).filter((row, index) => index !== 0 || !shouldSkipDescriptionRow(headers, row));
    
    // Validate row count
    validateRowCount(dataRows.length);

    // Validate content
    validateContent(headers, dataRows);

    const previewRows = dataRows.slice(0, previewLimit);
    const normalizedRows = previewRows.map((row) => normalizePreviewRow(headers, row));

    rowCount.textContent = String(dataRows.length);
    columnCount.textContent = String(headers.length);
    delimiterLabel.textContent = parsed.sourceLabel;
    previewCount.textContent = String(normalizedRows.length);

    renderTable(headers, normalizedRows);

    if (dataRows.length > previewLimit) {
      previewNote.textContent = `Showing the first ${previewLimit} data rows out of ${dataRows.length}.`;
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

async function parseFile(file) {
  if (isWorkbookFile(file)) {
    return parseWorkbook(file);
  }

  await assertSupportedTextFile(file);
  const text = await file.text();
  const delimiter = detectDelimiter(text);
  return {
    rows: parseCsv(text, delimiter),
    sourceLabel: describeDelimiter(delimiter),
  };
}

function isWorkbookFile(file) {
  const loweredName = file.name.toLowerCase();
  return ['.xlsx', '.xls', '.xlsm', '.ods'].some((extension) => loweredName.endsWith(extension));
}

async function parseWorkbook(file) {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array' });
  const sheetName = workbook.SheetNames[0];

  if (!sheetName) {
    throw new Error('The workbook does not contain any worksheets.');
  }

  const sheet = workbook.Sheets[sheetName];

  if (!sheet) {
    throw new Error('The workbook could not be read.');
  }

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    blankrows: false,
    defval: '',
    raw: false,
  });

  return {
    rows,
    sourceLabel: `Workbook: ${sheetName}`,
  };
}

async function assertSupportedTextFile(file) {
  const loweredName = file.name.toLowerCase();
  const allowedExtensions = ['.csv', '.tsv', '.txt'];

  if (!allowedExtensions.some((extension) => loweredName.endsWith(extension))) {
    throw new Error('Please upload a CSV or Excel workbook file.');
  }

  const header = new Uint8Array(await file.slice(0, 4).arrayBuffer());
  const isZipArchive = header[0] === 0x50 && header[1] === 0x4b;

  if (isZipArchive) {
    throw new Error('This file appears to be a workbook. Use the spreadsheet import flow for Excel files.');
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

function normalizePreviewRow(headers, row) {
  const paddedRow = padRow(row, headers.length);

  return paddedRow.map((cell, index) => {
    const header = headers[index] ?? '';

    if (!DATE_HEADERS.has(header)) {
      return cell;
    }

    return normalizeUkDate(cell);
  });
}

function normalizeUkDate(value) {
  const text = String(value ?? '').trim();

  if (!text) {
    return '';
  }

  const slashMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (slashMatch) {
    const first = Number(slashMatch[1]);
    const second = Number(slashMatch[2]);
    const year = normalizeYearPart(slashMatch[3]);

    if (year == null) {
      return text;
    }

    const ukDate = buildUkDate(year, second, first);
    if (ukDate) {
      return ukDate;
    }

    const alternateDate = buildUkDate(year, first, second);
    if (alternateDate) {
      return alternateDate;
    }
  }

  const isoMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) {
    const ukDate = buildUkDate(Number(isoMatch[1]), Number(isoMatch[2]), Number(isoMatch[3]));
    if (ukDate) {
      return ukDate;
    }
  }

  return text;
}

function normalizeYearPart(yearPart) {
  const year = Number(yearPart);

  if (Number.isNaN(year)) {
    return null;
  }

  if (yearPart.length === 2) {
    return year < 50 ? 2000 + year : 1900 + year;
  }

  return year;
}

function buildUkDate(year, month, day) {
  const parsedYear = Number(year);
  const parsedMonth = Number(month);
  const parsedDay = Number(day);

  if ([parsedYear, parsedMonth, parsedDay].some((value) => Number.isNaN(value))) {
    return null;
  }

  const date = new Date(parsedYear, parsedMonth - 1, parsedDay);

  if (
    date.getFullYear() !== parsedYear ||
    date.getMonth() !== parsedMonth - 1 ||
    date.getDate() !== parsedDay
  ) {
    return null;
  }

  const displayDay = String(parsedDay).padStart(2, '0');
  const displayMonth = String(parsedMonth).padStart(2, '0');
  const displayYear = String(parsedYear).padStart(4, '0');

  return `${displayDay}/${displayMonth}/${displayYear}`;
}

function setMessage(text, isError = false) {
  message.textContent = text;
  message.classList.toggle('error', isError);
}

function resetTable() {
  tableHead.innerHTML = '';
  tableBody.innerHTML = '<tr><td class="empty-state">Upload a CSV or Excel workbook to see the parsed data here.</td></tr>';
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

function shouldSkipDescriptionRow(headers, row) {
  if (!row || row.length === 0) {
    return false;
  }

  const comparedCells = padRow(row, headers.length).map((cell) => String(cell).trim());
  const markerPatterns = [
    /will always be/i,
    /either/i,
    /populated by applicant/i,
    /free field/i,
    /blank if not/i,
    /number in format/i,
    /yes\/no/i,
    /if uk then/i,
    /fixed field/i,
    /set fields/i,
    /removed below data protection/i,
    /student id/i,
    /ni number/i,
  ];

  let markerHits = 0;
  let dataLikeHits = 0;

  comparedCells.forEach((cell) => {
    if (!cell) {
      return;
    }

    if (markerPatterns.some((pattern) => pattern.test(cell))) {
      markerHits += 1;
    }

    if (looksLikeDataCell(cell)) {
      dataLikeHits += 1;
    }
  });

  return markerHits >= 3 || (markerHits >= 2 && dataLikeHits <= 5);
}

function looksLikeDataCell(value) {
  const cell = String(value).trim();

  if (!cell) {
    return false;
  }

  if (/^(yes|no|y|n|true|false|n\/a|na)$/i.test(cell)) {
    return true;
  }

  if (/^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$/.test(cell) || /^\d{4}-\d{2}-\d{2}$/.test(cell)) {
    return true;
  }

  if (/^\d+(?:\.\d+)?$/.test(cell)) {
    return true;
  }

  if (/^[A-Z]{1,4}\d{3,}$/.test(cell)) {
    return true;
  }

  return false;
}

function looksLikeDescriptionCell(value) {
  const cell = String(value).trim();

  if (!cell) {
    return false;
  }

  if (looksLikeDataCell(cell)) {
    return false;
  }

  return cell.length > 12 || /\s/.test(cell);
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
