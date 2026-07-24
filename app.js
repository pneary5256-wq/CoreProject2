import * as XLSX from 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm';

const csvInput = document.querySelector('#csvInput');
const dropzone = document.querySelector('#dropzone');
const message = document.querySelector('#message');
const fileMeta = document.querySelector('#fileMeta');
const totalApplicantsCount = document.querySelector('#totalApplicantsCount');
const eligibleCount = document.querySelector('#eligibleCount');
const ineligibleCount = document.querySelector('#ineligibleCount');
const reviewCount = document.querySelector('#reviewCount');
const missingDataCount = document.querySelector('#missingDataCount');
const applicantList = document.querySelector('#applicantList');
const applicantDetails = document.querySelector('#applicantDetails');
const summaryOutput = document.querySelector('#summaryOutput');
const configureRulesButton = document.querySelector('#configureRulesButton');
const copySummaryButton = document.querySelector('#copySummaryButton');
const clearButton = document.querySelector('#clearButton');
const rulesPanel = document.querySelector('#rulesPanel');
const closeRulesButton = document.querySelector('#closeRulesButton');
const rulesForm = document.querySelector('#rulesForm');
const saveRulesButton = document.querySelector('#saveRulesButton');
const updatePreferencesButton = document.querySelector('#updatePreferencesButton');
const resetRulesButton = document.querySelector('#resetRulesButton');
const ageCutoffInput = document.querySelector('#ageCutoffInput');
const workingHoursReviewInput = document.querySelector('#workingHoursReviewInput');
const workingHoursFailInput = document.querySelector('#workingHoursFailInput');
const countryOfResidenceInput = document.querySelector('#countryOfResidenceInput');
const nationalityInput = document.querySelector('#nationalityInput');
const countryOfBirthInput = document.querySelector('#countryOfBirthInput');
const rulesSummary = document.querySelector('#rulesSummary');
const yesNoRules = document.querySelector('#yesNoRules');

const YES_NO_RULE_DEFINITIONS = [
  { key: 'previousName', label: 'Have you used a previous name?', expected: 'no' },
  { key: 'inEmployment', label: 'In employment, including self-employment', expected: 'yes' },
  { key: 'ukEeaNational', label: 'UK/EEA National', expected: 'yes' },
  { key: 'withinEngland', label: 'More than 50% of role within England', expected: 'yes' },
  { key: 'residentThreeYears', label: 'Resident in the UK/EEA for 3 years', expected: 'yes' },
  { key: 'requiresWorkPermit', label: 'Requires a Work Permit', expected: 'no' },
  { key: 'otherGovernmentTraining', label: 'Other government-funded training', expected: 'no' },
  { key: 'contractedFullDuration', label: 'Contracted for full duration', expected: 'yes' },
  { key: 'paidMinimumWage', label: 'Paid minimum wage', expected: 'yes' },
  { key: 'permanentContract', label: 'Do you have a permanent contract?', expected: 'yes' },
  { key: 'anotherApprenticeship', label: 'Another apprenticeship', expected: 'yes' },
  { key: 'leedsBeckettPreviousStudent', label: 'Have you previously applied or studied with Leeds Beckett University?', expected: 'no' },
];

const DEFAULT_RULE_CONFIG = {
  ageCutoffDate: '2026-09-01',
  workingHoursReviewThreshold: 30,
  workingHoursFailThreshold: 48,
  countryOfResidence: 'UnitedKingdom',
  nationality: 'UKNational',
  countryOfBirth: 'UnitedKingdom',
  yesNoAnswers: YES_NO_RULE_DEFINITIONS.reduce((answers, definition) => {
    answers[definition.key] = definition.expected;
    return answers;
  }, {}),
};

const RULE_CONFIG_PATH = 'rules-config.json';

const ruleState = { ...DEFAULT_RULE_CONFIG };
let sharedRuleConfig = { ...DEFAULT_RULE_CONFIG };
const appState = {
  applicants: [],
  selectedApplicantIndex: -1,
};

// Validation constants
const MAX_FILE_SIZE = 50 * 1024 * 1024; // 50MB
const MAX_ROWS = 100000;
const REQUIRED_HEADERS = [
  'Id',
  'Type',
  'First name',
  'Last name',
  'Email',
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
  'Email',
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
configureRulesButton.addEventListener('click', (event) => {
  renderYesNoRules();
  syncRulesForm();
  renderRulesSummary();
});
closeRulesButton.addEventListener('click', () => {
  if (window.location.hash === '#rulesPanel') {
    history.replaceState(null, '', window.location.pathname + window.location.search);
  }
});
rulesForm.addEventListener('submit', handleRulesSubmit);
updatePreferencesButton.addEventListener('click', handleUpdatePreferences);
resetRulesButton.addEventListener('click', resetRulesToDefault);
copySummaryButton.addEventListener('click', () => {
  setMessage('Summary output is not enabled yet.');
});
applicantList.addEventListener('click', handleApplicantListClick);
void initApp();

async function initApp() {
  await loadSharedRuleConfig();
  renderRulesPanel();
  resetView();
}

async function loadSharedRuleConfig() {
  try {
    const response = await fetch(RULE_CONFIG_PATH, { cache: 'no-store' });

    if (!response.ok) {
      throw new Error(`Unable to load ${RULE_CONFIG_PATH}`);
    }

    const config = await response.json();
    sharedRuleConfig = normalizeRuleConfig(config);
  } catch (error) {
    console.warn(error);
    sharedRuleConfig = { ...DEFAULT_RULE_CONFIG };
  }

  applyRuleConfig(sharedRuleConfig);
}

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

    const emailIndex = getHeaderIndex('Email');
    if (emailIndex !== -1 && normalize(row[emailIndex])) {
      if (!normalize(row[emailIndex]).includes('@')) {
        errors.push(`Row ${rowNum}: Email "${row[emailIndex]}" must contain "@"`);
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
  return parseFlexibleDate(dateString) != null;
}

async function handleSelection(event) {
  const [file] = event.target.files;
  if (file) {
    await processFile(file);
  }
}

async function processFile(file) {
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

    validateHeaders(headers);

    const dataRows = parsed.rows.slice(1);
    if (dataRows.length > 0 && shouldSkipDescriptionRow(headers, dataRows[0])) {
      // Some exports include a short description row immediately after the header row.
      dataRows.shift();
    }

    validateRowCount(dataRows.length);
    validateContent(headers, dataRows);

    const applicants = dataRows.map((row, index) => evaluateApplicant(headers, row, index));
    appState.applicants = applicants;
    appState.selectedApplicantIndex = -1;

    fileMeta.textContent = `${file.name} · ${parsed.sourceLabel}`;
    renderDashboard(applicants);

    setMessage(`Loaded ${file.name} successfully.`);
  } catch (error) {
    resetView();
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
  csvInput.value = '';
  clearButton.disabled = true;
  fileMeta.textContent = 'Waiting for upload';
  totalApplicantsCount.textContent = '0';
  eligibleCount.textContent = '0';
  ineligibleCount.textContent = '0';
  reviewCount.textContent = '0';
  missingDataCount.textContent = '0';
  applicantList.innerHTML = '';
  applicantDetails.innerHTML = '<div class="empty-state-panel">Select an applicant to inspect the criteria breakdown.</div>';
  summaryOutput.value = '';
  setMessage('No file selected yet.');
  appState.applicants = [];
  appState.selectedApplicantIndex = -1;
}

function openRulesPanel() {
  syncRulesForm();
  renderRulesSummary();
  rulesPanel.scrollIntoView({ block: 'start', behavior: 'smooth' });
  ageCutoffInput.focus();
}

function closeRulesPanel() {
  configureRulesButton.focus();
}

function handleRulesSubmit(event) {
  event.preventDefault();

  try {
    const nextConfig = readRuleConfigFromForm();

    if (nextConfig.workingHoursReviewThreshold >= nextConfig.workingHoursFailThreshold) {
      throw new Error('The review threshold must be lower than the fail threshold.');
    }

    applyRuleConfig(nextConfig);
    renderRulesSummary();

    if (appState.applicants.length > 0) {
      renderDashboard(appState.applicants);
    }

    setMessage('Rule settings updated for this session. Use Update Preferences to export the JSON file.');
    closeRulesPanel();
  } catch (error) {
    setMessage(error.message || 'Unable to save rules.', true);
  }
}

function handleUpdatePreferences() {
  try {
    const nextConfig = readRuleConfigFromForm();

    if (nextConfig.workingHoursReviewThreshold >= nextConfig.workingHoursFailThreshold) {
      throw new Error('The review threshold must be lower than the fail threshold.');
    }

    const normalizedConfig = normalizeRuleConfig(nextConfig);
    downloadRuleConfig(normalizedConfig);
    setMessage('rules-config.json downloaded. Replace the shared file with this version to update preferences.');
  } catch (error) {
    setMessage(error.message || 'Unable to update preferences.', true);
  }
}

function resetRulesToDefault() {
  applyRuleConfig(sharedRuleConfig);
  syncRulesForm();
  renderRulesSummary();

  if (appState.applicants.length > 0) {
    renderDashboard(appState.applicants);
  }

  setMessage('Rule settings reset to the shared defaults.');
}

function downloadRuleConfig(config) {
  const blob = new Blob([`${JSON.stringify(config, null, 2)}\n`], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');

  anchor.href = url;
  anchor.download = 'rules-config.json';
  anchor.rel = 'noopener';
  document.body.append(anchor);
  anchor.click();
  anchor.remove();

  window.setTimeout(() => URL.revokeObjectURL(url), 0);
}

function applyRuleConfig(nextConfig) {
  ruleState.ageCutoffDate = nextConfig.ageCutoffDate;
  ruleState.workingHoursReviewThreshold = nextConfig.workingHoursReviewThreshold;
  ruleState.workingHoursFailThreshold = nextConfig.workingHoursFailThreshold;
  ruleState.countryOfResidence = nextConfig.countryOfResidence;
  ruleState.nationality = nextConfig.nationality;
  ruleState.countryOfBirth = nextConfig.countryOfBirth;
  ruleState.yesNoAnswers = { ...nextConfig.yesNoAnswers };
}

function normalizeRuleConfig(config) {
  const source = config && typeof config === 'object' ? config : {};
  const ageCutoffDate = typeof source.ageCutoffDate === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(source.ageCutoffDate.trim())
    ? source.ageCutoffDate.trim()
    : DEFAULT_RULE_CONFIG.ageCutoffDate;
  const workingHoursReviewThreshold = Number(source.workingHoursReviewThreshold);
  const workingHoursFailThreshold = Number(source.workingHoursFailThreshold);
  const countryOfResidence = typeof source.countryOfResidence === 'string' && source.countryOfResidence.trim()
    ? source.countryOfResidence.trim()
    : DEFAULT_RULE_CONFIG.countryOfResidence;
  const nationality = typeof source.nationality === 'string' && source.nationality.trim()
    ? source.nationality.trim()
    : DEFAULT_RULE_CONFIG.nationality;
  const countryOfBirth = typeof source.countryOfBirth === 'string' && source.countryOfBirth.trim()
    ? source.countryOfBirth.trim()
    : DEFAULT_RULE_CONFIG.countryOfBirth;
  const yesNoAnswers = YES_NO_RULE_DEFINITIONS.reduce((answers, definition) => {
    answers[definition.key] = normalizeYesNoAnswer(source.yesNoAnswers?.[definition.key], definition.expected);
    return answers;
  }, {});

  return {
    ageCutoffDate,
    workingHoursReviewThreshold: Number.isFinite(workingHoursReviewThreshold)
      ? workingHoursReviewThreshold
      : DEFAULT_RULE_CONFIG.workingHoursReviewThreshold,
    workingHoursFailThreshold: Number.isFinite(workingHoursFailThreshold)
      ? workingHoursFailThreshold
      : DEFAULT_RULE_CONFIG.workingHoursFailThreshold,
    countryOfResidence,
    nationality,
    countryOfBirth,
    yesNoAnswers,
  };
}

function renderRulesPanel() {
  renderYesNoRules();
  syncRulesForm();
  renderRulesSummary();
}

function syncRulesForm() {
  ageCutoffInput.value = ruleState.ageCutoffDate;
  workingHoursReviewInput.value = String(ruleState.workingHoursReviewThreshold);
  workingHoursFailInput.value = String(ruleState.workingHoursFailThreshold);
  countryOfResidenceInput.value = ruleState.countryOfResidence;
  nationalityInput.value = ruleState.nationality;
  countryOfBirthInput.value = ruleState.countryOfBirth;

  if (yesNoRules) {
    yesNoRules.querySelectorAll('[data-yes-no-key]').forEach((field) => {
      const key = field.dataset.yesNoKey;
      const select = field.querySelector('select');

      if (select) {
        select.value = normalizeYesNoAnswer(ruleState.yesNoAnswers?.[key], DEFAULT_RULE_CONFIG.yesNoAnswers[key]);
      }
    });
  }
}

function readRuleConfigFromForm() {
  const ageCutoffDate = ageCutoffInput.value.trim();
  const workingHoursReviewThreshold = Number(workingHoursReviewInput.value);
  const workingHoursFailThreshold = Number(workingHoursFailInput.value);
  const countryOfResidence = countryOfResidenceInput.value.trim();
  const nationality = nationalityInput.value.trim();
  const countryOfBirth = countryOfBirthInput.value.trim();
  const yesNoAnswers = {};

  if (!ageCutoffDate) {
    throw new Error('Age cutoff date is required.');
  }

  if (Number.isNaN(workingHoursReviewThreshold) || Number.isNaN(workingHoursFailThreshold)) {
    throw new Error('Working hour thresholds must be valid numbers.');
  }

  if (!countryOfResidence) {
    throw new Error('Expected country of residence is required.');
  }

  if (!nationality) {
    throw new Error('Nationality is required.');
  }

  if (!countryOfBirth) {
    throw new Error('Country of birth is required.');
  }

  YES_NO_RULE_DEFINITIONS.forEach((definition) => {
    const select = yesNoRules?.querySelector(`[data-yes-no-key="${definition.key}"] select`);
    yesNoAnswers[definition.key] = normalizeYesNoAnswer(select?.value, definition.expected);
  });

  return {
    ageCutoffDate,
    workingHoursReviewThreshold,
    workingHoursFailThreshold,
    countryOfResidence,
    nationality,
    countryOfBirth,
    yesNoAnswers,
  };
}

function renderRulesSummary() {
  if (!rulesSummary) {
    return;
  }

  const summaryCards = [
    {
      label: 'Age rule',
      value: formatRuleDate(ruleState.ageCutoffDate),
      note: 'Applicants must be 18 or older on this date.',
    },
    {
      label: 'Weekly hours',
      value: `Review below ${ruleState.workingHoursReviewThreshold}`,
      note: `Fail above ${ruleState.workingHoursFailThreshold}.`,
    },
    {
      label: 'Residence',
      value: ruleState.countryOfResidence,
      note: 'Country matching is normalized before comparison.',
    },
    {
      label: 'Nationality',
      value: ruleState.nationality,
      note: 'Nationality is compared as a normalized country value.',
    },
    {
      label: 'Country of birth',
      value: ruleState.countryOfBirth,
      note: 'Country of birth is compared as a normalized country value.',
    },
    {
      label: 'Yes/No presets',
      value: `${YES_NO_RULE_DEFINITIONS.length} editable fields`,
      note: 'Each preset keeps the eligible answer ready in the shared config.',
    },
    {
      label: 'Shared source',
      value: RULE_CONFIG_PATH,
      note: 'Update this file and republish the app to change the rules for everyone.',
    },
  ];

  rulesSummary.innerHTML = summaryCards
    .map(
      (card) => `
        <article class="rule-summary-card">
          <span class="rule-summary-card__label">${escapeHtml(card.label)}</span>
          <strong class="rule-summary-card__value">${escapeHtml(card.value)}</strong>
          <p class="rule-summary-card__note">${escapeHtml(card.note)}</p>
        </article>
      `,
    )
    .join('');
}

function renderYesNoRules() {
  if (!yesNoRules) {
    return;
  }

  yesNoRules.innerHTML = YES_NO_RULE_DEFINITIONS.map((definition) => `
    <div class="rule-field rule-field--inline" data-yes-no-key="${definition.key}">
      <label for="${definition.key}Input">${escapeHtml(definition.label)}</label>
      <select id="${definition.key}Input">
        <option value="yes">Yes</option>
        <option value="no">No</option>
      </select>
    </div>
  `).join('');
}

function formatRuleDate(value) {
  const date = new Date(`${value}T00:00:00Z`);

  if (Number.isNaN(date.getTime())) {
    return value;
  }

  return new Intl.DateTimeFormat(undefined, {
    dateStyle: 'medium',
  }).format(date);
}

function handleApplicantListClick(event) {
  const button = event.target.closest('[data-applicant-index]');

  if (!button) {
    return;
  }

  const index = Number(button.dataset.applicantIndex);

  if (Number.isNaN(index)) {
    return;
  }

  appState.selectedApplicantIndex = index;
  renderApplicantDetails();
  renderApplicantList();
}

function renderDashboard(applicants) {
  const metrics = applicants.reduce(
    (summary, applicant) => {
      summary.total += 1;
      summary[applicant.statusKey] += 1;
      return summary;
    },
    {
      total: 0,
      pass: 0,
      fail: 0,
      review: 0,
      missing: 0,
    },
  );

  totalApplicantsCount.textContent = String(metrics.total);
  eligibleCount.textContent = String(metrics.pass);
  ineligibleCount.textContent = String(metrics.fail);
  reviewCount.textContent = String(metrics.review);
  missingDataCount.textContent = String(metrics.missing);

  summaryOutput.value = '';
  renderApplicantList();
  renderApplicantDetails();
}

function renderApplicantList() {
  if (appState.applicants.length === 0) {
    applicantList.innerHTML = '<div class="empty-state-panel">Upload a file to see applicants listed here.</div>';
    return;
  }

  applicantList.innerHTML = appState.applicants
    .map(
      (applicant, index) => `
        <button
          type="button"
          class="applicant-item applicant-item--${applicant.statusKey}${index === appState.selectedApplicantIndex ? ' is-selected' : ''}"
          data-applicant-index="${index}"
        >
          <span class="applicant-item__name">${escapeHtml(applicant.name)}</span>
          <span class="status-pill status-pill--${applicant.statusKey}">${escapeHtml(applicant.statusLabel)}</span>
        </button>
      `,
    )
    .join('');
}

function renderApplicantDetails() {
  const applicant = appState.applicants[appState.selectedApplicantIndex];

  if (!applicant) {
    applicantDetails.innerHTML = '<div class="empty-state-panel">Select an applicant to inspect the criteria breakdown.</div>';
    return;
  }

  const criteriaMarkup = applicant.criteria
    .map(
      (criterion) => `
        <article class="criterion-card criterion-card--${criterion.status}">
          <div class="criterion-card__top">
            <span class="criterion-card__label">${escapeHtml(criterion.label)}</span>
            <span class="status-pill status-pill--${criterion.status}">${escapeHtml(criterion.statusLabel)}</span>
          </div>
          <p class="criterion-card__value">${escapeHtml(criterion.value)}</p>
          <p class="criterion-card__note">${escapeHtml(criterion.note)}</p>
        </article>
      `,
    )
    .join('');

  applicantDetails.innerHTML = `
    <div class="detail-summary detail-summary--${applicant.statusKey}">
      <div>
        <p class="detail-name">${escapeHtml(applicant.name)}</p>
        <p class="detail-subtitle">${escapeHtml(applicant.subTitle)}</p>
      </div>
      <span class="status-pill status-pill--${applicant.statusKey}">${escapeHtml(applicant.statusLabel)}</span>
    </div>
    <div class="criteria-grid">
      ${criteriaMarkup}
    </div>
  `;
}

function evaluateApplicant(headers, row, rowIndex) {
  const getValue = (headerName) => {
    const headerIndex = headers.findIndex((header) => header.toLowerCase() === headerName.toLowerCase());

    if (headerIndex === -1) {
      return '';
    }

    return String(row[headerIndex] ?? '').trim();
  };

  const firstName = getValue('First name');
  const lastName = getValue('Last name');
  const fallbackName = getValue('Email') || getValue('Id') || `Applicant ${rowIndex + 1}`;
  const applicantName = [firstName, lastName].filter(Boolean).join(' ') || fallbackName;
  const criteria = [];

  const addCriterion = (label, value, status, note, statusLabel = status.toUpperCase()) => {
    criteria.push({
      label,
      value: value || 'Not provided',
      status,
      statusLabel,
      note,
    });
  };

  const dob = getValue('DOB');
  if (!dob) {
    addCriterion('DOB', 'Not provided', 'missing', 'Date of birth is required.');
  } else {
    const parsedDob = parseFlexibleDate(dob);
    if (!parsedDob) {
      addCriterion('DOB', dob, 'fail', 'Date format is not recognised.');
    } else {
      const ageAtCutoff = getAgeOnDate(parsedDob, getAgeCutoffDate());
      const meetsAge = ageAtCutoff >= 18;
      addCriterion(
        'Minimum Age Requirement Met?',
        `${meetsAge ? 'Yes' : 'No'}, they will be ${ageAtCutoff} at the start of the apprenticeship.`,
        meetsAge ? 'pass' : 'fail',
        meetsAge ? 'Age requirement met.' : `Applicant will be under 18 on ${formatRuleDate(ruleState.ageCutoffDate)}.`,
      );
    }
  }

  addYesNoCriterion(criteria, 'Have you used a previous name?', getValue('Have you used a previous name?'), {
    yes: { status: 'review', note: 'Previous name declared and needs checking.' },
    no: { status: 'pass', note: 'No previous name declared.' },
    missing: { status: 'missing', note: 'Previous name response is required.' },
  });

  const workingHours = getValue('What will be your contracted weekly working hours?');
  if (!workingHours) {
    addCriterion('Working Hours', 'Not provided', 'missing', 'Weekly working hours are required.');
  } else {
    const hours = parseWeeklyHours(workingHours);
    if (hours == null) {
      addCriterion('Working Hours', workingHours, 'fail', 'Working hours could not be interpreted as a number.');
    } else if (hours < ruleState.workingHoursReviewThreshold) {
      addCriterion('Working Hours', String(hours), 'review', `Part-time hours below ${ruleState.workingHoursReviewThreshold} need review.`);
    } else if (hours > ruleState.workingHoursFailThreshold) {
      addCriterion('Working Hours', String(hours), 'fail', `Weekly hours exceed the fail threshold of ${ruleState.workingHoursFailThreshold}.`);
    } else {
      addCriterion('Working Hours', String(hours), 'pass', 'Working hours are within the expected range.');
    }
  }

  addYesNoCriterion(criteria, 'In employment, including self-employment', getValue('In employment, including self-employment'), {
    yes: { status: 'pass', note: 'Applicant is in employment.' },
    no: { status: 'fail', note: 'Applicant must be in employment.' },
    missing: { status: 'missing', note: 'Employment status is required.' },
  });

  addYesNoCriterion(criteria, 'UK/EEA National', getValue('UK/EEA National'), {
    yes: { status: 'pass', note: 'UK/EEA national declared.' },
    no: { status: 'review', note: 'Applicant is not marked as UK/EEA national and needs checking.' },
    missing: { status: 'missing', note: 'Nationality status is required.' },
  });

  addYesNoCriterion(criteria, 'Will you undertake more than 50% of your apprenticeship role within England?', getValue('Will you undertake more than 50% of your apprenticeship role within England?'), {
    yes: { status: 'pass', note: 'Role is mainly within England.' },
    no: { status: 'fail', note: 'Applicant must undertake more than 50% of the apprenticeship role within England.' },
    missing: { status: 'missing', note: 'England workplace check is required.' },
  });

  const countryOfResidence = getValue('Country of residence');
  if (!countryOfResidence) {
    addCriterion('Country of residence', 'Not provided', 'missing', 'Country of residence is required.');
  } else if (normalizeCountry(countryOfResidence) !== normalizeCountry(ruleState.countryOfResidence)) {
    addCriterion('Country of residence', countryOfResidence, 'fail', `Country of residence must be ${ruleState.countryOfResidence}.`);
  } else {
    addCriterion('Country of residence', countryOfResidence, 'pass', 'Country of residence matches the requirement.');
  }

  const nationality = getValue('Nationality');
  if (!nationality) {
    addCriterion('Nationality', 'Not provided', 'missing', 'Nationality is required.');
  } else if (normalizeCountry(nationality) !== 'unitedkingdom') {
    addCriterion('Nationality', nationality, 'review', 'Nationality is not UnitedKingdom and needs checking.');
  } else {
    addCriterion('Nationality', nationality, 'pass', 'Nationality matches the expected value.');
  }

  const countryOfBirth = getValue('Country of birth');
  if (!countryOfBirth) {
    addCriterion('Country of birth', 'Not provided', 'missing', 'Country of birth is required.');
  } else if (normalizeCountry(countryOfBirth) !== 'unitedkingdom') {
    addCriterion('Country of birth', countryOfBirth, 'review', 'Country of birth is not UnitedKingdom and needs checking.');
  } else {
    addCriterion('Country of birth', countryOfBirth, 'pass', 'Country of birth matches the expected value.');
  }

  addYesNoCriterion(criteria, 'Resident in the UK/EEA for 3 years', getValue('Resident in the UK/EEA for 3 years'), {
    yes: { status: 'pass', note: 'Residency requirement met.' },
    no: { status: 'fail', note: 'Applicant must be resident in the UK/EEA for 3 years.' },
    missing: { status: 'missing', note: 'Residency declaration is required.' },
  });

  addYesNoCriterion(criteria, 'Requires a Work Permit', getValue('Requires a Work Permit'), {
    yes: { status: 'review', note: 'Applicant requires a work permit and needs checking.' },
    no: { status: 'pass', note: 'No work permit required.' },
    missing: { status: 'missing', note: 'Work permit status is required.' },
  });

  addYesNoCriterion(criteria, 'Other government-funded training', getValue('In the last 12 months, have you undertaken, or are you planning to undertake, any other government-funded training (excluding apprenticeships)'), {
    yes: { status: 'review', note: 'Other government-funded training declared and needs further checks.' },
    no: { status: 'pass', note: 'No other government-funded training declared.' },
    missing: { status: 'missing', note: 'Training history is required.' },
  }, 'In the last 12 months, have you undertaken, or are you planning to undertake, any other government-funded training (excluding apprenticeships)');

  const evidence = getValue('Details of evidence presented');
  if (!evidence) {
    addCriterion('Details of evidence presented', 'Not provided', 'missing', 'Evidence details need to be completed.');
  } else {
    addCriterion('Details of evidence presented', evidence, 'review', 'Evidence details should be checked in Aptem.');
  }

  addYesNoCriterion(criteria, 'Contracted for full duration', getValue('Will you be contracted for the full duration of your apprenticeship, including the End-Point Assessment?'), {
    yes: { status: 'pass', note: 'Contract covers the full apprenticeship duration.' },
    no: { status: 'fail', note: 'Applicant must be contracted for the full duration of the apprenticeship.' },
    missing: { status: 'missing', note: 'Contract duration response is required.' },
  }, 'Will you be contracted for the full duration of your apprenticeship, including the End-Point Assessment?');

  addYesNoCriterion(criteria, 'Paid minimum wage', getValue('Will you be paid (at least) the apprenticeship minimum wage for the duration of the apprenticeship?'), {
    yes: { status: 'pass', note: 'Apprenticeship minimum wage will be paid.' },
    no: { status: 'fail', note: 'Applicant must be paid at least the apprenticeship minimum wage.' },
    missing: { status: 'missing', note: 'Pay response is required.' },
  }, 'Will you be paid (at least) the apprenticeship minimum wage for the duration of the apprenticeship?');

  const qualifications = getValue('Please list the full titles of the qualification(s) you will be using to meet the entry requirements of the apprenticeship, including their level and grade (e.g., Level 3 qualifications that hold UCAS points, a degree certificate).');
  if (!qualifications) {
    addCriterion('Qualifications used to meet entry requirements', 'Not provided', 'missing', 'Qualification titles need to be completed.');
  } else {
    addCriterion('Qualifications used to meet entry requirements', qualifications, 'review', 'Qualification details should be checked in Aptem.');
  }

  const additionalQualifications = getValue('If you hold any additional professional qualifications please list them here (e.g. role specific training, CPD).');
  if (additionalQualifications) {
    addCriterion('Additional professional qualifications', additionalQualifications, 'review', 'Additional professional qualifications should be checked in Aptem.');
  } else {
    addCriterion('Additional professional qualifications', 'Not provided', 'pass', 'No additional professional qualifications declared.');
  }

  addYesNoCriterion(criteria, 'Permanent contract', getValue('Do you have a permanent contract?'), {
    yes: { status: 'pass', note: 'Permanent contract declared.' },
    no: { status: 'review', note: 'Applicant does not have a permanent contract and needs review.' },
    missing: { status: 'missing', note: 'Permanent contract response is required.' },
  });

  addYesNoCriterion(criteria, 'Another apprenticeship', getValue('Apart from the apprenticeship you are currently applying for right now, are you enrolled on any another apprenticeship?'), {
    yes: { status: 'pass', note: 'Applicant is enrolled on another apprenticeship.' },
    no: { status: 'review', note: 'Further End Point Assessment details are required.' },
    missing: { status: 'missing', note: 'Another apprenticeship response is required.' },
  });

  addYesNoCriterion(criteria, 'Leeds Beckett previous student', getValue('Have you previously applied or studied with Leeds Beckett University?'), {
    yes: { status: 'review', note: 'Previous Leeds Beckett student declaration needs checking.' },
    no: { status: 'pass', note: 'No previous Leeds Beckett study declared.' },
    missing: { status: 'missing', note: 'Previous Leeds Beckett study response is required.' },
  });

  const studentNumber = getValue('If yes, please give your Leeds Beckett Student Number if known.');
  if (studentNumber) {
    addCriterion('Leeds Beckett Student Number', studentNumber, 'review', 'Student number is present and should be checked.');
  } else {
    addCriterion('Leeds Beckett Student Number', 'Not provided', 'pass', 'No student number was supplied.');
  }

  const niNumber = getValue('National insurance number');
  if (!niNumber) {
    addCriterion('National insurance number', 'Not provided', 'missing', 'National insurance number is required.');
  } else {
    addCriterion('National insurance number', niNumber, 'pass', 'National insurance number is present.');
  }

  const hasMissing = criteria.some((criterion) => criterion.status === 'missing');
  const hasFail = criteria.some((criterion) => criterion.status === 'fail');
  const hasReview = criteria.some((criterion) => criterion.status === 'review');
  const statusKey = hasMissing ? 'missing' : hasFail ? 'fail' : hasReview ? 'review' : 'pass';

  return {
    name: applicantName,
    subTitle: `Row ${rowIndex + 2}`,
    statusKey,
    statusLabel: statusKey === 'missing' ? 'MISSING DATA' : statusKey === 'fail' ? 'FAIL' : statusKey === 'review' ? 'REVIEW' : 'PASS',
    criteria,
  };
}

function addYesNoCriterion(criteria, label, value, outcomes, headerLabel = label) {
  const normalized = String(value ?? '').trim().toLowerCase();

  if (!normalized) {
    const missing = outcomes.missing || { status: 'missing', note: 'Response required.' };
    criteria.push({
      label: headerLabel,
      value: 'Not provided',
      status: missing.status,
      statusLabel: missing.status.toUpperCase(),
      note: missing.note,
    });
    return;
  }

  if (/^(yes|y|true)$/.test(normalized)) {
    const yes = outcomes.yes || { status: 'pass', note: 'Pass.' };
    criteria.push({
      label: headerLabel,
      value,
      status: yes.status,
      statusLabel: yes.status.toUpperCase(),
      note: yes.note,
    });
    return;
  }

  if (/^(no|n|false)$/.test(normalized)) {
    const no = outcomes.no || { status: 'review', note: 'Review required.' };
    criteria.push({
      label: headerLabel,
      value,
      status: no.status,
      statusLabel: no.status.toUpperCase(),
      note: no.note,
    });
    return;
  }

  criteria.push({
    label: headerLabel,
    value,
    status: 'missing',
    statusLabel: 'MISSING',
    note: 'Response must be Yes or No.',
  });
}

function getAgeOnDate(date, referenceDate) {
  const yearDifference = referenceDate.getFullYear() - date.getFullYear();
  const monthDifference = referenceDate.getMonth() - date.getMonth();
  const dayDifference = referenceDate.getDate() - date.getDate();

  let age = yearDifference;

  if (monthDifference < 0 || (monthDifference === 0 && dayDifference < 0)) {
    age -= 1;
  }

  return age;
}

function getAgeCutoffDate() {
  return new Date(`${ruleState.ageCutoffDate}T00:00:00Z`);
}

function parseWeeklyHours(value) {
  const text = String(value ?? '').trim();

  if (!text) {
    return null;
  }

  const normalized = text.replace(',', '.');
  const match = normalized.match(/\d+(?:\.\d+)?/);

  if (!match) {
    return null;
  }

  const hours = Number(match[0]);

  return Number.isNaN(hours) ? null : hours;
}

function normalizeCountry(value) {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z]/g, '');
}

function parseFlexibleDate(value) {
  const text = String(value ?? '').trim();

  if (!text) {
    return null;
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    const date = new Date(text);
    return Number.isNaN(date.getTime()) ? null : date;
  }

  if (/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(text)) {
    const parts = text.split('/');
    const first = Number(parts[0]);
    const second = Number(parts[1]);
    const year = normalizeYearPart(parts[2]);

    if (year == null) {
      return null;
    }

    const ukDate = buildDate(year, second, first);
    if (ukDate) {
      return ukDate;
    }

    return buildDate(year, first, second);
  }

  return null;
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
function buildDate(year, month, day) {
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

  return date;
}

function setMessage(text, isError = false) {
  message.textContent = text;
  message.classList.toggle('error', isError);
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

