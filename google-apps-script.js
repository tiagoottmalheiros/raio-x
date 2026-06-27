const SHEET_NAME = 'Respostas';

const COLUMNS = [
  'submittedAt',
  'name',
  'email',
  'businessName',
  'totalScore',
  'resultName',
  'leakLevel',
  'mainBottleneckKey',
  'mainBottleneckLabel',
  'diagnosis',
  'symptoms',
  'cost',
  'nextStep',
  'impactPhrase',
  'generatedPrompt',
  'answers'
];

for (let i = 1; i <= 12; i++) {
  COLUMNS.push(
    `q${i}Id`,
    `q${i}Category`,
    `q${i}CategoryLabel`,
    `q${i}Title`,
    `q${i}OptionIndex`,
    `q${i}OptionValue`,
    `q${i}AnswerText`,
    `q${i}Score`
  );
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    const sheet = getSheet_();
    ensureHeaders_(sheet);

    const row = COLUMNS.map((column) => formatValue_(payload[column]));
    sheet.appendRow(row);

    return json_({ success: true });
  } catch (error) {
    return json_({ success: false, error: String(error) });
  }
}

function getSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
}

function ensureHeaders_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), COLUMNS.length);
  const currentHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const needsUpdate = COLUMNS.some((column, index) => currentHeaders[index] !== column);

  if (needsUpdate) {
    sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
    sheet.setFrozenRows(1);
  }
}

function formatValue_(value) {
  if (value === undefined || value === null) return '';
  if (Array.isArray(value) || typeof value === 'object') return JSON.stringify(value);
  return value;
}

function json_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
