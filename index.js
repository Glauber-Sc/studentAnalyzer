console.log("Hello World");

const { google } = require('googleapis');
const fs = require('fs');

// Configuration
const SPREADSHEET_ID = '1iVNooq3kmfliABvtK6WXQ76zwkmrq2SqR-dLcD2b2FE';
const RANGE = 'Glauber!A4:H27'; // Defines the range to be updated

// Load the JSON credentials file
const credentials = require('./studentanalyzer-414402-b689a9819dc4.json');

// Authorization
const auth = new google.auth.JWT(
  credentials.client_email,
  null,
  credentials.private_key,
  ['https://www.googleapis.com/auth/spreadsheets']
);

// Google Sheets API initialization
const sheets = google.sheets({ version: 'v4', auth });

// Function to read data from the JSON file
function readData() {
  try {
    const data = fs.readFileSync('data.json', 'utf8');
    return JSON.parse(data);
  } catch (err) {
    console.error('Error reading data file:', err);
    return [];
  }
}

// Function to calculate the situation based on the average of the three tests (P1, P2, and P3)
function calculateSituation(student) {
  const { Faltas, P1, P2, P3 } = student;
  const totalClasses = 60;
  const attendanceThreshold = 0.25; // 25% attendance threshold

  const average = (P1 + P2 + P3) / 3;

  if (Faltas > totalClasses * attendanceThreshold) {
    return 'Reprovado por Falta';
  } else if (average < 5) {
    return 'Reprovado por Nota';
  } else if (average < 7) {
    return 'Exame Final';
  } else {
    return 'Aprovado';
  }
}

// Function to calculate the "Nota para Aprovacao Final" if the situation is "Exame Final"
function calculateNaf(student) {
  const { P1, P2, P3 } = student;
  const average = (P1 + P2 + P3) / 3;

  if (average >= 5) {
    const naf = 10 - average;
    return Math.ceil(naf);
  } else {
    return 0;
  }
}

// Function to update the sheet
async function updateSheet() {
  const newData = readData();
  const values = newData.map(student => [
    student.Matricula,
    student.Aluno,
    student.Faltas,
    student.P1,
    student.P2,
    student.P3,
    calculateSituation(student),
    calculateSituation(student) === 'Exame Final' ? calculateNaf(student) : 0
  ]);

  try {
    const result = await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: RANGE,
      valueInputOption: 'RAW',
      requestBody: {
        values: values
      }
    });
    console.log('%d cells updated.', result.data.updatedCells);
  } catch (err) {
    console.error('The API returned an error:', err);
  }
}

// Call the function to update the sheet
updateSheet();
