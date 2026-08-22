const express = require('express');
const cors = require('cors');
const path = require('path');
const { google } = require('googleapis');
const ExcelJS = require('exceljs');
require('dotenv').config();

const app = express();
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());
app.use(cors());  

const PORT = process.env.PORT || 3000;
const liveFleetTracker = {};

async function getGoogleSheetsClient() {
  let auth;
  if (process.env.GOOGLE_CREDENTIALS_JSON) {
    const credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON.trim());
    auth = new google.auth.GoogleAuth({ credentials, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
  } else {
    auth = new google.auth.GoogleAuth({ keyFile: 'credentials.json', scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
  }
  const client = await auth.getClient();
  return google.sheets({ version: 'v4', auth: client });
}

async function ensureDriversSheetExists(googleSheets, spreadsheetId) {
  const spreadsheet = await googleSheets.spreadsheets.get({ spreadsheetId });
  const sheetExists = spreadsheet.data.sheets.some(s => s.properties.title === 'Drivers');

  if (!sheetExists) {
    await googleSheets.spreadsheets.batchUpdate({
      spreadsheetId,
      resource: { requests: [{ addSheet: { properties: { title: 'Drivers' } } }] }
    });
    await googleSheets.spreadsheets.values.update({
      spreadsheetId,
      range: 'Drivers!A1:D1',
      valueInputOption: 'USER_ENTERED',
      resource: { values: [['Phone Number', 'Driver Name', 'Vehicle ID', 'Status']] }
    });
  }
}

function calculateDistance(lat1, lon1, lat2, lon2) {
  const R = 6371;
  const dLat = (lat2 - lat1) * (Math.PI / 180);
  const dLon = (lon2 - lon1) * (Math.PI / 180);
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) + Math.cos(lat1 * (Math.PI / 180)) * Math.cos(lat2 * (Math.PI / 180)) * Math.sin(dLon / 2) * Math.sin(dLon / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

async function appendToGoogleSheet(tripData) {
  try {
    const googleSheets = await getGoogleSheetsClient();
    const spreadsheetId = process.env.SPREADSHEET_ID;

    const date = new Date();
    const sheetName = date.toLocaleString('default', { month: 'long', year: 'numeric' });

    const spreadsheet = await googleSheets.spreadsheets.get({ spreadsheetId });
    const sheetExists = spreadsheet.data.sheets.some(s => s.properties.title === sheetName);

    if (!sheetExists) {
      await googleSheets.spreadsheets.batchUpdate({
        spreadsheetId,
        resource: { requests: [{ addSheet: { properties: { title: sheetName } } }] }
      });
      await googleSheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${sheetName}!A1:K1`,
        valueInputOption: 'USER_ENTERED',
        resource: { values: [['Start Time', 'End Time', 'Vehicle ID', 'Driver Name', 'Customer Name', 'Start Odo', 'End Odo', 'Manual Dist', 'Start GPS', 'End GPS', 'GPS Dist (km)']] }
      });
    }

    let manualDistance = '';
    if (tripData.start_odometer && tripData.end_odometer) {
      manualDistance = parseFloat(tripData.end_odometer) - parseFloat(tripData.start_odometer);
    }

    let gpsDistanceKm = '';
    if (tripData.start_gps && tripData.end_gps) {
      const [startLat, startLon] = tripData.start_gps.split(',').map(Number);
      const [endLat, endLon] = tripData.end_gps.split(',').map(Number);
      if (!isNaN(startLat) && !isNaN(startLon) && !isNaN(endLat) && !isNaN(endLon)) {
        gpsDistanceKm = calculateDistance(startLat, startLon, endLat, endLon).toFixed(2);
      }
    }

    const finalDriverName = (tripData.driver_name || tripData.root_driver_name || '').trim().toUpperCase();

    const rowData = [
      tripData.start_timestamp || '',
      tripData.timestamp || new Date().toLocaleString(),
      tripData.vehicle_id ? tripData.vehicle_id.trim().toUpperCase() : '',
      finalDriverName,
      tripData.customer_name ? tripData.customer_name.trim().toUpperCase() : '',
      tripData.start_odometer || '',
      tripData.end_odometer || '',
      manualDistance,
      tripData.start_gps || '',
      tripData.end_gps || '',
      gpsDistanceKm
    ];

    await googleSheets.spreadsheets.values.append({
      spreadsheetId,
      range: `${sheetName}!A:K`,
      valueInputOption: 'USER_ENTERED',
      resource: { values: [rowData] },
    });
  } catch (error) {
    console.error('Error writing to Google Sheet:', error);
  }
}

app.post('/api/drivers/register', async (req, res) => {
  try {
    const { driver_name, phone_number, vehicle_id } = req.body;
    const normalizedName = driver_name.trim().toUpperCase();
    const normalizedPhone = phone_number.trim();
    const normalizedVehicle = vehicle_id.trim().toUpperCase();

    const googleSheets = await getGoogleSheetsClient();
    const spreadsheetId = process.env.SPREADSHEET_ID;
    await ensureDriversSheetExists(googleSheets, spreadsheetId);

    const response = await googleSheets.spreadsheets.values.get({ spreadsheetId, range: 'Drivers!A:D' });
    const rows = response.data.values || [];
    let rowIndex = -1;
    let currentStatus = 'pending';

    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][0] || '').trim() === normalizedPhone) {
        rowIndex = i + 1;
        const oldVehicle = (rows[i][2] || '').trim().toUpperCase();
        let existingStatus = (rows[i][3] || 'pending').trim().toLowerCase();

        if (oldVehicle === normalizedVehicle) {
          currentStatus = existingStatus; 
        } else {
          currentStatus = 'pending';
          await googleSheets.spreadsheets.values.update({
            spreadsheetId,
            range: `Drivers!B${rowIndex}:D${rowIndex}`,
            valueInputOption: 'USER_ENTERED',
            resource: { values: [[normalizedName, normalizedVehicle, 'pending']] },
          });
        }
        break;
      }
    }

    if (rowIndex === -1) {
      await googleSheets.spreadsheets.values.append({
        spreadsheetId,
        range: 'Drivers!A:D',
        valueInputOption: 'USER_ENTERED',
        resource: { values: [[normalizedPhone, normalizedName, normalizedVehicle, 'pending']] },
      });
      currentStatus = 'pending';
    }
    return res.status(200).json({ status: currentStatus });
  } catch (error) {
    return res.status(500).json({ status: 'error' });
  }
});

app.get('/api/admin/export-excel/:monthYear', async (req, res) => {
  try {
    const sheetName = req.params.monthYear;
    const googleSheets = await getGoogleSheetsClient();
    const spreadsheetId = process.env.SPREADSHEET_ID;

    const response = await googleSheets.spreadsheets.values.get({ spreadsheetId, range: `${sheetName}!A:K` });
    const rows = response.data.values || [];
    const workbook = new ExcelJS.Workbook();

    const vehicleGroups = {};
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      const vehicleId = (r[2] || 'UNKNOWN').trim();
      if (!vehicleGroups[vehicleId]) vehicleGroups[vehicleId] = [];
      vehicleGroups[vehicleId].push({
        date: r[0]?.split(' ')[0] || '',
        start: r[0]?.split(' ')[1] || '',
        end: r[1]?.split(' ')[1] || '',
        startOdo: r[5] || 0,
        endOdo: r[6] || 0,
        dist: parseFloat(r[7]) || 0
      });
    }

    for (const [vehicleId, trips] of Object.entries(vehicleGroups)) {
      const ws = workbook.addWorksheet(vehicleId);
      ws.mergeCells('A1:H1');
      ws.getCell('A1').value = `VEHICLE NUMBER (${vehicleId}) FOR THE MONTH OF ${sheetName.toUpperCase()}`;
      ws.getCell('A1').font = { bold: true, size: 12 };
      ws.getCell('A1').alignment = { horizontal: 'center' };

      ws.addRow(['DATE', 'STARTING TIME', 'CLOSING TIME', 'TOTAL HOURS', 'EXTRA HOURS', 'STARTING Kms', 'CLOSING Kms', 'TOTAL Kms']);
      let totalKms = 0;
      trips.forEach(t => {
        ws.addRow([t.date, t.start, t.end, 12, 0, t.startOdo, t.endOdo, t.dist]);
        totalKms += t.dist;
      });
      const totalRow = ws.addRow(['TOTAL', '', '', '', '', '', '', totalKms]);
      totalRow.font = { bold: true };
      ws.columns = [{width: 15}, {width: 15}, {width: 15}, {width: 15}, {width: 15}, {width: 15}, {width: 15}, {width: 15}];
    }

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=Billing_Report_${sheetName.replace(' ', '_')}.xlsx`);
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    res.status(500).send('Error generating report: ' + error.message);
  }
});

app.post('/api/trips/sync', async (req, res) => {
  try {
    const { vehicle_id, driver_name, trips } = req.body;
    for (const trip of trips) {
      await appendToGoogleSheet({
        vehicle_id, driver_name: trip.driver_name || driver_name,
        customer_name: trip.customer_name, start_odometer: trip.start_odometer,
        end_odometer: trip.end_odometer, start_gps: trip.start_gps,
        end_gps: trip.end_gps, start_timestamp: trip.start_timestamp, timestamp: trip.timestamp
      });
    }
    res.status(200).json({ status: 'success' });
  } catch (error) {
    res.status(500).json({ error: 'Internal Server Error' });
  }
});

app.post('/api/fleet/location', (req, res) => {
  const { vehicle_id, latitude, longitude, speed } = req.body;
  if (vehicle_id) liveFleetTracker[vehicle_id.trim().toUpperCase()] = { latitude, longitude, speed };
  res.status(200).json({ status: 'success' });
});

app.listen(PORT, '0.0.0.0', () => console.log(`--- SERVER RUNNING ON PORT ${PORT} ---`));