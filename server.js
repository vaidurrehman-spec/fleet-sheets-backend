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

// Automatically ensures the 'Drivers' sheet and header row exist
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
  const a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(lat1 * (Math.PI / 180)) * Math.cos(lat2 * (Math.PI / 180)) *
    Math.sin(dLon / 2) * Math.sin(dLon / 2);
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
    if (!driver_name || !phone_number || !vehicle_id) {
      return res.status(400).json({ status: 'error', message: 'Missing required fields.' });
    }

    const normalizedName = driver_name.trim().toUpperCase();
    const normalizedPhone = phone_number.trim();
    const normalizedVehicle = vehicle_id.trim().toUpperCase();

    if (normalizedName === 'INPUT TEXT') {
      return res.status(400).json({ status: 'error', message: 'Invalid placeholder name.' });
    }

    const googleSheets = await getGoogleSheetsClient();
    const spreadsheetId = process.env.SPREADSHEET_ID;

    // Automatically create tab and headers if missing
    await ensureDriversSheetExists(googleSheets, spreadsheetId);

    const response = await googleSheets.spreadsheets.values.get({ spreadsheetId, range: 'Drivers!A:D' });
    const rows = response.data.values || [];
    let rowIndex = -1;
    let currentStatus = 'pending';

    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][0] || '').trim() === normalizedPhone) {
        rowIndex = i + 1;
        const oldVehicle = (rows[i][2] || '').trim().toUpperCase();
        currentStatus = (rows[i][3] || 'pending').trim().toLowerCase();

        if (oldVehicle !== normalizedVehicle) {
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

    return res.status(200).json({ status: currentStatus, message: `Current status: ${currentStatus}` });
  } catch (error) {
    console.error('Error in driver registration:', error);
    return res.status(500).json({ status: 'error', message: 'Internal server error' });
  }
});

app.get('/api/admin/drivers', async (req, res) => {
  try {
    const googleSheets = await getGoogleSheetsClient();
    const spreadsheetId = process.env.SPREADSHEET_ID;
    
    await ensureDriversSheetExists(googleSheets, spreadsheetId);

    const response = await googleSheets.spreadsheets.values.get({ spreadsheetId, range: 'Drivers!A:D' });
    const rows = response.data.values || [];
    
    const drivers = rows.slice(1).map(r => ({
      phone_number: (r[0] || '').trim(),
      driver_name: (r[1] || '').trim().toUpperCase(),
      vehicle_id: (r[2] || '').trim().toUpperCase(),
      status: (r[3] || 'pending').trim().toLowerCase(),
    }));

    return res.status(200).json({ status: 'success', drivers });
  } catch (error) {
    return res.status(500).json({ error: 'Failed to fetch drivers' });
  }
});

app.post('/api/admin/drivers/update-status', async (req, res) => {
  try {
    const { phone_number, status } = req.body;
    if (!phone_number || !status) {
      return res.status(400).json({ status: 'error', message: 'Phone number and status required.' });
    }

    const googleSheets = await getGoogleSheetsClient();
    const spreadsheetId = process.env.SPREADSHEET_ID;
    
    await ensureDriversSheetExists(googleSheets, spreadsheetId);

    const response = await googleSheets.spreadsheets.values.get({ spreadsheetId, range: 'Drivers!A:D' });
    const rows = response.data.values || [];

    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][0] || '').trim() === phone_number.trim()) {
        const rowIndex = i + 1;
        await googleSheets.spreadsheets.values.update({
          spreadsheetId,
          range: `Drivers!D${rowIndex}`,
          valueInputOption: 'USER_ENTERED',
          resource: { values: [[status.trim().toLowerCase()]] },
        });
        return res.status(200).json({ status: 'success', message: `Driver status updated to ${status}` });
      }
    }
    return res.status(404).json({ status: 'error', message: 'Driver not found.' });
  } catch (error) {
    return res.status(500).json({ error: 'Failed to update status' });
  }
});

app.get('/api/admin/driver-details/:driverName', async (req, res) => {
  try {
    const driverName = req.params.driverName.trim().toUpperCase();
    const phoneParam = (req.query.phone || '').trim();

    const googleSheets = await getGoogleSheetsClient();
    const spreadsheetId = process.env.SPREADSHEET_ID;
    
    await ensureDriversSheetExists(googleSheets, spreadsheetId);

    const response = await googleSheets.spreadsheets.values.get({ spreadsheetId, range: 'Drivers!A:D' });
    const rows = response.data.values || [];

    let matchedRow = rows.slice(1).find(r => (r[0] || '').trim() === phoneParam || (r[1] || '').trim().toUpperCase() === driverName);
    let matchedVehicleId = matchedRow ? (matchedRow[2] || '').trim().toUpperCase() : null;
    let matchedPhone = matchedRow ? (matchedRow[0] || '').trim() : 'N/A';

    let locationInfo = { latitude: 'N/A', longitude: 'N/A', speed: 0, heading: 'Stationary', timestamp: 'No transmission' };
    if (matchedVehicleId && liveFleetTracker[matchedVehicleId]) {
      const loc = liveFleetTracker[matchedVehicleId];
      locationInfo = { latitude: loc.latitude, longitude: loc.longitude, speed: loc.speed, heading: 'Active on road', timestamp: loc.last_updated };
    }
    
    res.status(200).json({ 
      driver_name: driverName, 
      phone_number: matchedPhone,
      vehicle_id: matchedVehicleId || 'N/A',
      ...locationInfo 
    });
  } catch (error) {
    res.status(500).json({ error: 'Failed to fetch driver details' });
  }
});

app.get('/api/admin/billing-summary/:monthYear', async (req, res) => {
  try {
    const sheetName = req.params.monthYear;
    const googleSheets = await getGoogleSheetsClient();
    const spreadsheetId = process.env.SPREADSHEET_ID;

    const response = await googleSheets.spreadsheets.values.get({ spreadsheetId, range: `${sheetName}!A:K` });
    const rows = response.data.values;
    if (!rows || rows.length < 2) {
      return res.status(200).json({ month: sheetName, client_totals: [], driver_totals: [] });
    }

    const clientMap = {};
    const driverMap = {};

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const customerName = (row[4] || 'GENERAL CLIENT').trim().toUpperCase();
      const driverName = (row[3] || 'UNKNOWN DRIVER').trim().toUpperCase();
      const manualDist = parseFloat(row[7]) || 0;

      if (!clientMap[customerName]) clientMap[customerName] = { total_trips: 0, total_km: 0 };
      clientMap[customerName].total_trips += 1;
      clientMap[customerName].total_km += manualDist;

      if (!driverMap[driverName]) driverMap[driverName] = { total_trips: 0, total_km: 0 };
      driverMap[driverName].total_trips += 1;
      driverMap[driverName].total_km += manualDist;
    }

    const clientTotals = Object.keys(clientMap).map(clientName => ({
      client_name: clientName,
      total_trips: clientMap[clientName].total_trips,
      total_km: parseFloat(clientMap[clientName].total_km.toFixed(2))
    }));

    const driverTotals = Object.keys(driverMap).map(driverName => ({
      driver_name: driverName,
      total_trips: driverMap[driverName].total_trips,
      total_km: parseFloat(driverMap[driverName].total_km.toFixed(2))
    }));

    return res.status(200).json({ month: sheetName, client_totals: clientTotals, driver_totals: driverTotals });
  } catch (error) {
    return res.status(500).json({ error: 'Failed to calculate summary data' });
  }
});

app.get('/api/admin/driver-summary/:monthYear', async (req, res) => {
  req.url = `/api/admin/billing-summary/${req.params.monthYear}`;
  return app._router.handle(req, res);
});

app.get('/api/admin/export-excel/:monthYear', async (req, res) => {
  try {
    const sheetName = req.params.monthYear;
    const googleSheets = await getGoogleSheetsClient();
    const spreadsheetId = process.env.SPREADSHEET_ID;

    const response = await googleSheets.spreadsheets.values.get({ spreadsheetId, range: `${sheetName}!A:K` });
    const rows = response.data.values;
    const workbook = new ExcelJS.Workbook();
    const logSheet = workbook.addWorksheet('Monthly Fleet Logs');

    logSheet.columns = [
      { header: 'Start Time', key: 'start_timestamp', width: 20 },
      { header: 'End Time', key: 'timestamp', width: 20 },
      { header: 'Vehicle ID', key: 'vehicle_id', width: 15 },
      { header: 'Driver Name', key: 'driver_name', width: 20 },
      { header: 'Customer Name', key: 'customer_name', width: 25 },
      { header: 'Start Odo', key: 'start_odometer', width: 15 },
      { header: 'End Odo', key: 'end_odometer', width: 15 },
      { header: 'Manual Dist (km)', key: 'manual_dist', width: 18 },
      { header: 'Start GPS', key: 'start_gps', width: 22 },
      { header: 'End GPS', key: 'end_gps', width: 22 },
      { header: 'GPS Dist (km)', key: 'gps_dist', width: 15 },
    ];

    if (rows && rows.length > 1) {
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        logSheet.addRow({
          start_timestamp: r[0] || '', timestamp: r[1] || '', vehicle_id: r[2] || '', driver_name: r[3] || '', customer_name: r[4] || '',
          start_odometer: r[5] || '', end_odometer: r[6] || '', manual_dist: r[7] || '',
          start_gps: r[8] || '', end_gps: r[9] || '', gps_dist: r[10] || '',
        });
      }
    }

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=Fleet_Report_${sheetName.replace(' ', '_')}.xlsx`);
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    res.status(500).send('Error generating Excel file: ' + error.message);
  }
});

app.get('/api/drivers/history/:monthYear/:driverName', async (req, res) => {
  try {
    const { monthYear } = req.params;
    const driverName = req.params.driverName.trim().toUpperCase();
    const googleSheets = await getGoogleSheetsClient();
    const spreadsheetId = process.env.SPREADSHEET_ID;

    const response = await googleSheets.spreadsheets.values.get({ spreadsheetId, range: `${monthYear}!A:K` });
    const rows = response.data.values;
    if (!rows || rows.length < 2) return res.status(200).json({ driver: driverName, month: monthYear, trips: [] });

    const driverTrips = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if ((row[3] || '').trim().toUpperCase() === driverName) {
        driverTrips.push({ start_timestamp: row[0], timestamp: row[1], vehicle_id: row[2], customer_name: row[4], start_odometer: row[5], end_odometer: row[6], manual_dist: row[7], gps_dist: row[10] });
      }
    }
    return res.status(200).json({ driver: driverName, month: monthYear, trips: driverTrips });
  } catch (error) {
    res.status(500).json({ error: 'Failed to fetch driver history' });
  }
});

app.post('/api/trips/sync', async (req, res) => {
  try {
    const { vehicle_id, driver_name, trips } = req.body;
    for (const trip of trips) {
      await appendToGoogleSheet({
        vehicle_id,
        root_driver_name: driver_name,
        driver_name: trip.driver_name || driver_name || 'WEB USER',
        customer_name: trip.customer_name || 'GENERAL ROUTE',
        start_odometer: trip.start_odometer,
        end_odometer: trip.end_odometer,
        start_gps: trip.start_gps || '',
        end_gps: trip.end_gps || '',
        start_timestamp: trip.start_timestamp || '',
        timestamp: trip.timestamp
      });
    }
    res.status(200).json({ status: 'success', message: 'Synced successfully!' });
  } catch (error) {
    res.status(500).json({ error: 'Internal Server Error' });
  }
});

app.post('/api/fleet/location', (req, res) => {
  const { vehicle_id, latitude, longitude, speed, timestamp } = req.body;
  if (vehicle_id) {
    const normalizedVid = vehicle_id.trim().toUpperCase();
    liveFleetTracker[normalizedVid] = { latitude, longitude, speed: speed || 0, last_updated: timestamp || new Date().toISOString() };
  }
  res.status(200).json({ status: 'success' });
});

app.get('/api/fleet/live-status', (req, res) => {
  res.status(200).json({ total_active_vehicles: Object.keys(liveFleetTracker).length, fleet_data: liveFleetTracker });
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`--- SERVER RUNNING ON PORT ${PORT} ---`);
});