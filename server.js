const express = require('express');
const cors = require('cors');
const path = require('path');
const { google } = require('googleapis');
const ExcelJS = require('exceljs');
const bcrypt = require('bcrypt');
require('dotenv').config();

const app = express();
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());
app.use(cors());  

const PORT = process.env.PORT || 3000;
const liveFleetTracker = {};

// --- DATE NORMALIZER HELPER FOR EXCEL EXPORT ---
function normalizeDateKey(dateStr) {
  if (!dateStr) return '';
  const cleanDate = dateStr.split(' ')[0].trim();
  let parts = [];
  
  if (cleanDate.includes('/')) {
    parts = cleanDate.split('/');
  } else if (cleanDate.includes('-')) {
    parts = cleanDate.split('-');
  } else {
    return cleanDate;
  }

  if (parts.length === 3) {
    if (parts[0].length === 4) {
      return `${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0')}`;
    } else {
      return `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
    }
  }
  return cleanDate;
}

// --- SECURE DYNAMIC ADMIN AUTH MIDDLEWARE ---
async function verifyAdminAuth(req, res, next) {
  try {
    const authHeader = req.headers['authorization'];
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
      return res.status(401).json({ status: 'error', message: 'Unauthorized: No token provided.' });
    }

    const tokenPassword = authHeader.split(' ')[1];
    const adminSecret = process.env.ADMIN_SECRET_KEY || 'FALLBACK_SECRET_PASSWORD';

    if (tokenPassword === adminSecret) {
      return next();
    }

    const googleSheets = await getGoogleSheetsClient();
    const spreadsheetId = process.env.SPREADSHEET_ID;
    await ensureAdminUsersSheetExists(googleSheets, spreadsheetId);

    const response = await googleSheets.spreadsheets.values.get({ spreadsheetId, range: 'AdminUsers!A:D' });
    const rows = response.data.values || [];

    for (let i = 1; i < rows.length; i++) {
      const storedHash = rows[i][2];
      if (storedHash) {
        const match = await bcrypt.compare(tokenPassword, storedHash);
        if (match) {
          return next();
        }
      }
    }

    return res.status(401).json({ status: 'error', message: 'Unauthorized: Invalid admin credentials.' });
  } catch (error) {
    console.error('Auth middleware error:', error);
    return res.status(500).json({ status: 'error', message: 'Internal server error during authentication' });
  }
}

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
      range: 'Drivers!A1:E1',
      valueInputOption: 'USER_ENTERED',
      resource: { values: [['Phone Number', 'Driver Name', 'Vehicle ID', 'Status', 'Password']] }
    });
  }
}

async function ensureAdminUsersSheetExists(googleSheets, spreadsheetId) {
  const spreadsheet = await googleSheets.spreadsheets.get({ spreadsheetId });
  const sheetExists = spreadsheet.data.sheets.some(s => s.properties.title === 'AdminUsers');

  if (!sheetExists) {
    await googleSheets.spreadsheets.batchUpdate({
      spreadsheetId,
      resource: { requests: [{ addSheet: { properties: { title: 'AdminUsers' } } }] }
    });
    const defaultMasterHash = await bcrypt.hash('MasterPassword123', 10);
    await googleSheets.spreadsheets.values.update({
      spreadsheetId,
      range: 'AdminUsers!A1:D1',
      valueInputOption: 'USER_ENTERED',
      resource: { values: [['Name', 'Username', 'PasswordHash', 'Role']] }
    });
    await googleSheets.spreadsheets.values.append({
      spreadsheetId,
      range: 'AdminUsers!A:D',
      valueInputOption: 'USER_ENTERED',
      resource: { values: [['Super Admin', 'admin', defaultMasterHash, 'master']] }
    });
  }
}

async function ensureFuelSheetExists(googleSheets, spreadsheetId) {
  const spreadsheet = await googleSheets.spreadsheets.get({ spreadsheetId });
  const sheetExists = spreadsheet.data.sheets.some(s => s.properties.title === 'FuelLogs');

  if (!sheetExists) {
    await googleSheets.spreadsheets.batchUpdate({
      spreadsheetId,
      resource: { requests: [{ addSheet: { properties: { title: 'FuelLogs' } } }] }
    });
    await googleSheets.spreadsheets.values.update({
      spreadsheetId,
      range: 'FuelLogs!A1:K1',
      valueInputOption: 'USER_ENTERED',
      resource: { values: [['Timestamp', 'Vehicle ID', 'Driver Name', 'Fuel Type', 'Shift Time', 'Quantity', 'Unit', 'Total Cost (₹)', 'Per Unit Rate (₹)', 'Odometer', 'Station Location']] }
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

// --- HEALTH & AUTH CHECK ENDPOINTS ---
app.get('/api/health', (req, res) => {
  res.status(200).json({ status: 'success', message: 'Fleet server is running.' });
});

app.post('/api/admin/login', (req, res) => {
  const { password } = req.body;
  const adminSecret = process.env.ADMIN_SECRET_KEY || 'FALLBACK_SECRET_PASSWORD';
  if (password === adminSecret) {
    return res.status(200).json({ status: 'success', message: 'Authenticated successfully' });
  }
  return res.status(401).json({ status: 'error', message: 'Invalid admin password' });
});

app.post('/api/admin/multi-login', async (req, res) => {
  try {
    const { username, password } = req.body;
    const googleSheets = await getGoogleSheetsClient();
    const spreadsheetId = process.env.SPREADSHEET_ID;
    await ensureAdminUsersSheetExists(googleSheets, spreadsheetId);

    const response = await googleSheets.spreadsheets.values.get({ spreadsheetId, range: 'AdminUsers!A:D' });
    const rows = response.data.values || [];

    let matchedUser = null;
    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][1] || '').trim().toLowerCase() === username.trim().toLowerCase()) {
        matchedUser = { rowIdx: i + 1, name: rows[i][0], username: rows[i][1], hash: rows[i][2], role: rows[i][3] };
        break;
      }
    }

    if (!matchedUser) {
      return res.status(404).json({ status: 'error', message: 'Admin user not found.' });
    }

    const passwordMatch = await bcrypt.compare(password, matchedUser.hash);
    if (!passwordMatch) {
      return res.status(401).json({ status: 'error', message: 'Incorrect password.' });
    }

    return res.status(200).json({ status: 'success', message: 'Login successful', name: matchedUser.name, role: matchedUser.role });
  } catch (error) {
    console.error('Login error:', error);
    return res.status(500).json({ status: 'error', message: 'Internal server error' });
  }
});

app.post('/api/admin/update-password', async (req, res) => {
  try {
    const { username, oldPassword, newPassword } = req.body;
    const googleSheets = await getGoogleSheetsClient();
    const spreadsheetId = process.env.SPREADSHEET_ID;

    const response = await googleSheets.spreadsheets.values.get({ spreadsheetId, range: 'AdminUsers!A:D' });
    const rows = response.data.values || [];

    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][1] || '').trim().toLowerCase() === username.trim().toLowerCase()) {
        const rowIdx = i + 1;
        const currentHash = rows[i][2];

        const match = await bcrypt.compare(oldPassword, currentHash);
        if (!match) {
          return res.status(401).json({ status: 'error', message: 'Old password is incorrect.' });
        }

        const newHash = await bcrypt.hash(newPassword, 10);
        await googleSheets.spreadsheets.values.update({
          spreadsheetId,
          range: `AdminUsers!C${rowIdx}`,
          valueInputOption: 'USER_ENTERED',
          resource: { values: [[newHash]] }
        });

        return res.status(200).json({ status: 'success', message: 'Password updated successfully!' });
      }
    }
    return res.status(404).json({ status: 'error', message: 'User not found.' });
  } catch (error) {
    return res.status(500).json({ status: 'error', message: 'Failed to update password' });
  }
});

// --- DRIVER REGISTRATION & PASSWORD SETUP ---
app.post('/api/drivers/register', async (req, res) => {
  try {
    const { driver_name, phone_number, vehicle_id, password } = req.body;
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
    await ensureDriversSheetExists(googleSheets, spreadsheetId);

    const response = await googleSheets.spreadsheets.values.get({ spreadsheetId, range: 'Drivers!A:E' });
    const rows = response.data.values || [];
    let rowIndex = -1;
    let currentStatus = 'pending';
    let existingPasswordHash = '';

    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][0] || '').trim() === normalizedPhone) {
        rowIndex = i + 1;
        const existingStatus = (rows[i][3] || 'pending').trim().toLowerCase();
        existingPasswordHash = rows[i][4] || '';
        currentStatus = (existingStatus === 'approved') ? 'approved' : 'pending';
        break;
      }
    }

    let finalPasswordHash = existingPasswordHash;
    if (password && password.trim() !== '') {
      finalPasswordHash = await bcrypt.hash(password.trim(), 10);
    }

    if (rowIndex !== -1) {
      await googleSheets.spreadsheets.values.update({
        spreadsheetId,
        range: `Drivers!B${rowIndex}:E${rowIndex}`,
        valueInputOption: 'USER_ENTERED',
        resource: { values: [[normalizedName, normalizedVehicle, currentStatus, finalPasswordHash]] },
      });
    } else {
      const defaultNewHash = password ? await bcrypt.hash(password.trim(), 10) : '';
      await googleSheets.spreadsheets.values.append({
        spreadsheetId,
        range: 'Drivers!A:E',
        valueInputOption: 'USER_ENTERED',
        resource: { values: [[normalizedPhone, normalizedName, normalizedVehicle, 'pending', defaultNewHash]] },
      });
      currentStatus = 'pending';
    }

    return res.status(200).json({ status: currentStatus, message: `Current status: ${currentStatus}` });
  } catch (error) {
    console.error('Error in driver registration:', error);
    return res.status(500).json({ status: 'error', message: 'Internal server error' });
  }
});

// --- DRIVER LOGIN ENDPOINT ---
app.post('/api/drivers/login', async (req, res) => {
  try {
    const { phone_number, password } = req.body;
    if (!phone_number || !password) {
      return res.status(400).json({ status: 'error', message: 'Phone number and password required.' });
    }

    const googleSheets = await getGoogleSheetsClient();
    const spreadsheetId = process.env.SPREADSHEET_ID;
    await ensureDriversSheetExists(googleSheets, spreadsheetId);

    const response = await googleSheets.spreadsheets.values.get({ spreadsheetId, range: 'Drivers!A:E' });
    const rows = response.data.values || [];

    let matchedDriver = null;
    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][0] || '').trim() === phone_number.trim()) {
        matchedDriver = {
          phone_number: rows[i][0],
          driver_name: rows[i][1],
          vehicle_id: rows[i][2],
          status: rows[i][3] || 'pending',
          passwordHash: rows[i][4] || ''
        };
        break;
      }
    }

    if (!matchedDriver) {
      return res.status(404).json({ status: 'error', message: 'Driver not found. Please register first.' });
    }

    const passwordMatch = await bcrypt.compare(password.trim(), matchedDriver.passwordHash);
    if (!passwordMatch) {
      return res.status(401).json({ status: 'error', message: 'Incorrect password.' });
    }

    return res.status(200).json({
      status: 'success',
      driver_name: matchedDriver.driver_name,
      vehicle_id: matchedDriver.vehicle_id,
      driver_status: matchedDriver.status
    });
  } catch (error) {
    console.error('Driver login error:', error);
    return res.status(500).json({ status: 'error', message: 'Internal server error' });
  }
});

// --- DRIVER FORGOT / RESET PASSWORD ENDPOINT ---
app.post('/api/drivers/reset-password', async (req, res) => {
  try {
    const { phone_number, new_password } = req.body;
    if (!phone_number || !new_password) {
      return res.status(400).json({ status: 'error', message: 'Phone number and new password required.' });
    }

    const googleSheets = await getGoogleSheetsClient();
    const spreadsheetId = process.env.SPREADSHEET_ID;
    await ensureDriversSheetExists(googleSheets, spreadsheetId);

    const response = await googleSheets.spreadsheets.values.get({ spreadsheetId, range: 'Drivers!A:E' });
    const rows = response.data.values || [];
    let rowIndex = -1;

    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][0] || '').trim() === phone_number.trim()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return res.status(404).json({ status: 'error', message: 'Phone number not registered.' });
    }

    const hashedNewPassword = await bcrypt.hash(new_password.trim(), 10);
    await googleSheets.spreadsheets.values.update({
      spreadsheetId,
      range: `Drivers!E${rowIndex}`,
      valueInputOption: 'USER_ENTERED',
      resource: { values: [[hashedNewPassword]] },
    });

    return res.status(200).json({ status: 'success', message: 'Password reset successfully!' });
  } catch (error) {
    console.error('Reset password error:', error);
    return res.status(500).json({ status: 'error', message: 'Failed to reset password' });
  }
});

// --- COMBINED DRIVER HISTORY ENDPOINT (Trips & Fuel Logs) ---
app.get('/api/drivers/history/:monthYear/:driverName', async (req, res) => {
  try {
    const { monthYear, driverName } = req.params;
    const targetDriver = driverName.trim().toUpperCase();
    
    const googleSheets = await getGoogleSheetsClient();
    const spreadsheetId = process.env.SPREADSHEET_ID;

    const spreadsheet = await googleSheets.spreadsheets.get({ spreadsheetId });
    
    let driverTrips = [];
    const tripSheetExists = spreadsheet.data.sheets.some(s => s.properties.title === monthYear);
    if (tripSheetExists) {
      const tripResponse = await googleSheets.spreadsheets.values.get({ spreadsheetId, range: `${monthYear}!A:K` });
      const tripRows = tripResponse.data.values || [];
      
      for (let i = 1; i < tripRows.length; i++) {
        const r = tripRows[i];
        const rowDriver = (r[3] || '').trim().toUpperCase();
        if (rowDriver === targetDriver) {
          driverTrips.push({
            start_timestamp: r[0] || '',
            timestamp: r[1] || '',
            vehicle_id: r[2] || '',
            driver_name: r[3] || '',
            customer_name: r[4] || '',
            start_odometer: r[5] || '',
            end_odometer: r[6] || '',
            manual_dist: r[7] || '',
            start_gps: r[8] || '',
            end_gps: r[9] || '',
            gps_dist: r[10] || '',
          });
        }
      }
    }

    let driverFuelLogs = [];
    const fuelSheetExists = spreadsheet.data.sheets.some(s => s.properties.title === 'FuelLogs');
    if (fuelSheetExists) {
      const fuelResponse = await googleSheets.spreadsheets.values.get({ spreadsheetId, range: 'FuelLogs!A:K' });
      const fuelRows = fuelResponse.data.values || [];

      for (let i = 1; i < fuelRows.length; i++) {
        const r = fuelRows[i];
        const rowDriver = (r[2] || '').trim().toUpperCase();
        if (rowDriver === targetDriver) {
          driverFuelLogs.push({
            timestamp: r[0] || '',
            vehicle_id: r[1] || '',
            driver_name: r[2] || '',
            fuel_type: r[3] || '',
            shift_time: r[4] || '',
            quantity: r[5] || '',
            unit: r[6] || '',
            total_cost: r[7] || '',
            per_unit_rate: r[8] || '',
            odometer: r[9] || '',
            station_location: r[10] || ''
          });
        }
      }
    }

    return res.status(200).json({ trips: driverTrips, fuel_logs: driverFuelLogs });
  } catch (error) {
    console.error('Error fetching driver history:', error);
    return res.status(500).json({ error: 'Failed to fetch driver history' });
  }
});

// --- PROTECTED ADMIN ENDPOINTS (REQUIRES verifyAdminAuth) ---
app.get('/api/admin/drivers', verifyAdminAuth, async (req, res) => {
  try {
    const googleSheets = await getGoogleSheetsClient();
    const spreadsheetId = process.env.SPREADSHEET_ID;
    await ensureDriversSheetExists(googleSheets, spreadsheetId);

    const response = await googleSheets.spreadsheets.values.get({ spreadsheetId, range: 'Drivers!A:D' });
    const rows = response.data.values || [];
    
    const drivers = rows.slice(1)
      .filter(r => r && r.length > 0 && r[0])
      .map(r => ({
        phone_number: (r[0] || '').trim(),
        driver_name: (r[1] || '').trim().toUpperCase(),
        vehicle_id: (r[2] || '').trim().toUpperCase(),
        status: (r[3] || 'pending').trim().toLowerCase(),
      }));

    return res.status(200).json({ status: 'success', drivers });
  } catch (error) {
    console.error('Error fetching drivers:', error);
    return res.status(500).json({ status: 'error', drivers: [], message: error.message });
  }
});

app.get('/api/admin/fuel-logs', verifyAdminAuth, async (req, res) => {
  try {
    const googleSheets = await getGoogleSheetsClient();
    const spreadsheetId = process.env.SPREADSHEET_ID;
    await ensureFuelSheetExists(googleSheets, spreadsheetId);

    const response = await googleSheets.spreadsheets.values.get({ spreadsheetId, range: 'FuelLogs!A:K' });
    const rows = response.data.values || [];

    const fuelLogs = rows.slice(1)
      .filter(r => r && r.length > 0 && r[0])
      .map(r => ({
        timestamp: r[0] || '',
        vehicle_id: (r[1] || '').trim().toUpperCase(),
        driver_name: (r[2] || '').trim().toUpperCase(),
        fuel_type: r[3] || 'CNG',
        shift_time: r[4] || '',
        quantity: r[5] || '',
        unit: r[6] || 'KG',
        total_cost: r[7] || '',
        per_unit_rate: r[8] || '',
        odometer: r[9] || '',
        station_location: r[10] || ''
      }));

    return res.status(200).json({ status: 'success', fuel_logs: fuelLogs });
  } catch (error) {
    console.error('Error fetching fuel logs for admin:', error);
    return res.status(500).json({ status: 'error', fuel_logs: [], message: error.message });
  }
});

app.post('/api/admin/drivers/update-status', verifyAdminAuth, async (req, res) => {
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

app.get('/api/admin/driver-details/:driverName', verifyAdminAuth, async (req, res) => {
  try {
    const driverName = req.params.driverName.trim().toUpperCase();
    const phoneParam = (req.query.phone || '').trim();

    const googleSheets = await getGoogleSheetsClient();
    const spreadsheetId = process.env.SPREADSHEET_ID;
    
    await ensureDriversSheetExists(googleSheets, spreadsheetId);

    const response = await googleSheets.spreadsheets.values.get({ spreadsheetId, range: 'Drivers!A:D' });
    const rows = response.data.values || [];

    let matchedRow = rows.slice(1).find(r => 
      (r[0] || '').trim() === phoneParam || 
      (r[1] || '').trim().toUpperCase() === driverName
    );

    let matchedVehicleId = matchedRow ? (matchedRow[2] || '').trim().toUpperCase().replace(/[\s-]/g, '') : null;
    let matchedPhone = matchedRow ? (matchedRow[0] || '').trim() : 'N/A';
    let matchedName = matchedRow ? (matchedRow[1] || '').trim().toUpperCase() : driverName;

    let foundLocation = null;
    for (const [vid, loc] of Object.entries(liveFleetTracker)) {
      const cleanVid = vid.trim().toUpperCase().replace(/[\s-]/g, '');
      if (cleanVid === matchedVehicleId || vid.trim().toUpperCase() === driverName) {
        foundLocation = loc;
        break;
      }
    }

    let locationInfo = { 
      latitude: foundLocation ? foundLocation.latitude : 'N/A', 
      longitude: foundLocation ? foundLocation.longitude : 'N/A', 
      speed: foundLocation ? foundLocation.speed : 0, 
      heading: foundLocation ? 'Active on road' : 'Stationary', 
      timestamp: foundLocation ? foundLocation.last_updated : 'No recent transmission' 
    };
    
    res.status(200).json({ 
      driver_name: matchedName, 
      phone_number: matchedPhone,
      vehicle_id: matchedVehicleId || 'N/A',
      ...locationInfo 
    });
  } catch (error) {
    console.error('Error fetching driver details:', error);
    res.status(500).json({ error: 'Failed to fetch driver details' });
  }
});

app.get('/api/admin/billing-summary/:monthYear', verifyAdminAuth, async (req, res) => {
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

app.get('/api/admin/driver-summary/:monthYear', verifyAdminAuth, async (req, res) => {
  req.url = `/api/admin/billing-summary/${req.params.monthYear}`;
  return app._router.handle(req, res);
});

// --- UPDATED EXCEL EXPORT ENDPOINT WITH NORMALIZED DATE MAPPING ---
app.get('/api/admin/export-excel/:monthYear', verifyAdminAuth, async (req, res) => {
  try {
    const sheetName = req.params.monthYear;
    const googleSheets = await getGoogleSheetsClient();
    const spreadsheetId = process.env.SPREADSHEET_ID;

    const spreadsheet = await googleSheets.spreadsheets.get({ spreadsheetId });
    
    // 1. Fetch Monthly Trips
    const tripResponse = await googleSheets.spreadsheets.values.get({ spreadsheetId, range: `${sheetName}!A:K` });
    const rows = tripResponse.data.values || [];

    // 2. Fetch Fuel Logs
    let fuelRows = [];
    const fuelSheetExists = spreadsheet.data.sheets.some(s => s.properties.title === 'FuelLogs');
    if (fuelSheetExists) {
      const fuelResponse = await googleSheets.spreadsheets.values.get({ spreadsheetId, range: 'FuelLogs!A:K' });
      fuelRows = fuelResponse.data.values || [];
    }

    const workbook = new ExcelJS.Workbook();

    // --- SHEET 1: Monthly Fleet Logs ---
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

    const clientMap = {};
    const driverMap = {};
    const vehicleTripsMap = {};
    const vehicleFuelMap = {};

    if (rows && rows.length > 1) {
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        logSheet.addRow({
          start_timestamp: r[0] || '', timestamp: r[1] || '', vehicle_id: r[2] || '', driver_name: r[3] || '', customer_name: r[4] || '',
          start_odometer: r[5] || '', end_odometer: r[6] || '', manual_dist: r[7] || '',
          start_gps: r[8] || '', end_gps: r[9] || '', gps_dist: r[10] || '',
        });

        const customerName = (r[4] || 'GENERAL CLIENT').trim().toUpperCase();
        const driverName = (r[3] || 'UNKNOWN DRIVER').trim().toUpperCase();
        const vehicleId = (r[2] || 'UNKNOWN').trim().toUpperCase();
        const manualDist = parseFloat(r[7]) || 0;

        if (!clientMap[customerName]) clientMap[customerName] = { total_trips: 0, total_km: 0 };
        clientMap[customerName].total_trips += 1;
        clientMap[customerName].total_km += manualDist;

        if (!driverMap[driverName]) driverMap[driverName] = { total_trips: 0, total_km: 0 };
        driverMap[driverName].total_trips += 1;
        driverMap[driverName].total_km += manualDist;

        if (!vehicleTripsMap[vehicleId]) vehicleTripsMap[vehicleId] = {};
        
        // --- USING NORMALIZED DATE KEY ---
        const rawDateStr = normalizeDateKey(r[0]);
        vehicleTripsMap[vehicleId][rawDateStr] = {
          date: r[0]?.split(' ')[0] || '',
          start: r[0]?.split(' ')[1] || '08:00 AM',
          end: r[1]?.split(' ')[1] || '20:00 PM',
          startOdo: r[5] || '',
          endOdo: r[6] || '',
          dist: manualDist
        };
      }
    }

    // Process Fuel Logs with Normalized Date Key
    if (fuelRows && fuelRows.length > 1) {
      for (let i = 1; i < fuelRows.length; i++) {
        const f = fuelRows[i];
        const fuelTimestamp = f[0] || '';
        const fuelVehicle = (f[1] || '').trim().toUpperCase();
        const fuelType = f[3] || 'CNG';
        const qty = parseFloat(f[5]) || 0;
        const unit = f[6] || 'KG';
        const cost = parseFloat(f[7]) || 0;
        const odo = f[9] || '';
        
        // --- USING NORMALIZED DATE KEY ---
        const dateKey = normalizeDateKey(fuelTimestamp);

        if (!vehicleFuelMap[fuelVehicle]) vehicleFuelMap[fuelVehicle] = {};
        if (!vehicleFuelMap[fuelVehicle][dateKey]) {
          vehicleFuelMap[fuelVehicle][dateKey] = { petrol: 0, totalCost: 0, cngFills: [], startingOdo: odo, closingOdo: odo };
        }

        const dayRecord = vehicleFuelMap[fuelVehicle][dateKey];
        dayRecord.totalCost += cost;
        dayRecord.closingOdo = odo || dayRecord.closingOdo;
        if (!dayRecord.startingOdo) dayRecord.startingOdo = odo;

        if (fuelType === 'Petrol') {
          dayRecord.petrol += qty;
        } else if (fuelType === 'CNG') {
          dayRecord.cngFills.push(`${qty} ${unit} (₹${cost})`);
        }
      }
    }

    // --- SHEET 2: Billing Summary ---
    const summarySheet = workbook.addWorksheet('Billing Summary');
    summarySheet.columns = [
      { header: 'Category / Name', key: 'name', width: 30 },
      { header: 'Total Trips', key: 'trips', width: 15 },
      { header: 'Total Distance (km)', key: 'km', width: 20 },
    ];

    summarySheet.addRow({ name: '--- DRIVER SUMMARIES ---', trips: '', km: '' });
    for (const [driver, data] of Object.entries(driverMap)) {
      summarySheet.addRow({ name: driver, trips: data.total_trips, km: parseFloat(data.total_km.toFixed(2)) });
    }

    summarySheet.addRow({ name: '', trips: '', km: '' });
    summarySheet.addRow({ name: '--- CLIENT / CUSTOMER SUMMARIES ---', trips: '', km: '' });
    for (const [client, data] of Object.entries(clientMap)) {
      summarySheet.addRow({ name: client, trips: data.total_trips, km: parseFloat(data.total_km.toFixed(2)) });
    }

    const [monthStr, yearStr] = sheetName.split(' ');
    const monthIndex = new Date(`${monthStr} 1, ${yearStr || new Date().getFullYear()}`).getMonth();
    const year = parseInt(yearStr) || new Date().getFullYear();
    const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();

    // --- SHEET 3+: Vehicle Trip Cards ---
    for (const [vehicleId, recordedTrips] of Object.entries(vehicleTripsMap)) {
      const ws = workbook.addWorksheet(vehicleId);
      
      ws.mergeCells('A1:H1');
      ws.getCell('A1').value = `VEHICLE NUMBER (${vehicleId}) FOR THE MONTH OF ${sheetName.toUpperCase()}`;
      ws.getCell('A1').font = { bold: true, size: 12 };
      ws.getCell('A1').alignment = { horizontal: 'center' };

      ws.addRow(['DATE', 'STARTING TIME', 'CLOSING TIME', 'TOTAL HOURS', 'EXTRA HOURS', 'STARTING Kms', 'CLOSING Kms', 'TOTAL Kms']);

      let totalKmsMonthly = 0;

      for (let day = 1; day <= daysInMonth; day++) {
        const currentDate = new Date(year, monthIndex, day);
        const dayFormatted = String(day).padStart(2, '0');
        const monthFormatted = String(monthIndex + 1).padStart(2, '0');
        const dateKey1 = `${dayFormatted}.${monthFormatted}.${year}`;
        const dateKey2 = `${year}-${monthFormatted}-${dayFormatted}`;

        const isSunday = currentDate.getDay() === 0;
        const isFirstMayHoliday = (monthIndex === 4 && day === 1);

        if (recordedTrips[dateKey1] || recordedTrips[dateKey2]) {
          const t = recordedTrips[dateKey1] || recordedTrips[dateKey2];
          ws.addRow([t.date, t.start, t.end, 12, 0, t.startOdo, t.endOdo, t.dist]);
          totalKmsMonthly += t.dist;
        } else if (isSunday) {
          ws.addRow([dateKey1, 'SUNDAY', '', '', '', '', '', 0]);
          ws.mergeCells(`B${ws.rowCount}:H${ws.rowCount}`);
          ws.getCell(`B${ws.rowCount}`).alignment = { horizontal: 'center' };
        } else if (isFirstMayHoliday) {
          ws.addRow([dateKey1, 'MAY FIRST HOLIDAY', '', '', '', '', '', 0]);
          ws.mergeCells(`B${ws.rowCount}:H${ws.rowCount}`);
          ws.getCell(`B${ws.rowCount}`).alignment = { horizontal: 'center' };
        } else {
          ws.addRow([dateKey1, '', '', '', '', '', '', 0]);
        }
      }

      const totalRow = ws.addRow(['TOTAL', '', '', '', '', '', '', totalKmsMonthly]);
      totalRow.font = { bold: true };
      ws.columns = [{width: 15}, {width: 15}, {width: 15}, {width: 15}, {width: 15}, {width: 15}, {width: 15}, {width: 15}];
    }

    // --- FUEL STATEMENT SHEET ---
    const fuelWs = workbook.addWorksheet('Fuel Statement');
    fuelWs.addRow(['Vehicle No', 'Month', 'Date', 'Received Amount', 'Petrol', 'GNG Filling 1', 'GNG Filling 2', 'GNG Filling 3', 'Total Amount', 'Starting Kms', 'Closing Kms', 'Total Kms']);

    for (const [vehicleId, datesObj] of Object.entries(vehicleFuelMap)) {
      for (const [dateVal, fData] of Object.entries(datesObj)) {
        const startKms = parseFloat(fData.startingOdo) || 0;
        const closeKms = parseFloat(fData.closingOdo) || 0;
        const totalKmsCalc = (closeKms > startKms) ? (closeKms - startKms).toFixed(1) : 0;

        fuelWs.addRow([
          vehicleId,
          sheetName,
          dateVal,
          fData.totalCost,
          fData.petrol > 0 ? fData.petrol : '',
          fData.cngFills[0] || '',
          fData.cngFills[1] || '',
          fData.cngFills[2] || '',
          fData.totalCost,
          fData.startingOdo,
          fData.closingOdo,
          totalKmsCalc
        ]);
      }
    }

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=Professional_Billing_Report_${sheetName.replace(' ', '_')}.xlsx`);
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    res.status(500).send('Error generating report: ' + error.message);
  }
});

// --- TRIPS & LIVE FLEET TRACKING ---
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

// --- FUEL / GAS FILLING LOGGING ENDPOINT ---
app.post('/api/fuel/sync', async (req, res) => {
  try {
    const { 
      vehicle_id, 
      driver_name, 
      fuel_type, 
      shift_time, 
      quantity, 
      unit, 
      total_cost, 
      per_unit_rate, 
      odometer, 
      timestamp, 
      station_location 
    } = req.body;
    
    if (!vehicle_id || !quantity || !total_cost) {
      return res.status(400).json({ status: 'error', message: 'Missing required fuel fields.' });
    }

    const googleSheets = await getGoogleSheetsClient();
    const spreadsheetId = process.env.SPREADSHEET_ID;
    
    await ensureFuelSheetExists(googleSheets, spreadsheetId);

    const rowData = [
      timestamp || new Date().toLocaleString(),
      vehicle_id.trim().toUpperCase(),
      (driver_name || '').trim().toUpperCase(),
      fuel_type || 'CNG',
      shift_time || 'Mid-Day',
      quantity,
      unit || 'KG',
      total_cost,
      per_unit_rate || '',
      odometer || '',
      station_location || ''
    ];

    await googleSheets.spreadsheets.values.append({
      spreadsheetId,
      range: 'FuelLogs!A:K',
      valueInputOption: 'USER_ENTERED',
      resource: { values: [rowData] },
    });

    return res.status(200).json({ status: 'success', message: 'Fuel record logged successfully!' });
  } catch (error) {
    console.error('Error syncing fuel log:', error);
    return res.status(500).json({ status: 'error', message: 'Internal Server Error while syncing fuel log' });
  }
});

// --- ROUTE BREADCRUMBS ENDPOINTS ---
app.post('/api/trips/breadcrumb', async (req, res) => {
  try {
    const breadcrumb = req.body;
    console.log(`Received trip breadcrumb for Vehicle: ${breadcrumb.vehicle_id}`);
    res.status(200).json({ status: 'success' });
  } catch (error) {
    res.status(500).json({ status: 'error', message: error.message });
  }
});

app.post('/api/fleet/breadcrumb', async (req, res) => {
  try {
    const breadcrumb = req.body;
    console.log(`Received shift breadcrumb for Vehicle: ${breadcrumb.vehicle_id}`);
    res.status(200).json({ status: 'success' });
  } catch (error) {
    res.status(500).json({ status: 'error', message: error.message });
  }
});

app.post('/api/fleet/location', (req, res) => {
  const { vehicle_id, latitude, longitude, speed } = req.body;
  if (vehicle_id) liveFleetTracker[vehicle_id.trim().toUpperCase()] = { latitude, longitude, speed, last_updated: new Date().toLocaleString() };
  res.status(200).json({ status: 'success' });
});

app.listen(PORT, '0.0.0.0', () => console.log(`--- SERVER RUNNING ON PORT ${PORT} ---`));