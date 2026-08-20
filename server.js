const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');
require('dotenv').config();

const app = express();
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());
app.use(cors());  


const PORT = process.env.PORT || 3000;

// In-memory storage for real-time Master tracking of active cars
const liveFleetTracker = {};

// ----------------------------------------------------
// 1. GOOGLE SHEETS FUNCTION (For permanent trip records)
// ----------------------------------------------------
async function appendToGoogleSheet(tripData) {
  try {
    const auth = new google.auth.GoogleAuth({
      keyFile: 'credentials.json',
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const client = await auth.getClient();
    const googleSheets = google.sheets({ version: 'v4', auth: client });

    const spreadsheetId = process.env.SPREADSHEET_ID;

    // Columns: Vehicle ID, Start Odometer, End Odometer, Distance, Timestamp
    const rowData = [
      tripData.vehicle_id,
      tripData.start_odometer,
      tripData.end_odometer,
      tripData.distance,
      tripData.timestamp || new Date().toLocaleString()
    ];

    await googleSheets.spreadsheets.values.append({
      spreadsheetId,
      range: 'Sheet1!A:E',
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [rowData],
      },
    });
    console.log(`[Google Sheets] Logged trip for ${tripData.vehicle_id}`);
  } catch (error) {
    console.error('Error writing to Google Sheet:', error);
  }
}

// ----------------------------------------------------
// 2. API ENDPOINTS
// ----------------------------------------------------

// Endpoint for Mobile App to sync completed trips to Google Sheets
app.post('/api/trips/sync', async (req, res) => {
  try {
    const { vehicle_id, trips } = req.body;

    if (!trips || trips.length === 0) {
      return res.status(400).json({ error: 'No trips provided' });
    }

    for (const trip of trips) {
      await appendToGoogleSheet({
        vehicle_id,
        start_odometer: trip.start_odometer,
        end_odometer: trip.end_odometer,
        distance: trip.end_odometer - trip.start_odometer,
        timestamp: trip.timestamp
      });
    }

    res.status(200).json({ status: 'success', message: 'Synced to Google Sheets!' });
  } catch (error) {
    console.error('Error syncing trips:', error);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});

// Endpoint for Mobile App to send live GPS coordinates (Master system)
app.post('/api/fleet/location', async (req, res) => {
  try {
    const { vehicle_id, latitude, longitude, speed, timestamp } = req.body;

    if (!vehicle_id || !latitude || !longitude) {
      return res.status(400).json({ error: 'Missing required location data' });
    }

    // Update Master live tracking dictionary
    liveFleetTracker[vehicle_id] = {
      latitude,
      longitude,
      speed: speed || 0,
      last_updated: timestamp || new Date().toISOString()
    };

    console.log(`[Master System] Location updated -> Vehicle: ${vehicle_id} | Lat: ${latitude} | Lng: ${longitude}`);

    return res.status(200).json({ status: 'success', message: 'Location recorded by master system' });
  } catch (error) {
    console.error('Master tracking error:', error);
    return res.status(500).json({ error: 'Server error processing location' });
  }
});

// Endpoint for your Master dashboard to view all active vehicles running
app.get('/api/fleet/live-status', (req, res) => {
  return res.status(200).json({
    total_active_vehicles: Object.keys(liveFleetTracker).length,
    fleet_data: liveFleetTracker
  });
});

// ----------------------------------------------------
// 3. START SERVER
// ----------------------------------------------------
app.listen(PORT, () => {
  console.log(`--- FLEET MASTER SERVER RUNNING ON PORT ${PORT} ---`);
});
