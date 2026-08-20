const express = require('express');
const cors = require('cors');
const path = require('path');
const { google } = require('googleapis');
require('dotenv').config();

const app = express();
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());
app.use(cors());  

const PORT = process.env.PORT || 3000;

// In-memory storage for real-time Master tracking of active cars
const liveFleetTracker = {};

// Helper function to calculate distance in kilometers using the Haversine formula
function calculateDistance(lat1, lon1, lat2, lon2) {
  const R = 6371; // Earth radius in km
  const dLat = (lat2 - lat1) * (Math.PI / 180);
  const dLon = (lon2 - lon1) * (Math.PI / 180);
  const a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(lat1 * (Math.PI / 180)) * Math.cos(lat2 * (Math.PI / 180)) *
    Math.sin(dLon / 2) * Math.sin(dLon / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c; // Distance in km
}

// ----------------------------------------------------
// 1. GOOGLE SHEETS FUNCTION (Enterprise Dual-Tracking)
// ----------------------------------------------------
async function appendToGoogleSheet(tripData) {
  try {
    let auth;
    if (process.env.GOOGLE_CREDENTIALS_JSON) {
      // Production on Render using Environment Variable with safe trimming
      const rawCreds = process.env.GOOGLE_CREDENTIALS_JSON.trim();
      const credentials = JSON.parse(rawCreds);

      auth = new google.auth.GoogleAuth({
        credentials,
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
      });
    } else {
      // Local development using credentials.json file
      auth = new google.auth.GoogleAuth({
        keyFile: 'credentials.json',
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
      });
    }

    const client = await auth.getClient();
    const googleSheets = google.sheets({ version: 'v4', auth: client });

    const spreadsheetId = process.env.SPREADSHEET_ID;

    // Calculate manual odometer distance
    let manualDistance = '';
    if (tripData.start_odometer && tripData.end_odometer) {
      manualDistance = parseFloat(tripData.end_odometer) - parseFloat(tripData.start_odometer);
    }

    // Calculate GPS distance if start and end coordinates are provided
    let gpsDistanceKm = '';
    if (tripData.start_gps && tripData.end_gps) {
      const [startLat, startLon] = tripData.start_gps.split(',').map(Number);
      const [endLat, endLon] = tripData.end_gps.split(',').map(Number);
      
      if (!isNaN(startLat) && !isNaN(startLon) && !isNaN(endLat) && !isNaN(endLon)) {
        gpsDistanceKm = calculateDistance(startLat, startLon, endLat, endLon).toFixed(2);
      }
    }

    // Columns: Date/Time, Vehicle No, Driver Name, Customer, Start Odo, End Odo, Manual Dist, Start GPS, End GPS, GPS Dist (km)
    const rowData = [
      tripData.timestamp || new Date().toLocaleString(),
      tripData.vehicle_id || '',
      tripData.driver_name || '',
      tripData.customer_name || '',
      tripData.start_odometer || '',
      tripData.end_odometer || '',
      manualDistance,
      tripData.start_gps || '',
      tripData.end_gps || '',
      gpsDistanceKm
    ];

    await googleSheets.spreadsheets.values.append({
      spreadsheetId,
      range: 'Sheet1!A:J', // Extended range from A through J for enterprise tracking
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [rowData],
      },
    });
    console.log(`[Google Sheets] Logged enterprise trip for Driver: ${tripData.driver_name} | Vehicle: ${tripData.vehicle_id}`);
  } catch (error) {
    console.error('Error writing to Google Sheet:', error);
  }
}

// ----------------------------------------------------
// 2. API ENDPOINTS
// ----------------------------------------------------

// Endpoint for Mobile App / Web to sync enterprise trips to Google Sheets
app.post('/api/trips/sync', async (req, res) => {
  try {
    const { vehicle_id, driver_name, trips } = req.body;

    if (!trips || trips.length === 0) {
      return res.status(400).json({ error: 'No trips provided' });
    }

    for (const trip of trips) {
      await appendToGoogleSheet({
        vehicle_id,
        driver_name: driver_name || 'Web User',
        customer_name: trip.customer_name || 'General Route',
        start_odometer: trip.start_odometer,
        end_odometer: trip.end_odometer,
        start_gps: trip.start_gps || '',
        end_gps: trip.end_gps || '',
        timestamp: trip.timestamp
      });
    }

    res.status(200).json({ status: 'success', message: 'Enterprise trip data synced to Google Sheets!' });
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