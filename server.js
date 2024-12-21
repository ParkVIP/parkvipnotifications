const express = require('express');
const { Twilio } = require('twilio');
const axios = require('axios');
const fs = require('fs');
require('dotenv').config();

const app = express();
const PORT = 3000;

// Middleware to serve static files
app.use(express.static('public'));

// Twilio Client Setup
const twilioClient = new Twilio(process.env.TWILIO_API_KEY, process.env.TWILIO_API_SECRET, {
  accountSid: process.env.TWILIO_ACCOUNT_SID,
});

// Load tokens from file on server start
if (fs.existsSync('tokens.json')) {
  const tokens = JSON.parse(fs.readFileSync('tokens.json', 'utf-8'));
  process.env.MICROSOFT_GRAPH_TOKEN = tokens.access_token;
  process.env.MICROSOFT_GRAPH_REFRESH_TOKEN = tokens.refresh_token;
}

// Function to refresh the Microsoft Graph token
async function refreshMicrosoftToken() {
  try {
    const response = await axios.post(
      `https://login.microsoftonline.com/${process.env.MICROSOFT_GRAPH_TENANT_ID}/oauth2/v2.0/token`,
      new URLSearchParams({
        client_id: process.env.MICROSOFT_GRAPH_CLIENT_ID,
        client_secret: process.env.MICROSOFT_GRAPH_CLIENT_SECRET,
        refresh_token: process.env.MICROSOFT_GRAPH_REFRESH_TOKEN,
        grant_type: 'refresh_token',
        scope: 'https://graph.microsoft.com/.default offline_access',
      }),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );

    const { access_token, refresh_token } = response.data;
    process.env.MICROSOFT_GRAPH_TOKEN = access_token;
    process.env.MICROSOFT_GRAPH_REFRESH_TOKEN = refresh_token;

    // Save new tokens to file
    fs.writeFileSync(
      'tokens.json',
      JSON.stringify({ access_token, refresh_token }, null, 2)
    );

    console.log('Microsoft token refreshed successfully');
  } catch (error) {
    console.error('Error refreshing Microsoft token:', error.response?.data || error.message);
  }
}
app.get('/view-emails', async (req, res) => {
  try {
    const response = await axios.get(
      `${process.env.MICROSOFT_GRAPH_API_URL}/users/${process.env.MICROSOFT_GRAPH_EMAIL}/messages`,
      {
        headers: { Authorization: `Bearer ${process.env.MICROSOFT_GRAPH_TOKEN}` },
      }
    );

    // Map the response to create logs with "to" (recipients), date, and body
    const emailLogs = response.data.value.map((email) => ({
      email: email.toRecipients.map((r) => r.emailAddress.address).join(', '), // Fetch all recipients
      date: email.receivedDateTime, // Date the email was received
      body: email.body?.content || "No content available", // Check for valid body content
    }));

    res.json(emailLogs);
  } catch (error) {
    if (error.response && error.response.status === 401) {
      console.log('Token expired. Refreshing token...');
      await refreshMicrosoftToken(); // Refresh token
      try {
        // Retry the API request after refreshing the token
        const retryResponse = await axios.get(
          `${process.env.MICROSOFT_GRAPH_API_URL}/users/${process.env.MICROSOFT_GRAPH_EMAIL}/messages`,
          {
            headers: { Authorization: `Bearer ${process.env.MICROSOFT_GRAPH_TOKEN}` },
          }
        );

        const emailLogs = retryResponse.data.value.map((email) => ({
          email: email.toRecipients.map((r) => r.emailAddress.address).join(', '),
          date: email.receivedDateTime,
          body: email.body?.content || "No content available", // Check for valid body content
        }));

        return res.json(emailLogs);
      } catch (retryError) {
        return res
          .status(500)
          .json({ error: `Error retrying Email logs: ${retryError.message}` });
      }
    }
    res.status(500).json({ error: `Error fetching Email logs: ${error.message}` });
  }
});



// SMS Logs Route
app.get('/view-sms', async (req, res) => {
  try {
    const messages = await twilioClient.messages.list({ limit: 10 });

    const smsLogs = messages.map((msg, index) => ({
      id: index + 1,
      phoneNumber: msg.to || "Unknown",
      date: msg.dateSent ? new Date(msg.dateSent).toLocaleString() : "Unknown",
      message: msg.body ? msg.body.replace(/\n/g, '<br>') : "No content", // Replace newlines with <br> for HTML
    }));

    res.json(smsLogs);
  } catch (error) {
    console.error('Error fetching SMS logs:', error.message);
    res.status(500).json({ error: `Error fetching SMS logs: ${error.message}` });
  }
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
