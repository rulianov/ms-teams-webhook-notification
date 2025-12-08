// server.js
const express = require('express');
const axios = require('axios');

const app = express();
const PORT = process.env.PORT || 8080;
const WEBHOOK_URL = process.env.TEAMS_WEBHOOK_URL;

app.use(express.json());

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// Main webhook endpoint
app.post('/alert', async (req, res) => {
  const alertmanagerPayload = req.body;

  // Log the original Alertmanager message to stdout
  console.log('='.repeat(80));
  console.log('RECEIVED ALERTMANAGER WEBHOOK:');
  console.log(JSON.stringify(alertmanagerPayload, null, 2));
  console.log('='.repeat(80));

  if (!WEBHOOK_URL) {
    console.error('ERROR: TEAMS_WEBHOOK_URL environment variable not set');
    return res.status(500).json({ error: 'Webhook URL not configured' });
  }

  try {
    await sendAlertsToTeams(alertmanagerPayload);
    res.status(200).json({ message: 'Alerts sent to Teams successfully' });
  } catch (error) {
    console.error('ERROR sending to Teams:');
    console.error(error.message);
    if (error.response) {
      console.error('Response status:', error.response.status);
      console.error('Response data:', error.response.data);
    }
    res.status(500).json({ error: 'Failed to send alert to Teams' });
  }
});

// Determine color based on status
function getColor(status) {
  switch (status.toLowerCase()) {
    case 'firing':
      return 'attention';
    case 'resolved':
      return 'good';
    default:
      return 'default';
  }
}

// Create adaptive card for a single alert
function createAdaptiveCard(alert, status) {
  const alertname = alert.labels?.alertname || 'Unknown Alert';
  const resourceName = alert.labels?.resourceName || 'N/A';
  const summary = alert.annotations?.summary || '';
  const description = alert.annotations?.description || '';
  const details = alert.labels?.details || '';

  return {
    type: 'message',
    summary: `Prometheus Alert: ${alertname} - ${status.toUpperCase()}`,
    attachments: [
      {
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: {
          $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
          type: 'AdaptiveCard',
          version: '1.4',
          body: [
            {
              type: 'Container',
              style: getColor(status),
              items: [
                {
                  type: 'TextBlock',
                  text: alertname,
                  weight: 'Bolder',
                  size: 'Large',
                  wrap: true,
                },
                {
                  type: 'TextBlock',
                  text: `Resource Name: ${resourceName}`,
                  wrap: true,
                  spacing: 'Small',
                },
                {
                  type: 'TextBlock',
                  text: `Summary: ${summary}`,
                  wrap: true,
                  spacing: 'Small',
                },
                {
                  type: 'TextBlock',
                  text: `Description: ${description}`,
                  wrap: true,
                  spacing: 'Small',
                },
                {
                  type: 'TextBlock',
                  text: `Details: ${details}`,
                  wrap: true,
                  spacing: 'Small',
                },
              ],
            },
          ],
        },
      },
    ],
  };
}

// Send all alerts to Teams
async function sendAlertsToTeams(alertmanagerPayload) {
  const alerts = alertmanagerPayload.alerts || [];
  const status = alertmanagerPayload.status || 'unknown';

  for (const alert of alerts) {
    const adaptiveCard = createAdaptiveCard(alert, status);

    // Log the adaptive card being sent
    console.log('SENDING ADAPTIVE CARD:');
    console.log(JSON.stringify(adaptiveCard, null, 2));
    console.log('='.repeat(80));

    const response = await axios.post(WEBHOOK_URL, adaptiveCard, {
      headers: { 'Content-Type': 'application/json' },
    });

    console.log(`✓ Successfully sent to Teams. Status: ${response.status}`);
  }
}

app.listen(PORT, () => {
  console.log(`Alertmanager to Teams adapter running on port ${PORT}`);
  console.log(`Webhook URL configured: ${WEBHOOK_URL ? 'Yes' : 'No'}`);
  console.log('Endpoints:');
  console.log(`  POST /alert - Receive Alertmanager webhooks`);
  console.log(`  GET /health - Health check`);
});
