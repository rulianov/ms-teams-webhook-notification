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

// Create adaptive card for a group of alerts
function createGroupedAdaptiveCard(alertmanagerPayload) {
  const alerts = alertmanagerPayload.alerts || [];
  const status = alertmanagerPayload.status || 'unknown';
  const groupName = alertmanagerPayload.groupLabels?.alertname || 'Alert Group';
  const color = getColor(status);

  // Card header
  const headerText = `${groupName} Alarm Group (${status.toUpperCase()})`;

  // Build alert items - each alert as a separate container with facts
  const alertItems = alerts.map((alert, index) => {
    const labels = alert.labels || {};
    const annotations = alert.annotations || {};
    const resourceName = labels.resourceName || 'N/A';
    const subject = labels.subject || 'N/A';
    const description = annotations.description || 'N/A';
    const details = labels.details || 'N/A';
    
    // Try to extract channel_id from details if present
    let channelId = '';
    const channelIdMatch = details.match(/channel_id:\s*([^|]+)/);
    if (channelIdMatch) {
      channelId = channelIdMatch[1].trim();
    }

    const facts = [
      { title: 'Resource Name', value: resourceName },
      { title: 'Subject', value: subject },
      { title: 'Description', value: description },
      { title: 'Details', value: details },
    ];

    // Only add Channel ID if it exists
    if (channelId) {
      facts.push({ title: 'Channel ID', value: channelId });
    }

    return {
      type: 'Container',
      separator: index > 0, // Add separator between alerts
      spacing: 'Medium',
      items: [
        {
          type: 'TextBlock',
          text: `Alert ${index + 1}`,
          weight: 'Bolder',
          size: 'Medium',
          wrap: true,
        },
        {
          type: 'FactSet',
          facts: facts,
        },
      ],
    };
  });

  return {
    type: 'message',
    summary: `Prometheus Alert: ${headerText}`,
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
              style: color,
              items: [
                {
                  type: 'TextBlock',
                  text: headerText,
                  weight: 'Bolder',
                  size: 'Large',
                  wrap: true,
                },
                {
                  type: 'TextBlock',
                  text: `Total Alerts: ${alerts.length}`,
                  weight: 'Bolder',
                  spacing: 'Small',
                  wrap: true,
                },
              ],
            },
            ...alertItems,
          ],
        },
      },
    ],
  };
}

// Send grouped alerts to Teams as a single message
async function sendAlertsToTeams(alertmanagerPayload) {
  const adaptiveCard = createGroupedAdaptiveCard(alertmanagerPayload);

  // Log the adaptive card being sent
  console.log('SENDING GROUPED ADAPTIVE CARD:');
  console.log(JSON.stringify(adaptiveCard, null, 2));
  console.log('='.repeat(80));

  const response = await axios.post(WEBHOOK_URL, adaptiveCard, {
    headers: { 'Content-Type': 'application/json' },
  });

  console.log(
    `✓ Successfully sent grouped alert to Teams. Status: ${response.status}`,
  );
}

app.listen(PORT, () => {
  console.log(`Alertmanager to Teams adapter running on port ${PORT}`);
  console.log(`Webhook URL configured: ${WEBHOOK_URL ? 'Yes' : 'No'}`);
  console.log('Endpoints:');
  console.log(`  POST /alert - Receive Alertmanager webhooks`);
  console.log(`  GET /health - Health check`);
});
