process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';

require('dotenv').config();

const express = require('express');
const { CloudAdapter, ConfigurationBotFrameworkAuthentication } = require('botbuilder');
const { AzureDevOpsBot } = require('./src/bot');

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication({
  MicrosoftAppId: process.env.BOT_APP_ID,
  MicrosoftAppPassword: process.env.BOT_APP_SECRET,
  MicrosoftAppTenantId: process.env.BOT_TENANT_ID,
  MicrosoftAppType: 'SingleTenant'
});

const adapter = new CloudAdapter(botFrameworkAuthentication);

adapter.onTurnError = async (context, error) => {
  console.error('[onTurnError]', error);
  await context.sendActivity('Something went wrong. Please try again.');
};

const bot = new AzureDevOpsBot();

const app = express();
app.use(express.json());

app.post('/api/messages', async (req, res) => {
  await adapter.process(req, res, (context) => bot.run(context));
});

app.listen(process.env.PORT || 3978, () => {
  console.log(`Bot is running on port ${process.env.PORT || 3978}`);
});
