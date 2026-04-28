process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";

require("dotenv").config();

const express = require("express");
const {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication,
} = require("botbuilder");
const { AzureDevOpsBot } = require("./src/bot");

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication({
  MicrosoftAppId: process.env.BOT_APP_ID,
  MicrosoftAppPassword: process.env.BOT_APP_SECRET,
  MicrosoftAppTenantId: process.env.BOT_TENANT_ID,
  MicrosoftAppType: "SingleTenant",
});

const adapter = new CloudAdapter(botFrameworkAuthentication);

adapter.onTurnError = async (context, error) => {
  console.error("[onTurnError]", error);
  await context.sendActivity("Something went wrong. Please try again.");
};

const bot = new AzureDevOpsBot();

const app = express();
app.use(express.json());

app.post("/api/messages", async (req, res) => {
  await adapter.process(req, res, (context) => bot.run(context));
});

// ── ADO Service Hook: pipeline build completed ──────────────────────
app.post("/api/webhooks/build-complete", async (req, res) => {
  try {
    const payload = req.body;
    const resource = payload?.resource;

    if (!resource) {
      res.status(200).send("OK — no resource in payload");
      return;
    }

    const buildId = resource.id;
    const buildNumber = resource.buildNumber || buildId;
    const result = resource.result || resource.status || "unknown";
    const pipelineName = resource.definition?.name || "Unknown Pipeline";
    const repoName = resource.repository?.name || "";
    const branch = (resource.sourceBranch || "").replace("refs/heads/", "");
    const requestedBy = resource.requestedFor?.displayName || "Unknown";
    const buildUrl = resource._links?.web?.href ||
      `https://dev.azure.com/${process.env.ADO_ORG}/${process.env.ADO_PROJECT}/_build/results?buildId=${buildId}`;

    const emoji = {
      succeeded: "✅",
      partiallySucceeded: "⚠️",
      failed: "❌",
      canceled: "🚫",
    }[result] || "🔵";

    const message =
      `${emoji} **Pipeline ${result.charAt(0).toUpperCase() + result.slice(1)}**\n\n` +
      `🚀 **${pipelineName}** · Build #${buildNumber}\n` +
      `🌿 Branch: \`${branch}\`\n` +
      (repoName ? `📁 Repo: **${repoName}**\n` : "") +
      `👤 Triggered by: ${requestedBy}\n\n` +
      `🔗 [View Build](${buildUrl})`;

    // Send proactive message to all active conversations
    const refs = bot.getConversationReferences();
    console.log(`[webhook] Build #${buildNumber} ${result}. Notifying ${refs.size} conversation(s)...`);

    if (refs.size === 0) {
      console.warn("[webhook] No conversation references stored. Someone must message the bot first after deploy.");
    }

    for (const [convId, ref] of refs) {
      try {
        console.log(`[webhook] Sending to conversation: ${convId}`);
        await adapter.continueConversationAsync(
          process.env.BOT_APP_ID,
          ref,
          async (turnContext) => {
            await turnContext.sendActivity(message);
          },
        );
        console.log(`[webhook] Successfully notified: ${convId}`);
      } catch (err) {
        console.error(`[webhook] Failed to notify ${convId}:`, err.message);
      }
    }

    res.status(200).send("OK");
  } catch (err) {
    console.error("Webhook error:", err.message);
    res.status(200).send("OK");
  }
});

app.get("/", (req, res) => {
  res.send("Hello! This is the Azure DevOps Bot running.");
});

app.listen(process.env.PORT || 3978, () => {
  console.log(`Bot is running on port ${process.env.PORT || 3978}`);
});