"use strict";

const axios = require("axios");
const { MessageFactory, CardFactory } = require("botbuilder");
const { ORG, PROJECT, BASE_URL, HEADERS } = require("./adoConfig");
const { generateText } = require("./aiProvider");

// ── Step 1: Select a repository ──────────────────────────────────────
async function showRepoSelectorForPipeline(context, session, intent) {
  try {
    await context.sendActivity({ type: "typing" });

    const res = await axios.get(
      `${BASE_URL}/git/repositories?api-version=7.1`,
      { headers: HEADERS },
    );
    const repos = res.data.value;

    if (!repos?.length) {
      await context.sendActivity("No repositories found in this project.");
      return;
    }

    session.repos = repos;

    const label =
      intent === "run"
        ? "🚀 Run Pipeline"
        : intent === "logs"
          ? "📋 Pipeline Failure Logs"
          : "📊 Pipeline Runs";

    const card = {
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
      type: "AdaptiveCard",
      version: "1.5",
      body: [
        { type: "TextBlock", text: label, size: "Large", weight: "Bolder" },
        {
          type: "TextBlock",
          text: "Step 1: Select a repository",
          wrap: true,
          spacing: "Small",
        },
        {
          type: "Input.ChoiceSet",
          id: "repoId",
          style: "compact",
          placeholder: "Choose a repository...",
          choices: repos.map((r) => ({ title: r.name, value: r.id })),
        },
      ],
      actions: [
        {
          type: "Action.Submit",
          title: "Next →",
          data: { action: "pipeline_select_repo", intent },
        },
      ],
    };

    await context.sendActivity(
      MessageFactory.attachment(CardFactory.adaptiveCard(card)),
    );
  } catch (err) {
    console.error(
      "showRepoSelectorForPipeline error:",
      err.response?.data || err.message,
    );
    await context.sendActivity("Could not fetch repositories.");
  }
}

// ── Step 2: List pipelines for the selected repo ─────────────────────
async function showPipelinesForRepo(context, session, repoId, intent) {
  try {
    await context.sendActivity({ type: "typing" });

    const repo = (session.repos || []).find((r) => r.id === repoId);
    const repoName = repo?.name || repoId;
    session.pipelineFlow = { repoId, repoName, intent };

    const res = await axios.get(
      `${BASE_URL}/build/definitions?repositoryId=${repoId}&repositoryType=TfsGit&api-version=7.1`,
      { headers: HEADERS },
    );
    const pipelines = (res.data.value || []).slice(0, 20);

    if (!pipelines.length) {
      await context.sendActivity(
        `No pipelines found linked to **${repoName}**.`,
      );
      return;
    }

    session.pipelineFlow.pipelines = pipelines;

    const label =
      intent === "run"
        ? `🚀 Run Pipeline — ${repoName}`
        : intent === "logs"
          ? `📋 Failure Logs — ${repoName}`
          : `📊 Pipeline Runs — ${repoName}`;

    const card = {
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
      type: "AdaptiveCard",
      version: "1.5",
      body: [
        { type: "TextBlock", text: label, size: "Large", weight: "Bolder" },
        {
          type: "TextBlock",
          text: "Step 2: Select a pipeline",
          wrap: true,
          spacing: "Small",
        },
        {
          type: "Input.ChoiceSet",
          id: "pipelineId",
          style: "compact",
          placeholder: "Choose a pipeline...",
          choices: pipelines.map((p) => ({
            title: p.name,
            value: String(p.id),
          })),
        },
      ],
      actions: [
        {
          type: "Action.Submit",
          title: intent === "run" ? "Next →" : "View",
          data: { action: "pipeline_select_pipeline" },
        },
      ],
    };

    await context.sendActivity(
      MessageFactory.attachment(CardFactory.adaptiveCard(card)),
    );
  } catch (err) {
    console.error(
      "showPipelinesForRepo error:",
      err.response?.data || err.message,
    );
    await context.sendActivity("Could not fetch pipelines for this repo.");
  }
}

// ── Step 2→ dispatch: run needs branch, status/logs go straight ──────
async function handlePipelineSelected(context, session, pipelineId) {
  const flow = session.pipelineFlow;
  flow.pipelineId = parseInt(pipelineId, 10);
  const pipeline = (flow.pipelines || []).find(
    (p) => p.id === flow.pipelineId,
  );
  flow.pipelineName = pipeline?.name || pipelineId;

  if (flow.intent === "run") {
    await showBranchSelectorForRun(context, session);
  } else if (flow.intent === "status") {
    await showPipelineRuns(context, session);
  } else if (flow.intent === "logs") {
    await showFailedRunLogs(context, session);
  }
}

// ── Run: Step 3 — pick a branch ──────────────────────────────────────
async function showBranchSelectorForRun(context, session) {
  try {
    await context.sendActivity({ type: "typing" });
    const flow = session.pipelineFlow;

    const res = await axios.get(
      `${BASE_URL}/git/repositories/${flow.repoId}/refs?filter=heads&api-version=7.1`,
      { headers: HEADERS },
    );
    const branches = res.data.value || [];

    if (!branches.length) {
      await context.sendActivity("No branches found in this repository.");
      return;
    }

    flow.branches = branches;

    const card = {
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
      type: "AdaptiveCard",
      version: "1.5",
      body: [
        {
          type: "TextBlock",
          text: `🚀 Run — ${flow.pipelineName}`,
          size: "Large",
          weight: "Bolder",
        },
        {
          type: "TextBlock",
          text: `Repo: **${flow.repoName}**`,
          wrap: true,
          spacing: "Small",
        },
        {
          type: "TextBlock",
          text: "Step 3: Select the branch to build",
          wrap: true,
          spacing: "Small",
        },
        {
          type: "Input.ChoiceSet",
          id: "branch",
          style: "compact",
          placeholder: "Choose branch...",
          choices: branches.map((b) => {
            const name = b.name.replace("refs/heads/", "");
            return { title: name, value: name };
          }),
        },
      ],
      actions: [
        {
          type: "Action.Submit",
          title: "Run Pipeline 🚀",
          data: { action: "pipeline_run" },
        },
      ],
    };

    await context.sendActivity(
      MessageFactory.attachment(CardFactory.adaptiveCard(card)),
    );
  } catch (err) {
    console.error(
      "showBranchSelectorForRun error:",
      err.response?.data || err.message,
    );
    await context.sendActivity("Could not fetch branches.");
  }
}

// ── Run: trigger the pipeline ────────────────────────────────────────
async function triggerPipelineRun(context, session, branch) {
  try {
    const flow = session.pipelineFlow;
    await context.sendActivity({ type: "typing" });
    await context.sendActivity(
      `🚀 Triggering **${flow.pipelineName}** on branch \`${branch}\`...`,
    );

    const res = await axios.post(
      `${BASE_URL}/build/builds?api-version=7.1`,
      {
        definition: { id: flow.pipelineId },
        sourceBranch: `refs/heads/${branch}`,
      },
      { headers: HEADERS },
    );

    const build = res.data;
    const buildUrl = build._links?.web?.href ||
      `https://dev.azure.com/${ORG}/${PROJECT}/_build/results?buildId=${build.id}`;

    session.pipelineFlow = {};

    await context.sendActivity(
      `✅ **Pipeline Run Queued!**\n\n` +
        `🔗 [View Build #${build.id}](${buildUrl})\n` +
        `📌 Pipeline: **${flow.pipelineName}**\n` +
        `🌿 Branch: \`${branch}\`\n` +
        `📁 Repo: **${flow.repoName}**`,
    );
  } catch (err) {
    console.error(
      "triggerPipelineRun error:",
      JSON.stringify(err.response?.data || err.message, null, 2),
    );
    await context.sendActivity(
      "Could not trigger pipeline run. Check that the PAT has **Build (Read & Execute)** permission.",
    );
  }
}

// ── Status: show recent runs for the selected pipeline ───────────────
async function showPipelineRuns(context, session) {
  try {
    const flow = session.pipelineFlow;
    await context.sendActivity({ type: "typing" });

    const res = await axios.get(
      `${BASE_URL}/build/builds?definitions=${flow.pipelineId}&$top=10&api-version=7.1`,
      { headers: HEADERS },
    );
    const builds = res.data.value || [];

    if (!builds.length) {
      await context.sendActivity(
        `No runs found for **${flow.pipelineName}**.`,
      );
      return;
    }

    const statusEmoji = (r) =>
      ({
        succeeded: "✅",
        partiallySucceeded: "⚠️",
        failed: "❌",
        canceled: "🚫",
      })[r] || "🔵";

    const statusColor = (r) =>
      ({
        succeeded: "Good",
        partiallySucceeded: "Warning",
        failed: "Attention",
        canceled: "Default",
      })[r] || "Accent";

    const card = {
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
      type: "AdaptiveCard",
      version: "1.5",
      body: [
        {
          type: "TextBlock",
          text: `📊 Runs — ${flow.pipelineName}`,
          size: "Large",
          weight: "Bolder",
        },
        {
          type: "TextBlock",
          text: `Repo: **${flow.repoName}** · Showing last ${builds.length} run(s)`,
          isSubtle: true,
          size: "Small",
          spacing: "None",
          wrap: true,
        },
        ...builds.map((b) => {
          const branch = (b.sourceBranch || "").replace("refs/heads/", "");
          const result = b.result || b.status || "inProgress";
          const date = new Date(b.queueTime).toLocaleDateString();
          const duration =
            b.startTime && b.finishTime
              ? formatDuration(
                  new Date(b.finishTime) - new Date(b.startTime),
                )
              : "running";
          const buildUrl =
            b._links?.web?.href ||
            `https://dev.azure.com/${ORG}/${PROJECT}/_build/results?buildId=${b.id}`;

          return {
            type: "ColumnSet",
            separator: true,
            selectAction: { type: "Action.OpenUrl", url: buildUrl },
            columns: [
              {
                type: "Column",
                width: "stretch",
                items: [
                  {
                    type: "TextBlock",
                    text: `${statusEmoji(result)} **#${b.id}** · \`${branch}\``,
                    wrap: true,
                    size: "Small",
                    weight: "Bolder",
                  },
                  {
                    type: "TextBlock",
                    text: `👤 ${b.requestedFor?.displayName || "Unknown"} · 📅 ${date} · ⏱ ${duration}`,
                    isSubtle: true,
                    size: "Small",
                    spacing: "None",
                  },
                ],
              },
              {
                type: "Column",
                width: "auto",
                items: [
                  {
                    type: "TextBlock",
                    text: capitalize(result),
                    size: "Small",
                    color: statusColor(result),
                  },
                ],
              },
            ],
          };
        }),
      ],
    };

    session.pipelineFlow = {};

    await context.sendActivity(
      MessageFactory.attachment(CardFactory.adaptiveCard(card)),
    );
  } catch (err) {
    console.error(
      "showPipelineRuns error:",
      err.response?.data || err.message,
    );
    await context.sendActivity("Could not fetch pipeline runs.");
  }
}

// ── Logs: fetch last failed run, extract errors, AI summary ──────────
async function showFailedRunLogs(context, session) {
  try {
    const flow = session.pipelineFlow;
    await context.sendActivity({ type: "typing" });
    await context.sendActivity(
      `🔍 Looking for failed runs of **${flow.pipelineName}**...`,
    );

    // Find the most recent failed build
    const buildsRes = await axios.get(
      `${BASE_URL}/build/builds?definitions=${flow.pipelineId}&resultFilter=failed&$top=1&api-version=7.1`,
      { headers: HEADERS },
    );
    const failedBuild = buildsRes.data.value?.[0];

    if (!failedBuild) {
      await context.sendActivity(
        `No failed runs found for **${flow.pipelineName}**. 🎉`,
      );
      session.pipelineFlow = {};
      return;
    }

    const buildUrl =
      failedBuild._links?.web?.href ||
      `https://dev.azure.com/${ORG}/${PROJECT}/_build/results?buildId=${failedBuild.id}`;

    await context.sendActivity(
      `❌ Found failed build **#${failedBuild.id}** (${new Date(failedBuild.queueTime).toLocaleDateString()}). Fetching logs...`,
    );
    await context.sendActivity({ type: "typing" });

    // Get timeline to find failed tasks
    const timelineRes = await axios.get(
      `${BASE_URL}/build/builds/${failedBuild.id}/timeline?api-version=7.1`,
      { headers: HEADERS },
    );
    const failedTasks = (timelineRes.data.records || []).filter(
      (r) => r.result === "failed" && r.log,
    );

    let errorLogs = "";
    for (const task of failedTasks.slice(0, 3)) {
      try {
        const logRes = await axios.get(task.log.url, { headers: HEADERS });
        const logText =
          typeof logRes.data === "string"
            ? logRes.data
            : JSON.stringify(logRes.data);
        // Keep last 150 lines per task to stay within token limits
        const lines = logText.split("\n");
        const tail = lines.slice(-150).join("\n");
        errorLogs += `\n--- Task: ${task.name} ---\n${tail}\n`;
      } catch {
        errorLogs += `\n--- Task: ${task.name} --- (could not fetch log)\n`;
      }
    }

    if (!errorLogs.trim()) {
      await context.sendActivity(
        `Could not extract error logs from build #${failedBuild.id}. [View in ADO](${buildUrl})`,
      );
      session.pipelineFlow = {};
      return;
    }

    // Truncate for display
    const logPreview =
      errorLogs.length > 2000
        ? errorLogs.substring(errorLogs.length - 2000)
        : errorLogs;

    await context.sendActivity(
      `📋 **Error Logs — Build #${failedBuild.id}:**\n\n\`\`\`\n${logPreview}\n\`\`\``,
    );

    // AI summary
    await context.sendActivity("🤖 Asking AI to diagnose the failure...");
    await context.sendActivity({ type: "typing" });

    const prompt = `You are a senior DevOps engineer. Analyze the following CI/CD pipeline failure logs and provide:

1. **Root Cause** — What specifically failed and why.
2. **Fix** — Concrete steps to resolve the issue.
3. **Prevention** — How to avoid this in the future.

Pipeline: ${flow.pipelineName}
Repository: ${flow.repoName}
Build #${failedBuild.id}

Logs:
${errorLogs.substring(0, 6000)}

Be concise and actionable.`;

    try {
      const aiResponse = await generateText(prompt);
      await context.sendActivity(
        `🤖 **AI Diagnosis — Build #${failedBuild.id}:**\n\n${aiResponse}`,
      );
    } catch (aiErr) {
      console.warn("AI unavailable for log analysis:", aiErr.message);
      await context.sendActivity(
        "AI analysis unavailable. Review the logs above manually.",
      );
    }

    await context.sendActivity(`🔗 [View full build in ADO](${buildUrl})`);
    session.pipelineFlow = {};
  } catch (err) {
    console.error(
      "showFailedRunLogs error:",
      err.response?.data || err.message,
    );
    await context.sendActivity("Could not fetch failure logs.");
  }
}

// ── Helpers ──────────────────────────────────────────────────────────
function formatDuration(ms) {
  const totalSec = Math.floor(ms / 1000);
  const m = Math.floor(totalSec / 60);
  const s = totalSec % 60;
  return m > 0 ? `${m}m ${s}s` : `${s}s`;
}

function capitalize(s) {
  return s ? s.charAt(0).toUpperCase() + s.slice(1) : "";
}

module.exports = {
  showRepoSelectorForPipeline,
  showPipelinesForRepo,
  handlePipelineSelected,
  showBranchSelectorForRun,
  triggerPipelineRun,
  showPipelineRuns,
  showFailedRunLogs,
};
