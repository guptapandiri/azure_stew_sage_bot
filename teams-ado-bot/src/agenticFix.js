"use strict";

const axios = require("axios");
const { MessageFactory, CardFactory } = require("botbuilder");
const { generateText } = require("./aiProvider");
const { ORG, PROJECT, BASE_URL, HEADERS } = require("./adoConfig");

async function showRepoSelectorForFix(context, bugDetails, repos) {
  const card = {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      { type: "TextBlock", text: "🔧 AI Code Fix — Select Repository", size: "Large", weight: "Bolder" },
      { type: "TextBlock", text: `Bug #${bugDetails.id}: ${bugDetails.title}`, wrap: true, spacing: "Small" },
      { type: "TextBlock", text: "Which repository contains the code for this bug?", wrap: true, spacing: "Small" },
      {
        type: "Input.ChoiceSet",
        id: "repoId",
        style: "compact",
        placeholder: "Choose a repository...",
        choices: repos.map((r) => ({ title: r.name, value: r.id })),
      },
    ],
    actions: [
      { type: "Action.Submit", title: "Generate Fix 🤖", data: { action: "select_repo_for_fix" } },
    ],
  };
  await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
}

async function continueFixBugFlow(context, session, repoId) {
  const repo = (session.repos || []).find((r) => r.id === repoId);
  if (!repo || !session.pendingFixBug) {
    await context.sendActivity("Session expired. Please type `show bugs` and try again.");
    return;
  }
  await runAgenticFixFlow(context, session, repo, session.pendingFixBug);
}

async function runAgenticFixFlow(context, session, repo, bugDetails) {
  try {
    await context.sendActivity(`📂 Scanning **${repo.name}** for relevant files...`);
    await context.sendActivity({ type: "typing" });

    const defaultBranch = (repo.defaultBranch || "refs/heads/main").replace("refs/heads/", "");
    const filePaths = await fetchRepoTree(repo.id, defaultBranch);

    if (!filePaths.length) {
      await context.sendActivity("Repository appears to be empty. Cannot generate a fix.");
      return;
    }

    await context.sendActivity(`🤖 **Step 1/2** — Identifying relevant files (${filePaths.length} total)...`);
    await context.sendActivity({ type: "typing" });

    let selectedPaths;
    try {
      selectedPaths = await selectFilesToInvestigate(bugDetails, filePaths);
    } catch (aiErr) {
      console.error("selectFilesToInvestigate failed:", aiErr.message);
      await context.sendActivity(`❌ AI file selection failed: ${aiErr.message}\n\nCheck that \`GOOGLE_API_KEY\` / \`OPENAI_API_KEY\` is set correctly in Railway.`);
      return;
    }

    if (!selectedPaths.length) {
      await context.sendActivity(
        "AI could not identify relevant files for this bug. Try adding more detail to the bug description.",
      );
      return;
    }

    const pathList = selectedPaths.map((p) => `\`${p}\``).join(", ");
    await context.sendActivity(`📄 Fetching ${selectedPaths.length} file(s): ${pathList}`);
    await context.sendActivity({ type: "typing" });

    const files = await fetchFileContents(repo.id, selectedPaths, defaultBranch);
    if (!files.length) {
      await context.sendActivity("Could not read the selected files. Please try again.");
      return;
    }

    await context.sendActivity("🤖 **Step 2/2** — Generating code fix...");
    await context.sendActivity({ type: "typing" });

    let fixedFiles;
    try {
      fixedFiles = await generateCodeFix(bugDetails, files);
    } catch (aiErr) {
      console.error("generateCodeFix failed:", aiErr.message);
      await context.sendActivity(`❌ AI code generation failed: ${aiErr.message}`);
      return;
    }

    if (!fixedFiles.length) {
      await context.sendActivity(
        "AI could not generate a code fix for this bug. Please fix it manually or refine the bug description.",
      );
      return;
    }

    session.fixedFiles = fixedFiles;
    session.fixRepo = repo;

    const fileList = fixedFiles.map((f) => `• \`${f.path}\``).join("\n");
    await context.sendActivity(`✅ **AI Code Fix Ready** for Bug #${bugDetails.id}:\n\n${fileList}\n\n🚀 Raising PR automatically...`);

    await createFixPR(context, session);
  } catch (err) {
    console.error("runAgenticFixFlow error:", err.response?.data || err.message);
    await context.sendActivity("Failed to generate code fix. Please try again.");
  }
}

async function createFixPR(context, session) {
  const { scheduledBug, fixedFiles, fixRepo } = session;

  await context.sendActivity({ type: "typing" });
  await context.sendActivity("📂 Preparing branch and commit...");

  const defaultBranch = fixRepo.defaultBranch || "refs/heads/main";
  const refsRes = await axios.get(
    `${BASE_URL}/git/repositories/${fixRepo.id}/refs?filter=heads&api-version=7.1`,
    { headers: HEADERS },
  );
  const mainRef =
    refsRes.data.value.find((r) => r.name === defaultBranch) ||
    refsRes.data.value[0];
  if (!mainRef) {
    await context.sendActivity("Could not find the default branch.");
    return;
  }

  const suffix = Date.now().toString(36);
  const branchName = `ai-fix/bug-${scheduledBug.id}-${suffix}`;

  await context.sendActivity(`🌿 Creating branch \`${branchName}\`...`);

  await axios.post(
    `${BASE_URL}/git/repositories/${fixRepo.id}/pushes?api-version=7.1`,
    {
      refUpdates: [{ name: `refs/heads/${branchName}`, oldObjectId: mainRef.objectId }],
      commits: [{
        comment: `AI fix for Bug #${scheduledBug.id}: ${scheduledBug.fields["System.Title"]}`,
        changes: fixedFiles.map((f) => ({
          changeType: "edit",
          item: { path: f.path },
          newContent: {
            content: Buffer.from(f.content).toString("base64"),
            contentType: "base64encoded",
          },
        })),
      }],
    },
    { headers: HEADERS },
  );

  await context.sendActivity("🔀 Creating Pull Request...");

  const fileList = fixedFiles.map((f) => `- \`${f.path}\``).join("\n");
  const prRes = await axios.post(
    `${BASE_URL}/git/repositories/${fixRepo.id}/pullrequests?api-version=7.1`,
    {
      title: `🤖 AI Fix: Bug #${scheduledBug.id} — ${scheduledBug.fields["System.Title"]}`,
      description: `## AI Generated Code Fix\n\nThis PR was automatically generated by StewSage.\n\n**Linked Bug:** #${scheduledBug.id} — ${scheduledBug.fields["System.Title"]}\n\n**Files Modified:**\n${fileList}`,
      sourceRefName: `refs/heads/${branchName}`,
      targetRefName: defaultBranch,
      workItemRefs: [{ id: String(scheduledBug.id) }],
    },
    { headers: HEADERS },
  );

  const prUrl = `https://dev.azure.com/${ORG}/${PROJECT}/_git/${fixRepo.name}/pullrequest/${prRes.data.pullRequestId}`;
  session.scheduledBug = null;
  session.fixedFiles = null;
  session.fixRepo = null;
  session.pendingFixBug = null;

  await context.sendActivity(
    `✅ **PR Created Successfully!**\n\n` +
      `🔗 [View PR #${prRes.data.pullRequestId}](${prUrl})\n` +
      `📌 \`${branchName}\` → \`${defaultBranch.replace("refs/heads/", "")}\`\n` +
      `📁 Repo: **${fixRepo.name}**\n` +
      `🐛 Linked to Bug #${scheduledBug.id}`,
  );
}

async function fetchRepoTree(repoId, branchName) {
  const res = await axios.get(
    `${BASE_URL}/git/repositories/${repoId}/items?recursionLevel=Full&versionDescriptor.version=${encodeURIComponent(branchName)}&versionDescriptor.versionType=branch&api-version=7.1`,
    { headers: HEADERS },
  );
  const items = res.data.value || [];
  const SKIP_EXTS = /\.(png|jpg|jpeg|gif|svg|ico|woff|woff2|ttf|eot|pdf|zip|tar|gz|lock|min\.js|min\.css|map)$/i;
  return items
    .filter((item) => item.gitObjectType === "blob" && !SKIP_EXTS.test(item.path))
    .map((item) => item.path);
}

async function selectFilesToInvestigate(bugDetails, filePaths) {
  const prompt = `You are a senior software engineer. Given the following bug and a repository file list, identify the ≤5 files most likely to contain the root cause.

Bug ID: ${bugDetails.id}
Title: ${bugDetails.title}
Type: ${bugDetails.type}
Description: ${bugDetails.description}
${bugDetails.reproSteps ? `Repro Steps: ${bugDetails.reproSteps}` : ""}
${bugDetails.acceptanceCriteria ? `Acceptance Criteria: ${bugDetails.acceptanceCriteria}` : ""}

Repository files:
${filePaths.join("\n")}

Respond with ONLY a JSON array of file paths (no markdown, no explanation):
["path/to/file1.js", "path/to/file2.js"]

Select at most 5 files. Return [] if no files seem relevant.`;

  const response = await generateText(prompt);
  return parseJsonFromLlm(response);
}

async function fetchFileContents(repoId, filePaths, branchName) {
  const files = [];
  for (const path of filePaths) {
    try {
      const res = await axios.get(
        `${BASE_URL}/git/repositories/${repoId}/items?path=${encodeURIComponent(path)}&versionDescriptor.version=${encodeURIComponent(branchName)}&versionDescriptor.versionType=branch&$format=text&api-version=7.1`,
        { headers: HEADERS, responseType: "text" },
      );
      const content = typeof res.data === "string" ? res.data : JSON.stringify(res.data);
      files.push({ path, content });
    } catch (err) {
      console.warn(`Could not fetch file ${path}:`, err.response?.data || err.message);
    }
  }
  return files;
}

async function generateCodeFix(bugDetails, files) {
  const filesSection = files
    .map((f) => `### ${f.path}\n\`\`\`\n${f.content}\n\`\`\``)
    .join("\n\n");

  const prompt = `You are a senior software engineer. Fix the following bug by modifying the relevant source files.

Bug ID: ${bugDetails.id}
Title: ${bugDetails.title}
Type: ${bugDetails.type}
Description: ${bugDetails.description}
${bugDetails.reproSteps ? `Repro Steps: ${bugDetails.reproSteps}` : ""}
${bugDetails.acceptanceCriteria ? `Acceptance Criteria: ${bugDetails.acceptanceCriteria}` : ""}

Source files:
${filesSection}

Respond with ONLY a JSON array of modified files. Each element must have "path" (exact original path) and "content" (complete updated file content). Only include files that actually need changes.

[
  {"path": "src/utils/validation.js", "content": "...complete file content..."}
]

Return [] if no code changes are needed.`;

  const response = await generateText(prompt);
  return parseJsonFromLlm(response);
}

function parseJsonFromLlm(text) {
  try {
    const match = text.match(/\[[\s\S]*\]/);
    if (!match) return [];
    return JSON.parse(match[0]);
  } catch (err) {
    console.warn("Failed to parse LLM JSON response:", err.message);
    return [];
  }
}

module.exports = {
  showRepoSelectorForFix,
  continueFixBugFlow,
  runAgenticFixFlow,
  fetchRepoTree,
  selectFilesToInvestigate,
  fetchFileContents,
  generateCodeFix,
  parseJsonFromLlm,
};
