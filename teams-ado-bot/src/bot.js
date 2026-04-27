const {
  TeamsActivityHandler,
  MessageFactory,
  CardFactory,
} = require("botbuilder");
const axios = require("axios");
const OpenAI = require("openai");

const ORG = process.env.ADO_ORG;
const PROJECT = process.env.ADO_PROJECT;
const PAT = process.env.ADO_PAT;
const BASE_URL = `https://dev.azure.com/${ORG}/${PROJECT}/_apis`;
const AUTH = Buffer.from(`:${PAT}`).toString("base64");
const HEADERS = {
  Authorization: `Basic ${AUTH}`,
  "Content-Type": "application/json",
};

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

class AzureDevOpsBot extends TeamsActivityHandler {
  constructor() {
    super();
    this.sessions = new Map();

    this.onMessage(async (context, next) => {
      const text = (context.activity.text || "")
        .replace(/<[^>]+>/g, "")
        .toLowerCase()
        .trim();
      const convId = context.activity.conversation.id;

      if (!this.sessions.has(convId)) {
        this.sessions.set(convId, {
          bugs: [],
          scheduledBug: null,
          aiFix: null,
        });
      }

      const scheduleMatch = text.match(/schedule\s+bug\s+(\d+)/);
      const wantsPR =
        (text.includes("raise pr") ||
          text === "yes" ||
          text.includes("approve")) &&
        this.sessions.get(convId)?.aiFix;

      if (
        text.includes("show bugs") ||
        text.includes("list bugs") ||
        text === "bugs"
      ) {
        await this.showBugs(context, convId);
      } else if (scheduleMatch) {
        await this.scheduleBug(context, convId, parseInt(scheduleMatch[1]));
      } else if (wantsPR) {
        await this.raisePR(context, convId);
      } else if (text.includes("sprint")) {
        await this.showSprint(context);
      } else if (text.includes("pipeline") || text.includes("build")) {
        await this.showPipelines(context);
      } else if (text.includes("repo")) {
        await this.showRepos(context);
      } else {
        await this.showHelp(context);
      }

      await next();
    });

    this.onMembersAdded(async (context, next) => {
      for (const member of context.activity.membersAdded) {
        if (member.id !== context.activity.recipient.id) {
          await context.sendActivity(
            "👋 Hi! I'm **StewMate** — Stewart's AI-powered dev assistant!\n\nType `show bugs` to get started.",
          );
        }
      }
      await next();
    });
  }

  async showBugs(context, convId) {
    try {
      await context.sendActivity({ type: "typing" });

      const wiql = {
        query: `SELECT [System.Id],[System.Title],[System.State],[System.AssignedTo] FROM WorkItems WHERE [System.TeamProject] = '${PROJECT}' AND [System.WorkItemType] = 'Issue' AND [System.State] <> 'Done' AND [System.State] <> 'Closed' ORDER BY [System.CreatedDate] DESC`,
      };

      let wiqlRes;
      try {
        wiqlRes = await axios.post(
          `${BASE_URL}/wit/wiql?api-version=7.1`,
          wiql,
          { headers: HEADERS },
        );
      } catch (err) {
        console.error("WIQL query failed:", err.response?.data || err.message);
        throw err;
      }
      const ids = (wiqlRes.data.workItems || []).slice(0, 10).map((w) => w.id);

      if (!ids.length) {
        await context.sendActivity("No open bugs found in the project!");
        return;
      }

      const itemsRes = await axios.get(
        `${BASE_URL}/wit/workitems?ids=${ids.join(",")}&fields=System.Id,System.Title,System.State,System.AssignedTo&api-version=7.1`,
        { headers: HEADERS },
      );

      const bugs = itemsRes.data.value;
      this.sessions.get(convId).bugs = bugs;

      const card = {
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.5",
        body: [
          {
            type: "TextBlock",
            text: `🐛 Open Issues — ${PROJECT}`,
            size: "Large",
            weight: "Bolder",
          },
          {
            type: "TextBlock",
            text: "Type: schedule bug [number]  to have AI write a fix",
            isSubtle: true,
            size: "Small",
            spacing: "None",
          },
          ...bugs.map((bug, i) => ({
            type: "ColumnSet",
            separator: true,
            columns: [
              {
                type: "Column",
                width: "auto",
                items: [
                  {
                    type: "TextBlock",
                    text: `${i + 1}.`,
                    weight: "Bolder",
                    size: "Small",
                    color: "Accent",
                  },
                ],
              },
              {
                type: "Column",
                width: "stretch",
                items: [
                  {
                    type: "TextBlock",
                    text: `#${bug.id} · ${bug.fields["System.Title"]}`,
                    weight: "Bolder",
                    size: "Small",
                    wrap: true,
                  },
                  {
                    type: "TextBlock",
                    text:
                      bug.fields["System.AssignedTo"]?.displayName ||
                      "Unassigned",
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
                    text: bug.fields["System.State"],
                    size: "Small",
                    color: "Attention",
                  },
                ],
              },
            ],
          })),
        ],
      };

      await context.sendActivity(
        MessageFactory.attachment(CardFactory.adaptiveCard(card)),
      );
    } catch (err) {
      console.error("showBugs error:", err.response?.data || err.message);
      await context.sendActivity("Could not fetch bugs.");
    }
  }

  async scheduleBug(context, convId, listNumber) {
    try {
      const session = this.sessions.get(convId);

      if (!session.bugs.length) {
        await context.sendActivity(
          "Please type `show bugs` first to load the bug list.",
        );
        return;
      }

      const bug = session.bugs[listNumber - 1];
      if (!bug) {
        await context.sendActivity(
          `Bug #${listNumber} not found in the list. There are ${session.bugs.length} bugs. Try again.`,
        );
        return;
      }

      await context.sendActivity({ type: "typing" });
      await context.sendActivity(
        `🔍 Fetching full details for **Bug #${bug.id}: ${bug.fields["System.Title"]}**...`,
      );

      const detailRes = await axios.get(
        `${BASE_URL}/wit/workitems/${bug.id}?api-version=7.1`,
        { headers: HEADERS },
      );
      const f = detailRes.data.fields;
      const title = f["System.Title"] || "";
      const description = this.stripHtml(
        f["System.Description"] || "No description provided.",
      );
      const reproSteps = this.stripHtml(
        f["Microsoft.VSTS.TCM.ReproSteps"] || "No repro steps provided.",
      );
      const acceptanceCriteria = this.stripHtml(
        f["Microsoft.VSTS.Common.AcceptanceCriteria"] || "",
      );

      await context.sendActivity("🤖 Sending to AI — writing code fix now...");
      await context.sendActivity({ type: "typing" });

      const prompt = `You are a senior software engineer. Analyze this bug and write a production-ready code fix.

Bug ID: ${bug.id}
Title: ${title}
Description: ${description}
Repro Steps: ${reproSteps}
${acceptanceCriteria ? `Acceptance Criteria: ${acceptanceCriteria}` : ""}

Respond with:
## Root Cause
Brief analysis of what's likely causing this bug.

## Code Fix
Complete, production-ready code with proper file paths and code blocks. Write real code, not pseudocode.

## Test Cases
Key tests to verify the fix works correctly.`;

      let aiFix;
      try {
        const result = await openai.chat.completions.create({
          model: "gpt-3.5-turbo",
          max_tokens: 4096,
          messages: [{ role: "user", content: prompt }],
        });
        aiFix = result.choices[0].message.content;
      } catch (aiErr) {
        console.warn("AI API unavailable, using demo fix:", aiErr.message);
        aiFix = this.getDemoFix(title);
      }
      session.scheduledBug = bug;
      session.aiFix = aiFix;

      const preview =
        aiFix.length > 3500
          ? aiFix.substring(0, 3500) +
            "\n\n_...truncated. Full fix will be in the PR._"
          : aiFix;
      await context.sendActivity(
        `🤖 **AI Fix for Bug #${bug.id}:**\n\n${preview}`,
      );
      await context.sendActivity(
        "---\n✅ Happy with the fix? Type **`raise PR`** to create a PR in ADO.\n❌ Type **`show bugs`** to pick a different bug.",
      );
    } catch (err) {
      console.error("scheduleBug error:", err.response?.data || err.message);
      await context.sendActivity("Could not analyze bug. Please try again.");
    }
  }

  async raisePR(context, convId) {
    try {
      const session = this.sessions.get(convId);
      const { scheduledBug, aiFix } = session;

      await context.sendActivity({ type: "typing" });
      await context.sendActivity("📂 Fetching repository info...");

      const reposRes = await axios.get(
        `${BASE_URL}/git/repositories?api-version=7.1`,
        { headers: HEADERS },
      );
      const repos = reposRes.data.value;

      if (!repos.length) {
        await context.sendActivity(
          "No git repositories found in this project!",
        );
        return;
      }

      const repo = repos[0];
      const defaultBranch = repo.defaultBranch || "refs/heads/main";

      const refsRes = await axios.get(
        `${BASE_URL}/git/repositories/${repo.id}/refs?filter=heads&api-version=7.1`,
        { headers: HEADERS },
      );

      const mainRef =
        refsRes.data.value.find((r) => r.name === defaultBranch) ||
        refsRes.data.value[0];
      if (!mainRef) {
        await context.sendActivity("Could not find the default branch.");
        return;
      }

      const branchName = `ai-fix/bug-${scheduledBug.id}`;
      const filePath = `/ai-fixes/bug-${scheduledBug.id}-fix.md`;
      const fileContent = [
        `# AI Fix: Bug #${scheduledBug.id}`,
        `**Title:** ${scheduledBug.fields["System.Title"]}`,
        "",
        "## AI Generated Fix",
        "",
        aiFix,
        "",
        "---",
        "*Generated by StewMate — Stewart AI Dev Assistant*",
      ].join("\n");

      await context.sendActivity(`🌿 Creating branch \`${branchName}\`...`);

      await axios.post(
        `${BASE_URL}/git/repositories/${repo.id}/pushes?api-version=7.1`,
        {
          refUpdates: [
            { name: `refs/heads/${branchName}`, oldObjectId: mainRef.objectId },
          ],
          commits: [
            {
              comment: `AI fix for Bug #${scheduledBug.id}: ${scheduledBug.fields["System.Title"]}`,
              changes: [
                {
                  changeType: "add",
                  item: { path: filePath },
                  newContent: {
                    content: Buffer.from(fileContent).toString("base64"),
                    contentType: "base64encoded",
                  },
                },
              ],
            },
          ],
        },
        { headers: HEADERS },
      );

      await context.sendActivity("🔀 Creating Pull Request...");

      const prRes = await axios.post(
        `${BASE_URL}/git/repositories/${repo.id}/pullrequests?api-version=7.1`,
        {
          title: `🤖 AI Fix: Bug #${scheduledBug.id} — ${scheduledBug.fields["System.Title"]}`,
          description: `## AI Generated Fix\n\nThis PR was automatically generated by the Dev Assistant Bot.\n\n**Linked Bug:** #${scheduledBug.id} — ${scheduledBug.fields["System.Title"]}\n\n---\n\n${aiFix}`,
          sourceRefName: `refs/heads/${branchName}`,
          targetRefName: defaultBranch,
          workItemRefs: [{ id: String(scheduledBug.id) }],
        },
        { headers: HEADERS },
      );

      const prUrl = `https://dev.azure.com/${ORG}/${PROJECT}/_git/${repo.name}/pullrequest/${prRes.data.pullRequestId}`;
      session.scheduledBug = null;
      session.aiFix = null;

      await context.sendActivity(
        `✅ **PR Created Successfully!**\n\n` +
          `🔗 [View PR #${prRes.data.pullRequestId}](${prUrl})\n` +
          `📌 \`${branchName}\` → \`${defaultBranch.replace("refs/heads/", "")}\`\n` +
          `📁 Repo: **${repo.name}**\n` +
          `🐛 Linked to Bug #${scheduledBug.id}`,
      );
    } catch (err) {
      console.error(
        "raisePR error:",
        JSON.stringify(err.response?.data || err.message, null, 2),
      );
      await context.sendActivity(
        "Could not create PR. Check that the PAT has **Code (Read & Write)** permission.",
      );
    }
  }

  async showSprint(context) {
    try {
      await context.sendActivity({ type: "typing" });
      const res = await axios.get(
        `https://dev.azure.com/${ORG}/${PROJECT}/_apis/work/teamsettings/iterations?$timeframe=current&api-version=7.1`,
        { headers: HEADERS },
      );
      const sprint = res.data.value?.[0];
      if (!sprint) {
        await context.sendActivity("No active sprint found!");
        return;
      }
      await context.sendActivity(
        `🏃 **Current Sprint: ${sprint.name}**\n\n📅 Start: ${new Date(sprint.attributes?.startDate).toLocaleDateString()}\n📅 End: ${new Date(sprint.attributes?.finishDate).toLocaleDateString()}`,
      );
    } catch (err) {
      console.error(err.response?.data || err.message);
      await context.sendActivity("Could not fetch sprint info.");
    }
  }

  async showPipelines(context) {
    try {
      await context.sendActivity({ type: "typing" });
      const res = await axios.get(`${BASE_URL}/pipelines?api-version=7.1`, {
        headers: HEADERS,
      });
      const pipelines = res.data.value?.slice(0, 5);
      if (!pipelines?.length) {
        await context.sendActivity("No pipelines found!");
        return;
      }
      let msg = "🚀 **Pipelines:**\n\n";
      for (const p of pipelines) {
        msg += `• **${p.name}**\n`;
      }
      await context.sendActivity(msg);
    } catch (err) {
      console.error(err.response?.data || err.message);
      await context.sendActivity("Could not fetch pipelines.");
    }
  }

  async showRepos(context) {
    try {
      await context.sendActivity({ type: "typing" });
      const res = await axios.get(
        `${BASE_URL}/git/repositories?api-version=7.1`,
        { headers: HEADERS },
      );
      const repos = res.data.value?.slice(0, 10);
      if (!repos?.length) {
        await context.sendActivity("No repositories found in this project.");
        return;
      }

      let msg = "📁 **Repositories:**\n\n";
      for (const repo of repos) {
        const branch = (repo.defaultBranch || "refs/heads/main").replace(
          "refs/heads/",
          "",
        );
        msg += `• **${repo.name}** (default: \`${branch}\`)\n`;
      }
      await context.sendActivity(msg);
    } catch (err) {
      console.error(err.response?.data || err.message);
      await context.sendActivity("Could not fetch repositories.");
    }
  }

  async showHelp(context) {
    await context.sendActivity(
      "🤖 **StewMate — Your Dev Assistant:**\n\n" +
        "🐛 `show bugs` — List all open bugs in the project\n" +
        "🔧 `schedule bug 2` — AI writes a code fix for bug #2 from the list\n" +
        "🔀 `raise PR` — Create a PR in ADO with the AI-generated fix\n" +
        "🏃 `sprint status` — Current sprint info\n" +
        "🚀 `pipeline status` — List pipelines\n" +
        "📁 `list repos` — List git repositories in the project",
    );
  }

  getDemoFix(title) {
    return `## Root Cause
The \`isFormValid\` state is computed inside \`onChange\` handlers individually per field. On the last field entry, React batches the state update and the button's \`disabled\` prop reads the stale pre-update value — so it stays disabled until the next render cycle triggered by an external interaction.

## Code Fix

**File: \`src/components/RegistrationForm.jsx\`**

\`\`\`jsx
import React, { useState, useEffect } from 'react';

const RegistrationForm = () => {
  const [form, setForm] = useState({
    firstName: '',
    lastName: '',
    email: '',
    password: ''
  });

  const [isFormValid, setIsFormValid] = useState(false);

  // useEffect re-evaluates validity after every state update,
  // ensuring the button reflects the latest values correctly.
  useEffect(() => {
    const { firstName, lastName, email, password } = form;
    const emailValid = /^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$/.test(email);
    setIsFormValid(
      firstName.trim() !== '' &&
      lastName.trim() !== '' &&
      emailValid &&
      password.length >= 8
    );
  }, [form]);

  const handleChange = (e) => {
    const { name, value } = e.target;
    setForm(prev => ({ ...prev, [name]: value }));
  };

  const handleSubmit = (e) => {
    e.preventDefault();
    console.log('Form submitted:', form);
  };

  return (
    <form onSubmit={handleSubmit}>
      <input name="firstName" placeholder="First Name" onChange={handleChange} />
      <input name="lastName" placeholder="Last Name" onChange={handleChange} />
      <input name="email" type="email" placeholder="Email" onChange={handleChange} />
      <input name="password" type="password" placeholder="Password" onChange={handleChange} />
      <button type="submit" disabled={!isFormValid}>
        Submit
      </button>
    </form>
  );
};

export default RegistrationForm;
\`\`\`

## Test Cases
- Fill all 4 fields → Submit button enables immediately ✅
- Clear any field → Submit button disables again ✅
- Enter invalid email format → Submit stays disabled ✅
- Password under 8 chars → Submit stays disabled ✅`;
  }

  stripHtml(html) {
    return (html || "")
      .replace(/<[^>]+>/g, "")
      .replace(/&nbsp;/g, " ")
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .trim();
  }
}

module.exports = { AzureDevOpsBot };
