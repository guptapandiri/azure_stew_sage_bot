const {
  TeamsActivityHandler,
  MessageFactory,
  CardFactory,
  TurnContext,
} = require("botbuilder");
const axios = require("axios");
const { ORG, PROJECT, BASE_URL, HEADERS } = require("./adoConfig");
const { generateText } = require("./aiProvider");
const agenticFix = require("./agenticFix");
const pipelineCmd = require("./pipelineCommands");
const intentParser = require("./intentParser");


class AzureDevOpsBot extends TeamsActivityHandler {
  constructor() {
    super();
    this.sessions = new Map();
    this.conversationReferences = new Map();

    // Store conversation reference on every incoming activity
    this.onConversationUpdate(async (context, next) => {
      this._addConversationReference(context.activity);
      await next();
    });

    this.onMessage(async (context, next) => {
      this._addConversationReference(context.activity);
      const convId = context.activity.conversation.id;

      if (!this.sessions.has(convId)) {
        this.sessions.set(convId, {
          workItems: [],
          scheduledBug: null,
          aiFix: null,
          fixedFiles: null,
          fixRepo: null,
          pendingFixBug: null,
          repos: [],
          prFlow: {},
          pipelineFlow: {},
        });
      }

      // Handle Adaptive Card submit actions (e.g. repo selector for PR list)
      const cardValue = context.activity.value;
      if (cardValue && cardValue.action === "select_repo") {
        await this.showPRsForRepo(context, convId, cardValue.repoId);
        await next();
        return;
      }
      if (cardValue && cardValue.action === "select_repo_for_pr") {
        await this.showSourceBranchSelector(context, convId, cardValue.repoId);
        await next();
        return;
      }
      if (cardValue && cardValue.action === "select_source_branch") {
        await this.showTargetBranchSelector(context, convId, cardValue.sourceBranch);
        await next();
        return;
      }
      if (cardValue && cardValue.action === "select_target_branch") {
        await this.createPRFromBranches(context, convId, cardValue.targetBranch);
        await next();
        return;
      }
      if (cardValue && cardValue.action === "select_repo_for_fix") {
        await agenticFix.continueFixBugFlow(context, this.sessions.get(convId), cardValue.repoId);
        await next();
        return;
      }
      if (cardValue && cardValue.action === "pipeline_select_repo") {
        await pipelineCmd.showPipelinesForRepo(context, this.sessions.get(convId), cardValue.repoId, cardValue.intent);
        await next();
        return;
      }
      if (cardValue && cardValue.action === "pipeline_select_pipeline") {
        await pipelineCmd.handlePipelineSelected(context, this.sessions.get(convId), cardValue.pipelineId);
        await next();
        return;
      }
      if (cardValue && cardValue.action === "pipeline_run") {
        await pipelineCmd.triggerPipelineRun(context, this.sessions.get(convId), cardValue.branch);
        await next();
        return;
      }

      const text = (context.activity.text || "")
        .replace(/<[^>]+>/g, "")
        .trim();

      await context.sendActivity({ type: "typing" });

      const parsed = await intentParser.parseIntent(text);

      if (parsed) {
        await this.dispatch(context, convId, parsed, text);
      } else {
        // Fallback regex routing when AI parser is unavailable
        console.warn("[bot] AI intent parser unavailable — using regex fallback");
        const ltext = text.toLowerCase();
        const scheduleMatch = ltext.match(/(?:fix|schedule)\s+bug\s+(\d+)/);
        if (scheduleMatch) {
          await this.scheduleBug(context, convId, parseInt(scheduleMatch[1]));
        } else if (ltext.includes("show bugs") || ltext.includes("list bugs") || ltext === "bugs") {
          await this.showWorkItems(context, convId, ["Bug"]);
        } else if (ltext.includes("raise pr") || ltext.includes("create pr")) {
          await this.showRepoSelectorForPR(context, convId);
        } else if (ltext.includes("list pr") || ltext.includes("show pr") || ltext === "prs") {
          await this.showRepoSelector(context, convId);
        } else if (ltext.includes("sprint")) {
          await this.showSprint(context);
        } else if (ltext.includes("run pipeline") || ltext.includes("trigger pipeline") || ltext.includes("run build")) {
          await pipelineCmd.showRepoSelectorForPipeline(context, this.sessions.get(convId), "run");
        } else if (ltext.includes("pipeline log") || ltext.includes("build log") || ltext.includes("failed log") || ltext.includes("failure log")) {
          await pipelineCmd.showRepoSelectorForPipeline(context, this.sessions.get(convId), "logs");
        } else if (ltext.includes("pipeline") || ltext.includes("build")) {
          await pipelineCmd.showRepoSelectorForPipeline(context, this.sessions.get(convId), "status");
        } else if (ltext.includes("repo")) {
          await this.showRepos(context);
        } else {
          await this.showHelp(context);
        }
      }

      await next();
    });

    this.onMembersAdded(async (context, next) => {
      this._addConversationReference(context.activity);
      for (const member of context.activity.membersAdded) {
        if (member.id !== context.activity.recipient.id) {
          await context.sendActivity(
            "👋 Hi! I'm **StewSage** — Stewart's AI-powered dev assistant!\n\nType `show bugs` to see open bugs and user stories, or type `help` to see all commands.",
          );
        }
      }
      await next();
    });
  }

  _addConversationReference(activity) {
    const ref = TurnContext.getConversationReference(activity);
    this.conversationReferences.set(ref.conversation.id, ref);
    console.log(`[bot] Stored conversation ref: ${ref.conversation.id} (total: ${this.conversationReferences.size})`);
  }

  getConversationReferences() {
    return this.conversationReferences;
  }

  async dispatch(context, convId, { intent, params = {} }, originalText = "") {
    const session = this.sessions.get(convId);
    switch (intent) {
      case "show_work_items": {
        const types = Array.isArray(params.types) && params.types.length
          ? params.types
          : ["Bug", "User Story"];
        await this.showWorkItems(context, convId, types);
        break;
      }
      case "fix_bug":
        await this.scheduleBug(context, convId, params.bugNumber || 1);
        break;
      case "raise_pr":
        await this.showRepoSelectorForPR(context, convId);
        break;
      case "list_prs":
        await this.showRepoSelector(context, convId);
        break;
      case "sprint_status":
        await this.showSprint(context);
        break;
      case "run_pipeline":
        await pipelineCmd.showRepoSelectorForPipeline(context, session, "run");
        break;
      case "pipeline_logs":
        await pipelineCmd.showRepoSelectorForPipeline(context, session, "logs");
        break;
      case "pipeline_status":
        await pipelineCmd.showRepoSelectorForPipeline(context, session, "status");
        break;
      case "list_repos":
        await this.showRepos(context);
        break;
      case "chat":
        await this.handleChat(context, originalText);
        break;
      default:
        await this.showHelp(context);
    }
  }

  async handleChat(context, text) {
    try {
      await context.sendActivity({ type: "typing" });
      const prompt = `You are StewSage, an AI developer assistant embedded in Microsoft Teams for the Stewart engineering team. You help developers with code questions, debugging, architecture decisions, code reviews, and general software engineering topics.

Be concise and practical. Use markdown formatting — code blocks for code, bullet points for lists. If the question is vague, give a focused answer and offer to go deeper.

Developer's message:
${text}`;
      const response = await generateText(prompt);
      await context.sendActivity(response);
    } catch (err) {
      console.error("handleChat error:", err.message);
      await context.sendActivity(`❌ AI unavailable: ${err.message}\n\nMake sure \`AI_PROVIDER\` is set correctly in your environment.`);
    }
  }

  async showWorkItems(context, convId, types) {
    try {
      await context.sendActivity({ type: "typing" });

      const typeList = types.map((t) => `'${t}'`).join(",");
      const wiql = {
        query: `SELECT [System.Id],[System.Title],[System.State],[System.AssignedTo],[System.WorkItemType] FROM WorkItems WHERE [System.TeamProject] = '${PROJECT}' AND [System.WorkItemType] IN (${typeList}) AND [System.State] <> 'Done' AND [System.State] <> 'Closed' AND [System.State] <> 'Resolved' ORDER BY [System.WorkItemType] ASC,[System.CreatedDate] DESC`,
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
        await context.sendActivity(`No open ${types.join(" / ")} found in the project!`);
        return;
      }

      const itemsRes = await axios.get(
        `${BASE_URL}/wit/workitems?ids=${ids.join(",")}&fields=System.Id,System.Title,System.State,System.AssignedTo,System.WorkItemType&api-version=7.1`,
        { headers: HEADERS },
      );

      const workItems = itemsRes.data.value;
      this.sessions.get(convId).workItems = workItems;

      const TYPE_LABEL = {
        "Bug": "🐛 Bug",
        "User Story": "📖 User Story",
        "Task": "✅ Task",
        "Feature": "⭐ Feature",
        "Epic": "🚀 Epic",
      };
      const cardTitle = types.length >= 5
        ? `🗂️ All Work Items — ${PROJECT}`
        : `${types.map((t) => TYPE_LABEL[t] || t).join(" & ")} — ${PROJECT}`;

      const card = {
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.5",
        body: [
          {
            type: "TextBlock",
            text: cardTitle,
            size: "Large",
            weight: "Bolder",
          },
          {
            type: "TextBlock",
            text: "Type: fix bug [number]  —  AI will write a code fix for that item",
            isSubtle: true,
            size: "Small",
            spacing: "None",
          },
          ...workItems.map((bug, i) => {
            const type = bug.fields["System.WorkItemType"];
            const typeLabel = TYPE_LABEL[type] || type;
            return {
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
                      text: `${typeLabel} · ${bug.fields["System.AssignedTo"]?.displayName || "Unassigned"}`,
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
            };
          }),
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

      if (!session.workItems.length) {
        await context.sendActivity("Please show work items first (e.g. `show bugs`) to load the list.");
        return;
      }

      const bug = session.workItems[listNumber - 1];
      if (!bug) {
        await context.sendActivity(
          `Item #${listNumber} not found. There are ${session.workItems.length} items loaded. Try again.`,
        );
        return;
      }

      await context.sendActivity({ type: "typing" });
      await context.sendActivity(
        `🔍 Fetching full details for **#${bug.id}: ${bug.fields["System.Title"]}**...`,
      );

      const detailRes = await axios.get(
        `${BASE_URL}/wit/workitems/${bug.id}?api-version=7.1`,
        { headers: HEADERS },
      );
      const f = detailRes.data.fields;
      const bugDetails = {
        id: bug.id,
        title: f["System.Title"] || "",
        type: f["System.WorkItemType"] || "Bug",
        description: this.stripHtml(f["System.Description"] || "No description provided."),
        reproSteps: this.stripHtml(f["Microsoft.VSTS.TCM.ReproSteps"] || ""),
        acceptanceCriteria: this.stripHtml(f["Microsoft.VSTS.Common.AcceptanceCriteria"] || ""),
      };

      session.scheduledBug = bug;
      session.pendingFixBug = bugDetails;
      // Clear any previous fix state
      session.fixedFiles = null;
      session.fixRepo = null;

      const reposRes = await axios.get(`${BASE_URL}/git/repositories?api-version=7.1`, { headers: HEADERS });
      const repos = reposRes.data.value || [];
      session.repos = repos;

      if (!repos.length) {
        await context.sendActivity("No repositories found in this project.");
        return;
      }

      if (repos.length === 1) {
        await agenticFix.runAgenticFixFlow(context, session, repos[0], bugDetails);
      } else {
        await agenticFix.showRepoSelectorForFix(context, bugDetails, repos);
      }
    } catch (err) {
      console.error("scheduleBug error:", err.response?.data || err.message);
      await context.sendActivity("Could not analyze bug. Please try again.");
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

  async showRepoSelector(context, convId) {
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

      // Cache repos in session for later lookup
      this.sessions.get(convId).repos = repos;

      const card = {
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.5",
        body: [
          {
            type: "TextBlock",
            text: "🔀 List Pull Requests",
            size: "Large",
            weight: "Bolder",
          },
          {
            type: "TextBlock",
            text: "Select a repository to view its pull requests:",
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
            title: "View PRs",
            data: { action: "select_repo" },
          },
        ],
      };

      await context.sendActivity(
        MessageFactory.attachment(CardFactory.adaptiveCard(card)),
      );
    } catch (err) {
      console.error(
        "showRepoSelector error:",
        err.response?.data || err.message,
      );
      await context.sendActivity("Could not fetch repositories.");
    }
  }

  async showPRsForRepo(context, convId, repoId) {
    try {
      await context.sendActivity({ type: "typing" });

      const session = this.sessions.get(convId);
      const repo = (session.repos || []).find((r) => r.id === repoId);
      const repoName = repo?.name || repoId;

      const res = await axios.get(
        `${BASE_URL}/git/repositories/${repoId}/pullrequests?searchCriteria.status=all&$top=20&api-version=7.1`,
        { headers: HEADERS },
      );
      const prs = res.data.value;

      if (!prs?.length) {
        await context.sendActivity(
          `No pull requests found for **${repoName}**.`,
        );
        return;
      }

      const statusColor = (s) =>
        ({ active: "Accent", completed: "Good", abandoned: "Attention" })[s] ||
        "Default";

      const statusEmoji = (s) =>
        ({ active: "🟡", completed: "✅", abandoned: "❌" })[s] || "⚪";

      const card = {
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.5",
        body: [
          {
            type: "TextBlock",
            text: `🔀 Pull Requests — ${repoName}`,
            size: "Large",
            weight: "Bolder",
          },
          {
            type: "TextBlock",
            text: `Showing ${prs.length} most recent PR(s)`,
            isSubtle: true,
            size: "Small",
            spacing: "None",
          },
          ...prs.map((pr) => {
            const src = pr.sourceRefName.replace("refs/heads/", "");
            const tgt = pr.targetRefName.replace("refs/heads/", "");
            const date = new Date(pr.creationDate).toLocaleDateString();
            const prUrl = `https://dev.azure.com/${ORG}/${PROJECT}/_git/${repoName}/pullrequest/${pr.pullRequestId}`;

            return {
              type: "ColumnSet",
              separator: true,
              selectAction: { type: "Action.OpenUrl", url: prUrl },
              columns: [
                {
                  type: "Column",
                  width: "stretch",
                  items: [
                    {
                      type: "TextBlock",
                      text: `${statusEmoji(pr.status)} **#${pr.pullRequestId}** · ${pr.title}`,
                      wrap: true,
                      size: "Small",
                      weight: "Bolder",
                    },
                    {
                      type: "TextBlock",
                      text: `\`${src}\` → \`${tgt}\``,
                      isSubtle: true,
                      size: "Small",
                      spacing: "None",
                      wrap: true,
                    },
                    {
                      type: "TextBlock",
                      text: `👤 ${pr.createdBy?.displayName || "Unknown"} · 📅 ${date}`,
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
                      text:
                        pr.status.charAt(0).toUpperCase() + pr.status.slice(1),
                      size: "Small",
                      color: statusColor(pr.status),
                    },
                  ],
                },
              ],
            };
          }),
        ],
      };

      await context.sendActivity(
        MessageFactory.attachment(CardFactory.adaptiveCard(card)),
      );
    } catch (err) {
      console.error(
        "showPRsForRepo error:",
        err.response?.data || err.message,
      );
      await context.sendActivity(
        "Could not fetch pull requests. Please try again.",
      );
    }
  }

  async showRepoSelectorForPR(context, convId) {
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

      this.sessions.get(convId).repos = repos;

      const card = {
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.5",
        body: [
          {
            type: "TextBlock",
            text: "🔀 Raise a Pull Request",
            size: "Large",
            weight: "Bolder",
          },
          {
            type: "TextBlock",
            text: "Step 1: Select the repository",
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
            data: { action: "select_repo_for_pr" },
          },
        ],
      };

      await context.sendActivity(
        MessageFactory.attachment(CardFactory.adaptiveCard(card)),
      );
    } catch (err) {
      console.error(
        "showRepoSelectorForPR error:",
        err.response?.data || err.message,
      );
      await context.sendActivity("Could not fetch repositories.");
    }
  }

  async showSourceBranchSelector(context, convId, repoId) {
    try {
      await context.sendActivity({ type: "typing" });

      const session = this.sessions.get(convId);
      const repo = (session.repos || []).find((r) => r.id === repoId);
      session.prFlow = { repoId, repoName: repo?.name || repoId };

      const res = await axios.get(
        `${BASE_URL}/git/repositories/${repoId}/refs?filter=heads&api-version=7.1`,
        { headers: HEADERS },
      );
      const branches = res.data.value || [];

      if (!branches.length) {
        await context.sendActivity("No branches found in this repository.");
        return;
      }

      session.prFlow.branches = branches;

      const card = {
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.5",
        body: [
          {
            type: "TextBlock",
            text: `🔀 Raise PR — ${session.prFlow.repoName}`,
            size: "Large",
            weight: "Bolder",
          },
          {
            type: "TextBlock",
            text: "Step 2: Select the **source** branch (branch with your changes)",
            wrap: true,
            spacing: "Small",
          },
          {
            type: "Input.ChoiceSet",
            id: "sourceBranch",
            style: "compact",
            placeholder: "Choose source branch...",
            choices: branches.map((b) => {
              const name = b.name.replace("refs/heads/", "");
              return { title: name, value: name };
            }),
          },
        ],
        actions: [
          {
            type: "Action.Submit",
            title: "Next →",
            data: { action: "select_source_branch" },
          },
        ],
      };

      await context.sendActivity(
        MessageFactory.attachment(CardFactory.adaptiveCard(card)),
      );
    } catch (err) {
      console.error(
        "showSourceBranchSelector error:",
        err.response?.data || err.message,
      );
      await context.sendActivity("Could not fetch branches.");
    }
  }

  async showTargetBranchSelector(context, convId, sourceBranch) {
    try {
      const session = this.sessions.get(convId);
      session.prFlow.sourceBranch = sourceBranch;

      const branches = session.prFlow.branches || [];
      const targetBranches = branches.filter(
        (b) => b.name.replace("refs/heads/", "") !== sourceBranch,
      );

      if (!targetBranches.length) {
        await context.sendActivity(
          "No other branches available as target.",
        );
        return;
      }

      const card = {
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.5",
        body: [
          {
            type: "TextBlock",
            text: `🔀 Raise PR — ${session.prFlow.repoName}`,
            size: "Large",
            weight: "Bolder",
          },
          {
            type: "TextBlock",
            text: `Source: \`${sourceBranch}\``,
            wrap: true,
            spacing: "Small",
          },
          {
            type: "TextBlock",
            text: "Step 3: Select the **target** branch (branch to merge into)",
            wrap: true,
            spacing: "Small",
          },
          {
            type: "Input.ChoiceSet",
            id: "targetBranch",
            style: "compact",
            placeholder: "Choose target branch...",
            choices: targetBranches.map((b) => {
              const name = b.name.replace("refs/heads/", "");
              return { title: name, value: name };
            }),
          },
        ],
        actions: [
          {
            type: "Action.Submit",
            title: "Create PR 🚀",
            data: { action: "select_target_branch" },
          },
        ],
      };

      await context.sendActivity(
        MessageFactory.attachment(CardFactory.adaptiveCard(card)),
      );
    } catch (err) {
      console.error(
        "showTargetBranchSelector error:",
        err.response?.data || err.message,
      );
      await context.sendActivity("Could not prepare branch selection.");
    }
  }

  async createPRFromBranches(context, convId, targetBranch) {
    try {
      const session = this.sessions.get(convId);
      const { repoId, repoName, sourceBranch } = session.prFlow;
      session.prFlow.targetBranch = targetBranch;

      await context.sendActivity({ type: "typing" });
      await context.sendActivity(
        `🔍 Comparing \`${sourceBranch}\` → \`${targetBranch}\` in **${repoName}**...`,
      );

      // Fetch diff between branches
      const diffRes = await axios.get(
        `${BASE_URL}/git/repositories/${repoId}/diffs/commits?baseVersion=${encodeURIComponent(targetBranch)}&targetVersion=${encodeURIComponent(sourceBranch)}&baseVersionType=branch&targetVersionType=branch&api-version=7.1`,
        { headers: HEADERS },
      );
      const changes = (diffRes.data.changes || []).slice(0, 50);
      const changeSummary = changes
        .map((c) => `${c.changeType}: ${c.item?.path || "unknown"}`)
        .join("\n");

      // Fetch commits between branches
      const commitsRes = await axios.get(
        `${BASE_URL}/git/repositories/${repoId}/commits?searchCriteria.compareVersion.version=${encodeURIComponent(targetBranch)}&searchCriteria.compareVersion.versionType=branch&searchCriteria.itemVersion.version=${encodeURIComponent(sourceBranch)}&searchCriteria.itemVersion.versionType=branch&$top=30&api-version=7.1`,
        { headers: HEADERS },
      );
      const commits = commitsRes.data.value || [];
      const commitSummary = commits.map((c) => `- ${c.comment}`).join("\n");

      if (!changes.length && !commits.length) {
        await context.sendActivity(
          `No differences found between \`${sourceBranch}\` and \`${targetBranch}\`. Nothing to merge.`,
        );
        session.prFlow = {};
        return;
      }

      await context.sendActivity(
        "🤖 Generating PR description from branch differences...",
      );
      await context.sendActivity({ type: "typing" });

      const prompt = `You are a senior software engineer. Generate a clear, professional pull request title and description based on the following branch comparison.

Source branch: ${sourceBranch}
Target branch: ${targetBranch}
Repository: ${repoName}

Commits (${commits.length}):
${commitSummary || "No commit messages available."}

File changes (${changes.length}):
${changeSummary || "No file changes detected."}

Respond in EXACTLY this format:
PR_TITLE: <a concise PR title>

PR_DESCRIPTION:
<a well-structured PR description in markdown, including:
- Summary of changes
- List of key modifications
- Any notable files changed>`;

      let prTitle, prDescription;
      try {
        const aiResponse = await generateText(prompt);

        const titleMatch = aiResponse.match(/PR_TITLE:\s*(.+)/);
        prTitle = titleMatch
          ? titleMatch[1].trim()
          : `Merge ${sourceBranch} into ${targetBranch}`;

        const descMatch = aiResponse.match(/PR_DESCRIPTION:\s*([\s\S]+)/);
        prDescription = descMatch ? descMatch[1].trim() : aiResponse;
      } catch (aiErr) {
        console.warn(
          "AI unavailable, using auto-generated description:",
          aiErr.message,
        );
        prTitle = `Merge ${sourceBranch} into ${targetBranch}`;
        prDescription = `## Summary\n\nMerge \`${sourceBranch}\` → \`${targetBranch}\`\n\n### Commits (${commits.length})\n${commitSummary || "N/A"}\n\n### Files Changed (${changes.length})\n${changeSummary || "N/A"}`;
      }

      await context.sendActivity("🔀 Creating Pull Request...");

      const prRes = await axios.post(
        `${BASE_URL}/git/repositories/${repoId}/pullrequests?api-version=7.1`,
        {
          title: prTitle,
          description: prDescription,
          sourceRefName: `refs/heads/${sourceBranch}`,
          targetRefName: `refs/heads/${targetBranch}`,
        },
        { headers: HEADERS },
      );

      const prUrl = `https://dev.azure.com/${ORG}/${PROJECT}/_git/${repoName}/pullrequest/${prRes.data.pullRequestId}`;
      session.prFlow = {};

      await context.sendActivity(
        `✅ **PR Created Successfully!**\n\n` +
          `🔗 [View PR #${prRes.data.pullRequestId}](${prUrl})\n` +
          `📌 \`${sourceBranch}\` → \`${targetBranch}\`\n` +
          `📁 Repo: **${repoName}**\n\n` +
          `**Title:** ${prTitle}`,
      );
    } catch (err) {
      console.error(
        "createPRFromBranches error:",
        JSON.stringify(err.response?.data || err.message, null, 2),
      );
      await context.sendActivity(
        "Could not create PR. Check that the PAT has **Code (Read & Write)** permission.",
      );
    }
  }

  async showHelp(context) {
    await context.sendActivity(
      "🤖 **StewSage — Your AI Dev Workspace:**\n\n" +
        "🐛 `show bugs` — List open bugs and user stories from the current sprint\n" +
        "🔧 `fix bug 2` — AI scans the repo, generates a code fix, and raises a PR automatically\n" +
        "🔀 `raise PR` — Pick repo & branches, AI writes the PR title and description\n" +
        "📋 `list PRs` — Browse and review pull requests across your repositories\n" +
        "🏃 `sprint status` — View current sprint name, dates, and progress\n" +
        "🚀 `run pipeline` — Select repo → pipeline → branch, then trigger a build\n" +
        "📊 `pipeline status` — View recent runs for a pipeline with results and duration\n" +
        "📋 `pipeline logs` — Fetch failed run logs with AI-powered root cause analysis\n" +
        "📁 `list repos` — Browse all repositories in the Azure DevOps project",
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
