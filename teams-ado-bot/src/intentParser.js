"use strict";

const { generateText } = require("./aiProvider");

const SYSTEM_PROMPT = `You are a command parser for a Microsoft Teams bot that integrates with Azure DevOps.
Parse the user message and return a JSON object with "intent" and "params".

Available intents:
- "show_work_items": View work items. params: { types: string[] } — subset of ["Bug","User Story","Task","Feature","Epic"]. Default to ["Bug","User Story"] when user says "show bugs" / "show stories" / doesn't specify. Use all five only when user explicitly asks for everything or all work items.
- "fix_bug": Fix a specific item by its list number. params: { bugNumber: number }
- "raise_pr": Create a pull request. params: {}
- "list_prs": View pull requests. params: {}
- "sprint_status": View sprint info. params: {}
- "run_pipeline": Trigger or run a pipeline build. params: {}
- "pipeline_logs": View failed pipeline logs. params: {}
- "pipeline_status": View pipeline or build status. params: {}
- "list_repos": View repositories. params: {}
- "help": Unknown request or explicit help. params: {}

Examples:
"show bugs" → {"intent":"show_work_items","params":{"types":["Bug"]}}
"show user stories" → {"intent":"show_work_items","params":{"types":["User Story"]}}
"show bugs and user stories" → {"intent":"show_work_items","params":{"types":["Bug","User Story"]}}
"I want to view bugs and user stories" → {"intent":"show_work_items","params":{"types":["Bug","User Story"]}}
"show tasks" → {"intent":"show_work_items","params":{"types":["Task"]}}
"show bugs and tasks" → {"intent":"show_work_items","params":{"types":["Bug","Task"]}}
"show everything" → {"intent":"show_work_items","params":{"types":["Bug","User Story","Task","Feature","Epic"]}}
"show all work items" → {"intent":"show_work_items","params":{"types":["Bug","User Story","Task","Feature","Epic"]}}
"list bugs" → {"intent":"show_work_items","params":{"types":["Bug"]}}
"bugs" → {"intent":"show_work_items","params":{"types":["Bug","User Story"]}}
"fix bug 3" → {"intent":"fix_bug","params":{"bugNumber":3}}
"schedule bug 2" → {"intent":"fix_bug","params":{"bugNumber":2}}
"fix the second item" → {"intent":"fix_bug","params":{"bugNumber":2}}
"raise a PR" → {"intent":"raise_pr","params":{}}
"create pull request" → {"intent":"raise_pr","params":{}}
"show PRs" → {"intent":"list_prs","params":{}}
"list pull requests" → {"intent":"list_prs","params":{}}
"sprint status" → {"intent":"sprint_status","params":{}}
"what sprint are we in?" → {"intent":"sprint_status","params":{}}
"run the build" → {"intent":"run_pipeline","params":{}}
"trigger pipeline" → {"intent":"run_pipeline","params":{}}
"build logs" → {"intent":"pipeline_logs","params":{}}
"show failed build logs" → {"intent":"pipeline_logs","params":{}}
"pipeline status" → {"intent":"pipeline_status","params":{}}
"build status" → {"intent":"pipeline_status","params":{}}
"list repos" → {"intent":"list_repos","params":{}}
"help" → {"intent":"help","params":{}}
"banana" → {"intent":"help","params":{}}

Respond with ONLY a JSON object, no markdown, no explanation.

<user_message>
{{MESSAGE}}
</user_message>`;

async function parseIntent(text) {
  try {
    const prompt = SYSTEM_PROMPT.replace("{{MESSAGE}}", text);
    const response = await generateText(prompt);
    const match = response.match(/\{[\s\S]*\}/);
    if (!match) return null;
    const parsed = JSON.parse(match[0]);
    if (!parsed.intent) return null;
    return parsed;
  } catch (err) {
    console.warn("[intentParser] AI parse failed:", err.message);
    return null;
  }
}

module.exports = { parseIntent };
