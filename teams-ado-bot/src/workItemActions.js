"use strict";

const axios = require("axios");
const { ORG, PROJECT, BASE_URL, HEADERS } = require("./adoConfig");

// Valid states per work item type (ADO Agile process)
const VALID_STATES = {
  Bug: ["New", "Active", "Resolved", "Closed"],
  "User Story": ["New", "Active", "Resolved", "Closed"],
  Task: ["To Do", "In Progress", "Done"],
  Feature: ["New", "Active", "Resolved", "Closed"],
  Epic: ["New", "Active", "Resolved", "Closed"],
};

// Resolve a display name fragment to an ADO unique name (email)
async function resolveADOIdentity(nameFragment) {
  try {
    const res = await axios.get(
      `https://vssps.dev.azure.com/${ORG}/_apis/identities?searchFilter=General&filterValue=${encodeURIComponent(nameFragment)}&api-version=6.0`,
      { headers: HEADERS },
    );
    const identities = res.data.value || [];
    if (!identities.length) return null;
    // Prefer exact display name match, fall back to first result
    const exact = identities.find(
      (i) => i.providerDisplayName?.toLowerCase() === nameFragment.toLowerCase(),
    );
    const identity = exact || identities[0];
    return identity.properties?.Account?.["$value"] || identity.subjectDescriptor || null;
  } catch (err) {
    console.warn("[workItemActions] Identity lookup failed:", err.response?.data || err.message);
    return null;
  }
}

function findWorkItem(session, itemNumber) {
  return session.workItems[itemNumber - 1] ||
    session.workItems.find((w) => w.id === itemNumber) ||
    null;
}

async function assignWorkItem(context, session, itemNumber, assignTo) {
  const workItem = findWorkItem(session, itemNumber);
  if (!workItem) {
    await context.sendActivity(`Item #${itemNumber} not found. Run \`show bugs\` to reload the list.`);
    return;
  }

  await context.sendActivity({ type: "typing" });

  let assignee;
  if (assignTo === "me") {
    assignee = context.activity.from.name;
  } else {
    assignee = assignTo;
  }

  // Try to resolve to ADO identity (email) for reliable assignment
  const uniqueName = await resolveADOIdentity(assignee);
  const assigneeValue = uniqueName || assignee;

  try {
    await axios.patch(
      `${BASE_URL}/wit/workitems/${workItem.id}?api-version=7.1`,
      [{ op: "add", path: "/fields/System.AssignedTo", value: assigneeValue }],
      { headers: { ...HEADERS, "Content-Type": "application/json-patch+json" } },
    );

    const displayAssignee = assignTo === "me" ? `you (${assignee})` : assignee;
    await context.sendActivity(
      `✅ **#${workItem.id}** — ${workItem.fields["System.Title"]}\nassigned to **${displayAssignee}**`,
    );
  } catch (err) {
    console.error("[workItemActions] assignWorkItem error:", err.response?.data || err.message);
    const detail = err.response?.data?.message || err.message;
    if (detail.toLowerCase().includes("unknown identity")) {
      await context.sendActivity(
        `❌ Could not find **"${assignTo}"** in Azure DevOps.\n\nTry using their full display name, e.g. \`assign #${workItem.id} to Gupta Pandiri\``
      );
    } else {
      await context.sendActivity(`❌ Could not assign work item: ${detail}`);
    }
  }
}

async function updateWorkItemStatus(context, session, itemNumber, newStatus) {
  const workItem = findWorkItem(session, itemNumber);
  if (!workItem) {
    await context.sendActivity(`Item #${itemNumber} not found. Run \`show bugs\` to reload the list.`);
    return;
  }

  const itemType = workItem.fields["System.WorkItemType"] || "Bug";
  const validStates = VALID_STATES[itemType] || VALID_STATES["Bug"];

  // Case-insensitive match against valid states
  const matched = validStates.find((s) => s.toLowerCase() === newStatus.toLowerCase());
  if (!matched) {
    await context.sendActivity(
      `❌ **"${newStatus}"** is not a valid state for **${itemType}**.\n\nValid states: ${validStates.join(", ")}`,
    );
    return;
  }

  await context.sendActivity({ type: "typing" });

  try {
    await axios.patch(
      `${BASE_URL}/wit/workitems/${workItem.id}?api-version=7.1`,
      [{ op: "add", path: "/fields/System.State", value: matched }],
      { headers: { ...HEADERS, "Content-Type": "application/json-patch+json" } },
    );

    await context.sendActivity(
      `✅ **#${workItem.id}** — ${workItem.fields["System.Title"]}\nStatus updated to **${matched}**`,
    );
  } catch (err) {
    console.error("[workItemActions] updateWorkItemStatus error:", err.response?.data || err.message);
    const detail = err.response?.data?.message || err.message;
    await context.sendActivity(`❌ Could not update status: ${detail}`);
  }
}

async function addWorkItemComment(context, session, itemNumber, comment) {
  const workItem = findWorkItem(session, itemNumber);
  if (!workItem) {
    await context.sendActivity(`Item #${itemNumber} not found. Run \`show bugs\` to reload the list.`);
    return;
  }

  await context.sendActivity({ type: "typing" });

  try {
    await axios.post(
      `${BASE_URL}/wit/workitems/${workItem.id}/comments?api-version=7.1-preview.3`,
      { text: comment },
      { headers: HEADERS },
    );

    await context.sendActivity(
      `💬 Comment added to **#${workItem.id}** — ${workItem.fields["System.Title"]}`,
    );
  } catch (err) {
    console.error("[workItemActions] addWorkItemComment error:", err.response?.data || err.message);
    const detail = err.response?.data?.message || err.message;
    await context.sendActivity(`❌ Could not add comment: ${detail}`);
  }
}

const VALID_TYPES = ["Bug", "User Story", "Task", "Feature", "Epic"];

async function createWorkItem(context, type, title, description, assignTo) {
  if (!title?.trim()) {
    await context.sendActivity("❌ Please include a title, e.g. `create a bug: Login button is broken`");
    return;
  }

  const matchedType = VALID_TYPES.find((t) => t.toLowerCase() === (type || "").toLowerCase()) || "Task";

  await context.sendActivity({ type: "typing" });

  const ops = [
    { op: "add", path: "/fields/System.Title", value: title.trim() },
  ];

  if (description?.trim()) {
    ops.push({ op: "add", path: "/fields/System.Description", value: description.trim() });
  }

  if (assignTo) {
    const assignee = assignTo === "me" ? context.activity.from.name : assignTo;
    const uniqueName = await resolveADOIdentity(assignee);
    ops.push({ op: "add", path: "/fields/System.AssignedTo", value: uniqueName || assignee });
  }

  try {
    const res = await axios.patch(
      `${BASE_URL}/wit/workitems/${encodeURIComponent("$" + matchedType)}?api-version=7.1`,
      ops,
      { headers: { ...HEADERS, "Content-Type": "application/json-patch+json" } },
    );

    const item = res.data;
    const itemUrl = `https://dev.azure.com/${ORG}/${PROJECT}/_workitems/edit/${item.id}`;
    const assignedTo = item.fields["System.AssignedTo"]?.displayName || "Unassigned";

    await context.sendActivity(
      `✅ **${matchedType} Created!**\n\n` +
      `**#${item.id}** · ${item.fields["System.Title"]}\n` +
      `👤 Assigned to: ${assignedTo}\n` +
      `🔗 [View in ADO](${itemUrl})`,
    );
  } catch (err) {
    console.error("[workItemActions] createWorkItem error:", err.response?.data || err.message);
    const detail = err.response?.data?.message || err.message;
    await context.sendActivity(`❌ Could not create work item: ${detail}`);
  }
}

module.exports = { assignWorkItem, updateWorkItemStatus, addWorkItemComment, createWorkItem };
