import { z } from "zod";
import { ToolDefinition } from "../registry.js";

const BASE = "https://graph.microsoft.com/v1.0";

function buildHeaders(): Record<string, string> {
  return {
    Authorization: "Bearer {token}",
    "Content-Type": "application/json",
  };
}

const PLANNER_READ_PERMISSIONS = {
  delegated: ["Tasks.Read"],
  application: ["Tasks.Read.All"],
};

const PLANNER_WRITE_PERMISSIONS = {
  delegated: ["Tasks.ReadWrite"],
  application: ["Tasks.ReadWrite.All"],
};

const TODO_READ_PERMISSIONS = {
  delegated: ["Tasks.Read"],
  application: ["Tasks.Read.All"],
};

const TODO_WRITE_PERMISSIONS = {
  delegated: ["Tasks.ReadWrite"],
  application: ["Tasks.ReadWrite.All"],
};

export const tasksTools: ToolDefinition[] = [
  {
    name: "graph_tasks_list_plans",
    description: "List Planner plans for a group",
    category: "tasks",
    zodShape: {
      groupId: z.string().describe("Microsoft 365 group ID that owns the plans"),
    },
    handler: (args: { groupId: string }) => {
      const endpoint = `${BASE}/groups/${args.groupId}/planner/plans`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { groupId: args.groupId },
        queryParams: {},
        body: null,
        description: `List Planner plans for group ${args.groupId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/plannergroup-list-plans",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst data = await response.json();`,
        requiredPermissions: PLANNER_READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_tasks_get_plan",
    description: "Get a Planner plan by ID",
    category: "tasks",
    zodShape: {
      planId: z.string().describe("Planner plan ID"),
    },
    handler: (args: { planId: string }) => {
      const endpoint = `${BASE}/planner/plans/${args.planId}`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { planId: args.planId },
        queryParams: {},
        body: null,
        description: `Get plan ${args.planId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/plannerplan-get",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst plan = await response.json();`,
        requiredPermissions: PLANNER_READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_tasks_create_plan",
    description: "Create a Planner plan",
    category: "tasks",
    zodShape: {
      title: z.string().describe("Title of the new plan"),
      ownerId: z.string().describe("Group ID that owns the plan"),
    },
    handler: (args: { title: string; ownerId: string }) => {
      const endpoint = `${BASE}/planner/plans`;
      const body = { title: args.title, owner: args.ownerId };
      return {
        endpoint,
        method: "POST",
        headers: buildHeaders(),
        pathParams: {},
        queryParams: {},
        body,
        description: `Create plan '${args.title}' owned by group ${args.ownerId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/planner-post-plans",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'POST',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(body)})\n});\nconst plan = await response.json();`,
        requiredPermissions: PLANNER_WRITE_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_tasks_list_tasks",
    description: "List tasks in a Planner plan",
    category: "tasks",
    zodShape: {
      planId: z.string().describe("Planner plan ID"),
    },
    handler: (args: { planId: string }) => {
      const endpoint = `${BASE}/planner/plans/${args.planId}/tasks`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { planId: args.planId },
        queryParams: {},
        body: null,
        description: `List tasks in plan ${args.planId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/plannerplan-list-tasks",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst data = await response.json();`,
        requiredPermissions: PLANNER_READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_tasks_get_task",
    description: "Get a Planner task by ID",
    category: "tasks",
    zodShape: {
      taskId: z.string().describe("Planner task ID"),
    },
    handler: (args: { taskId: string }) => {
      const endpoint = `${BASE}/planner/tasks/${args.taskId}`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { taskId: args.taskId },
        queryParams: {},
        body: null,
        description: `Get task ${args.taskId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/plannertask-get",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst task = await response.json();`,
        requiredPermissions: PLANNER_READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_tasks_create_task",
    description: "Create a Planner task",
    category: "tasks",
    zodShape: {
      planId: z.string().describe("Planner plan ID to create the task in"),
      title: z.string().describe("Title of the new task"),
      bucketId: z.string().optional().describe("Planner bucket ID to place the task in"),
      assignedToIds: z.array(z.string()).optional().describe("List of user IDs to assign the task to"),
      dueDateTime: z.string().optional().describe("Due date — ISO 8601 datetime in UTC (e.g. '2024-12-31T23:59:59Z')"),
    },
    handler: (args: {
      planId: string;
      title: string;
      bucketId?: string;
      assignedToIds?: string[];
      dueDateTime?: string;
    }) => {
      const endpoint = `${BASE}/planner/tasks`;
      const body: Record<string, unknown> = { planId: args.planId, title: args.title };
      if (args.bucketId) body.bucketId = args.bucketId;
      if (args.dueDateTime) body.dueDateTime = args.dueDateTime;
      if (args.assignedToIds?.length) {
        body.assignments = Object.fromEntries(
          args.assignedToIds.map((id) => [id, { "@odata.type": "#microsoft.graph.plannerAssignment", orderHint: " !" }])
        );
      }
      return {
        endpoint,
        method: "POST",
        headers: buildHeaders(),
        pathParams: {},
        queryParams: {},
        body,
        description: `Create task '${args.title}' in plan ${args.planId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/planner-post-tasks",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'POST',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(body, null, 2)})\n});\nconst task = await response.json();`,
        requiredPermissions: PLANNER_WRITE_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_tasks_update_task",
    description: "Update a Planner task",
    category: "tasks",
    zodShape: {
      taskId: z.string().describe("Planner task ID to update"),
      eTag: z.string().describe("ETag from the task (required for updates) — retrieve the task first to get the current value"),
      title: z.string().optional().describe("New task title"),
      percentComplete: z.number().optional().describe("Completion percentage — 0 (not started), 50 (in progress), or 100 (complete)"),
      dueDateTime: z.string().optional().describe("New due date — ISO 8601 datetime in UTC (e.g. '2024-12-31T23:59:59Z')"),
      bucketId: z.string().optional().describe("Planner bucket ID to move the task to"),
    },
    handler: (args: {
      taskId: string;
      eTag: string;
      title?: string;
      percentComplete?: number;
      dueDateTime?: string;
      bucketId?: string;
    }) => {
      const endpoint = `${BASE}/planner/tasks/${args.taskId}`;
      const body: Record<string, unknown> = {};
      if (args.title) body.title = args.title;
      if (args.percentComplete !== undefined) body.percentComplete = args.percentComplete;
      if (args.dueDateTime) body.dueDateTime = args.dueDateTime;
      if (args.bucketId) body.bucketId = args.bucketId;
      return {
        endpoint,
        method: "PATCH",
        headers: {
          ...buildHeaders(),
          "If-Match": args.eTag,
        },
        pathParams: { taskId: args.taskId },
        queryParams: {},
        body,
        description: `Update task ${args.taskId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/plannertask-update",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'PATCH',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json', 'If-Match': '${args.eTag}' },\n  body: JSON.stringify(${JSON.stringify(body, null, 2)})\n});`,
        requiredPermissions: PLANNER_WRITE_PERMISSIONS,
        notes: "The If-Match eTag header is required. Retrieve the task first to get the current eTag value. Submitting a stale eTag returns 412 Precondition Failed.",
      };
    },
  },
  {
    name: "graph_tasks_delete_task",
    description: "Delete a Planner task",
    category: "tasks",
    zodShape: {
      taskId: z.string().describe("Planner task ID to delete"),
      eTag: z.string().describe("ETag from the task — retrieve the task first to get the current value"),
    },
    handler: (args: { taskId: string; eTag: string }) => {
      const endpoint = `${BASE}/planner/tasks/${args.taskId}`;
      return {
        endpoint,
        method: "DELETE",
        headers: {
          ...buildHeaders(),
          "If-Match": args.eTag,
        },
        pathParams: { taskId: args.taskId },
        queryParams: {},
        body: null,
        description: `Delete task ${args.taskId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/plannertask-delete",
        codeExample: `await fetch('${endpoint}', {\n  method: 'DELETE',\n  headers: { 'Authorization': 'Bearer {token}', 'If-Match': '${args.eTag}' }\n});`,
        requiredPermissions: PLANNER_WRITE_PERMISSIONS,
        notes: "The If-Match eTag header is required. Retrieve the task first to get the current eTag value.",
      };
    },
  },
  {
    name: "graph_todo_list_lists",
    description: "List Microsoft To Do task lists",
    category: "tasks",
    zodShape: {
      userId: z.string().describe("User ID or UPN (e.g. user@contoso.com)"),
    },
    handler: (args: { userId: string }) => {
      const endpoint = `${BASE}/users/${args.userId}/todo/lists`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { userId: args.userId },
        queryParams: {},
        body: null,
        description: `List To Do task lists for user ${args.userId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/todo-list-lists",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst data = await response.json();`,
        requiredPermissions: TODO_READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_todo_create_task",
    description: "Create a Microsoft To Do task",
    category: "tasks",
    zodShape: {
      userId: z.string().describe("User ID or UPN (e.g. user@contoso.com)"),
      listId: z.string().describe("Microsoft To Do list ID to create the task in"),
      title: z.string().describe("Title of the new task"),
      dueDateTime: z.string().optional().describe("Due date — ISO 8601 datetime in UTC (e.g. '2024-12-31T23:59:59Z')"),
      importance: z.enum(["low", "normal", "high"]).optional().describe("Task importance level — 'low', 'normal', or 'high'"),
    },
    handler: (args: {
      userId: string;
      listId: string;
      title: string;
      dueDateTime?: string;
      importance?: "low" | "normal" | "high";
    }) => {
      const endpoint = `${BASE}/users/${args.userId}/todo/lists/${args.listId}/tasks`;
      const body: Record<string, unknown> = { title: args.title };
      if (args.dueDateTime) body.dueDateTime = { dateTime: args.dueDateTime, timeZone: "UTC" };
      if (args.importance) body.importance = args.importance;
      return {
        endpoint,
        method: "POST",
        headers: buildHeaders(),
        pathParams: { userId: args.userId, listId: args.listId },
        queryParams: {},
        body,
        description: `Create To Do task '${args.title}' in list ${args.listId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/todotasklist-post-tasks",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'POST',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(body, null, 2)})\n});\nconst task = await response.json();`,
        requiredPermissions: TODO_WRITE_PERMISSIONS,
        notes: null,
      };
    },
  },
];
