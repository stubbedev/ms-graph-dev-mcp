import { z } from "zod";
import { ToolDefinition } from "../registry.js";

const BASE = "https://graph.microsoft.com/v1.0";

function buildHeaders(): Record<string, string> {
  return {
    Authorization: "Bearer {token}",
    "Content-Type": "application/json",
  };
}

export const tasksTools: ToolDefinition[] = [
  {
    name: "graph_tasks_list_plans",
    description: "List Planner plans for a group",
    category: "tasks",
    zodShape: {
      groupId: z.string(),
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
      };
    },
  },
  {
    name: "graph_tasks_get_plan",
    description: "Get a Planner plan by ID",
    category: "tasks",
    zodShape: {
      planId: z.string(),
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
      };
    },
  },
  {
    name: "graph_tasks_create_plan",
    description: "Create a Planner plan",
    category: "tasks",
    zodShape: {
      title: z.string(),
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
      };
    },
  },
  {
    name: "graph_tasks_list_tasks",
    description: "List tasks in a Planner plan",
    category: "tasks",
    zodShape: {
      planId: z.string(),
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
      };
    },
  },
  {
    name: "graph_tasks_get_task",
    description: "Get a Planner task by ID",
    category: "tasks",
    zodShape: {
      taskId: z.string(),
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
      };
    },
  },
  {
    name: "graph_tasks_create_task",
    description: "Create a Planner task",
    category: "tasks",
    zodShape: {
      planId: z.string(),
      title: z.string(),
      bucketId: z.string().optional(),
      assignedToIds: z.array(z.string()).optional(),
      dueDateTime: z.string().optional(),
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
      };
    },
  },
  {
    name: "graph_tasks_update_task",
    description: "Update a Planner task",
    category: "tasks",
    zodShape: {
      taskId: z.string(),
      eTag: z.string().describe("ETag from the task (required for updates)"),
      title: z.string().optional(),
      percentComplete: z.number().optional(),
      dueDateTime: z.string().optional(),
      bucketId: z.string().optional(),
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
      };
    },
  },
  {
    name: "graph_tasks_delete_task",
    description: "Delete a Planner task",
    category: "tasks",
    zodShape: {
      taskId: z.string(),
      eTag: z.string().describe("ETag from the task"),
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
      };
    },
  },
  {
    name: "graph_todo_list_lists",
    description: "List Microsoft To Do task lists",
    category: "tasks",
    zodShape: {
      userId: z.string(),
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
      };
    },
  },
  {
    name: "graph_todo_create_task",
    description: "Create a Microsoft To Do task",
    category: "tasks",
    zodShape: {
      userId: z.string(),
      listId: z.string(),
      title: z.string(),
      dueDateTime: z.string().optional(),
      importance: z.enum(["low", "normal", "high"]).optional(),
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
      };
    },
  },
];
