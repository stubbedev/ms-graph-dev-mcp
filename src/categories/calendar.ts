import { z } from "zod";
import { ToolDefinition } from "../registry.js";

const BASE = "https://graph.microsoft.com/v1.0";

function buildHeaders(): Record<string, string> {
  return {
    Authorization: "Bearer {token}",
    "Content-Type": "application/json",
  };
}

export const calendarTools: ToolDefinition[] = [
  {
    name: "graph_calendar_list_events",
    description: "List calendar events for a user",
    category: "calendar",
    zodShape: {
      userId: z.string(),
      startDateTime: z.string().optional().describe("ISO 8601 datetime"),
      endDateTime: z.string().optional().describe("ISO 8601 datetime"),
      top: z.number().optional(),
    },
    handler: (args: { userId: string; startDateTime?: string; endDateTime?: string; top?: number }) => {
      const params: string[] = [];
      if (args.startDateTime) params.push(`startDateTime=${args.startDateTime}`);
      if (args.endDateTime) params.push(`endDateTime=${args.endDateTime}`);
      if (args.top) params.push(`$top=${args.top}`);
      const qs = params.length ? `?${params.join("&")}` : "";
      const endpoint = `${BASE}/users/${args.userId}/calendarView${qs}`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { userId: args.userId },
        queryParams: {},
        body: null,
        description: `List calendar events for user ${args.userId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-list-calendarview",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst data = await response.json();`,
      };
    },
  },
  {
    name: "graph_calendar_get_event",
    description: "Get a specific calendar event",
    category: "calendar",
    zodShape: {
      userId: z.string(),
      eventId: z.string(),
    },
    handler: (args: { userId: string; eventId: string }) => {
      const endpoint = `${BASE}/users/${args.userId}/events/${args.eventId}`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { userId: args.userId, eventId: args.eventId },
        queryParams: {},
        body: null,
        description: `Get event ${args.eventId} for user ${args.userId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/event-get",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst event = await response.json();`,
      };
    },
  },
  {
    name: "graph_calendar_create_event",
    description: "Create a calendar event",
    category: "calendar",
    zodShape: {
      userId: z.string(),
      subject: z.string(),
      startDateTime: z.string(),
      endDateTime: z.string(),
      timeZone: z.string().optional(),
      attendees: z.array(z.string()).optional(),
      bodyContent: z.string().optional(),
      location: z.string().optional(),
      isOnlineMeeting: z.boolean().optional(),
    },
    handler: (args: {
      userId: string;
      subject: string;
      startDateTime: string;
      endDateTime: string;
      timeZone?: string;
      attendees?: string[];
      bodyContent?: string;
      location?: string;
      isOnlineMeeting?: boolean;
    }) => {
      const tz = args.timeZone ?? "UTC";
      const endpoint = `${BASE}/users/${args.userId}/events`;
      const body: Record<string, unknown> = {
        subject: args.subject,
        start: { dateTime: args.startDateTime, timeZone: tz },
        end: { dateTime: args.endDateTime, timeZone: tz },
      };
      if (args.attendees?.length) {
        body.attendees = args.attendees.map((e) => ({
          emailAddress: { address: e },
          type: "required",
        }));
      }
      if (args.bodyContent) body.body = { contentType: "HTML", content: args.bodyContent };
      if (args.location) body.location = { displayName: args.location };
      if (args.isOnlineMeeting !== undefined) body.isOnlineMeeting = args.isOnlineMeeting;
      return {
        endpoint,
        method: "POST",
        headers: buildHeaders(),
        pathParams: { userId: args.userId },
        queryParams: {},
        body,
        description: `Create calendar event '${args.subject}' for user ${args.userId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-post-events",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'POST',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(body, null, 2)})\n});\nconst event = await response.json();`,
      };
    },
  },
  {
    name: "graph_calendar_update_event",
    description: "Update a calendar event",
    category: "calendar",
    zodShape: {
      userId: z.string(),
      eventId: z.string(),
      subject: z.string().optional(),
      startDateTime: z.string().optional(),
      endDateTime: z.string().optional(),
      timeZone: z.string().optional(),
      location: z.string().optional(),
      bodyContent: z.string().optional(),
      isOnlineMeeting: z.boolean().optional(),
    },
    handler: (args: {
      userId: string;
      eventId: string;
      subject?: string;
      startDateTime?: string;
      endDateTime?: string;
      timeZone?: string;
      location?: string;
      bodyContent?: string;
      isOnlineMeeting?: boolean;
    }) => {
      const endpoint = `${BASE}/users/${args.userId}/events/${args.eventId}`;
      const tz = args.timeZone ?? "UTC";
      const body: Record<string, unknown> = {};
      if (args.subject) body.subject = args.subject;
      if (args.startDateTime) body.start = { dateTime: args.startDateTime, timeZone: tz };
      if (args.endDateTime) body.end = { dateTime: args.endDateTime, timeZone: tz };
      if (args.location) body.location = { displayName: args.location };
      if (args.bodyContent) body.body = { contentType: "HTML", content: args.bodyContent };
      if (args.isOnlineMeeting !== undefined) body.isOnlineMeeting = args.isOnlineMeeting;
      return {
        endpoint,
        method: "PATCH",
        headers: buildHeaders(),
        pathParams: { userId: args.userId, eventId: args.eventId },
        queryParams: {},
        body,
        description: `Update event ${args.eventId} for user ${args.userId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/event-update",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'PATCH',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(body, null, 2)})\n});\nconst updated = await response.json();`,
      };
    },
  },
  {
    name: "graph_calendar_delete_event",
    description: "Delete a calendar event",
    category: "calendar",
    zodShape: {
      userId: z.string(),
      eventId: z.string(),
    },
    handler: (args: { userId: string; eventId: string }) => {
      const endpoint = `${BASE}/users/${args.userId}/events/${args.eventId}`;
      return {
        endpoint,
        method: "DELETE",
        headers: buildHeaders(),
        pathParams: { userId: args.userId, eventId: args.eventId },
        queryParams: {},
        body: null,
        description: `Delete event ${args.eventId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/event-delete",
        codeExample: `await fetch('${endpoint}', {\n  method: 'DELETE',\n  headers: { 'Authorization': 'Bearer {token}' }\n});`,
      };
    },
  },
  {
    name: "graph_calendar_find_meeting_times",
    description: "Find available meeting times for a set of attendees",
    category: "calendar",
    zodShape: {
      userId: z.string(),
      attendees: z.array(z.string()),
      duration: z.string().describe("ISO 8601 duration, e.g. PT1H"),
      timeConstraints: z.record(z.unknown()).optional(),
    },
    handler: (args: { userId: string; attendees: string[]; duration: string; timeConstraints?: Record<string, unknown> }) => {
      const endpoint = `${BASE}/users/${args.userId}/findMeetingTimes`;
      const body: Record<string, unknown> = {
        attendees: args.attendees.map((e) => ({ emailAddress: { address: e }, type: "required" })),
        meetingDuration: args.duration,
      };
      if (args.timeConstraints) body.timeConstraint = args.timeConstraints;
      return {
        endpoint,
        method: "POST",
        headers: buildHeaders(),
        pathParams: { userId: args.userId },
        queryParams: {},
        body,
        description: "Find meeting time suggestions for a set of attendees.",
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-findmeetingtimes",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'POST',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(body, null, 2)})\n});\nconst suggestions = await response.json();`,
      };
    },
  },
  {
    name: "graph_calendar_get_schedule",
    description: "Get free/busy schedule for users",
    category: "calendar",
    zodShape: {
      userId: z.string(),
      schedules: z.array(z.string()),
      startDateTime: z.string(),
      endDateTime: z.string(),
    },
    handler: (args: { userId: string; schedules: string[]; startDateTime: string; endDateTime: string }) => {
      const endpoint = `${BASE}/users/${args.userId}/calendar/getSchedule`;
      const body = {
        schedules: args.schedules,
        startTime: { dateTime: args.startDateTime, timeZone: "UTC" },
        endTime: { dateTime: args.endDateTime, timeZone: "UTC" },
        availabilityViewInterval: 30,
      };
      return {
        endpoint,
        method: "POST",
        headers: buildHeaders(),
        pathParams: { userId: args.userId },
        queryParams: {},
        body,
        description: "Get free/busy schedule for the specified email addresses.",
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/calendar-getschedule",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'POST',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(body, null, 2)})\n});\nconst schedule = await response.json();`,
      };
    },
  },
];
