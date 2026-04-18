import { z } from "zod";
import { ToolDefinition } from "../registry.js";

const BASE = "https://graph.microsoft.com/v1.0";

function buildHeaders(): Record<string, string> {
  return {
    Authorization: "Bearer {token}",
    "Content-Type": "application/json",
  };
}

const READ_PERMISSIONS = {
  delegated: ["Calendars.Read"],
  application: ["Calendars.Read"],
};

const READWRITE_PERMISSIONS = {
  delegated: ["Calendars.ReadWrite"],
  application: ["Calendars.ReadWrite"],
};

export const calendarTools: ToolDefinition[] = [
  {
    name: "graph_calendar_list_events",
    description: "List calendar events for a user",
    category: "calendar",
    zodShape: {
      userId: z.string().describe("User ID or UPN (e.g. user@contoso.com)"),
      startDateTime: z.string().optional().describe("Start of the time window — ISO 8601 datetime (e.g. 2024-01-15T09:00:00)"),
      endDateTime: z.string().optional().describe("End of the time window — ISO 8601 datetime (e.g. 2024-01-15T17:00:00)"),
      top: z.number().optional().describe("Maximum number of events to return"),
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
        requiredPermissions: READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_calendar_get_event",
    description: "Get a specific calendar event",
    category: "calendar",
    zodShape: {
      userId: z.string().describe("User ID or UPN (e.g. user@contoso.com)"),
      eventId: z.string().describe("Calendar event ID"),
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
        requiredPermissions: READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_calendar_create_event",
    description: "Create a calendar event",
    category: "calendar",
    zodShape: {
      userId: z.string().describe("User ID or UPN (e.g. user@contoso.com)"),
      subject: z.string().describe("Event title/subject"),
      startDateTime: z.string().describe("Event start — ISO 8601 datetime (e.g. 2024-01-15T09:00:00)"),
      endDateTime: z.string().describe("Event end — ISO 8601 datetime (e.g. 2024-01-15T10:00:00)"),
      timeZone: z.string().optional().describe("IANA time zone name (e.g. 'America/New_York'). Defaults to UTC."),
      attendees: z.array(z.string()).optional().describe("List of attendee email addresses"),
      bodyContent: z.string().optional().describe("HTML content for the event body"),
      location: z.string().optional().describe("Display name of the meeting location"),
      isOnlineMeeting: z.boolean().optional().describe("Whether to create a Teams online meeting link"),
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
        requiredPermissions: READWRITE_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_calendar_update_event",
    description: "Update a calendar event",
    category: "calendar",
    zodShape: {
      userId: z.string().describe("User ID or UPN (e.g. user@contoso.com)"),
      eventId: z.string().describe("Calendar event ID to update"),
      subject: z.string().optional().describe("New event title/subject"),
      startDateTime: z.string().optional().describe("New event start — ISO 8601 datetime (e.g. 2024-01-15T09:00:00)"),
      endDateTime: z.string().optional().describe("New event end — ISO 8601 datetime (e.g. 2024-01-15T10:00:00)"),
      timeZone: z.string().optional().describe("IANA time zone name (e.g. 'America/New_York'). Defaults to UTC."),
      location: z.string().optional().describe("Display name of the meeting location"),
      bodyContent: z.string().optional().describe("HTML content for the event body"),
      isOnlineMeeting: z.boolean().optional().describe("Whether to add or remove a Teams online meeting link"),
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
        requiredPermissions: READWRITE_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_calendar_delete_event",
    description: "Delete a calendar event",
    category: "calendar",
    zodShape: {
      userId: z.string().describe("User ID or UPN (e.g. user@contoso.com)"),
      eventId: z.string().describe("Calendar event ID to delete"),
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
        requiredPermissions: READWRITE_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_calendar_find_meeting_times",
    description: "Find available meeting times for a set of attendees",
    category: "calendar",
    zodShape: {
      userId: z.string().describe("User ID or UPN of the organizer (e.g. user@contoso.com)"),
      attendees: z.array(z.string()).describe("List of attendee email addresses to find a time for"),
      duration: z.string().describe("Required meeting duration — ISO 8601 duration (e.g. 'PT1H' for 1 hour, 'PT30M' for 30 minutes)"),
      timeConstraints: z.record(z.unknown()).optional().describe("Time constraint object per Graph API schema — restricts the search window"),
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
        requiredPermissions: READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_calendar_get_schedule",
    description: "Get free/busy schedule for users",
    category: "calendar",
    zodShape: {
      userId: z.string().describe("User ID or UPN of the requesting user (e.g. user@contoso.com)"),
      schedules: z.array(z.string()).describe("List of email addresses to retrieve free/busy information for"),
      startDateTime: z.string().describe("Start of the schedule window — ISO 8601 datetime (e.g. 2024-01-15T00:00:00)"),
      endDateTime: z.string().describe("End of the schedule window — ISO 8601 datetime (e.g. 2024-01-15T23:59:59)"),
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
        requiredPermissions: READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_calendar_get_delta",
    description: "Get changes to calendar events since a previous delta token — returns only new, updated, or deleted events.",
    category: "calendar",
    zodShape: {
      userId: z.string().describe("User ID or UPN (e.g. user@contoso.com)"),
      startDateTime: z.string().optional().describe("Start of the tracked window — ISO 8601 datetime. Required on first call, ignored when deltaToken is provided."),
      endDateTime: z.string().optional().describe("End of the tracked window — ISO 8601 datetime. Required on first call, ignored when deltaToken is provided."),
      deltaToken: z.string().optional().describe("Token from a previous delta response (@odata.deltaLink). Omit for initial sync."),
    },
    handler: (args: { userId: string; startDateTime?: string; endDateTime?: string; deltaToken?: string }) => {
      let endpoint: string;
      if (args.deltaToken) {
        endpoint = args.deltaToken;
      } else {
        const params: string[] = [];
        if (args.startDateTime) params.push(`startDateTime=${args.startDateTime}`);
        if (args.endDateTime) params.push(`endDateTime=${args.endDateTime}`);
        const qs = params.length ? `?${params.join("&")}` : "";
        endpoint = `${BASE}/users/${args.userId}/calendarView/delta${qs}`;
      }
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { userId: args.userId },
        queryParams: {},
        body: null,
        description: `Get delta changes to calendar events for user ${args.userId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/event-delta",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst data = await response.json();\n// Store data['@odata.deltaLink'] for next call`,
        requiredPermissions: READ_PERMISSIONS,
        notes: "startDateTime and endDateTime are required on the initial call to define the time window. They are not needed on subsequent calls using the deltaToken. Events outside the window are not tracked.",
      };
    },
  },
];
