// Example: a typed API client built from the generated types.
// Every path, parameter, request body and response below is checked
// against openapi/openapi.yaml at compile time.
import createClient, { type Middleware } from "openapi-fetch";
import type { components, paths } from "./generated/api";

type TokenPair = components["schemas"]["TokenPair"];
type Task = components["schemas"]["Task"];

const client = createClient<paths>({ baseUrl: "http://localhost:3000/v1" });

// ---- JWT handling ----------------------------------------------------------

let tokens: TokenPair | undefined;

// Attach "Authorization: Bearer <JWT>" to every request once logged in.
const authMiddleware: Middleware = {
  onRequest({ request }) {
    if (tokens) {
      request.headers.set("Authorization", `Bearer ${tokens.accessToken}`);
    }
    return request;
  },
};
client.use(authMiddleware);

export async function login(email: string, password: string): Promise<void> {
  const { data, error } = await client.POST("/auth/login", {
    body: { email, password },
  });
  if (error) throw new Error(`login failed: ${error.title}`);
  tokens = data;
}

export async function refresh(): Promise<void> {
  if (!tokens) throw new Error("not logged in");
  const { data, error } = await client.POST("/auth/refresh", {
    body: { refreshToken: tokens.refreshToken },
  });
  if (error) throw new Error(`refresh failed: ${error.title}`);
  tokens = data;
}

// ---- API calls -------------------------------------------------------------

export async function listMyTodoTasks(assigneeId: string): Promise<Task[]> {
  const { data, error } = await client.GET("/tasks", {
    params: {
      query: { status: ["todo", "doing"], assigneeId, sort: "dueDate", perPage: 50 },
    },
  });
  if (error) throw new Error(error.title);
  return data.items;
}

export async function listMyTodoTitles(assigneeId: string): Promise<string[]> {
  const tasks = await listMyTodoTasks(assigneeId);
  return tasks.map((task) => `${task.title}（期限: ${task.dueDate ?? "なし"}）`);
}

export async function completeTask(taskId: string): Promise<Task> {
  const { data, error } = await client.PATCH("/tasks/{taskId}", {
    params: { path: { taskId } },
    body: { status: "done", dueDate: null },
    // The spec declares application/merge-patch+json for this operation.
    headers: { "Content-Type": "application/merge-patch+json" },
  });
  if (error) throw new Error(error.title);
  return data;
}

export async function describeActivities(taskId: string): Promise<string[]> {
  const { data, error } = await client.GET("/tasks/{taskId}/activities", {
    params: { path: { taskId } },
  });
  if (error) throw new Error(error.title);
  // `type` is the discriminator, so each branch narrows to its own shape.
  return data.map((activity) => {
    switch (activity.type) {
      case "comment":
        return `${activity.actor.name}: ${activity.body}`;
      case "status_changed":
        return `${activity.actor.name}: ${activity.from} -> ${activity.to}`;
      case "assignee_changed":
        return `${activity.actor.name}: ${activity.from?.name ?? "-"} -> ${activity.to?.name ?? "-"}`;
      default: {
        // Compile error here if a new activity type is added to the spec.
        const unhandled: never = activity;
        return String(unhandled);
      }
    }
  });
}
