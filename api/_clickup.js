const CLICKUP_BASE = "https://api.clickup.com/api/v2";

function getEnv(name) {
  const value = process.env[name];
  if (!value) {
    throw new Error(`Variable d'environnement manquante: ${name}`);
  }
  return value;
}

async function clickupFetch(path) {
  const token = getEnv("CLICKUP_TOKEN");
  const response = await fetch(`${CLICKUP_BASE}${path}`, {
    headers: {
      "Authorization": token,
      "Accept": "application/json"
    }
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`ClickUp ${response.status}: ${text}`);
  }

  return response.json();
}

function normalizeTask(task) {
  return {
    id: task.id,
    name: task.name,
    description: task.description || task.markdown_description || "",
    markdown_description: task.markdown_description || "",
    status: task.status || null,
    priority: task.priority || null,
    tags: task.tags || [],
    assignees: task.assignees || [],
    date_created: task.date_created || null,
    due_date: task.due_date || null,
    url: task.url || null
  };
}

function normalizeAttachment(att) {
  return {
    id: att.id,
    title: att.title,
    url: att.url,
    size: att.size,
    date: att.date,
    extension: att.extension
  };
}

export async function getNewsTasks() {
  const listId = getEnv("CLICKUP_LIST_NEWS");
  const data = await clickupFetch(`/list/${listId}/task?include_closed=true&subtasks=true&page=0`);
  const rawTasks = data.tasks || [];

  const tasks = await Promise.all(
    rawTasks.map(async (task) => {
      let attachments = [];

      try {
        const fullTask = await clickupFetch(`/task/${task.id}`);
        attachments = (fullTask.attachments || []).map(normalizeAttachment);
      } catch {
        attachments = [];
      }

      return {
        ...normalizeTask(task),
        _attachments: attachments
      };
    })
  );

  tasks.sort((a, b) => {
    const aOpen = !["closed", "done", "completed"].includes((a.status?.status || "").toLowerCase());
    const bOpen = !["closed", "done", "completed"].includes((b.status?.status || "").toLowerCase());
    if (aOpen !== bOpen) return aOpen ? -1 : 1;
    return Number(b.date_created || 0) - Number(a.date_created || 0);
  });

  return { tasks };
}

export async function getChallengeTasks() {
  const listId = getEnv("CLICKUP_LIST_CHALLENGES");
  const data = await clickupFetch(`/list/${listId}/task?include_closed=true&subtasks=true&page=0`);
  const rawTasks = data.tasks || [];

  const tasks = await Promise.all(
    rawTasks.map(async (task) => {
      let attachments = [];

      try {
        const fullTask = await clickupFetch(`/task/${task.id}`);
        attachments = (fullTask.attachments || []).map(normalizeAttachment);
      } catch {
        attachments = [];
      }

      return {
        ...normalizeTask(task),
        _attachments: attachments
      };
    })
  );

  tasks.sort((a, b) => {
    const aOpen = !["closed", "done", "completed"].includes((a.status?.status || "").toLowerCase());
    const bOpen = !["closed", "done", "completed"].includes((b.status?.status || "").toLowerCase());
    if (aOpen !== bOpen) return aOpen ? -1 : 1;
    return Number(b.date_created || 0) - Number(a.date_created || 0);
  });

  return { tasks };
}
