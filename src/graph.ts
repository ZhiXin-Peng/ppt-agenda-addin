// src/graph.ts
// ⚠️ 不要引入 'isomorphic-fetch' —— 浏览器/Office 任务窗格自带 fetch

type GraphEventWire = {
  subject?: string;
  start?: { dateTime?: string; timeZone?: string };
  end?: { dateTime?: string; timeZone?: string };
  location?: { displayName?: string };
  attendees?: Array<{ emailAddress: { name?: string; address: string }; type?: string }>;
  body?: { contentType: "Text" | "HTML"; content: string };
};

export type CalendarEvent = {
  subject: string;
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  location?: { displayName?: string };
  attendees?: Array<{ emailAddress: { name?: string; address: string }; type?: string }>;
  body?: { contentType: "Text" | "HTML"; content: string };
};

async function gfetch<T>(token: string, url: string, init?: RequestInit): Promise<T> {
  const res = await fetch(`https://graph.microsoft.com/v1.0${url}`, {
    ...init,
    headers: {
      "Authorization": `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(init?.headers ?? {})
    }
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Graph ${url} failed: ${res.status} ${text}`);
  }
  return res.json();
}

/** 未来 7 天会议（默认取 5 条） */
export async function getUpcomingEvents(token: string, top = 5): Promise<CalendarEvent[]> {
  const now = new Date();
  const end = new Date(now.getTime() + 7 * 24 * 3600 * 1000);

  const data = await gfetch<{ value: GraphEventWire[] }>(
    token,
    `/me/calendarview?startDateTime=${encodeURIComponent(now.toISOString())}` +
      `&endDateTime=${encodeURIComponent(end.toISOString())}` +
      `&$top=${top}&$orderby=start/dateTime`,
    { method: "GET", headers: { Prefer: 'outlook.timezone="UTC"' } }
  );

  return (data.value ?? []).map(ev => ({
    subject: ev.subject ?? "(无主题)",
    start: { dateTime: ev.start?.dateTime ?? "", timeZone: ev.start?.timeZone ?? "UTC" },
    end: { dateTime: ev.end?.dateTime ?? "", timeZone: ev.end?.timeZone ?? "UTC" },
    location: ev.location,
    attendees: ev.attendees,
    body: undefined
  }));
}

/** 最近 7 天随机创建 N 条事件（E5 活跃度） */
export async function seedRandomEvents(token: string, count = 5): Promise<number> {
  let created = 0;
  for (let i = 0; i < count; i++) {
    const now = new Date();
    const pastDays = Math.floor(Math.random() * 7) + 1;
    const start = new Date(now.getTime() - pastDays * 24 * 3600 * 1000);
    start.setHours(9 + Math.floor(Math.random() * 9), 0, 0, 0); // 9~17 点
    const end = new Date(start.getTime() + [30, 45, 60, 75, 90][Math.floor(Math.random() * 5)] * 60000);

    const ev: CalendarEvent = {
      subject: `Daily Sync (auto) #${i + 1}`,
      start: { dateTime: start.toISOString(), timeZone: "UTC" },
      end: { dateTime: end.toISOString(), timeZone: "UTC" },
      location: { displayName: ["Teams", "Zoom", "Room A", "Room B"][Math.floor(Math.random() * 4)] },
      attendees: [],
      body: { contentType: "Text", content: "Auto-generated for dev activity." }
    };

    const r = await gfetch<any>(token, "/me/events", { method: "POST", body: JSON.stringify(ev) });
    if (r?.id) created++;
  }
  return created;
}

/** 兼容旧命名（有人 import 过 getUpcomingMeetings 也能用） */
export const getUpcomingMeetings = getUpcomingEvents;
