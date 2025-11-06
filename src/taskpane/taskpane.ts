// src/taskpane/taskpane.ts
import { getAccessToken } from "../auth";
import { getUpcomingEvents, type CalendarEvent } from "../graph";
import { buildAndOpenAgendaPpt } from "../ppt";

function setLog(msg: string) {
  const el = document.getElementById("log");
  if (el) el.textContent = msg;
}

async function doGenerateAndOpenPpt() {
  try {
    setLog("登录并读取你的日历...");
    const token = await getAccessToken();
    const meetings: CalendarEvent[] = await getUpcomingEvents(token, 1);

    if (!meetings.length) {
      setLog("未来7天没有会议，请创建一个。");
      return;
    }

    const m = meetings[0];
    const attendees: string[] = (m.attendees ?? [])
      .map((a: { emailAddress: { name?: string; address: string } }) =>
        a.emailAddress?.name || a.emailAddress?.address
      )
      .filter((x: string | undefined): x is string => !!x);

    const agendaLines = ["Opening", "Discussion", "Closing"];
    const meta = {
      subject: m.subject || "Untitled Meeting",
      time: `${m.start.dateTime} - ${m.end.dateTime} (${m.start.timeZone})`,
      location: m.location?.displayName || "未填写",
    };

    setLog("正在生成并打开 PowerPoint...");
    await buildAndOpenAgendaPpt(agendaLines, attendees, meta);
    setLog("✅ 已在 PowerPoint 中打开议程幻灯片！");
  } catch (e: any) {
    setLog(`❌ 出错: ${e.message || e}`);
  }
}

function wireUI() {
  const btn = document.getElementById("btnGeneratePpt");
  if (btn) btn.addEventListener("click", () => void doGenerateAndOpenPpt());
}

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(() => wireUI());
} else {
  window.addEventListener("DOMContentLoaded", wireUI);
}
