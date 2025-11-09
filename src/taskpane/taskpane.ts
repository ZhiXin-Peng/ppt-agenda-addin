// src/taskpane/taskpane.ts
import { getAccessToken } from "../auth";
import { getUpcomingEvents, seedRandomEvents, type CalendarEvent } from "../graph";
import { buildAndOpenAgendaPpt } from "../ppt";

// 显示日志信息
function setLog(msg: string) {
  const el = document.getElementById("log");
  if (el) el.textContent = msg;
}

// 创建随机会议事件
async function createRandomEvents(token: string, count: number) {
  const createdEventCount = await seedRandomEvents(token, count); // 创建随机事件
  setLog(`成功创建了 ${createdEventCount} 个随机会议`);

  return createdEventCount;
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

    // 获取刚刚生成的随机会议
    const randomMeetings = await getUpcomingEvents(token, 3);  // 获取最近的3个会议事件
    const randomAgendaLines = randomMeetings.map(meeting => `${meeting.subject} - ${meeting.start.dateTime}`);

    // 合并原有议程与随机生成的会议
    const fullAgendaLines = [...agendaLines, ...randomAgendaLines];

    // 生成PPT并打开
    await buildAndOpenAgendaPpt(fullAgendaLines, attendees, meta);
    setLog("✅ 已在 PowerPoint 中打开议程幻灯片！");
  } catch (e: any) {
    setLog(`❌ 出错: ${e.message || e}`);
  }
}

async function handleGeneratePpt() {
  try {
    const token = await getAccessToken();
    
    // 先创建随机事件
    const createdEventCount = await createRandomEvents(token, 3); // 创建 3 个随机事件

    if (createdEventCount === 0) {
      setLog("❌ 未能创建任何会议，请检查权限或网络连接。");
      return;
    }

    // 创建成功后，生成并打开PPT
    await doGenerateAndOpenPpt();

  } catch (e: any) {
    setLog(`❌ 出错: ${e.message || e}`);
  }
}

function wireUI() {
  const btn = document.getElementById("btnGeneratePpt");
  if (btn) btn.addEventListener("click", () => void handleGeneratePpt());
}

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(() => wireUI());
} else {
  window.addEventListener("DOMContentLoaded", wireUI);
}
