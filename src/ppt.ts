// src/ppt.ts
import PptxGenJS from "pptxgenjs";

export interface AgendaMeta {
  subject?: string;
  time?: string;
  location?: string;
}

/** 生成两页 PPT（议程 + 参会人），返回 Base64 字符串 */
export async function buildAgendaPpt(
  agendaLines: string[],
  attendees: string[],
  meta: AgendaMeta = {},
  title = "Meeting Agenda"
): Promise<string> {
  const pptx = new PptxGenJS();

  // Slide 1：议程页
  {
    const slide = pptx.addSlide();
    slide.addText(title, { x: 0.6, y: 0.5, fontSize: 30, bold: true });

    // 元信息
    const metaLines: string[] = [];
    if (meta.subject) metaLines.push(`Subject: ${meta.subject}`);
    if (meta.time) metaLines.push(`Time: ${meta.time}`);
    if (meta.location) metaLines.push(`Location: ${meta.location}`);

    if (metaLines.length > 0) {
      slide.addText(metaLines.join("\n"), {
        x: 0.6, y: 1.2, fontSize: 16, color: "666666", lineSpacing: 18,
      });
    }

    // 议程列表
    const agendaText = (agendaLines?.length
      ? agendaLines.map((t, i) => `${i + 1}. ${t}`).join("\n")
      : "1. Opening\n2. Discussion\n3. Closing");

    slide.addText(agendaText, {
      x: 0.6,
      y: metaLines.length ? 1.9 : 1.6,
      fontSize: 20,
      lineSpacing: 22,
    });

    // ✅ 添加简单的背景简介文字
    slide.addText("This is the meeting agenda to discuss the topics in detail and align on next steps.", {
      x: 0.6, y: 5.6, fontSize: 14, color: "888888",
    });

    // ✅ 增加议程表格
    const tableData: { text: string }[][] = [
      [{ text: "Time" }, { text: "Agenda Topic" }, { text: "Owner" }],
      [{ text: "10:00 - 10:30" }, { text: "Opening Remarks" }, { text: "Alice" }],
      [{ text: "10:30 - 11:00" }, { text: "Product Update" }, { text: "Bob" }],
      [{ text: "11:00 - 11:30" }, { text: "Q&A Session" }, { text: "Charlie" }],
    ];

    slide.addTable(tableData, {
      x: 0.6, y: 6.2, w: 9, colW: [2, 5, 2],
      fontSize: 14,
      border: { type: "solid", color: "000000" }, // 使用正确的边框配置
      fill: { color: "F7F7F7" }, // 使用正确的背景色配置
    });
  }

  // Slide 2：参会人页
  {
    const slide = pptx.addSlide();
    slide.addText("Attendees", { x: 0.6, y: 0.5, fontSize: 30, bold: true });

    const list = attendees?.length ? attendees.join("\n") : "N/A";
    slide.addText(list, { x: 0.6, y: 1.2, fontSize: 20, lineSpacing: 20 });

    // ✅ 添加参会人角色说明
    slide.addText("Please note: Each attendee will be responsible for discussing their respective topics.", {
      x: 0.6, y: 5.6, fontSize: 14, color: "888888",
    });

    // ✅ 增加参会人角色表格
    const roleTableData: { text: string }[][] = [
      [{ text: "Name" }, { text: "Role" }, { text: "Responsibilities" }],
      [{ text: "Alice" }, { text: "Host" }, { text: "Lead the meeting and set the tone." }],
      [{ text: "Bob" }, { text: "Presenter" }, { text: "Give the product update and demo." }],
      [{ text: "Charlie" }, { text: "Q&A Moderator" }, { text: "Facilitate the Q&A session and address concerns." }],
    ];

    slide.addTable(roleTableData, {
      x: 0.6, y: 2.5, w: 9, colW: [3, 3, 3],
      fontSize: 14,
      border: { type: "solid", color: "000000" }, // 使用正确的边框配置
      fill: { color: "F7F7F7" }, // 使用正确的背景色配置
    });
  }

  // 输出 base64，并修复类型错误
  const base64 = (await pptx.write({ outputType: "base64" })) as string;
  return base64;
}

/** 在 PowerPoint 中打开 Base64（新建演示文稿） */
export async function openInPowerPoint(base64: string): Promise<void> {
  const pp: any = (globalThis as any).PowerPoint || (window as any).PowerPoint;
  if (!pp || typeof pp.createPresentation !== "function") {
    throw new Error("PowerPoint.createPresentation 不可用：请在 PowerPoint 加载项环境中调用。");
  }
  await pp.createPresentation(base64);
}

/** 一键生成并打开 */
export async function buildAndOpenAgendaPpt(
  agendaLines: string[],
  attendees: string[],
  meta: AgendaMeta = {},
  title = "Meeting Agenda"
): Promise<void> {
  const base64 = await buildAgendaPpt(agendaLines, attendees, meta, title);
  await openInPowerPoint(base64);
}
