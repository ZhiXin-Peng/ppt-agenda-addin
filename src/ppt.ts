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

  // Slide 1
  {
    const slide = pptx.addSlide();
    slide.addText(title, { x: 0.6, y: 0.5, fontSize: 30, bold: true });

    const metaLines: string[] = [];
    if (meta.subject)  metaLines.push(`Subject: ${meta.subject}`);
    if (meta.time)     metaLines.push(`Time: ${meta.time}`);
    if (meta.location) metaLines.push(`Location: ${meta.location}`);
    if (metaLines.length) {
      slide.addText(metaLines.join("\n"), {
        x: 0.6, y: 1.2, fontSize: 16, color: "666666", lineSpacing: 18
      });
    }

    const agendaText = (agendaLines?.length
      ? agendaLines.map((t, i) => `${i + 1}. ${t}`).join("\n")
      : "1. Opening\n2. Discussion\n3. Closing");
    slide.addText(agendaText, {
      x: 0.6, y: metaLines.length ? 1.9 : 1.6, fontSize: 20, lineSpacing: 22
    });
  }

  // Slide 2
  {
    const slide = pptx.addSlide();
    slide.addText("Attendees", { x: 0.6, y: 0.5, fontSize: 30, bold: true });

    const list = (attendees?.length ? attendees.join("\n") : "N/A");
    slide.addText(list, { x: 0.6, y: 1.2, fontSize: 20, lineSpacing: 20 });
  }

  // ✅ 新版 pptxgenjs: 用对象参数
  const base64 = await pptx.write({ outputType: "base64" });
  return base64 as string;
}

/** 在 PowerPoint 中打开 Base64（新建演示文稿） */
export async function openInPowerPoint(base64: string): Promise<void> {
  const pp: any = (globalThis as any).PowerPoint || (window as any).PowerPoint;
  if (!pp || typeof pp.createPresentation !== "function") {
    throw new Error("PowerPoint.createPresentation 不可用：请在 PowerPoint 加载项环境中调用。");
  }
  await pp.createPresentation(base64);
}

/** 便捷方法：一键生成并打开 */
export async function buildAndOpenAgendaPpt(
  agendaLines: string[],
  attendees: string[],
  meta: AgendaMeta = {},
  title = "Meeting Agenda"
): Promise<void> {
  const b64 = await buildAgendaPpt(agendaLines, attendees, meta, title);
  await openInPowerPoint(b64);
}
