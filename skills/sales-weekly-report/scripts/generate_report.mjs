import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
import sharp from "sharp";

const args = parseArgs(process.argv.slice(2));
if (!args.xlsx || !args.output) {
  console.error("Usage: node generate_report.mjs --xlsx <input.xlsx> --output <report.png> [--editorial editorial.json] [--photos dir] [--audit audit.json]");
  process.exit(2);
}

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(path.resolve(args.xlsx)));
const editorial = args.editorial ? JSON.parse(await fs.readFile(path.resolve(args.editorial), "utf8")) : {};
const photosDir = path.resolve(args.photos || path.dirname(args.xlsx));
const data = readWorkbook(workbook);
const model = await prepareModel(data, editorial, photosDir);
const svg = renderSvg(model);

await fs.mkdir(path.dirname(path.resolve(args.output)), { recursive: true });
await sharp(Buffer.from(svg)).png({ compressionLevel: 9 }).toFile(path.resolve(args.output));
if (args.audit) {
  await fs.writeFile(path.resolve(args.audit), JSON.stringify({ generated_at: new Date().toISOString(), source_xlsx: path.resolve(args.xlsx), ...model, images: undefined }, null, 2), "utf8");
}
const outputHeight = (await sharp(path.resolve(args.output)).metadata()).height;
console.log(JSON.stringify({ output: path.resolve(args.output), width: 1080, height: outputHeight }, null, 2));

function parseArgs(input) {
  const out = {};
  for (let i = 0; i < input.length; i += 1) {
    if (!input[i].startsWith("--")) continue;
    out[input[i].slice(2)] = input[i + 1] && !input[i + 1].startsWith("--") ? input[++i] : true;
  }
  return out;
}

function sheetValues(workbookRef, sheetName, range) {
  return workbookRef.worksheets.getItem(sheetName).getRange(range).values;
}

function optionalSheetValues(workbookRef, sheetName, range) {
  try {
    return sheetValues(workbookRef, sheetName, range);
  } catch {
    return [];
  }
}

function readWorkbook(workbookRef) {
  const basicsRows = sheetValues(workbookRef, "02_基本信息", "A5:B14");
  const basics = Object.fromEntries(basicsRows.filter((r) => r[0]).map((r) => [String(r[0]).trim(), r[1]]));

  const teams = sheetValues(workbookRef, "03_团队业绩", "A5:F60")
    .filter((r) => r[0] && toNumber(r[1]) !== null)
    .map((r) => ({ name: clean(r[0]), weekly: toNumber(r[1]) || 0, target: toNumber(r[2]) || 0, monthly: toNumber(r[3]) || 0, previous: toNumber(r[4]) || 0, note: clean(r[5]) }));

  const institutions = optionalSheetValues(workbookRef, "03A_机构业绩", "A5:C60")
    .filter((r) => r[0] && toNumber(r[1]) !== null)
    .map((r) => ({ name: clean(r[0]), weekly: toNumber(r[1]) || 0, note: clean(r[2]) }));

  const people = sheetValues(workbookRef, "04_个人业绩", "A5:F500")
    .filter((r) => r[0] && clean(r[5]) !== "否")
    .map((r) => ({ name: clean(r[0]), team: clean(r[1]), weekly: toNumber(r[2]) || 0, orders: toNumber(r[3]) || 0, highlight: clean(r[4]) }));

  const daily = sheetValues(workbookRef, "05_每日业绩", "A5:D20")
    .filter((r) => r[0] && toNumber(r[2]) !== null)
    .map((r) => ({ date: formatDate(r[0]), weekday: clean(r[1]), value: toNumber(r[2]) || 0, target: toNumber(r[3]) || 0 }));

  const frontline = sheetValues(workbookRef, "06_一线有话说", "A5:G30")
    .filter((r) => r[0] && clean(r[6]) !== "否")
    .map((r) => ({ name: clean(r[0]), team: clean(r[1]), role: clean(r[2]), raw: clean(r[3]), emphasis: clean(r[4]), photo: clean(r[5]) }));

  const activities = sheetValues(workbookRef, "07_团队建设荣誉", "A5:H50")
    .filter((r) => r[0] && clean(r[7]) !== "否")
    .map((r) => ({ type: clean(r[0]), title: clean(r[1]), date: formatDate(r[2]), place: clean(r[3]), raw: clean(r[4]), photos: clean(r[5]).split(/[;；]/).map((x) => x.trim()).filter(Boolean), subject: clean(r[6]) }));

  return { basics, teams, institutions, people, daily, frontline, activities };
}

async function prepareModel(data, edit, photosDirRef) {
  const teams = [...data.teams].sort((a, b) => b.weekly - a.weekly);
  const institutions = [...data.institutions].sort((a, b) => b.weekly - a.weekly);
  const people = [...data.people].sort((a, b) => b.weekly - a.weekly);
  const total = sum(teams.map((x) => x.weekly));
  const target = sum(teams.map((x) => x.target));
  const previous = sum(teams.map((x) => x.previous));
  const achievement = target ? total / target : 0;
  const growth = previous ? total / previous - 1 : 0;
  const orders = sum(people.map((x) => x.orders));
  const hasFrontline = edit.frontline?.show !== false && data.frontline.length > 0;
  const frontline = data.frontline[0] || {};
  const activity = data.activities.find((x) => x.type !== "荣誉") || {};
  const honor = data.activities.find((x) => x.type === "荣誉") || null;

  const portraitFile = clean(edit.frontline?.photo) || clean(frontline.photo);
  const activityPhoto = clean(edit.team_building?.photo) || clean(edit.team_building?.photos?.[0]) || clean(activity.photos?.[0]);
  const portrait = await prepareImage(portraitFile ? path.join(photosDirRef, portraitFile) : null, 180, 240);
  const activityMeta = activityPhoto ? await sharp(path.join(photosDirRef, activityPhoto)).metadata() : null;
  const activityHeight = hasFrontline ? 300 : activityMeta ? Math.round(968 * activityMeta.height / activityMeta.width) : 726;
  const activityImage = await prepareImage(activityPhoto ? path.join(photosDirRef, activityPhoto) : null, hasFrontline ? 400 : 968, activityHeight, "contain");
  const periodStart = formatDate(data.basics["统计开始日期"]);
  const periodEnd = formatDate(data.basics["统计结束日期"]);
  const sample = clean(data.basics["是否显示示例标识"]) === "是";
  if (!/^\d{4}-\d{2}-\d{2}$/.test(periodStart) || !/^\d{4}-\d{2}-\d{2}$/.test(periodEnd) || periodStart > periodEnd) {
    throw new Error("请在基本信息中填写有效的统计开始日期和统计结束日期。");
  }
  if (!sample && !data.daily.length) {
    throw new Error("每日业绩缺少实际保费；请填写‘当日保费’，不能用当日目标代替。");
  }
  const news = Array.isArray(edit.news) ? edit.news.slice(0, 3) : [];
  if (!sample && news.length !== 3) {
    throw new Error(`正式周报必须提供 3 条统计周内资讯；当前为 ${news.length} 条。`);
  }
  const invalidNews = news.filter((item) => {
    const publicationDate = formatDate(item?.date);
    return !publicationDate || publicationDate < periodStart || publicationDate > periodEnd;
  });
  if (invalidNews.length) {
    const details = invalidNews.map((item) => `${formatDate(item?.date) || "缺少日期"}《${clean(item?.title) || "未命名资讯"}》`).join("；");
    throw new Error(`行业动态发布日期必须位于统计周 ${periodStart} 至 ${periodEnd}：${details}`);
  }
  while (news.length < 3) news.push({ tag: "待检索", date: "", title: "等待 Agent 检索本周动态", summary: "正式生成前应从监管、协会、政府或公司官方渠道选择最新信息，并重新撰写本周摘要。", source: "", url: "" });

  return {
    company: clean(data.basics["公司名称"]) || "亚太财险深圳分公司",
    reportName: clean(data.basics["周报名称"]) || "销售周报",
    issue: clean(data.basics["刊期"]) || "销售周报",
    period: `${periodStart}—${periodEnd}`,
    theme: "深耕一线，聚力业务突破",
    sample,
    total, target, previous, achievement, growth, orders,
    teams, institutions, people, daily: data.daily,
    news,
    hasFrontline,
    dataNotes: edit.data_notes || [],
    frontline: {
      ...frontline,
      quote: clean(edit.frontline?.quote) || clean(frontline.emphasis) || "准确，是速度的前提；专业，是长期服务的底气。",
      body: clean(edit.frontline?.body) || clean(frontline.raw) || "本周暂无一线分享。",
      photo: portraitFile,
    },
    activity: {
      ...activity,
      type: clean(edit.team_building?.type) || clean(activity.type) || "培训",
      date: formatDate(edit.team_building?.date) || formatDate(activity.date),
      lecturer: clean(edit.team_building?.lecturer),
      format: clean(edit.team_building?.format),
      title: clean(edit.team_building?.title) || clean(activity.title) || "团队建设动态",
      subtitle: clean(edit.team_building?.subtitle) || clean(activity.type) || "会议 / 培训通报",
      body: clean(edit.team_building?.body) || clean(activity.raw) || "本周暂无会议或培训材料。",
      photo: activityPhoto,
    },
    honor: edit.honor || honor,
    activityHeight,
    images: { portrait, activity: activityImage },
  };
}

async function prepareImage(file, width, height, fit = "cover") {
  if (!file) return null;
  try {
    const buffer = await fs.readFile(file);
    const jpg = await sharp(buffer).resize(width, height, {
      fit,
      position: "attention",
      background: { r: 247, g: 243, b: 235, alpha: 1 },
    }).jpeg({ quality: 88 }).toBuffer();
    return `data:image/jpeg;base64,${jpg.toString("base64")}`;
  } catch (error) {
    throw new Error(`无法读取照片 ${path.basename(file)}：${error.message}`);
  }
}

function renderSvg(m) {
  const W = 1080;
  const wideCompanyPhoto = !m.hasFrontline && Boolean(m.images.activity);
  const companyBodyLines = wideCompanyPhoto ? wrapText(m.activity.body, 860, 21, true).length : 0;
  const wideCompanyCardH = m.activityHeight + 132 + 68 + Math.max(0, companyBodyLines - 1) * 35 + 50;
  const companyCardY = m.hasFrontline ? 2675 : 2090;
  const pairedBodyLines = wrapText(m.activity.body, 900, 16, true).length;
  const plainBodyLines = wrapText(m.activity.body, 860, 25, true).length;
  const companyCardH = wideCompanyPhoto ? wideCompanyCardH : m.images.activity
    ? Math.max(650, 427 + Math.max(0, pairedBodyLines - 1) * 28 + 50)
    : Math.max(650, 365 + Math.max(0, plainBodyLines - 1) * 44 + 65);
  const honorY = companyCardY + companyCardH + 20;
  const contentEnd = m.honor ? honorY + 110 : companyCardY + companyCardH;
  const footerTop = Math.max(m.hasFrontline ? 3420 : 0, contentEnd + 70);
  const H = footerTop + 180;
  const red = "#C8102E";
  const deep = "#710014";
  const cream = "#F7F3EB";
  const gold = "#C6A15B";
  const ink = "#252525";
  const muted = "#6B625B";
  const white = "#FFFFFF";
  const line = "#DED6CA";
  const font = "PingFang SC,Microsoft YaHei,Noto Sans CJK SC,sans-serif";
  const s = [];
  const rect = (x, y, w, h, fill, rx = 0, stroke = "none", sw = 0) => s.push(`<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="${rx}" fill="${fill}" stroke="${stroke}" stroke-width="${sw}"/>`);
  const txt = (x, y, value, size, fill = ink, weight = 400, anchor = "start", extra = "") => s.push(`<text x="${x}" y="${y}" font-family="${font}" font-size="${size}" font-weight="${weight}" fill="${fill}" text-anchor="${anchor}" ${extra}>${esc(value)}</text>`);
  const wrap = (x, y, value, maxWidth, size, lineHeight, fill = ink, weight = 400, maxLines = 99, paragraphIndent = false) => {
    const allLines = wrapText(clean(value), maxWidth, size, paragraphIndent);
    if (allLines.length > maxLines) throw new Error(`文字排版超过 ${maxLines} 行：${clean(value).slice(0, 40)}`);
    const lines = allLines;
    lines.forEach((lineText, i) => {
      const isParagraphStart = lineText.startsWith("\u0001");
      txt(x + (isParagraphStart ? 2 * size : 0), y + i * lineHeight, isParagraphStart ? lineText.slice(1) : lineText, size, fill, weight);
    });
    return lines.length * lineHeight;
  };
  const section = (y, kicker, title, no) => {
    txt(56, y, kicker, 18, red, 800, "start", 'letter-spacing="4"');
    txt(56, y + 52, title, 40, ink, 900);
    txt(1022, y + 44, no, 56, "#DDD1C4", 700, "end", `style="font-family:Georgia,serif"`);
  };

  s.push(`<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}"><defs><linearGradient id="head" x1="0" x2="1"><stop offset="0" stop-color="${deep}"/><stop offset="1" stop-color="${red}"/></linearGradient><linearGradient id="bar" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#E43143"/><stop offset="1" stop-color="${red}"/></linearGradient><linearGradient id="photo" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#342F2C"/><stop offset="1" stop-color="#806E63"/></linearGradient></defs>`);
  rect(0, 0, W, H, cream);

  // Masthead
  rect(0, 0, W, 430, "url(#head)");
  s.push(`<circle cx="1030" cy="-80" r="280" fill="none" stroke="rgba(255,255,255,.13)" stroke-width="2"/>`);
  rect(66, 78, 72, 5, gold);
  txt(156, 91, m.company, 28, white, 750, "start", 'letter-spacing="5"');
  if (m.sample) {
    rect(832, 60, 182, 46, "none", 23, "rgba(255,255,255,.55)", 1);
    txt(923, 91, "视觉原型 · 示例数据", 16, white, 500, "middle", 'letter-spacing="2"');
  }
  txt(68, 168, "WEEKLY SALES BRIEF", 15, "#E9BCC2", 650, "start", 'letter-spacing="5"');
  txt(66, 258, m.reportName, 88, white, 900, "start", 'letter-spacing="6"');
  txt(66, 352, m.issue, 34, white, 850);
  rect(330, 313, 4, 46, gold);
  txt(366, 349, m.theme, 27, "#FBE9EB", 700, "start", 'letter-spacing="1"');
  txt(1010, 418, "WEEKLY", 116, "rgba(255,255,255,.07)", 900, "end", 'style="font-family:Arial,sans-serif" letter-spacing="-5"');

  // Industry news section
  section(500, "INDUSTRY NEWS", "行业动态", "01");
  const newsY = 595;
  m.news.forEach((n, i) => {
    const x = 56 + i * 322;
    rect(x, newsY, 306, 530, white);
    rect(x, newsY + 525, 306, 5, red);
    rect(x + 24, newsY + 25, 82, 30, cream);
    txt(x + 65, newsY + 47, n.tag || "行业动态", 14, red, 750, "middle");
    txt(x + 24, newsY + 88, formatDate(n.date), 13, muted, 500);
    wrap(x + 24, newsY + 132, n.title, 258, 21.5, 30, ink, 850, 3);
    rect(x + 24, newsY + 220, 44, 3, gold);
    wrap(x + 24, newsY + 258, n.summary, 258, 14, 21.5, "#4A4540", 400, 12, true);
  });

  // Performance section
  section(1170, "WEEKLY SNAPSHOT", "业绩通报", "02");
  rect(56, 1260, 968, 350, white);
  rect(56, 1260, 968, 5, red);
  txt(84, 1310, "本周业绩走势", 26, ink, 850);
  txt(272, 1310, "单位：万元", 14, muted, 500);
  const chartX = 100, chartY = 1358, chartW = 880, chartH = 190;
  s.push(`<line x1="${chartX}" y1="${chartY + chartH}" x2="${chartX + chartW}" y2="${chartY + chartH}" stroke="${line}"/><line x1="${chartX}" y1="${chartY}" x2="${chartX}" y2="${chartY + chartH}" stroke="${line}"/>`);
  [1, 2, 3].forEach((i) => s.push(`<line x1="${chartX}" y1="${chartY + i * 47.5}" x2="${chartX + chartW}" y2="${chartY + i * 47.5}" stroke="#EEE9E2"/>`));
  const daily = m.daily.length ? m.daily.slice(0, 7) : [{ weekday: "无数据", value: 0 }];
  const maxDaily = Math.max(1, ...daily.map((d) => d.value), ...daily.map((d) => d.target || 0));
  const step = chartW / daily.length;
  daily.forEach((d, i) => {
    const h = Math.max(2, (d.value / maxDaily) * 150);
    const x = chartX + step * i + step / 2 - 30;
    rect(x, chartY + chartH - h, 60, h, "url(#bar)", 8);
    txt(x + 30, chartY + chartH - h - 10, formatNum(d.value), 15, deep, 800, "middle");
    txt(x + 30, chartY + chartH + 30, d.weekday || d.date, 14, muted, 500, "middle");
  });

  // Team / institution / individual rankings
  rankingCard(56, 1645, "团队 TOP 3", m.teams);
  rankingCard(380, 1645, "机构 TOP 3", m.institutions, "等待填入机构数据");
  rankingCard(704, 1645, "个人 TOP 3", m.people);
  function rankingCard(x, y, title, rows, emptyText = "暂无排行数据") {
    rect(x, y, 304, 285, white);
    rect(x, y, 304, 5, red);
    txt(x + 24, y + 48, title, 23, ink, 850);
    if (!rows.length) {
      txt(x + 152, y + 164, emptyText, 16, muted, 500, "middle");
      return;
    }
    rows.slice(0, 3).forEach((row, i) => {
      const rowY = y + 103 + i * 59;
      txt(x + 24, rowY, i + 1, 24, gold, 800, "start", 'style="font-family:Georgia,serif"');
      txt(x + 62, rowY, row.name, 17, ink, 750);
      txt(x + 278, rowY, `${formatNum(row.weekly)}万`, 17, red, 850, "end");
      if (i < 2) s.push(`<line x1="${x + 24}" y1="${rowY + 22}" x2="${x + 280}" y2="${rowY + 22}" stroke="#EEE8DF"/>`);
    });
  }

  // Frontline section
  if (m.hasFrontline) {
  section(1995, "VOICE FROM THE FRONTLINE", "一线有话说", "03");
  rect(56, 2090, 968, 430, white);
  rect(56, 2090, 116, 430, deep);
  if (m.images.portrait) {
    rect(82, 2140, 184, 244, white);
    s.push(`<image href="${m.images.portrait}" x="90" y="2148" width="168" height="228" preserveAspectRatio="xMidYMid slice"/>`);
  }
  else {
    rect(82, 2140, 184, 244, "#E5DCCF");
    s.push(`<circle cx="174" cy="2240" r="50" fill="${deep}"/>`);
    txt(174, 2257, (m.frontline.name || "一").slice(0, 1), 42, white, 900, "middle");
    txt(174, 2350, "照片位", 13, muted, 500, "middle", 'letter-spacing="2"');
  }
  txt(322, 2150, m.frontline.name || "一线员工", 30, ink, 900);
  txt(492, 2150, `${m.frontline.team || ""}${m.frontline.role ? ` · ${m.frontline.role}` : ""}`, 16, muted, 600);
  wrap(322, 2218, `“${m.frontline.quote}”`, 630, 27, 39, deep, 650, 2);
  wrap(322, 2326, m.frontline.body, 630, 15, 26, "#45403B", 400, 8, true);
  }

  // Company update section
  const companySectionY = m.hasFrontline ? 2580 : 1995;
  section(companySectionY, "COMPANY UPDATE", "公司动态", m.hasFrontline ? "04" : "03");
  rect(56, companyCardY, 968, companyCardH, white);
  if (m.images.activity && !m.hasFrontline) {
    s.push(`<image href="${m.images.activity}" x="56" y="${companyCardY}" width="968" height="${m.activityHeight}" preserveAspectRatio="xMidYMid meet"/>`);
    const bandY = companyCardY + m.activityHeight;
    rect(56, bandY, 968, 132, red);
    txt(92, bandY + 53, m.activity.title, 34, white, 850);
    txt(92, bandY + 94, m.activity.subtitle, 16, "#F5D6D9", 600);
    wrap(92, bandY + 200, m.activity.body, 860, 21, 35, "#45403B", 400, 24, true);
  } else if (m.images.activity) {
    s.push(`<image href="${m.images.activity}" x="56" y="${companyCardY}" width="400" height="300" preserveAspectRatio="xMidYMid meet"/>`);
    rect(456, companyCardY, 568, 300, red);
    wrap(500, companyCardY + 105, panelTitle(m.activity.title), 470, 29, 42, white, 850, 2);
    wrap(500, companyCardY + 235, m.activity.subtitle, 470, 15, 22, "#F5D6D9", 500, 2);
    txt(92, companyCardY + 363, singleLineTitle(m.activity.title), 30, deep, 850);
    wrap(92, companyCardY + 427, m.activity.body, 900, 16, 28, "#45403B", 400, 24, true);
  } else {
    rect(56, companyCardY, 10, companyCardH, red);
    wrap(92, companyCardY + 67, m.activity.subtitle, 860, 19, 29, red, 650, 2);
    wrap(92, companyCardY + 165, m.activity.title, 860, 44, 58, deep, 900, 2);
    rect(92, companyCardY + 280, 92, 5, gold);
    wrap(92, companyCardY + 365, m.activity.body, 860, 25, 44, "#45403B", 400, 24, true);
  }
  if (m.honor) {
    rect(56, honorY, 968, 110, deep);
    txt(102, honorY + 72, "✦", 46, white, 700);
    txt(210, honorY + 52, clean(m.honor.title || "荣誉通报"), 21, white, 800);
    wrap(210, honorY + 84, clean(m.honor.body || m.honor.raw || m.honor.subject), 600, 14, 21, "#E9C9CA", 400, 2, true);
    txt(972, honorY + 56, "HONOR", 22, "#FFD992", 900, "end", 'style="font-family:Georgia,serif"');
    txt(972, honorY + 86, "BOARD", 22, "#FFD992", 900, "end", 'style="font-family:Georgia,serif"');
  }

  // Footer
  const footerY = footerTop;
  rect(0, footerY, 1080, 180, "#211E1C");
  txt(540, footerY + 72, "专业立业 · 合规立身 · 服务制胜 · 业绩争先", 25, white, 800, "middle", 'letter-spacing="7"');
  rect(518, footerY + 100, 44, 4, gold);
  s.push("</svg>");
  return s.join("");
}

function wrapText(value, maxWidth, fontSize, paragraphIndent = false) {
  if (!value) return [];
  const out = [];
  const paragraphs = paragraphIndent
    ? String(value).split(/\n+/).map((item) => item.trim()).filter(Boolean)
    : [String(value)];
  for (const paragraph of paragraphs) {
    let line = paragraphIndent ? "\u0001" : "";
    let width = paragraphIndent ? 2 * fontSize : 0;
    for (const ch of paragraph) {
      if (ch === "\n") {
        if (line.trim()) out.push(line.trimEnd());
        line = "";
        width = 0;
        continue;
      }
      const unit = /[\u2E80-\u9FFF\uF900-\uFAFF\uFF00-\uFFEF\u2010-\u2027]/.test(ch) ? 1 : /\s/.test(ch) ? 0.35 : 0.58;
      const chWidth = unit * fontSize;
      if (line && width + chWidth > maxWidth) {
        // Keep closing punctuation with preceding text and opening marks with following text.
        const carry = /[，。！？；：、）》】」』”’]/.test(ch) || /[（《【「『“‘]$/.test(line);
        if (carry && line.length > 2) {
          const last = line.slice(-1);
          out.push(line.slice(0, -1).trimEnd());
          line = last + ch;
          width = fontSize + chWidth;
        } else {
          out.push(line.trimEnd());
          line = ch;
          width = chWidth;
        }
      } else {
        line += ch;
        width += chWidth;
      }
    }
    if (line.trim()) out.push(line.trimEnd());
  }
  return out;
}

function panelTitle(value) {
  const title = clean(value);
  const parts = title.split(/\s+/).filter(Boolean);
  if (parts.length === 2 && parts.every((part) => [...part].length <= 11)) {
    return parts.join("\n");
  }
  return title;
}

function singleLineTitle(value) {
  return clean(value).replace(/\s+/g, "");
}

function esc(value) {
  return String(value ?? "").replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;");
}
function clean(value) { return value == null ? "" : String(value).trim(); }
function sum(values) { return values.reduce((a, b) => a + (Number(b) || 0), 0); }
function toNumber(value) {
  if (value == null || value === "") return null;
  const n = Number(String(value).replaceAll(",", "").replaceAll("%", ""));
  return Number.isFinite(n) ? n : null;
}
function formatNum(value, digits = 2) {
  const n = Number(value) || 0;
  return n.toLocaleString("zh-CN", { minimumFractionDigits: digits, maximumFractionDigits: digits });
}
function formatPct(value) { return `${Math.round((Number(value) || 0) * 100)}%`; }
function signedPct(value) {
  const n = Math.round((Number(value) || 0) * 1000) / 10;
  return `${n > 0 ? "+" : ""}${n.toFixed(1)}%`;
}
function formatDate(value) {
  if (!value) return "";
  if (value instanceof Date && !Number.isNaN(value.valueOf())) return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, "0")}-${String(value.getDate()).padStart(2, "0")}`;
  if (typeof value === "number" && value > 30000) {
    const date = new Date(Math.round((value - 25569) * 86400 * 1000));
    return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`;
  }
  const str = clean(value);
  const match = str.match(/(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})/);
  return match ? `${match[1]}-${match[2].padStart(2, "0")}-${match[3].padStart(2, "0")}` : str;
}
