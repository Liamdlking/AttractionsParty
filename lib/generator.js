const path = require("path");
const fs = require("fs");
const ExcelJS = require("exceljs");
const JSZip = require("jszip");

function partyKind(ptype) {
  if (!ptype) return null;
  const s = String(ptype).toLowerCase();
  if (s.includes("stomp")) return "stomp";
  if (s.includes("tag x") || s.includes("tagx")) return "tagx";
  return null;
}

function pizzaSplit(attendees) {
  if (!attendees) return { marg: null, pep: null, chips: null, cans: null };
  let marg, pep, chips;
  if (attendees <= 10) { marg = 3; pep = 2; chips = 4; }
  else if (attendees <= 15) { marg = 3; pep = 3; chips = 5; }
  else if (attendees <= 20) { marg = 4; pep = 3; chips = 8; }
  else if (attendees <= 25) { marg = 4; pep = 4; chips = 9; }
  else { marg = 5; pep = 5; chips = 10; }
  return { marg, pep, chips, cans: attendees };
}

function canonicalMinutes(val) {
  if (val == null) return null;
  if (val instanceof Date) return val.getHours() * 60 + val.getMinutes();
  const s0 = String(val).trim().toLowerCase();
  if (!s0) return null;
  let s = s0.replace(/\./g, ":");
  const am = s.includes("am");
  const pm = s.includes("pm");
  s = s.replace(/am|pm/g, "");
  s = s.replace(/[^0-9:]/g, "");
  if (!s) return null;
  let h, m;
  if (s.includes(":")) {
    const parts = s.split(":");
    h = parseInt(parts[0] || "0", 10);
    m = parseInt(parts[1] || "0", 10);
  } else {
    h = parseInt(s, 10);
    m = 0;
  }
  if (pm && h < 12) h += 12;
  if (am && h === 12) h = 0;
  return h * 60 + m;
}

function extractFirstName(val, upper=false) {
  if (!val) return null;
  let t = String(val).trim();
  t = t.split("(")[0].split("-")[0].trim();
  if (!t) return null;
  let first = t.split(/\s+/)[0];
  first = first.replace(/[^A-Za-z'\-]/g, "");
  if (!first) return null;
  return upper ? first.toUpperCase() : (first.charAt(0).toUpperCase() + first.slice(1).toLowerCase());
}

function buildAdditionalInfo(row) {
  const parts = [];
  const food = row["Food Any Allergies"];
  const notes = row["Food Notes (inc Allergies)"];
  const tel = row["Telephone"];
  const email = row["Email"];
  if (food && String(food).trim()) parts.push(`Food: ${String(food).trim()}`);
  if (notes && String(notes).trim()) parts.push(`Notes: ${String(notes).trim()}`);
  if (tel && String(tel).trim()) parts.push(`Tel: ${String(tel).trim()}`);
  if (email && String(email).trim()) parts.push(`Email: ${String(email).trim()}`);
  return parts.length ? parts.join(" | ") : "";
}

function groupByDate(rows, dateKey) {
  const map = new Map();
  for (const r of rows) {
    const d = r[dateKey];
    if (!d) continue;
    const dateObj = d instanceof Date ? d : new Date(d);
    if (Number.isNaN(dateObj.getTime())) continue;
    const key = dateObj.toISOString().slice(0,10);
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(r);
  }
  return map;
}

async function replaceDocxPlaceholders(docxBuffer, mapping) {
  const zip = await JSZip.loadAsync(docxBuffer);
  const docPath = "word/document.xml";
  const xml = await zip.file(docPath).async("string");
  let out = xml;
  for (const [k, v] of Object.entries(mapping)) {
    out = out.split(k).join(v);
  }
  zip.file(docPath, out);
  return await zip.generateAsync({ type: "nodebuffer" });
}

async function generatePartySheets(rows, templatesDir) {
  const templatePath = path.join(templatesDir, "PARTY SHEET TEMPLATE.xlsx");

  const wbT = new ExcelJS.Workbook();
  await wbT.xlsx.readFile(templatePath);
  const wsT = wbT.worksheets[0];

  const timeRows = new Map();
  for (let r = 4; r <= 13; r++) {
    const mins = canonicalMinutes(wsT.getCell(r, 4).value);
    if (mins != null) timeRows.set(mins, r);
  }

  const sample = rows[0] || {};
  const dateKey = sample["Date of Party"] != null ? "Date of Party" : (sample["Party Date"] != null ? "Party Date" : "Date");
  const timeKey = sample["Party Start Time"] != null ? "Party Start Time" : (sample["Party Time"] != null ? "Party Time" : "Time");

  const byDate = groupByDate(rows, dateKey);
  const outputs = [];

  for (const [dateStr, dayRows] of byDate.entries()) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(templatePath);
    const ws = wb.worksheets[0];

    for (const row of dayRows) {
      const mins = canonicalMinutes(row[timeKey]);
      if (mins == null) continue;
      const tr = timeRows.get(mins);
      if (!tr) continue;

      const ptype = row["Party Type"];
      const kind = partyKind(ptype);

      ws.getCell(tr, 16).value = buildAdditionalInfo(row);
      ws.getCell(tr, 2).value = ptype || "";
      ws.getCell(tr, 5).value = row["Name"] || "";
      ws.getCell(tr, 6).value = row["Child Details Name/Age"] || "";

      const m = String(ptype || "").match(/(\d+)/);
      const attendees = m ? parseInt(m[1], 10) : null;
      ws.getCell(tr, 7).value = attendees || "";

      ws.getCell(tr, 8).value = row["PartyLocation"] || "";

      if (kind === "tagx" && attendees) {
        const {marg, pep, chips, cans} = pizzaSplit(attendees);
        ws.getCell(tr, 9).value = marg;
        ws.getCell(tr, 10).value = pep;
        ws.getCell(tr, 11).value = chips;
        ws.getCell(tr, 12).value = cans;
      }
    }

    const buf = await wb.xlsx.writeBuffer();
    outputs.push({ name: `PartySheet_${dateStr}.xlsx`, buffer: Buffer.from(buf) });
  }

  return outputs;
}

async function generateSigns(rows, templatesDir) {
  const tagTemplate = fs.readFileSync(path.join(templatesDir, "New Tag X Name Sign 2025.docx"));
  const stompTemplate = fs.readFileSync(path.join(templatesDir, "Stompers_Template_2PP.docx"));

  const tagNames = [];
  const stompNames = [];

  for (const r of rows) {
    const kind = partyKind(r["Party Type"]);
    if (kind === "tagx") {
      const n = extractFirstName(r["Child Details Name/Age"], false);
      if (n && !tagNames.includes(n)) tagNames.push(n);
    } else if (kind === "stomp") {
      const n = extractFirstName(r["Child Details Name/Age"], true);
      if (n && !stompNames.includes(n)) stompNames.push(n);
    }
  }

  // Output as multiple docx files (one per page) to avoid docx merge issues.
  const tagFiles = [];
  for (let i=0; i<Math.max(1, tagNames.length); i+=4) {
    const chunk = tagNames.slice(i, i+4);
    const mapping = {"NAME 1": chunk[0] || "", "NAME 2": chunk[1] || "", "NAME 3": chunk[2] || "", "NAME 4": chunk[3] || ""};
    const buf = tagNames.length ? await replaceDocxPlaceholders(tagTemplate, mapping) : tagTemplate;
    tagFiles.push({ name: `TagX_Signs_${(tagFiles.length+1)}.docx`, buffer: buf });
  }

  const stompFiles = [];
  for (let i=0; i<Math.max(1, stompNames.length); i+=2) {
    const chunk = stompNames.slice(i, i+2);
    const mapping = {"NAME 1": chunk[0] || "", "NAME 2": chunk[1] || ""};
    const buf = stompNames.length ? await replaceDocxPlaceholders(stompTemplate, mapping) : stompTemplate;
    stompFiles.push({ name: `Stompers_Signs_${(stompFiles.length+1)}.docx`, buffer: buf });
  }

  return { tagFiles, stompFiles };
}

module.exports = { generatePartySheets, generateSigns };
