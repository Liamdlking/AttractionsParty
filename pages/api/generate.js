import formidable from "formidable";
import archiver from "archiver";
import ExcelJS from "exceljs";
import fs from "fs";
import { generatePartySheets, generateSigns } from "../../lib/generator";

export const config = {
  api: { bodyParser: false },
};

function parseForm(req) {
  const form = formidable({
    multiples: false,
    fileWriteStreamHandler: () => {
      const chunks = [];
      return {
        write(chunk) { chunks.push(chunk); },
        end() { this.buffer = Buffer.concat(chunks); },
      };
    },
  });

  return new Promise((resolve, reject) => {
    form.parse(req, (err, fields, files) => {
      if (err) reject(err);
      else resolve({ fields, files, form });
    });
  });
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).send("Method not allowed");
    return;
  }

  const required = (process.env.APP_PASSWORD || "").trim();

  try {
    const { fields, files } = await parseForm(req);

    const password = (fields.password || "").toString().trim();
    if (required && password !== required) {
      res.status(401).send("Invalid password");
      return;
    }

    const file = files.book1;
    if (!file) {
      res.status(400).send("Missing Book1 file");
      return;
    }

    // ✅ Read buffer directly (Vercel-safe)
    const data = file._writeStream?.buffer;
    if (!data) {
      res.status(400).send("Could not read uploaded file.");
      return;
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(data);

    const sheet = workbook.worksheets[0];

    const headers = [];
    sheet.getRow(1).eachCell((cell, col) => {
      headers[col - 1] = String(cell.value || "").trim();
    });

    const rows = [];
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const obj = {};
      row.eachCell((cell, col) => {
        const key = headers[col - 1];
        if (key) obj[key] = cell.value;
      });
      if (Object.keys(obj).length) rows.push(obj);
    });

    const partyFiles = await generatePartySheets(rows, "templates");
    const { tagFiles, stompFiles } = await generateSigns(rows, "templates");

    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", 'attachment; filename="TagX_Output.zip"');

    const archive = archiver("zip");
    archive.pipe(res);

    for (const f of partyFiles) archive.append(f.buffer, { name: f.name });
    for (const f of tagFiles) archive.append(f.buffer, { name: f.name });
    for (const f of stompFiles) archive.append(f.buffer, { name: f.name });

    await archive.finalize();

  } catch (e) {
    res.status(500).send(`Generation failed: ${e.message}`);
  }
}
