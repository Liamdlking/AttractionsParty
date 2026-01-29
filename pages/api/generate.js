import formidable from "formidable";
import path from "path";
import archiver from "archiver";
import ExcelJS from "exceljs";
import fs from "fs";
import { generatePartySheets, generateSigns } from "../../lib/generator";

export const config = {
  api: { bodyParser: false },
};

function parseForm(req) {
  const form = formidable({ multiples: false });
  return new Promise((resolve, reject) => {
    form.parse(req, (err, fields, files) => {
      if (err) return reject(err);
      resolve({ fields, files });
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

    // Read uploaded file as a Buffer (works reliably on Vercel)
    const uploadPath = file.filepath || file.path;
    if (!uploadPath) {
      res.status(400).send("Upload did not include a readable file path.");
      return;
    }

    const data = await fs.promises.readFile(uploadPath);

    // Load Excel from buffer
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(data);

    const sheet = workbook.worksheets[0];
    if (!sheet) {
      res.status(400).send("No worksheet found in Book1.");
      return;
    }

    // Convert sheet to JSON rows using header row (assumes row 1 = headers)
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

    if (!rows.length) {
      res.status(400).send("No rows found in Book1 (after header row).");
      return;
    }

    const templatesDir = path.join(process.cwd(), "templates");

    const partyFiles = await generatePartySheets(rows, templatesDir);
    const { tagFiles, stompFiles } = await generateSigns(rows, templatesDir);

    res.setHeader("Content-Type", "application/zip");
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="TagX_Output.zip"'
    );

    const archive = archiver("zip", { zlib: { level: 9 } });

    archive.on("error", (err) => {
      res.status(500).send(String(err));
    });

    archive.pipe(res);

    for (const f of partyFiles) archive.append(f.buffer, { name: f.name });
    for (const f of tagFiles) archive.append(f.buffer, { name: f.name });
    for (const f of stompFiles) archive.append(f.buffer, { name: f.name });

    await archive.finalize();
  } catch (e) {
    res.status(500).send(`Generation failed: ${e?.message || e}`);
  }
}
