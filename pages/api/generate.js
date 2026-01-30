import formidable from "formidable";
import archiver from "archiver";
import fs from "fs";
import path from "path";
import * as XLSX from "xlsx";
import { generatePartySheets, generateSigns } from "../../lib/generator";

export const config = {
  api: { bodyParser: false },
};

function parseForm(req) {
  const form = formidable({
    multiples: false,
    keepExtensions: true,
    uploadDir: "/tmp",
  });

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

    // Optional password
    const password = (fields.password || "").toString().trim();
    if (required && password !== required) {
      res.status(401).send("Invalid password");
      return;
    }

    // Formidable can return array even with multiples:false
    const fileRaw = files.book1;
    const file = Array.isArray(fileRaw) ? fileRaw[0] : fileRaw;

    if (!file) {
      res.status(400).send("Missing upload (book1)");
      return;
    }

    const uploadPath =
      file.filepath || file.path || file.tempFilePath || file?.toJSON?.().filepath;

    if (!uploadPath) {
      res.status(400).send("Upload did not include a readable file path.");
      return;
    }

    const data = await fs.promises.readFile(uploadPath);

    // ✅ Reads XLSX, XLS, CSV
    const workbook = XLSX.read(data, {
      type: "buffer",
      cellDates: true,
    });

    const sheetName = workbook.SheetNames?.[0];
    const sheet = workbook.Sheets?.[sheetName];

    if (!sheet) {
      res.status(400).send("No sheet found in uploaded file.");
      return;
    }

    // Convert to rows using header row 1
    const rows = XLSX.utils.sheet_to_json(sheet, {
      defval: "",
      raw: true, // keep dates/serials as-is; generator handles them
    });

    if (!rows.length) {
      res.status(400).send("No data rows found (after header row).");
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

    // tidy up temp upload
    fs.promises.unlink(uploadPath).catch(() => {});
  } catch (e) {
    res.status(500).send(`Generation failed: ${e?.message || e}`);
  }
}
