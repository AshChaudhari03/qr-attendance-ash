const express = require("express");
const Database = require("better-sqlite3");
const ExcelJS = require("exceljs");

const app = express();
const db = new Database("attendance.db");

const PASSWORD = "team123";

app.use(express.json());
app.use(express.static("public"));

db.prepare(`
CREATE TABLE IF NOT EXISTS attendance (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  employee_id TEXT,
  employee_name TEXT,
  date TEXT,
  clock_in_ts TEXT,
  clock_out_ts TEXT
)
`).run();

// LOGIN
app.post("/login", (req, res) => {
  req.body.password === PASSWORD ? res.sendStatus(200) : res.sendStatus(401);
});

// SCAN
app.post("/scan", (req, res) => {
  const { qr, type } = req.body;
  const [employee_id, employee_name] = qr.split("|");
  const date = new Date().toISOString().slice(0,10);
  const now = new Date().toISOString();

  const record = db.prepare(
    "SELECT * FROM attendance WHERE employee_id=? AND date=?"
  ).get(employee_id, date);

  if (type === "IN") {
    if (record) return res.json({ error: "Already clocked in today" });

    db.prepare(`
      INSERT INTO attendance (employee_id, employee_name, date, clock_in_ts)
      VALUES (?,?,?,?)
    `).run(employee_id, employee_name, date, now);

    return res.json({ ok: `Clock In successful` });
  }

  if (type === "OUT") {
    if (!record) return res.json({ error: "Clock In not found" });
    if (record.clock_out_ts) return res.json({ error: "Already clocked out" });

    db.prepare(`
      UPDATE attendance SET clock_out_ts=? WHERE id=?
    `).run(now, record.id);

    return res.json({ ok: `Clock Out successful` });
  }
});

// TODAY LIST
app.get("/today", (req, res) => {
  const date = new Date().toISOString().slice(0,10);
  const rows = db.prepare(
    "SELECT * FROM attendance WHERE date=? ORDER BY id DESC"
  ).all(date);
  res.json(rows);
});

//Today in / out count
app.get("/stats/today", (req, res) => {
  const date = new Date().toISOString().slice(0,10);
  const data = db.prepare(`
    SELECT
      COUNT(clock_in_ts) AS inCount,
      COUNT(clock_out_ts) AS outCount
    FROM attendance WHERE date=?
  `).get(date);
  res.json(data);
});

// DELETE
app.delete("/attendance/:id", (req, res) => {
  db.prepare("DELETE FROM attendance WHERE id=?").run(req.params.id);
  res.sendStatus(200);
});

// UNDO
app.post("/undo", (req, res) => {
  db.prepare(
    "DELETE FROM attendance WHERE id=(SELECT id FROM attendance ORDER BY id DESC LIMIT 1)"
  ).run();
  res.sendStatus(200);
});

// EXPORT EXCEL
app.get("/export", async (req, res) => {
  const { from, to } = req.query;
  const rows = from && to
    ? db.prepare("SELECT * FROM attendance WHERE date BETWEEN ? AND ?").all(from,to)
    : db.prepare("SELECT * FROM attendance").all();

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Attendance");

  // ---------------- Title row ----------------
  ws.addRow([]);
  const title = ws.addRow(["Rombola Family Farms Attendance Report"]);
  title.font = { bold: true, size: 20 };
  ws.mergeCells(`A2:F2`);
  title.alignment = { vertical: 'middle', horizontal: 'center' };
  ws.addRow([]);

  // ---------------- Column headers ----------------
  const headerRow = ws.addRow(["Employee ID","Name","Date","Clock In","Clock Out","Total Hours"]);
  headerRow.font = { bold: true };
  headerRow.alignment = { horizontal: "center" };
  headerRow.eachCell(cell=>{
    cell.fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{argb:'FF27AE60'}
    };
    cell.font = { color:{argb:'FFFFFFFF'}, bold:true };
    cell.border = {
      top:{style:"thin"}, left:{style:"thin"}, bottom:{style:"thin"}, right:{style:"thin"}
    };
  });

  let totalHours = 0;

  // ---------------- Data rows ----------------
  rows.forEach(r => {
    let hours = "";
    if(r.clock_in_ts && r.clock_out_ts){
      hours = ((new Date(r.clock_out_ts) - new Date(r.clock_in_ts)) / 36e5);
      totalHours += hours;
      hours = hours.toFixed(2);
    }

    const dt = new Date(r.date);
    const formattedDate = `${String(dt.getDate()).padStart(2,'0')}-${String(dt.getMonth()+1).padStart(2,'0')}-${dt.getFullYear()}`;

    ws.addRow([
      r.employee_id,
      r.employee_name,
      formattedDate,
      r.clock_in_ts ? new Date(r.clock_in_ts).toLocaleTimeString('en-US',{hour:'2-digit',minute:'2-digit',second:'2-digit',hour12:true}) : "",
      r.clock_out_ts ? new Date(r.clock_out_ts).toLocaleTimeString('en-US',{hour:'2-digit',minute:'2-digit',second:'2-digit',hour12:true}) : "",
      hours
    ]);
  });

  // ---------------- TOTAL row ----------------
  const totalRow = ws.addRow(["","","","","TOTAL",totalHours.toFixed(2)]);
  totalRow.font = { bold:true };
  totalRow.alignment = { horizontal:"center" };

  // ---------------- Auto column width ----------------
  ws.columns.forEach(c=>{
    let maxLength = 0;
    c.eachCell({ includeEmpty:true }, cell=>{
      const len = cell.value ? cell.value.toString().length : 10;
      if(len > maxLength) maxLength = len;
    });
    c.width = maxLength + 5;
  });

  // ---------------- Send file ----------------
  res.setHeader("Content-Disposition","attachment; filename=attendance.xlsx");
  await wb.xlsx.write(res);
  res.end();
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("Running on port", PORT));
