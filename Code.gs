// ══════════════════════════════════════════════════════════════════════
// Google Apps Script — IC Weekly Report Backend
// วิธีใช้: Copy ทั้งหมดนี้ไปวางใน Google Apps Script Editor
//          Deploy → Web App → Anyone → Copy URL ไปใส่ใน index.html
// ══════════════════════════════════════════════════════════════════════

const SPREADSHEET_ID = ""; // ใส่ ID ของ Google Sheets ที่ต้องการ หรือปล่อยว่างให้ระบบสร้างใหม่

// ── MAIN ENTRY POINT ──────────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = getSpreadsheet();
    writeWeeklySheet(ss, data);
    updateMonthlySummary(ss, data);
    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", week: data.week }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    const action = e.parameter.action || "getHistory";
    const ss = getSpreadsheet();
    if (action === "getHistory") {
      return ContentService
        .createTextOutput(JSON.stringify(getHistory(ss)))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (action === "getMonthly") {
      const month = e.parameter.month || getCurrentMonth();
      return ContentService
        .createTextOutput(JSON.stringify(getMonthlyData(ss, month)))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── GET OR CREATE SPREADSHEET ─────────────────────────────────────────
function getSpreadsheet() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  // ถ้าไม่ได้ระบุ ID ให้เปิด Spreadsheet ที่ผูกกับ Script นี้
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ── WRITE WEEKLY SHEET ────────────────────────────────────────────────
function writeWeeklySheet(ss, data) {
  const sheetName = "W_" + data.week.replace("-", "_"); // เช่น W_2568_W03
  let ws = ss.getSheetByName(sheetName);

  if (!ws) {
    ws = ss.insertSheet(sheetName);
    // ย้าย Sheet ใหม่ไปตำแหน่งที่เหมาะสม (เรียงตามสัปดาห์)
    const sheets = ss.getSheets();
    const weekSheets = sheets.filter(s => s.getName().startsWith("W_")).map(s => s.getName()).sort();
    const targetIdx = weekSheets.indexOf(sheetName);
    ss.moveActiveSheet(targetIdx + 1);
  }

  ws.clearContents();
  ws.setColumnWidth(1, 200);
  ws.setColumnWidth(2, 400);

  // ── HEADER ──
  const headerData = [
    ["I&C Weekly Report — BPK Power Plant", ""],
    ["สัปดาห์", data.weekLabel],
    ["ผู้จัดทำ", data.author],
    ["สถานะโดยรวม", data.overallStatus],
    ["หมายเหตุ", data.note || ""],
    ["Timestamp", data.timestamp],
    ["", ""],
  ];

  // ── I&C STATUS ──
  const icData = [
    ["=== สถานะโรงไฟฟ้า (I&C) ===", ""],
    ["I&C-related Events", data.icEvents],
    ["Protection System Readiness (%)", data.protection],
    ["Critical WO เกิน Due Date", data.overdueWo],
    ["Bypass Systems", data.bypass || "ไม่มี"],
    ["YTD I&C Events", data.ytd],
    ["เป้า I&C Events ทั้งปี", data.target],
    ["", ""],
  ];

  // ── ESCALATIONS ──
  const escHeader = [["=== Escalation ===", ""]];
  const escRows = (data.escalations || []).map(e =>
    [e.type + " (" + e.urgency + ")", e.detail]
  );

  // ── EVENTS ──
  const evHeader = [["", ""], ["=== Highlight Events ===", ""]];
  const evCols = [["วันที่", "ระบบ", "ประเภท", "รายละเอียด", "WO No.", "สถานะ"]];
  const evRows = (data.events || []).map(e =>
    [e.date, e.sys, e.type, e.desc, e.wo, e.status]
  );

  // ── RISKS ──
  const riskHeader = [["", ""], ["=== Risk Register Status ===", ""]];
  const riskCols = [["ระบบ", "ความเสี่ยง", "Score", "สถานะ", "เจ้าของ"]];
  const riskRows = (data.risks || []).map(r =>
    [r.sys, r.risk, r.score, r.status, r.owner]
  );

  // ── TEAM ──
  const teamHeader = [["", ""], ["=== สถานะทีมงาน ===", ""]];
  const teamRows = (data.team || []).map(p => [p.name, p.status]);

  // ── SPARES ──
  const spareHeader = [["", ""], ["=== Spare Parts ===", ""]];
  const spareCols = [["ชื่ออะไหล่", "จำนวน", "Min Stock", "สถานะ"]];
  const spareRows = (data.spares || []).map(s => [s.name, s.qty, s.min, s.status]);

  // ── PLANS ──
  const planHeader = [["", ""], ["=== แผนสัปดาห์หน้า ===", ""]];
  const planRows = (data.plans || []).map((p, i) =>
    [i + 1 + ". " + p.desc, p.owner + " [" + p.pri + "]"]
  );

  // รวมทุก Section แล้วเขียนลง Sheet
  const allData = [
    ...headerData,
    ...icData,
    ...escHeader, ...escRows,
    ...evHeader, ...evCols, ...evRows,
    ...riskHeader, ...riskCols, ...riskRows,
    ...teamHeader, ...teamRows,
    ...spareHeader, ...spareCols, ...spareRows,
    ...planHeader, ...planRows,
  ];

  if (allData.length > 0) {
    ws.getRange(1, 1, allData.length, 2).setValues(allData);
  }

  // จัดรูปแบบ Header Row
  ws.getRange("A1").setFontWeight("bold").setFontSize(13);
  ws.getRange("A1:B1").setBackground("#1A2E4A").setFontColor("#FFFFFF");

  formatSectionHeaders(ws, allData);
}

// ── FORMAT SECTION HEADERS ────────────────────────────────────────────
function formatSectionHeaders(ws, data) {
  data.forEach((row, i) => {
    if (typeof row[0] === "string" && row[0].startsWith("===")) {
      ws.getRange(i + 1, 1, 1, 2)
        .setBackground("#E0F4F6")
        .setFontWeight("bold")
        .setFontColor("#0D7C8C");
    }
  });
}

// ── UPDATE MONTHLY SUMMARY ────────────────────────────────────────────
function updateMonthlySummary(ss, data) {
  // หา Month จาก Week
  const weekStr = data.week; // "2025-W03"
  const month = getMonthFromWeek(weekStr);
  const sheetName = "Monthly_" + month;

  let ms = ss.getSheetByName(sheetName);
  if (!ms) {
    ms = ss.insertSheet(sheetName);
    // สร้าง Header สำหรับ Monthly Summary
    const monthHeaders = [
      ["Monthly Summary — " + month, "", "", "", "", "", ""],
      ["สัปดาห์", "I&C Events", "Protection %", "WO Overdue", "สถานะ", "ผู้จัดทำ", "หมายเหตุ"],
    ];
    ms.getRange(1, 1, 2, 7).setValues(monthHeaders);
    ms.getRange("A1:G1").setBackground("#1A2E4A").setFontColor("#FFFFFF").setFontWeight("bold");
    ms.getRange("A2:G2").setBackground("#E0F4F6").setFontWeight("bold");

    // สูตร Monthly Summary ที่ Auto-calculate
    ms.getRange("I1").setValue("Monthly KPI Summary");
    ms.getRange("I1").setFontWeight("bold").setBackground("#1A2E4A").setFontColor("#FFFFFF");
    ms.getRange("I2").setValue("Total I&C Events:");
    ms.getRange("J2").setFormula('=SUM(B3:B99)');
    ms.getRange("I3").setValue("Avg Protection %:");
    ms.getRange("J3").setFormula('=IFERROR(AVERAGE(C3:C99),"N/A")');
    ms.getRange("I4").setValue("Max WO Overdue:");
    ms.getRange("J4").setFormula('=IFERROR(MAX(D3:D99),"N/A")');
    ms.getRange("I5").setValue("จำนวนสัปดาห์ที่รายงาน:");
    ms.getRange("J5").setFormula('=COUNTA(A3:A99)');
  }

  // เพิ่มหรืออัปเดตแถวของสัปดาห์นี้
  const lastRow = ms.getLastRow();
  const existingWeeks = lastRow > 2
    ? ms.getRange(3, 1, lastRow - 2, 1).getValues().map(r => r[0])
    : [];
  const existingIdx = existingWeeks.indexOf(data.weekLabel);
  const targetRow = existingIdx >= 0 ? existingIdx + 3 : lastRow + 1;

  ms.getRange(targetRow, 1, 1, 7).setValues([[
    data.weekLabel,
    data.icEvents,
    data.protection,
    data.overdueWo,
    data.overallStatus,
    data.author,
    data.note || "",
  ]]);

  // ระบายสี Row ตาม status
  const rowColor = data.overallStatus === "ok" ? "#E8F5EE"
    : data.overallStatus === "warn" ? "#FEF3E2" : "#FDECEA";
  ms.getRange(targetRow, 1, 1, 7).setBackground(rowColor);
}

// ── GET HISTORY ───────────────────────────────────────────────────────
function getHistory(ss) {
  const sheets = ss.getSheets();
  const weekSheets = sheets.filter(s => s.getName().startsWith("W_")).reverse();

  return weekSheets.slice(0, 12).map(ws => {
    const vals = ws.getRange(1, 1, 10, 2).getValues();
    const toVal = (label) => {
      const row = vals.find(r => r[0] === label);
      return row ? row[1] : "";
    };
    return {
      week: ws.getName().replace("W_", "").replace("_", "-"),
      weekLabel: toVal("สัปดาห์"),
      author: toVal("ผู้จัดทำ"),
      overallStatus: toVal("สถานะโดยรวม"),
      icEvents: toVal("I&C-related Events"),
      protection: toVal("Protection System Readiness (%)"),
      overdueWo: toVal("Critical WO เกิน Due Date"),
    };
  });
}

// ── GET MONTHLY DATA ──────────────────────────────────────────────────
function getMonthlyData(ss, month) {
  const ms = ss.getSheetByName("Monthly_" + month);
  if (!ms) return { status: "not_found", month };
  const lastRow = ms.getLastRow();
  if (lastRow < 3) return { status: "empty", month, rows: [] };
  const data = ms.getRange(3, 1, lastRow - 2, 7).getValues();
  return {
    status: "ok", month,
    rows: data.map(r => ({
      weekLabel: r[0], icEvents: r[1], protection: r[2],
      overdueWo: r[3], status: r[4], author: r[5], note: r[6],
    })),
    summary: {
      totalEvents: ms.getRange("J2").getValue(),
      avgProtection: ms.getRange("J3").getValue(),
      maxOverdue: ms.getRange("J4").getValue(),
      weeksReported: ms.getRange("J5").getValue(),
    }
  };
}

// ── HELPERS ───────────────────────────────────────────────────────────
function getMonthFromWeek(weekStr) {
  // "2025-W03" → คำนวณว่า Week นั้นอยู่ในเดือนอะไร
  const [year, week] = weekStr.split("-W").map(Number);
  const d = new Date(year, 0, 1 + (week - 1) * 7);
  const thYear = year + 543;
  const monthNames = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
    "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
  return monthNames[d.getMonth()] + thYear;
}

function getCurrentMonth() {
  const now = new Date();
  const thYear = now.getFullYear() + 543;
  const monthNames = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
    "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
  return monthNames[now.getMonth()] + thYear;
}

// ── TEST FUNCTION (รันจาก Apps Script Editor เพื่อทดสอบ) ──────────────
function testWithMockData() {
  const mockData = {
    week: "2025-W03",
    weekLabel: "สัปดาห์ที่ 3 · ปี 2568",
    author: "หัวหน้าแผนก I&C",
    overallStatus: "warn",
    note: "สุเทพลา 2 วัน",
    icEvents: 2,
    protection: 83,
    overdueWo: 5,
    bypass: "Gas Detect 96HT-2 U2",
    ytd: 3,
    target: 4,
    team: [
      { name: "เพียรศิริ", status: "ok" },
      { name: "สุเทพ", status: "warn" },
      { name: "ดาวุธ", status: "ok" },
    ],
    risks: [
      { sys: "CDM/MBM", risk: "CDM Probe เสื่อม", score: 20, status: "ongoing", owner: "ดาวุธ" },
      { sys: "Gas Detect", risk: "EMI Trip", score: 20, status: "escalate", owner: "ยศธร" },
    ],
    events: [
      { date: "2025-01-15", sys: "CDM/MBM", type: "derate", desc: "Runback U2 398→368MW", wo: "WO-MBM-0147", status: "done" },
    ],
    escalations: [
      { type: "ขอ Budget", urgency: "high", detail: "Spare CDM Probe U1+U2" },
    ],
    spares: [
      { name: "CDM Probe U1+U2", qty: 0, min: 2, status: "out" },
      { name: "Gas Detect Sensor", qty: 1, min: 2, status: "low" },
    ],
    plans: [
      { desc: "KT Session W1 DCS Mark IVe", owner: "สุเทพ", pri: "high" },
      { desc: "ปิด WO MBH Stage 3 CV", owner: "ดาวุธ", pri: "med" },
    ],
    timestamp: new Date().toISOString(),
  };

  const ss = getSpreadsheet();
  writeWeeklySheet(ss, mockData);
  updateMonthlySummary(ss, mockData);
  Logger.log("Test completed successfully");
}
