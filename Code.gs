// ══════════════════════════════════════════════════════════════════════
// IC Weekly Report — Google Apps Script Backend v3
// รองรับ FormData (no-cors) + JSON
// ══════════════════════════════════════════════════════════════════════
 
const SPREADSHEET_ID = "";  // ใส่ ID ถ้ามีหลาย Spreadsheet ปล่อยว่าง = ใช้ที่ผูกกับ Script
 
// ── ROUTING ──────────────────────────────────────────────────────────
function doPost(e) {
  try {
    // รับได้ทั้ง FormData (no-cors) และ JSON
    let data;
    if (e.postData && e.postData.type === "application/json") {
      data = JSON.parse(e.postData.contents);
    } else {
      const raw = e.parameter.data || (e.parameters.data && e.parameters.data[0]) || "{}";
      data = JSON.parse(raw);
    }
 
    const ss = getSpreadsheet();
    switch (data.action) {
      case "saveDraft":      return ok(saveDraft(ss, data));
      case "mergeAndSubmit": return ok(mergeAndSubmit(ss, data.week));
      default:
        writeWeeklySheet(ss, data);
        updateMonthlySummary(ss, data);
        return ok({ status: "ok", week: data.week });
    }
  } catch (err) {
    return ok({ status: "error", message: err.message });
  }
}
 
function doGet(e) {
  try {
    const ss  = getSpreadsheet();
    const act = (e.parameter && e.parameter.action) || "getHistory";
    switch (act) {
      case "getDraft":      return ok(getDraft(ss, e.parameter.week));
      case "getHistory":    return ok(getHistory(ss));
      case "getMonthly":    return ok(getMonthlyData(ss, e.parameter.month || getCurrentMonth()));
      case "getWeekDetail": return ok(getWeekDetail(ss, e.parameter.week));
      default:              return ok({ status: "error", message: "Unknown action" });
    }
  } catch (err) {
    return ok({ status: "error", message: err.message });
  }
}
 
function ok(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
 
// ── SPREADSHEET ───────────────────────────────────────────────────────
function getSpreadsheet() {
  return SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}
 
// ══════════════════════════════════════════════════════════════════════
// DRAFT SYSTEM
// ══════════════════════════════════════════════════════════════════════
function saveDraft(ss, data) {
  const sheetName = "DRAFT_" + data.week.replace("-", "_");
  let ds = ss.getSheetByName(sheetName);
  if (!ds) {
    ds = ss.insertSheet(sheetName);
    ds.getRange("A1:F1").setValues([["timestamp","author","section","key","value","device"]]);
    ds.getRange("A1:F1").setBackground("#FEF3E2").setFontWeight("bold");
    ds.setTabColor("#E8A020");
  }
  deleteAuthorRows(ds, data.author);
  const ts   = new Date().toISOString();
  const rows = flattenData(data, ts);
  if (rows.length > 0) {
    ds.getRange(Math.max(ds.getLastRow(), 1) + 1, 1, rows.length, 6).setValues(rows);
  }
  return { status: "ok", sheet: sheetName, rows: rows.length, author: data.author };
}
 
function flattenData(data, ts) {
  const author = data.author;
  const rows   = [];
  const add    = (sec, key, val) =>
    rows.push([ts, author, sec, String(key), JSON.stringify(val), data.deviceInfo || ""]);
 
  add("meta","week",          data.week);
  add("meta","overallStatus", data.overallStatus);
  add("meta","note",          data.note || "");
  add("ic",  "icEvents",      data.icEvents   || 0);
  add("ic",  "protection",    data.protection || 100);
  add("ic",  "overdueWo",     data.overdueWo  || 0);
  add("ic",  "bypass",        data.bypass     || "");
  add("ic",  "ytd",           data.ytd        || 0);
  add("ic",  "target",        data.target     || 4);
 
  (data.team        || []).forEach((p,i) => add("team",       i+"_"+p.name, p));
  (data.risks       || []).forEach((r,i) => add("risk",       i+"_"+r.sys,  r));
  (data.events      || []).forEach((e,i) => add("event",      i,            e));
  (data.escalations || []).forEach((e,i) => add("escalation", i,            e));
  (data.spares      || []).forEach((s,i) => add("spare",      i,            s));
  (data.plans       || []).forEach((p,i) => add("plan",       i,            p));
  return rows;
}
 
function deleteAuthorRows(sheet, author) {
  const last = sheet.getLastRow();
  if (last < 2) return;
  const vals = sheet.getRange(2, 1, last - 1, 6).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i][1] === author) sheet.deleteRow(i + 2);
  }
}
 
// ── GET DRAFT ─────────────────────────────────────────────────────────
function getDraft(ss, week) {
  const sheetName = "DRAFT_" + week.replace("-", "_");
  const ds = ss.getSheetByName(sheetName);
  if (!ds || ds.getLastRow() < 2) return { status: "empty", week, contributors: [] };
 
  const rows = ds.getRange(2, 1, ds.getLastRow() - 1, 6).getValues();
  const byAuthor = {};
  rows.forEach(row => {
    const [ts, author, section, key, value] = row;
    if (!byAuthor[author]) byAuthor[author] = { author, lastUpdate: ts, sections: {} };
    if (!byAuthor[author].sections[section]) byAuthor[author].sections[section] = {};
    try { byAuthor[author].sections[section][key] = JSON.parse(value); }
    catch (e) { byAuthor[author].sections[section][key] = value; }
    if (ts > byAuthor[author].lastUpdate) byAuthor[author].lastUpdate = ts;
  });
  return { status: "ok", week, contributors: Object.values(byAuthor), rawRows: rows.length };
}
 
// ── MERGE & SUBMIT ────────────────────────────────────────────────────
function mergeAndSubmit(ss, week) {
  const draft = getDraft(ss, week);
  if (draft.status === "empty") return { status: "error", message: "ไม่พบ Draft สำหรับสัปดาห์ " + week };
  const merged = buildMergedData(draft.contributors, week);
  writeWeeklySheet(ss, merged);
  updateMonthlySummary(ss, merged);
  const ds = ss.getSheetByName("DRAFT_" + week.replace("-", "_"));
  if (ds) { ds.setName("SUBMITTED_" + week.replace("-", "_")); ds.setTabColor("#1E7A46"); }
  return { status: "ok", week, contributors: draft.contributors.map(c => c.author),
           message: "Merge สำเร็จ จาก " + draft.contributors.length + " คน" };
}
 
function buildMergedData(contributors, week) {
  const merged = {
    week, weekLabel: "สัปดาห์ " + week,
    author: "ทีม I&C (" + contributors.map(c => c.author).join(", ") + ")",
    overallStatus: "ok", note: "", icEvents: 0, protection: 100,
    overdueWo: 0, bypass: "", ytd: 0, target: 4,
    team: [], risks: [], events: [], escalations: [], spares: [], plans: [],
    timestamp: new Date().toISOString()
  };
  const priority = { danger: 2, warn: 1, ok: 0 };
  contributors.forEach(c => {
    const s = c.sections;
    if (s.meta) {
      if ((priority[s.meta.overallStatus] || 0) > (priority[merged.overallStatus] || 0))
        merged.overallStatus = s.meta.overallStatus;
      if (s.meta.note) merged.note += (merged.note ? " | " : "") + c.author + ": " + s.meta.note;
    }
    if (s.ic) {
      merged.icEvents   = Math.max(merged.icEvents,   s.ic.icEvents   || 0);
      merged.protection = Math.min(merged.protection, s.ic.protection || 100);
      merged.overdueWo  = Math.max(merged.overdueWo,  s.ic.overdueWo  || 0);
      merged.ytd        = Math.max(merged.ytd,         s.ic.ytd        || 0);
      merged.target     = s.ic.target || merged.target;
      if (s.ic.bypass) merged.bypass += (merged.bypass ? ", " : "") + s.ic.bypass;
    }
    if (s.team) Object.values(s.team).forEach(p => {
      const ex = merged.team.find(t => t.name === p.name);
      if (ex) ex.status = p.status; else merged.team.push(p);
    });
    if (s.risk) Object.values(s.risk).forEach((r, i) => {
      if (!merged.risks[i]) merged.risks[i] = r; else merged.risks[i].status = r.status;
    });
    ["event","escalation","spare","plan"].forEach(type => {
      if (s[type]) Object.values(s[type]).forEach(item => {
        if (item && typeof item === "object") merged[type + "s"].push(item);
      });
    });
  });
  return merged;
}
 
// ══════════════════════════════════════════════════════════════════════
// WEEKLY SHEET
// ══════════════════════════════════════════════════════════════════════
function writeWeeklySheet(ss, data) {
  const sheetName = "W_" + data.week.replace("-", "_");
  let ws = ss.getSheetByName(sheetName);
  if (!ws) { ws = ss.insertSheet(sheetName); ws.setTabColor("#0D7C8C"); }
  ws.clearContents();
  ws.setColumnWidth(1, 220); ws.setColumnWidth(2, 420);
 
  const rows = [
    ["I&C Weekly Report — BPK Power Plant", ""],
    ["สัปดาห์",      data.weekLabel || data.week],
    ["ผู้จัดทำ",     data.author],
    ["สถานะโดยรวม",  data.overallStatus],
    ["หมายเหตุ",     data.note || ""],
    ["Timestamp",   data.timestamp],
    ["",""],
    ["■ สถานะโรงไฟฟ้า (I&C)",""],
    ["I&C Events สัปดาห์นี้",            data.icEvents],
    ["Protection System Readiness (%)", data.protection],
    ["Critical WO เกิน Due Date",       data.overdueWo],
    ["Bypass Systems",                  data.bypass || "ไม่มี"],
    ["YTD I&C Events",                  data.ytd],
    ["เป้า I&C Events ทั้งปี",          data.target],
    ["",""],
    ["■ Escalation",""],
    ...(data.escalations||[]).map(e => [e.type+" ("+e.urgency+")", e.detail]),
    ["",""],
    ["■ Highlight Events",""],
    ["วันที่","ระบบ | ประเภท | รายละเอียด | WO | สถานะ"],
    ...(data.events||[]).map(e =>
      [e.date, [e.sys,e.type,e.desc,e.wo,e.status].filter(Boolean).join(" | ")]),
    ["",""],
    ["■ Risk Register Status",""],
    ["ระบบ","ความเสี่ยง | Score | สถานะ | เจ้าของ"],
    ...(data.risks||[]).filter(r=>r&&r.sys).map(r =>
      [r.sys, [r.risk,"Score:"+r.score,r.status,r.owner].filter(Boolean).join(" | ")]),
    ["",""],
    ["■ ทีมงาน",""],
    ...(data.team||[]).map(p => [p.name, p.status]),
    ["",""],
    ["■ Spare Parts",""],
    ["ชื่ออะไหล่","จำนวน | Min | สถานะ"],
    ...(data.spares||[]).filter(s=>s&&s.name).map(s =>
      [s.name, ["Qty:"+s.qty,"Min:"+s.min,s.status].join(" | ")]),
    ["",""],
    ["■ แผนสัปดาห์หน้า",""],
    ...(data.plans||[]).filter(p=>p&&p.desc).map((p,i) =>
      [(i+1)+". "+p.desc, p.owner+" ["+p.pri+"]"]),
  ];
 
  ws.getRange(1, 1, rows.length, 2).setValues(rows);
  ws.getRange("A1:B1").setBackground("#1A2E4A").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(12);
  rows.forEach((row, i) => {
    if (typeof row[0]==="string" && row[0].startsWith("■"))
      ws.getRange(i+1,1,1,2).setBackground("#E0F4F6").setFontWeight("bold").setFontColor("#0D7C8C");
  });
}
 
// ══════════════════════════════════════════════════════════════════════
// MONTHLY SUMMARY
// ══════════════════════════════════════════════════════════════════════
function updateMonthlySummary(ss, data) {
  const month     = getMonthFromWeek(data.week);
  const sheetName = "Monthly_" + month;
  let ms = ss.getSheetByName(sheetName);
  if (!ms) {
    ms = ss.insertSheet(sheetName);
    ms.getRange("A1:G1").setValues([["Monthly Summary — "+month,"","","","","",""]]);
    ms.getRange("A2:G2").setValues([["สัปดาห์","I&C Events","Protection %","WO Overdue","สถานะ","ผู้จัดทำ","หมายเหตุ"]]);
    ms.getRange("A1:G1").setBackground("#1A2E4A").setFontColor("#FFFFFF").setFontWeight("bold");
    ms.getRange("A2:G2").setBackground("#E0F4F6").setFontWeight("bold");
    ms.setTabColor("#1A2E4A");
    ms.getRange(1,9).setValue("Monthly KPI").setBackground("#1A2E4A").setFontColor("#FFFFFF").setFontWeight("bold");
    ms.getRange(2,9).setValue("Total I&C Events:");
    ms.getRange(2,10).setFormula("=SUM(B3:B99)");
    ms.getRange(3,9).setValue("Avg Protection %:");
    ms.getRange(3,10).setFormula("=IFERROR(AVERAGE(C3:C99),\"N/A\")");
    ms.getRange(4,9).setValue("Max WO Overdue:");
    ms.getRange(4,10).setFormula("=IFERROR(MAX(D3:D99),\"N/A\")");
    ms.getRange(5,9).setValue("สัปดาห์ที่รายงาน:");
    ms.getRange(5,10).setFormula("=COUNTA(A3:A99)");
  }
  const label   = data.weekLabel || data.week;
  const lastRow = ms.getLastRow();
  const existing = lastRow > 2 ? ms.getRange(3,1,lastRow-2,1).getValues().map(r=>r[0]) : [];
  const idx       = existing.indexOf(label);
  const targetRow = idx >= 0 ? idx + 3 : lastRow + 1;
  ms.getRange(targetRow,1,1,7).setValues([[
    label, data.icEvents, data.protection,
    data.overdueWo, data.overallStatus, data.author, data.note||""
  ]]);
  const rowColor = data.overallStatus==="ok"?"#E8F5EE":data.overallStatus==="warn"?"#FEF3E2":"#FDECEA";
  ms.getRange(targetRow,1,1,7).setBackground(rowColor);
}
 
// ══════════════════════════════════════════════════════════════════════
// HISTORY & HELPERS
// ══════════════════════════════════════════════════════════════════════
function getHistory(ss) {
  return ss.getSheets()
    .filter(s => s.getName().startsWith("W_"))
    .sort((a,b) => b.getName().localeCompare(a.getName()))
    .slice(0, 12)
    .map(ws => {
      const vals  = ws.getRange(1,1,12,2).getValues();
      const toVal = lbl => { const r=vals.find(r=>r[0]===lbl); return r?r[1]:""; };
      return {
        week:          ws.getName().replace("W_","").replace("_","-"),
        weekLabel:     toVal("สัปดาห์"),
        author:        toVal("ผู้จัดทำ"),
        overallStatus: toVal("สถานะโดยรวม"),
        icEvents:      toVal("I&C Events สัปดาห์นี้"),
        protection:    toVal("Protection System Readiness (%)"),
        overdueWo:     toVal("Critical WO เกิน Due Date"),
      };
    });
}
 
function getMonthlyData(ss, month) {
  const ms = ss.getSheetByName("Monthly_" + month);
  if (!ms) return { status: "not_found", month };
  const last = ms.getLastRow();
  if (last < 3) return { status: "empty", month, rows: [] };
  return {
    status: "ok", month,
    rows: ms.getRange(3,1,last-2,7).getValues().map(r => ({
      weekLabel:r[0], icEvents:r[1], protection:r[2],
      overdueWo:r[3], status:r[4], author:r[5], note:r[6]
    })),
    summary: {
      totalEvents:   ms.getRange("J2").getValue(),
      avgProtection: ms.getRange("J3").getValue(),
      maxOverdue:    ms.getRange("J4").getValue(),
      weeksReported: ms.getRange("J5").getValue(),
    }
  };
}
 
 
// ── GET WEEK DETAIL (อ่าน Weekly Sheet ทั้งหมดมาแสดงในเว็บ) ──────────
function getWeekDetail(ss, week) {
  const sheetName = "W_" + week.replace("-", "_");
  const ws = ss.getSheetByName(sheetName);
  if (!ws) return { status: "not_found", week };
  const last = ws.getLastRow();
  if (last < 1) return { status: "empty", week };
  const rows = ws.getRange(1, 1, last, 2).getValues()
    .map(r => [String(r[0]), String(r[1])]);
  return { status: "ok", week, rows };
}
 
function getMonthFromWeek(weekStr) {
  const [year, week] = weekStr.split("-W").map(Number);
  const d = new Date(Date.UTC(year, 0, 1 + (week-1)*7));
  const months = ["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."];
  return months[d.getMonth()] + (year + 543);
}
 
function getCurrentMonth() {
  const now = new Date();
  const months = ["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."];
  return months[now.getMonth()] + (now.getFullYear() + 543);
}
 
// ── TEST ──────────────────────────────────────────────────────────────
function testSaveDraft() {
  const ss   = getSpreadsheet();
  const mock = {
    action:"saveDraft", week:"2025-W28",
    weekLabel:"สัปดาห์ที่ 28 · ปี 2568",
    author:"เพียรศิริ", overallStatus:"warn",
    note:"ทดสอบระบบ", icEvents:1, protection:83,
    overdueWo:5, bypass:"Gas Detect 96HT-2", ytd:3, target:4,
    team:[{name:"เพียรศิริ",status:"ok"},{name:"สุเทพ",status:"warn"}],
    risks:[{sys:"CDM/MBM",risk:"CDM Probe เสื่อม",score:20,status:"ongoing",owner:"ดาวุธ"}],
    events:[{date:"2025-07-07",sys:"CDM",type:"derate",desc:"Runback U2",wo:"WO-001",status:"done"}],
    escalations:[{type:"ขอ Budget",urgency:"high",detail:"CDM Probe"}],
    spares:[{name:"CDM Probe",qty:0,min:2,status:"out"}],
    plans:[{desc:"KT Session W1",owner:"สุเทพ",pri:"high"}],
    timestamp:new Date().toISOString(), deviceInfo:"Desktop"
  };
  Logger.log(JSON.stringify(saveDraft(ss, mock)));
}
