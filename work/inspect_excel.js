const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");
const { XMLParser } = require("fast-xml-parser");

var XLSX = path.resolve(__dirname, "D587项目整车平台化第十一轮产业化通讯协议_V1.1.8_20260326.xlsx(LSCM).xlsx");
var TMP = path.resolve(__dirname, "_xlsx_tmp");

var P = new XMLParser({ ignoreNamespaces: true,
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  textNodeName: "#text",
  isArray: function(n) { return ["sheet","si","row","c","Relationship"].indexOf(n) >= 0; }
});
function setup() {
  if (fs.existsSync(TMP)) fs.rmSync(TMP, { recursive: true, force: true });
  fs.mkdirSync(TMP, { recursive: true });
  console.log("Setup done.");
}

function extract() {
  console.log("Extracting XML...");
  var files = ["xl/workbook.xml", "xl/_rels/workbook.xml.rels", "xl/sharedStrings.xml"];
  for (var i = 0; i < files.length; i++) {
    var f = files[i];
    var d = path.dirname(path.join(TMP, f));
    if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true });
    try {
      execSync("unzip -o \"" + XLSX + "\" \"" + f + "\" -d \"" + TMP + "\"", { stdio: "pipe" });
    } catch (e) {
      console.error("  Error:", f, e.message.substring(0, 60));
    }
  }
}
function readXML(rp) {
  var fp = path.join(TMP, rp);
  if (!fs.existsSync(fp)) { console.log("  NOT FOUND:", fp); return null; }
  return P.parse(fs.readFileSync(fp, "utf-8"));
}

function getSheets() {
  console.log("");
  console.log("========== STEP 1: Sheet Names ==========");
  var wb = readXML("xl/workbook.xml");
  if (!wb || !wb.workbook || !wb.workbook.sheets) { console.log("  None."); return []; }
  var ss = wb.workbook.sheets.sheet || [];
  console.log("  Count: " + ss.length);
  for (var i = 0; i < ss.length; i++) {
    var s = ss[i];
    console.log("  [" + (i+1) + "] \"" + (s["@_name"]||"?") + "\"  id=" + (s["@_sheetId"]||"?") + "  r:id=" + (s["@_r:id"]||"?"));
  }
  return ss;
}
function getRels(sheets) {
  console.log("");
  console.log("========== STEP 2: Relationships ==========");
  var rels = readXML("xl/_rels/workbook.xml.rels");
  if (!rels || !rels.Relationships) { console.log("  None."); return {}; }
  var ra = Array.isArray(rels.Relationships.Relationship) ? rels.Relationships.Relationship : [rels.Relationships.Relationship];
  var rm = {};
  console.log("  All relationships:");
  for (var i = 0; i < ra.length; i++) {
    var r = ra[i];
    rm[r["@_Id"]] = r["@_Target"] || null;
    console.log("    " + r["@_Id"] + " -> " + (r["@_Target"]||"?") + "  (" + (r["@_Type"]||"").split("/").pop() + ")");
  }
  console.log("");
  console.log("  Sheet-to-file:");
  var sm = {};
  var sa = Array.isArray(sheets) ? sheets : [sheets];
  for (var i = 0; i < sa.length; i++) {
    var s = sa[i];
    var nm = s["@_name"] || "?";
    var tg = rm[s["@_r:id"]] || "?";
    var fp = tg !== "?" ? "xl/" + tg : "?";
    sm[nm] = fp;
    console.log("    \"" + nm + "\" -> " + fp);
  }
  return sm;
}
function getStrings() {
  console.log("");
  console.log("========== STEP 3: Shared Strings ==========");
  var ss = readXML("xl/sharedStrings.xml");
  if (!ss || !ss.sst) { console.log("  Failed."); return []; }
  var sst = ss.sst;
  console.log("  count=" + (sst["@_count"]||"?") + "  uniqueCount=" + (sst["@_uniqueCount"]||"?"));
  var items = Array.isArray(sst.si) ? sst.si : (sst.si ? [sst.si] : []);
  var strs = [];
  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    var txt = "";
    if (it && it.t !== undefined && it.t !== null) {
      txt = typeof it.t === "object" ? (it.t["#text"] || "") : String(it.t);
    } else if (it && it.r) {
      var runs = Array.isArray(it.r) ? it.r : [it.r];
      for (var j = 0; j < runs.length; j++) {
        var ru = runs[j];
        if (ru && ru.t !== undefined && ru.t !== null) {
          txt += typeof ru.t === "object" ? (ru.t["#text"] || "") : String(ru.t);
        }
      }
    }
    strs.push(txt);
  }
  var n = Math.min(10, strs.length);
  console.log("  First " + n + " of " + strs.length + ":");
  for (var i = 0; i < n; i++) {
    var s = strs[i] || "";
    if (s.length > 80) s = s.substring(0,80) + "...";
    console.log("    [" + i + "] \"" + s + "\"");
  }
  if (strs.length > 10) console.log("    ... and " + (strs.length-10) + " more");
  return strs;
}
function readSheet(sname, spath, strs) {
  console.log("");
  console.log("========== STEP 4: Sheet \"" + sname + "\" ==========");
  var fp = path.join(TMP, spath);
  if (!fs.existsSync(fp)) {
    try { execSync("unzip -o \"" + XLSX + "\" \"" + spath + "\" -d \"" + TMP + "\"", { stdio: "pipe" }); }
    catch (e) { console.log("  Extract fail: " + e.message.substring(0,50)); return; }
  }
  var sh = readXML(spath);
  if (!sh || !sh.worksheet) { console.log("  Cannot parse."); return; }
  var rows = (sh.worksheet.sheetData && sh.worksheet.sheetData.row) || [];
  var ra = Array.isArray(rows) ? rows : [rows];
  console.log("  Total rows: " + ra.length);
  var n = Math.min(3, ra.length);
  for (var r = 0; r < n; r++) {
    var row = ra[r];
    var rn = row["@_r"] || "?";
    var cells = Array.isArray(row.c) ? row.c : (row.c ? [row.c] : []);
    var vals = [];
    for (var c = 0; c < cells.length; c++) {
      var cl = cells[c];
      var ref = cl["@_r"] || "?";
      var typ = cl["@_t"] || "";
      var val = "";
      if (cl.v !== undefined && cl.v !== null) {
        var raw = typeof cl.v === "object" ? (cl.v["#text"] || "") : String(cl.v);
        if (typ === "s") { val = strs[parseInt(raw)] || raw; }
        else { val = raw; }
      }
      if (val.length > 50) val = val.substring(0,50) + "...";
      vals.push(ref + "=\"" + val + "\"");
    }
    console.log("  Row " + rn + ": " + vals.join(", "));
  }
  if (ra.length > 3) console.log("  ... (" + (ra.length-3) + " more rows)");
}
function main() {
  console.log("========================================================");
  console.log("  XLSX FILE INSPECTOR (Node.js)");
  console.log("  File: " + XLSX);
  console.log("========================================================");
  setup();
  extract();
  var sheets = getSheets();
  var sheetMap = getRels(sheets);
  var strings = getStrings();
  console.log("");
  console.log("========== STEP 4: Reading Sheets ==========");
  var names = Array.isArray(sheets) ? sheets.map(function(s) { return s["@_name"]; }) : [];
  var wanted = ["Matrix", "CAN", "LIN", "ETH"];
  var toRead = [];
  for (var w = 0; w < wanted.length; w++) {
    for (var n = 0; n < names.length; n++) {
      if (names[n].toLowerCase().indexOf(wanted[w].toLowerCase()) >= 0 && toRead.indexOf(names[n]) < 0) {
        toRead.push(names[n]);
      }
    }
  }
  if (names.length > 0 && toRead.indexOf(names[0]) < 0) toRead.unshift(names[0]);
  if (names.length > 1 && toRead.indexOf(names[1]) < 0) toRead.push(names[1]);
  var slice = toRead.slice(0, 4);
  for (var i = 0; i < slice.length; i++) {
    var nm = slice[i];
    var sp = sheetMap[nm];
    if (sp && sp !== "?") readSheet(nm, sp, strings);
  }
  console.log("");
  console.log("========== Cleanup ==========");
  if (fs.existsSync(TMP)) fs.rmSync(TMP, { recursive: true, force: true });
  console.log("  Temp cleaned.");
  console.log("");
  console.log("  Done.");
}

main();