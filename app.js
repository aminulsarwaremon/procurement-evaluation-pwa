let criteria = [];
let scoreRows = [];
let commercialRows = [];
let finalResults = [];
let deferredPrompt = null;

const $ = (id) => document.getElementById(id);

function escapeXmlText(input) {
  return String(input ?? "").replace(/[<>&'"]/g, c => ({
    "<": "&lt;", ">": "&gt;", "&": "&amp;", "'": "&#39;", '"': "&quot;"
  }[c]));
}

function normalizeHeader(value) {
  return String(value || "").trim().toLowerCase().replace(/\s+/g, " ");
}

function getConfig() {
  return {
    techWeight: Number($("techWeight").value || 0),
    funcWeight: Number($("funcWeight").value || 0),
    commWeight: Number($("commWeight").value || 0),
    techThreshold: Number($("techThreshold").value || 0),
    funcThreshold: Number($("funcThreshold").value || 0)
  };
}

function validateWeights() {
  const c = getConfig();
  const total = c.techWeight + c.funcWeight + c.commWeight;
  const el = $("weightStatus");
  el.textContent = `Total weight: ${total}`;
  el.className = "status " + (total === 100 ? "good" : "bad");
  return total === 100;
}

["techWeight", "funcWeight", "commWeight"].forEach(id => $(id).addEventListener("input", validateWeights));
validateWeights();

function extractTextFromCsv(text) {
  return text.split(/\r?\n/).map(line => line.split(",").join(" ")).join("\n");
}

async function parseScopeFile(file) {
  const name = file.name.toLowerCase();

  if (name.endsWith(".txt") || name.endsWith(".csv")) {
    const text = await file.text();
    return name.endsWith(".csv") ? extractTextFromCsv(text) : text;
  }

  if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer);
    let out = [];
    wb.SheetNames.forEach(sheet => {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1 });
      rows.forEach(r => out.push(r.join(" ")));
    });
    return out.join("\n");
  }

  if (name.endsWith(".docx")) {
    const buffer = await file.arrayBuffer();
    const result = await mammoth.extractRawText({ arrayBuffer: buffer });
    return result.value;
  }

  if (name.endsWith(".pptx")) {
    const buffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(buffer);
    const slideFiles = Object.keys(zip.files).filter(x => /^ppt\/slides\/slide\d+\.xml$/.test(x)).sort();
    const texts = [];
    for (const sf of slideFiles) {
      const xml = await zip.files[sf].async("text");
      const matches = [...xml.matchAll(/<a:t>(.*?)<\/a:t>/g)].map(m => m[1]);
      texts.push(matches.join(" "));
    }
    return texts.join("\n");
  }

  if (name.endsWith(".pdf")) {
    if (!window.pdfjsLib) {
      throw new Error("PDF parser did not load. Check your internet connection and retry.");
    }
    const buffer = await file.arrayBuffer();
    const pdf = await window.pdfjsLib.getDocument({ data: buffer }).promise;
    const out = [];
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      out.push(content.items.map(item => item.str).join(" "));
    }
    return out.join("\n");
  }

  throw new Error("Unsupported file type.");
}

$("scopeFile").addEventListener("change", async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  $("scopeText").value = "Reading file...";
  try {
    $("scopeText").value = await parseScopeFile(file);
  } catch (err) {
    $("scopeText").value = "";
    alert(err.message);
  }
});

function classifyLine(line) {
  const l = line.toLowerCase();
  const technicalWords = ["server", "storage", "network", "database", "security", "encryption", "api", "integration", "architecture", "availability", "backup", "disaster", "performance", "latency", "throughput", "cloud", "on premise", "license", "hardware", "software", "endpoint", "firewall", "siem"];
  const functionalWords = ["workflow", "approval", "report", "dashboard", "user", "role", "process", "customer", "case", "service", "module", "business", "functional", "form", "screen", "notification", "audit trail"];
  const techHit = technicalWords.some(w => l.includes(w));
  const funcHit = functionalWords.some(w => l.includes(w));
  if (techHit && !funcHit) return "Technical";
  if (funcHit && !techHit) return "Functional";
  if (l.includes("shall") || l.includes("must") || l.includes("system")) return "Technical";
  return "Functional";
}

function cleanLine(line) {
  return line.replace(/^[\s•●▪◦\d.\)\(]+/, "").replace(/\s+/g, " ").trim();
}

function generateCriteria() {
  const text = $("scopeText").value;
  const rawLines = text.split(/\r?\n|;/).map(cleanLine).filter(x => x.length > 18);
  const unique = [];
  const seen = new Set();

  rawLines.forEach(line => {
    const compact = line.toLowerCase().slice(0, 130);
    if (!seen.has(compact)) {
      seen.add(compact);
      unique.push(line);
    }
  });

  criteria = unique.slice(0, 160).map((line, index) => ({
    id: `C${String(index + 1).padStart(3, "0")}`,
    type: classifyLine(line),
    parameter: line,
    maxMark: 10,
    guidance: "Fully compliant: full mark. Partial compliance: evaluator judgement. Non compliant: zero."
  }));

  if (criteria.length === 0) {
    criteria = [
      { id: "C001", type: "Technical", parameter: "Core technical compliance with the stated scope", maxMark: 10, guidance: "Assess against submitted technical proposal." },
      { id: "C002", type: "Functional", parameter: "Functional fit with business requirements", maxMark: 10, guidance: "Assess against user journey and process requirements." }
    ];
  }

  renderCriteria();
}

function renderCriteria() {
  const tbody = $("criteriaTable").querySelector("tbody");
  tbody.innerHTML = "";
  criteria.forEach((c, idx) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><input value="${escapeXmlText(c.id)}" data-field="id" data-idx="${idx}"></td>
      <td>
        <select data-field="type" data-idx="${idx}">
          <option ${c.type === "Technical" ? "selected" : ""}>Technical</option>
          <option ${c.type === "Functional" ? "selected" : ""}>Functional</option>
        </select>
      </td>
      <td><input value="${escapeXmlText(c.parameter)}" data-field="parameter" data-idx="${idx}"></td>
      <td><input type="number" value="${Number(c.maxMark)}" data-field="maxMark" data-idx="${idx}"></td>
      <td><input value="${escapeXmlText(c.guidance)}" data-field="guidance" data-idx="${idx}"></td>
      <td><button class="secondary" data-delete="${idx}">Delete</button></td>
    `;
    tbody.appendChild(tr);
  });

  tbody.querySelectorAll("input, select").forEach(el => {
    el.addEventListener("input", () => {
      const idx = Number(el.dataset.idx);
      const field = el.dataset.field;
      criteria[idx][field] = field === "maxMark" ? Number(el.value || 0) : el.value;
    });
  });

  tbody.querySelectorAll("[data-delete]").forEach(btn => {
    btn.addEventListener("click", () => {
      criteria.splice(Number(btn.dataset.delete), 1);
      renderCriteria();
    });
  });
}

$("extractCriteriaBtn").addEventListener("click", generateCriteria);
$("clearScopeBtn").addEventListener("click", () => { $("scopeText").value = ""; criteria = []; renderCriteria(); });
$("addCriterionBtn").addEventListener("click", () => {
  criteria.push({
    id: `C${String(criteria.length + 1).padStart(3, "0")}`,
    type: "Technical",
    parameter: "New evaluation parameter",
    maxMark: 10,
    guidance: "Define scoring guidance."
  });
  renderCriteria();
});

function workbookFromRows(sheets) {
  const wb = XLSX.utils.book_new();
  Object.entries(sheets).forEach(([name, rows]) => {
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  });
  return wb;
}

function exportBlankScoreSheet() {
  if (!criteria.length) {
    alert("Please generate or add criteria first.");
    return;
  }
  const config = getConfig();
  const scoreTemplate = criteria.map(c => ({
    Vendor: "",
    Type: c.type,
    "Criterion ID": c.id,
    Parameter: c.parameter,
    "Max Mark": c.maxMark,
    "Obtained Mark": "",
    Compliance: "",
    Comment: ""
  }));

  const commercialTemplate = [{ Vendor: "", "Bid Price": "" }];

  const criteriaRows = criteria.map(c => ({
    ID: c.id, Type: c.type, Parameter: c.parameter, "Max Mark": c.maxMark, Guidance: c.guidance
  }));

  const configRows = [
    { Item: "Technical Weight", Value: config.techWeight },
    { Item: "Functional Weight", Value: config.funcWeight },
    { Item: "Commercial Weight", Value: config.commWeight },
    { Item: "Technical Threshold", Value: config.techThreshold },
    { Item: "Functional Threshold", Value: config.funcThreshold }
  ];

  const wb = workbookFromRows({
    "Configuration": configRows,
    "Criteria": criteriaRows,
    "Score Sheet": scoreTemplate,
    "Commercial Offers": commercialTemplate
  });
  XLSX.writeFile(wb, "blank_evaluation_score_sheet.xlsx");
}

$("exportBlankBtn").addEventListener("click", exportBlankScoreSheet);

async function parseTableFile(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith(".csv")) {
    const text = await file.text();
    const wb = XLSX.read(text, { type: "string" });
    return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
  }
  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer);
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
}

$("scoreFile").addEventListener("change", async e => {
  const file = e.target.files[0];
  if (!file) return;
  scoreRows = await parseTableFile(file);
  alert(`Loaded ${scoreRows.length} score rows.`);
});

$("commercialFile").addEventListener("change", async e => {
  const file = e.target.files[0];
  if (!file) return;
  commercialRows = await parseTableFile(file);
  alert(`Loaded ${commercialRows.length} commercial rows.`);
});

function readValue(row, candidates) {
  const keys = Object.keys(row);
  for (const c of candidates) {
    const key = keys.find(k => normalizeHeader(k) === normalizeHeader(c));
    if (key) return row[key];
  }
  return "";
}

function loadSampleScores() {
  if (!criteria.length) generateCriteria();
  const vendors = ["Vendor A", "Vendor B", "Vendor C"];
  scoreRows = [];
  vendors.forEach((v, vi) => {
    criteria.forEach((c, ci) => {
      const base = vi === 0 ? 8 : vi === 1 ? 7 : 6;
      const mark = Math.max(0, Math.min(Number(c.maxMark), base + ((ci + vi) % 3) - 1));
      scoreRows.push({
        Vendor: v,
        Type: c.type,
        "Criterion ID": c.id,
        Parameter: c.parameter,
        "Max Mark": c.maxMark,
        "Obtained Mark": mark,
        Compliance: mark === Number(c.maxMark) ? "Fully compliant" : "Partial",
        Comment: "Sample score"
      });
    });
  });
  commercialRows = [
    { Vendor: "Vendor A", "Bid Price": 1050000 },
    { Vendor: "Vendor B", "Bid Price": 1000000 },
    { Vendor: "Vendor C", "Bid Price": 1150000 }
  ];
  alert("Sample scores and commercial offers loaded.");
}

$("loadSampleBtn").addEventListener("click", loadSampleScores);

function runEvaluation() {
  if (!validateWeights()) {
    alert("Total weight must be 100.");
    return;
  }
  if (!scoreRows.length) {
    alert("Please upload completed score sheet or load sample scores.");
    return;
  }

  const config = getConfig();
  const byVendor = {};

  scoreRows.forEach(row => {
    const vendor = String(readValue(row, ["Vendor"]) || "").trim();
    const type = String(readValue(row, ["Type"]) || "").trim();
    const maxMark = Number(readValue(row, ["Max Mark", "Max"]) || 0);
    const obtained = Number(readValue(row, ["Obtained Mark", "Score", "Marks Obtained"]) || 0);
    if (!vendor || !type) return;

    if (!byVendor[vendor]) {
      byVendor[vendor] = { vendor, techMax: 0, techObt: 0, funcMax: 0, funcObt: 0, bidPrice: null };
    }

    if (type.toLowerCase().startsWith("tech")) {
      byVendor[vendor].techMax += maxMark;
      byVendor[vendor].techObt += obtained;
    } else if (type.toLowerCase().startsWith("func")) {
      byVendor[vendor].funcMax += maxMark;
      byVendor[vendor].funcObt += obtained;
    }
  });

  commercialRows.forEach(row => {
    const vendor = String(readValue(row, ["Vendor"]) || "").trim();
    const price = Number(readValue(row, ["Bid Price", "Price", "Commercial Offer", "Quoted Price"]) || 0);
    if (!vendor || !price) return;
    if (!byVendor[vendor]) byVendor[vendor] = { vendor, techMax: 0, techObt: 0, funcMax: 0, funcObt: 0, bidPrice: null };
    byVendor[vendor].bidPrice = price;
  });

  const vendors = Object.values(byVendor).map(v => {
    const techPct = v.techMax ? (v.techObt / v.techMax) * 100 : 0;
    const funcPct = v.funcMax ? (v.funcObt / v.funcMax) * 100 : 0;
    const qualified = techPct >= config.techThreshold && funcPct >= config.funcThreshold;
    return { ...v, techPct, funcPct, qualified };
  });

  const qualifiedWithPrice = vendors.filter(v => v.qualified && v.bidPrice > 0);
  const lowest = qualifiedWithPrice.length ? Math.min(...qualifiedWithPrice.map(v => v.bidPrice)) : null;

  finalResults = vendors.map(v => {
    const techWeighted = (v.techPct / 100) * config.techWeight;
    const funcWeighted = (v.funcPct / 100) * config.funcWeight;
    const commercialWeighted = v.qualified && v.bidPrice > 0 && lowest ? (lowest / v.bidPrice) * config.commWeight : 0;
    const total = techWeighted + funcWeighted + commercialWeighted;
    return {
      Vendor: v.vendor,
      "Technical %": Number(v.techPct.toFixed(2)),
      "Functional %": Number(v.funcPct.toFixed(2)),
      Qualified: v.qualified ? "Yes" : "No",
      "Technical Weighted": Number(techWeighted.toFixed(2)),
      "Functional Weighted": Number(funcWeighted.toFixed(2)),
      "Commercial Weighted": Number(commercialWeighted.toFixed(2)),
      "Total Score": Number(total.toFixed(2)),
      "Bid Price": v.bidPrice || ""
    };
  }).sort((a, b) => b["Total Score"] - a["Total Score"]).map((r, i) => ({ Rank: i + 1, ...r }));

  renderResults();
}

function renderResults() {
  const tbody = $("resultTable").querySelector("tbody");
  tbody.innerHTML = "";
  finalResults.forEach(r => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${r.Rank}</td>
      <td>${escapeXmlText(r.Vendor)}</td>
      <td>${r["Technical %"]}</td>
      <td>${r["Functional %"]}</td>
      <td>${r.Qualified}</td>
      <td>${r["Technical Weighted"]}</td>
      <td>${r["Functional Weighted"]}</td>
      <td>${r["Commercial Weighted"]}</td>
      <td><strong>${r["Total Score"]}</strong></td>
      <td>${r["Bid Price"]}</td>
    `;
    tbody.appendChild(tr);
  });

  const winner = finalResults[0];
  $("resultSummary").textContent = winner ? `Highest ranked bidder: ${winner.Vendor} with total score ${winner["Total Score"]}.` : "No valid result.";
  $("exportResultBtn").disabled = finalResults.length === 0;
}

$("runEvaluationBtn").addEventListener("click", runEvaluation);

function exportResult() {
  if (!finalResults.length) return;
  const criteriaRows = criteria.map(c => ({
    ID: c.id, Type: c.type, Parameter: c.parameter, "Max Mark": c.maxMark, Guidance: c.guidance
  }));
  const wb = workbookFromRows({
    "Final Ranking": finalResults,
    "Detailed Scores": scoreRows,
    "Commercial Offers": commercialRows,
    "Criteria": criteriaRows
  });
  XLSX.writeFile(wb, "full_evaluation_result.xlsx");
}

$("exportResultBtn").addEventListener("click", exportResult);

window.addEventListener("beforeinstallprompt", (e) => {
  e.preventDefault();
  deferredPrompt = e;
  $("installBtn").classList.remove("hidden");
});

$("installBtn").addEventListener("click", async () => {
  if (!deferredPrompt) return;
  deferredPrompt.prompt();
  await deferredPrompt.userChoice;
  deferredPrompt = null;
  $("installBtn").classList.add("hidden");
});

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("service-worker.js");
  });
}
