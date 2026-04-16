const DATA_URL = "./data/ted_results.xlsx";
const NEW_SHEET = "Agent_2";
const RESULTS_SHEET = "Agent_2_Results";

const LOCATION_FILTERS = [
  ["All", "all"],
  ["Berlin", "berlin"],
  ["Stuttgart", "stuttgart"],
  ["Hamburg", "hamburg"],
  ["Deutschland", "deutschland"],
  ["Frankreich", "frankreich"],
  ["Spanien", "spanien"],
  ["Britanien", "britanien"],
  ["Nordics & Baltics", "region_nordics_baltics"],
  ["Eastern Europe & Balkans", "region_eastern_balkans"],
  ["Central Europe", "region_central_europe"],
  ["Southern Europe", "region_southern_europe"],
  ["Global rest", "global_rest"],
];

const LOCATION_KEYWORDS = {
  berlin: ["berlin"],
  stuttgart: ["stuttgart"],
  hamburg: ["hamburg"],
  deutschland: ["deutschland", "germany"],
  frankreich: ["frankreich", "france"],
  spanien: ["spanien", "spain"],
};

const REGION_KEYWORDS = {
  britanien: ["scotland", "wales", "northern ireland", "irland", "ireland"],
  region_nordics_baltics: [
    "norwegen", "norway", "schweden", "sweden", "finnland", "finland", "dänemark", "daenemark", "denmark",
    "lettland", "latvia", "litauen", "lithuania", "estland", "estonia"
  ],
  region_eastern_balkans: [
    "polen", "poland", "ungarn", "hungary", "slowakei", "slovakia", "tschechien", "czech", "czech republic",
    "slowenien", "slovenia", "kroatien", "croatia", "serbien", "serbia", "rumänien", "rumaenien", "romania",
    "bulgarien", "bulgaria", "moldau", "moldova", "griechenland", "thessaloniki"
  ],
  region_central_europe: [
    "deutschland", "germany", "österreich", "oesterreich", "austria", "schweiz", "switzerland", "frankreich", "france",
    "belgien", "belgium", "niederlande", "netherlands", "luxemburg", "luxembourg"
  ],
  region_southern_europe: ["spanien", "spain", "portugal", "italien", "italy", "zypern", "cyprus"],
};

const state = {
  allRows: [],
  filteredRows: [],
  selectedKeys: new Set(),
  filters: {
    type: "all",
    location: "all",
    category: "all",
    score: "all",
    query: "",
  },
};

const dom = {
  loadStatus: document.getElementById("loadStatus"),
  loadWarnings: document.getElementById("loadWarnings"),
  searchInput: document.getElementById("searchInput"),
  typeFilter: document.getElementById("typeFilter"),
  locationFilter: document.getElementById("locationFilter"),
  categoryFilter: document.getElementById("categoryFilter"),
  scoreFilter: document.getElementById("scoreFilter"),
  clearFiltersBtn: document.getElementById("clearFiltersBtn"),
  clearSelectionBtn: document.getElementById("clearSelectionBtn"),
  submitBtn: document.getElementById("submitBtn"),
  tableMeta: document.getElementById("tableMeta"),
  overviewBody: document.getElementById("overviewBody"),
  submitInfo: document.getElementById("submitInfo"),
  cardsContainer: document.getElementById("cardsContainer"),
};

function normalize(value) {
  if (value === undefined || value === null) {
    return "";
  }
  return String(value).trim();
}

function escapeHtml(value) {
  return normalize(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function parseSheetRows(workbook, sheetName, sourceType) {
  if (!workbook.SheetNames.includes(sheetName)) {
    return [];
  }
  const ws = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (!rows.length) {
    return [];
  }

  const headers = rows[0].map((h) => normalize(h).toLowerCase());
  const out = [];

  for (let i = 1; i < rows.length; i += 1) {
    const values = rows[i] || [];
    const row = {};
    let hasAny = false;
    headers.forEach((header, idx) => {
      if (!header) {
        return;
      }
      const v = values[idx];
      if (normalize(v)) {
        hasAny = true;
      }
      row[header] = v;
    });
    if (!hasAny) {
      continue;
    }
    row._source_type = sourceType;
    row._key = buildRowKey(row, i, sourceType);
    out.push(row);
  }
  return out;
}

function buildRowKey(row, idx, sourceType) {
  const id = normalize(row.id);
  const date = normalize(row.date);
  const title = normalize(row.titel || row.title);
  return `${sourceType}::${id}::${date}::${title}::${idx}`;
}

function parseRelevanzScore(value) {
  if (value === null || value === undefined || normalize(value) === "") {
    return -1;
  }
  const raw = normalize(value).replace(",", ".");
  const num = Number.parseFloat(raw);
  return Number.isFinite(num) ? num : -1;
}

function parseRowDate(value) {
  if (!value && value !== 0) {
    return null;
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }

  if (typeof value === "number") {
    const serial = value;
    const utcDays = Math.floor(serial - 25569);
    const utcValue = utcDays * 86400;
    const dateInfo = new Date(utcValue * 1000);
    return Number.isNaN(dateInfo.getTime()) ? null : dateInfo;
  }

  const raw = normalize(value);
  if (!raw) {
    return null;
  }

  const formats = [
    /^([0-9]{4})-([0-9]{2})-([0-9]{2})$/,
    /^([0-9]{2})\.([0-9]{2})\.([0-9]{4})$/,
    /^([0-9]{2})\/([0-9]{2})\/([0-9]{4})$/,
  ];

  for (const pattern of formats) {
    const m = raw.match(pattern);
    if (m) {
      if (pattern === formats[0]) {
        const d = new Date(`${m[1]}-${m[2]}-${m[3]}T00:00:00`);
        return Number.isNaN(d.getTime()) ? null : d;
      }
      const d = new Date(`${m[3]}-${m[2]}-${m[1]}T00:00:00`);
      return Number.isNaN(d.getTime()) ? null : d;
    }
  }

  const fallback = new Date(raw);
  return Number.isNaN(fallback.getTime()) ? null : fallback;
}

function parseDisplayDate(value) {
  const d = parseRowDate(value);
  if (!d) {
    return normalize(value) || "-";
  }
  const day = String(d.getDate()).padStart(2, "0");
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const year = d.getFullYear();
  return `${day}.${month}.${year}`;
}

function scoreBadgeClass(score) {
  if (score >= 8) {
    return "high";
  }
  if (score >= 5) {
    return "mid";
  }
  return "low";
}

function sourceLabel(sourceType) {
  return sourceType === "results" ? "Results" : "New Competition";
}

function normalizeSourceType(value) {
  return normalize(value).toLowerCase() === "results" ? "results" : "new_competition";
}

function extractFirstNumber(value) {
  const raw = normalize(value).toLowerCase().replace(/\s+/g, "");
  if (!raw) {
    return null;
  }

  const match = raw.match(/[0-9.,-]+/);
  if (!match) {
    return null;
  }

  let token = match[0];
  if (token.includes(",") && token.includes(".")) {
    token = token.replace(/\./g, "").replace(",", ".");
  } else if (token.includes(",")) {
    token = token.replace(",", ".");
  } else if ((token.match(/\./g) || []).length > 1) {
    token = token.replace(/\./g, "");
  }

  const num = Number.parseFloat(token);
  return Number.isFinite(num) ? num : null;
}

function formatMioEur(value) {
  const raw = normalize(value);
  if (!raw) {
    return "-";
  }

  const n = extractFirstNumber(raw);
  if (n === null) {
    return raw;
  }

  const lower = raw.toLowerCase();
  const isMio = lower.includes("mio") || lower.includes("million");
  const mio = isMio ? n : (n / 1000000);
  return `${mio.toFixed(2)} Mio €`;
}

function buildNoticeLinks(project) {
  const noticeId = normalize(project.id);
  const detailLink = normalize(project.link);
  if (detailLink) {
    const pdfLink = noticeId ? `https://ted.europa.eu/de/notice/${noticeId}/pdf` : detailLink;
    return [detailLink, pdfLink];
  }
  if (noticeId) {
    return [
      `https://ted.europa.eu/en/notice/-/detail/${noticeId}`,
      `https://ted.europa.eu/de/notice/${noticeId}/pdf`,
    ];
  }
  return ["#", "#"];
}

function buildGoogleMapsLink(location) {
  const query = normalize(location);
  if (!query || query === "-") {
    return "";
  }
  return `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(query)}`;
}

function buildLocationTags(location) {
  const raw = normalize(location).toLowerCase();
  if (!raw || raw === "-") {
    return new Set(["global_rest"]);
  }
  const tags = new Set();

  for (const [, value] of LOCATION_FILTERS) {
    if (value === "all" || value === "global_rest") {
      continue;
    }
    const keywords = REGION_KEYWORDS[value] || LOCATION_KEYWORDS[value] || [value];
    if (keywords.some((kw) => raw.includes(kw))) {
      tags.add(value);
    }
  }

  if (!tags.size) {
    tags.add("global_rest");
  }
  return tags;
}

function renderNamedRows(fields) {
  const rows = fields.map(([label, value]) => {
    const safeLabel = escapeHtml(label);
    const safeValue = escapeHtml(value || "-").replace(/\n/g, "<br>");
    return `<tr><th>${safeLabel}</th><td>${safeValue}</td></tr>`;
  });
  return `<table class="field-table">${rows.join("")}</table>`;
}

function enrichForSearch(row) {
  const titel = normalize(row.titel || row.title);
  const lage = normalize(row.projektlage);
  const category = normalize(row.category);
  const leistungen = normalize(row.leistungen);
  const wettbewerb = normalize(row.wettbewerb_art);
  const winner = normalize(row.gewinner);
  const winnerRole = normalize(row.gewinner_rolle);
  row._search = `${titel} ${lage} ${category} ${leistungen} ${wettbewerb} ${winner} ${winnerRole}`.toLowerCase();
  row._score = parseRelevanzScore(row.relevanzbewertung);
  row._locationTags = buildLocationTags(row.projektlage);
  row._categoryLower = category.toLowerCase();
}

function setupFilterOptions() {
  const categories = new Set(["all"]);
  state.allRows.forEach((row) => {
    const cat = normalize(row.category);
    if (cat) {
      categories.add(cat.toLowerCase());
    }
  });

  dom.locationFilter.innerHTML = LOCATION_FILTERS.map(([label, value]) =>
    `<option value="${escapeHtml(value)}">${escapeHtml(label)}</option>`
  ).join("");

  dom.categoryFilter.innerHTML = Array.from(categories)
    .map((cat) => {
      if (cat === "all") {
        return '<option value="all">All</option>';
      }
      return `<option value="${escapeHtml(cat)}">${escapeHtml(cat)}</option>`;
    })
    .join("");

  const scoreOptions = ["all", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"];
  dom.scoreFilter.innerHTML = scoreOptions
    .map((v) => {
      if (v === "all") {
        return '<option value="all">All</option>';
      }
      return `<option value="${v}">${v}</option>`;
    })
    .join("");
}

function rowMatchesFilters(row) {
  const { type, location, category, score, query } = state.filters;
  if (type !== "all" && normalizeSourceType(row._source_type) !== type) {
    return false;
  }
  if (location !== "all" && !row._locationTags.has(location)) {
    return false;
  }
  if (category !== "all" && row._categoryLower !== category) {
    return false;
  }
  if (score !== "all") {
    const rounded = row._score < 0 ? "" : String(Math.min(10, Math.max(1, Math.round(row._score))));
    if (rounded !== score) {
      return false;
    }
  }
  if (query && !row._search.includes(query)) {
    return false;
  }
  return true;
}

function applyFilters() {
  state.filteredRows = state.allRows.filter((row) => rowMatchesFilters(row));
  renderTable();
}

function renderTable() {
  if (!state.filteredRows.length) {
    dom.overviewBody.innerHTML = '<tr><td colspan="8" class="empty">No rows found for current filters.</td></tr>';
    dom.tableMeta.textContent = `Visible rows: 0 / ${state.allRows.length} | Selected rows: ${state.selectedKeys.size}`;
    return;
  }

  const htmlRows = state.filteredRows.map((row) => {
    const key = row._key;
    const selected = state.selectedKeys.has(key) ? "checked" : "";
    const type = sourceLabel(normalizeSourceType(row._source_type));
    const score = normalize(row.relevanzbewertung) || "-";
    const date = parseDisplayDate(row.date);
    const title = escapeHtml(normalize(row.titel || row.title) || "-");
    const location = escapeHtml(normalize(row.projektlage) || "-");
    const category = escapeHtml(normalize(row.category) || "-");
    const id = escapeHtml(normalize(row.id) || "-");

    return `
      <tr>
        <td><input type="checkbox" data-row-key="${escapeHtml(key)}" ${selected}></td>
        <td>${escapeHtml(type)}</td>
        <td>${escapeHtml(score)}</td>
        <td>${escapeHtml(date)}</td>
        <td>${title}</td>
        <td>${location}</td>
        <td>${category}</td>
        <td>${id}</td>
      </tr>
    `;
  });

  dom.overviewBody.innerHTML = htmlRows.join("");
  dom.tableMeta.textContent = `Visible rows: ${state.filteredRows.length} / ${state.allRows.length} | Selected rows: ${state.selectedKeys.size}`;

  dom.overviewBody.querySelectorAll('input[type="checkbox"][data-row-key]').forEach((input) => {
    input.addEventListener("change", () => {
      const key = input.getAttribute("data-row-key");
      if (!key) {
        return;
      }
      if (input.checked) {
        state.selectedKeys.add(key);
      } else {
        state.selectedKeys.delete(key);
      }
      dom.tableMeta.textContent = `Visible rows: ${state.filteredRows.length} / ${state.allRows.length} | Selected rows: ${state.selectedKeys.size}`;
    });
  });
}

function renderCards(rows) {
  if (!rows.length) {
    dom.cardsContainer.innerHTML = "";
    dom.submitInfo.textContent = "No rows selected. Please select at least one row and click Submit.";
    return;
  }

  dom.submitInfo.textContent = `Rendered ${rows.length} selected project(s).`;

  const cards = rows.map((project) => {
    const scoreRaw = normalize(project.relevanzbewertung) || "-";
    const scoreValue = parseRelevanzScore(project.relevanzbewertung);
    const scoreClass = scoreBadgeClass(Math.max(scoreValue, 0));

    const nummer = escapeHtml(normalize(project.id) || "-");
    const datum = escapeHtml(parseDisplayDate(project.date));
    const abgabefrist = escapeHtml(parseDisplayDate(project.abgabefrist));
    const titel = escapeHtml(normalize(project.titel || project.title) || "-");
    const kurzbeschreibung = escapeHtml(normalize(project.kurzbeschreibung) || "-").replace(/\n/g, "<br>");
    const lageText = normalize(project.projektlage) || "-";
    const mapsLink = buildGoogleMapsLink(lageText);
    const lage = mapsLink
      ? `<a href="${escapeHtml(mapsLink)}" target="_blank" rel="noopener noreferrer">${escapeHtml(lageText)}</a>`
      : escapeHtml(lageText);

    const category = escapeHtml(normalize(project.category) || "-");
    const leistungen = escapeHtml(normalize(project.leistungen) || "-").replace(/\n/g, "<br>");
    const wettbewerbsart = escapeHtml(normalize(project.wettbewerb_art) || "-").replace(/\n/g, "<br>");
    const gewinner = escapeHtml(normalize(project.gewinner) || "-").replace(/\n/g, "<br>");
    const gewinnerRolle = escapeHtml(normalize(project.gewinner_rolle) || "-").replace(/\n/g, "<br>");
    const erklaerung = escapeHtml(
      normalize(project.relevanzbewertung_erklaerung) || normalize(project.relevanzbewertung_begruendung) || "-"
    ).replace(/\n/g, "<br>");

    const [detailLink, pdfLink] = buildNoticeLinks(project);

    const sourceType = normalizeSourceType(project._source_type);
    const sourceTypeLabel = sourceLabel(sourceType);

    let mainLabel = "Leistungen";
    let mainValue = leistungen;
    let resultsMainFields = "";

    if (sourceType === "results") {
      mainLabel = "Wettbewerbsart";
      mainValue = wettbewerbsart;
      resultsMainFields = `
        <p><strong>Gewinner:</strong><br>${gewinner}</p>
        <p><strong>Gewinner Rolle:</strong><br>${gewinnerRolle}</p>
      `;
    }

    const kostenTable = sourceType === "results"
      ? renderNamedRows([
        ["Baukosten kg300/400", formatMioEur(project.baukosten_kg300_400)],
        ["Erklaerung der Baukosten", normalize(project.baukosten_erklaerung) || "-"],
      ])
      : renderNamedRows([
        ["Baukosten kg300/400", formatMioEur(project.baukosten_kg300_400)],
        ["Erklaerung der Baukosten", normalize(project.baukosten_erklaerung) || "-"],
        ["Honorar sbp", formatMioEur(project.geschaetztes_honorar_sbp)],
        ["Erklaerung Honorar SBP", normalize(project.honorar_erklaerung) || "-"],
      ]);

    const weitereTable = sourceType === "results"
      ? renderNamedRows([
        ["Wettbewerbsart", normalize(project.wettbewerb_art) || "-"],
        ["Gewinner", normalize(project.gewinner) || "-"],
        ["Gewinner Rolle", normalize(project.gewinner_rolle) || "-"],
        ["Gewinner Kontakt", normalize(project.gewinner_kontakt) || "-"],
        ["Projektbeteiligte", normalize(project.projektbeteiligte) || "-"],
        ["Naechste Schritte", normalize(project.naechste_schritte) || "-"],
        ["Notes", normalize(project.notes) || "-"],
      ])
      : renderNamedRows([
        ["Abgabefrist", parseDisplayDate(project.abgabefrist)],
        ["Leistungen", normalize(project.leistungen) || "-"],
        ["Umfang", normalize(project.umfang) || "-"],
        ["Zuschlagskriterien", normalize(project.zuschlagskriterien) || "-"],
        ["Referenzen/Qualifikationen", normalize(project.referenzen_qualifikationen) || "-"],
        ["Auftraggeber", normalize(project.auftraggeber) || "-"],
        ["Notes", normalize(project.notes) || "-"],
      ]);

    return `
      <article class="project ${scoreClass}">
        <div class="source-badge">${escapeHtml(sourceTypeLabel)}</div>
        <header class="project-head">
          <div class="score-pill ${scoreClass}">${escapeHtml(scoreRaw)}</div>
          <div class="head-main">
            <h3>${titel}</h3>
            <div class="head-grid">
              <p><strong>Number:</strong> ${nummer}</p>
              <p><strong>Datum der Veröffentlichung:</strong> ${datum}</p>
              <p><strong>Abgabefrist:</strong> ${abgabefrist}</p>
              <p><strong>Lage:</strong> ${lage}</p>
              <p><strong>Category:</strong> ${category}</p>
              <p><strong>Links:</strong> <a href="${escapeHtml(detailLink)}" target="_blank" rel="noopener noreferrer">Notice</a> | <a href="${escapeHtml(pdfLink)}" target="_blank" rel="noopener noreferrer">PDF</a></p>
            </div>
          </div>
        </header>

        <section class="always-visible">
          <p><strong>Kurzbeschreibung:</strong><br>${kurzbeschreibung}</p>
          <p><strong>${escapeHtml(mainLabel)}:</strong><br>${mainValue}</p>
          ${resultsMainFields}
          <p><strong>Relevanzbewertung Erklaerung:</strong><br>${erklaerung}</p>
        </section>

        <details class="details-block">
          <summary>Kostenschaetzung</summary>
          <div class="details-content">${kostenTable}</div>
        </details>

        <details class="details-block">
          <summary>Weitere Informationen</summary>
          <div class="details-content">${weitereTable}</div>
        </details>
      </article>
    `;
  });

  dom.cardsContainer.innerHTML = cards.join("\n");
}

function bindEvents() {
  dom.searchInput.addEventListener("input", () => {
    state.filters.query = normalize(dom.searchInput.value).toLowerCase();
    applyFilters();
  });

  dom.typeFilter.addEventListener("change", () => {
    state.filters.type = dom.typeFilter.value;
    applyFilters();
  });

  dom.locationFilter.addEventListener("change", () => {
    state.filters.location = dom.locationFilter.value;
    applyFilters();
  });

  dom.categoryFilter.addEventListener("change", () => {
    state.filters.category = dom.categoryFilter.value;
    applyFilters();
  });

  dom.scoreFilter.addEventListener("change", () => {
    state.filters.score = dom.scoreFilter.value;
    applyFilters();
  });

  dom.clearFiltersBtn.addEventListener("click", () => {
    state.filters = { type: "all", location: "all", category: "all", score: "all", query: "" };
    dom.searchInput.value = "";
    dom.typeFilter.value = "all";
    dom.locationFilter.value = "all";
    dom.categoryFilter.value = "all";
    dom.scoreFilter.value = "all";
    applyFilters();
  });

  dom.clearSelectionBtn.addEventListener("click", () => {
    state.selectedKeys.clear();
    renderTable();
    dom.submitInfo.textContent = "Selection cleared.";
    dom.cardsContainer.innerHTML = "";
  });

  dom.submitBtn.addEventListener("click", () => {
    const selectedRows = state.allRows.filter((row) => state.selectedKeys.has(row._key));
    renderCards(selectedRows);
  });
}

function renderWarnings(warnings) {
  if (!warnings.length) {
    dom.loadWarnings.textContent = "";
    return;
  }
  dom.loadWarnings.innerHTML = warnings.map((w) => `<div>${escapeHtml(w)}</div>`).join("");
}

async function loadWorkbook() {
  try {
    if (!window.XLSX) {
      throw new Error("SheetJS library not loaded.");
    }

    const response = await fetch(DATA_URL, { cache: "no-cache" });
    if (!response.ok) {
      throw new Error(`Could not fetch ${DATA_URL}. HTTP ${response.status}`);
    }

    const data = await response.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array", cellDates: true });

    const warnings = [];
    const newRows = parseSheetRows(workbook, NEW_SHEET, "new_competition");
    const resultRows = parseSheetRows(workbook, RESULTS_SHEET, "results");

    if (!workbook.SheetNames.includes(NEW_SHEET)) {
      warnings.push(`Worksheet '${NEW_SHEET}' was not found.`);
    }
    if (!workbook.SheetNames.includes(RESULTS_SHEET)) {
      warnings.push(`Worksheet '${RESULTS_SHEET}' was not found.`);
    }

    const allRows = [...newRows, ...resultRows];
    allRows.forEach((row) => enrichForSearch(row));

    state.allRows = allRows;
    setupFilterOptions();
    applyFilters();

    dom.loadStatus.textContent = `Workbook loaded: ${state.allRows.length} rows.`;
    renderWarnings(warnings);

    if (!state.allRows.length) {
      dom.overviewBody.innerHTML = '<tr><td colspan="8" class="empty">No rows found in workbook sheets.</td></tr>';
    }
  } catch (err) {
    dom.loadStatus.textContent = "Workbook load failed.";
    dom.loadWarnings.textContent = String(err && err.message ? err.message : err);
    dom.overviewBody.innerHTML = '<tr><td colspan="8" class="empty">Data could not be loaded. Check README and Excel path.</td></tr>';
  }
}

bindEvents();
loadWorkbook();
