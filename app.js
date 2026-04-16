const DATA_URL = "./data/ted_results.xlsx";
const NEW_SHEET = "Agent_2";
const RESULTS_SHEET = "Agent_2_Results";

const LOCATION_FILTERS = [
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
  rows: [],
  filters: {
    type: "all",
    location: "all",
    category: "all",
    scores: new Set(),
    query: "",
    startDate: "",
    endDate: "",
  },
  hasSubmitted: false,
};

function normalize(v) {
  if (v === undefined || v === null) {
    return "";
  }
  return String(v).trim();
}

function esc(v) {
  return normalize(v)
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
    const vals = rows[i] || [];
    const row = {};
    let hasAny = false;
    headers.forEach((h, idx) => {
      if (!h) {
        return;
      }
      const val = vals[idx];
      if (normalize(val)) {
        hasAny = true;
      }
      row[h] = val;
    });
    if (!hasAny) {
      continue;
    }
    row._source_type = sourceType;
    row._key = `${sourceType}::${normalize(row.id)}::${normalize(row.date)}::${normalize(row.titel || row.title)}::${i}`;
    out.push(row);
  }
  return out;
}

function parseRelevanzScore(value) {
  if (value === null || value === undefined) {
    return -1;
  }
  const raw = normalize(value).replace(",", ".");
  if (!raw) {
    return -1;
  }
  const parsed = Number.parseFloat(raw);
  return Number.isFinite(parsed) ? parsed : -1;
}

function parseRowDate(value) {
  if (value === null || value === undefined || normalize(value) === "") {
    return null;
  }
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }
  if (typeof value === "number") {
    const utcDays = Math.floor(value - 25569);
    const utcValue = utcDays * 86400;
    const d = new Date(utcValue * 1000);
    return Number.isNaN(d.getTime()) ? null : d;
  }

  const raw = normalize(value);
  const patterns = [
    /^([0-9]{4})-([0-9]{2})-([0-9]{2})$/,
    /^([0-9]{2})\.([0-9]{2})\.([0-9]{4})$/,
    /^([0-9]{2})\/([0-9]{2})\/([0-9]{4})$/,
  ];

  for (const p of patterns) {
    const m = raw.match(p);
    if (!m) {
      continue;
    }
    if (p === patterns[0]) {
      return new Date(`${m[1]}-${m[2]}-${m[3]}T00:00:00`);
    }
    return new Date(`${m[3]}-${m[2]}-${m[1]}T00:00:00`);
  }

  const fallback = new Date(raw);
  return Number.isNaN(fallback.getTime()) ? null : fallback;
}

function formatDateInput(dateObj) {
  const y = dateObj.getFullYear();
  const m = String(dateObj.getMonth() + 1).padStart(2, "0");
  const d = String(dateObj.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
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

function scoreFilterValue(score) {
  if (score < 0) {
    return "";
  }
  return String(Math.min(10, Math.max(1, Math.round(score))));
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

function normalizeSourceType(value) {
  return normalize(value).toLowerCase() === "results" ? "results" : "new_competition";
}

function sourceLabel(sourceType) {
  return sourceType === "results" ? "Results" : "New Competition";
}

function buildLocationTags(location) {
  const raw = normalize(location).toLowerCase();
  if (!raw || raw === "-") {
    return new Set(["global_rest"]);
  }

  const tags = new Set();
  LOCATION_FILTERS.forEach(([, value]) => {
    if (value === "global_rest") {
      return;
    }
    const keywords = REGION_KEYWORDS[value] || LOCATION_KEYWORDS[value] || [value];
    if (keywords.some((kw) => raw.includes(kw))) {
      tags.add(value);
    }
  });

  if (!tags.size) {
    tags.add("global_rest");
  }
  return tags;
}

function extractFirstNumber(value) {
  const raw = normalize(value).toLowerCase().replace(/\s+/g, "");
  if (!raw) {
    return null;
  }

  const m = raw.match(/[0-9.,-]+/);
  if (!m) {
    return null;
  }

  let token = m[0];
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

  const number = extractFirstNumber(raw);
  if (number === null) {
    return raw;
  }

  const lower = raw.toLowerCase();
  const isMio = lower.includes("mio") || lower.includes("million");
  const mioValue = isMio ? number : (number / 1000000);
  return `${mioValue.toFixed(2)} Mio €`;
}

function renderNamedRows(fields) {
  const rows = fields.map(([label, value]) => {
    const safeLabel = esc(label);
    const safeValue = esc(value || "-").replace(/\n/g, "<br>");
    return `<tr><th>${safeLabel}</th><td>${safeValue}</td></tr>`;
  });
  return `<table class="field-table">${rows.join("")}</table>`;
}

function enrichRow(row) {
  const title = normalize(row.titel || row.title);
  const lage = normalize(row.projektlage);
  const category = normalize(row.category);
  const leistungen = normalize(row.leistungen);
  const wettbewerb = normalize(row.wettbewerb_art);
  const winner = normalize(row.gewinner);
  const winnerRole = normalize(row.gewinner_rolle);

  row._score = parseRelevanzScore(row.relevanzbewertung);
  row._scoreFilter = scoreFilterValue(row._score);
  row._locationTags = buildLocationTags(lage);
  row._category = category.toLowerCase();
  row._source = normalizeSourceType(row._source_type);
  row._search = `${title} ${lage} ${category} ${leistungen} ${wettbewerb} ${winner} ${winnerRole} ${row._source}`.toLowerCase();
  row._dateObj = parseRowDate(row.date);
}

function addDynamicFilterButtons() {
  const locationWrap = document.getElementById("locationFilters");
  LOCATION_FILTERS.forEach(([label, value]) => {
    locationWrap.insertAdjacentHTML(
      "beforeend",
      `<button class="filter-btn" data-filter-group="location" data-value="${esc(value)}">${esc(label)}</button>`
    );
  });

  const categories = [...new Set(state.rows.map((r) => normalize(r.category)).filter(Boolean))].sort();
  const categoryWrap = document.getElementById("categoryFilters");
  categories.forEach((cat) => {
    categoryWrap.insertAdjacentHTML(
      "beforeend",
      `<button class="filter-btn" data-filter-group="category" data-value="${esc(cat.toLowerCase())}">${esc(cat)}</button>`
    );
  });

  const scoreWrap = document.getElementById("scoreFilters");
  for (let i = 1; i <= 10; i += 1) {
    scoreWrap.insertAdjacentHTML(
      "beforeend",
      `<button class="filter-btn" data-filter-group="score" data-value="${i}">${i}</button>`
    );
  }
}

function dateWithinRange(rowDate, startDateRaw, endDateRaw) {
  const hasStart = normalize(startDateRaw) !== "";
  const hasEnd = normalize(endDateRaw) !== "";
  if (!hasStart && !hasEnd) {
    return true;
  }
  if (!rowDate) {
    return false;
  }

  const startDate = hasStart ? new Date(`${startDateRaw}T00:00:00`) : null;
  const endDate = hasEnd ? new Date(`${endDateRaw}T23:59:59`) : null;

  if (startDate && rowDate < startDate) {
    return false;
  }
  if (endDate && rowDate > endDate) {
    return false;
  }
  return true;
}

function matchesFilters(row, override = null) {
  const base = {
    type: state.filters.type,
    location: state.filters.location,
    category: state.filters.category,
    scores: state.filters.scores,
    query: state.filters.query,
    startDate: state.filters.startDate,
    endDate: state.filters.endDate,
  };

  const f = { ...base, ...(override || {}) };

  const typeMatch = f.type === "all" || row._source === f.type;
  const locMatch = f.location === "all" || row._locationTags.has(f.location);
  const catMatch = f.category === "all" || row._category.includes(f.category);
  const scoreMatch = f.scores.size === 0 || f.scores.has(row._scoreFilter);
  const searchMatch = !f.query || row._search.includes(f.query);
  const dateMatch = dateWithinRange(row._dateObj, f.startDate, f.endDate);
  return typeMatch && locMatch && catMatch && scoreMatch && searchMatch && dateMatch;
}

function computeVisibleCount(override = null) {
  return state.rows.filter((row) => matchesFilters(row, override)).length;
}

function ensureButtonBadge(btn) {
  if (btn.querySelector(".filter-count")) {
    return;
  }
  const label = normalize(btn.textContent);
  btn.setAttribute("data-label", label);
  btn.textContent = "";
  const labelSpan = document.createElement("span");
  labelSpan.textContent = label;
  const countSpan = document.createElement("span");
  countSpan.className = "filter-count";
  countSpan.textContent = "0";
  btn.appendChild(labelSpan);
  btn.appendChild(countSpan);
}

function updateFilterCountBadges() {
  document.querySelectorAll(".filter-btn").forEach((btn) => {
    ensureButtonBadge(btn);
    const group = btn.getAttribute("data-filter-group");
    const value = btn.getAttribute("data-value") || "all";

    let count = 0;
    if (group === "type") {
      count = computeVisibleCount({ type: value });
    } else if (group === "location") {
      count = computeVisibleCount({ location: value });
    } else if (group === "category") {
      count = computeVisibleCount({ category: value });
    } else if (group === "score") {
      const scoreSet = value === "all" ? new Set() : new Set([value]);
      count = computeVisibleCount({ scores: scoreSet });
    }

    const badge = btn.querySelector(".filter-count");
    if (badge) {
      badge.textContent = String(count);
    }
  });
}

function renderCard(project) {
  const scoreRaw = normalize(project.relevanzbewertung) || "-";
  const scoreClass = scoreBadgeClass(Math.max(project._score, 0));

  const nummer = esc(normalize(project.id) || "-");
  const datum = esc(parseDisplayDate(project.date));
  const abgabefrist = esc(parseDisplayDate(project.abgabefrist));
  const titel = esc(normalize(project.titel || project.title) || "-");
  const kurzbeschreibung = esc(normalize(project.kurzbeschreibung) || "-").replace(/\n/g, "<br>");

  const lageText = normalize(project.projektlage) || "-";
  const mapsLink = buildGoogleMapsLink(lageText);
  const lage = mapsLink
    ? `<a href="${esc(mapsLink)}" target="_blank" rel="noopener noreferrer">${esc(lageText)}</a>`
    : esc(lageText);

  const categoryValue = esc(normalize(project.category) || "-");
  const leistungen = esc(normalize(project.leistungen) || "-").replace(/\n/g, "<br>");
  const wettbewerbsart = esc(normalize(project.wettbewerb_art) || "-").replace(/\n/g, "<br>");
  const gewinner = esc(normalize(project.gewinner) || "-").replace(/\n/g, "<br>");
  const gewinnerRolle = esc(normalize(project.gewinner_rolle) || "-").replace(/\n/g, "<br>");
  const erklaerung = esc(
    normalize(project.relevanzbewertung_erklaerung) || normalize(project.relevanzbewertung_begruendung) || "-"
  ).replace(/\n/g, "<br>");

  const [detailLink, pdfLink] = buildNoticeLinks(project);
  const sourceType = project._source;
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
    <article class="project ${scoreClass}" data-key="${esc(project._key)}">
      <div class="card-topline">
        <div class="source-badge">${esc(sourceTypeLabel)}</div>
      </div>
      <header class="project-head">
        <div class="score-pill ${scoreClass}">${esc(scoreRaw)}</div>
        <div class="head-main">
          <h2>${titel}</h2>
          <div class="head-grid">
            <p><strong>Number:</strong> ${nummer}</p>
            <p><strong>Datum der Veröffentlichung:</strong> ${datum}</p>
            <p><strong>Abgabefrist:</strong> ${abgabefrist}</p>
            <p><strong>Lage:</strong> ${lage}</p>
            <p><strong>Category:</strong> ${categoryValue}</p>
            <p><strong>Links:</strong> <a href="${esc(detailLink)}" target="_blank" rel="noopener noreferrer">Notice</a> | <a href="${esc(pdfLink)}" target="_blank" rel="noopener noreferrer">PDF</a></p>
          </div>
        </div>
      </header>

      <section class="always-visible">
        <p><strong>Kurzbeschreibung:</strong><br>${kurzbeschreibung}</p>
        <p><strong>${esc(mainLabel)}:</strong><br>${mainValue}</p>
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
}

function activeFilteredRows() {
  return state.rows.filter((row) => matchesFilters(row));
}

function updateCounts() {
  const countEl = document.getElementById("resultsCount");
  const count = activeFilteredRows().length;
  countEl.textContent = `Filter preview: ${count} / ${state.rows.length} cards`;

  const meta = document.getElementById("metaInfo");
  meta.textContent = `Treffer: ${state.rows.length} (all loaded rows)`;
}

function renderCardsArea() {
  const pool = document.getElementById("cardsPool");
  if (!state.hasSubmitted) {
    pool.innerHTML = "<p>Keine Ergebnisse angezeigt. Klicke auf <strong>Submit Selection</strong>, um die aktuell gefilterten Karten zu laden.</p>";
    return;
  }

  const filtered = activeFilteredRows();
  if (!filtered.length) {
    pool.innerHTML = "<p>Keine passenden Eintraege gefunden.</p>";
    return;
  }

  pool.innerHTML = filtered.map((row) => renderCard(row)).join("\n");
}

function refreshUI() {
  updateCounts();
  updateFilterCountBadges();
  renderCardsArea();
}

function activateSingleSelectFilter(group, value, clickedButton) {
  if (group === "type") {
    state.filters.type = value;
  }
  if (group === "location") {
    state.filters.location = value;
  }
  if (group === "category") {
    state.filters.category = value;
  }

  document.querySelectorAll(`.filter-btn[data-filter-group="${group}"]`).forEach((b) => b.classList.remove("active"));
  clickedButton.classList.add("active");
}

function bindUi() {
  document.getElementById("liveSearch").addEventListener("input", (e) => {
    state.filters.query = normalize(e.target.value).toLowerCase();
    refreshUI();
  });

  document.getElementById("startDate").addEventListener("change", (e) => {
    state.filters.startDate = normalize(e.target.value);
    refreshUI();
  });

  document.getElementById("endDate").addEventListener("change", (e) => {
    state.filters.endDate = normalize(e.target.value);
    refreshUI();
  });

  document.querySelectorAll(".filter-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const group = btn.getAttribute("data-filter-group");
      const value = btn.getAttribute("data-value") || "all";
      if (!group) {
        return;
      }

      if (group === "score") {
        if (value === "all") {
          state.filters.scores.clear();
          document.querySelectorAll('.filter-btn[data-filter-group="score"]').forEach((b) => b.classList.remove("active"));
          btn.classList.add("active");
        } else {
          const allBtn = document.querySelector('.filter-btn[data-filter-group="score"][data-value="all"]');
          if (state.filters.scores.has(value)) {
            state.filters.scores.delete(value);
            btn.classList.remove("active");
          } else {
            state.filters.scores.add(value);
            btn.classList.add("active");
          }
          if (allBtn) {
            if (state.filters.scores.size === 0) {
              allBtn.classList.add("active");
            } else {
              allBtn.classList.remove("active");
            }
          }
        }
      } else {
        activateSingleSelectFilter(group, value, btn);
      }

      refreshUI();
    });
  });

  document.getElementById("expandAll").addEventListener("click", () => {
    document.querySelectorAll("details").forEach((d) => { d.open = true; });
  });

  document.getElementById("collapseAll").addEventListener("click", () => {
    document.querySelectorAll("details").forEach((d) => { d.open = false; });
  });

  document.getElementById("clearFilters").addEventListener("click", () => {
    state.filters.type = "all";
    state.filters.location = "all";
    state.filters.category = "all";
    state.filters.scores.clear();
    state.filters.query = "";
    document.getElementById("liveSearch").value = "";

    const startEl = document.getElementById("startDate");
    const endEl = document.getElementById("endDate");
    state.filters.startDate = normalize(startEl.value);
    state.filters.endDate = normalize(endEl.value);

    document.querySelectorAll(".filter-btn").forEach((b) => b.classList.remove("active"));
    document.querySelectorAll('.filter-btn[data-value="all"]').forEach((b) => b.classList.add("active"));
    refreshUI();
  });

  document.getElementById("submitSelection").addEventListener("click", () => {
    state.hasSubmitted = true;
    refreshUI();
  });
}

function initializeDateInputs() {
  const datedRows = state.rows.filter((r) => r._dateObj instanceof Date);
  if (!datedRows.length) {
    return;
  }

  datedRows.sort((a, b) => a._dateObj - b._dateObj);
  const minDate = formatDateInput(datedRows[0]._dateObj);
  const maxDate = formatDateInput(datedRows[datedRows.length - 1]._dateObj);

  const startEl = document.getElementById("startDate");
  const endEl = document.getElementById("endDate");
  startEl.value = minDate;
  endEl.value = maxDate;

  state.filters.startDate = minDate;
  state.filters.endDate = maxDate;
}

async function loadWorkbook() {
  const status = document.getElementById("loadStatus");
  const warningsEl = document.getElementById("loadWarnings");

  try {
    if (!window.XLSX) {
      throw new Error("SheetJS library not loaded.");
    }

    const response = await fetch(DATA_URL, { cache: "no-cache" });
    if (!response.ok) {
      throw new Error(`Could not fetch ${DATA_URL}. HTTP ${response.status}`);
    }

    const buffer = await response.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: true });

    const warnings = [];
    const newRows = parseSheetRows(workbook, NEW_SHEET, "new_competition");
    const resultRows = parseSheetRows(workbook, RESULTS_SHEET, "results");

    if (!workbook.SheetNames.includes(NEW_SHEET)) {
      warnings.push(`Worksheet '${NEW_SHEET}' was not found.`);
    }
    if (!workbook.SheetNames.includes(RESULTS_SHEET)) {
      warnings.push(`Worksheet '${RESULTS_SHEET}' was not found.`);
    }

    state.rows = [...newRows, ...resultRows];
    state.rows.forEach((row) => enrichRow(row));
    state.rows.sort((a, b) => b._score - a._score);

    addDynamicFilterButtons();
    bindUi();
    initializeDateInputs();
    refreshUI();

    status.textContent = `Workbook loaded: ${state.rows.length} rows from '${NEW_SHEET}' + '${RESULTS_SHEET}'.`;
    warningsEl.innerHTML = warnings.map((w) => `<div>${esc(w)}</div>`).join("");
  } catch (err) {
    status.textContent = "Workbook load failed.";
    warningsEl.textContent = String(err && err.message ? err.message : err);
    document.getElementById("cardsPool").innerHTML = "<p>Data could not be loaded. Check Excel path.</p>";
  }
}

loadWorkbook();
