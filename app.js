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
    "norwegen", "norway", "schweden", "sweden", "finnland", "finland", "daenemark", "denmark", "Dänemark",
    "lettland", "latvia", "litauen", "lithuania", "estland", "estonia",
  ],
  region_eastern_balkans: [
    "polen", "poland", "ungarn", "hungary", "slowakei", "slovakia", "tschechien", "czech", "czech republic",
    "slowenien", "slovenia", "kroatien", "croatia", "serbien", "serbia", "rumaenien", "romania", "Rumänien",
    "bulgarien", "bulgaria", "moldau", "moldova", "griechenland", "thessaloniki",
  ],
  region_central_europe: [
    "deutschland", "germany", "oesterreich", "austria", "schweiz", "switzerland", "frankreich", "france",
    "belgien", "belgium", "niederlande", "netherlands", "luxemburg", "luxembourg",
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
    onlyVerified: false,
    onlyWithActivity: false,
    onlyWithOpenRequest: false,
    sortOrder: "score_desc",
  },
  auth: {
    enabled: false,
    client: null,
    user: null,
  },
  remote: {
    commentsByKey: new Map(),
    overridesByKey: new Map(),
    verificationsByKey: new Map(),
    approvalRequestsByKey: new Map(),
    fieldEditsByKey: new Map(),
    loadRequestId: 0,
  },
  ui: {
    authCollapsed: false,
    hasAutoCollapsedAuth: false,
    openCommentForms: new Set(),
    openOverrideForms: new Set(),
    openFieldEditForms: new Set(),
    pendingForms: new Set(),
  },
};

function normalize(v) {
  if (v === undefined || v === null) return "";
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

function authMessage(text, isError = false) {
  const el = document.getElementById("authMessage");
  if (!el) return;
  el.textContent = text;
  el.classList.toggle("error", isError);
}

function applyAuthPanelVisibility() {
  const panel = document.getElementById("authPanel");
  const toggleBtn = document.getElementById("authToggleBtn");
  const body = document.body;
  if (!panel || !toggleBtn || !body) return;
  panel.classList.toggle("collapsed", state.ui.authCollapsed);
  body.classList.toggle("auth-collapsed", state.ui.authCollapsed);
  toggleBtn.textContent = state.ui.authCollapsed ? "Einblenden" : "Ausblenden";
}

function applyDashboardGate() {
  document.body.classList.toggle("dashboard-locked", !state.auth.user);
}

function updateAuthStatus() {
  const status = document.getElementById("authStatus");
  const logoutBtn = document.getElementById("logoutBtn");
  if (!status || !logoutBtn) return;

  if (!state.auth.enabled) {
    status.textContent = "Supabase nicht konfiguriert - Login ist erforderlich.";
    logoutBtn.disabled = true;
    applyAuthPanelVisibility();
    applyDashboardGate();
    return;
  }

  if (state.auth.user) {
    status.textContent = `Angemeldet als ${state.auth.user.email}`;
    logoutBtn.disabled = false;
    if (!state.ui.hasAutoCollapsedAuth) {
      state.ui.authCollapsed = true;
      state.ui.hasAutoCollapsedAuth = true;
    }
  } else {
    status.textContent = "Nicht angemeldet";
    logoutBtn.disabled = true;
    state.ui.authCollapsed = false;
    state.ui.hasAutoCollapsedAuth = false;
  }
  applyAuthPanelVisibility();
  applyDashboardGate();
}

function toAsciiKey(raw) {
  return raw
    .toLowerCase()
    .replace(/[^a-z0-9_-]+/g, "-")
    .replace(/^-+/, "")
    .replace(/-+$/, "");
}

function buildTenderKey(row, sourceType, index) {
  const noticeId = normalize(row.id);
  if (noticeId) return `ted:${noticeId}`;

  const date = normalize(row.date);
  const title = normalize(row.titel || row.title);
  const fallback = toAsciiKey(`${sourceType}-${date}-${title}`);
  return fallback ? `fallback:${fallback}` : `fallback:${sourceType}-${index}`;
}

function parseSheetRows(workbook, sheetName, sourceType) {
  if (!workbook.SheetNames.includes(sheetName)) return [];
  const ws = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (!rows.length) return [];

  const headers = rows[0].map((h) => normalize(h).toLowerCase());
  const out = [];
  for (let i = 1; i < rows.length; i += 1) {
    const vals = rows[i] || [];
    const row = {};
    let hasAny = false;
    headers.forEach((h, idx) => {
      if (!h) return;
      const val = vals[idx];
      if (normalize(val)) hasAny = true;
      row[h] = val;
    });
    if (!hasAny) continue;
    row._source_type = sourceType;
    row._tenderKey = buildTenderKey(row, sourceType, i);
    row._key = `${row._tenderKey}::${i}`;
    out.push(row);
  }
  return out;
}

function parseRelevanzScore(value) {
  if (value === null || value === undefined) return -1;
  const raw = normalize(value).replace(",", ".");
  if (!raw) return -1;
  const parsed = Number.parseFloat(raw);
  return Number.isFinite(parsed) ? parsed : -1;
}

function parseRowDate(value) {
  if (value === null || value === undefined || normalize(value) === "") return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
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
    if (!m) continue;
    if (p === patterns[0]) return new Date(`${m[1]}-${m[2]}-${m[3]}T00:00:00`);
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

function formatDateTime(value) {
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return "-";
  return d.toLocaleString("de-DE", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function clearDateRangeQuickSelection() {
  document.querySelectorAll(".date-range-btn").forEach((btn) => btn.classList.remove("active"));
}

function applyQuickDateRange(days) {
  if (!Number.isFinite(days) || days < 1) return;

  const startDate = new Date();
  startDate.setHours(0, 0, 0, 0);
  startDate.setDate(startDate.getDate() - (days - 1));

  const endDate = new Date();
  endDate.setHours(0, 0, 0, 0);

  const startValue = formatDateInput(startDate);
  const endValue = formatDateInput(endDate);

  const startEl = document.getElementById("startDate");
  const endEl = document.getElementById("endDate");
  startEl.value = startValue;
  endEl.value = endValue;

  state.filters.startDate = startValue;
  state.filters.endDate = endValue;
}

function parseDisplayDate(value) {
  const d = parseRowDate(value);
  if (!d) return normalize(value) || "-";
  const day = String(d.getDate()).padStart(2, "0");
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const year = d.getFullYear();
  return `${day}.${month}.${year}`;
}

function scoreBadgeClass(score) {
  if (score >= 8) return "high";
  if (score >= 5) return "mid";
  return "low";
}

function scoreFilterValue(score) {
  if (score < 0) return "";
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
  if (!query || query === "-") return "";
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
  if (!raw || raw === "-") return new Set(["global_rest"]);

  const tags = new Set();
  LOCATION_FILTERS.forEach(([, value]) => {
    if (value === "global_rest") return;
    const keywords = REGION_KEYWORDS[value] || LOCATION_KEYWORDS[value] || [value];
    if (keywords.some((kw) => raw.includes(kw))) tags.add(value);
  });

  if (!tags.size) tags.add("global_rest");
  return tags;
}

function extractFirstNumber(value) {
  const raw = normalize(value).toLowerCase().replace(/\s+/g, "");
  if (!raw) return null;
  const m = raw.match(/[0-9.,-]+/);
  if (!m) return null;

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
  if (!raw) return "-";

  const number = extractFirstNumber(raw);
  if (number === null) return raw;

  const lower = raw.toLowerCase();
  const isMio = lower.includes("mio") || lower.includes("million");
  const mioValue = isMio ? number : number / 1000000;
  return `${mioValue.toFixed(2)} Mio EUR`;
}

function parseCostSortable(value) {
  const raw = normalize(value);
  if (!raw) return Number.NEGATIVE_INFINITY;
  const number = extractFirstNumber(raw);
  if (number === null) return Number.NEGATIVE_INFINITY;
  const lower = raw.toLowerCase();
  const isMio = lower.includes("mio") || lower.includes("million");
  return isMio ? number : number / 1000000;
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

  row._baseScore = parseRelevanzScore(row.relevanzbewertung);
  row._effectiveScore = row._baseScore;
  row._effectiveScoreRaw = normalize(row.relevanzbewertung) || "-";
  row._scoreFilter = scoreFilterValue(row._effectiveScore);
  row._locationTags = buildLocationTags(lage);
  row._category = category.toLowerCase();
  row._source = normalizeSourceType(row._source_type);
  row._search = `${title} ${lage} ${category} ${leistungen} ${wettbewerb} ${winner} ${winnerRole} ${row._source}`.toLowerCase();
  row._dateObj = parseRowDate(row.date);
  row._deadlineObj = parseRowDate(row.abgabefrist);
  row._costValue = parseCostSortable(row.baukosten_kg300_400);
}

function getOverrideHistory(tenderKey) {
  return state.remote.overridesByKey.get(tenderKey) || [];
}

function getLatestOverride(tenderKey) {
  const history = getOverrideHistory(tenderKey);
  if (!history.length) return null;
  return history[history.length - 1];
}

function getVerificationHistory(tenderKey) {
  return state.remote.verificationsByKey.get(tenderKey) || [];
}

function getLatestVerification(tenderKey) {
  const history = getVerificationHistory(tenderKey);
  if (!history.length) return null;
  return history[history.length - 1];
}

function isVerified(tenderKey) {
  const latest = getLatestVerification(tenderKey);
  return !!(latest && latest.is_verified);
}

function getApprovalRequestHistory(tenderKey) {
  return state.remote.approvalRequestsByKey.get(tenderKey) || [];
}

function getLatestApprovalRequest(tenderKey) {
  const history = getApprovalRequestHistory(tenderKey);
  if (!history.length) return null;
  return history[history.length - 1];
}

function hasOpenApprovalRequest(tenderKey) {
  const latest = getLatestApprovalRequest(tenderKey);
  return !!(latest && normalize(latest.status).toLowerCase() === "open");
}

function getFieldEditHistory(tenderKey) {
  return state.remote.fieldEditsByKey.get(tenderKey) || [];
}

function getLatestFieldEdit(tenderKey) {
  const history = getFieldEditHistory(tenderKey);
  if (!history.length) return null;
  return history[history.length - 1];
}

function hasFieldEdits(tenderKey) {
  return getFieldEditHistory(tenderKey).length > 0;
}

function getEffectiveProjectData(project) {
  const latestEdit = getLatestFieldEdit(project._tenderKey);
  const payload = latestEdit && typeof latestEdit.edit_payload === "object" && latestEdit.edit_payload
    ? latestEdit.edit_payload
    : null;
  return payload ? { ...project, ...payload } : project;
}

function rebuildDerivedFields(row, sourceData) {
  const title = normalize(sourceData.titel || sourceData.title);
  const lage = normalize(sourceData.projektlage);
  const category = normalize(sourceData.category);
  const leistungen = normalize(sourceData.leistungen);
  const wettbewerb = normalize(sourceData.wettbewerb_art);
  const winner = normalize(sourceData.gewinner);
  const winnerRole = normalize(sourceData.gewinner_rolle);

  row._locationTags = buildLocationTags(lage);
  row._category = category.toLowerCase();
  row._search = `${title} ${lage} ${category} ${leistungen} ${wettbewerb} ${winner} ${winnerRole} ${row._source}`.toLowerCase();
  row._dateObj = parseRowDate(sourceData.date);
  row._deadlineObj = parseRowDate(sourceData.abgabefrist);
  row._costValue = parseCostSortable(sourceData.baukosten_kg300_400);
}

function applyEffectiveScores() {
  state.rows.forEach((row) => {
    row._effectiveData = getEffectiveProjectData(row);
    rebuildDerivedFields(row, row._effectiveData);

    row._baseScore = parseRelevanzScore(row._effectiveData.relevanzbewertung);
    const latest = getLatestOverride(row._tenderKey);
    if (latest) {
      row._effectiveScore = parseRelevanzScore(latest.score_value);
      row._effectiveScoreRaw = normalize(latest.score_value) || "-";
    } else {
      row._effectiveScore = row._baseScore;
      row._effectiveScoreRaw = normalize(row._effectiveData.relevanzbewertung) || "-";
    }
    row._scoreFilter = scoreFilterValue(row._effectiveScore);
  });
}

function dateValueOrFallback(dateObj, fallback) {
  if (!(dateObj instanceof Date) || Number.isNaN(dateObj.getTime())) return fallback;
  return dateObj.getTime();
}

function sortRows(rows) {
  const order = state.filters.sortOrder;
  const sorted = rows.slice();

  sorted.sort((a, b) => {
    if (order === "date_desc") return dateValueOrFallback(b._dateObj, -1) - dateValueOrFallback(a._dateObj, -1);
    if (order === "date_asc") return dateValueOrFallback(a._dateObj, Number.POSITIVE_INFINITY) - dateValueOrFallback(b._dateObj, Number.POSITIVE_INFINITY);
    if (order === "deadline_desc") return dateValueOrFallback(b._deadlineObj, -1) - dateValueOrFallback(a._deadlineObj, -1);
    if (order === "deadline_asc") return dateValueOrFallback(a._deadlineObj, Number.POSITIVE_INFINITY) - dateValueOrFallback(b._deadlineObj, Number.POSITIVE_INFINITY);
    if (order === "cost_desc") return b._costValue - a._costValue;
    if (order === "cost_asc") return a._costValue - b._costValue;
    return b._effectiveScore - a._effectiveScore;
  });

  return sorted;
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
  if (!hasStart && !hasEnd) return true;
  if (!rowDate) return false;

  const startDate = hasStart ? new Date(`${startDateRaw}T00:00:00`) : null;
  const endDate = hasEnd ? new Date(`${endDateRaw}T23:59:59`) : null;
  if (startDate && rowDate < startDate) return false;
  if (endDate && rowDate > endDate) return false;
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
    onlyVerified: state.filters.onlyVerified,
    onlyWithActivity: state.filters.onlyWithActivity,
    onlyWithOpenRequest: state.filters.onlyWithOpenRequest,
  };
  const f = { ...base, ...(override || {}) };

  const typeMatch = f.type === "all" || row._source === f.type;
  const locMatch = f.location === "all" || row._locationTags.has(f.location);
  const catMatch = f.category === "all" || row._category.includes(f.category);
  const scoreMatch = f.scores.size === 0 || f.scores.has(row._scoreFilter);
  const searchMatch = !f.query || row._search.includes(f.query);
  const dateMatch = dateWithinRange(row._dateObj, f.startDate, f.endDate);
  const verifiedMatch = !f.onlyVerified || isVerified(row._tenderKey);
  const activityMatch = !f.onlyWithActivity || (
    (state.remote.commentsByKey.get(row._tenderKey) || []).length > 0
    || (state.remote.overridesByKey.get(row._tenderKey) || []).length > 0
    || hasFieldEdits(row._tenderKey)
  );
  const openRequestMatch = !f.onlyWithOpenRequest || hasOpenApprovalRequest(row._tenderKey);
  return typeMatch && locMatch && catMatch && scoreMatch && searchMatch && dateMatch && verifiedMatch && activityMatch && openRequestMatch;
}

function computeVisibleCount(override = null) {
  return state.rows.filter((row) => matchesFilters(row, override)).length;
}

function ensureButtonBadge(btn) {
  if (btn.querySelector(".filter-count")) return;
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
    const group = btn.getAttribute("data-filter-group");
    if (!group) return;
    ensureButtonBadge(btn);
    const value = btn.getAttribute("data-value") || "all";

    let count = 0;
    if (group === "type") count = computeVisibleCount({ type: value });
    else if (group === "location") count = computeVisibleCount({ location: value });
    else if (group === "category") count = computeVisibleCount({ category: value });
    else if (group === "score") {
      const scoreSet = value === "all" ? new Set() : new Set([value]);
      count = computeVisibleCount({ scores: scoreSet });
    }

    const badge = btn.querySelector(".filter-count");
    if (badge) badge.textContent = String(count);
  });
}

function renderComments(project) {
  if (!state.auth.user) return "";

  const comments = state.remote.commentsByKey.get(project._tenderKey) || [];
  const myUserId = state.auth.user ? state.auth.user.id : "";
  const isFormOpen = state.ui.openCommentForms.has(project._tenderKey);
  const hasComments = comments.length > 0;

  const listHtml = comments.length
    ? comments.map((comment) => {
      const canDelete = comment.user_id === myUserId;
      return `
        <li class="comment-item">
          <div class="comment-head">
            <span>${esc(comment.user_email || "Unbekannt")}</span>
            <span>${esc(formatDateTime(comment.created_at))}</span>
          </div>
          <p>${esc(comment.comment_text).replace(/\n/g, "<br>")}</p>
          ${canDelete ? `<button class="action-btn comment-delete" type="button" data-comment-id="${esc(comment.id)}" data-tender-key="${esc(project._tenderKey)}">Kommentar loeschen</button>` : ""}
        </li>
      `;
    }).join("\n")
    : "";

  const formHtml = isFormOpen
    ? `
      <form class="comment-form" data-tender-key="${esc(project._tenderKey)}">
        <label>Kommentar</label>
        <textarea name="comment_text" required maxlength="2000"></textarea>
        <button class="action-btn" type="submit">Kommentar speichern</button>
      </form>
    `
    : "";

  if (!hasComments && !isFormOpen) {
    return "";
  }

  return `
    <section class="card-section">
      <h3>Kommentare</h3>
      <ul class="comment-list">${listHtml}</ul>
      ${formHtml}
    </section>
  `;
}

function renderOverrides(project) {
  if (!state.auth.user) return "";

  const history = getOverrideHistory(project._tenderKey);
  const latest = getLatestOverride(project._tenderKey);
  const isFormOpen = state.ui.openOverrideForms.has(project._tenderKey);
  const hasOverrides = history.length > 0;
  const myUserId = state.auth.user ? state.auth.user.id : "";

  const historyHtml = history.length
    ? history.slice().reverse().map((entry, idx) => {
      const isLatest = idx === 0;
      const reason = normalize(entry.reason_text) || "-";
      return `
        <li class="override-item ${isLatest ? "latest" : ""}">
          <div class="override-head">
            <span>${esc(entry.user_email || "Unbekannt")}</span>
            <span>${esc(formatDateTime(entry.created_at))}</span>
          </div>
          <p><strong>Score:</strong> ${esc(entry.score_value)}</p>
          <p><strong>Begruendung:</strong> ${esc(reason).replace(/\n/g, "<br>")}</p>
          ${isLatest ? '<span class="latest-badge">Aktiver Override</span>' : ""}
          ${entry.user_id === myUserId ? `<button class="action-btn override-delete" type="button" data-override-id="${esc(entry.id)}" data-tender-key="${esc(project._tenderKey)}">Override loeschen</button>` : ""}
        </li>
      `;
    }).join("\n")
    : "";

  const formHtml = isFormOpen
    ? `
      <form class="override-form" data-tender-key="${esc(project._tenderKey)}">
        <label>Neuer globaler Relevanzwert (1-10)</label>
        <input name="score_value" type="number" min="1" max="10" step="1" required>
        <label>Begruendung (optional)</label>
        <textarea name="reason_text" maxlength="2000"></textarea>
        <button class="action-btn" type="submit">Override speichern</button>
      </form>
    `
    : "";

  const scoreState = latest
    ? `<p><strong>Aktiver Score:</strong> ${esc(project._effectiveScoreRaw)} (Override, AI: ${esc(normalize((project._effectiveData || project).relevanzbewertung) || "-")})</p>`
    : `<p><strong>Aktiver Score:</strong> ${esc(project._effectiveScoreRaw)} (AI-Original)</p>`;

  if (!hasOverrides && !isFormOpen) {
    return "";
  }

  return `
    <section class="card-section">
      <h3>Relevanz-Overrides</h3>
      ${scoreState}
      <div class="override-history">${historyHtml}</div>
      ${formHtml}
    </section>
  `;
}

function editableFieldKeys(project) {
  const blocked = new Set([
    "_source_type",
    "_tenderKey",
    "_key",
    "_baseScore",
    "_effectiveScore",
    "_effectiveScoreRaw",
    "_scoreFilter",
    "_locationTags",
    "_category",
    "_source",
    "_search",
    "_dateObj",
    "_deadlineObj",
    "_costValue",
    "_effectiveData",
  ]);
  return Object.keys(project).filter((key) => !blocked.has(key) && !key.startsWith("_")).sort();
}

function prettyFieldLabel(key) {
  return key
    .replace(/_/g, " ")
    .replace(/\b\w/g, (c) => c.toUpperCase());
}

function renderFieldEdits(project) {
  if (!state.auth.user) return "";

  const editHistory = getFieldEditHistory(project._tenderKey);
  const isFormOpen = state.ui.openFieldEditForms.has(project._tenderKey);
  const hasEdits = editHistory.length > 0;
  const myUserId = state.auth.user.id;
  const effective = project._effectiveData || project;

  if (!hasEdits && !isFormOpen) return "";

  const historyHtml = hasEdits
    ? editHistory.slice().reverse().map((entry, idx) => {
      const isLatest = idx === 0;
      return `
        <li class="override-item ${isLatest ? "latest" : ""}">
          <div class="override-head">
            <span>${esc(entry.user_email || "Unbekannt")}</span>
            <span>${esc(formatDateTime(entry.created_at))}</span>
          </div>
          <p>${isLatest ? "<strong>Aktive Feldkorrektur</strong>" : "Fruehere Feldkorrektur"}</p>
          ${entry.user_id === myUserId ? `<button class="action-btn field-edit-delete" type="button" data-field-edit-id="${esc(entry.id)}" data-tender-key="${esc(project._tenderKey)}">Korrektur loeschen</button>` : ""}
        </li>
      `;
    }).join("\n")
    : "";

  const fields = editableFieldKeys(project);
  const fieldsHtml = fields.map((fieldKey) => {
    const val = normalize(effective[fieldKey]);
    const inputId = `field_${toAsciiKey(`${project._tenderKey}_${fieldKey}`)}`;
    const isLong = val.length > 120 || /\n/.test(val);
    if (isLong) {
      return `
        <div class="field-edit-field">
          <label for="${esc(inputId)}">${esc(prettyFieldLabel(fieldKey))}</label>
          <textarea id="${esc(inputId)}" name="${esc(fieldKey)}">${esc(val)}</textarea>
        </div>
      `;
    }
    return `
      <div class="field-edit-field">
        <label for="${esc(inputId)}">${esc(prettyFieldLabel(fieldKey))}</label>
        <input id="${esc(inputId)}" name="${esc(fieldKey)}" type="text" value="${esc(val)}">
      </div>
    `;
  }).join("\n");

  const formHtml = isFormOpen
    ? `
      <form class="field-edit-form" data-tender-key="${esc(project._tenderKey)}">
        <div class="field-edit-grid">${fieldsHtml}</div>
        <div class="field-edit-actions">
          <button class="action-btn submit-btn" type="submit">Felder speichern</button>
        </div>
      </form>
    `
    : "";

  return `
    <section class="card-section">
      <h3>Feldkorrekturen</h3>
      <ul class="override-history">${historyHtml}</ul>
      ${formHtml}
    </section>
  `;
}

function renderCardActions(project) {
  if (!state.auth.user) return "";
  const isCommentOpen = state.ui.openCommentForms.has(project._tenderKey);
  const isOverrideOpen = state.ui.openOverrideForms.has(project._tenderKey);
  const isFieldEditOpen = state.ui.openFieldEditForms.has(project._tenderKey);
  const verified = isVerified(project._tenderKey);
  const hasOpenRequest = hasOpenApprovalRequest(project._tenderKey);
  const latestRequest = getLatestApprovalRequest(project._tenderKey);
  const canResolveRequest = hasOpenRequest && latestRequest && latestRequest.user_id === state.auth.user.id;
  const requestLabel = hasOpenRequest
    ? (canResolveRequest ? "Resolve my AKQ request" : "AKQ request open")
    : "Request approval by AKQ";

  return `
    <div class="card-actions">
      <button class="action-btn comment-toggle ${isCommentOpen ? "active" : ""}" type="button" data-tender-key="${esc(project._tenderKey)}">${isCommentOpen ? "Kommentarfeld ausblenden" : "Add a comment"}</button>
      <button class="action-btn override-toggle ${isOverrideOpen ? "active" : ""}" type="button" data-tender-key="${esc(project._tenderKey)}">${isOverrideOpen ? "Override-Feld ausblenden" : "Override score"}</button>
      <button class="action-btn field-edit-toggle ${isFieldEditOpen ? "active" : ""}" type="button" data-tender-key="${esc(project._tenderKey)}">${isFieldEditOpen ? "Edit fields ausblenden" : "Edit card fields"}</button>
      <button class="action-btn verify-toggle ${verified ? "active" : ""}" type="button" data-tender-key="${esc(project._tenderKey)}">Verified by Akquistions team</button>
      <button class="action-btn request-approval-toggle ${hasOpenRequest ? "active" : ""}" type="button" data-tender-key="${esc(project._tenderKey)}" ${hasOpenRequest && !canResolveRequest ? "disabled" : ""}>${requestLabel}</button>
    </div>
  `;
}

function renderCard(project) {
  const view = project._effectiveData || project;
  const scoreClass = scoreBadgeClass(Math.max(project._effectiveScore, 0));

  const nummer = esc(normalize(view.id) || "-");
  const datum = esc(parseDisplayDate(view.date));
  const abgabefrist = esc(parseDisplayDate(view.abgabefrist));
  const titel = esc(normalize(view.titel || view.title) || "-");
  const kurzbeschreibung = esc(normalize(view.kurzbeschreibung) || "-").replace(/\n/g, "<br>");

  const lageText = normalize(view.projektlage) || "-";
  const mapsLink = buildGoogleMapsLink(lageText);
  const lage = mapsLink
    ? `<a href="${esc(mapsLink)}" target="_blank" rel="noopener noreferrer">${esc(lageText)}</a>`
    : esc(lageText);

  const categoryValue = esc(normalize(view.category) || "-");
  const leistungen = esc(normalize(view.leistungen) || "-").replace(/\n/g, "<br>");
  const wettbewerbsart = esc(normalize(view.wettbewerb_art) || "-").replace(/\n/g, "<br>");
  const gewinner = esc(normalize(view.gewinner) || "-").replace(/\n/g, "<br>");
  const gewinnerRolle = esc(normalize(view.gewinner_rolle) || "-").replace(/\n/g, "<br>");
  const erklaerung = esc(
    normalize(view.relevanzbewertung_erklaerung) || normalize(view.relevanzbewertung_begruendung) || "-"
  ).replace(/\n/g, "<br>");

  const [detailLink, pdfLink] = buildNoticeLinks(view);
  const sourceType = project._source;
  const sourceTypeLabel = sourceLabel(sourceType);
  const latestOverride = getLatestOverride(project._tenderKey);
  const overwrittenByHtml = latestOverride
    ? `<p class="override-byline">Overwritten by ${esc(latestOverride.user_email || "Unbekannt")}</p>`
    : "";
  const verifiedBadgeHtml = isVerified(project._tenderKey) ? '<span class="confirm-badge">Verified by AKQ</span>' : "";
  const requestBadgeHtml = hasOpenApprovalRequest(project._tenderKey) ? '<span class="status-badge request-open">AKQ Request Open</span>' : "";
  const editedBadgeHtml = hasFieldEdits(project._tenderKey) ? '<span class="status-badge edited">Fields Edited</span>' : "";

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
      ["Baukosten kg300/400", formatMioEur(view.baukosten_kg300_400)],
      ["Erklaerung der Baukosten", normalize(view.baukosten_erklaerung) || "-"],
    ])
    : renderNamedRows([
      ["Baukosten kg300/400", formatMioEur(view.baukosten_kg300_400)],
      ["Erklaerung der Baukosten", normalize(view.baukosten_erklaerung) || "-"],
      ["Honorar sbp", formatMioEur(view.geschaetztes_honorar_sbp)],
      ["Erklaerung Honorar SBP", normalize(view.honorar_erklaerung) || "-"],
    ]);

  const weitereTable = sourceType === "results"
    ? renderNamedRows([
      ["Wettbewerbsart", normalize(view.wettbewerb_art) || "-"],
      ["Gewinner", normalize(view.gewinner) || "-"],
      ["Gewinner Rolle", normalize(view.gewinner_rolle) || "-"],
      ["Gewinner Kontakt", normalize(view.gewinner_kontakt) || "-"],
      ["Projektbeteiligte", normalize(view.projektbeteiligte) || "-"],
      ["Naechste Schritte", normalize(view.naechste_schritte) || "-"],
      ["Notes", normalize(view.notes) || "-"],
    ])
    : renderNamedRows([
      ["Abgabefrist", parseDisplayDate(view.abgabefrist)],
      ["Leistungen", normalize(view.leistungen) || "-"],
      ["Umfang", normalize(view.umfang) || "-"],
      ["Zuschlagskriterien", normalize(view.zuschlagskriterien) || "-"],
      ["Referenzen/Qualifikationen", normalize(view.referenzen_qualifikationen) || "-"],
      ["Auftraggeber", normalize(view.auftraggeber) || "-"],
      ["Notes", normalize(view.notes) || "-"],
    ]);

  return `
    <article class="project ${scoreClass}" data-key="${esc(project._key)}" data-tender-key="${esc(project._tenderKey)}">
      <div class="card-topline">
        <div class="card-topline-left">
          <div class="source-badge">${esc(sourceTypeLabel)}</div>
          ${verifiedBadgeHtml}
          ${requestBadgeHtml}
          ${editedBadgeHtml}
        </div>
        <div class="tender-key">${esc(project._tenderKey)}</div>
      </div>
      <header class="project-head">
        <div class="score-pill ${scoreClass}">${esc(project._effectiveScoreRaw)}</div>
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
        ${overwrittenByHtml}
      </section>

      ${renderCardActions(project)}

      <details class="details-block">
        <summary>Kostenschaetzung</summary>
        <div class="details-content">${kostenTable}</div>
      </details>

      <details class="details-block">
        <summary>Weitere Informationen</summary>
        <div class="details-content">${weitereTable}</div>
      </details>

      ${renderComments(project)}
      ${renderOverrides(project)}
      ${renderFieldEdits(project)}
    </article>
  `;
}

function ensureCardsEmptyState() {
  const pool = document.getElementById("cardsPool");
  if (!pool) return;
  const hasCards = !!pool.querySelector("article.project");
  if (!hasCards) {
    pool.innerHTML = "<p>Keine passenden Eintraege gefunden.</p>";
  }
}

function findCardElementByTenderKey(tenderKey) {
  const pool = document.getElementById("cardsPool");
  if (!pool) return null;
  return [...pool.querySelectorAll("article.project")].find((el) => normalize(el.getAttribute("data-tender-key")) === tenderKey) || null;
}

function refreshSingleCardByTenderKey(tenderKey) {
  const row = state.rows.find((r) => r._tenderKey === tenderKey);
  if (!row) {
    refreshUI();
    return;
  }

  const existing = findCardElementByTenderKey(tenderKey);
  const shouldShow = !!state.auth.user && matchesFilters(row);

  if (!existing && shouldShow) {
    // Positioning depends on global sort/filter order.
    refreshUI();
    return;
  }
  if (!existing && !shouldShow) return;

  if (!shouldShow) {
    existing.remove();
    updateCounts();
    ensureCardsEmptyState();
    return;
  }

  const tmp = document.createElement("div");
  tmp.innerHTML = renderCard(row);
  const nextCard = tmp.firstElementChild;
  if (!nextCard) {
    refreshUI();
    return;
  }
  existing.replaceWith(nextCard);
  updateCounts();
}

function activeFilteredRows() {
  if (!state.auth.user) return [];
  return sortRows(state.rows.filter((row) => matchesFilters(row)));
}

function countFilteredRows() {
  if (!state.auth.user) return 0;
  return state.rows.reduce((acc, row) => acc + (matchesFilters(row) ? 1 : 0), 0);
}

function updateCounts() {
  const countEl = document.getElementById("resultsCount");
  const count = countFilteredRows();
  countEl.textContent = `Filter preview: ${count} / ${state.rows.length} cards`;

  const meta = document.getElementById("metaInfo");
  meta.textContent = `Treffer: ${state.rows.length} (all loaded rows)`;
}

function renderCardsArea() {
  const pool = document.getElementById("cardsPool");
  if (!state.auth.user) {
    pool.innerHTML = "";
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
  applyDashboardGate();
  applyEffectiveScores();
  updateCounts();
  updateFilterCountBadges();
  renderCardsArea();
}

function activateSingleSelectFilter(group, value, clickedButton) {
  if (group === "type") state.filters.type = value;
  if (group === "location") state.filters.location = value;
  if (group === "category") state.filters.category = value;

  document.querySelectorAll(`.filter-btn[data-filter-group="${group}"]`).forEach((b) => b.classList.remove("active"));
  clickedButton.classList.add("active");
}

function splitChunks(values, size) {
  const out = [];
  for (let i = 0; i < values.length; i += size) out.push(values.slice(i, i + size));
  return out;
}

async function fetchByTenderKeys(table, selectColumns, tenderKeys) {
  const all = [];
  const chunks = splitChunks(tenderKeys, 200);
  for (const chunk of chunks) {
    const { data, error } = await state.auth.client
      .from(table)
      .select(selectColumns)
      .in("tender_key", chunk)
      .order("created_at", { ascending: true });

    if (error) throw error;
    all.push(...(data || []));
  }
  return all;
}

function mapByTenderKey(rows) {
  const out = new Map();
  rows.forEach((row) => {
    const key = normalize(row.tender_key);
    if (!out.has(key)) out.set(key, []);
    out.get(key).push(row);
  });
  return out;
}

function mergeByTenderKeys(existingMap, replacementMap, tenderKeys) {
  const merged = new Map(existingMap);
  tenderKeys.forEach((key) => {
    merged.delete(key);
    const rows = replacementMap.get(key) || [];
    if (rows.length) merged.set(key, rows);
  });
  return merged;
}

function setFormPending(form, isPending, pendingLabel = "Speichern...") {
  if (!(form instanceof HTMLFormElement)) return;
  const submitBtn = form.querySelector('button[type="submit"]');
  if (isPending) {
    if (state.ui.pendingForms.has(form)) return;
    state.ui.pendingForms.add(form);
    form.setAttribute("data-busy", "1");
    if (submitBtn) {
      submitBtn.setAttribute("data-original-label", submitBtn.textContent || "");
      submitBtn.disabled = true;
      submitBtn.textContent = pendingLabel;
    }
    return;
  }

  state.ui.pendingForms.delete(form);
  form.removeAttribute("data-busy");
  if (submitBtn) {
    submitBtn.disabled = false;
    const original = submitBtn.getAttribute("data-original-label");
    if (original !== null) {
      submitBtn.textContent = original;
      submitBtn.removeAttribute("data-original-label");
    }
  }
}

async function runWithButtonLock(button, pendingLabel, task) {
  if (!(button instanceof HTMLButtonElement)) return;
  if (button.getAttribute("data-busy") === "1") return;
  const originalText = button.textContent || "";
  button.setAttribute("data-busy", "1");
  button.disabled = true;
  if (pendingLabel) button.textContent = pendingLabel;
  try {
    await task();
  } finally {
    button.disabled = false;
    button.removeAttribute("data-busy");
    button.textContent = originalText;
  }
}

async function loadRemoteDataForTenderKeys(tenderKeys, options = {}) {
  const { replaceAll = false, refresh = true } = options;
  const uniqueKeys = [...new Set((tenderKeys || []).map((k) => normalize(k)).filter(Boolean))];
  if (!uniqueKeys.length) {
    if (replaceAll) {
      state.remote.commentsByKey = new Map();
      state.remote.overridesByKey = new Map();
      state.remote.verificationsByKey = new Map();
      state.remote.approvalRequestsByKey = new Map();
      state.remote.fieldEditsByKey = new Map();
      if (refresh) refreshUI();
    }
    return;
  }

  const requestId = state.remote.loadRequestId + 1;
  state.remote.loadRequestId = requestId;
  const [comments, overrides, verifications, approvalRequests, fieldEdits] = await Promise.all([
    fetchByTenderKeys("tender_comments", "id,tender_key,user_id,user_email,comment_text,created_at,updated_at", uniqueKeys),
    fetchByTenderKeys("tender_score_overrides", "id,tender_key,user_id,user_email,score_value,reason_text,created_at", uniqueKeys),
    fetchByTenderKeys("tender_verifications", "id,tender_key,user_id,user_email,is_verified,created_at", uniqueKeys),
    fetchByTenderKeys("tender_approval_requests", "id,tender_key,user_id,user_email,status,created_at,updated_at", uniqueKeys),
    fetchByTenderKeys("tender_field_edits", "id,tender_key,user_id,user_email,edit_payload,created_at", uniqueKeys),
  ]);

  if (requestId !== state.remote.loadRequestId) return;

  const commentsMap = mapByTenderKey(comments || []);
  const overridesMap = mapByTenderKey(overrides || []);
  const verificationsMap = mapByTenderKey(verifications || []);
  const requestsMap = mapByTenderKey(approvalRequests || []);
  const fieldEditsMap = mapByTenderKey(fieldEdits || []);

  if (replaceAll) {
    state.remote.commentsByKey = commentsMap;
    state.remote.overridesByKey = overridesMap;
    state.remote.verificationsByKey = verificationsMap;
    state.remote.approvalRequestsByKey = requestsMap;
    state.remote.fieldEditsByKey = fieldEditsMap;
  } else {
    state.remote.commentsByKey = mergeByTenderKeys(state.remote.commentsByKey, commentsMap, uniqueKeys);
    state.remote.overridesByKey = mergeByTenderKeys(state.remote.overridesByKey, overridesMap, uniqueKeys);
    state.remote.verificationsByKey = mergeByTenderKeys(state.remote.verificationsByKey, verificationsMap, uniqueKeys);
    state.remote.approvalRequestsByKey = mergeByTenderKeys(state.remote.approvalRequestsByKey, requestsMap, uniqueKeys);
    state.remote.fieldEditsByKey = mergeByTenderKeys(state.remote.fieldEditsByKey, fieldEditsMap, uniqueKeys);
  }

  if (refresh) refreshUI();
}

async function loadRemoteData() {
  if (!state.auth.enabled || !state.auth.user) {
    state.remote.commentsByKey = new Map();
    state.remote.overridesByKey = new Map();
    state.remote.verificationsByKey = new Map();
    state.remote.approvalRequestsByKey = new Map();
    state.remote.fieldEditsByKey = new Map();
    state.ui.openCommentForms.clear();
    state.ui.openOverrideForms.clear();
    state.ui.openFieldEditForms.clear();
    refreshUI();
    return;
  }

  const tenderKeys = [...new Set(state.rows.map((row) => row._tenderKey))];
  if (!tenderKeys.length) {
    await loadRemoteDataForTenderKeys([], { replaceAll: true, refresh: true });
    return;
  }

  try {
    await loadRemoteDataForTenderKeys(tenderKeys, { replaceAll: true, refresh: true });
  } catch (err) {
    authMessage(`Supabase-Daten konnten nicht geladen werden: ${err.message || err}`, true);
  }
}

async function submitComment(form) {
  if (!state.auth.user) {
    authMessage("Bitte zuerst anmelden.", true);
    return;
  }

  const tenderKey = normalize(form.getAttribute("data-tender-key"));
  const fd = new FormData(form);
  const commentText = normalize(fd.get("comment_text"));
  if (!tenderKey || !commentText) {
    authMessage("Kommentar darf nicht leer sein.", true);
    return;
  }

  const { data, error } = await state.auth.client.from("tender_comments")
    .insert({
      tender_key: tenderKey,
      user_id: state.auth.user.id,
      user_email: state.auth.user.email,
      comment_text: commentText,
    })
    .select("id,tender_key,user_id,user_email,comment_text,created_at,updated_at")
    .single();

  if (error) throw error;
  const existing = state.remote.commentsByKey.get(tenderKey) || [];
  state.remote.commentsByKey.set(tenderKey, [...existing, data]);
  state.ui.openCommentForms.delete(tenderKey);
  form.reset();
  authMessage("Kommentar gespeichert.");
  refreshSingleCardByTenderKey(tenderKey);
}

async function deleteComment(commentId, tenderKey) {
  const { error } = await state.auth.client.from("tender_comments").delete().eq("id", commentId);
  if (error) throw error;

  const key = tenderKey || "";
  if (key) {
    const existing = state.remote.commentsByKey.get(key) || [];
    const next = existing.filter((entry) => normalize(entry.id) !== commentId);
    if (next.length) state.remote.commentsByKey.set(key, next);
    else state.remote.commentsByKey.delete(key);
  }

  authMessage("Kommentar geloescht.");
  if (key) refreshSingleCardByTenderKey(key);
  else await loadRemoteData();
}

async function submitOverride(form) {
  if (!state.auth.user) {
    authMessage("Bitte zuerst anmelden.", true);
    return;
  }

  const tenderKey = normalize(form.getAttribute("data-tender-key"));
  const fd = new FormData(form);
  const scoreValueRaw = normalize(fd.get("score_value"));
  const reasonText = normalize(fd.get("reason_text"));
  const scoreValue = Number(scoreValueRaw);
  if (!tenderKey || !Number.isInteger(scoreValue) || scoreValue < 1 || scoreValue > 10) {
    authMessage("Score muss zwischen 1 und 10 liegen.", true);
    return;
  }

  const { error } = await state.auth.client.from("tender_score_overrides").insert({
    tender_key: tenderKey,
    user_id: state.auth.user.id,
    user_email: state.auth.user.email,
    score_value: scoreValue,
    reason_text: reasonText || null,
  });

  if (error) throw error;
  state.ui.openOverrideForms.delete(tenderKey);
  form.reset();
  authMessage("Override gespeichert.");
  await loadRemoteDataForTenderKeys([tenderKey], { refresh: true });
}

async function deleteOverride(overrideId, tenderKey) {
  const { error } = await state.auth.client.from("tender_score_overrides").delete().eq("id", overrideId);
  if (error) throw error;
  authMessage("Override geloescht.");
  if (tenderKey) await loadRemoteDataForTenderKeys([tenderKey], { refresh: true });
  else await loadRemoteData();
}

async function toggleVerification(tenderKey) {
  if (!state.auth.user) {
    authMessage("Bitte zuerst anmelden.", true);
    return;
  }
  const currentlyVerified = isVerified(tenderKey);
  const { error } = await state.auth.client.from("tender_verifications").insert({
    tender_key: tenderKey,
    user_id: state.auth.user.id,
    user_email: state.auth.user.email,
    is_verified: !currentlyVerified,
  });
  if (error) throw error;
  authMessage(!currentlyVerified ? "Karte verifiziert." : "Verifizierung entfernt.");
  await loadRemoteDataForTenderKeys([tenderKey], { refresh: true });
}

async function toggleApprovalRequest(tenderKey) {
  if (!state.auth.user) {
    authMessage("Bitte zuerst anmelden.", true);
    return;
  }

  const latest = getLatestApprovalRequest(tenderKey);
  const isOpen = !!(latest && normalize(latest.status).toLowerCase() === "open");
  if (isOpen) {
    if (latest.user_id !== state.auth.user.id) {
      authMessage("Nur der Ersteller kann diese Anfrage schliessen.", true);
      return;
    }
    const { error } = await state.auth.client.from("tender_approval_requests")
      .update({ status: "resolved" })
      .eq("id", latest.id);
    if (error) throw error;
    authMessage("AKQ-Anfrage geschlossen.");
    await loadRemoteDataForTenderKeys([tenderKey], { refresh: true });
    return;
  }

  const { error } = await state.auth.client.from("tender_approval_requests").insert({
    tender_key: tenderKey,
    user_id: state.auth.user.id,
    user_email: state.auth.user.email,
    status: "open",
  });
  if (error) throw error;
  authMessage("AKQ-Anfrage erstellt.");
  await loadRemoteDataForTenderKeys([tenderKey], { refresh: true });
}

async function submitFieldEdit(form) {
  if (!state.auth.user) {
    authMessage("Bitte zuerst anmelden.", true);
    return;
  }

  const tenderKey = normalize(form.getAttribute("data-tender-key"));
  if (!tenderKey) return;

  const fd = new FormData(form);
  const payload = {};
  for (const [key, value] of fd.entries()) {
    payload[normalize(key)] = normalize(value);
  }

  const { error } = await state.auth.client.from("tender_field_edits").insert({
    tender_key: tenderKey,
    user_id: state.auth.user.id,
    user_email: state.auth.user.email,
    edit_payload: payload,
  });
  if (error) throw error;
  state.ui.openFieldEditForms.delete(tenderKey);
  authMessage("Feldkorrektur gespeichert.");
  await loadRemoteDataForTenderKeys([tenderKey], { refresh: true });
}

async function deleteFieldEdit(fieldEditId, tenderKey) {
  const { error } = await state.auth.client.from("tender_field_edits").delete().eq("id", fieldEditId);
  if (error) throw error;
  authMessage("Feldkorrektur geloescht.");
  if (tenderKey) await loadRemoteDataForTenderKeys([tenderKey], { refresh: true });
  else await loadRemoteData();
}

function bindUi() {
  document.getElementById("authToggleBtn").addEventListener("click", () => {
    state.ui.authCollapsed = !state.ui.authCollapsed;
    applyAuthPanelVisibility();
  });

  document.getElementById("liveSearch").addEventListener("input", (e) => {
    state.filters.query = normalize(e.target.value).toLowerCase();
    refreshUI();
  });

  document.getElementById("startDate").addEventListener("change", (e) => {
    clearDateRangeQuickSelection();
    state.filters.startDate = normalize(e.target.value);
    refreshUI();
  });

  document.getElementById("endDate").addEventListener("change", (e) => {
    clearDateRangeQuickSelection();
    state.filters.endDate = normalize(e.target.value);
    refreshUI();
  });

  document.querySelectorAll(".date-range-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const days = Number.parseInt(btn.getAttribute("data-days") || "", 10);
      applyQuickDateRange(days);
      clearDateRangeQuickSelection();
      btn.classList.add("active");
      refreshUI();
    });
  });

  document.querySelectorAll(".filter-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const group = btn.getAttribute("data-filter-group");
      const value = btn.getAttribute("data-value") || "all";
      if (!group) return;

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
            if (state.filters.scores.size === 0) allBtn.classList.add("active");
            else allBtn.classList.remove("active");
          }
        }
      } else {
        activateSingleSelectFilter(group, value, btn);
      }

      refreshUI();
    });
  });

  document.getElementById("verifiedFilterBtn").addEventListener("click", (e) => {
    state.filters.onlyVerified = !state.filters.onlyVerified;
    e.currentTarget.classList.toggle("active", state.filters.onlyVerified);
    refreshUI();
  });

  document.getElementById("activityFilterBtn").addEventListener("click", (e) => {
    state.filters.onlyWithActivity = !state.filters.onlyWithActivity;
    e.currentTarget.classList.toggle("active", state.filters.onlyWithActivity);
    refreshUI();
  });

  document.getElementById("approvalRequestFilterBtn").addEventListener("click", (e) => {
    state.filters.onlyWithOpenRequest = !state.filters.onlyWithOpenRequest;
    e.currentTarget.classList.toggle("active", state.filters.onlyWithOpenRequest);
    refreshUI();
  });

  document.getElementById("sortOrder").addEventListener("change", (e) => {
    state.filters.sortOrder = normalize(e.target.value) || "score_desc";
    refreshUI();
  });
  document.getElementById("sortOrder").value = state.filters.sortOrder;

  document.getElementById("cardsPool").addEventListener("submit", async (e) => {
    const form = e.target;
    if (!(form instanceof HTMLFormElement)) return;
    if (form.getAttribute("data-busy") === "1") {
      e.preventDefault();
      return;
    }

    if (form.classList.contains("comment-form")) {
      e.preventDefault();
      setFormPending(form, true, "Speichert...");
      try {
        await submitComment(form);
      } catch (err) {
        authMessage(`Kommentar konnte nicht gespeichert werden: ${err.message || err}`, true);
      } finally {
        setFormPending(form, false);
      }
      return;
    }

    if (form.classList.contains("override-form")) {
      e.preventDefault();
      setFormPending(form, true, "Speichert...");
      try {
        await submitOverride(form);
      } catch (err) {
        authMessage(`Override konnte nicht gespeichert werden: ${err.message || err}`, true);
      } finally {
        setFormPending(form, false);
      }
      return;
    }

    if (form.classList.contains("field-edit-form")) {
      e.preventDefault();
      setFormPending(form, true, "Speichert...");
      try {
        await submitFieldEdit(form);
      } catch (err) {
        authMessage(`Feldkorrektur konnte nicht gespeichert werden: ${err.message || err}`, true);
      } finally {
        setFormPending(form, false);
      }
    }
  });

  document.getElementById("cardsPool").addEventListener("click", async (e) => {
    const target = e.target;
    if (!(target instanceof HTMLElement)) return;

    if (target.classList.contains("comment-toggle")) {
      const tenderKey = normalize(target.getAttribute("data-tender-key"));
      if (!tenderKey) return;
      if (state.ui.openCommentForms.has(tenderKey)) state.ui.openCommentForms.delete(tenderKey);
      else state.ui.openCommentForms.add(tenderKey);
      refreshUI();
      return;
    }

    if (target.classList.contains("override-toggle")) {
      const tenderKey = normalize(target.getAttribute("data-tender-key"));
      if (!tenderKey) return;
      if (state.ui.openOverrideForms.has(tenderKey)) state.ui.openOverrideForms.delete(tenderKey);
      else state.ui.openOverrideForms.add(tenderKey);
      refreshUI();
      return;
    }

    if (target.classList.contains("field-edit-toggle")) {
      const tenderKey = normalize(target.getAttribute("data-tender-key"));
      if (!tenderKey) return;
      if (state.ui.openFieldEditForms.has(tenderKey)) state.ui.openFieldEditForms.delete(tenderKey);
      else state.ui.openFieldEditForms.add(tenderKey);
      refreshUI();
      return;
    }

    if (target.classList.contains("verify-toggle")) {
      const tenderKey = normalize(target.getAttribute("data-tender-key"));
      if (!tenderKey) return;
      await runWithButtonLock(target, "Speichert...", async () => {
        try {
          await toggleVerification(tenderKey);
        } catch (err) {
          authMessage(`Verifizierung konnte nicht gespeichert werden: ${err.message || err}`, true);
        }
      });
      return;
    }

    if (target.classList.contains("request-approval-toggle")) {
      const tenderKey = normalize(target.getAttribute("data-tender-key"));
      if (!tenderKey) return;
      await runWithButtonLock(target, "Speichert...", async () => {
        try {
          await toggleApprovalRequest(tenderKey);
        } catch (err) {
          authMessage(`AKQ-Anfrage konnte nicht gespeichert werden: ${err.message || err}`, true);
        }
      });
      return;
    }

    if (target.classList.contains("override-delete")) {
      const overrideId = normalize(target.getAttribute("data-override-id"));
      const tenderKey = normalize(target.getAttribute("data-tender-key"));
      if (!overrideId) return;
      await runWithButtonLock(target, "Loescht...", async () => {
        try {
          await deleteOverride(overrideId, tenderKey);
        } catch (err) {
          authMessage(`Override konnte nicht geloescht werden: ${err.message || err}`, true);
        }
      });
      return;
    }

    if (target.classList.contains("field-edit-delete")) {
      const fieldEditId = normalize(target.getAttribute("data-field-edit-id"));
      const tenderKey = normalize(target.getAttribute("data-tender-key"));
      if (!fieldEditId) return;
      await runWithButtonLock(target, "Loescht...", async () => {
        try {
          await deleteFieldEdit(fieldEditId, tenderKey);
        } catch (err) {
          authMessage(`Feldkorrektur konnte nicht geloescht werden: ${err.message || err}`, true);
        }
      });
      return;
    }

    if (!target.classList.contains("comment-delete")) return;

    const commentId = normalize(target.getAttribute("data-comment-id"));
    const tenderKey = normalize(target.getAttribute("data-tender-key"));
    if (!commentId) return;
    await runWithButtonLock(target, "Loescht...", async () => {
      try {
        await deleteComment(commentId, tenderKey);
      } catch (err) {
        authMessage(`Kommentar konnte nicht geloescht werden: ${err.message || err}`, true);
      }
    });
  });
}

function initializeDateInputs() {
  const today = formatDateInput(new Date());

  const startEl = document.getElementById("startDate");
  const endEl = document.getElementById("endDate");
  startEl.value = today;
  endEl.value = today;

  state.filters.startDate = today;
  state.filters.endDate = today;
}

async function loadWorkbook() {
  const status = document.getElementById("loadStatus");
  const warningsEl = document.getElementById("loadWarnings");

  try {
    if (!window.XLSX) throw new Error("SheetJS library not loaded.");
    const response = await fetch(DATA_URL, { cache: "no-cache" });
    if (!response.ok) throw new Error(`Could not fetch ${DATA_URL}. HTTP ${response.status}`);

    const buffer = await response.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: true });

    const warnings = [];
    const newRows = parseSheetRows(workbook, NEW_SHEET, "new_competition");
    const resultRows = parseSheetRows(workbook, RESULTS_SHEET, "results");

    if (!workbook.SheetNames.includes(NEW_SHEET)) warnings.push(`Worksheet '${NEW_SHEET}' was not found.`);
    if (!workbook.SheetNames.includes(RESULTS_SHEET)) warnings.push(`Worksheet '${RESULTS_SHEET}' was not found.`);

    state.rows = [...newRows, ...resultRows];
    state.rows.forEach((row) => enrichRow(row));

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

function initializeSupabaseClient() {
  const cfg = window.SUPABASE_CONFIG || {};
  const url = normalize(cfg.url);
  const anonKey = normalize(cfg.anonKey);

  if (!url || !anonKey) {
    state.auth.enabled = false;
    state.auth.client = null;
    updateAuthStatus();
    authMessage("Supabase-Konfiguration fehlt. Setze URL + anonKey in supabase-config.js.", true);
    return;
  }

  if (!window.supabase || typeof window.supabase.createClient !== "function") {
    state.auth.enabled = false;
    state.auth.client = null;
    updateAuthStatus();
    authMessage("Supabase library wurde nicht geladen.", true);
    return;
  }

  state.auth.client = window.supabase.createClient(url, anonKey);
  state.auth.enabled = true;
  updateAuthStatus();
}

function bindAuthUi() {
  document.getElementById("registerForm").addEventListener("submit", async (e) => {
    e.preventDefault();
    if (!state.auth.enabled) return;

    const email = normalize(document.getElementById("registerEmail").value);
    const password = normalize(document.getElementById("registerPassword").value);
    const { error } = await state.auth.client.auth.signUp({ email, password });
    if (error) {
      authMessage(`Registrierung fehlgeschlagen: ${error.message}`, true);
      return;
    }
    authMessage("Registrierung gestartet. Bitte E-Mail-Bestaetigung pruefen.");
  });

  document.getElementById("loginForm").addEventListener("submit", async (e) => {
    e.preventDefault();
    if (!state.auth.enabled) return;

    const email = normalize(document.getElementById("loginEmail").value);
    const password = normalize(document.getElementById("loginPassword").value);
    const { error } = await state.auth.client.auth.signInWithPassword({ email, password });
    if (error) {
      authMessage(`Anmeldung fehlgeschlagen: ${error.message}`, true);
      return;
    }
    authMessage("Anmeldung erfolgreich.");
  });

  document.getElementById("logoutBtn").addEventListener("click", async () => {
    if (!state.auth.enabled) return;
    const { error } = await state.auth.client.auth.signOut();
    if (error) {
      authMessage(`Abmeldung fehlgeschlagen: ${error.message}`, true);
      return;
    }
    authMessage("Abgemeldet.");
  });
}

async function initAuthFlow() {
  initializeSupabaseClient();
  bindAuthUi();

  if (!state.auth.enabled) return;

  const { data, error } = await state.auth.client.auth.getSession();
  if (error) {
    authMessage(`Session konnte nicht geladen werden: ${error.message}`, true);
    return;
  }

  state.auth.user = data.session ? data.session.user : null;
  updateAuthStatus();
  await loadRemoteData();

  state.auth.client.auth.onAuthStateChange(async (_event, session) => {
    state.auth.user = session ? session.user : null;
    updateAuthStatus();
    await loadRemoteData();
  });
}

async function bootstrap() {
  applyDashboardGate();
  await loadWorkbook();
  await initAuthFlow();
}

bootstrap();
