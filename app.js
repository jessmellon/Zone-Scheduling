const SHEET_GVIZ_URL =
  "https://docs.google.com/spreadsheets/d/1_KPdGkIe-tQrKEFxqAfJL87VvX73aEPuwVW5G_b4zOI/gviz/tq?tqx=out:json&sheet=Copy%20of%20Dates";
const LIMITS_API_URL =
  "https://script.google.com/macros/s/AKfycby6Jm_C3Z536OSDitlorRvFrNmSFKtqt1TE1-oCSrphTzduvO2DvWa5TWnAe3Up7gu8/exec";
const ALL_ZONES = ["Z1", "Z2", "Z3", "Z4", "Z5"];

const state = {
  events: [],
  currentMonth: null,
  selectedCategory: null,
  selectedEventId: null,
  searchTerm: "",
  viewMode: "events",
  categoryColors: new Map(),
  limits: emptyLimits(),
};

const monthLabel = document.querySelector("#month-label");
const calendarGrid = document.querySelector("#calendar-grid");
const categorySummary = document.querySelector("#category-summary");
const filterList = document.querySelector("#filter-list");
const selectionDetails = document.querySelector("#selection-details");
const statusBanner = document.querySelector("#status-banner");
const lastUpdated = document.querySelector("#last-updated");
const searchInput = document.querySelector("#school-search");
const searchResults = document.querySelector("#search-results");

document.querySelector("#prev-month").addEventListener("click", () => {
  state.currentMonth = addMonths(state.currentMonth, -1);
  render();
});

document.querySelector("#next-month").addEventListener("click", () => {
  state.currentMonth = addMonths(state.currentMonth, 1);
  render();
});

document.querySelector("#today-button").addEventListener("click", () => {
  state.currentMonth = startOfMonth(new Date());
  render();
});

document.querySelector("#reset-filter").addEventListener("click", () => {
  state.selectedCategory = null;
  render();
});

document.querySelector("#events-view-button").addEventListener("click", () => {
  state.viewMode = "events";
  state.selectedEventId = null;
  render();
});

document.querySelector("#staffing-view-button").addEventListener("click", () => {
  state.viewMode = "staffing";
  state.selectedEventId = null;
  render();
});

searchInput.addEventListener("input", (event) => {
  state.searchTerm = event.target.value.trim();
  renderSearchResults();
});

loadCalendar();

async function loadCalendar() {
  setStatus("Loading live data from Google Sheets...");

  try {
    const [sheetData, limits] = await Promise.all([loadGvizData(), loadLimits()]);
    const events = buildEvents(sheetData);

    if (!events.length) {
      throw new Error("No dated rows were found in the Dates tab.");
    }

    state.events = events.sort((left, right) => left.date - right.date);
    state.limits = limits;
    state.currentMonth = startOfMonth(state.events[0].date);
    state.categoryColors = buildCategoryColors(state.events);
    lastUpdated.textContent = `Loaded ${state.events.length} calendar entries`;
    render();
  } catch (error) {
    console.error(error);
    setStatus(`Unable to load the Google Sheet: ${error.message}`);
    lastUpdated.textContent = "Sheet load failed";
  }
}

function loadGvizData() {
  return new Promise((resolve, reject) => {
    const timeoutId = window.setTimeout(() => {
      cleanup();
      reject(new Error("Google Sheets request timed out."));
    }, 15000);

    const cleanup = () => {
      window.clearTimeout(timeoutId);
      if (script.parentNode) {
        script.parentNode.removeChild(script);
      }
      delete window.google;
    };

    const script = document.createElement("script");
    script.src = SHEET_GVIZ_URL;
    script.async = true;
    script.onerror = () => {
      cleanup();
      reject(new Error("The browser could not load the Google Sheets feed."));
    };

    window.google = {
      visualization: {
        Query: {
          setResponse(response) {
            cleanup();
            resolve(response);
          },
        },
      },
    };

    document.head.appendChild(script);
  });
}

async function loadLimits() {
  const response = await fetch(LIMITS_API_URL, {
    method: "GET",
  });

  if (!response.ok) {
    throw new Error(`Limits request failed with status ${response.status}`);
  }

  const payload = await response.json();
  return normalizeLimitsPayload(payload);
}

function buildEvents(sheetData) {
  const columns = sheetData?.table?.cols || [];
  const rows = sheetData?.table?.rows || [];

  if (!columns.length || !rows.length) {
    return [];
  }

  const headers = columns.map((column) => (column.label || column.id || "").trim());
  const indexes = {
    title: findHeaderIndex(headers, ["School"]),
    category: findHeaderIndex(headers, ["Zone"]),
    stars: findHeaderIndex(headers, ["Stars"]),
    date: findHeaderIndex(headers, ["Date"]),
    sent: findHeaderIndex(headers, ["Sent?"]),
    photographers: findHeaderIndex(headers, ["Photographers"]),
    type: findHeaderIndex(headers, ["Type"]),
  };

  return rows
    .map((row, rowIndex) => {
      const parsedDate = parseGvizDate(readCell(row.c, indexes.date, true));
      if (!parsedDate) {
        return null;
      }

      const title = readCell(row.c, indexes.title) || "Untitled";
      const schoolName = normalizeSchoolName(title);
      const category = normalizeZone(readCell(row.c, indexes.category)) || "Unassigned";

      if (category === "Z6") {
        return null;
      }

      return {
        id: `${title}-${rowIndex}-${formatDateKey(parsedDate)}`,
        title,
        schoolName,
        category,
        date: parsedDate,
        stars: readCell(row.c, indexes.stars),
        sent: readCell(row.c, indexes.sent),
        photographers: readCell(row.c, indexes.photographers),
        type: readCell(row.c, indexes.type),
      };
    })
    .filter(Boolean);
}

function render() {
  if (!state.currentMonth || !state.events.length) {
    return;
  }

  renderViewSwitch();
  renderHeader();
  renderSummary();
  renderFilters();
  renderCalendar();
  renderDetails();
  renderSearchResults();
}

function renderViewSwitch() {
  document
    .querySelector("#events-view-button")
    .classList.toggle("active", state.viewMode === "events");
  document
    .querySelector("#staffing-view-button")
    .classList.toggle("active", state.viewMode === "staffing");
}

function renderHeader() {
  monthLabel.textContent = state.currentMonth.toLocaleDateString(undefined, {
    month: "long",
    year: "numeric",
  });

  const visibleEvents = getVisibleEvents();
  if (state.viewMode === "staffing") {
    const totalPhotographers = visibleEvents.reduce(
      (sum, event) => sum + getPhotographerCount(event),
      0
    );
    setStatus(
      `${totalPhotographers} Photog${
        totalPhotographers === 1 ? "" : "s"
      } scheduled${
        state.selectedCategory ? ` for ${state.selectedCategory}` : ""
      }`
    );
    return;
  }

  setStatus(
    `${visibleEvents.length} event${visibleEvents.length === 1 ? "" : "s"} shown${
      state.selectedCategory ? ` for ${state.selectedCategory}` : ""
    }`
  );
}

function renderSummary() {
  const counts =
    state.viewMode === "staffing"
      ? sumPhotographersByCategory(state.events)
      : countByCategory(state.events);

  categorySummary.innerHTML = "";
  counts.forEach(({ category, value }) => {
    const item = document.createElement("div");
    item.className = "summary-item";
    item.innerHTML = `
      <span><span class="dot" style="background:${getCategoryColor(category)}"></span> ${escapeHtml(
        category
      )}</span>
      <span class="summary-count">${value}</span>
    `;
    categorySummary.appendChild(item);
  });
}

function renderFilters() {
  const counts =
    state.viewMode === "staffing"
      ? sumPhotographersByCategory(state.events)
      : countByCategory(state.events);

  filterList.innerHTML = "";
  counts.forEach(({ category, value }) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `filter-chip${
      state.selectedCategory === category ? " active" : ""
    }`;
    button.innerHTML = `
      <span class="dot" style="background:${getCategoryColor(category)}"></span>
      <span>${escapeHtml(category)}</span>
      <span class="summary-count">${value}</span>
    `;
    button.addEventListener("click", () => {
      state.selectedCategory =
        state.selectedCategory === category ? null : category;
      render();
    });
    filterList.appendChild(button);
  });
}

function renderCalendar() {
  const monthStart = startOfMonth(state.currentMonth);
  const monthEnd = endOfMonth(state.currentMonth);
  const gridStart = startOfWeek(monthStart);
  const gridEnd = endOfWeek(monthEnd);
  const visibleEvents = getVisibleEvents();
  const todayKey = formatDateKey(new Date());

  calendarGrid.innerHTML = "";

  for (
    let cursor = new Date(gridStart);
    cursor <= gridEnd;
    cursor = addDays(cursor, 1)
  ) {
    const dayKey = formatDateKey(cursor);
    const dayEvents = visibleEvents.filter(
      (event) => formatDateKey(event.date) === dayKey
    );
    const cell = document.createElement("div");
    cell.className = "day-cell";
    const isWeekend = cursor.getDay() === 0 || cursor.getDay() === 6;

    if (cursor.getMonth() !== monthStart.getMonth()) {
      cell.classList.add("is-other-month");
    }

    if (dayKey === todayKey) {
      cell.classList.add("is-today");
    }

    if (isWeekend) {
      cell.classList.add("is-weekend");
    }

    if (state.viewMode === "staffing") {
      applyCapacityClass(
        cell,
        getCapacityStatus(
          getTotalPhotographers(dayEvents),
          getDayLimit(dayKey)
        )
      );
    }

    const groupsMarkup =
      state.viewMode === "staffing"
        ? renderStaffingGroups(dayKey, dayEvents)
        : renderEventGroups(dayEvents);

    cell.innerHTML = `
      <div class="day-heading">
        <div class="day-number">${cursor.getDate()}</div>
        <div class="day-name">${escapeHtml(
          cursor.toLocaleDateString(undefined, { weekday: "short" })
        )}</div>
      </div>
      <div class="day-groups">${groupsMarkup}</div>
    `;

    if (state.viewMode === "events") {
      cell.querySelectorAll("[data-event-id]").forEach((button) => {
        button.addEventListener("click", () => {
          state.selectedEventId = button.dataset.eventId;
          renderDetails();
          renderCalendar();
        });
      });
    } else {
      cell.querySelectorAll("[data-day-limit]").forEach((input) => {
        input.addEventListener("change", async (event) => {
          await setDayLimit(dayKey, event.target.value);
          render();
        });
      });

      cell.querySelectorAll("[data-zone-limit]").forEach((input) => {
        input.addEventListener("change", async (event) => {
          await setZoneLimit(dayKey, event.target.dataset.zone, event.target.value);
          render();
        });
      });
    }

    calendarGrid.appendChild(cell);
  }
}

function renderEventGroups(dayEvents) {
  return countByCategory(dayEvents)
    .map(({ category }) => {
      const categoryEvents = dayEvents.filter((event) => event.category === category);

      const itemsMarkup = categoryEvents
        .map((event) => {
          return `
            <button
              class="event-pill${
                state.selectedEventId === event.id ? " is-selected" : ""
              }"
              type="button"
              data-event-id="${escapeHtml(event.id)}"
              style="border-left-color:${getCategoryColor(event.category)}"
            >
              <span class="event-title">${escapeHtml(event.title)}</span>
              <span class="event-meta">${escapeHtml(
                [event.type, event.stars].filter(Boolean).join(" • ")
              )}</span>
            </button>
          `;
        })
        .join("");

      return `
        <div class="group-block">
          <div class="group-label">${escapeHtml(category)}</div>
          ${itemsMarkup}
        </div>
      `;
    })
    .join("");
}

function renderStaffingGroups(dayKey, dayEvents) {
  const totalPhotographers = getTotalPhotographers(dayEvents);
  const dayLimit = getDayLimit(dayKey);
  const dayStatus = getCapacityStatus(totalPhotographers, dayLimit);
  const displayedZones = getDisplayedZones();

  const plannerMarkup = `
    <div class="planner-card ${getCapacityClassName(dayStatus)}">
      <div class="planner-row">
        <span class="planner-title">Day total</span>
        <span class="planner-total">${totalPhotographers} scheduled</span>
      </div>
      <label class="limit-field">
        <span class="limit-label">Daily limit</span>
        <input
          class="limit-input"
          type="number"
          min="0"
          step="1"
          value="${escapeHtml(dayLimit ?? "")}"
          data-day-limit="true"
        />
      </label>
    </div>
  `;

  const zoneTotals = buildZoneTotals(dayEvents);
  const zoneMarkup = displayedZones
    .map((category) => {
      const value = zoneTotals.get(category)?.photographers || 0;
      const totalEvents = zoneTotals.get(category)?.schools || 0;
      const zoneLimit = getZoneLimit(dayKey, category);
      const zoneStatus = getCapacityStatus(value, zoneLimit);
      return `
        <div class="group-block">
          <div class="group-label">${escapeHtml(category)}</div>
          <div
            class="event-pill summary-pill ${getCapacityClassName(zoneStatus)}"
            style="border-left-color:${getCategoryColor(category)}"
          >
            <span class="event-title">${value} Photog${
              value === 1 ? "" : "s"
            }</span>
            <span class="event-meta">${totalEvents} school${
              totalEvents === 1 ? "" : "s"
            }</span>
            <label class="limit-field zone-limit-field">
              <span class="limit-label">Limit</span>
              <input
                class="limit-input"
                type="number"
                min="0"
                step="1"
                value="${escapeHtml(zoneLimit ?? "")}"
                data-zone-limit="true"
                data-zone="${escapeHtml(category)}"
              />
            </label>
          </div>
        </div>
      `;
    })
    .join("");

  return `${plannerMarkup}${zoneMarkup}`;
}

function renderDetails() {
  if (state.viewMode === "staffing") {
    selectionDetails.className = "details-card";
    selectionDetails.innerHTML = `
      <h3>Zone staffing view</h3>
      <p>
        Each day shows the total number of photographers scheduled in each zone.
      </p>
      <div class="detail-grid">
        ${renderDetailItem("Basis", "Sum of Photographers column")}
        ${renderDetailItem("Warnings", "Yellow = within 2, red = at limit, dark red = over")}
        ${renderDetailItem(
          "Filter",
          state.selectedCategory || "All zones"
        )}
      </div>
    `;
    return;
  }

  const event = state.events.find((item) => item.id === state.selectedEventId);

  if (!event) {
    selectionDetails.className = "details-empty";
    selectionDetails.textContent =
      "Select an event on the calendar to inspect it.";
    return;
  }

  selectionDetails.className = "details-card";
  selectionDetails.innerHTML = `
    <h3>${escapeHtml(event.title)}</h3>
    <p>${escapeHtml(
      event.date.toLocaleDateString(undefined, {
        weekday: "long",
        month: "long",
        day: "numeric",
        year: "numeric",
      })
    )}</p>
    <div class="detail-grid">
      ${renderDetailItem("Zone", event.category)}
      ${renderDetailItem("Stars", event.stars)}
      ${renderDetailItem("Type", event.type)}
      ${renderDetailItem("Sent?", event.sent)}
      ${renderDetailItem("Photographers", event.photographers)}
    </div>
  `;
}

function renderSearchResults() {
  const term = state.searchTerm.trim().toLowerCase();

  if (!term) {
    searchResults.textContent = "Start typing to see a school's dates.";
    return;
  }

  const matchingEvents = state.events.filter((event) =>
    event.schoolName.toLowerCase().includes(term)
  );

  if (!matchingEvents.length) {
    searchResults.textContent = "No matching schools found.";
    return;
  }

  const grouped = groupEventsBySchool(matchingEvents);
  const cardsMarkup = grouped
    .map(({ schoolName, events }) => {
      const itemsMarkup = events
        .sort((left, right) => left.date - right.date)
        .map((event) => {
          return `
            <div class="search-result-item">
              <span class="search-result-date">${escapeHtml(
                event.date.toLocaleDateString(undefined, {
                  month: "short",
                  day: "numeric",
                  year: "numeric",
                })
              )}</span>
              <span class="search-result-meta">${escapeHtml(
                [event.category, event.type].filter(Boolean).join(" • ")
              )}</span>
            </div>
          `;
        })
        .join("");

      return `
        <div class="search-result-card">
          <h3>${escapeHtml(schoolName)}</h3>
          <div class="search-result-list">${itemsMarkup}</div>
        </div>
      `;
    })
    .join("");

  searchResults.innerHTML = cardsMarkup;
}

function renderDetailItem(label, value) {
  return `
    <div class="detail-item">
      <strong>${escapeHtml(label)}</strong>
      <span>${escapeHtml(value || "Not provided")}</span>
    </div>
  `;
}

function getVisibleEvents() {
  if (!state.selectedCategory) {
    return state.events;
  }

  return state.events.filter((event) => event.category === state.selectedCategory);
}

function countByCategory(events) {
  const counts = new Map();

  events.forEach((event) => {
    counts.set(event.category, (counts.get(event.category) || 0) + 1);
  });

  return [...counts.entries()]
    .map(([category, count]) => ({ category, value: count }))
    .sort(compareZoneEntries);
}

function groupEventsBySchool(events) {
  const groups = new Map();

  events.forEach((event) => {
    if (!groups.has(event.schoolName)) {
      groups.set(event.schoolName, []);
    }
    groups.get(event.schoolName).push(event);
  });

  return [...groups.entries()]
    .map(([schoolName, groupedEvents]) => ({ schoolName, events: groupedEvents }))
    .sort((left, right) => left.schoolName.localeCompare(right.schoolName));
}

function sumPhotographersByCategory(events) {
  const totals = new Map();

  events.forEach((event) => {
    totals.set(
      event.category,
      (totals.get(event.category) || 0) + getPhotographerCount(event)
    );
  });

  return [...totals.entries()]
    .map(([category, value]) => ({ category, value }))
    .sort(compareZoneEntries);
}

function buildZoneTotals(events) {
  const totals = new Map(
    ALL_ZONES.map((zone) => [zone, { photographers: 0, schools: 0 }])
  );

  events.forEach((event) => {
    const zone = totals.get(event.category);
    if (!zone) {
      return;
    }
    zone.photographers += getPhotographerCount(event);
    zone.schools += 1;
  });

  return totals;
}

function getDisplayedZones() {
  if (state.viewMode === "staffing" && state.selectedCategory) {
    return [state.selectedCategory];
  }

  return ALL_ZONES;
}

function buildCategoryColors(events) {
  const palette = [
    "#2563eb",
    "#16a34a",
    "#dc2626",
    "#9333ea",
    "#ea580c",
    "#0f766e",
    "#c026d3",
    "#65a30d",
  ];

  const categories = [...new Set(events.map((event) => event.category))].sort();
  return new Map(
    categories.map((category, index) => [category, palette[index % palette.length]])
  );
}

function getCategoryColor(category) {
  return state.categoryColors.get(category) || "#475569";
}

function getPhotographerCount(event) {
  const parsed = Number.parseFloat(event.photographers);
  return Number.isFinite(parsed) ? parsed : 0;
}

function getTotalPhotographers(events) {
  return events.reduce((sum, event) => sum + getPhotographerCount(event), 0);
}

function getCapacityStatus(actual, limit) {
  if (!Number.isFinite(limit)) {
    return "none";
  }
  if (actual > limit) {
    return "critical";
  }
  if (actual === limit) {
    return "danger";
  }
  if (actual === limit - 1 || actual === limit - 2) {
    return "warning";
  }
  return "none";
}

function getCapacityClassName(status) {
  if (status === "warning") {
    return "capacity-warning";
  }
  if (status === "danger") {
    return "capacity-danger";
  }
  if (status === "critical") {
    return "capacity-critical";
  }
  return "";
}

function applyCapacityClass(element, status) {
  element.classList.remove("capacity-warning", "capacity-danger", "capacity-critical");
  const className = getCapacityClassName(status);
  if (className) {
    element.classList.add(className);
  }
}

function getDayLimit(dayKey) {
  const value = state.limits.days?.[dayKey];
  return parseLimitValue(value);
}

function getZoneLimit(dayKey, zone) {
  const value = state.limits.zones?.[dayKey]?.[zone];
  return parseLimitValue(value);
}

async function setDayLimit(dayKey, rawValue) {
  if (!state.limits.days) {
    state.limits.days = {};
  }

  const normalized = normalizeLimitInput(rawValue);
  const previousValue = state.limits.days[dayKey];
  if (normalized === null) {
    delete state.limits.days[dayKey];
  } else {
    state.limits.days[dayKey] = normalized;
  }

  try {
    await persistLimit(dayKey, "DAY", normalized);
  } catch (error) {
    if (previousValue === undefined) {
      delete state.limits.days[dayKey];
    } else {
      state.limits.days[dayKey] = previousValue;
    }
    setStatus(`Unable to save daily limit: ${error.message}`);
  }
}

async function setZoneLimit(dayKey, zone, rawValue) {
  if (!state.limits.zones) {
    state.limits.zones = {};
  }

  if (!state.limits.zones[dayKey]) {
    state.limits.zones[dayKey] = {};
  }

  const normalized = normalizeLimitInput(rawValue);
  const previousValue = state.limits.zones[dayKey][zone];
  if (normalized === null) {
    delete state.limits.zones[dayKey][zone];
    if (!Object.keys(state.limits.zones[dayKey]).length) {
      delete state.limits.zones[dayKey];
    }
  } else {
    state.limits.zones[dayKey][zone] = normalized;
  }

  try {
    await persistLimit(dayKey, zone, normalized);
  } catch (error) {
    if (!state.limits.zones[dayKey]) {
      state.limits.zones[dayKey] = {};
    }
    if (previousValue === undefined) {
      delete state.limits.zones[dayKey][zone];
      if (!Object.keys(state.limits.zones[dayKey]).length) {
        delete state.limits.zones[dayKey];
      }
    } else {
      state.limits.zones[dayKey][zone] = previousValue;
    }
    setStatus(`Unable to save zone limit: ${error.message}`);
  }
}

function normalizeLimitInput(rawValue) {
  if (rawValue === "") {
    return null;
  }

  const parsed = Number.parseInt(rawValue, 10);
  if (!Number.isFinite(parsed) || parsed < 0) {
    return null;
  }

  return parsed;
}

function parseLimitValue(value) {
  return Number.isFinite(value) ? value : null;
}

function emptyLimits() {
  return { days: {}, zones: {} };
}

function normalizeLimitsPayload(payload) {
  const normalized = emptyLimits();
  const dayEntries = Object.entries(payload.days || {});
  const zoneEntries = Object.entries(payload.zones || {});

  dayEntries.forEach(([rawDateKey, value]) => {
    const dateKey = normalizeDateKey(rawDateKey);
    if (!dateKey) {
      return;
    }
    const parsedValue = Number(value);
    if (Number.isFinite(parsedValue)) {
      normalized.days[dateKey] = parsedValue;
    }
  });

  zoneEntries.forEach(([rawDateKey, zones]) => {
    const dateKey = normalizeDateKey(rawDateKey);
    if (!dateKey || !zones || typeof zones !== "object") {
      return;
    }

    Object.entries(zones).forEach(([zone, value]) => {
      const parsedValue = Number(value);
      if (!Number.isFinite(parsedValue)) {
        return;
      }
      if (!normalized.zones[dateKey]) {
        normalized.zones[dateKey] = {};
      }
      normalized.zones[dateKey][zone] = parsedValue;
    });
  });

  return normalized;
}

function normalizeDateKey(rawDateKey) {
  if (!rawDateKey) {
    return null;
  }

  const trimmed = String(rawDateKey).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
    return trimmed;
  }

  const parsed = new Date(trimmed);
  if (Number.isNaN(parsed.getTime())) {
    return null;
  }

  return formatDateKey(parsed);
}

async function persistLimit(dateKey, zone, limit) {
  const response = await fetch(LIMITS_API_URL, {
    method: "POST",
    headers: {
      "Content-Type": "text/plain;charset=utf-8",
    },
    body: JSON.stringify({
      dateKey,
      zone,
      limit,
    }),
  });

  if (!response.ok) {
    throw new Error(`save failed with status ${response.status}`);
  }

  const payload = await response.json();
  if (!payload.ok) {
    throw new Error(payload.error || "unknown save error");
  }
}

function setStatus(message) {
  statusBanner.textContent = message;
  statusBanner.classList.toggle("is-hidden", !message);
}

function findHeaderIndex(headers, options) {
  const normalizedHeaders = headers.map(normalizeHeader);
  return options
    .map((option) => normalizedHeaders.indexOf(normalizeHeader(option)))
    .find((index) => index >= 0);
}

function readCell(row, index, raw = false) {
  if (!Array.isArray(row) || index === undefined || index < 0) {
    return raw ? null : "";
  }

  const cell = row[index];
  if (raw) {
    return cell || null;
  }

  if (!cell) {
    return "";
  }

  if (typeof cell.f === "string" && cell.f.trim()) {
    return cell.f.trim();
  }

  if (cell.v === null || cell.v === undefined) {
    return "";
  }

  return String(cell.v).trim();
}

function parseGvizDate(cell) {
  if (!cell) {
    return null;
  }

  if (typeof cell.v === "string") {
    const match = cell.v.match(/^Date\((\d+),(\d+),(\d+)\)$/);
    if (match) {
      const [, year, month, day] = match;
      return new Date(Number(year), Number(month), Number(day));
    }
  }

  if (typeof cell.f === "string") {
    const parsed = new Date(cell.f);
    if (!Number.isNaN(parsed.getTime())) {
      return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
    }
  }

  return null;
}

function normalizeHeader(value) {
  return value.trim().toLowerCase().replace(/\s+/g, " ");
}

function normalizeZone(value) {
  const trimmed = String(value || "").trim().toUpperCase();
  const match = trimmed.match(/^Z([1-5])/);
  if (match) {
    return `Z${match[1]}`;
  }
  return trimmed;
}

function normalizeSchoolName(value) {
  return String(value || "")
    .replace(/\s+Z[1-6][A-Z]*$/i, "")
    .trim();
}

function compareZoneEntries(left, right) {
  return compareZones(left.category, right.category);
}

function compareZones(left, right) {
  const leftMatch = String(left).match(/^Z(\d+)$/);
  const rightMatch = String(right).match(/^Z(\d+)$/);

  if (leftMatch && rightMatch) {
    return Number(leftMatch[1]) - Number(rightMatch[1]);
  }

  if (leftMatch) {
    return -1;
  }

  if (rightMatch) {
    return 1;
  }

  return String(left).localeCompare(String(right));
}

function startOfMonth(date) {
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

function endOfMonth(date) {
  return new Date(date.getFullYear(), date.getMonth() + 1, 0);
}

function startOfWeek(date) {
  return addDays(date, -date.getDay());
}

function endOfWeek(date) {
  return addDays(date, 6 - date.getDay());
}

function addDays(date, amount) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate() + amount);
}

function addMonths(date, amount) {
  return new Date(date.getFullYear(), date.getMonth() + amount, 1);
}

function formatDateKey(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
