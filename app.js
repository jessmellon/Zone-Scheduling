const SHEET_GVIZ_URL =
  "https://docs.google.com/spreadsheets/d/1_KPdGkIe-tQrKEFxqAfJL87VvX73aEPuwVW5G_b4zOI/gviz/tq?tqx=out:json&sheet=Copy%20of%20Dates";
const LIMITS_API_URL =
  "https://script.google.com/macros/s/AKfycbw9uRUe811I05XrPwi1r1nDU045a9WlCUgLYL_WJqeHTaFb5zbcxbKX3veshhZBkz7x/exec";
const ALL_ZONES = ["Z1", "Z2", "Z3", "Z4", "Z5"];

const state = {
  events: [],
  currentMonth: null,
  selectedCategory: null,
  selectedStaffingZone: null,
  selectedEventId: null,
  focusedDayKey: null,
  searchTerm: "",
  viewMode: "staffing",
  categoryColors: new Map(),
  limits: emptyLimits(),
  notes: {},
  noteMeta: {},
  noteHistory: {},
  dateHistory: {},
  selectedNoteDateKey: null,
  saveSequence: 0,
  latestSaveByKey: {},
};

const monthLabel = document.querySelector("#month-label");
const calendarGrid = document.querySelector("#calendar-grid");
const filterList = document.querySelector("#filter-list");
const selectionDetails = document.querySelector("#selection-details");
const statusBanner = document.querySelector("#status-banner");
const lastUpdated = document.querySelector("#last-updated");
const searchInput = document.querySelector("#school-search");
const searchResults = document.querySelector("#search-results");
const dayNoteInput = document.querySelector("#day-note-input");
const noteDateLabel = document.querySelector("#note-date-label");
const currentNoteMeta = document.querySelector("#current-note-meta");
const dayNotesPanelAnchor = document.querySelector("#day-notes-panel-anchor");
const dayNotesPanel = document.querySelector("#day-notes-panel");
const saveNoteButton = document.querySelector("#save-note");
const noteHistory = document.querySelector("#note-history");

document.querySelector("#save-note").addEventListener("click", () => {
  saveSelectedNote();
});

document.querySelector("#clear-note").addEventListener("click", () => {
  clearSelectedNote();
});

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
  state.selectedStaffingZone = null;
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

dayNoteInput.addEventListener("input", () => {
  updateNoteSaveButtonState();
});

loadCalendar();

async function loadCalendar() {
  setStatus("Loading live data from Google Sheets...");

  try {
    const [sheetData, sharedData] = await Promise.all([loadGvizData(), loadSharedData()]);
    const events = buildEvents(sheetData);

    if (!events.length) {
      throw new Error("No dated rows were found in the Dates tab.");
    }

    state.events = events.sort((left, right) => left.date - right.date);
    state.limits = sharedData.limits;
    state.notes = sharedData.notes;
    state.noteMeta = sharedData.noteMeta;
    state.noteHistory = sharedData.noteHistory;
    state.dateHistory = sharedData.dateHistory;
    state.currentMonth = startOfMonth(new Date());
    state.categoryColors = buildCategoryColors(state.events);
    lastUpdated.textContent = `Loaded ${state.events.length} calendar entries`;
  } catch (error) {
    console.error(error);
    setStatus(`Unable to load the Google Sheet: ${error.message}`);
    lastUpdated.textContent = "Sheet load failed";
    return;
  }

  try {
    render();
  } catch (error) {
    console.error(error);
    setStatus(`Unable to render the page: ${error.message}`);
    lastUpdated.textContent = "Render failed";
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

async function loadSharedData() {
  const response = await fetch(LIMITS_API_URL, {
    method: "GET",
  });

  if (!response.ok) {
    throw new Error(`Limits request failed with status ${response.status}`);
  }

  const payload = await response.json();
  return normalizeSharedPayload(payload);
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
  renderFilters();
  renderCalendar();
  renderDetails();
  renderNotesPanel();
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
        state.selectedStaffingZone ? ` for ${state.selectedStaffingZone}` : ""
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

function renderFilters() {
  const counts =
    state.viewMode === "staffing"
      ? sumPhotographersByCategory(state.events)
      : countByCategory(state.events);

  filterList.innerHTML = "";
  counts.forEach(({ category, value }) => {
    const isActive =
      state.viewMode === "staffing"
        ? state.selectedStaffingZone === category
        : state.selectedCategory === category;
    const button = document.createElement("button");
    button.type = "button";
    button.className = `filter-chip${isActive ? " active" : ""}`;
    button.innerHTML = `
      <span class="dot" style="background:${getCategoryColor(category)}"></span>
      <span>${escapeHtml(category)}</span>
      <span class="summary-count">${value}</span>
    `;
    button.addEventListener("click", () => {
      if (state.viewMode === "staffing") {
        state.selectedStaffingZone =
          state.selectedStaffingZone === category ? null : category;
      } else {
        state.selectedCategory =
          state.selectedCategory === category ? null : category;
      }
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
  const allEvents = state.events;
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
    const allDayEvents = allEvents.filter(
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

    if (dayKey === state.focusedDayKey) {
      cell.classList.add("is-focused");
      cell.dataset.focusedDay = dayKey;
    }

    if (isWeekend) {
      cell.classList.add("is-weekend");
    }

    if (isWeekend && allDayEvents.length > 0) {
      cell.classList.add("is-weekend-booked");
    }

    if (state.viewMode === "staffing") {
      applyCapacityClass(
        cell,
        getCapacityStatus(
          getTotalPhotographers(allDayEvents),
          getDayLimit(dayKey)
        )
      );
    }

    const groupsMarkup =
      state.viewMode === "staffing"
        ? renderStaffingGroups(dayKey, allDayEvents)
        : renderEventGroups(dayEvents);
    const hasNote = hasDayNote(dayKey);

    cell.innerHTML = `
      <div class="day-heading">
        <button
          class="day-note-button"
          type="button"
          data-note-day="${dayKey}"
          aria-label="Open note for ${escapeHtml(dayKey)}"
        >
          <div class="day-number${hasNote ? " has-note" : ""}">${cursor.getDate()}</div>
        </button>
        <div class="day-name">${escapeHtml(
          cursor.toLocaleDateString(undefined, { weekday: "short" })
        )}</div>
      </div>
      <div class="day-groups">${groupsMarkup}</div>
    `;

    cell.querySelectorAll("[data-note-day]").forEach((button) => {
      button.addEventListener("click", () => {
        state.selectedNoteDateKey = button.dataset.noteDay;
        renderNotesPanel();
        scrollNotesPanelIntoView();
      });
    });

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
        input.addEventListener("change", (event) => {
          setDayLimit(dayKey, event.target.value);
        });
      });

      cell.querySelectorAll("[data-zone-limit]").forEach((input) => {
        input.addEventListener("change", (event) => {
          setZoneLimit(dayKey, event.target.dataset.zone, event.target.value);
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
    <div class="planner-card planner-total-card ${getCapacityClassName(dayStatus)}">
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
      <div class="planner-row">
        <span class="planner-title">Day total</span>
        <span class="planner-total-number">${totalPhotographers}</span>
      </div>
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
            <span class="event-title">${value} Photog${
              value === 1 ? "" : "s"
            }</span>
            <span class="event-meta">${totalEvents} school${
              totalEvents === 1 ? "" : "s"
            }</span>
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
        ${renderDetailItem(
          "Filter",
          state.selectedStaffingZone || "All zones"
        )}
      </div>
      ${renderColorLegend()}
    `;
    return;
  }

  const event = state.events.find((item) => item.id === state.selectedEventId);

  if (!event) {
    selectionDetails.className = "details-card";
    selectionDetails.innerHTML = `
      <h3>Details</h3>
      <p>Select an event on the calendar to inspect it.</p>
      ${renderColorLegend()}
    `;
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

function renderNotesPanel() {
  if (!state.selectedNoteDateKey) {
    noteDateLabel.textContent = "Select a day to add a note.";
    currentNoteMeta.textContent = "";
    dayNoteInput.value = "";
    dayNoteInput.disabled = true;
    dayNotesPanel.classList.remove("has-note");
    noteHistory.innerHTML = '<div class="note-history-empty">Select a day to view note history.</div>';
    updateNoteSaveButtonState();
    return;
  }

  const selectedDate = parseDateKey(state.selectedNoteDateKey);
  noteDateLabel.textContent = selectedDate.toLocaleDateString(undefined, {
    weekday: "long",
    month: "long",
    day: "numeric",
    year: "numeric",
  });
  dayNoteInput.disabled = false;
  const noteValue = state.notes[state.selectedNoteDateKey] || "";
  dayNoteInput.value = noteValue;
  dayNotesPanel.classList.toggle("has-note", Boolean(noteValue.trim()));
  currentNoteMeta.textContent = buildCurrentNoteMeta(state.selectedNoteDateKey);
  renderNoteHistory(state.selectedNoteDateKey);
  updateNoteSaveButtonState();
}

function renderNoteHistory(dateKey) {
  const entries = state.noteHistory[dateKey] || [];

  if (!entries.length) {
    noteHistory.innerHTML = `
      <p class="note-history-title">History</p>
      <div class="note-history-empty">No previous notes for this day.</div>
    `;
    return;
  }

  const itemsMarkup = entries
    .map((entry) => {
      return `
        <div class="note-history-item">
          <span class="note-history-meta">${escapeHtml(formatHistoryMeta(entry))}</span>
          <p class="note-history-text">${escapeHtml(entry.note || "")}</p>
        </div>
      `;
    })
    .join("");

  noteHistory.innerHTML = `
    <p class="note-history-title">History</p>
    ${itemsMarkup}
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
      const historyMarkup = renderSchoolDateHistory(schoolName);
      const itemsMarkup = events
        .sort((left, right) => left.date - right.date)
        .map((event) => {
          return `
            <button
              class="search-result-item"
              type="button"
              data-search-event="${escapeHtml(event.id)}"
            >
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
            </button>
          `;
        })
        .join("");

      return `
        <div class="search-result-card">
          <h3>${escapeHtml(schoolName)}</h3>
          <div class="search-result-list">${itemsMarkup}</div>
          ${historyMarkup}
        </div>
      `;
    })
    .join("");

  searchResults.innerHTML = cardsMarkup;
  searchResults.querySelectorAll("[data-search-event]").forEach((button) => {
    button.addEventListener("click", () => {
      const event = state.events.find((item) => item.id === button.dataset.searchEvent);
      if (!event) {
        return;
      }
      jumpToEventInStaffingView(event);
    });
  });
}

function renderDetailItem(label, value) {
  return `
    <div class="detail-item">
      <strong>${escapeHtml(label)}</strong>
      <span>${escapeHtml(value || "Not provided")}</span>
    </div>
  `;
}

function renderColorLegend() {
  return `
    <div class="details-legend">
      <div class="legend-item">
        <span class="legend-swatch warning"></span>
        <span>Within 2 below the limit</span>
      </div>
      <div class="legend-item">
        <span class="legend-swatch danger"></span>
        <span>At the limit</span>
      </div>
      <div class="legend-item">
        <span class="legend-swatch critical"></span>
        <span>Over the limit</span>
      </div>
    </div>
  `;
}

function getVisibleEvents() {
  const activeCategory =
    state.viewMode === "staffing" ? state.selectedStaffingZone : state.selectedCategory;

  if (!activeCategory) {
    return state.events;
  }

  return state.events.filter((event) => event.category === activeCategory);
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
  if (state.viewMode === "staffing" && state.selectedStaffingZone) {
    return [state.selectedStaffingZone];
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
  if (actual === 0 && limit === 0) {
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

function hasDayNote(dayKey) {
  return Boolean((state.notes[dayKey] || "").trim());
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
  const saveKey = getLimitSaveKey(dayKey, "DAY");
  const saveToken = registerSaveToken(saveKey);
  if (normalized === null) {
    delete state.limits.days[dayKey];
  } else {
    state.limits.days[dayKey] = normalized;
  }
  render();

  try {
    await persistLimit(dayKey, "DAY", normalized);
  } catch (error) {
    if (!isLatestSaveToken(saveKey, saveToken)) {
      return;
    }
    if (previousValue === undefined) {
      delete state.limits.days[dayKey];
    } else {
      state.limits.days[dayKey] = previousValue;
    }
    setStatus(`Unable to save daily limit: ${error.message}`);
    render();
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
  const saveKey = getLimitSaveKey(dayKey, zone);
  const saveToken = registerSaveToken(saveKey);
  if (normalized === null) {
    delete state.limits.zones[dayKey][zone];
    if (!Object.keys(state.limits.zones[dayKey]).length) {
      delete state.limits.zones[dayKey];
    }
  } else {
    state.limits.zones[dayKey][zone] = normalized;
  }
  render();

  try {
    await persistLimit(dayKey, zone, normalized);
  } catch (error) {
    if (!isLatestSaveToken(saveKey, saveToken)) {
      return;
    }
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
    render();
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

async function saveSelectedNote() {
  if (!state.selectedNoteDateKey) {
    return;
  }

  const value = dayNoteInput.value.trim();
  const previousValue = state.notes[state.selectedNoteDateKey];
  const previousMeta = state.noteMeta[state.selectedNoteDateKey]
    ? { ...state.noteMeta[state.selectedNoteDateKey] }
    : undefined;
  const previousHistory = cloneHistoryEntries(state.noteHistory[state.selectedNoteDateKey]);
  const saveKey = getLimitSaveKey(state.selectedNoteDateKey, "NOTE");
  const saveToken = registerSaveToken(saveKey);
  updateLocalNoteState(state.selectedNoteDateKey, previousValue, value);

  render();

  try {
    await persistNote(state.selectedNoteDateKey, value);
  } catch (error) {
    if (!isLatestSaveToken(saveKey, saveToken)) {
      return;
    }
    restoreLocalNoteState(
      state.selectedNoteDateKey,
      previousValue,
      previousMeta,
      previousHistory
    );
    setStatus(`Unable to save note: ${error.message}`);
    render();
  }
}

function clearSelectedNote() {
  if (!state.selectedNoteDateKey) {
    return;
  }

  dayNoteInput.value = "";
  saveSelectedNote();
}

function updateNoteSaveButtonState() {
  if (!state.selectedNoteDateKey || dayNoteInput.disabled) {
    saveNoteButton.disabled = true;
    return;
  }

  const currentValue = dayNoteInput.value.trim();
  const savedValue = (state.notes[state.selectedNoteDateKey] || "").trim();
  saveNoteButton.disabled = currentValue === savedValue;
}

function jumpToEventInStaffingView(event) {
  state.viewMode = "staffing";
  state.selectedCategory = null;
  state.selectedStaffingZone = null;
  state.selectedEventId = null;
  state.currentMonth = startOfMonth(event.date);
  state.focusedDayKey = formatDateKey(event.date);
  render();
  window.requestAnimationFrame(() => {
    const focusedDay = document.querySelector("[data-focused-day]");
    if (focusedDay) {
      focusedDay.scrollIntoView({
        behavior: "smooth",
        block: "center",
      });
    }
  });
}

function scrollNotesPanelIntoView() {
  if (!dayNotesPanelAnchor) {
    return;
  }

  dayNotesPanelAnchor.scrollIntoView({
    behavior: "smooth",
    block: "start",
  });
}

function getLimitSaveKey(dateKey, zone) {
  return `${dateKey}::${zone}`;
}

function registerSaveToken(saveKey) {
  state.saveSequence += 1;
  state.latestSaveByKey[saveKey] = state.saveSequence;
  return state.saveSequence;
}

function isLatestSaveToken(saveKey, saveToken) {
  return state.latestSaveByKey[saveKey] === saveToken;
}

function normalizeSharedPayload(payload) {
  const normalized = {
    limits: emptyLimits(),
    notes: {},
    noteMeta: {},
    noteHistory: {},
    dateHistory: {},
  };
  const dayEntries = Object.entries(payload.days || {});
  const zoneEntries = Object.entries(payload.zones || {});
  const noteEntries = Object.entries(payload.notes || {});
  const noteMetaEntries = Object.entries(payload.noteMeta || {});
  const noteHistoryEntries = Object.entries(payload.noteHistory || {});
  const dateHistoryEntries = normalizeDateHistoryEntries(payload.dateHistory);

  dayEntries.forEach(([rawDateKey, value]) => {
    const dateKey = normalizeDateKey(rawDateKey);
    if (!dateKey) {
      return;
    }
    const parsedValue = Number(value);
    if (Number.isFinite(parsedValue)) {
      normalized.limits.days[dateKey] = parsedValue;
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
      if (!normalized.limits.zones[dateKey]) {
        normalized.limits.zones[dateKey] = {};
      }
      normalized.limits.zones[dateKey][zone] = parsedValue;
    });
  });

  noteEntries.forEach(([rawDateKey, note]) => {
    const dateKey = normalizeDateKey(rawDateKey);
    if (!dateKey) {
      return;
    }
    const normalizedNote = String(note || "").trim();
    if (normalizedNote) {
      normalized.notes[dateKey] = normalizedNote;
    }
  });

  noteMetaEntries.forEach(([rawDateKey, meta]) => {
    const dateKey = normalizeDateKey(rawDateKey);
    if (!dateKey || !meta || typeof meta !== "object") {
      return;
    }

    normalized.noteMeta[dateKey] = {
      updatedAt: String(meta.updatedAt || "").trim(),
    };
  });

  noteHistoryEntries.forEach(([rawDateKey, entries]) => {
    const dateKey = normalizeDateKey(rawDateKey);
    if (!dateKey || !Array.isArray(entries)) {
      return;
    }

    normalized.noteHistory[dateKey] = entries
      .map((entry) => ({
        note: String(entry.note || "").trim(),
        status: String(entry.status || "").trim().toUpperCase(),
        updatedAt: String(entry.updatedAt || "").trim(),
      }))
      .filter((entry) => entry.note);
  });

  dateHistoryEntries.forEach(([schoolName, entries]) => {
    const normalizedSchool = normalizeSchoolName(schoolName);
    if (!normalizedSchool || !Array.isArray(entries)) {
      return;
    }

    normalized.dateHistory[normalizedSchool] = entries
      .map((entry) => normalizeDateHistoryEntry(entry))
      .filter(Boolean)
      .sort(compareDateHistoryEntries);
  });

  return normalized;
}

function normalizeDateHistoryEntries(rawDateHistory) {
  if (Array.isArray(rawDateHistory)) {
    const grouped = new Map();

    rawDateHistory.forEach((entry) => {
      const schoolName = normalizeSchoolName(entry?.school || entry?.School || "");
      if (!schoolName) {
        return;
      }

      if (!grouped.has(schoolName)) {
        grouped.set(schoolName, []);
      }

      grouped.get(schoolName).push(entry);
    });

    return [...grouped.entries()];
  }

  if (rawDateHistory && typeof rawDateHistory === "object") {
    return Object.entries(rawDateHistory);
  }

  return [];
}

function normalizeDateHistoryEntry(entry) {
  if (!entry || typeof entry !== "object") {
    return null;
  }

  const oldDate = normalizeDateKey(
    entry.oldEffectiveDate || entry.OldEffectiveDate || entry.oldDate || entry.old
  );
  const newDate = normalizeDateKey(
    entry.newEffectiveDate || entry.NewEffectiveDate || entry.newDate || entry.new
  );
  const changedAt = String(
    entry.timestamp || entry.Timestamp || entry.changedAt || entry.updatedAt || ""
  ).trim();
  const editedBy = String(entry.editedBy || entry.EditedBy || "").trim();
  const sourceColumn = String(
    entry.sourceColumn || entry.SourceColumn || ""
  ).trim();

  if (!oldDate && !newDate) {
    return null;
  }

  return {
    oldDate,
    newDate,
    changedAt,
    editedBy,
    sourceColumn,
  };
}

function compareDateHistoryEntries(left, right) {
  const leftTime = Date.parse(left.changedAt || "") || 0;
  const rightTime = Date.parse(right.changedAt || "") || 0;
  return rightTime - leftTime;
}

function renderSchoolDateHistory(schoolName) {
  const entries = state.dateHistory[normalizeSchoolName(schoolName)] || [];

  if (!entries.length) {
    return "";
  }

  const itemsMarkup = entries
    .map((entry) => {
      return `
        <div class="search-history-item">
          <span class="search-history-dates">${escapeHtml(
            formatDateHistoryRange(entry)
          )}</span>
          <span class="search-history-meta">${escapeHtml(
            formatDateHistoryMeta(entry)
          )}</span>
        </div>
      `;
    })
    .join("");

  return `
    <div class="search-history">
      <p class="search-history-title">Moved Dates</p>
      <div class="search-history-list">${itemsMarkup}</div>
    </div>
  `;
}

function formatDateHistoryRange(entry) {
  const oldLabel = entry.oldDate ? formatShortDate(parseDateKey(entry.oldDate)) : "No date";
  const newLabel = entry.newDate ? formatShortDate(parseDateKey(entry.newDate)) : "Cleared";
  return `${oldLabel} -> ${newLabel}`;
}

function formatDateHistoryMeta(entry) {
  const parts = [];

  if (entry.changedAt) {
    const parsed = new Date(entry.changedAt);
    if (!Number.isNaN(parsed.getTime())) {
      parts.push(
        parsed.toLocaleString(undefined, {
          month: "short",
          day: "numeric",
          year: "numeric",
          hour: "numeric",
          minute: "2-digit",
        })
      );
    } else {
      parts.push(entry.changedAt);
    }
  }

  if (entry.editedBy) {
    parts.push(entry.editedBy);
  }

  if (entry.sourceColumn) {
    parts.push(`from ${entry.sourceColumn}`);
  }

  return parts.join(" • ");
}

function formatShortDate(date) {
  return date.toLocaleDateString(undefined, {
    month: "short",
    day: "numeric",
    year: "numeric",
  });
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

function parseDateKey(dateKey) {
  const match = String(dateKey).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (match) {
    const [, year, month, day] = match;
    return new Date(Number(year), Number(month) - 1, Number(day));
  }

  const parsed = new Date(dateKey);
  return Number.isNaN(parsed.getTime()) ? new Date() : parsed;
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

async function persistNote(dateKey, note) {
  const response = await fetch(LIMITS_API_URL, {
    method: "POST",
    headers: {
      "Content-Type": "text/plain;charset=utf-8",
    },
    body: JSON.stringify({
      dateKey,
      zone: "NOTE",
      note: note || "",
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

function updateLocalNoteState(dateKey, previousValue, nextValue) {
  if (previousValue && previousValue !== nextValue) {
    appendNoteHistoryEntry(dateKey, {
      note: previousValue,
      status: "ARCHIVED",
      updatedAt: new Date().toISOString(),
    });
  }

  if (nextValue) {
    state.notes[dateKey] = nextValue;
    state.noteMeta[dateKey] = {
      updatedAt: new Date().toISOString(),
    };
  } else {
    delete state.notes[dateKey];
    delete state.noteMeta[dateKey];
  }
}

function restoreLocalNoteState(dateKey, previousValue, previousMeta, previousHistory) {
  if (previousValue === undefined) {
    delete state.notes[dateKey];
  } else {
    state.notes[dateKey] = previousValue;
  }

  if (previousMeta === undefined) {
    delete state.noteMeta[dateKey];
  } else {
    state.noteMeta[dateKey] = previousMeta;
  }

  if (previousHistory) {
    state.noteHistory[dateKey] = previousHistory;
  } else {
    delete state.noteHistory[dateKey];
  }
}

function appendNoteHistoryEntry(dateKey, entry) {
  if (!state.noteHistory[dateKey]) {
    state.noteHistory[dateKey] = [];
  }

  state.noteHistory[dateKey] = [entry, ...state.noteHistory[dateKey]];
}

function cloneHistoryEntries(entries) {
  if (!Array.isArray(entries) || !entries.length) {
    return null;
  }

  return entries.map((entry) => ({ ...entry }));
}

function formatHistoryMeta(entry) {
  const parts = [];
  if (entry.status) {
    parts.push(entry.status === "ARCHIVED" ? "Cleared" : entry.status);
  }
  if (entry.updatedAt) {
    const parsed = new Date(entry.updatedAt);
    if (!Number.isNaN(parsed.getTime())) {
      parts.push(
        parsed.toLocaleString(undefined, {
          month: "short",
          day: "numeric",
          year: "numeric",
          hour: "numeric",
          minute: "2-digit",
        })
      );
    }
  }
  return parts.join(" • ") || "Previous note";
}

function buildCurrentNoteMeta(dateKey) {
  const noteValue = (state.notes[dateKey] || "").trim();
  if (!noteValue) {
    return "";
  }

  const updatedAt = state.noteMeta[dateKey]?.updatedAt || "";
  if (!updatedAt) {
    return "Current note saved";
  }

  const parsed = new Date(updatedAt);
  if (Number.isNaN(parsed.getTime())) {
    return "Current note saved";
  }

  return `Current note saved ${parsed.toLocaleString(undefined, {
    month: "short",
    day: "numeric",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
  })}`;
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
