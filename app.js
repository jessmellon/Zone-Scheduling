const SHEET_GVIZ_URL =
  "https://docs.google.com/spreadsheets/d/1_KPdGkIe-tQrKEFxqAfJL87VvX73aEPuwVW5G_b4zOI/gviz/tq?tqx=out:json&sheet=Copy%20of%20Dates";
const LIMITS_API_URL =
  "https://script.google.com/macros/s/AKfycbzR4cRjr4WDOkV8UQ2g2HuWlxdeyOfkRUZ1mw6sRIK8hlsmZ5CcomH-LWLtGuMe9tHf/exec";
const ALL_ZONES = ["Z1", "Z2", "Z3", "Z4", "Z5"];
const SITE_PASSWORD = "VOS4437";
const PASSWORD_STORAGE_KEY = "zone_scheduling_unlocked";

const state = {
  events: [],
  currentMonth: null,
  selectedCategory: null,
  selectedStaffingZone: null,
  selectedStars: null,
  selectedType: null,
  selectedSent: null,
  selectedConfirmed: null,
  layoutMode: "standard",
  selectedStaffingDayKey: null,
  selectedStaffingZoneKey: null,
  selectedEventsDayKey: null,
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
  confirmedOverrides: {},
  selectedNoteDateKey: null,
  selectedNoteZone: null,
  saveSequence: 0,
  latestSaveByKey: {},
};

const monthLabel = document.querySelector("#month-label");
const calendarGrid = document.querySelector("#calendar-grid");
const filterList = document.querySelector("#filter-list");
const activeFilters = document.querySelector("#active-filters");
const layoutRoot = document.querySelector(".layout");
const weekdayRow = document.querySelector(".weekday-row");
const starsFilterList = document.querySelector("#stars-filter-list");
const typeFilterList = document.querySelector("#type-filter-list");
const sentFilterList = document.querySelector("#sent-filter-list");
const confirmedFilterList = document.querySelector("#confirmed-filter-list");
const selectionDetails = document.querySelector("#selection-details");
const detailsPanel = selectionDetails ? selectionDetails.closest(".panel") : null;
const statusBanner = document.querySelector("#status-banner");
const lastUpdated = document.querySelector("#last-updated");
const searchInput = document.querySelector("#school-search");
const searchResults = document.querySelector("#search-results");
const dayNoteInput = document.querySelector("#day-note-input");
const noteDateLabel = document.querySelector("#note-date-label");
const currentNoteMeta = document.querySelector("#current-note-meta");
const holidayNote = document.querySelector("#holiday-note");
const dayNotesPanelAnchor = document.querySelector("#day-notes-panel-anchor");
const dayNotesPanel = document.querySelector("#day-notes-panel");
const saveNoteButton = document.querySelector("#save-note");
const noteHistory = document.querySelector("#note-history");
const calendarPanel = document.querySelector(".calendar-panel");
const passwordGate = document.querySelector("#password-gate");
const passwordForm = document.querySelector("#password-form");
const passwordInput = document.querySelector("#password-input");
const passwordError = document.querySelector("#password-error");

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
  state.selectedStars = null;
  state.selectedType = null;
  state.selectedSent = null;
  state.selectedConfirmed = null;
  render();
});

document.querySelector("#events-view-button").addEventListener("click", () => {
  state.viewMode = "events";
  state.selectedEventId = null;
  state.selectedStaffingDayKey = null;
  render();
});

document.querySelector("#staffing-view-button").addEventListener("click", () => {
  state.viewMode = "staffing";
  state.selectedEventId = null;
  state.selectedEventsDayKey = null;
  render();
});

document.querySelector("#picture-days-view-button").addEventListener("click", () => {
  state.viewMode = "picture-days";
  state.selectedEventId = null;
  state.selectedStaffingDayKey = null;
  state.selectedStaffingZoneKey = null;
  render();
});

document.querySelector("#standard-layout-button").addEventListener("click", () => {
  state.layoutMode = "standard";
  render();
});

document.querySelector("#full-month-layout-button").addEventListener("click", () => {
  state.layoutMode = "full-month";
  render();
});

searchInput.addEventListener("input", (event) => {
  state.searchTerm = event.target.value.trim();
  renderSearchResults();
});

dayNoteInput.addEventListener("input", () => {
  updateNoteSaveButtonState();
});

passwordForm.addEventListener("submit", (event) => {
  event.preventDefault();
  unlockSite();
});

initializePasswordGate();

async function loadCalendar() {
  setStatus("Loading live data from Google Sheets...");

  try {
    const [sheetData, sharedData] = await Promise.all([loadGvizData(), loadSharedData()]);
    const events = applyConfirmedOverrides(buildEvents(sheetData));

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

function initializePasswordGate() {
  const alreadyUnlocked = window.localStorage.getItem(PASSWORD_STORAGE_KEY) === "true";

  if (alreadyUnlocked) {
    passwordGate.classList.add("is-hidden");
    loadCalendar();
    return;
  }

  passwordGate.classList.remove("is-hidden");
  passwordInput.focus();
}

function unlockSite() {
  const enteredPassword = passwordInput.value;

  if (enteredPassword !== SITE_PASSWORD) {
    passwordError.classList.remove("is-hidden");
    passwordInput.select();
    return;
  }

  passwordError.classList.add("is-hidden");
  window.localStorage.setItem(PASSWORD_STORAGE_KEY, "true");
  passwordGate.classList.add("is-hidden");
  loadCalendar();
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
    confirmed: findHeaderIndex(headers, ["Confirmed", "Roosted", "Added To Roosted", "Added to Roosted"]),
    rowNumber: findHeaderIndex(headers, ["RowNumber", "Source Row", "Row"]),
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
        confirmed: parseConfirmedCell(readCell(row.c, indexes.confirmed, true)),
        rowNumber: parseSourceRow(readCell(row.c, indexes.rowNumber, true)),
      };
    })
    .filter(Boolean);
}

function render() {
  if (!state.currentMonth || !state.events.length) {
    return;
  }

  renderViewSwitch();
  renderLayoutSwitch();
  renderLayoutMode();
  renderHeader();
  renderFilters();
  renderActiveFilters();
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
  document
    .querySelector("#picture-days-view-button")
    .classList.toggle("active", state.viewMode === "picture-days");
}

function renderLayoutSwitch() {
  document
    .querySelector("#standard-layout-button")
    .classList.toggle("active", state.layoutMode === "standard");
  document
    .querySelector("#full-month-layout-button")
    .classList.toggle("active", state.layoutMode === "full-month");
}

function renderLayoutMode() {
  const isFullMonthLayout = state.layoutMode === "full-month";
  layoutRoot?.classList.toggle("full-month-layout", isFullMonthLayout);
  calendarGrid.classList.toggle("full-month-layout", isFullMonthLayout);
  weekdayRow?.classList.toggle("is-hidden", isFullMonthLayout);
}

function renderHeader() {
  monthLabel.textContent = state.currentMonth.toLocaleDateString(undefined, {
    month: "long",
    year: "numeric",
  });
  setStatus("");
}

function renderFilters() {
  const attributeFilteredEvents = getAttributeFilteredEvents();
  const counts =
    isStaffingLikeView()
      ? sumPhotographersByCategory(attributeFilteredEvents)
      : countByCategory(attributeFilteredEvents);

  filterList.innerHTML = "";
  counts.forEach(({ category, value }) => {
    const isActive =
      isStaffingLikeView()
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

  renderAttributeFilters();
}

function renderAttributeFilters() {
  renderAttributeFilterList(starsFilterList, getUniqueAttributeValues("stars"), state.selectedStars, (value) => {
    state.selectedStars = state.selectedStars === value ? null : value;
    render();
  });

  renderAttributeFilterList(typeFilterList, getUniqueAttributeValues("type"), state.selectedType, (value) => {
    state.selectedType = state.selectedType === value ? null : value;
    render();
  });

  renderAttributeFilterList(sentFilterList, getUniqueSentValues(), state.selectedSent, (value) => {
    state.selectedSent = state.selectedSent === value ? null : value;
    render();
  });

  renderAttributeFilterList(
    confirmedFilterList,
    ["Confirmed", "Not confirmed"],
    state.selectedConfirmed,
    (value) => {
      state.selectedConfirmed = state.selectedConfirmed === value ? null : value;
      render();
    }
  );
}

function renderActiveFilters() {
  if (!activeFilters) {
    return;
  }

  const chips = [];
  const attributeFilterChips = [];
  const zoneFilter =
    isStaffingLikeView() ? state.selectedStaffingZone : state.selectedCategory;

  if (zoneFilter) {
    chips.push(`Zone: ${zoneFilter}`);
  }
  if (state.selectedStars) {
    const chip = `Stars: ${state.selectedStars}`;
    chips.push(chip);
    attributeFilterChips.push(chip);
  }
  if (state.selectedType) {
    const chip = `Type: ${state.selectedType}`;
    chips.push(chip);
    attributeFilterChips.push(chip);
  }
  if (state.selectedSent) {
    const chip = `Sent: ${state.selectedSent}`;
    chips.push(chip);
    attributeFilterChips.push(chip);
  }
  if (state.selectedConfirmed) {
    const chip = `Confirmed: ${state.selectedConfirmed}`;
    chips.push(chip);
    attributeFilterChips.push(chip);
  }

  if (!chips.length) {
    activeFilters.innerHTML = "";
    activeFilters.classList.add("is-hidden");
    if (calendarPanel) {
      calendarPanel.classList.remove("is-filtered");
    }
    return;
  }

  activeFilters.innerHTML = `
    <span class="active-filters-label">Filtered View</span>
    ${chips
      .map((chip) => `<span class="active-filter-chip">${escapeHtml(chip)}</span>`)
      .join("")}
  `;
  activeFilters.classList.remove("is-hidden");
  if (calendarPanel) {
    calendarPanel.classList.toggle("is-filtered", attributeFilterChips.length > 0);
  }
}

function renderAttributeFilterList(container, values, activeValue, onSelect) {
  if (!container) {
    return;
  }

  container.innerHTML = "";
  values.forEach((value) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `attribute-chip${activeValue === value ? " active" : ""}`;
    button.textContent = value;
    button.addEventListener("click", () => onSelect(value));
    container.appendChild(button);
  });
}

function renderCalendar() {
  const monthStart = startOfMonth(state.currentMonth);
  const monthEnd = endOfMonth(state.currentMonth);
  const isFullMonthLayout = state.layoutMode === "full-month";
  const gridStart = startOfWeek(monthStart);
  const gridEnd = endOfWeek(monthEnd);
  const rangeStart = isFullMonthLayout ? monthStart : gridStart;
  const rangeEnd = isFullMonthLayout ? monthEnd : gridEnd;
  const visibleEvents = getVisibleEvents();
  const allEvents = getAttributeFilteredEvents();
  const todayKey = formatDateKey(new Date());

  calendarGrid.innerHTML = "";
  calendarGrid.style.setProperty("--month-day-count", String(monthEnd.getDate()));

  for (
    let cursor = new Date(rangeStart);
    cursor <= rangeEnd;
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

    if (isStaffingLikeView() && dayKey === state.selectedStaffingDayKey) {
      cell.classList.add("is-selected");
    }

    if (state.viewMode === "events" && dayKey === state.selectedEventsDayKey) {
      cell.classList.add("is-selected");
    }

    if (isWeekend) {
      cell.classList.add("is-weekend");
    }

    if (isWeekend && allDayEvents.length > 0) {
      cell.classList.add("is-weekend-booked");
    }

    if (isStaffingLikeView()) {
      applyCapacityClass(
        cell,
        getCapacityStatus(
          getTotalPhotographers(allDayEvents),
          getDayLimit(dayKey)
        )
      );
    }

    const groupsMarkup =
      isStaffingLikeView()
        ? renderStaffingGroups(dayKey, allDayEvents)
        : state.viewMode === "picture-days"
          ? renderPictureDayGroups(dayEvents)
          : renderEventGroups(dayEvents);
    const hasNote = hasDayNote(dayKey);
    const holidayNames = getHolidayNames(dayKey);

    if (hasNote) {
      cell.classList.add("has-note");
    }

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
      ${
        holidayNames.length
          ? `<div class="holiday-chip">${escapeHtml(holidayNames.join(" • "))}</div>`
          : ""
      }
      <div class="day-groups">${groupsMarkup}</div>
    `;

    cell.querySelectorAll("[data-note-day]").forEach((button) => {
      button.addEventListener("click", () => {
        state.selectedNoteDateKey = button.dataset.noteDay;
        state.selectedNoteZone = null;
        renderNotesPanel();
        scrollNotesPanelIntoView();
      });
    });

    if (!isStaffingLikeView()) {
      cell.addEventListener("click", (event) => {
        if (event.target.closest("[data-note-day]") || event.target.closest("[data-event-id]")) {
          return;
        }

        state.selectedEventsDayKey = dayKey;
        state.selectedEventId = null;
        renderDetails();
        renderCalendar();
        scrollDetailsIntoView();
      });

      cell.querySelectorAll("[data-event-id]").forEach((button) => {
        button.addEventListener("click", () => {
          state.selectedEventId = button.dataset.eventId;
          state.selectedEventsDayKey = null;
          renderDetails();
          renderCalendar();
          scrollDetailsIntoView();
        });
      });
    } else {
      cell.addEventListener("click", (event) => {
        if (
          event.target.closest("[data-note-day]") ||
          event.target.closest("[data-zone-note]") ||
          event.target.closest("[data-zone-card]") ||
          event.target.closest("[data-day-limit]") ||
          event.target.closest("[data-zone-limit]")
        ) {
          return;
        }

        state.selectedStaffingDayKey = dayKey;
        state.selectedStaffingZoneKey = null;
        renderDetails();
        renderCalendar();
      });

      cell.querySelectorAll("[data-zone-note]").forEach((button) => {
        button.addEventListener("click", (event) => {
          event.stopPropagation();
          state.selectedNoteDateKey = dayKey;
          state.selectedNoteZone = button.dataset.zoneNote;
          renderNotesPanel();
          scrollNotesPanelIntoView();
        });
      });

      cell.querySelectorAll("[data-zone-card]").forEach((button) => {
        button.addEventListener("click", () => {
          state.selectedStaffingDayKey = dayKey;
          state.selectedStaffingZoneKey = button.dataset.zoneCard;
          renderDetails();
          renderCalendar();
          scrollDetailsIntoView();
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

function renderPictureDayGroups(dayEvents) {
  return dayEvents
    .slice()
    .sort((left, right) => {
      const rowCompare = (left.rowNumber || Number.MAX_SAFE_INTEGER) - (right.rowNumber || Number.MAX_SAFE_INTEGER);
      if (rowCompare !== 0) {
        return rowCompare;
      }
      return left.title.localeCompare(right.title);
    })
    .map((event) => {
      return `
        <button
          class="event-pill picture-day-pill${
            state.selectedEventId === event.id ? " is-selected" : ""
          }"
          type="button"
          data-event-id="${escapeHtml(event.id)}"
        >
          <span class="event-title">${escapeHtml(event.schoolName)}</span>
          <span class="event-meta">${escapeHtml(
            [
              event.stars,
              event.type,
              event.sent,
              isConfirmedEvent(event) ? "Confirmed" : "Not confirmed",
              event.category,
              `${getPhotographerCount(event)} Photographer${
                getPhotographerCount(event) === 1 ? "" : "s"
              }`,
            ]
              .filter(Boolean)
              .join(" • ")
          )}</span>
        </button>
      `;
    })
    .join("");
}

function renderDetailSchoolLink(schoolName, metaText) {
  return `
    <button
      class="detail-school-item detail-school-link"
      type="button"
      data-detail-school="${escapeHtml(schoolName)}"
    >
      <strong>${escapeHtml(schoolName)}</strong>
      <span>${escapeHtml(metaText)}</span>
    </button>
  `;
}

function renderStaffingGroups(dayKey, dayEvents) {
  const totalPhotographers = getTotalPhotographers(dayEvents);
  const dayLimit = getDayLimit(dayKey);
  const dayStatus = getCapacityStatus(totalPhotographers, dayLimit);
  const displayedZones = getDisplayedZones();

  const plannerMarkup = `
    <div class="group-block">
      <div class="group-label">Day Total</div>
      <div class="planner-card planner-total-card ${getCapacityClassName(dayStatus)}">
        <span class="planner-total-number">${totalPhotographers}</span>
      </div>
    </div>
  `;

  const zoneTotals = buildZoneTotals(dayEvents);
  const zoneMarkup = displayedZones
    .map((category) => {
      const value = zoneTotals.get(category)?.photographers || 0;
      const zoneLimit = getZoneLimit(dayKey, category);
      const zoneStatus = getCapacityStatus(value, zoneLimit);
      return `
        <div class="group-block">
          <div class="group-heading">
            <div class="group-label">${escapeHtml(category)}</div>
            <button
              class="zone-note-chip${hasZoneNote(dayKey, category) ? " has-note" : ""}"
              type="button"
              data-zone-note="${escapeHtml(category)}"
            >
              Note
            </button>
          </div>
          <div
            class="event-pill summary-pill zone-summary-card ${getCapacityClassName(zoneStatus)}${
              state.selectedStaffingDayKey === dayKey &&
              state.selectedStaffingZoneKey === category
                ? " is-selected"
                : ""
            }"
            style="border-left-color:${getCategoryColor(category)}"
            data-zone-card="${escapeHtml(category)}"
          >
            <span class="summary-number">${value}</span>
          </div>
        </div>
      `;
    })
    .join("");

  return `${zoneMarkup}${plannerMarkup}`;
}

function renderDetails() {
  if (isStaffingLikeView()) {
    const detailEvents = getAttributeFilteredEvents();
    const selectedDayEvents = state.selectedStaffingDayKey
      ? detailEvents
          .filter((event) => formatDateKey(event.date) === state.selectedStaffingDayKey)
          .sort((left, right) => {
            const zoneCompare = compareZoneEntries(
              { category: left.category },
              { category: right.category }
            );
            if (zoneCompare !== 0) {
              return zoneCompare;
            }
            return left.title.localeCompare(right.title);
          })
      : [];

    if (state.selectedStaffingDayKey) {
      const selectedDate = parseDateKey(state.selectedStaffingDayKey);
      const selectedZone = state.selectedStaffingZoneKey;

      if (selectedZone) {
        const zoneEvents = selectedDayEvents.filter((event) => event.category === selectedZone);
        const zoneLimit = getZoneLimit(state.selectedStaffingDayKey, selectedZone);
        const zonePhotogs = getTotalPhotographers(zoneEvents);

        selectionDetails.className = "details-card";
        selectionDetails.innerHTML = `
          <h3>${escapeHtml(
            `${selectedDate.toLocaleDateString(undefined, {
              weekday: "long",
              month: "long",
              day: "numeric",
              year: "numeric",
            })} • ${selectedZone}`
          )}</h3>
          <div class="detail-grid">
            ${renderDetailItem("Schools", String(zoneEvents.length))}
            ${renderDetailItem("Photographers", String(zonePhotogs))}
          </div>
          <div class="detail-actions">
            <label class="limit-field detail-limit-field">
              <span class="limit-label">Limit</span>
              <input
                id="details-zone-limit"
                class="limit-input"
                type="number"
                min="0"
                step="1"
                value="${escapeHtml(zoneLimit ?? "")}"
              />
            </label>
          </div>
          <div class="detail-school-list">
            ${
              zoneEvents.length
                ? zoneEvents
                    .map((event) => {
                      return renderDetailSchoolLink(
                        event.schoolName,
                        [event.type, `${getPhotographerCount(event)} Photographer${
                          getPhotographerCount(event) === 1 ? "" : "s"
                        }`]
                          .filter(Boolean)
                          .join(" • ")
                      );
                    })
                    .join("")
                : '<p class="details-empty-copy">No schools scheduled in this zone.</p>'
            }
          </div>
          ${renderColorLegend()}
        `;

        const detailsZoneLimit = document.querySelector("#details-zone-limit");
        if (detailsZoneLimit) {
          detailsZoneLimit.addEventListener("change", (event) => {
            setZoneLimit(state.selectedStaffingDayKey, selectedZone, event.target.value);
          });
        }
        bindDetailSchoolLinks();
        return;
      }

      const itemsMarkup = selectedDayEvents.length
        ? selectedDayEvents
            .map((event) => {
              return renderDetailSchoolLink(
                event.schoolName,
                [event.category, event.type, `${getPhotographerCount(event)} Photographer${
                  getPhotographerCount(event) === 1 ? "" : "s"
                }`]
                  .filter(Boolean)
                  .join(" • ")
              );
            })
            .join("")
        : '<p class="details-empty-copy">No schools scheduled for this day.</p>';

      selectionDetails.className = "details-card";
      selectionDetails.innerHTML = `
        <h3>${escapeHtml(
          selectedDate.toLocaleDateString(undefined, {
            weekday: "long",
            month: "long",
            day: "numeric",
            year: "numeric",
          })
        )}</h3>
        <div class="detail-grid">
          ${renderDetailItem(
            "Schools",
            `${selectedDayEvents.length} scheduled school${
              selectedDayEvents.length === 1 ? "" : "s"
            }`
          )}
          ${renderDetailItem(
            "Photographers",
            `${getTotalPhotographers(selectedDayEvents)} total`
          )}
        </div>
        <div class="detail-actions">
          <label class="limit-field detail-limit-field">
            <span class="limit-label">Daily limit</span>
            <input
              id="details-day-limit"
              class="limit-input"
              type="number"
              min="0"
              step="1"
              value="${escapeHtml(getDayLimit(state.selectedStaffingDayKey) ?? "")}"
            />
          </label>
        </div>
        <div class="detail-school-list">${itemsMarkup}</div>
        ${renderColorLegend()}
      `;

      const detailsDayLimit = document.querySelector("#details-day-limit");
      if (detailsDayLimit) {
        detailsDayLimit.addEventListener("change", (event) => {
          setDayLimit(state.selectedStaffingDayKey, event.target.value);
        });
      }
      bindDetailSchoolLinks();
      return;
    }

    selectionDetails.className = "details-card";
    selectionDetails.innerHTML = `
      <h3>Zone staffing view</h3>
      <p>
        Each day shows the total number of photographers scheduled in each zone.
        Click a day to see the schools scheduled there.
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

  if (state.selectedEventsDayKey) {
    const selectedDate = parseDateKey(state.selectedEventsDayKey);
    const dayEvents = getAttributeFilteredEvents()
      .filter((event) => formatDateKey(event.date) === state.selectedEventsDayKey)
      .sort((left, right) => {
        const zoneCompare = compareZoneEntries(
          { category: left.category },
          { category: right.category }
        );
        if (zoneCompare !== 0) {
          return zoneCompare;
        }
        return left.title.localeCompare(right.title);
      });
    const historyEntries = getDateHistoryForDay(state.selectedEventsDayKey);

    selectionDetails.className = "details-card";
    selectionDetails.innerHTML = `
      <h3>${escapeHtml(
        selectedDate.toLocaleDateString(undefined, {
          weekday: "long",
          month: "long",
          day: "numeric",
          year: "numeric",
        })
      )}</h3>
      <p>Click a school for event details, or review what moved to or from this day below.</p>
      <div class="detail-school-list">
        ${
          dayEvents.length
            ? dayEvents
                .map((event) => {
                  return renderDetailSchoolLink(
                    event.schoolName,
                    [event.category, event.type].filter(Boolean).join(" • ")
                  );
                })
                .join("")
            : '<p class="details-empty-copy">No schools currently scheduled for this day.</p>'
        }
      </div>
      ${renderDayHistoryDetails(historyEntries)}
    `;
    bindDetailSchoolLinks();
    return;
  }

  const event = getAttributeFilteredEvents().find((item) => item.id === state.selectedEventId);

  if (!event) {
    selectionDetails.className = "details-card";
    selectionDetails.innerHTML = `
      <h3>Details</h3>
      <p>${
        state.viewMode === "picture-days"
          ? "Select a picture day entry on the calendar to inspect it."
          : "Select an event on the calendar to inspect it."
      }</p>
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
      ${renderDetailItem("Confirmed", isConfirmedEvent(event) ? "Yes" : "No")}
      ${renderDetailItem("Photographers", event.photographers)}
    </div>
    <div class="detail-actions">
      <button
        class="primary-button"
        type="button"
        data-confirm-toggle="${escapeHtml(event.id)}"
        ${event.rowNumber ? "" : "disabled"}
      >
        ${isConfirmedEvent(event) ? "Mark Not Confirmed" : "Mark Confirmed"}
      </button>
      ${
        event.rowNumber
          ? ""
          : '<p class="details-empty-copy">Add a RowNumber column from Table Main Tab before confirming from the site.</p>'
      }
    </div>
  `;

  const confirmButton = selectionDetails.querySelector("[data-confirm-toggle]");
  if (confirmButton) {
    confirmButton.addEventListener("click", () => {
      setConfirmedStatus(event.id, !isConfirmedEvent(event));
    });
  }
}

function bindDetailSchoolLinks() {
  selectionDetails.querySelectorAll("[data-detail-school]").forEach((button) => {
    button.addEventListener("click", () => {
      openSchoolSearch(button.dataset.detailSchool);
    });
  });
}

function openSchoolSearch(schoolName) {
  state.searchTerm = schoolName;
  if (searchInput) {
    searchInput.value = schoolName;
  }
  renderSearchResults();
  window.requestAnimationFrame(() => {
    searchInput?.scrollIntoView({ behavior: "smooth", block: "start" });
  });
}

function renderNotesPanel() {
  if (!state.selectedNoteDateKey) {
    noteDateLabel.textContent = "Select a day or zone to add a note.";
    currentNoteMeta.textContent = "";
    holidayNote.textContent = "";
    holidayNote.classList.add("is-hidden");
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
  if (state.selectedNoteZone) {
    noteDateLabel.textContent += ` • ${state.selectedNoteZone}`;
  }
  const holidayNames = getHolidayNames(state.selectedNoteDateKey);
  dayNoteInput.disabled = false;
  const noteKey = getSelectedNoteStorageKey();
  const noteValue = state.notes[noteKey] || "";
  dayNoteInput.value = noteValue;
  holidayNote.textContent = holidayNames.length ? `Holiday: ${holidayNames.join(" • ")}` : "";
  holidayNote.classList.toggle("is-hidden", !holidayNames.length);
  dayNotesPanel.classList.toggle("has-note", Boolean(noteValue.trim()) || holidayNames.length > 0);
  currentNoteMeta.textContent = buildCurrentNoteMeta(noteKey);
  renderNoteHistory(noteKey);
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
  const baseEvents = getAttributeFilteredEvents();
  const activeCategory =
    isStaffingLikeView() ? state.selectedStaffingZone : state.selectedCategory;

  if (!activeCategory) {
    return baseEvents;
  }

  return baseEvents.filter((event) => event.category === activeCategory);
}

function getAttributeFilteredEvents() {
  return state.events.filter((event) => {
    if (state.selectedStars && normalizeAttributeValue(event.stars) !== state.selectedStars) {
      return false;
    }

    if (state.selectedType && normalizeTypeFilterValue(event.type) !== state.selectedType) {
      return false;
    }

    if (state.selectedSent && normalizeSentValue(event.sent) !== state.selectedSent) {
      return false;
    }

    if (state.selectedConfirmed) {
      const isConfirmed = isConfirmedEvent(event);
      if (state.selectedConfirmed === "Confirmed" && !isConfirmed) {
        return false;
      }
      if (state.selectedConfirmed === "Not confirmed" && isConfirmed) {
        return false;
      }
    }

    return true;
  });
}

function getUniqueAttributeValues(field) {
  if (field === "type") {
    return [...new Set(
      state.events
        .map((event) => normalizeTypeFilterValue(event.type))
        .filter(Boolean)
    )].sort((left, right) => left.localeCompare(right, undefined, { numeric: true }));
  }

  return [...new Set(
    state.events
      .map((event) => normalizeAttributeValue(event[field]))
      .filter(Boolean)
  )].sort((left, right) => left.localeCompare(right, undefined, { numeric: true }));
}

function getUniqueSentValues() {
  return [...new Set(
    state.events
      .map((event) => normalizeSentValue(event.sent))
      .filter(Boolean)
  )].sort((left, right) => left.localeCompare(right));
}

function normalizeAttributeValue(value) {
  return String(value || "").trim();
}

function normalizeSentValue(value) {
  const normalized = String(value || "").trim();
  if (!normalized) {
    return "";
  }

  return normalized.toLowerCase() === "sent" ? "Sent" : normalized;
}

function normalizeTypeFilterValue(value) {
  const normalized = String(value || "").trim();
  if (!normalized) {
    return "";
  }

  const lower = normalized.toLowerCase();
  if (lower.includes("senior session") || lower.includes("seniors session")) {
    return "Seniors";
  }
  if (lower.includes("underclass")) {
    return "Underclass";
  }
  if (lower.includes("panoramic")) {
    return "Panoramic";
  }

  return normalized;
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
  if (isStaffingLikeView() && state.selectedStaffingZone) {
    return [state.selectedStaffingZone];
  }

  return ALL_ZONES;
}

function buildCategoryColors(events) {
  const palette = [
    "#2f2f2f",
    "#bcbcbc",
    "#565656",
    "#d6d6d6",
    "#7c7c7c",
    "#e7e7e7",
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

function isConfirmedEvent(event) {
  const value = event?.confirmed;
  if (value === true) {
    return true;
  }

  const normalized = String(value || "").trim().toLowerCase();
  return normalized === "true" || normalized === "yes" || normalized === "1";
}

function parseConfirmedCell(cell) {
  if (!cell) {
    return false;
  }

  if (typeof cell.v === "boolean") {
    return cell.v;
  }

  if (typeof cell.f === "string" && cell.f.trim()) {
    const formatted = cell.f.trim().toLowerCase();
    return formatted === "true" || formatted === "yes" || formatted === "1";
  }

  if (cell.v === null || cell.v === undefined) {
    return false;
  }

  const normalized = String(cell.v).trim().toLowerCase();
  return normalized === "true" || normalized === "yes" || normalized === "1";
}

function parseSourceRow(cell) {
  if (!cell) {
    return null;
  }

  const rawValue =
    cell.v !== null && cell.v !== undefined
      ? cell.v
      : typeof cell.f === "string"
        ? cell.f.trim()
        : null;

  const parsed = Number.parseInt(String(rawValue || "").trim(), 10);
  return Number.isFinite(parsed) && parsed > 1 ? parsed : null;
}

function applyConfirmedOverrides(events) {
  return events.map((event) => {
    if (!event.rowNumber) {
      return event;
    }

    const override = state.confirmedOverrides[event.rowNumber];
    if (override === undefined) {
      return event;
    }

    if (isConfirmedEvent(event) === override) {
      delete state.confirmedOverrides[event.rowNumber];
      return event;
    }

    return {
      ...event,
      confirmed: override,
    };
  });
}

function isStaffingLikeView() {
  return state.viewMode === "staffing";
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

function hasZoneNote(dayKey, zone) {
  return Boolean((state.notes[buildNoteStorageKey(dayKey, zone)] || "").trim());
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

async function setConfirmedStatus(eventId, confirmed) {
  const event = state.events.find((item) => item.id === eventId);
  if (!event || !event.rowNumber) {
    setStatus("Unable to update confirmation: missing source row number.");
    return;
  }

  const previousValue = event.confirmed;
  state.confirmedOverrides[event.rowNumber] = confirmed;
  event.confirmed = confirmed;
  render();

  try {
    await persistConfirmed(event.rowNumber, confirmed);
    setStatus("");
  } catch (error) {
    delete state.confirmedOverrides[event.rowNumber];
    event.confirmed = previousValue;
    setStatus(`Unable to update confirmation: ${error.message}`);
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

  const noteKey = getSelectedNoteStorageKey();
  const value = dayNoteInput.value.trim();
  const previousValue = state.notes[noteKey];
  const previousMeta = state.noteMeta[noteKey]
    ? { ...state.noteMeta[noteKey] }
    : undefined;
  const previousHistory = cloneHistoryEntries(state.noteHistory[noteKey]);
  const saveKey = getLimitSaveKey(noteKey, "NOTE");
  const saveToken = registerSaveToken(saveKey);
  updateLocalNoteState(noteKey, previousValue, value);

  render();

  try {
    await persistNote(noteKey, value);
  } catch (error) {
    if (!isLatestSaveToken(saveKey, saveToken)) {
      return;
    }
    restoreLocalNoteState(
      noteKey,
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

  const noteKey = getSelectedNoteStorageKey();
  const currentValue = dayNoteInput.value.trim();
  const savedValue = (state.notes[noteKey] || "").trim();
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

function scrollDetailsIntoView() {
  if (!detailsPanel) {
    return;
  }

  detailsPanel.scrollIntoView({
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
    const dateKey = normalizeNoteStorageKey(rawDateKey);
    if (!dateKey) {
      return;
    }
    const normalizedNote = String(note || "").trim();
    if (normalizedNote) {
      normalized.notes[dateKey] = normalizedNote;
    }
  });

  noteMetaEntries.forEach(([rawDateKey, meta]) => {
    const dateKey = normalizeNoteStorageKey(rawDateKey);
    if (!dateKey || !meta || typeof meta !== "object") {
      return;
    }

    normalized.noteMeta[dateKey] = {
      updatedAt: String(meta.updatedAt || "").trim(),
    };
  });

  noteHistoryEntries.forEach(([rawDateKey, entries]) => {
    const dateKey = normalizeNoteStorageKey(rawDateKey);
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

function getDateHistoryForDay(dateKey) {
  const entries = [];

  Object.entries(state.dateHistory).forEach(([schoolName, schoolEntries]) => {
    schoolEntries.forEach((entry) => {
      if (entry.oldDate === dateKey || entry.newDate === dateKey) {
        entries.push({
          ...entry,
          schoolName,
        });
      }
    });
  });

  return entries.sort(compareDateHistoryEntries);
}

function renderDayHistoryDetails(entries) {
  if (!entries.length) {
    return `
      <div class="detail-history">
        <h3>Change History</h3>
        <p class="details-empty-copy">No date changes recorded for this day.</p>
      </div>
    `;
  }

  const itemsMarkup = entries
    .map((entry) => {
      return `
        <div class="detail-school-item">
          <strong>${escapeHtml(entry.schoolName)}</strong>
          <span>${escapeHtml(formatDateHistoryRange(entry))}</span>
          <span>${escapeHtml(formatDateHistoryMeta(entry))}</span>
        </div>
      `;
    })
    .join("");

  return `
    <div class="detail-history">
      <h3>Change History</h3>
      <div class="detail-school-list">${itemsMarkup}</div>
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

function normalizeNoteStorageKey(rawKey) {
  if (!rawKey) {
    return null;
  }

  const trimmed = String(rawKey).trim();
  const [rawDateKey, rawZone] = trimmed.split("::");
  const dateKey = normalizeDateKey(rawDateKey);

  if (!dateKey) {
    return null;
  }

  if (!rawZone) {
    return dateKey;
  }

  return `${dateKey}::${String(rawZone).trim().toUpperCase()}`;
}

function buildNoteStorageKey(dateKey, zone) {
  if (!zone) {
    return dateKey;
  }

  return `${dateKey}::${zone}`;
}

function getSelectedNoteStorageKey() {
  if (!state.selectedNoteDateKey) {
    return "";
  }

  return buildNoteStorageKey(state.selectedNoteDateKey, state.selectedNoteZone);
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

async function persistConfirmed(rowNumber, confirmed) {
  const response = await fetch(LIMITS_API_URL, {
    method: "POST",
    headers: {
      "Content-Type": "text/plain;charset=utf-8",
    },
    body: JSON.stringify({
      action: "setConfirmed",
      rowNumber,
      confirmed,
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

function getHolidayNames(dateKey) {
  const date = parseDateKey(dateKey);
  const year = date.getFullYear();
  const holidays = buildUsHolidayMap(year);
  return holidays[dateKey] || [];
}

function buildUsHolidayMap(year) {
  const holidays = {};

  addHoliday(holidays, observedDate(year, 0, 1), "New Year's Day");
  addHoliday(holidays, nthWeekdayOfMonth(year, 0, 1, 3), "Martin Luther King Jr. Day");
  addHoliday(holidays, nthWeekdayOfMonth(year, 1, 1, 3), "Presidents Day");
  addHoliday(holidays, lastWeekdayOfMonth(year, 4, 1), "Memorial Day");
  addHoliday(holidays, observedDate(year, 5, 19), "Juneteenth");
  addHoliday(holidays, observedDate(year, 6, 4), "Independence Day");
  addHoliday(holidays, nthWeekdayOfMonth(year, 8, 1, 1), "Labor Day");
  addHoliday(holidays, nthWeekdayOfMonth(year, 9, 1, 2), "Columbus Day");
  addHoliday(holidays, observedDate(year, 10, 11), "Veterans Day");
  addHoliday(holidays, nthWeekdayOfMonth(year, 10, 4, 4), "Thanksgiving");
  addHoliday(holidays, observedDate(year, 11, 25), "Christmas Day");
  addSecondaryHolidayEntries(holidays, year);

  return holidays;
}

function addHoliday(map, date, label) {
  const key = formatDateKey(date);
  if (!map[key]) {
    map[key] = [];
  }
  map[key].push(label);
}

function addSecondaryHolidayEntries(map, year) {
  const entries = {
    2025: [
      ["2025-01-29", "Lunar New Year"],
      ["2025-03-30", "Eid al-Fitr"],
      ["2025-03-31", "Eid al-Fitr"],
      ["2025-10-20", "Diwali"],
      ["2025-09-23", "Rosh Hashanah"],
      ["2025-09-24", "Rosh Hashanah"],
      ["2025-10-02", "Yom Kippur"],
    ],
    2026: [
      ["2026-02-17", "Lunar New Year"],
      ["2026-03-20", "Eid al-Fitr"],
      ["2026-09-12", "Rosh Hashanah"],
      ["2026-09-13", "Rosh Hashanah"],
      ["2026-09-21", "Yom Kippur"],
      ["2026-11-08", "Diwali"],
    ],
    2027: [
      ["2027-02-06", "Lunar New Year"],
      ["2027-03-10", "Eid al-Fitr"],
      ["2027-10-02", "Rosh Hashanah"],
      ["2027-10-03", "Rosh Hashanah"],
      ["2027-10-11", "Yom Kippur"],
      ["2027-10-28", "Diwali"],
    ],
    2028: [
      ["2028-01-26", "Lunar New Year"],
      ["2028-02-27", "Eid al-Fitr"],
      ["2028-09-21", "Rosh Hashanah"],
      ["2028-09-22", "Rosh Hashanah"],
      ["2028-09-30", "Yom Kippur"],
      ["2028-10-17", "Diwali"],
    ],
    2029: [
      ["2029-02-13", "Lunar New Year"],
      ["2029-02-15", "Eid al-Fitr"],
      ["2029-09-10", "Rosh Hashanah"],
      ["2029-09-11", "Rosh Hashanah"],
      ["2029-09-19", "Yom Kippur"],
      ["2029-11-05", "Diwali"],
    ],
    2030: [
      ["2030-02-03", "Lunar New Year"],
      ["2030-02-05", "Eid al-Fitr"],
      ["2030-09-28", "Rosh Hashanah"],
      ["2030-09-29", "Rosh Hashanah"],
      ["2030-10-07", "Yom Kippur"],
      ["2030-10-25", "Diwali"],
    ],
  };

  (entries[year] || []).forEach(([dateKey, label]) => {
    if (!map[dateKey]) {
      map[dateKey] = [];
    }
    map[dateKey].push(label);
  });
}

function observedDate(year, monthIndex, day) {
  const date = new Date(year, monthIndex, day);
  const weekday = date.getDay();

  if (weekday === 6) {
    return new Date(year, monthIndex, day - 1);
  }

  if (weekday === 0) {
    return new Date(year, monthIndex, day + 1);
  }

  return date;
}

function nthWeekdayOfMonth(year, monthIndex, weekday, nth) {
  const firstDay = new Date(year, monthIndex, 1);
  const offset = (7 + weekday - firstDay.getDay()) % 7;
  return new Date(year, monthIndex, 1 + offset + (nth - 1) * 7);
}

function lastWeekdayOfMonth(year, monthIndex, weekday) {
  const lastDay = new Date(year, monthIndex + 1, 0);
  const offset = (7 + lastDay.getDay() - weekday) % 7;
  return new Date(year, monthIndex, lastDay.getDate() - offset);
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
