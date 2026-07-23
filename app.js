const state = {
  events: [],
  dataVersion: null,
  currentDate: new Date(),
  viewMode: "month",
  filters: {
    region: "",
    status: "",
    type: "",
    owner: "",
    search: "",
  },
};

const els = {
  calendarPanel: document.querySelector("#calendarPanel"),
  calendarGrid: document.querySelector("#calendarGrid"),
  yearPanel: document.querySelector("#yearPanel"),
  yearGrid: document.querySelector("#yearGrid"),
  currentMonth: document.querySelector("#currentMonth"),
  prevMonth: document.querySelector("#prevMonth"),
  nextMonth: document.querySelector("#nextMonth"),
  monthView: document.querySelector("#monthView"),
  yearView: document.querySelector("#yearView"),
  regionFilter: document.querySelector("#regionFilter"),
  statusFilter: document.querySelector("#statusFilter"),
  typeFilter: document.querySelector("#typeFilter"),
  ownerFilter: document.querySelector("#ownerFilter"),
  searchInput: document.querySelector("#searchInput"),
  syncStatus: document.querySelector("#syncStatus"),
  lastUpdated: document.querySelector("#lastUpdated"),
  summaryStrip: document.querySelector("#summaryStrip"),
  allEventsCount: document.querySelector("#allEventsCount"),
  allEventsList: document.querySelector("#allEventsList"),
  detailDrawer: document.querySelector("#detailDrawer"),
  detailContent: document.querySelector("#detailContent"),
  closeDrawer: document.querySelector("#closeDrawer"),
  drawerScrim: document.querySelector("#drawerScrim"),
  emptyStateTemplate: document.querySelector("#emptyStateTemplate"),
};

function parseDate(value) {
  if (!value) return null;
  const date = new Date(`${value}T00:00:00`);
  return Number.isNaN(date.getTime()) ? null : date;
}

function formatMonth(date) {
  return `${date.getFullYear()}年${String(date.getMonth() + 1).padStart(2, "0")}月`;
}

function formatYear(date) {
  return `${date.getFullYear()}年`;
}

function formatPeriod(date) {
  return state.viewMode === "year" ? formatYear(date) : formatMonth(date);
}

function formatDate(value) {
  const date = parseDate(value);
  if (!date) return value || "未填写";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function uniqueValues(key) {
  return [...new Set(state.events.map((event) => event[key]).filter(Boolean))].sort((a, b) => a.localeCompare(b, "zh-CN"));
}

function setOptions(select, values, allLabel) {
  const current = select.value;
  select.innerHTML = [`<option value="">${allLabel}</option>`, ...values.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`)].join("");
  select.value = values.includes(current) ? current : "";
}

function refreshFilterOptions() {
  setOptions(els.regionFilter, uniqueValues("region"), "全部区域");
  setOptions(els.statusFilter, uniqueValues("status"), "全部状态");
  setOptions(els.typeFilter, uniqueValues("projectType"), "全部类型");
  setOptions(els.ownerFilter, uniqueValues("owner"), "全部负责人");
}

function eventSearchText(event) {
  return [
    event.projectId,
    event.title,
    event.region,
    event.location,
    event.owner,
    event.participation,
    event.guestText,
    event.status,
    event.projectType,
    event.projectName,
  ]
    .join(" ")
    .toLowerCase();
}

function getFilteredEvents() {
  const term = state.filters.search.trim().toLowerCase();
  return state.events.filter((event) => {
    if (state.filters.region && event.region !== state.filters.region) return false;
    if (state.filters.status && event.status !== state.filters.status) return false;
    if (state.filters.type && event.projectType !== state.filters.type) return false;
    if (state.filters.owner && event.owner !== state.filters.owner) return false;
    if (term && !eventSearchText(event).includes(term)) return false;
    return true;
  });
}

function monthEvents(events) {
  const year = state.currentDate.getFullYear();
  const month = state.currentDate.getMonth();
  return events.filter((event) => {
    const date = parseDate(event.date);
    return date && date.getFullYear() === year && date.getMonth() === month;
  });
}

function yearEvents(events) {
  const year = state.currentDate.getFullYear();
  return events.filter((event) => {
    const date = parseDate(event.date);
    return date && date.getFullYear() === year;
  });
}

function visiblePeriodEvents(events) {
  return state.viewMode === "year" ? yearEvents(events) : monthEvents(events);
}

function compareEvents(a, b) {
  const dateA = parseDate(a.date);
  const dateB = parseDate(b.date);
  if (dateA && dateB) return dateA - dateB;
  if (dateA) return -1;
  if (dateB) return 1;
  return a.title.localeCompare(b.title, "zh-CN");
}

function statusClass(status) {
  if (!status) return "";
  if (status.includes("完成")) return "done";
  if (status.includes("进行") || status.includes("待") || status.includes("筹备")) return "pending";
  return "";
}

function renderSummary(events) {
  const visible = visiblePeriodEvents(events);
  const withFolders = visible.filter((event) => event.folderMatched).length;
  const attachments = visible.reduce((sum, event) => sum + event.attachments.length, 0);
  const guests = new Set(visible.flatMap((event) => event.guests || [])).size;
  const active = visible.filter((event) => !event.status.includes("完成")).length;
  const labelPrefix = state.viewMode === "year" ? "本年" : "本月";
  const metrics = [
    [`${labelPrefix}会议`, visible.length],
    ["未完成/推进中", active],
    ["涉及嘉宾", guests],
    ["已匹配资料夹", withFolders],
    ["关联文件", attachments],
  ];
  els.summaryStrip.innerHTML = metrics
    .map(([label, value]) => `<div class="metric"><span>${label}</span><strong>${value}</strong></div>`)
    .join("");
}

function renderAllEvents(events) {
  els.allEventsCount.textContent = `${events.length} 场`;
  if (!events.length) {
    els.allEventsList.innerHTML = `<div class="empty-inline">当前筛选下没有会议。</div>`;
    return;
  }

  els.allEventsList.innerHTML = "";
  [...events].sort(compareEvents).forEach((event) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "event-row";
    const metaText = [event.region, event.location, event.owner].filter(Boolean).join(" · ") || "未填写";
    const progressText = event.progress ? ` · 最新进度：${event.progress}` : "";
    button.innerHTML = `
      <span class="row-date">${event.date ? formatDate(event.date) : "未排期"}</span>
      <span class="row-main">
        <strong>${escapeHtml(event.title)}</strong>
        <span>${escapeHtml(`${metaText}${progressText}`)}</span>
      </span>
      <span class="row-tags">
        <span class="tag">${escapeHtml(event.status || "未填状态")}</span>
        <span class="tag">${event.attachments.length} 文件</span>
      </span>
    `;
    button.addEventListener("click", () => openDetail(event.id));
    els.allEventsList.appendChild(button);
  });
}

function renderCalendar(events) {
  const year = state.currentDate.getFullYear();
  const month = state.currentDate.getMonth();
  const first = new Date(year, month, 1);
  const mondayOffset = (first.getDay() + 6) % 7;
  const start = new Date(year, month, 1 - mondayOffset);
  const days = [];
  for (let index = 0; index < 42; index += 1) {
    const day = new Date(start);
    day.setDate(start.getDate() + index);
    days.push(day);
  }

  const byDate = new Map();
  for (const event of events) {
    const key = event.date;
    if (!byDate.has(key)) byDate.set(key, []);
    byDate.get(key).push(event);
  }

  els.calendarGrid.innerHTML = "";
  for (const day of days) {
    const key = `${day.getFullYear()}-${String(day.getMonth() + 1).padStart(2, "0")}-${String(day.getDate()).padStart(2, "0")}`;
    const dayEvents = byDate.get(key) || [];
    const cell = document.createElement("div");
    cell.className = `day-cell${day.getMonth() === month ? "" : " outside"}`;
    cell.innerHTML = `
      <div class="day-header">
        <span class="day-number">${day.getDate()}</span>
        ${dayEvents.length ? `<span class="count-badge">${dayEvents.length}</span>` : ""}
      </div>
    `;
    dayEvents.forEach((event) => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = `event-button ${statusClass(event.status)}`;
      button.innerHTML = `
        <span class="event-title">${escapeHtml(event.title)}</span>
        <span class="event-meta">
          <span>${escapeHtml(event.region || "未填区域")}</span>
          <span>·</span>
          <span>${escapeHtml(event.status || "未填状态")}</span>
        </span>
      `;
      button.addEventListener("click", () => openDetail(event.id));
      cell.appendChild(button);
    });
    els.calendarGrid.appendChild(cell);
  }

  if (!monthEvents(events).length) {
    els.calendarGrid.appendChild(els.emptyStateTemplate.content.cloneNode(true));
  }
}

function renderYear(events) {
  const year = state.currentDate.getFullYear();
  const months = Array.from({ length: 12 }, (_, month) => ({
    month,
    events: events
      .filter((event) => {
        const date = parseDate(event.date);
        return date && date.getFullYear() === year && date.getMonth() === month;
      })
      .sort(compareEvents),
  }));

  els.yearGrid.innerHTML = "";
  months.forEach(({ month, events: monthItems }) => {
    const card = document.createElement("section");
    card.className = "year-card";
    card.innerHTML = `
      <div class="year-month-header">
        <strong>${month + 1}月</strong>
        <span>${monthItems.length} 场</span>
      </div>
      <div class="year-month-body"></div>
    `;
    const body = card.querySelector(".year-month-body");
    if (!monthItems.length) {
      body.innerHTML = `<div class="year-empty">暂无会议</div>`;
    } else {
      monthItems.forEach((event) => {
        const date = parseDate(event.date);
        const button = document.createElement("button");
        button.type = "button";
        button.className = `year-event ${statusClass(event.status)}`;
        button.innerHTML = `
          <span>${date ? `${date.getMonth() + 1}/${date.getDate()}` : "未排期"}</span>
          <strong>${escapeHtml(event.title)}</strong>
          <em>${escapeHtml([event.region, event.status || "未填状态"].filter(Boolean).join(" · "))}</em>
        `;
        button.addEventListener("click", () => openDetail(event.id));
        body.appendChild(button);
      });
    }
    els.yearGrid.appendChild(card);
  });
}

function render() {
  const filtered = getFilteredEvents();
  els.currentMonth.textContent = formatPeriod(state.currentDate);
  els.monthView.classList.toggle("active", state.viewMode === "month");
  els.yearView.classList.toggle("active", state.viewMode === "year");
  els.calendarPanel.hidden = state.viewMode !== "month";
  els.yearPanel.hidden = state.viewMode !== "year";

  renderSummary(filtered);
  renderAllEvents(filtered);
  if (state.viewMode === "year") {
    renderYear(filtered);
  } else {
    renderCalendar(filtered);
  }
}

function groupAttachments(files) {
  return files.reduce((groups, file) => {
    const key = file.category || "其他";
    if (!groups[key]) groups[key] = [];
    groups[key].push(file);
    return groups;
  }, {});
}

function renderFileSections(files) {
  if (!files.length) {
    return `<div class="info-box"><p>暂无关联文件。</p></div>`;
  }
  const groups = groupAttachments(files);
  return `
    <div class="file-section">
      ${Object.entries(groups)
        .map(
          ([category, items]) => `
            <div class="file-group">
              <h4>${escapeHtml(category)}</h4>
              ${items
                .map(
                  (file) => `
                    <a class="file-link" href="${file.url}" target="_blank" rel="noreferrer">
                      <span>${escapeHtml(file.name)}</span>
                      <span>${escapeHtml(file.modifiedAt)}</span>
                    </a>
                  `,
                )
                .join("")}
            </div>
          `,
        )
        .join("")}
    </div>
  `;
}

function renderLocalTables(tables) {
  if (!tables.length) return "";
  return `
    <h3 class="section-title">补充信息表</h3>
    ${tables
      .map((table) => {
        if (table.error) {
          return `<div class="info-box"><p>${escapeHtml(table.fileName)} 读取失败：${escapeHtml(table.error)}</p></div>`;
        }
        return `
          <div class="table-section">
            <h4><a href="${table.url}" target="_blank" rel="noreferrer">${escapeHtml(table.fileName)}</a> · ${escapeHtml(table.sheetName)}</h4>
            <table>
              <tbody>
                ${table.rows
                  .map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`)
                  .join("")}
              </tbody>
            </table>
          </div>
        `;
      })
      .join("")}
  `;
}

function openDetail(eventId) {
  const event = state.events.find((item) => item.id === eventId);
  if (!event) return;
  els.detailContent.innerHTML = `
    <header class="detail-head">
      <div class="detail-meta">
        <span class="tag">${escapeHtml(event.status || "未填状态")}</span>
        <span class="tag">${escapeHtml(event.region || "未填区域")}</span>
        <span class="tag">${escapeHtml(event.projectType || "未填类型")}</span>
      </div>
      <h2 id="detailTitle">${escapeHtml(event.title)}</h2>
      <div class="detail-meta">
        <span>${escapeHtml(event.projectId)}</span>
        <span>${formatDate(event.date)}</span>
      </div>
    </header>
    <section class="detail-grid">
      <div class="info-box"><span>会议时间</span><strong>${formatDate(event.date)}</strong></div>
      <div class="info-box"><span>地理位置</span><strong>${escapeHtml(event.location || "未填写")}</strong></div>
      <div class="info-box"><span>项目负责人</span><strong>${escapeHtml(event.owner || "未填写")}</strong></div>
      <div class="info-box"><span>我司参与环节</span><strong>${escapeHtml(event.participation || "未填写")}</strong></div>
      <div class="info-box"><span>参会嘉宾</span><p>${escapeHtml(event.guestText || "未填写")}</p></div>
      <div class="info-box"><span>项目名称</span><p>${escapeHtml(event.projectName || "未填写")}</p></div>
      <div class="info-box"><span>关键进度/问题</span><p>${escapeHtml(event.progress || "未填写")}</p></div>
      <div class="info-box"><span>讲题/内容</span><p>${escapeHtml(event.topic || "未填写")}</p></div>
    </section>
    <h3 class="section-title">资料文件夹</h3>
    <div class="info-box">
      <p>${event.folderMatched ? escapeHtml(event.folderName) : "未找到资料文件夹"}</p>
    </div>
    <h3 class="section-title">相关文件</h3>
    ${renderFileSections(event.attachments)}
    ${renderLocalTables(event.localTables || [])}
  `;
  els.detailDrawer.classList.add("open");
  els.detailDrawer.setAttribute("aria-hidden", "false");
}

function closeDetail() {
  els.detailDrawer.classList.remove("open");
  els.detailDrawer.setAttribute("aria-hidden", "true");
}

async function loadEvents({ silent = false } = {}) {
  if (!silent) {
    els.syncStatus.className = "sync-dot";
    els.lastUpdated.textContent = "正在读取数据";
  }
  try {
    const response = await fetch(`/api/events?ts=${Date.now()}`);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const payload = await response.json();
    if (payload.error) throw new Error(payload.error);
    state.events = payload.events || [];
    state.dataVersion = payload.dataVersion;
    refreshFilterOptions();
    render();
    els.syncStatus.className = "sync-dot ready";
    els.lastUpdated.textContent = `已同步 ${payload.count} 场 · ${payload.generatedAt}`;
  } catch (error) {
    els.syncStatus.className = "sync-dot error";
    els.lastUpdated.textContent = `读取失败：${error.message}`;
  }
}

function setViewMode(viewMode) {
  state.viewMode = viewMode;
  render();
}

function bindEvents() {
  els.prevMonth.addEventListener("click", () => {
    const delta = state.viewMode === "year" ? -12 : -1;
    state.currentDate = new Date(state.currentDate.getFullYear(), state.currentDate.getMonth() + delta, 1);
    render();
  });
  els.nextMonth.addEventListener("click", () => {
    const delta = state.viewMode === "year" ? 12 : 1;
    state.currentDate = new Date(state.currentDate.getFullYear(), state.currentDate.getMonth() + delta, 1);
    render();
  });
  els.monthView.addEventListener("click", () => setViewMode("month"));
  els.yearView.addEventListener("click", () => setViewMode("year"));
  [
    [els.regionFilter, "region"],
    [els.statusFilter, "status"],
    [els.typeFilter, "type"],
    [els.ownerFilter, "owner"],
  ].forEach(([element, key]) => {
    element.addEventListener("change", () => {
      state.filters[key] = element.value;
      render();
    });
  });
  els.searchInput.addEventListener("input", () => {
    state.filters.search = els.searchInput.value;
    render();
  });
  els.closeDrawer.addEventListener("click", closeDetail);
  els.drawerScrim.addEventListener("click", closeDetail);
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") closeDetail();
  });
}

bindEvents();
loadEvents();
setInterval(() => loadEvents({ silent: true }), 30000);
