const state = {
  events: [],
  dataVersion: null,
  cloudFiles: {},
  activeDetailEventId: "",
  cloudRefreshTimer: null,
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
  summaryPanel: document.querySelector("#summaryPanel"),
  analysisGrid: document.querySelector("#analysisGrid"),
  currentMonth: document.querySelector("#currentMonth"),
  prevMonth: document.querySelector("#prevMonth"),
  nextMonth: document.querySelector("#nextMonth"),
  monthView: document.querySelector("#monthView"),
  yearView: document.querySelector("#yearView"),
  summaryView: document.querySelector("#summaryView"),
  regionFilter: document.querySelector("#regionFilter"),
  statusFilter: document.querySelector("#statusFilter"),
  typeFilter: document.querySelector("#typeFilter"),
  ownerFilter: document.querySelector("#ownerFilter"),
  searchInput: document.querySelector("#searchInput"),
  syncStatus: document.querySelector("#syncStatus"),
  lastUpdated: document.querySelector("#lastUpdated"),
  summaryStrip: document.querySelector("#summaryStrip"),
  allEventsPanel: document.querySelector(".all-events-panel"),
  allEventsCount: document.querySelector("#allEventsCount"),
  allEventsList: document.querySelector("#allEventsList"),
  eventDataTools: document.querySelector("#eventDataTools"),
  addEventButton: document.querySelector("#addEventButton"),
  eventImportInput: document.querySelector("#eventImportInput"),
  eventImportButton: document.querySelector("#eventImportButton"),
  eventExportButton: document.querySelector("#eventExportButton"),
  detailDrawer: document.querySelector("#detailDrawer"),
  detailContent: document.querySelector("#detailContent"),
  closeDrawer: document.querySelector("#closeDrawer"),
  drawerScrim: document.querySelector("#drawerScrim"),
  emptyStateTemplate: document.querySelector("#emptyStateTemplate"),
};

const supportedFileAccept = ".pdf,.doc,.docx,.ppt,.pptx,.xls,.xlsx,.jpg,.jpeg,.png";
const supportedFileExtensions = ["pdf", "doc", "docx", "ppt", "pptx", "xls", "xlsx", "jpg", "jpeg", "png"];
const supportedFileLabel = "PDF、Word、PPT、Excel、JPG、PNG";
const cloudUploadLimitBytes = 50 * 1024 * 1024;
const meetingFormFields = [
  ["projectId", "项目编号", "text", "留空将自动生成"],
  ["title", "活动/会议名称", "text", "必填"],
  ["date", "日期", "date", ""],
  ["region", "区域", "text", ""],
  ["location", "地理位置", "text", ""],
  ["owner", "项目负责人", "text", ""],
  ["projectType", "项目类型", "text", ""],
  ["projectName", "项目名称", "text", ""],
  ["status", "项目状态", "text", "进行中 / 已完成"],
  ["participation", "我司参与环节", "text", ""],
  ["guestText", "参与嘉宾", "text", ""],
  ["topic", "讲题/内容", "text", ""],
  ["progress", "关键进度/问题", "text", ""],
  ["notes", "备注", "text", ""],
];

function parseDate(value) {
  if (!value) return null;
  const date = new Date(`${value}T00:00:00`);
  return Number.isNaN(date.getTime()) ? null : date;
}

function formatFileSize(bytes) {
  if (!Number.isFinite(bytes)) return "";
  const units = ["B", "KB", "MB", "GB"];
  let value = bytes;
  let index = 0;
  while (value >= 1024 && index < units.length - 1) {
    value /= 1024;
    index += 1;
  }
  return `${value.toFixed(index === 0 ? 0 : 1)} ${units[index]}`;
}

async function readApiResponse(response) {
  const text = await response.text();
  if (!text.trim()) return {};
  try {
    return JSON.parse(text);
  } catch (error) {
    const message = text.replace(/\s+/g, " ").trim();
    return { error: message.length > 300 ? `${message.slice(0, 300)}...` : message || `HTTP ${response.status}` };
  }
}

async function uploadFileToSignedUrl(signedUrl, file) {
  const formData = new FormData();
  formData.append("cacheControl", "3600");
  formData.append("", file);
  const response = await fetch(signedUrl, {
    method: "PUT",
    headers: { "x-upsert": "true" },
    body: formData,
  });
  if (!response.ok) {
    const message = (await response.text()).replace(/\s+/g, " ").trim();
    throw new Error(message || `HTTP ${response.status}`);
  }
}

function formatMonth(date) {
  return `${date.getFullYear()}年${String(date.getMonth() + 1).padStart(2, "0")}月`;
}

function formatYear(date) {
  return `${date.getFullYear()}年`;
}

function formatPeriod(date) {
  if (state.viewMode === "summary") return `${date.getFullYear()}年汇总`;
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

function isCloudSnapshot() {
  return window.location.hostname.endsWith("github.io") || !window.location.pathname.startsWith("/");
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
  if (state.viewMode === "summary") return events;
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
  const attachments = visible.reduce((sum, event) => sum + event.attachments.length, 0);
  const guests = new Set(visible.flatMap((event) => event.guests || [])).size;
  const active = visible.filter((event) => !event.status.includes("完成")).length;
  const labelPrefix = state.viewMode === "summary" ? "筛选后" : state.viewMode === "year" ? "本年" : "本月";
  const metrics = [
    [`${labelPrefix}会议`, visible.length],
    ["未完成/推进中", active],
    ["涉及嘉宾", guests],
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
    cell.className = `day-cell${day.getMonth() === month ? "" : " outside"}${dayEvents.length ? "" : " empty"}`;
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

function countBy(events, getter, fallback = "未填写") {
  return events.reduce((counts, event) => {
    const rawValue = getter(event);
    const value = rawValue && String(rawValue).trim() ? String(rawValue).trim() : fallback;
    counts.set(value, (counts.get(value) || 0) + 1);
    return counts;
  }, new Map());
}

function sortedCounts(counts) {
  return [...counts.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], "zh-CN"));
}

function renderDistribution(title, items, total) {
  const max = Math.max(1, ...items.map(([, value]) => value));
  if (!items.length) {
    return `
      <section class="analysis-card">
        <div class="analysis-card-head">
          <h3>${escapeHtml(title)}</h3>
          <span>0 项</span>
        </div>
        <div class="empty-inline">当前筛选下暂无数据。</div>
      </section>
    `;
  }
  return `
    <section class="analysis-card">
      <div class="analysis-card-head">
        <h3>${escapeHtml(title)}</h3>
        <span>${items.length} 项</span>
      </div>
      <div class="bar-list">
        ${items
          .map(([label, value]) => {
            const percent = total ? Math.round((value / total) * 100) : 0;
            const width = Math.max(8, Math.round((value / max) * 100));
            return `
              <div class="bar-item">
                <div class="bar-label">
                  <span>${escapeHtml(label)}</span>
                  <strong>${value} 场 · ${percent}%</strong>
                </div>
                <div class="bar-track" aria-label="${escapeHtml(label)} ${value} 场">
                  <span style="width: ${width}%"></span>
                </div>
              </div>
            `;
          })
          .join("")}
      </div>
    </section>
  `;
}

function renderMonthDistribution(events) {
  const year = state.currentDate.getFullYear();
  const counts = Array.from({ length: 12 }, (_, month) => {
    const value = events.filter((event) => {
      const date = parseDate(event.date);
      return date && date.getFullYear() === year && date.getMonth() === month;
    }).length;
    return [`${month + 1}月`, value];
  });
  return renderDistribution(`${year}年时间分布`, counts, counts.reduce((sum, [, value]) => sum + value, 0));
}

function renderAnalysis(events) {
  const total = events.length;
  const complete = events.filter((event) => event.status.includes("完成")).length;
  const active = total - complete;
  const owners = countBy(events, (event) => event.owner);
  const regions = countBy(events, (event) => event.region);
  const projectTypes = countBy(events, (event) => event.projectType);
  const projects = countBy(events, (event) => event.projectName || event.projectType);
  const attachments = events.reduce((sum, event) => sum + event.attachments.length, 0);
  const guests = new Set(events.flatMap((event) => event.guests || [])).size;

  els.analysisGrid.innerHTML = `
    <section class="analysis-overview" aria-label="汇总指标">
      <div class="analysis-metric"><span>会议总数</span><strong>${total}</strong></div>
      <div class="analysis-metric"><span>推进中/未完成</span><strong>${active}</strong></div>
      <div class="analysis-metric"><span>已完成</span><strong>${complete}</strong></div>
      <div class="analysis-metric"><span>相关文件</span><strong>${attachments}</strong></div>
      <div class="analysis-metric"><span>涉及嘉宾</span><strong>${guests}</strong></div>
    </section>
    <section class="analysis-layout">
      ${renderDistribution("区域分布", sortedCounts(regions), total)}
      ${renderDistribution("项目分布", sortedCounts(projects), total)}
      ${renderMonthDistribution(events)}
      ${renderDistribution("负责人分布", sortedCounts(owners), total)}
      ${renderDistribution("项目类型分布", sortedCounts(projectTypes), total)}
    </section>
  `;
}

function render() {
  const filtered = getFilteredEvents();
  els.currentMonth.textContent = formatPeriod(state.currentDate);
  els.monthView.classList.toggle("active", state.viewMode === "month");
  els.yearView.classList.toggle("active", state.viewMode === "year");
  els.summaryView.classList.toggle("active", state.viewMode === "summary");
  els.calendarPanel.hidden = state.viewMode !== "month";
  els.yearPanel.hidden = state.viewMode !== "year";
  els.summaryPanel.hidden = state.viewMode !== "summary";
  els.allEventsPanel.hidden = state.viewMode === "summary";

  renderSummary(filtered);
  if (state.viewMode === "summary") {
    renderAnalysis(filtered);
  } else if (state.viewMode === "year") {
    renderYear(filtered);
  } else {
    renderCalendar(filtered);
  }
  if (state.viewMode !== "summary") {
    renderAllEvents(filtered);
  }
}

function renderFileTools(event) {
  if (isCloudSnapshot()) return "";
  const exportUrl = `/api/export-files?eventId=${encodeURIComponent(event.id)}`;
  return `
    <div class="file-tools">
      <input id="cloudUploadInput" type="file" accept="${supportedFileAccept}" hidden />
      <button class="primary-button" id="cloudUploadButton" type="button">上传文件</button>
      <a class="secondary-button" href="${exportUrl}" target="_blank" rel="noreferrer">下载全部</a>
      <span id="cloudUploadStatus">支持 ${supportedFileLabel}，单个文件不超过 50MB。</span>
    </div>
  `;
}

function renderCloudFileList(eventId) {
  const cache = state.cloudFiles[eventId];
  if (!cache) {
    return `<div class="info-box cloud-file-list" id="cloudFiles"><p>正在读取云端文件...</p></div>`;
  }
  if (cache.error) {
    return `<div class="info-box cloud-file-list" id="cloudFiles"><p>${escapeHtml(cache.error)}</p></div>`;
  }
  if (!cache.configured) {
    return `<div class="info-box cloud-file-list" id="cloudFiles"><p>尚未配置 Supabase Storage，配置后可上传和下载课件/附件。</p></div>`;
  }
  if (!cache.files.length) {
    return `<div class="info-box cloud-file-list" id="cloudFiles"><p>云端暂无文件。</p></div>`;
  }
  return `
    <div class="file-section cloud-file-list" id="cloudFiles">
      <div class="file-group">
        <h4>课件/附件</h4>
        ${cache.files
          .map(
            (file) => `
              <div class="file-link">
                <a class="file-name" href="${file.url}" target="_blank" rel="noreferrer">${escapeHtml(file.name)}</a>
                <span>${escapeHtml(file.modifiedAt || "可下载")}</span>
                <a class="file-action" href="${file.url}" download>下载</a>
              </div>
            `,
          )
          .join("")}
      </div>
    </div>
  `;
}

function renderAttachmentArea(event) {
  return `
    <h3 class="section-title">课件/附件</h3>
    ${renderFileTools(event)}
    ${renderCloudFileList(event.id)}
  `;
}

async function loadCloudFiles(eventId) {
  if (isCloudSnapshot()) return;
  try {
    const response = await fetch(`/api/cloud-files?eventId=${encodeURIComponent(eventId)}&ts=${Date.now()}`);
    const payload = await readApiResponse(response);
    state.cloudFiles[eventId] = {
      configured: Boolean(payload.configured),
      files: payload.files || [],
      error: payload.error || "",
    };
  } catch (error) {
    state.cloudFiles[eventId] = { configured: false, files: [], error: `云端文件读取失败：${error.message}` };
  }
  const container = document.querySelector("#cloudFiles");
  if (container) {
    const wrapper = document.createElement("div");
    wrapper.innerHTML = renderCloudFileList(eventId).trim();
    container.replaceWith(wrapper.firstElementChild);
  }
}

function bindUploadControls(eventId) {
  const input = document.querySelector("#cloudUploadInput");
  const button = document.querySelector("#cloudUploadButton");
  const status = document.querySelector("#cloudUploadStatus");
  if (!input || !button || !status) return;
  button.addEventListener("click", () => input.click());
  input.addEventListener("change", async () => {
    const file = input.files && input.files[0];
    if (!file) return;
    const extension = file.name.split(".").pop().toLowerCase();
    if (!supportedFileExtensions.includes(extension)) {
      status.textContent = `文件格式不支持，请选择 ${supportedFileLabel}。`;
      input.value = "";
      return;
    }
    if (file.size > cloudUploadLimitBytes) {
      status.textContent = `文件过大：${formatFileSize(file.size)}，单个文件不能超过 ${formatFileSize(cloudUploadLimitBytes)}。`;
      input.value = "";
      return;
    }
    button.disabled = true;
    status.textContent = "正在准备上传...";
    try {
      const response = await fetch(`/api/upload-url?eventId=${encodeURIComponent(eventId)}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ filename: file.name, contentType: file.type || "application/octet-stream" }),
      });
      const payload = await readApiResponse(response);
      if (!response.ok || payload.error) throw new Error(payload.error || `HTTP ${response.status}`);
      status.textContent = "正在上传到云端...";
      await uploadFileToSignedUrl(payload.upload.signedUrl, file);
      status.textContent = "上传完成，云端文件列表已更新。";
      input.value = "";
      await loadCloudFiles(eventId);
      await loadEvents({ silent: true });
    } catch (error) {
      status.textContent = `上传失败：${error.message}`;
    } finally {
      button.disabled = false;
    }
  });
}

function startCloudFileSync(eventId) {
  if (state.cloudRefreshTimer) {
    clearInterval(state.cloudRefreshTimer);
  }
  state.activeDetailEventId = eventId;
  loadCloudFiles(eventId);
  state.cloudRefreshTimer = window.setInterval(() => {
    if (state.activeDetailEventId === eventId && els.detailDrawer.classList.contains("open")) {
      loadCloudFiles(eventId);
    }
  }, 30000);
}

function stopCloudFileSync() {
  state.activeDetailEventId = "";
  if (state.cloudRefreshTimer) {
    clearInterval(state.cloudRefreshTimer);
    state.cloudRefreshTimer = null;
  }
}

function renderMeetingForm() {
  return `
    <header class="detail-head">
      <div class="detail-meta">
        <span class="tag">会议数据</span>
        <span class="tag">Excel同步</span>
      </div>
      <h2 id="detailTitle">新增会议</h2>
      <div class="detail-meta">
        <span>保存后会写入会议汇总 Excel</span>
      </div>
    </header>
    <form id="eventForm" class="event-form">
      ${meetingFormFields
        .map(
          ([key, label, type, placeholder]) => `
            <label class="form-field">
              <span>${escapeHtml(label)}</span>
              ${key === "progress" || key === "notes" ? `<textarea name="${key}" rows="3" placeholder="${escapeHtml(placeholder)}"></textarea>` : `<input name="${key}" type="${type}" placeholder="${escapeHtml(placeholder)}" />`}
            </label>
          `,
        )
        .join("")}
      <div class="form-actions">
        <button class="primary-button" type="submit">保存会议</button>
        <button class="secondary-button" type="button" id="cancelEventForm">取消</button>
        <span id="eventFormStatus"></span>
      </div>
    </form>
  `;
}

function openEventForm() {
  if (isCloudSnapshot()) return;
  els.detailContent.innerHTML = renderMeetingForm();
  els.detailDrawer.classList.add("open");
  els.detailDrawer.setAttribute("aria-hidden", "false");
  bindEventForm();
}

function collectEventForm(form) {
  return Object.fromEntries(meetingFormFields.map(([key]) => [key, form.elements[key]?.value || ""]));
}

function bindEventForm() {
  const form = document.querySelector("#eventForm");
  const cancelButton = document.querySelector("#cancelEventForm");
  const status = document.querySelector("#eventFormStatus");
  if (!form || !status) return;
  cancelButton?.addEventListener("click", closeDetail);
  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    const payload = collectEventForm(form);
    if (!payload.title.trim()) {
      status.textContent = "请填写活动/会议名称。";
      return;
    }
    status.textContent = "正在保存...";
    try {
      const response = await fetch("/api/events", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const result = await readApiResponse(response);
      if (!response.ok || result.error) throw new Error(result.error || `HTTP ${response.status}`);
      status.textContent = "保存完成。";
      await loadEvents({ silent: true });
      closeDetail();
    } catch (error) {
      status.textContent = `保存失败：${error.message}`;
    }
  });
}

async function deleteEvent(eventId) {
  if (!eventId || !window.confirm("确认删除这场会议吗？该操作会从会议汇总 Excel 中移除该行。")) return;
  try {
    const response = await fetch(`/api/events?eventId=${encodeURIComponent(eventId)}`, { method: "DELETE" });
    const payload = await readApiResponse(response);
    if (!response.ok || payload.error) throw new Error(payload.error || `HTTP ${response.status}`);
    await loadEvents({ silent: true });
    closeDetail();
  } catch (error) {
    window.alert(`删除失败：${error.message}`);
  }
}

function bindEventDataTools() {
  if (isCloudSnapshot()) {
    if (els.eventDataTools) els.eventDataTools.hidden = true;
    return;
  }
  els.addEventButton?.addEventListener("click", openEventForm);
  els.eventImportButton?.addEventListener("click", () => els.eventImportInput?.click());
  els.eventImportInput?.addEventListener("change", async () => {
    const file = els.eventImportInput.files && els.eventImportInput.files[0];
    if (!file) return;
    const extension = file.name.split(".").pop().toLowerCase();
    if (!["xlsx", "xlsm"].includes(extension)) {
      window.alert("请导入 .xlsx / .xlsm 格式的会议 Excel。");
      els.eventImportInput.value = "";
      return;
    }
    const formData = new FormData();
    formData.append("file", file);
    els.eventImportButton.disabled = true;
    els.eventImportButton.textContent = "导入中...";
    try {
      const response = await fetch("/api/import-events", {
        method: "POST",
        body: formData,
      });
      const payload = await readApiResponse(response);
      if (!response.ok || payload.error) throw new Error(payload.error || `HTTP ${response.status}`);
      await loadEvents({ silent: true });
      window.alert(`导入完成：新增 ${payload.created} 条，更新 ${payload.updated} 条，跳过 ${payload.skipped} 条。`);
    } catch (error) {
      window.alert(`导入失败：${error.message}`);
    } finally {
      els.eventImportInput.value = "";
      els.eventImportButton.disabled = false;
      els.eventImportButton.textContent = "导入会议Excel";
    }
  });
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
      ${isCloudSnapshot() ? "" : `<div class="detail-actions"><button class="danger-button" id="deleteEventButton" type="button">删除会议</button></div>`}
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
    ${renderAttachmentArea(event)}
  `;
  els.detailDrawer.classList.add("open");
  els.detailDrawer.setAttribute("aria-hidden", "false");
  document.querySelector("#deleteEventButton")?.addEventListener("click", () => deleteEvent(event.id));
  bindUploadControls(event.id);
  startCloudFileSync(event.id);
}

function closeDetail() {
  stopCloudFileSync();
  els.detailDrawer.classList.remove("open");
  els.detailDrawer.setAttribute("aria-hidden", "true");
}

async function loadEvents({ silent = false } = {}) {
  if (!silent) {
    els.syncStatus.className = "sync-dot";
    els.lastUpdated.textContent = "正在读取数据";
  }
  try {
    const isGithubPages = window.location.hostname.endsWith("github.io");
    const primaryUrl = isGithubPages ? `./events.json?ts=${Date.now()}` : `/api/events?ts=${Date.now()}`;
    let response = await fetch(primaryUrl);
    if (!response.ok && !isGithubPages) {
      response = await fetch(`./events.json?ts=${Date.now()}`);
    }
    const payload = await readApiResponse(response);
    if (!response.ok || payload.error) throw new Error(payload.error || `HTTP ${response.status}`);
    state.events = payload.events || [];
    state.dataVersion = payload.dataVersion;
    refreshFilterOptions();
    render();
    els.syncStatus.className = "sync-dot ready";
    els.lastUpdated.textContent = isGithubPages
      ? `云端快照 ${payload.count} 场 · ${payload.generatedAt}`
      : `已同步 ${payload.count} 场 · ${payload.generatedAt}`;
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
    const delta = state.viewMode === "month" ? -1 : -12;
    state.currentDate = new Date(state.currentDate.getFullYear(), state.currentDate.getMonth() + delta, 1);
    render();
  });
  els.nextMonth.addEventListener("click", () => {
    const delta = state.viewMode === "month" ? 1 : 12;
    state.currentDate = new Date(state.currentDate.getFullYear(), state.currentDate.getMonth() + delta, 1);
    render();
  });
  els.monthView.addEventListener("click", () => setViewMode("month"));
  els.yearView.addEventListener("click", () => setViewMode("year"));
  els.summaryView.addEventListener("click", () => setViewMode("summary"));
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
  bindEventDataTools();
}

bindEvents();
loadEvents();
setInterval(() => loadEvents({ silent: true }), 30000);
