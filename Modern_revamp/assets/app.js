(function () {
  const DATA = window.MODERN_DASHBOARD_DATA || {};
  const COLORS = {
    navy: "#00353e",
    teal: "#008768",
    blue: "#109cde",
    gold: "#aecc53",
    green: "#29a36a",
    red: "#c64a3c",
    orange: "#d88928",
    gray: "#8a98a5",
    line: "#d7dddc"
  };

  const projects = DATA.projects || [];
  const allProjectIds = projects.map((project) => project.project_id);
  const state = {
    selectedIds: new Set(allProjectIds),
    inspectedId: allProjectIds[0],
    dateStart: toInputDate(DATA.meta && DATA.meta.date_min) || "2026-01-01",
    dateEnd: toInputDate(DATA.meta && DATA.meta.date_max) || "2026-12-31",
    view: "overview",
    selectionMode: "multi",
    portfolioMode: "all"
  };

  const tooltip = document.getElementById("tooltip");

  function init() {
    wireControls();
    renderProjectPicker();
    document.getElementById("date-start").value = state.dateStart;
    document.getElementById("date-end").value = state.dateEnd;
    renderAll();
  }

  function wireControls() {
    document.addEventListener("click", (event) => {
      const action = event.target.closest("[data-action]");
      const viewButton = event.target.closest("[data-view]");
      const projectChip = event.target.closest("[data-select-project-id]");
      const drill = event.target.closest("[data-drill]");
      if (projectChip) handleProjectSelection(Number(projectChip.dataset.selectProjectId));
      if (drill) openDetail(drill.dataset.drill, Number(drill.dataset.projectId || state.inspectedId));
      if (viewButton) {
        state.view = viewButton.dataset.view;
        renderAll();
      }
      if (action) handleAction(action.dataset.action);
    });
    document.addEventListener("mouseover", (event) => {
      const target = event.target.closest("[data-tooltip]");
      if (target) showTooltip(target.dataset.tooltip, event);
    });
    document.addEventListener("mousemove", (event) => {
      if (!tooltip.hidden) moveTooltip(event);
    });
    document.addEventListener("mouseout", (event) => {
      const target = event.target.closest("[data-tooltip]");
      if (target) hideTooltip();
    });

    document.getElementById("date-start").addEventListener("change", (event) => {
      state.dateStart = event.target.value;
      renderAll();
    });
    document.getElementById("date-end").addEventListener("change", (event) => {
      state.dateEnd = event.target.value;
      renderAll();
    });
  }

  function handleProjectSelection(id) {
    if (!id) return;
    if (state.selectionMode === "single") {
      state.selectedIds = new Set([id]);
    } else if (state.selectedIds.has(id) && state.selectedIds.size > 1) {
      state.selectedIds.delete(id);
    } else {
      state.selectedIds.add(id);
    }
    state.inspectedId = id;
    renderAll();
  }

  function handleAction(action) {
    if (action === "all-projects" || action === "reset-filter") {
      state.selectionMode = "multi";
      state.portfolioMode = "all";
      state.selectedIds = new Set(allProjectIds);
      state.inspectedId = allProjectIds[0];
    }
    if (action === "construction-projects") {
      state.selectionMode = "multi";
      state.portfolioMode = "construction";
      state.selectedIds = new Set(
        projects
          .filter((project) => String(project.project_stage_name || "").toLowerCase().includes("construction"))
          .map((project) => project.project_id)
      );
      state.inspectedId = [...state.selectedIds][0] || allProjectIds[0];
    }
    if (action === "export-json") {
      const blob = new Blob([JSON.stringify(filteredPayload(), null, 2)], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = "spct-modern-dashboard-view.json";
      link.click();
      URL.revokeObjectURL(url);
    }
    if (action === "toggle-selection-mode") {
      state.selectionMode = state.selectionMode === "multi" ? "single" : "multi";
      if (state.selectionMode === "single") {
        const id = state.inspectedId || [...state.selectedIds][0] || allProjectIds[0];
        state.selectedIds = new Set([id]);
      }
    }
    if (action === "window-week" || action === "window-month" || action === "window-3mo") {
      applyDatePreset(action);
    }
    if (action === "close-drawer") {
      closeDetail();
    }
    renderAll();
  }

  function applyDatePreset(action) {
    const end = parseDate(DATA.meta && DATA.meta.date_max) || new Date();
    const days = action === "window-week" ? 7 : action === "window-month" ? 31 : 92;
    const start = new Date(end.getTime() - days * 86400000);
    state.dateStart = toInputDate(start);
    state.dateEnd = toInputDate(end);
    document.getElementById("date-start").value = state.dateStart;
    document.getElementById("date-end").value = state.dateEnd;
  }

  function renderAll() {
    renderProjectPicker();
    renderViewState();
    renderSelectionMode();
    renderPortfolioMode();
    renderHero();
    renderKpis();
    renderFinanceBars();
    renderInspector();
    renderControlDonuts();
    renderProjectGantt();
    renderFinancialWaterfall();
    renderFinancialTable();
    renderControlsCharts();
    renderHseTable();
    renderSourceMedia();
    renderRiskGantt();
    renderRiskMix();
    renderActionTable();
    renderCorrespondencePulse();
    renderWeeklyNarrative();
    renderCorrespondenceTable();
    renderFeedNote();
  }

  function selectedProjects() {
    return projects.filter((project) => state.selectedIds.has(project.project_id));
  }

  function selectedRows(rows, dateField) {
    const ids = state.selectedIds;
    return (rows || []).filter((row) => {
      if (!ids.has(row.project_id)) return false;
      if (!dateField) return true;
      return dateInWindow(row[dateField]);
    });
  }

  function rowsFor(tableName) {
    const lookup = {
      financials: DATA.financials,
      rfis: DATA.rfis,
      submittals: DATA.submittals,
      hse: DATA.hse,
      schedule: DATA.schedule,
      weekly_updates: DATA.weekly_updates
    };
    return selectedRows(lookup[tableName] || []);
  }

  function byProject(tableName) {
    return new Map(rowsFor(tableName).map((row) => [row.project_id, row]));
  }

  function selectedActions() {
    return selectedRows(DATA.action_items || [], "start").filter((row) => {
      if (!row.start && !row.end) return true;
      return dateInWindow(row.start) || dateInWindow(row.end);
    });
  }

  function selectedCorrespondence() {
    return selectedRows(DATA.recent_correspondence || [], "date");
  }

  function totals() {
    const financials = rowsFor("financials");
    const rfis = rowsFor("rfis");
    const submittals = rowsFor("submittals");
    const hse = rowsFor("hse");
    return {
      approvedBudget: sum(financials, "approved_budget"),
      commitments: sum(financials, "commitment_value"),
      invoiced: sum(financials, "invoiced_to_date"),
      unpaid: sum(financials, "unpaid_amount"),
      openRfis: sum(rfis, "open_rfis"),
      overdueRfis: sum(rfis, "overdue_rfis"),
      openSubmittals: sum(submittals, "open_submittals"),
      overdueSubmittals: sum(submittals, "overdue_submittals"),
      incidents: sum(hse, "weekly_incident_count"),
      observations: sum(hse, "weekly_observation_count"),
      actionItems: selectedActions().length,
      correspondence: selectedCorrespondence().length
    };
  }

  function renderProjectPicker() {
    const picker = document.getElementById("project-picker");
    picker.innerHTML = projects.map((project, index) => `
      <button class="project-chip ${state.selectedIds.has(project.project_id) ? "active" : ""}" data-select-project-id="${project.project_id}">
        <span class="project-dot" style="color:${projectColor(index)}"></span>
        <span>${escapeHtml(shortProjectName(project))}</span>
      </button>
    `).join("");
  }

  function renderSelectionMode() {
    const button = document.getElementById("selection-mode");
    if (!button) return;
    button.classList.toggle("single", state.selectionMode === "single");
    button.setAttribute("aria-pressed", state.selectionMode === "single" ? "true" : "false");
    button.querySelector("strong").textContent = state.selectionMode === "single" ? "Single" : "Multi";
  }

  function renderPortfolioMode() {
    document.querySelectorAll("[data-action='all-projects']").forEach((button) => {
      button.classList.toggle("active-filter", state.portfolioMode === "all");
    });
    document.querySelectorAll("[data-action='construction-projects']").forEach((button) => {
      button.classList.toggle("active-filter", state.portfolioMode === "construction");
    });
  }

  function renderViewState() {
    document.querySelectorAll("[data-view]").forEach((button) => {
      button.classList.toggle("active", button.dataset.view === state.view);
    });
    document.querySelectorAll(".view-section").forEach((section) => {
      section.classList.toggle("active", section.classList.contains(`${state.view}-view`));
    });
  }

  function renderFeedNote() {
    const note = document.getElementById("data-feed-note");
    const meta = DATA.meta || {};
    const sync = meta.delta_sync || {};
    note.innerHTML = `
      Source: ${escapeHtml(meta.source || "output")}<br>
      Generated: ${formatDateTime(meta.generated_at)}<br>
      Projects: ${meta.project_count || projects.length}<br>
      Source tables: ${meta.source_table_count || "local snapshot"}${sync.lookback_months ? `<br>Lookback: ${sync.lookback_months} months` : ""}
    `;
  }

  function renderHero() {
    const selected = selectedProjects();
    const t = totals();
    document.getElementById("portfolio-summary").textContent =
      `${selected.length} selected projects. ${fmtMoney(t.approvedBudget)} approved budget, ${fmtMoney(t.commitments)} committed, ${fmtMoney(t.invoiced)} invoiced to date. ${t.openRfis + t.openSubmittals} open control items and ${t.actionItems} executive action records in the selected window.`;
  }

  function renderKpis() {
    const t = totals();
    const burn = t.approvedBudget ? (t.invoiced / t.approvedBudget) : 0;
    const exposure = t.openRfis + t.openSubmittals + t.actionItems;
    const cards = [
      ["Approved Budget", fmtMoney(t.approvedBudget), `${fmtPct(burn)} invoiced against approved budget`],
      ["Commitments", fmtMoney(t.commitments), `${fmtPct(t.approvedBudget ? t.commitments / t.approvedBudget : 0)} of approved budget`],
      ["Open Controls", fmtNumber(t.openRfis + t.openSubmittals), `${t.overdueRfis + t.overdueSubmittals} overdue RFIs/submittals`],
      ["Executive Actions", fmtNumber(t.actionItems), `${selectedCorrespondence().length} correspondence records in window`],
      ["HSE Weekly Activity", fmtNumber(t.incidents + t.observations), `${t.incidents} incidents, ${t.observations} observations`]
    ];
    document.getElementById("kpi-grid").innerHTML = cards.map(([label, value, note]) => `
      <article class="kpi-card">
        <p class="eyebrow">${escapeHtml(label)}</p>
        <div class="kpi-value">${escapeHtml(value)}</div>
        <div class="kpi-note">${escapeHtml(note)}</div>
      </article>
    `).join("");
    document.getElementById("finance-subtitle").textContent = `${fmtMoney(t.unpaid)} unpaid balance`;
  }

  function renderFinanceBars() {
    const rows = rowsFor("financials").filter((row) => getProject(row.project_id));
    const series = [
      { key: "approved_budget", label: "Approved", color: COLORS.navy },
      { key: "commitment_value", label: "Committed", color: COLORS.teal },
      { key: "invoiced_to_date", label: "Invoiced", color: COLORS.gold }
    ];
    drawGroupedBars("finance-bars", rows, series, { label: (row) => shortProjectName(getProject(row.project_id)), valueFormat: fmtCompactMoney });
  }

  function renderFinancialWaterfall() {
    const rows = rowsFor("financials").filter((row) => getProject(row.project_id));
    const series = [
      { key: "approved_budget", label: "Budget", color: COLORS.navy },
      { key: "commitment_value", label: "Commit", color: COLORS.teal },
      { key: "invoiced_to_date", label: "Invoice", color: COLORS.gold },
      { key: "unpaid_amount", label: "Unpaid", color: COLORS.red }
    ];
    drawGroupedBars("financial-waterfall", rows, series, { label: (row) => shortProjectName(getProject(row.project_id)), valueFormat: fmtCompactMoney });
  }

  function renderFinancialTable() {
    const rows = rowsFor("financials")
      .filter((row) => getProject(row.project_id))
      .sort((a, b) => num(b.approved_budget) - num(a.approved_budget));
    renderTable("financial-table", ["Project", "Budget", "Committed", "Invoiced"], rows.map((row) => [
      shortProjectName(getProject(row.project_id)),
      fmtMoney(row.approved_budget),
      fmtMoney(row.commitment_value),
      fmtMoney(row.invoiced_to_date)
    ]));
  }

  function renderControlDonuts() {
    const rfis = byProject("rfis");
    const submittals = byProject("submittals");
    const selected = selectedProjects();
    document.getElementById("controls-donuts").innerHTML = selected.map((project, index) => {
      const rfi = rfis.get(project.project_id) || {};
      const sub = submittals.get(project.project_id) || {};
      const open = num(rfi.open_rfis) + num(sub.open_submittals);
      const total = num(rfi.total_rfis) + num(sub.total_submittals);
      const overdue = num(rfi.overdue_rfis) + num(sub.overdue_submittals);
      const deg = total ? Math.max(0, Math.min(360, open / total * 360)) : 0;
      return `
        <article class="donut-card">
          <div class="donut-ring" style="background:conic-gradient(${projectColor(index)} 0deg, ${projectColor(index)} ${deg}deg, #dfe8ee ${deg}deg, #dfe8ee 360deg)">
            <span>${fmtNumber(open)}</span>
          </div>
          <div>
            <h3>${escapeHtml(shortProjectName(project))}</h3>
            <div class="metric-list">
              <div class="kpi-note">Total: ${fmtNumber(total)}</div>
              <div class="kpi-note">Overdue: ${fmtNumber(overdue)}</div>
              <div class="progress-track"><span style="width:${total ? Math.min(100, open / total * 100) : 0}%"></span></div>
            </div>
            <div class="drill-actions">
              <button class="mini-button" data-drill="rfis" data-project-id="${project.project_id}">RFIs</button>
              <button class="mini-button" data-drill="submittals" data-project-id="${project.project_id}">Submittals</button>
            </div>
          </div>
        </article>
      `;
    }).join("");
  }

  function renderProjectGantt() {
    const rows = selectedProjects().map((project) => ({
      label: shortProjectName(project),
      start: project.actual_start_date || project.estimated_start_date,
      end: project.actual_completion_date || project.estimated_completion_date,
      color: projectColor(projects.indexOf(project)),
      detail: project.project_stage_name || ""
    }));
    drawGantt("project-gantt", rows, { fallbackDays: 210 });
  }

  function renderControlsCharts() {
    const rfiRows = rowsFor("rfis").filter((row) => getProject(row.project_id));
    const subRows = rowsFor("submittals").filter((row) => getProject(row.project_id));
    drawGroupedBars("rfi-chart", rfiRows, [
      { key: "open_rfis", label: "Open", color: COLORS.blue },
      { key: "overdue_rfis", label: "Overdue", color: COLORS.red }
    ], { label: (row) => shortProjectName(getProject(row.project_id)), valueFormat: fmtNumber, drill: "rfis" });
    drawGroupedBars("submittal-chart", subRows, [
      { key: "open_submittals", label: "Open", color: COLORS.teal },
      { key: "overdue_submittals", label: "Overdue", color: COLORS.red }
    ], { label: (row) => shortProjectName(getProject(row.project_id)), valueFormat: fmtNumber, drill: "submittals" });
  }

  function renderSourceMedia() {
    const target = document.getElementById("source-photo-feed");
    if (!target) return;
    const rows = selectedRows(DATA.source_photos || [], "date").slice(0, 9);
    if (!rows.length) {
      target.innerHTML = `<div class="empty-state">No bucket photo or attachment records are loaded for the selected window. Run the Delta sync to populate source-backed media metadata.</div>`;
      return;
    }
    target.innerHTML = rows.map((row) => {
      const link = row.open_url
        ? `<a href="${escapeAttribute(row.open_url)}" target="_blank" rel="noreferrer">Open source record</a>`
        : `<span class="kpi-note">No direct image URL exposed in the shared table.</span>`;
      return `
        <article class="source-media-card">
          <h3>${escapeHtml(row.filename || row.title || "Source media record")}</h3>
          <div class="media-meta">${escapeHtml(shortProjectName(getProject(row.project_id)))} | ${escapeHtml(formatDate(row.date))}</div>
          <div class="media-meta">${escapeHtml([row.source_table, row.content_type, row.uploader].filter(Boolean).join(" | "))}</div>
          ${link}
        </article>
      `;
    }).join("");
  }

  function renderHseTable() {
    const incidents = selectedRows(DATA.incident_details || []);
    const observations = selectedRows(DATA.observation_details || []);
    if (!incidents.length && !observations.length) {
      document.getElementById("hse-table").innerHTML = `<div class="empty-state">No incident or observation detail rows were present in the current dashboard period.</div>`;
      return;
    }
    const rows = incidents.map((row) => [
      shortProjectName(getProject(row.project_id)),
      row.incident_number || "",
      row.incident_title || "",
      formatDate(row.event_date)
    ]).concat(observations.map((row) => [
      shortProjectName(getProject(row.project_id)),
      row.observation_number || "",
      row.observation_name || "",
      formatDate(row.due_date)
    ]));
    renderTable("hse-table", ["Project", "No.", "Item", "Date"], rows);
  }

  function renderRiskGantt() {
    const rows = selectedActions().map((item) => ({
      label: item.subject,
      start: item.start,
      end: item.end,
      color: riskColor(item.risk_class),
      detail: [shortProjectName(getProject(item.project_id)), item.category, item.priority].filter(Boolean).join(" | ")
    }));
    if (!rows.length) {
      document.getElementById("risk-gantt").innerHTML = `<div class="empty-state">No executive action records matched the selected filters.</div>`;
      return;
    }
    drawGantt("risk-gantt", rows.slice(0, 26), { fallbackDays: 14 });
  }

  function renderRiskMix() {
    const counts = selectedActions().reduce((acc, item) => {
      acc[item.risk_class || "action"] = (acc[item.risk_class || "action"] || 0) + 1;
      return acc;
    }, {});
    const rows = Object.entries(counts).map(([key, value]) => ({ label: key, value, color: riskColor(key) }));
    drawDonutChart("risk-mix", rows);
  }

  function renderActionTable() {
    const rows = selectedActions().slice(0, 20).map((item) => [
      escapeHtml(shortProjectName(getProject(item.project_id))),
      escapeHtml(item.subject),
      `<span class="pill ${escapeAttribute(item.risk_class)}">${escapeHtml(item.category || item.risk_class)}</span>`,
      escapeHtml(formatDate(item.start))
    ]);
    renderTable("action-table", ["Project", "Item", "Type", "Date"], rows, true);
  }

  function renderCorrespondencePulse() {
    drawPulse("correspondence-pulse", selectedCorrespondence());
  }

  function renderWeeklyNarrative() {
    const narrative = buildNarratives();
    const html = `
      <section class="narrative-section">
        <h3>Summary of Work Completed</h3>
        <p class="source-note">Source: Correspondence tool, Weekly Project Update fields. Fallback: Program Action & Coordination Item when weekly fields are absent.</p>
        ${renderNarrativeCards(narrative.summary, "No weekly summary found.")}
      </section>
      <section class="narrative-section">
        <h3>Current Issues</h3>
        <p class="source-note">Source: Current Issues field from Weekly Project Update. Fallback: latest coordination item subjects and notes.</p>
        ${renderNarrativeCards(narrative.issues, "No current issues found.")}
      </section>
      <section class="narrative-section">
        <h3>Top Wins</h3>
        <p class="source-note">Source: Top 3 Wins field from Weekly Project Update. Fallback: completed or progress-oriented coordination items.</p>
        ${renderNarrativeCards(narrative.wins, "No Top 3 Wins field found.")}
      </section>
      <section class="narrative-section">
        <h3>Top Risks</h3>
        <p class="source-note">Source: Top 3 Risks field from Weekly Project Update. Fallback: high-priority or risk-class coordination items.</p>
        ${renderNarrativeCards(narrative.risks, "No Top 3 Risks field found.")}
      </section>
      <section class="narrative-section">
        <h3>Two Week Lookahead</h3>
        <p class="source-note">Source: 2 Weeks Lookahead field or lookahead attachment from Weekly Project Update.</p>
        ${renderNarrativeCards(narrative.lookahead, "No two week lookahead field found.")}
      </section>
    `;
    ["weekly-narrative", "weekly-narrative-overview"].forEach((id) => {
      const target = document.getElementById(id);
      if (target) target.innerHTML = html;
    });
  }

  function buildNarratives() {
    const rows = rowsFor("weekly_updates");
    const curated = {
      summary: fieldCards(rows, "latest_weekly_project_update_text", "Weekly summary"),
      issues: fieldCards(rows, "current_issues", "Current issue"),
      wins: fieldCards(rows, "top_wins", "Top win"),
      risks: fieldCards(rows, "top_risks", "Top risk"),
      lookahead: fieldCards(rows, "two_week_lookahead", "Two week lookahead", "asc")
    };
    if (curated.summary.length || curated.issues.length || curated.wins.length || curated.risks.length || curated.lookahead.length) {
      return curated;
    }

    const actions = selectedActions();
    const highPriority = actions.filter((item) => String(item.priority || "").toLowerCase().startsWith("high") || item.risk_class === "risk");
    const mediumPriority = actions.filter((item) => String(item.priority || "").toLowerCase().startsWith("med"));
    const winsSource = actions.filter((item) => {
      const text = cleanText(`${item.subject || ""} ${item.description || ""} ${item.category || ""}`).toLowerCase();
      return text.includes("done") || text.includes("complete") || text.includes("ready") || text.includes("uploaded") || text.includes("win") || item.risk_class === "win";
    });
    return {
      summary: sortNarrativeCards(actions.slice(0, 6).map((item) => actionCard(item, "Summary"))),
      issues: sortNarrativeCards(actions.slice(0, 6).map((item) => actionCard(item, "Issue"))),
      wins: sortNarrativeCards((winsSource.length ? winsSource : actions).slice(0, 4).map((item) => actionCard(item, "Win"))),
      risks: sortNarrativeCards((highPriority.length ? highPriority : actions).slice(0, 4).map((item) => actionCard(item, "Risk"))),
      lookahead: sortNarrativeCards(uniqueActionItems([...highPriority, ...mediumPriority, ...actions]).slice(0, 6).map((item) => actionCard(item, "Lookahead")), "asc")
    };
  }

  function fieldCards(rows, field, fallbackTitle, direction) {
    const cards = [];
    rows.forEach((row) => {
      splitNarrative(row[field]).forEach((text, index) => {
        cards.push({
          title: index === 0 ? cleanText(row.latest_correspondence_subject) || fallbackTitle : fallbackTitle,
          body: text,
          projectId: row.project_id,
          start: row.latest_correspondence_date || row.snapshot_generated_at,
          end: row.latest_correspondence_date || row.snapshot_generated_at,
          category: fallbackTitle,
          priority: "",
          source: "Weekly Project Update"
        });
      });
    });
    return sortNarrativeCards(cards, direction);
  }

  function actionCard(item, fallbackCategory) {
    return {
      title: cleanText(item.subject) || fallbackCategory,
      body: cleanText(item.meeting_notes || item.description || item.category || ""),
      projectId: item.project_id,
      start: item.start,
      end: item.end || item.start,
      category: cleanText(item.category || fallbackCategory),
      priority: cleanText(item.priority || ""),
      phase: cleanText(item.phase || ""),
      source: "Program Action & Coordination Item",
      riskClass: item.risk_class || "action"
    };
  }

  function sortNarrativeCards(cards, direction) {
    const sorted = (cards || []).filter((card) => cleanText(card.title) || cleanText(card.body));
    sorted.sort((left, right) => {
      const leftTime = (parseDate(left.end) || parseDate(left.start) || new Date(0)).getTime();
      const rightTime = (parseDate(right.end) || parseDate(right.start) || new Date(0)).getTime();
      return direction === "asc" ? leftTime - rightTime : rightTime - leftTime;
    });
    return sorted;
  }

  function uniqueActionItems(items) {
    const seen = new Set();
    return items.filter((item) => {
      const key = String(item.id || `${item.project_id}-${item.subject}-${item.start}`);
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  function renderNarrativeCards(cards, emptyMessage) {
    const cleanCards = (cards || []).filter((card) => cleanText(card.title) || cleanText(card.body));
    if (!cleanCards.length) return `<div class="empty-state">${escapeHtml(emptyMessage)}</div>`;
    return `
      <div class="narrative-card-grid">
        ${cleanCards.map((card) => `
          <article class="narrative-item-card">
            <div class="narrative-card-topline">
              <span class="pill ${escapeAttribute(card.riskClass || "")}">${escapeHtml(card.category || "Update")}</span>
              ${card.priority ? `<span class="narrative-priority">${escapeHtml(card.priority)}</span>` : ""}
            </div>
            <h4>${escapeHtml(card.title)}</h4>
            <div class="narrative-card-meta">
              <span>${escapeHtml(formatDateRange(card.start, card.end))}</span>
              <span>${escapeHtml(shortProjectName(getProject(card.projectId)))}</span>
              ${card.phase ? `<span>${escapeHtml(card.phase)}</span>` : ""}
            </div>
            ${card.body ? `<p>${escapeHtml(card.body)}</p>` : ""}
            <div class="source-note">Source: ${escapeHtml(card.source || "Dashboard data")}</div>
          </article>
        `).join("")}
      </div>
    `;
  }

  function splitNarrative(text) {
    return cleanText(text)
      .split(/\n{2,}|\n(?=\d{1,2}\/\d{1,2}|\w{3}\s+\d{1,2}|[-•])/)
      .map((item) => item.replace(/^[-•]\s*/, "").trim())
      .filter(Boolean);
  }

  function renderCorrespondenceTable() {
    const rows = selectedCorrespondence().slice(0, 28).map((row) => [
      formatDate(row.date),
      shortProjectName(getProject(row.project_id)),
      row.type || "",
      row.subject || ""
    ]);
    renderTable("correspondence-table", ["Date", "Project", "Type", "Subject"], rows);
  }

  function renderInspector() {
    const selected = selectedProjects();
    const target = document.getElementById("project-inspector");
    if (state.selectionMode === "multi" && selected.length > 1) {
      target.innerHTML = `<div class="inspector-stack">${selected.map((project) => renderInspectorCard(project, true)).join("")}</div>`;
      return;
    }
    const project = getProject(state.inspectedId) || selected[0] || projects[0];
    if (!project) return;
    target.innerHTML = renderInspectorCard(project, false);
  }

  function renderInspectorCard(project, compact) {
    const financial = rowForProject("financials", project.project_id);
    const rfi = rowForProject("rfis", project.project_id);
    const sub = rowForProject("submittals", project.project_id);
    const committedRatio = num(financial.commitment_value) / Math.max(1, num(financial.approved_budget));
    return `
      <div class="inspector-card ${compact ? "compact" : ""}">
        <div>
          <h3>${escapeHtml(project.project_display_name || project.project_name)}</h3>
          <span class="pill">${escapeHtml(project.project_stage_name || "Stage unavailable")}</span>
        </div>
        <p class="project-description">${escapeHtml(project.project_description || "No project description available.")}</p>
        <div>
          <div class="kpi-note">Committed vs Approved Budget</div>
          <div class="progress-track"><span style="width:${Math.min(100, committedRatio * 100)}%"></span></div>
        </div>
        <div class="metric-list">
          <div class="metric-row"><span>Approved Budget</span><strong>${fmtMoney(financial.approved_budget)}</strong><span></span><span></span></div>
          <div class="metric-row"><span>Invoiced</span><strong>${fmtMoney(financial.invoiced_to_date)}</strong><span></span><span></span></div>
          <div class="metric-row"><span>Open RFIs</span><strong>${fmtNumber(rfi.open_rfis)}</strong><span>Overdue</span><strong>${fmtNumber(rfi.overdue_rfis)}</strong></div>
          <div class="metric-row"><span>Open Submittals</span><strong>${fmtNumber(sub.open_submittals)}</strong><span>Overdue</span><strong>${fmtNumber(sub.overdue_submittals)}</strong></div>
        </div>
      </div>
    `;
  }

  function openDetail(kind, projectId) {
    const project = getProject(projectId) || selectedProjects()[0] || projects[0];
    if (!project) return;
    state.inspectedId = project.project_id;
    const drawer = document.getElementById("detail-drawer");
    const title = document.getElementById("drawer-title");
    const eyebrow = document.getElementById("drawer-eyebrow");
    const body = document.getElementById("drawer-body");
    const isRfi = kind === "rfis";
    const label = isRfi ? "RFIs" : "Submittals";
    const summary = rowForProject(isRfi ? "rfis" : "submittals", project.project_id);
    const detailRows = detailRowsFor(kind, project.project_id);
    const metrics = isRfi
      ? [
          ["Total", summary.total_rfis],
          ["Open", summary.open_rfis],
          ["Overdue", summary.overdue_rfis],
          ["Loaded Details", detailRows.length]
        ]
      : [
          ["Total", summary.total_submittals],
          ["Open", summary.open_submittals],
          ["Overdue", summary.overdue_submittals],
          ["Loaded Details", detailRows.length]
        ];
    eyebrow.textContent = `${shortProjectName(project)} Detail`;
    title.textContent = label;
    body.innerHTML = `
      <div class="detail-grid">
        ${metrics.map(([name, value]) => `<div class="detail-metric"><span class="kpi-note">${escapeHtml(name)}</span><strong>${fmtNumber(value)}</strong></div>`).join("")}
      </div>
      ${detailRows.length ? detailRows.slice(0, 28).map((row) => isRfi ? rfiRecord(row) : submittalRecord(row)).join("") : detailEmptyState(label)}
    `;
    drawer.hidden = false;
  }

  function closeDetail() {
    const drawer = document.getElementById("detail-drawer");
    if (drawer) drawer.hidden = true;
  }

  function detailRowsFor(kind, projectId) {
    const table = kind === "rfis" ? DATA.rfi_details : DATA.submittal_details;
    return (table || [])
      .filter((row) => row.project_id === projectId)
      .filter((row) => dateInWindow(row.date || row.due_date || row.updated_at || row.created_at || row.initiated_at || row.final_due_date))
      .sort((a, b) => String(b.date || b.updated_at || b.created_at || "").localeCompare(String(a.date || a.updated_at || a.created_at || "")));
  }

  function rfiRecord(row) {
    const responses = (DATA.rfi_responses || []).filter((item) => sameId(item.rfi_header_id, row.rfi_header_id));
    const assignees = (DATA.rfi_assignees || [])
      .filter((item) => sameId(item.rfi_header_id, row.rfi_header_id))
      .map((item) => item.assignee_name || item.name)
      .filter(Boolean);
    const latestResponse = responses[0] && (responses[0].response || responses[0].plain_text_response);
    return `
      <article class="detail-record">
        <div>
          <h3>${escapeHtml(cleanText(row.subject || row.title || `RFI ${row.number || row.rfi_header_id || ""}`))}</h3>
          <span class="pill">${escapeHtml(row.status || "Status unavailable")}</span>
          ${row.overdue ? `<span class="pill risk">Overdue</span>` : ""}
        </div>
        <div class="kpi-note">No. ${escapeHtml(row.number || "")} | Due ${escapeHtml(formatDate(row.due_date))} | ${responses.length} response records</div>
        ${assignees.length ? `<p class="kpi-note">Assignees: ${escapeHtml(assignees.slice(0, 4).join(", "))}</p>` : ""}
        ${row.question ? `<p>${escapeHtml(cleanText(row.question))}</p>` : ""}
        ${latestResponse ? `<p class="kpi-note">Latest response: ${escapeHtml(trimLabel(cleanText(latestResponse), 260))}</p>` : ""}
        ${row.rfi_resource_url ? `<a class="mini-button" href="${escapeAttribute(sourceUrl(row.rfi_resource_url))}" target="_blank" rel="noreferrer">Open RFI</a>` : ""}
      </article>
    `;
  }

  function submittalRecord(row) {
    const approvers = (DATA.submittal_approvers || [])
      .filter((item) => sameId(item.submittal_log_id, row.submittal_log_id))
      .map((item) => item.approver_name || item.name || item.login)
      .filter(Boolean);
    return `
      <article class="detail-record">
        <div>
          <h3>${escapeHtml(cleanText(row.title || row.description || `Submittal ${row.number || row.submittal_log_id || ""}`))}</h3>
          <span class="pill">${escapeHtml(row.status || row.status_category || "Status unavailable")}</span>
          ${row.overdue ? `<span class="pill risk">Overdue</span>` : ""}
        </div>
        <div class="kpi-note">No. ${escapeHtml(row.number || "")} | Due ${escapeHtml(formatDate(row.final_due_date || row.due_date))} | ${approvers.length} approver records</div>
        ${approvers.length ? `<p class="kpi-note">Approvers: ${escapeHtml(approvers.slice(0, 4).join(", "))}</p>` : ""}
        ${row.submittal_resource_url ? `<a class="mini-button" href="${escapeAttribute(sourceUrl(row.submittal_resource_url))}" target="_blank" rel="noreferrer">Open Submittal</a>` : ""}
      </article>
    `;
  }

  function detailEmptyState(label) {
    return `<div class="empty-state">${escapeHtml(label)} summary metrics are loaded, but detail rows are not present in the current local bundle. Run the Delta sync with a larger detail limit to deepen this drawer.</div>`;
  }

  function sourceUrl(value) {
    const text = String(value || "");
    if (!text) return "";
    if (/^https?:\/\//i.test(text)) return text;
    return `https://app.procore.com${text.startsWith("/") ? "" : "/"}${text}`;
  }

  function sameId(left, right) {
    return left != null && right != null && String(left) === String(right);
  }

  function showTooltip(text, event) {
    if (!tooltip || !text) return;
    tooltip.textContent = text;
    tooltip.hidden = false;
    moveTooltip(event);
  }

  function moveTooltip(event) {
    if (!tooltip) return;
    tooltip.style.left = `${Math.min(window.innerWidth - 340, event.clientX + 14)}px`;
    tooltip.style.top = `${Math.min(window.innerHeight - 120, event.clientY + 14)}px`;
  }

  function hideTooltip() {
    if (tooltip) tooltip.hidden = true;
  }

  function drawGroupedBars(targetId, rows, series, options) {
    const target = document.getElementById(targetId);
    if (!rows.length) {
      target.innerHTML = `<div class="empty-state">No rows for the selected projects.</div>`;
      return;
    }
    if (state.selectionMode === "multi" && rows.length > 1) {
      drawVerticalGroupedBars(target, rows, series, options);
      return;
    }
    const width = Math.max(target.clientWidth || 760, 620);
    const height = target.classList.contains("tall") ? 370 : 300;
    const margin = { top: 24, right: 96, bottom: 56, left: 142 };
    const innerW = width - margin.left - margin.right;
    const innerH = height - margin.top - margin.bottom;
    const maxValue = Math.max(...rows.flatMap((row) => series.map((s) => num(row[s.key]))), 1);
    const rowH = innerH / rows.length;
    const gap = 5;
    const groupH = Math.min(rowH - 8, series.length * 18 + (series.length - 1) * gap);
    const barH = Math.min(16, groupH / series.length - gap);
    const valueFormat = options.valueFormat || fmtNumber;

    let html = `<svg class="svg-chart" viewBox="0 0 ${width} ${height}" role="img">`;
    html += `<g transform="translate(${margin.left},${margin.top})">`;
    rows.forEach((row, i) => {
      const y = i * rowH;
      html += `<text class="axis-label" x="-12" y="${y + rowH / 2 + 4}" text-anchor="end">${escapeSvg(trimLabel(options.label(row), 21))}</text>`;
      series.forEach((s, j) => {
        const value = num(row[s.key]);
        const barW = value / maxValue * innerW;
        const by = y + (rowH - groupH) / 2 + j * (barH + gap);
        const drillAttrs = options.drill ? ` data-drill="${escapeAttribute(options.drill)}" data-project-id="${row.project_id || ""}"` : "";
        const tooltipText = `${options.label(row)} | ${s.label}: ${valueFormat(value)}`;
        html += `<rect class="${options.drill ? "clickable" : ""}"${drillAttrs} data-tooltip="${escapeAttribute(tooltipText)}" x="0" y="${by}" width="${barW}" height="${barH}" rx="4" fill="${s.color}"></rect>`;
        html += `<text class="chart-value-label" x="${barW + 8}" y="${by + barH - 4}">${escapeSvg(valueFormat(value))}</text>`;
      });
    });
    series.forEach((s, index) => {
      const x = index * 130;
      html += `<g transform="translate(${x},${innerH + 34})"><rect width="10" height="10" rx="2" fill="${s.color}"></rect><text class="tiny-label" x="16" y="10">${escapeSvg(s.label)}</text></g>`;
    });
    html += `</g></svg>`;
    target.innerHTML = html;
  }

  function drawVerticalGroupedBars(target, rows, series, options) {
    const width = Math.max(target.clientWidth || 760, rows.length * 126 + 120);
    const height = target.classList.contains("tall") ? 380 : 315;
    const margin = { top: 30, right: 24, bottom: 78, left: 58 };
    const innerW = width - margin.left - margin.right;
    const innerH = height - margin.top - margin.bottom;
    const maxValue = Math.max(...rows.flatMap((row) => series.map((s) => num(row[s.key]))), 1);
    const groupW = innerW / rows.length;
    const gap = 6;
    const barW = Math.max(10, Math.min(26, (groupW - 24) / series.length - gap));
    const valueFormat = options.valueFormat || fmtNumber;
    let html = `<svg class="svg-chart wide-chart" style="min-width:${width}px" viewBox="0 0 ${width} ${height}" role="img">`;
    html += `<g transform="translate(${margin.left},${margin.top})">`;
    html += `<line x1="0" x2="${innerW}" y1="${innerH}" y2="${innerH}" stroke="${COLORS.line}"></line>`;
    rows.forEach((row, i) => {
      const center = i * groupW + groupW / 2;
      const totalBarW = series.length * barW + (series.length - 1) * gap;
      html += `<text class="axis-label" x="${center}" y="${innerH + 30}" text-anchor="middle">${escapeSvg(trimLabel(options.label(row), 15))}</text>`;
      series.forEach((s, j) => {
        const value = num(row[s.key]);
        const h = value / maxValue * innerH;
        const x = center - totalBarW / 2 + j * (barW + gap);
        const y = innerH - h;
        const drillAttrs = options.drill ? ` data-drill="${escapeAttribute(options.drill)}" data-project-id="${row.project_id || ""}"` : "";
        const tooltipText = `${options.label(row)} | ${s.label}: ${valueFormat(value)}`;
        html += `<rect class="${options.drill ? "clickable" : ""}"${drillAttrs} data-tooltip="${escapeAttribute(tooltipText)}" x="${x}" y="${y}" width="${barW}" height="${Math.max(1, h)}" rx="4" fill="${s.color}"></rect>`;
        html += `<text class="chart-value-label" x="${x + barW / 2}" y="${Math.max(12, y - 7)}" text-anchor="middle">${escapeSvg(valueFormat(value))}</text>`;
      });
    });
    series.forEach((s, index) => {
      const x = index * 130;
      html += `<g transform="translate(${x},${height - margin.top - 20})"><rect width="10" height="10" rx="2" fill="${s.color}"></rect><text class="tiny-label" x="16" y="10">${escapeSvg(s.label)}</text></g>`;
    });
    html += `</g></svg>`;
    target.innerHTML = html;
  }

  function drawGantt(targetId, rows, options) {
    const target = document.getElementById(targetId);
    if (!rows.length) {
      target.innerHTML = `<div class="empty-state">No timeline rows available.</div>`;
      return;
    }
    const parsed = rows.map((row) => {
      const start = parseDate(row.start) || new Date();
      const end = parseDate(row.end) || new Date(start.getTime() + (options.fallbackDays || 30) * 86400000);
      return { ...row, startDate: start, endDate: end < start ? start : end };
    });
    const min = new Date(Math.min(...parsed.map((row) => row.startDate.getTime())));
    const max = new Date(Math.max(...parsed.map((row) => row.endDate.getTime())));
    const margin = { top: 24, right: 24, bottom: 40, left: 182 };
    const spanDays = Math.max(14, Math.ceil((max - min) / 86400000));
    const tickCount = Math.ceil(spanDays / 14) + 1;
    const width = Math.max(target.clientWidth || 860, parsed.length > 14 ? 1120 : 860, tickCount * 70 + margin.left + margin.right);
    const height = Math.max(target.classList.contains("xlarge") ? 500 : 280, parsed.length * 28 + 94);
    const innerW = width - margin.left - margin.right;
    const rowH = 26;
    const span = Math.max(1, max - min);
    const xScale = (date) => margin.left + (date - min) / span * innerW;
    let html = `<svg class="svg-chart wide-chart" style="min-width:${width}px" viewBox="0 0 ${width} ${height}">`;
    html += `<line x1="${margin.left}" x2="${margin.left + innerW}" y1="${height - margin.bottom}" y2="${height - margin.bottom}" stroke="${COLORS.line}"></line>`;
    for (let tick = new Date(min.getTime()), index = 0; tick <= max; tick = new Date(tick.getTime() + 14 * 86400000), index += 1) {
      const x = xScale(tick);
      html += `<line x1="${x}" x2="${x}" y1="${margin.top - 8}" y2="${height - margin.bottom}" stroke="${COLORS.line}" opacity="${index % 2 === 0 ? ".55" : ".32"}"></line>`;
      html += `<text class="tiny-label" x="${x}" y="${height - 14}" text-anchor="middle">${formatShortDate(tick)}</text>`;
    }
    parsed.forEach((row, index) => {
      const y = margin.top + index * rowH;
      const x = xScale(row.startDate);
      const w = Math.max(10, xScale(row.endDate) - x);
      html += `<text class="axis-label" x="${margin.left - 10}" y="${y + 15}" text-anchor="end">${escapeSvg(trimLabel(row.label, 28))}</text>`;
      html += `<rect data-tooltip="${escapeAttribute(`${row.label} | ${row.detail || ""} | ${formatDate(row.startDate)} to ${formatDate(row.endDate)}`)}" x="${x}" y="${y}" width="${w}" height="16" rx="5" fill="${row.color || COLORS.teal}" opacity=".92"></rect>`;
      if (row.detail) html += `<text class="tiny-label" x="${Math.min(x + w + 8, margin.left + innerW - 120)}" y="${y + 13}">${escapeSvg(trimLabel(row.detail, 24))}</text>`;
    });
    html += `</svg>`;
    target.innerHTML = html;
  }

  function drawDonutChart(targetId, rows) {
    const target = document.getElementById(targetId);
    if (!rows.length) {
      target.innerHTML = `<div class="empty-state">No category rows available.</div>`;
      return;
    }
    const total = rows.reduce((acc, row) => acc + row.value, 0) || 1;
    let start = 0;
    const segments = rows.map((row) => {
      const degrees = row.value / total * 360;
      const segment = `${row.color} ${start}deg ${start + degrees}deg`;
      start += degrees;
      return segment;
    });
    target.innerHTML = `
      <div style="display:grid;place-items:center;min-height:260px">
        <div class="donut-ring" style="width:180px;height:180px;background:conic-gradient(${segments.join(",")})"><span style="width:112px;height:112px;font-size:30px">${total}</span></div>
        <div class="metric-list" style="width:100%;margin-top:14px">
          ${rows.map((row) => `<div class="metric-row"><span><span class="pill ${escapeAttribute(row.label)}">${escapeHtml(row.label)}</span></span><strong>${row.value}</strong><span></span><span></span></div>`).join("")}
        </div>
      </div>
    `;
  }

  function drawPulse(targetId, rows) {
    const target = document.getElementById(targetId);
    if (!rows.length) {
      target.innerHTML = `<div class="empty-state">No correspondence records in the selected window.</div>`;
      return;
    }
    const parsed = rows.map((row, index) => ({ ...row, dateObj: parseDate(row.date), index })).filter((row) => row.dateObj);
    if (!parsed.length) {
      target.innerHTML = `<div class="empty-state">Correspondence rows do not have usable dates.</div>`;
      return;
    }
    const types = [...new Set(parsed.map((row) => row.type || "Other"))].slice(0, 8);
    const min = new Date(Math.min(...parsed.map((row) => row.dateObj.getTime())));
    const max = new Date(Math.max(...parsed.map((row) => row.dateObj.getTime())));
    const width = target.clientWidth || 820;
    const height = 370;
    const margin = { top: 28, right: 20, bottom: 46, left: 138 };
    const innerW = width - margin.left - margin.right;
    const rowH = (height - margin.top - margin.bottom) / types.length;
    const span = Math.max(1, max - min);
    const xScale = (date) => margin.left + (date - min) / span * innerW;
    let html = `<svg class="svg-chart" viewBox="0 0 ${width} ${height}">`;
    types.forEach((type, i) => {
      const y = margin.top + i * rowH + rowH / 2;
      html += `<text class="axis-label" x="${margin.left - 10}" y="${y + 4}" text-anchor="end">${escapeSvg(trimLabel(type, 22))}</text>`;
      html += `<line x1="${margin.left}" x2="${margin.left + innerW}" y1="${y}" y2="${y}" stroke="${COLORS.line}"></line>`;
    });
    parsed.forEach((row) => {
      const i = types.indexOf(row.type || "Other");
      if (i < 0) return;
      const y = margin.top + i * rowH + rowH / 2;
      const x = xScale(row.dateObj);
      const tip = `${shortProjectName(getProject(row.project_id))} | ${formatDate(row.date)} | ${row.type || "Other"} | ${cleanText(row.subject || "")}`;
      html += `<circle class="clickable" data-tooltip="${escapeAttribute(tip)}" cx="${x}" cy="${y}" r="6" fill="${projectColor(projects.findIndex((p) => p.project_id === row.project_id))}" opacity=".84"></circle>`;
    });
    html += `<text class="tiny-label" x="${margin.left}" y="${height - 16}">${formatDate(min)}</text>`;
    html += `<text class="tiny-label" x="${margin.left + innerW}" y="${height - 16}" text-anchor="end">${formatDate(max)}</text>`;
    html += `</svg>`;
    target.innerHTML = html;
  }

  function renderTable(targetId, headers, rows, allowHtml) {
    const target = document.getElementById(targetId);
    if (!rows.length) {
      target.innerHTML = `<div class="empty-state">No records for the selected filters.</div>`;
      return;
    }
    target.innerHTML = `
      <div class="table-row header">${headers.map((h) => `<span>${escapeHtml(h)}</span>`).join("")}</div>
      ${rows.map((row) => `<div class="table-row">${row.map((cell) => `<span>${allowHtml ? cell : escapeHtml(cell)}</span>`).join("")}</div>`).join("")}
    `;
  }

  function filteredPayload() {
    return {
      meta: DATA.meta,
      selected_project_ids: [...state.selectedIds],
      projects: selectedProjects(),
      totals: totals(),
      actions: selectedActions(),
      correspondence: selectedCorrespondence()
    };
  }

  function rowForProject(tableName, projectId) {
    return (DATA[tableName] || []).find((row) => row.project_id === projectId) || {};
  }

  function getProject(projectId) {
    return projects.find((project) => project.project_id === projectId) || null;
  }

  function shortProjectName(project) {
    if (!project) return "Unassigned";
    return String(project.project_display_name || project.project_name || project.project_id).replace(/^00-/, "");
  }

  function projectColor(index) {
    const colors = [COLORS.gold, COLORS.teal, COLORS.blue, COLORS.green, COLORS.orange, COLORS.navy];
    return colors[Math.max(0, index) % colors.length];
  }

  function riskColor(key) {
    return { risk: COLORS.red, decision: COLORS.orange, win: COLORS.green, action: COLORS.blue }[key] || COLORS.gray;
  }

  function sum(rows, key) {
    return rows.reduce((acc, row) => acc + num(row[key]), 0);
  }

  function uniqueValues(values) {
    const seen = new Set();
    return values
      .map(cleanText)
      .filter(Boolean)
      .filter((value) => {
        const key = value.toLowerCase();
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
  }

  function num(value) {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function parseDate(value) {
    if (!value) return null;
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }

  function dateInWindow(value) {
    const date = parseDate(value);
    if (!date) return true;
    const start = parseDate(state.dateStart);
    const end = parseDate(state.dateEnd);
    if (start && date < start) return false;
    if (end) {
      const endOfDay = new Date(end.getTime() + 86400000 - 1);
      if (date > endOfDay) return false;
    }
    return true;
  }

  function toInputDate(value) {
    const date = parseDate(value);
    return date ? date.toISOString().slice(0, 10) : "";
  }

  function formatDate(value) {
    const date = value instanceof Date ? value : parseDate(value);
    if (!date) return "";
    return date.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" });
  }

  function formatDateRange(startValue, endValue) {
    const start = parseDate(startValue);
    const end = parseDate(endValue);
    if (start && end && start.toDateString() !== end.toDateString()) {
      return `${formatDate(start)} to ${formatDate(end)}`;
    }
    if (start) return formatDate(start);
    if (end) return formatDate(end);
    return "Date unavailable";
  }

  function formatShortDate(value) {
    const date = value instanceof Date ? value : parseDate(value);
    if (!date) return "";
    return date.toLocaleDateString(undefined, { month: "short", day: "numeric" });
  }

  function formatDateTime(value) {
    const date = parseDate(value);
    if (!date) return "";
    return date.toLocaleString(undefined, { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" });
  }

  function fmtMoney(value) {
    return new Intl.NumberFormat(undefined, { style: "currency", currency: "USD", maximumFractionDigits: 0 }).format(num(value));
  }

  function fmtCompactMoney(value) {
    return new Intl.NumberFormat(undefined, { style: "currency", currency: "USD", notation: "compact", maximumFractionDigits: 1 }).format(num(value));
  }

  function fmtNumber(value) {
    return new Intl.NumberFormat(undefined, { maximumFractionDigits: 0 }).format(num(value));
  }

  function fmtPct(value) {
    return new Intl.NumberFormat(undefined, { style: "percent", maximumFractionDigits: 1 }).format(num(value));
  }

  function trimLabel(value, length) {
    const text = String(value || "");
    return text.length > length ? `${text.slice(0, length - 1)}...` : text;
  }

  function cleanText(value) {
    const raw = String(value == null ? "" : value)
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<\/p>/gi, "\n")
      .replace(/<[^>]+>/g, "")
      .replace(/\r\n?/g, "\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
    const textarea = document.createElement("textarea");
    textarea.innerHTML = raw;
    return textarea.value.replace(/\u00a0/g, " ").trim();
  }

  function escapeHtml(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function escapeSvg(value) {
    return escapeHtml(value);
  }

  function escapeAttribute(value) {
    return escapeHtml(value).replace(/`/g, "&#096;");
  }

  window.addEventListener("resize", () => renderAll());
  init();
})();
