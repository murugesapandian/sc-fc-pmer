const statusBadge = document.getElementById("status-badge");
const lastReportInfo = document.getElementById("last-report-info");
const generateBtn = document.getElementById("generate-btn");
const forceBtn = document.getElementById("force-btn");
const refreshBtn = document.getElementById("refresh-btn");
const actionMessage = document.getElementById("action-message");
const reportsList = document.getElementById("reports-list");

const glanceSection = document.getElementById("glance");
const glanceDot = document.getElementById("glance-dot");
const glanceStatusWord = document.getElementById("glance-status-word");
const glancePhase = document.getElementById("glance-phase");
const glanceTime = document.getElementById("glance-time");
const statDays = document.getElementById("stat-days");
const statPct = document.getElementById("stat-pct");
const statPctBar = document.getElementById("stat-pct-bar");
const statRaid = document.getElementById("stat-raid");
const statBudget = document.getElementById("stat-budget");

const STATUS_TONE = { GREEN: "green", AMBER: "amber", RED: "red" };
const POLL_INTERVAL_MS = 15000;

// Tracks the last report id this page has actually rendered, so the silent
// background poll can tell "nothing changed" from "a new report showed up"
// and only touch the DOM in the second case — per spec, the page must not
// visibly refresh just because a poll tick happened.
let knownReportId = undefined; // undefined = not loaded yet, null = loaded but no reports exist

function fmtDateTime(iso) {
  try {
    const d = new Date(iso);
    return d.toLocaleString(undefined, {
      year: "numeric", month: "short", day: "numeric", hour: "2-digit", minute: "2-digit",
    });
  } catch {
    return iso;
  }
}

function triggerTag(trigger) {
  const label = { manual: "Manual", scheduled: "Scheduled", forced: "Forced" }[trigger] || trigger;
  return `<span class="trigger-tag trigger-${trigger}">${label}</span>`;
}

function showMessage(text, kind) {
  actionMessage.textContent = text;
  actionMessage.className = `action-message ${kind}`;
  actionMessage.hidden = false;
}

function renderGlance(kpi, generatedAt) {
  if (!kpi) {
    glanceSection.hidden = true;
    return;
  }
  glanceSection.hidden = false;
  const tone = STATUS_TONE[kpi.status] || "";
  glanceDot.className = `status-dot ${tone}`;
  glanceStatusWord.textContent = kpi.status || "—";
  glancePhase.textContent = kpi.phase ? `Current phase: ${kpi.phase}` : "";
  glanceTime.textContent = fmtDateTime(generatedAt);

  statDays.textContent = kpi.days_to_go_live ?? "—";

  const pct = kpi.pct_complete != null ? Math.round(kpi.pct_complete * 100) : null;
  statPct.textContent = pct != null ? `${pct}%` : "—";
  statPctBar.style.width = pct != null ? `${pct}%` : "0%";

  statRaid.textContent = kpi.open_raid ?? "—";

  if (kpi.budget_variance_pct != null) {
    const v = kpi.budget_variance_pct * 100;
    statBudget.textContent = `${v >= 0 ? "+" : ""}${v.toFixed(1)}%`;
  } else {
    statBudget.textContent = "—";
  }
}

async function loadStatus() {
  statusBadge.textContent = "Checking…";
  statusBadge.className = "badge badge-neutral";
  try {
    const res = await fetch("/api/status");
    if (!res.ok) throw new Error(await res.text());
    const data = await res.json();

    if (data.up_to_date) {
      statusBadge.textContent = "✓ Up to date";
      statusBadge.className = "badge badge-green";
      generateBtn.disabled = true;
      generateBtn.title = "No changes to the workbook or manual inputs since the last report.";
    } else {
      statusBadge.textContent = "● Changes pending";
      statusBadge.className = "badge badge-amber";
      generateBtn.disabled = false;
      generateBtn.title = "";
    }

    if (data.last_report) {
      lastReportInfo.innerHTML =
        `Last generated report: <strong>${data.last_report.filename}</strong> ` +
        `on ${fmtDateTime(data.last_report.generated_at)} (${triggerTag(data.last_report.trigger)})`;
      renderGlance(data.last_report.kpi, data.last_report.generated_at);
      knownReportId = data.last_report.id;
    } else {
      lastReportInfo.textContent = "No report has been generated yet.";
      generateBtn.disabled = false;
      renderGlance(null);
      knownReportId = null;
    }
  } catch (err) {
    statusBadge.textContent = "Error";
    statusBadge.className = "badge badge-red";
    lastReportInfo.textContent = "Could not reach the report service: " + err.message;
    renderGlance(null);
  }
}

async function loadReports() {
  try {
    const res = await fetch("/api/reports");
    const reports = await res.json();
    if (!reports.length) {
      reportsList.innerHTML = `<div class="empty-state muted">No reports generated yet — click <strong>Generate report</strong> above to create the first one.</div>`;
      return;
    }
    reportsList.innerHTML = reports.map(r => `
      <div class="report-row">
        <div class="report-main">
          <div class="report-icon">📊</div>
          <div class="report-meta">
            <div class="report-filename">${r.filename}</div>
            <div class="report-sub">${fmtDateTime(r.generated_at)} &middot; ${triggerTag(r.trigger)}</div>
          </div>
        </div>
        <a class="download-link" href="/api/reports/${r.id}/download">Download</a>
      </div>
    `).join("");
  } catch (err) {
    reportsList.innerHTML = `<div class="empty-state muted">Failed to load report history.</div>`;
  }
}

async function generate(force) {
  const btn = force ? forceBtn : generateBtn;
  const originalHtml = btn.innerHTML;
  btn.disabled = true;
  generateBtn.disabled = true;
  btn.innerHTML = `<span class="spinner"></span> Generating…`;
  actionMessage.hidden = true;

  try {
    const res = await fetch(`/api/generate?force=${force}`, { method: "POST" });
    const data = await res.json();
    if (data.generated) {
      showMessage(`✓ Generated ${data.report.filename}. Review it before sharing with executives.`, "success");
    } else {
      showMessage(data.reason, "info");
    }
    await loadStatus();
    await loadReports();
  } catch (err) {
    showMessage("Generation failed: " + err.message, "error");
  } finally {
    btn.innerHTML = originalHtml;
    await loadStatus();
  }
}

async function refreshAll() {
  await loadStatus();
  await loadReports();
}

// Silent background poll: the server auto-checks the workbook on its own
// timer and writes a new report only when something actually changed. This
// just asks "is there a report I haven't rendered yet" — if not, it touches
// nothing at all (no re-render, no flicker); if so, it's the one case that
// refreshes the page, plus a small toast so it's obvious *why* the numbers
// just moved.
async function checkForNewReport() {
  try {
    const res = await fetch("/api/status");
    if (!res.ok) return;
    const data = await res.json();
    const latestId = data.last_report ? data.last_report.id : null;
    if (knownReportId !== undefined && latestId !== knownReportId) {
      await refreshAll();
      if (data.last_report) {
        showMessage(`New report generated automatically: ${data.last_report.filename}`, "success");
      }
    }
  } catch {
    // Transient failure — leave the page as-is, the next tick will retry.
  }
}

generateBtn.addEventListener("click", () => generate(false));
forceBtn.addEventListener("click", () => generate(true));
refreshBtn.addEventListener("click", refreshAll);

// Smart refresh: the server re-checks disk on every call and drops any
// history entry whose file was deleted outside the app, so re-hitting the
// API whenever the user comes back to this tab is enough to make a manually
// deleted report disappear on its own — no manual refresh click needed.
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "visible") refreshAll();
});
window.addEventListener("focus", refreshAll);

setInterval(checkForNewReport, POLL_INTERVAL_MS);

refreshAll();
