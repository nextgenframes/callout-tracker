import React, { useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import "./styles.css";

const shifts = ["AM Shift", "PM Shift", "Overnight Shift"];
const days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const today = new Date().toISOString().slice(0, 10);
const todayDay = days[new Date().getDay()];
const workforceStatuses = ["Scheduled", "Present", "Called Out", "Late", "No Show", "Training", "Off Shift"];
const coverageStatuses = ["Pending", "Coverage Needed", "Covered", "Closed"];
const reasons = ["Sick", "Emergency", "Transport", "Family", "Other"];
const shiftStart = { "AM Shift": 6, "PM Shift": 14, "Overnight Shift": 22 };

const starterContractors = [
  makeContractor("Jordan Lee", "Northline AV", "AM Shift", ["Mon", "Tue", "Wed", "Thu", "Fri"], "Present", 97, 0, 1, 22),
  makeContractor("Maya Santos", "StageOps", "PM Shift", ["Tue", "Wed", "Thu", "Fri", "Sat"], "Present", 93, 1, 2, 14),
  makeContractor("Eli Morgan", "ClearCom Crew", "Overnight Shift", ["Sun", "Mon", "Tue", "Wed"], "Training", 88, 0, 0, 6),
  makeContractor("Nina Patel", "Northline AV", "AM Shift", ["Mon", "Wed", "Fri"], "Off Shift", 84, 2, 3, 4, false)
];

const starterLaptops = [
  { id: crypto.randomUUID(), asset: "RIV-LT-101", contractor: "Jordan Lee", status: "Assigned", assignedBy: "Ops Admin", assignedDate: today, returnedBy: "", returnedDate: "", dueDate: today, history: "Assigned for AM shift" },
  { id: crypto.randomUUID(), asset: "RIV-LT-102", contractor: "Maya Santos", status: "Overdue", assignedBy: "Ops Admin", assignedDate: today, returnedBy: "", returnedDate: "", dueDate: today, history: "Return overdue" },
  { id: crypto.randomUUID(), asset: "RIV-LT-103", contractor: "", status: "Returned", assignedBy: "Ops Admin", assignedDate: today, returnedBy: "Site Lead", returnedDate: today, dueDate: today, history: "Returned clean" }
];

function makeContractor(name, company, shift, scheduledDays, status, attendance, noShows, lateArrivals, attendanceDays, active = true) {
  return {
    id: crypto.randomUUID(),
    name,
    email: `${name.toLowerCase().replaceAll(" ", ".")}@example.com`,
    phone: "555-0100",
    company,
    shift,
    scheduledDays,
    active,
    status,
    attendance,
    noShows,
    lateArrivals,
    attendanceDays,
    trainingCompletion: status === "Training" ? 62 : 100
  };
}

function readSaved(key, fallback) {
  try {
    return JSON.parse(localStorage.getItem(key)) ?? fallback;
  } catch {
    return fallback;
  }
}

function save(key, value, setter) {
  setter(value);
  localStorage.setItem(key, JSON.stringify(value));
}

function normalizeShift(shift) {
  if (shift === "Morning") return "AM Shift";
  if (shift === "Swing") return "PM Shift";
  if (shift === "Night") return "Overnight Shift";
  return shifts.includes(shift) ? shift : "AM Shift";
}

function shiftLabel(shift) {
  return normalizeShift(shift);
}

function escapeValue(value) {
  return String(value ?? "").replaceAll('"', '""');
}

function downloadCsv(filename, headers, rows) {
  const csv = [headers, ...rows].map((row) => row.map((cell) => `"${escapeValue(cell)}"`).join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function escapeSheetValue(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function tableHtml(title, headers, rows) {
  return `
    <h2>${escapeSheetValue(title)}</h2>
    <table>
      <thead><tr>${headers.map((header) => `<th>${escapeSheetValue(header)}</th>`).join("")}</tr></thead>
      <tbody>${rows.map((row) => `<tr>${row.map((cell) => `<td>${escapeSheetValue(cell)}</td>`).join("")}</tr>`).join("")}</tbody>
    </table>
  `;
}

function downloadSpreadsheet(data) {
  const { contractors, callouts, tickets, laptops, notes, stats } = data;
  const workbook = `
    <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
      <head><meta charset="UTF-8" /><style>
        body { font-family: Arial, sans-serif; color: #1d2725; }
        h1, h2 { color: #244238; }
        table { border-collapse: collapse; margin-bottom: 20px; }
        th { background: #244238; color: #fffdf8; font-weight: bold; }
        th, td { border: 1px solid #d9d1c2; padding: 7px 10px; mso-number-format:"\\@"; }
      </style></head>
      <body>
        <h1>FleetOps Leadership Review</h1>
        ${tableHtml("Executive KPIs", ["Metric", "Value"], Object.entries(stats))}
        ${tableHtml("Contractors", ["Name", "Company", "Shift", "Status", "Attendance", "No Shows", "Late", "Training"], contractors.map((c) => [c.name, c.company, shiftLabel(c.shift), c.status, `${c.attendance ?? 95}%`, c.noShows ?? 0, c.lateArrivals ?? 0, `${c.trainingCompletion ?? 100}%`]))}
        ${tableHtml("Callout Log", ["Date", "Time", "Name", "Company", "Shift", "Reason", "Hours Before Shift", "Coverage Status", "Replacement", "Status"], callouts.map((c) => [c.date, c.submittedTime, c.name, c.company, shiftLabel(c.shift), c.reason, c.hoursBeforeShiftStart, c.coverageStatus, c.replacementAssigned, c.status]))}
        ${tableHtml("Coverage Tickets", ["Date", "Shift", "Contractor", "Status", "Replacement", "Response Time", "Notes"], tickets.map((t) => [t.date, shiftLabel(t.shift), t.contractorName, t.status, t.replacementAssigned, t.responseTimeMinutes ? `${t.responseTimeMinutes} min` : "", t.notes]))}
        ${tableHtml("Laptop Accountability", ["Asset", "Contractor", "Status", "Assigned By", "Assigned Date", "Returned By", "Returned Date", "Due Date"], laptops.map((l) => [l.asset, l.contractor, l.status, l.assignedBy, l.assignedDate, l.returnedBy, l.returnedDate, l.dueDate]))}
        ${tableHtml("Operations Notes", ["Date", "Note"], notes.map((n) => [n.date, n.text]))}
      </body>
    </html>
  `;
  const blob = new Blob([workbook], { type: "application/vnd.ms-excel" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `fleetops-leadership-${today}.xls`;
  link.click();
  URL.revokeObjectURL(url);
}

function App() {
  const calloutOnly = new URLSearchParams(window.location.search).get("callout") === "1";
  const calloutLink = useMemo(() => {
    const url = new URL(window.location.href);
    url.search = "";
    url.hash = "";
    url.searchParams.set("callout", "1");
    return url.toString();
  }, []);
  const [role, setRole] = useState(calloutOnly ? "operator" : "admin");
  const [contractors, setContractors] = useState(() => readSaved("contractors", starterContractors).map((c) => ({ ...c, shift: normalizeShift(c.shift), status: c.status ?? "Scheduled", attendance: c.attendance ?? 95, noShows: c.noShows ?? 0, lateArrivals: c.lateArrivals ?? 0, attendanceDays: c.attendanceDays ?? 0, trainingCompletion: c.trainingCompletion ?? 100 })));
  const [callouts, setCallouts] = useState(() => readSaved("callouts", []));
  const [tickets, setTickets] = useState(() => readSaved("coverageTickets", []));
  const [laptops, setLaptops] = useState(() => readSaved("laptops", starterLaptops));
  const [notes, setNotes] = useState(() => readSaved("opsNotes", [{ id: crypto.randomUUID(), date: today, text: "AM shift short 2 operators." }]));
  const [notifications, setNotifications] = useState(() => readSaved("opsNotifications", []));

  function saveContractors(next) {
    save("contractors", next, setContractors);
  }

  function pushNotification(message) {
    const next = [{ id: crypto.randomUUID(), time: new Date().toLocaleTimeString(), message }, ...notifications];
    save("opsNotifications", next, setNotifications);
  }

  function submitCallout(callout) {
    const contractor = contractors.find((c) => c.name.toLowerCase() === callout.name.toLowerCase());
    const shift = normalizeShift(callout.shift);
    const submitted = new Date();
    const start = new Date(`${callout.date}T${String(shiftStart[shift]).padStart(2, "0")}:00:00`);
    const hoursBefore = Math.max(0, Math.round((start - submitted) / 36e5));
    const enriched = {
      id: crypto.randomUUID(),
      submittedAt: submitted.toISOString(),
      submittedTime: submitted.toLocaleTimeString(),
      company: contractor?.company ?? "",
      shift,
      hoursBeforeShiftStart: hoursBefore,
      coverageStatus: "Coverage Needed",
      replacementAssigned: "",
      status: "Coverage Needed",
      ...callout
    };
    const ticket = {
      id: crypto.randomUUID(),
      calloutId: enriched.id,
      date: callout.date,
      shift,
      contractorName: callout.name,
      company: contractor?.company ?? "",
      reason: callout.reason,
      status: "Coverage Needed",
      replacementAssigned: "",
      notes: "",
      openedAt: submitted.toISOString(),
      responseTimeMinutes: ""
    };
    save("callouts", [enriched, ...callouts], setCallouts);
    save("coverageTickets", [ticket, ...tickets], setTickets);
    pushNotification(`Callout submitted: ${callout.name} / ${shift}. Coverage ticket created.`);
  }

  const state = { contractors, callouts, tickets, laptops, notes, notifications };
  const actions = { saveContractors, setCallouts, setTickets, setLaptops, setNotes, setNotifications, pushNotification };

  return (
    <main className="app-shell">
      <header className="topbar">
        <div className="brand-lockup">
          <img className="brand-logo" src="/rivian-logo.png" alt="Rivian" />
          <div>
            <p className="eyebrow">FleetOps Command Center</p>
            <h1>Workforce Operations</h1>
          </div>
        </div>
        {!calloutOnly && (
          <div className="role-switch" aria-label="Role switcher">
            <button className={role === "operator" ? "active" : ""} onClick={() => setRole("operator")}>Operator</button>
            <button className={role === "admin" ? "active" : ""} onClick={() => setRole("admin")}>Admin</button>
          </div>
        )}
      </header>

      {role === "operator" ? (
        <OperatorForm contractors={contractors} onSubmit={submitCallout} calloutLink={calloutLink} showShareLink={!calloutOnly} />
      ) : (
        <AdminCommandCenter state={state} actions={actions} />
      )}
    </main>
  );
}

function OperatorForm({ contractors, onSubmit, calloutLink, showShareLink }) {
  const [form, setForm] = useState({ name: "", date: today, shift: "AM Shift", reason: "Sick", notes: "" });
  const [sent, setSent] = useState(false);
  const [copied, setCopied] = useState(false);
  const activeNames = contractors.filter((c) => c.active).map((c) => c.name);

  function update(field, value) {
    setForm((current) => ({ ...current, [field]: value }));
    setSent(false);
  }

  function handleSubmit(event) {
    event.preventDefault();
    if (!form.name.trim() || !form.reason.trim()) return;
    onSubmit(form);
    setForm({ name: "", date: today, shift: "AM Shift", reason: "Sick", notes: "" });
    setSent(true);
  }

  async function copyCalloutLink(event) {
    try {
      await navigator.clipboard.writeText(calloutLink);
    } catch {
      const input = event.currentTarget.previousElementSibling;
      input.focus();
      input.select();
      document.execCommand("copy");
    }
    setCopied(true);
    window.setTimeout(() => setCopied(false), 1600);
  }

  return (
    <div className="operator-stack">
      {showShareLink && (
        <section className="panel narrow">
          <div className="section-heading">
            <div><p className="eyebrow">Contractor Link</p><h2>Callout Landing Page</h2></div>
            {copied && <span className="success-pill">Copied</span>}
          </div>
          <div className="link-box">
            <input value={calloutLink} readOnly aria-label="Contractor callout link" />
            <button className="primary-button" type="button" onClick={copyCalloutLink}>Copy</button>
          </div>
        </section>
      )}

      <section className="panel narrow">
        <div className="section-heading">
          <div><p className="eyebrow">Contractor</p><h2>Submit Callout</h2></div>
          {sent && <span className="success-pill">Submitted</span>}
        </div>
        <form className="form-grid" onSubmit={handleSubmit}>
          <label>Name<input list="contractor-names" value={form.name} onChange={(event) => update("name", event.target.value)} placeholder="Contractor name" required /></label>
          <datalist id="contractor-names">{activeNames.map((name) => <option value={name} key={name} />)}</datalist>
          <label>Date<input type="date" value={form.date} onChange={(event) => update("date", event.target.value)} required /></label>
          <label>Shift<select value={form.shift} onChange={(event) => update("shift", event.target.value)}>{shifts.map((shift) => <option key={shift}>{shift}</option>)}</select></label>
          <label>Reason<select value={form.reason} onChange={(event) => update("reason", event.target.value)}>{reasons.map((reason) => <option key={reason}>{reason}</option>)}</select></label>
          <label className="wide">Notes<textarea value={form.notes} onChange={(event) => update("notes", event.target.value)} placeholder="Optional context for operations" /></label>
          <button className="primary-button" type="submit">Submit Callout</button>
        </form>
      </section>
    </div>
  );
}

function AdminCommandCenter({ state, actions }) {
  const [page, setPage] = useState("Daily Ops");
  const pages = ["Daily Ops", "Availability", "Callouts", "Coverage", "Laptops", "Performance", "Analytics", "Notes", "Notifications", "AI Assistant", "Executive", "Contractors"];
  const metrics = useMemo(() => buildMetrics(state), [state]);

  return (
    <div className="dashboard">
      <nav className="ops-nav" aria-label="Operations views">
        {pages.map((item) => <button key={item} className={page === item ? "active" : ""} onClick={() => setPage(item)}>{item}</button>)}
      </nav>
      {page === "Daily Ops" && <DailyOperations state={state} metrics={metrics} />}
      {page === "Availability" && <AvailabilityBoard contractors={state.contractors} callouts={state.callouts} />}
      {page === "Callouts" && <CalloutCenter state={state} actions={actions} />}
      {page === "Coverage" && <CoverageWorkflow state={state} actions={actions} metrics={metrics} />}
      {page === "Laptops" && <LaptopModule state={state} actions={actions} />}
      {page === "Performance" && <PerformanceScorecard contractors={state.contractors} callouts={state.callouts} />}
      {page === "Analytics" && <WorkforceAnalytics state={state} />}
      {page === "Notes" && <OperationsNotes notes={state.notes} setNotes={actions.setNotes} />}
      {page === "Notifications" && <NotificationsFeed notifications={state.notifications} />}
      {page === "AI Assistant" && <AiAssistant state={state} metrics={metrics} />}
      {page === "Executive" && <ExecutiveDashboard state={state} metrics={metrics} />}
      {page === "Contractors" && <ContractorAdmin contractors={state.contractors} callouts={state.callouts} onSaveContractors={actions.saveContractors} />}
    </div>
  );
}

function buildMetrics({ contractors, callouts, tickets, laptops }) {
  const active = contractors.filter((c) => c.active);
  const scheduled = active.filter((c) => c.scheduledDays.includes(todayDay));
  const todayCallouts = callouts.filter((c) => c.date === today);
  const openTickets = tickets.filter((t) => ["Pending", "Coverage Needed"].includes(t.status));
  const coveredTickets = tickets.filter((t) => t.status === "Covered" || t.status === "Closed");
  const overdue = laptops.filter((l) => l.status === "Overdue").length;
  const returned = laptops.filter((l) => l.status === "Returned").length;
  const assigned = laptops.filter((l) => l.status === "Assigned" || l.status === "Overdue").length;
  const available = scheduled.length - todayCallouts.length;
  const coveragePct = scheduled.length ? Math.round((available / scheduled.length) * 100) : 100;
  const attendance = active.length ? Math.round(active.reduce((sum, c) => sum + (Number(c.attendance) || 0), 0) / active.length) : 100;
  const laptopCompliance = assigned + returned ? Math.round((returned / (assigned + returned)) * 100) : 100;
  const readiness = Math.max(0, Math.round((coveragePct * 0.45) + (attendance * 0.3) + (laptopCompliance * 0.2) - (openTickets.length * 2)));
  const avgCoverage = coveredTickets.length ? Math.round(coveredTickets.reduce((sum, t) => sum + (Number(t.responseTimeMinutes) || 0), 0) / coveredTickets.length) : 0;
  return { active, scheduled, todayCallouts, openTickets, coveredTickets, overdue, returned, assigned, available, coveragePct, attendance, laptopCompliance, readiness, avgCoverage };
}

function Metric({ label, value }) {
  return <article className="metric-card"><p>{label}</p><strong>{value}</strong></article>;
}

function DailyOperations({ state, metrics }) {
  const level = metrics.readiness >= 90 ? "Healthy" : metrics.readiness >= 75 ? "Watch" : "Critical";
  return (
    <>
      <section className="metric-grid">
        <Metric label="Total Scheduled" value={metrics.scheduled.length} />
        <Metric label="Total Available" value={metrics.available} />
        <Metric label="Total Callouts" value={metrics.todayCallouts.length} />
        <Metric label="Open Coverage Gaps" value={metrics.openTickets.length} />
        <Metric label="Contractors In Training" value={state.contractors.filter((c) => c.status === "Training").length} />
        <Metric label="Laptop Issues" value={metrics.overdue} />
      </section>
      <section className="panel readiness-panel">
        <div><p className="eyebrow">Operational Readiness</p><h2>{metrics.readiness}%</h2></div>
        <span className={`readiness ${level.toLowerCase()}`}>{level}</span>
      </section>
      <NotificationsFeed notifications={state.notifications} />
    </>
  );
}

function AvailabilityBoard({ contractors, callouts }) {
  const calledOut = new Set(callouts.filter((c) => c.date === today).map((c) => c.name.toLowerCase()));
  function statusFor(contractor) {
    if (!contractor.active) return "Off Shift";
    if (calledOut.has(contractor.name.toLowerCase())) return "Called Out";
    if (!contractor.scheduledDays.includes(todayDay)) return "Off Shift";
    return contractor.status ?? "Scheduled";
  }
  return (
    <section className="panel">
      <div className="section-heading"><div><p className="eyebrow">Real-Time Staffing</p><h2>Contractor Availability Board</h2></div><span className="date-pill">{todayDay} {today}</span></div>
      <div className="availability-grid">
        {shifts.map((shift) => {
          const shiftContractors = contractors.filter((c) => normalizeShift(c.shift) === shift);
          return (
            <article className="board-column" key={shift}>
              <h3>{shift}</h3>
              {workforceStatuses.map((status) => {
                const count = shiftContractors.filter((c) => statusFor(c) === status).length;
                return <div className="board-row" key={status}><span>{status}</span><strong>{count}</strong></div>;
              })}
            </article>
          );
        })}
      </div>
    </section>
  );
}

function CalloutCenter({ state }) {
  const [filters, setFilters] = useState({ date: "", company: "", shift: "", reason: "" });
  const companies = [...new Set(state.contractors.map((c) => c.company).filter(Boolean))];
  const filtered = state.callouts.filter((c) =>
    (!filters.date || c.date === filters.date) &&
    (!filters.company || c.company === filters.company) &&
    (!filters.shift || normalizeShift(c.shift) === filters.shift) &&
    (!filters.reason || c.reason === filters.reason)
  );
  function exportFiltered() {
    downloadCsv(`callouts-${today}.csv`, ["Date", "Time", "Name", "Company", "Shift", "Reason", "Hours Before Shift", "Coverage Status", "Replacement", "Status"], filtered.map((c) => [c.date, c.submittedTime, c.name, c.company, normalizeShift(c.shift), c.reason, c.hoursBeforeShiftStart, c.coverageStatus, c.replacementAssigned, c.status]));
  }
  return (
    <section className="panel">
      <div className="section-heading"><div><p className="eyebrow">Callout Management</p><h2>Callout Center</h2></div><button className="primary-button" onClick={exportFiltered}>Export CSV</button></div>
      <div className="filters-grid">
        <label>Date<input type="date" value={filters.date} onChange={(e) => setFilters({ ...filters, date: e.target.value })} /></label>
        <label>Company<select value={filters.company} onChange={(e) => setFilters({ ...filters, company: e.target.value })}><option value="">All</option>{companies.map((c) => <option key={c}>{c}</option>)}</select></label>
        <label>Shift<select value={filters.shift} onChange={(e) => setFilters({ ...filters, shift: e.target.value })}><option value="">All</option>{shifts.map((s) => <option key={s}>{s}</option>)}</select></label>
        <label>Reason<select value={filters.reason} onChange={(e) => setFilters({ ...filters, reason: e.target.value })}><option value="">All</option>{reasons.map((r) => <option key={r}>{r}</option>)}</select></label>
      </div>
      <CalloutTable callouts={filtered} />
    </section>
  );
}

function CalloutTable({ callouts }) {
  return (
    <div className="table-wrap">
      <table>
        <thead><tr><th>Date</th><th>Time</th><th>Name</th><th>Company</th><th>Shift</th><th>Reason</th><th>Hours Before</th><th>Coverage</th><th>Replacement</th><th>Status</th></tr></thead>
        <tbody>{callouts.length ? callouts.map((c) => <tr key={c.id}><td>{c.date}</td><td>{c.submittedTime}</td><td>{c.name}</td><td>{c.company}</td><td>{normalizeShift(c.shift)}</td><td>{c.reason}</td><td>{c.hoursBeforeShiftStart}</td><td>{c.coverageStatus}</td><td>{c.replacementAssigned || "Unassigned"}</td><td><span className="status called-out">{c.status}</span></td></tr>) : <tr><td colSpan="10">No callouts match filters.</td></tr>}</tbody>
      </table>
    </div>
  );
}

function CoverageWorkflow({ state, actions, metrics }) {
  function updateTicket(id, patch) {
    const next = state.tickets.map((t) => {
      if (t.id !== id) return t;
      const responseTimeMinutes = patch.status === "Covered" && !t.responseTimeMinutes ? Math.max(1, Math.round((Date.now() - new Date(t.openedAt).getTime()) / 60000)) : t.responseTimeMinutes;
      return { ...t, ...patch, responseTimeMinutes };
    });
    save("coverageTickets", next, actions.setTickets);
    if (patch.status === "Covered") actions.pushNotification("Coverage assigned and shift marked covered.");
  }
  return (
    <>
      <section className="metric-grid">
        <Metric label="Open Coverage Requests" value={metrics.openTickets.length} />
        <Metric label="Covered Shifts" value={metrics.coveredTickets.length} />
        <Metric label="Uncovered Shifts" value={metrics.openTickets.length} />
        <Metric label="Avg Coverage Time" value={`${metrics.avgCoverage}m`} />
      </section>
      <section className="panel">
        <div className="section-heading"><div><p className="eyebrow">Workflow</p><h2>Coverage Needed Tickets</h2></div></div>
        <div className="ticket-grid">
          {state.tickets.length ? state.tickets.map((ticket) => (
            <article className="ticket-card" key={ticket.id}>
              <div><strong>{ticket.contractorName}</strong><span>{ticket.date} / {normalizeShift(ticket.shift)} / {ticket.reason}</span></div>
              <label>Replacement<select value={ticket.replacementAssigned} onChange={(e) => updateTicket(ticket.id, { replacementAssigned: e.target.value })}><option value="">Unassigned</option>{state.contractors.filter((c) => c.active && c.name !== ticket.contractorName).map((c) => <option key={c.id}>{c.name}</option>)}</select></label>
              <label>Status<select value={ticket.status} onChange={(e) => updateTicket(ticket.id, { status: e.target.value })}>{coverageStatuses.map((s) => <option key={s}>{s}</option>)}</select></label>
              <label>Notes<textarea value={ticket.notes} onChange={(e) => updateTicket(ticket.id, { notes: e.target.value })} /></label>
              <span className="date-pill">Response: {ticket.responseTimeMinutes ? `${ticket.responseTimeMinutes} min` : "Open"}</span>
            </article>
          )) : <p>No coverage tickets yet.</p>}
        </div>
      </section>
    </>
  );
}

function LaptopModule({ state, actions }) {
  const assigned = state.laptops.filter((l) => l.status === "Assigned" || l.status === "Overdue").length;
  const returned = state.laptops.filter((l) => l.status === "Returned").length;
  const overdue = state.laptops.filter((l) => l.status === "Overdue").length;
  function updateLaptop(id, patch) {
    save("laptops", state.laptops.map((l) => l.id === id ? { ...l, ...patch, history: `${l.history}; ${patch.status ?? "Updated"} ${new Date().toLocaleTimeString()}` } : l), actions.setLaptops);
    if (patch.status === "Overdue") actions.pushNotification("Laptop became overdue.");
  }
  return (
    <>
      <section className="metric-grid laptop-grid">
        <Metric label="Assigned" value={assigned} /><Metric label="Returned" value={returned} /><Metric label="Overdue" value={overdue} />
      </section>
      <section className="panel">
        <div className="section-heading"><div><p className="eyebrow">Laptop Accountability</p><h2>Compliance, Alerts, Audit Log</h2></div></div>
        <div className="table-wrap"><table><thead><tr><th>Asset</th><th>Contractor</th><th>Status</th><th>Assigned By</th><th>Assigned Date</th><th>Returned By</th><th>Returned Date</th><th>History</th></tr></thead><tbody>
          {state.laptops.map((l) => <tr key={l.id}><td>{l.asset}</td><td>{l.contractor || "Unassigned"}</td><td><select value={l.status} onChange={(e) => updateLaptop(l.id, { status: e.target.value })}><option>Assigned</option><option>Returned</option><option>Overdue</option><option>Missing</option></select></td><td>{l.assignedBy}</td><td>{l.assignedDate}</td><td>{l.returnedBy || "-"}</td><td>{l.returnedDate || "-"}</td><td>{l.history}</td></tr>)}
        </tbody></table></div>
      </section>
    </>
  );
}

function PerformanceScorecard({ contractors, callouts }) {
  function badge(c) {
    if ((c.attendance ?? 0) >= 96 && (c.noShows ?? 0) === 0) return "Excellent";
    if ((c.attendance ?? 0) >= 90) return "Good";
    if ((c.noShows ?? 0) > 1 || (c.attendance ?? 0) < 85) return "Needs Attention";
    return "Watchlist";
  }
  return (
    <section className="panel">
      <div className="section-heading"><div><p className="eyebrow">Contractor Performance</p><h2>Scorecard</h2></div></div>
      <div className="table-wrap"><table><thead><tr><th>Name</th><th>Attendance</th><th>Total Callouts</th><th>No Shows</th><th>Late</th><th>Consecutive Days</th><th>Training</th><th>Badge</th></tr></thead><tbody>
        {contractors.map((c) => <tr key={c.id}><td>{c.name}</td><td>{c.attendance ?? 0}%</td><td>{callouts.filter((x) => x.name === c.name).length}</td><td>{c.noShows ?? 0}</td><td>{c.lateArrivals ?? 0}</td><td>{c.attendanceDays ?? 0}</td><td>{c.trainingCompletion ?? 100}%</td><td><span className={`status ${badge(c).toLowerCase().replaceAll(" ", "-")}`}>{badge(c)}</span></td></tr>)}
      </tbody></table></div>
    </section>
  );
}

function WorkforceAnalytics({ state }) {
  const byCompany = groupCount(state.callouts, (c) => c.company || "Unknown");
  const byShift = groupCount(state.callouts, (c) => normalizeShift(c.shift));
  const byReason = groupCount(state.callouts, (c) => c.reason || "Other");
  const byDay = groupCount(state.callouts, (c) => days[new Date(`${c.date}T00:00:00`).getDay()]);
  return (
    <section className="analytics-grid">
      <BarPanel title="Callouts by Day" data={byDay} />
      <BarPanel title="Callouts by Company" data={byCompany} />
      <BarPanel title="Shift Coverage Trends" data={byShift} />
      <BarPanel title="No-Show / Reason Trends" data={byReason} />
    </section>
  );
}

function groupCount(items, keyFn) {
  return items.reduce((acc, item) => {
    const key = keyFn(item);
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
}

function BarPanel({ title, data }) {
  const max = Math.max(1, ...Object.values(data));
  return (
    <section className="panel">
      <div className="section-heading"><div><p className="eyebrow">Analytics</p><h2>{title}</h2></div></div>
      <div className="bar-list">{Object.keys(data).length ? Object.entries(data).map(([label, value]) => <div className="bar-row" key={label}><span>{label}</span><div><i style={{ width: `${(value / max) * 100}%` }} /></div><strong>{value}</strong></div>) : <p>No data yet.</p>}</div>
    </section>
  );
}

function OperationsNotes({ notes, setNotes }) {
  const [draft, setDraft] = useState("");
  const [date, setDate] = useState(today);
  const [search, setSearch] = useState("");
  const filtered = notes.filter((n) => (!search || n.text.toLowerCase().includes(search.toLowerCase())) && (!date || n.date === date));
  function addNote(event) {
    event.preventDefault();
    if (!draft.trim()) return;
    save("opsNotes", [{ id: crypto.randomUUID(), date, text: draft }, ...notes], setNotes);
    setDraft("");
  }
  return (
    <section className="panel">
      <div className="section-heading"><div><p className="eyebrow">Daily Log</p><h2>Operations Notes</h2></div></div>
      <form className="form-grid" onSubmit={addNote}>
        <label>Date<input type="date" value={date} onChange={(e) => setDate(e.target.value)} /></label>
        <label>Search<input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search notes" /></label>
        <label className="wide">Note<textarea value={draft} onChange={(e) => setDraft(e.target.value)} placeholder="AM shift short 2 operators." /></label>
        <button className="primary-button">Add Note</button>
      </form>
      <div className="notes-list">{filtered.map((n) => <article key={n.id}><strong>{n.date}</strong><p>{n.text}</p></article>)}</div>
    </section>
  );
}

function NotificationsFeed({ notifications }) {
  return (
    <section className="panel callout-log">
      <div className="section-heading"><div><p className="eyebrow">Slack-Ready Feed</p><h2>Operations Notifications</h2></div></div>
      <div className="notes-list">{notifications.length ? notifications.map((n) => <article key={n.id}><strong>{n.time}</strong><p>{n.message}</p></article>) : <p>No notifications yet.</p>}</div>
    </section>
  );
}

function AiAssistant({ state, metrics }) {
  const [prompt, setPrompt] = useState("Who is available to cover tonight's PM shift?");
  const answer = useMemo(() => generateAiAnswer(prompt, state, metrics), [prompt, state, metrics]);
  const examples = [
    "Who is available to cover tonight's PM shift?",
    "Which contractors have the highest attendance?",
    "Which contractor companies have the most callouts?",
    "Show staffing shortages for this week.",
    "Generate daily operations summary."
  ];
  return (
    <section className="panel">
      <div className="section-heading"><div><p className="eyebrow">AI Workforce Assistant</p><h2>Planning Recommendations</h2></div></div>
      <div className="assistant-layout">
        <div className="notes-list">{examples.map((item) => <button key={item} onClick={() => setPrompt(item)}>{item}</button>)}</div>
        <label>Prompt<textarea value={prompt} onChange={(e) => setPrompt(e.target.value)} /></label>
        <article className="assistant-answer"><strong>Recommendation</strong><p>{answer}</p></article>
      </div>
    </section>
  );
}

function generateAiAnswer(prompt, state, metrics) {
  const lower = prompt.toLowerCase();
  if (lower.includes("available")) return state.contractors.filter((c) => c.active && c.status !== "Called Out" && c.status !== "No Show").map((c) => `${c.name} (${normalizeShift(c.shift)})`).join(", ") || "No available contractors found.";
  if (lower.includes("highest attendance")) return [...state.contractors].sort((a, b) => (b.attendance ?? 0) - (a.attendance ?? 0)).slice(0, 3).map((c) => `${c.name}: ${c.attendance}%`).join(", ");
  if (lower.includes("companies")) return Object.entries(groupCount(state.callouts, (c) => c.company || "Unknown")).sort((a, b) => b[1] - a[1]).map(([company, count]) => `${company}: ${count}`).join(", ") || "No callouts by company yet.";
  if (lower.includes("95")) return `${Math.max(0, Math.ceil((metrics.scheduled.length * 0.95) - metrics.available))} additional contractors needed to maintain 95% coverage.`;
  return `Today: ${metrics.scheduled.length} scheduled, ${metrics.available} available, ${metrics.todayCallouts.length} callouts, ${metrics.openTickets.length} open coverage gaps, readiness ${metrics.readiness}%.`;
}

function ExecutiveDashboard({ state, metrics }) {
  const stats = {
    "Total Contractors": state.contractors.length,
    "Active Contractors": metrics.active.length,
    "Staffing Coverage %": `${metrics.coveragePct}%`,
    "Callout Rate %": `${metrics.scheduled.length ? Math.round((metrics.todayCallouts.length / metrics.scheduled.length) * 100) : 0}%`,
    "Attendance %": `${metrics.attendance}%`,
    "Open Coverage Tickets": metrics.openTickets.length,
    "Laptop Compliance %": `${metrics.laptopCompliance}%`,
    "Operational Readiness Score": `${metrics.readiness}%`
  };
  return (
    <>
      <section className="metric-grid executive-grid">{Object.entries(stats).map(([label, value]) => <Metric key={label} label={label} value={value} />)}</section>
      <section className="panel export-panel">
        <div><p className="eyebrow">Leadership Reporting</p><h2>Export Full Operations Workbook</h2></div>
        <button className="primary-button" onClick={() => downloadSpreadsheet({ ...state, stats })}>Download Spreadsheet</button>
      </section>
      <WorkforceAnalytics state={state} />
    </>
  );
}

function ContractorAdmin({ contractors, callouts, onSaveContractors }) {
  const [editingId, setEditingId] = useState(null);
  const blank = { name: "", email: "", phone: "", company: "", shift: "AM Shift", scheduledDays: ["Mon", "Tue", "Wed", "Thu", "Fri"], active: true, status: "Scheduled", attendance: 95, noShows: 0, lateArrivals: 0, attendanceDays: 0, trainingCompletion: 100 };
  const [draft, setDraft] = useState(blank);
  function startEdit(contractor) {
    setEditingId(contractor.id);
    setDraft({ ...blank, ...contractor, shift: normalizeShift(contractor.shift) });
  }
  function resetForm() {
    setEditingId(null);
    setDraft(blank);
  }
  function saveDraft(event) {
    event.preventDefault();
    if (!draft.name.trim()) return;
    const next = editingId ? contractors.map((c) => c.id === editingId ? draft : c) : [{ ...draft, id: crypto.randomUUID() }, ...contractors];
    onSaveContractors(next);
    resetForm();
  }
  function removeContractor(id) {
    onSaveContractors(contractors.filter((c) => c.id !== id));
  }
  function toggleDay(day) {
    setDraft({ ...draft, scheduledDays: draft.scheduledDays.includes(day) ? draft.scheduledDays.filter((d) => d !== day) : [...draft.scheduledDays, day] });
  }
  return (
    <section className="layout-grid">
      <div className="panel">
        <div className="section-heading"><div><p className="eyebrow">Admin</p><h2>Contractors</h2></div><span className="date-pill">{todayDay} {today}</span></div>
        <ContractorTable contractors={contractors} callouts={callouts} onEdit={startEdit} onRemove={removeContractor} />
      </div>
      <aside className="panel">
        <div className="section-heading"><div><p className="eyebrow">{editingId ? "Edit" : "Add"}</p><h2>Contractor</h2></div></div>
        <form className="form-grid compact" onSubmit={saveDraft}>
          <label>Name<input value={draft.name} onChange={(e) => setDraft({ ...draft, name: e.target.value })} required /></label>
          <label>Email<input type="email" value={draft.email} onChange={(e) => setDraft({ ...draft, email: e.target.value })} /></label>
          <label>Phone<input value={draft.phone} onChange={(e) => setDraft({ ...draft, phone: e.target.value })} /></label>
          <label>Company<input value={draft.company} onChange={(e) => setDraft({ ...draft, company: e.target.value })} /></label>
          <label>Shift<select value={draft.shift} onChange={(e) => setDraft({ ...draft, shift: e.target.value })}>{shifts.map((s) => <option key={s}>{s}</option>)}</select></label>
          <label>Status<select value={draft.status} onChange={(e) => setDraft({ ...draft, status: e.target.value })}>{workforceStatuses.map((s) => <option key={s}>{s}</option>)}</select></label>
          <fieldset><legend>Scheduled Days</legend><div className="day-grid">{days.map((day) => <label className="check-label" key={day}><input type="checkbox" checked={draft.scheduledDays.includes(day)} onChange={() => toggleDay(day)} />{day}</label>)}</div></fieldset>
          <label className="check-label inline"><input type="checkbox" checked={draft.active} onChange={(e) => setDraft({ ...draft, active: e.target.checked })} />Active</label>
          <div className="button-row"><button className="primary-button">{editingId ? "Save" : "Add"}</button><button className="ghost-button" type="button" onClick={resetForm}>Clear</button></div>
        </form>
      </aside>
    </section>
  );
}

function ContractorTable({ contractors, callouts, onEdit, onRemove }) {
  function lastCalloutFor(name) {
    return callouts.find((callout) => callout.name.toLowerCase() === name.toLowerCase())?.date ?? "None";
  }
  return (
    <div className="table-wrap"><table><thead><tr><th>Name</th><th>Shift</th><th>Scheduled Today</th><th>Status</th><th>Last Callout</th><th>Actions</th></tr></thead><tbody>
      {contractors.map((c) => <tr key={c.id}><td><strong>{c.name}</strong><span>{c.company}</span></td><td>{normalizeShift(c.shift)}</td><td>{c.active && c.scheduledDays.includes(todayDay) ? "Yes" : "No"}</td><td><span className={`status ${(c.status ?? "scheduled").toLowerCase().replaceAll(" ", "-")}`}>{c.active ? c.status : "Inactive"}</span></td><td>{lastCalloutFor(c.name)}</td><td><div className="row-actions"><button onClick={() => onEdit(c)}>Edit</button><button className="danger" onClick={() => onRemove(c.id)}>Remove</button></div></td></tr>)}
    </tbody></table></div>
  );
}

createRoot(document.getElementById("root")).render(<App />);
