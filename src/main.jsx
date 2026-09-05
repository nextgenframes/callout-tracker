import React, { useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import * as XLSX from "xlsx";
import { hasSupabaseConfig, supabase } from "./supabaseClient";
import "./styles.css";

function localDateString(date = new Date()) {
  const offsetDate = new Date(date.getTime() - date.getTimezoneOffset() * 60000);
  return offsetDate.toISOString().slice(0, 10);
}

function dateOnly(value) {
  if (!value) return "";
  return String(value).slice(0, 10);
}

function addDays(dateText, amount) {
  const date = new Date(`${dateText}T00:00:00`);
  date.setDate(date.getDate() + amount);
  return localDateString(date);
}

const shifts = ["AM Shift", "PM Shift", "Overnight Shift"];
const operationShifts = [
  { label: "1st Shift", value: "AM Shift" },
  { label: "2nd Shift", value: "PM Shift" }
];
const sites = ["San Francisco", "Miami", "Chicago"];
const defaultSite = "San Francisco";
const days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const today = localDateString();
const todayDay = days[new Date().getDay()];
const workforceStatuses = ["Scheduled", "Present", "Called Out", "Late", "Leaving Early", "Day Off Requested", "No Show", "Training", "Off Shift", "Offboarding"];
const attendanceActions = ["Scheduled", "Late", "No Show", "Leaving Early", "Day Off Requested"];
const dailyResetStatuses = ["Present", "Called Out", "Late", "Leaving Early", "Day Off Requested", "No Show"];
const attendanceStatuses = ["P", "A", "T", "WO"];
const reportTypes = ["Daily attendance", "Weekly attendance", "Monthly attendance", "Vehicle utilization", "Operator utilization", "Callout trends", "Late trends"];
const coverageStatuses = ["Pending", "Coverage Needed", "Covered", "Closed"];
const reasons = ["Sick", "Emergency", "Transport", "Family", "Other"];
const shiftStart = { "AM Shift": 6, "PM Shift": 14, "Overnight Shift": 22 };
const vehicleModels = ["R1S", "R1T", "R2"];
const vehicleStatuses = ["Available", "Assigned", "In Service", "Maintenance", "Out of Service"];
const laptopStatuses = ["Assigned", "Returned", "Overdue", "Missing"];

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

const starterVehicles = [
  { id: crypto.randomUUID(), name: "Fleet R1T 01", vin: "7FCTGAAA0NN000001", status: "Available", licensePlate: "RIV-101", model: "R1T", year: "2025", location: "Main Lot" },
  { id: crypto.randomUUID(), name: "Fleet R1S 02", vin: "7PDSGABA0NN000002", status: "In Service", licensePlate: "RIV-202", model: "R1S", year: "2025", location: "Bay 3" }
];

function makeContractor(name, company, shift, scheduledDays, status, attendance, noShows, lateArrivals, attendanceDays, active = true) {
  return {
    id: crypto.randomUUID(),
    name,
    email: `${name.toLowerCase().replaceAll(" ", ".")}@example.com`,
    phone: "555-0100",
    company,
    startDate: today,
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

function siteFor(item) {
  return sites.includes(item?.site) ? item.site : defaultSite;
}

function statusDateMap() {
  return readSaved("contractorStatusDates", {});
}

function statusDateFor(id) {
  return statusDateMap()[id] ?? "";
}

function saveStatusDates(contractors) {
  const current = statusDateMap();
  const next = { ...current };
  contractors.forEach((contractor) => {
    if (contractor.statusDate) next[contractor.id] = contractor.statusDate;
  });
  localStorage.setItem("contractorStatusDates", JSON.stringify(next));
}

function isOffboarding(contractor) {
  return contractor?.status === "Offboarding";
}

function isOperationsContractor(contractor) {
  return Boolean(contractor?.active) && !isOffboarding(contractor);
}

function resetDailyStatusForDate(contractor, date = today) {
  if (dailyResetStatuses.includes(contractor.status) && contractor.statusDate && contractor.statusDate !== date) {
    return { ...contractor, status: "Scheduled", statusDate: "" };
  }
  return contractor;
}

async function confirmDelete(label) {
  const password = await requestDeletePassword(label);
  if (!password) return false;

  if (hasSupabaseConfig) {
    const { data } = await supabase.auth.getSession();
    const email = data.session?.user?.email;
    if (!email) {
      await showDeleteMessage("Please log in as an admin before deleting.");
      return false;
    }
    const { error } = await supabase.auth.signInWithPassword({ email, password });
    if (error) {
      await showDeleteMessage("Password incorrect. Delete cancelled.");
      return false;
    }
    return true;
  }

  const localDeletePassword = localStorage.getItem("deletePassword") || "admin";
  if (password !== localDeletePassword) {
    await showDeleteMessage("Password incorrect. Delete cancelled.");
    return false;
  }
  return true;
}

function requestDeletePassword(label) {
  return new Promise((resolve) => {
    const root = document.createElement("div");
    root.className = "confirm-modal-root";
    const backdrop = document.createElement("div");
    backdrop.className = "modal-backdrop confirm-delete-backdrop";
    const panel = document.createElement("form");
    panel.className = "modal-panel confirm-delete-panel";
    panel.innerHTML = `
      <div class="section-heading">
        <div><p class="eyebrow">Protected Action</p><h2>Confirm Delete</h2></div>
      </div>
      <p class="confirm-delete-copy"></p>
      <label>Password<input type="password" autocomplete="current-password" required placeholder="Enter account password" /></label>
      <div class="button-row">
        <button class="danger" type="submit">Delete</button>
        <button class="ghost-button" type="button" data-cancel>Cancel</button>
      </div>
    `;
    panel.querySelector(".confirm-delete-copy").textContent = `Delete ${label}? This cannot be undone.`;
    backdrop.append(panel);
    root.append(backdrop);
    document.body.append(root);
    const input = panel.querySelector("input");
    input.focus();

    function close(value = "") {
      root.remove();
      resolve(value);
    }

    panel.addEventListener("submit", (event) => {
      event.preventDefault();
      close(input.value);
    });
    panel.querySelector("[data-cancel]").addEventListener("click", () => close(""));
    backdrop.addEventListener("click", (event) => {
      if (event.target === backdrop) close("");
    });
  });
}

function showDeleteMessage(message) {
  return new Promise((resolve) => {
    const root = document.createElement("div");
    root.className = "confirm-modal-root";
    const backdrop = document.createElement("div");
    backdrop.className = "modal-backdrop confirm-delete-backdrop";
    const panel = document.createElement("div");
    panel.className = "modal-panel confirm-delete-panel";
    panel.innerHTML = `
      <div class="section-heading"><div><p class="eyebrow">Delete Cancelled</p><h2>Unable to Delete</h2></div></div>
      <p class="confirm-delete-copy"></p>
      <div class="button-row"><button class="primary-button" type="button">OK</button></div>
    `;
    panel.querySelector(".confirm-delete-copy").textContent = message;
    backdrop.append(panel);
    root.append(backdrop);
    document.body.append(root);
    const button = panel.querySelector("button");
    button.focus();
    button.addEventListener("click", () => {
      root.remove();
      resolve();
    });
  });
}

function displayNameForSession(session) {
  return session?.user?.user_metadata?.display_name || session?.user?.email || "Admin User";
}

function dbStatusClass(status) {
  const value = String(status).toLowerCase();
  if (value.includes("error")) return "error";
  if (value.includes("connected") || value.includes("synced")) return "connected";
  return "neutral";
}

function dbToContractor(row) {
  return {
    id: row.id,
    site: row.site ?? defaultSite,
    name: row.name ?? "",
    email: row.email ?? "",
    phone: row.phone ?? "",
    company: row.contractor_company ?? row.company ?? "",
    startDate: row.start_date ?? row.startDate ?? row.created_at?.slice(0, 10) ?? today,
    shift: normalizeShift(row.assigned_shift ?? row.shift),
    scheduledDays: row.scheduled_days ?? row.scheduledDays ?? [],
    active: (row.status ?? "active") !== "inactive",
    status: row.workforce_status ?? "Scheduled",
    statusDate: statusDateFor(row.id),
    attendance: row.attendance ?? 95,
    noShows: row.no_shows ?? 0,
    lateArrivals: row.late_arrivals ?? 0,
    attendanceDays: row.attendance_days ?? 0,
    trainingCompletion: row.training_completion ?? 100
  };
}

function contractorToDb(contractor) {
  return {
    id: contractor.id,
    site: contractor.site ?? defaultSite,
    name: contractor.name,
    email: contractor.email,
    phone: contractor.phone,
    contractor_company: contractor.company,
    start_date: contractor.startDate || null,
    assigned_shift: normalizeShift(contractor.shift),
    scheduled_days: contractor.scheduledDays,
    status: contractor.active ? "active" : "inactive",
    workforce_status: contractor.status ?? "Scheduled",
    training_completion: contractor.trainingCompletion ?? 100
  };
}

function dbToCallout(row) {
  return {
    id: row.id,
    site: row.site ?? defaultSite,
    contractorId: row.contractor_id ?? "",
    name: row.contractor_name ?? row.name ?? "",
    date: row.callout_date ?? row.date ?? today,
    submittedAt: row.submitted_at ?? "",
    submittedTime: row.submitted_at ? new Date(row.submitted_at).toLocaleTimeString() : "",
    company: row.company ?? "",
    shift: normalizeShift(row.shift),
    reason: row.reason ?? "",
    notes: row.notes ?? "",
    hoursBeforeShiftStart: row.hours_before_shift ?? 0,
    coverageStatus: row.coverage_status ?? "Coverage Needed",
    replacementAssigned: row.replacement_assigned ?? "",
    status: row.status ?? "Pending"
  };
}

function calloutToDb(callout) {
  return {
    id: callout.id,
    site: callout.site ?? defaultSite,
    contractor_id: callout.contractorId || null,
    contractor_name: callout.name,
    callout_date: callout.date,
    submitted_at: callout.submittedAt || new Date().toISOString(),
    shift: normalizeShift(callout.shift),
    reason: callout.reason,
    notes: callout.notes,
    hours_before_shift: callout.hoursBeforeShiftStart ?? 0,
    coverage_status: callout.coverageStatus ?? "Coverage Needed",
    replacement_assigned: callout.replacementAssigned ?? "",
    status: callout.status ?? "Pending"
  };
}

function dbToTicket(row) {
  return {
    id: row.id,
    site: row.site ?? defaultSite,
    calloutId: row.callout_id ?? "",
    contractorId: row.contractor_id ?? "",
    date: row.date ?? today,
    shift: normalizeShift(row.shift),
    contractorName: row.contractor_name ?? "",
    company: row.company ?? "",
    reason: row.reason ?? "",
    status: row.ticket_status ?? row.status ?? "Coverage Needed",
    replacementAssigned: row.replacement_assigned ?? "",
    notes: row.notes ?? "",
    openedAt: row.opened_at ?? "",
    coveredAt: row.covered_at ?? "",
    responseTimeMinutes: row.response_minutes ?? ""
  };
}

function ticketToDb(ticket) {
  return {
    id: ticket.id,
    site: ticket.site ?? defaultSite,
    callout_id: ticket.calloutId || null,
    contractor_id: ticket.contractorId || null,
    shift: normalizeShift(ticket.shift),
    ticket_status: ticket.status ?? "Coverage Needed",
    replacement_contractor_id: null,
    notes: ticket.notes ?? "",
    opened_at: ticket.openedAt || new Date().toISOString(),
    covered_at: ticket.status === "Covered" ? new Date().toISOString() : null,
    response_minutes: Number(ticket.responseTimeMinutes) || null
  };
}

function dbToLaptop(row) {
  return {
    id: row.id,
    site: row.site ?? defaultSite,
    asset: row.asset_tag ?? row.asset ?? "",
    contractorId: row.contractor_id ?? "",
    contractor: row.contractor ?? "",
    status: row.status ?? "Assigned",
    assignedBy: row.assigned_by ?? "",
    assignedDate: row.assigned_date ?? "",
    returnedBy: row.returned_by ?? "",
    returnedDate: row.returned_date ?? "",
    dueDate: row.due_date ?? row.assigned_date ?? "",
    history: row.audit_log ?? ""
  };
}

function laptopToDb(laptop) {
  return {
    id: laptop.id,
    site: laptop.site ?? defaultSite,
    asset_tag: laptop.asset,
    contractor_id: laptop.contractorId || null,
    status: laptop.status,
    assigned_by: laptop.assignedBy,
    assigned_date: laptop.assignedDate || null,
    returned_by: laptop.returnedBy,
    returned_date: laptop.returnedDate || null,
    audit_log: laptop.history
  };
}

function dbToNote(row) {
  return {
    id: row.id,
    site: row.site ?? defaultSite,
    date: row.note_date ?? row.date ?? today,
    text: row.note ?? row.text ?? "",
    createdBy: row.created_by ?? ""
  };
}

function noteToDb(note) {
  return {
    id: note.id,
    site: note.site ?? defaultSite,
    note_date: note.date,
    note: note.text,
    created_by: note.createdBy ?? "Ops Manager"
  };
}

function dbToNotification(row) {
  return {
    id: row.id,
    site: row.site ?? defaultSite,
    time: row.created_at ? new Date(row.created_at).toLocaleTimeString() : row.time ?? "",
    message: row.message ?? "",
    eventType: row.event_type ?? "operations",
    status: row.status ?? "New"
  };
}

function notificationToDb(notification) {
  return {
    id: notification.id,
    site: notification.site ?? defaultSite,
    event_type: notification.eventType ?? "operations",
    message: notification.message,
    status: notification.status ?? "New"
  };
}

function dbToAttendance(row) {
  return {
    id: row.id,
    site: row.site ?? defaultSite,
    shift: normalizeShift(row.shift),
    weekStart: row.week_start ?? today,
    contractorId: row.contractor_id ?? "",
    contractorName: row.contractor_name ?? "",
    company: row.company ?? "",
    monday: row.monday ?? "",
    tuesday: row.tuesday ?? "",
    wednesday: row.wednesday ?? "",
    thursday: row.thursday ?? "",
    friday: row.friday ?? "",
    comments: row.comments ?? ""
  };
}

function attendanceToDb(record) {
  return {
    id: record.id,
    site: record.site ?? defaultSite,
    shift: normalizeShift(record.shift),
    week_start: record.weekStart,
    contractor_id: record.contractorId || null,
    contractor_name: record.contractorName,
    company: record.company,
    monday: record.monday,
    tuesday: record.tuesday,
    wednesday: record.wednesday,
    thursday: record.thursday,
    friday: record.friday,
    comments: record.comments
  };
}

function dbToVehicle(row) {
  return {
    id: row.id,
    site: row.site ?? defaultSite,
    name: row.name ?? "",
    vin: row.vin ?? "",
    status: row.status ?? "Available",
    licensePlate: row.license_plate ?? row.licensePlate ?? "",
    model: row.model ?? "R1T",
    year: String(row.year ?? ""),
    location: row.location ?? "",
    assignedContractorId: row.assigned_contractor_id ?? "",
    assignedContractorName: row.assigned_contractor_name ?? ""
  };
}

function vehicleToDb(vehicle) {
  return {
    id: vehicle.id,
    site: vehicle.site ?? defaultSite,
    name: vehicle.name,
    vin: vehicle.vin,
    status: vehicle.status,
    license_plate: vehicle.licensePlate,
    model: vehicle.model,
    year: Number(vehicle.year) || null,
    location: vehicle.location,
    assigned_contractor_id: vehicle.assignedContractorId || null,
    assigned_contractor_name: vehicle.assignedContractorName || ""
  };
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

function calloutsForContractor(contractor, callouts) {
  return callouts.filter((callout) =>
    (contractor.id && callout.contractorId === contractor.id) ||
    callout.name?.toLowerCase() === contractor.name.toLowerCase()
  );
}

function scheduledDaysSinceStart(contractor) {
  if (!contractor.startDate) return 0;
  const start = new Date(`${contractor.startDate}T00:00:00`);
  const end = new Date(`${today}T00:00:00`);
  if (Number.isNaN(start.getTime()) || start > end) return 0;

  let count = 0;
  for (let date = new Date(start); date <= end; date.setDate(date.getDate() + 1)) {
    if (contractor.scheduledDays.includes(days[date.getDay()])) count += 1;
  }
  return count;
}

function attendanceRecordsForContractor(contractor, attendanceRecords = []) {
  return attendanceRecords.filter((record) =>
    (contractor.id && record.contractorId === contractor.id) ||
    record.contractorName?.toLowerCase() === contractor.name.toLowerCase()
  );
}

function attendanceStatsForContractor(contractor, callouts, attendanceRecords = []) {
  const records = attendanceRecordsForContractor(contractor, attendanceRecords);
  const marks = records.flatMap((record) => attendanceMarks(record)).filter((mark) => mark && mark !== "WO");
  const tardy = marks.filter((mark) => mark === "T").length;
  const absences = marks.filter((mark) => mark === "A").length;
  const present = marks.filter((mark) => mark === "P" || mark === "T").length;
  if (marks.length) {
    return {
      attendance: Math.max(0, Math.min(100, Math.round((present / marks.length) * 100))),
      absences,
      tardy,
      trackedDays: marks.length
    };
  }

  const scheduled = scheduledDaysSinceStart(contractor);
  if (!scheduled) return { attendance: 100, absences: Number(contractor.noShows) || 0, tardy: Number(contractor.lateArrivals) || 0, trackedDays: 0 };
  const fallbackAbsences = calloutsForContractor(contractor, callouts).length + (Number(contractor.noShows) || 0);
  return {
    attendance: Math.max(0, Math.min(100, Math.round(((scheduled - fallbackAbsences) / scheduled) * 100))),
    absences: fallbackAbsences,
    tardy: Number(contractor.lateArrivals) || 0,
    trackedDays: scheduled
  };
}

function attendancePercent(contractor, callouts, attendanceRecords = []) {
  return attendanceStatsForContractor(contractor, callouts, attendanceRecords).attendance;
}

function weeklyAttendancePercent(record, status) {
  const values = [record.monday, record.tuesday, record.wednesday, record.thursday, record.friday].filter((value) => value && value !== "WO");
  if (!values.length) return 0;
  return Math.round((values.filter((value) => value === status).length / values.length) * 100);
}

function escapeValue(value) {
  return String(value ?? "").replaceAll('"', '""');
}

function compareText(a, b) {
  return String(a ?? "").trim().localeCompare(String(b ?? "").trim(), undefined, { numeric: true, sensitivity: "base" });
}

function compareDate(a, b) {
  const aTime = Date.parse(a || "");
  const bTime = Date.parse(b || "");
  if (Number.isNaN(aTime) && Number.isNaN(bTime)) return 0;
  if (Number.isNaN(aTime)) return 1;
  if (Number.isNaN(bTime)) return -1;
  return aTime - bTime;
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

function downloadJson(filename, data) {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];
    if (char === '"' && quoted && next === '"') {
      cell += '"';
      index += 1;
    } else if (char === '"') {
      quoted = !quoted;
    } else if (char === "," && !quoted) {
      row.push(cell.trim());
      cell = "";
    } else if ((char === "\n" || char === "\r") && !quoted) {
      if (char === "\r" && next === "\n") index += 1;
      row.push(cell.trim());
      if (row.some(Boolean)) rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += char;
    }
  }

  row.push(cell.trim());
  if (row.some(Boolean)) rows.push(row);
  return rows;
}

function normalizeHeader(header) {
  return String(header).toLowerCase().replace(/[^a-z0-9]/g, "");
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
  const { contractors, callouts, tickets, laptops, vehicles, notes, attendanceRecords = [], stats } = data;
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
        ${tableHtml("Contractors", ["Name", "Company", "Start Date", "Shift", "Status", "Attendance", "No Shows", "Late", "Training"], contractors.map((c) => {
          const attendanceStats = attendanceStatsForContractor(c, callouts, attendanceRecords);
          return [c.name, c.company, c.startDate, shiftLabel(c.shift), c.status, `${attendanceStats.attendance}%`, attendanceStats.absences, attendanceStats.tardy, `${c.trainingCompletion ?? 100}%`];
        }))}
        ${tableHtml("Callout Log", ["Date", "Time", "Name", "Company", "Shift", "Reason", "Hours Before Shift", "Coverage Status", "Replacement", "Status"], callouts.map((c) => [c.date, c.submittedTime, c.name, c.company, shiftLabel(c.shift), c.reason, c.hoursBeforeShiftStart, c.coverageStatus, c.replacementAssigned, c.status]))}
        ${tableHtml("Coverage Tickets", ["Date", "Shift", "Contractor", "Status", "Replacement", "Response Time", "Notes"], tickets.map((t) => [t.date, shiftLabel(t.shift), t.contractorName, t.status, t.replacementAssigned, t.responseTimeMinutes ? `${t.responseTimeMinutes} min` : "", t.notes]))}
        ${tableHtml("Laptop Accountability", ["Asset", "Contractor", "Status", "Assigned By", "Assigned Date", "Returned By", "Returned Date", "Due Date"], laptops.map((l) => [l.asset, l.contractor, l.status, l.assignedBy, l.assignedDate, l.returnedBy, l.returnedDate, l.dueDate]))}
        ${tableHtml("Vehicles", ["Name", "VIN", "Status", "License Plate", "Model", "Year", "Location", "Assigned Contractor"], vehicles.map((v) => [v.name, v.vin, v.status, v.licensePlate, v.model, v.year, v.location, v.assignedContractorName]))}
        ${tableHtml("Weekly Attendance", ["Week Start", "Name", "Company", "Shift", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "P%", "A%", "Tardy %", "Comments"], attendanceRecords.map((r) => [r.weekStart, r.contractorName, r.company, shiftLabel(r.shift), r.monday, r.tuesday, r.wednesday, r.thursday, r.friday, `${weeklyAttendancePercent(r, "P")}%`, `${weeklyAttendancePercent(r, "A")}%`, `${weeklyAttendancePercent(r, "T")}%`, r.comments]))}
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
  const isRecoveryLink = new URLSearchParams(window.location.search).get("reset") === "1" || window.location.hash.includes("type=recovery");
  const calloutLink = useMemo(() => {
    const url = new URL(window.location.href);
    url.search = "";
    url.hash = "";
    url.searchParams.set("callout", "1");
    return url.toString();
  }, []);
  const [role, setRole] = useState(calloutOnly ? "operator" : "admin");
  const [contractors, setContractors] = useState(() => readSaved("contractors", starterContractors).map((c) => resetDailyStatusForDate({ ...c, site: siteFor(c), startDate: c.startDate ?? today, shift: normalizeShift(c.shift), status: c.status ?? "Scheduled", statusDate: c.statusDate ?? statusDateFor(c.id), attendance: c.attendance ?? 95, noShows: c.noShows ?? 0, lateArrivals: c.lateArrivals ?? 0, attendanceDays: c.attendanceDays ?? 0, trainingCompletion: c.trainingCompletion ?? 100 })));
  const [callouts, setCallouts] = useState(() => readSaved("callouts", []).map((item) => ({ ...item, site: siteFor(item) })));
  const [tickets, setTickets] = useState(() => readSaved("coverageTickets", []).map((item) => ({ ...item, site: siteFor(item) })));
  const [laptops, setLaptops] = useState(() => readSaved("laptops", starterLaptops).map((item) => ({ ...item, site: siteFor(item) })));
  const [vehicles, setVehicles] = useState(() => readSaved("vehicles", starterVehicles).map((item) => ({ ...item, site: siteFor(item) })));
  const [attendanceRecords, setAttendanceRecords] = useState(() => readSaved("attendanceRecords", []).map((item) => ({ ...item, site: siteFor(item), shift: normalizeShift(item.shift) })));
  const [notes, setNotes] = useState(() => readSaved("opsNotes", [{ id: crypto.randomUUID(), site: defaultSite, date: today, text: "AM shift short 2 operators." }]).map((item) => ({ ...item, site: siteFor(item) })));
  const [notifications, setNotifications] = useState(() => readSaved("opsNotifications", []).map((item) => ({ ...item, site: siteFor(item) })));
  const [dbStatus, setDbStatus] = useState(hasSupabaseConfig ? "Connecting to Supabase" : "Local browser save only");
  const [adminSession, setAdminSession] = useState(null);
  const [localDisplayName, setLocalDisplayName] = useState(() => localStorage.getItem("adminDisplayName") || "Admin User");
  const [theme, setTheme] = useState(() => localStorage.getItem("fleetopsTheme") || "dark");
  const [authReady, setAuthReady] = useState(!hasSupabaseConfig);
  const [passwordResetMode, setPasswordResetMode] = useState(isRecoveryLink);

  useEffect(() => {
    if (!hasSupabaseConfig) return;
    async function loadSession() {
      const { data } = await supabase.auth.getSession();
      setAdminSession(data.session);
      setAuthReady(true);
    }
    const { data: listener } = supabase.auth.onAuthStateChange((event, session) => {
      setAdminSession(session);
      if (event === "PASSWORD_RECOVERY") setPasswordResetMode(true);
      setAuthReady(true);
    });
    loadSession();
    return () => listener.subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (!hasSupabaseConfig) return;
    async function loadTable(table, key, setter, mapper, order) {
      let query = supabase.from(table).select("*");
      if (order) query = query.order(order, { ascending: false });
      const { data, error } = await query;
      if (error) {
        setDbStatus(`Supabase ${table} error: ${error.message}`);
        return;
      }
      const next = (data ?? []).map(mapper);
      setter(next);
      localStorage.setItem(key, JSON.stringify(next));
      setDbStatus("Supabase connected");
    }
    loadTable("contractors", "contractors", setContractors, dbToContractor, "name");
    loadTable("callouts", "callouts", setCallouts, dbToCallout, "submitted_at");
    loadTable("coverage_tickets", "coverageTickets", setTickets, dbToTicket, "opened_at");
    loadTable("laptops", "laptops", setLaptops, dbToLaptop, "assigned_date");
    loadTable("vehicles", "vehicles", setVehicles, dbToVehicle, "name");
    loadTable("attendance_records", "attendanceRecords", setAttendanceRecords, dbToAttendance, "week_start");
    loadTable("operations_notes", "opsNotes", setNotes, dbToNote, "note_date");
    loadTable("operations_notifications", "opsNotifications", setNotifications, dbToNotification, "created_at");
  }, []);

  useEffect(() => {
    const resetContractors = contractors.map((contractor) => resetDailyStatusForDate(contractor));
    if (resetContractors.some((contractor, index) => contractor.status !== contractors[index].status || contractor.statusDate !== contractors[index].statusDate)) {
      saveContractors(resetContractors);
    }
  }, [contractors]);

  async function syncRows(table, key, current, next, setter, mapper, label) {
    save(key, next, setter);
    if (!hasSupabaseConfig) return;

    const removedIds = current.filter((item) => !next.some((nextItem) => nextItem.id === item.id)).map((item) => item.id);
    if (removedIds.length) {
      const { error } = await supabase.from(table).delete().in("id", removedIds);
      if (error) {
        setDbStatus(`Supabase ${label} delete error: ${error.message}`);
        return;
      }
    }

    if (!next.length) {
      setDbStatus(`Supabase ${label} synced`);
      return;
    }

    const { error } = await supabase.from(table).upsert(next.map(mapper));
    setDbStatus(error ? `Supabase ${label} save error: ${error.message}` : `Supabase ${label} synced`);
  }

  async function saveContractors(next) {
    saveStatusDates(next);
    syncRows("contractors", "contractors", contractors, next, setContractors, contractorToDb, "contractors");
  }

  function saveCallouts(next) {
    syncRows("callouts", "callouts", callouts, next, setCallouts, calloutToDb, "callouts");
  }

  function saveTickets(next) {
    syncRows("coverage_tickets", "coverageTickets", tickets, next, setTickets, ticketToDb, "coverage tickets");
  }

  function saveLaptops(next) {
    syncRows("laptops", "laptops", laptops, next, setLaptops, laptopToDb, "laptops");
  }

  function saveVehicles(next) {
    syncRows("vehicles", "vehicles", vehicles, next, setVehicles, vehicleToDb, "vehicles");
  }

  function saveAttendanceRecords(next) {
    syncRows("attendance_records", "attendanceRecords", attendanceRecords, next, setAttendanceRecords, attendanceToDb, "attendance records");
  }

  function saveNotes(next) {
    syncRows("operations_notes", "opsNotes", notes, next, setNotes, noteToDb, "operations notes");
  }

  function saveNotifications(next) {
    syncRows("operations_notifications", "opsNotifications", notifications, next, setNotifications, notificationToDb, "notifications");
  }

  function pushNotification(message, site = defaultSite) {
    const next = [{ id: crypto.randomUUID(), site, time: new Date().toLocaleTimeString(), message }, ...notifications];
    saveNotifications(next);
  }

  function submitCallout(callout) {
    const site = siteFor(callout);
    const contractor = contractors.find((c) => siteFor(c) === site && c.name.toLowerCase() === callout.name.toLowerCase());
    const shift = normalizeShift(callout.shift);
    const submitted = new Date();
    const start = new Date(`${callout.date}T${String(shiftStart[shift]).padStart(2, "0")}:00:00`);
    const hoursBefore = Math.max(0, Math.round((start - submitted) / 36e5));
    const enriched = {
      id: crypto.randomUUID(),
      site,
      submittedAt: submitted.toISOString(),
      submittedTime: submitted.toLocaleTimeString(),
      contractorId: contractor?.id ?? "",
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
      site,
      calloutId: enriched.id,
      contractorId: contractor?.id ?? "",
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
    saveCallouts([enriched, ...callouts]);
    saveTickets([ticket, ...tickets]);
    pushNotification(`Callout submitted: ${callout.name} / ${site} / ${shift}. Coverage ticket created.`, site);
  }

  const state = { contractors, callouts, tickets, laptops, vehicles, attendanceRecords, notes, notifications };
  const actions = { saveContractors, saveCallouts, saveTickets, saveLaptops, saveVehicles, saveAttendanceRecords, saveNotes, saveNotifications, pushNotification, submitCallout };
  const adminAllowed = !hasSupabaseConfig || Boolean(adminSession);

  async function signOutAdmin() {
    if (hasSupabaseConfig) await supabase.auth.signOut();
    setAdminSession(null);
    setRole("operator");
  }

  function updateLocalDisplayName(name) {
    setLocalDisplayName(name);
    localStorage.setItem("adminDisplayName", name);
  }

  function toggleTheme() {
    setTheme((current) => {
      const next = current === "light" ? "dark" : "light";
      localStorage.setItem("fleetopsTheme", next);
      return next;
    });
  }

  return (
    <main className="app-shell" data-theme={theme}>
      <header className="topbar">
        <div className="brand-lockup">
          <img className="brand-logo" src="/rivian-mark.webp" alt="Rivian" />
          <div>
            <h1>FleetOps Command Center</h1>
          </div>
        </div>
        {!calloutOnly && (
          <div className="top-actions">
            <button className="theme-toggle" type="button" onClick={toggleTheme}>{theme === "light" ? "Dark Mode" : "Light Mode"}</button>
            <div className="role-switch" aria-label="Role switcher">
              <button className={role === "operator" ? "active" : ""} onClick={() => setRole("operator")}>Operator</button>
              <button className={role === "admin" ? "active" : ""} onClick={() => setRole("admin")}>Admin</button>
            </div>
            {(adminAllowed && role === "admin") || adminSession ? (
              <div className="account-actions">
                {adminAllowed && role === "admin" && <span className="profile-chip">{hasSupabaseConfig ? displayNameForSession(adminSession) : localDisplayName}</span>}
                {adminSession && <button className="logout-button" onClick={signOutAdmin}>Logout</button>}
              </div>
            ) : null}
          </div>
        )}
      </header>

      {role === "operator" ? (
        <OperatorForm contractors={contractors} onSubmit={submitCallout} calloutLink={calloutLink} showShareLink={!calloutOnly} />
      ) : !authReady ? (
        <section className="panel narrow"><p>Checking admin access...</p></section>
      ) : passwordResetMode ? (
        <AdminLogin resetMode onLoggedIn={setAdminSession} onResetComplete={() => setPasswordResetMode(false)} />
      ) : adminAllowed ? (
        <AdminCommandCenter state={state} actions={actions} dbStatus={dbStatus} session={adminSession} localDisplayName={localDisplayName} onSessionChange={setAdminSession} onLocalDisplayNameChange={updateLocalDisplayName} />
      ) : (
        <AdminLogin onLoggedIn={setAdminSession} />
      )}
    </main>
  );
}

function AdminLogin({ onLoggedIn, resetMode = false, onResetComplete = () => {} }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [newPassword, setNewPassword] = useState("");
  const [error, setError] = useState("");
  const [message, setMessage] = useState("");
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    if (!resetMode) return;
    const cleanUrl = new URL(window.location.href);
    cleanUrl.searchParams.delete("reset");
    window.history.replaceState({}, "", cleanUrl.toString());
  }, [resetMode]);

  async function submit(event) {
    event.preventDefault();
    setBusy(true);
    setError("");
    setMessage("");
    const { data, error } = await supabase.auth.signInWithPassword({ email, password });
    setBusy(false);
    if (error) {
      setError(error.message);
      return;
    }
    onLoggedIn(data.session);
  }

  async function sendResetEmail() {
    if (!email.trim()) {
      setError("Enter your admin email first.");
      return;
    }
    setBusy(true);
    setError("");
    setMessage("");
    const { error } = await supabase.auth.resetPasswordForEmail(email, {
      redirectTo: "https://callout-tracker.vercel.app?reset=1"
    });
    setBusy(false);
    if (error) {
      setError(error.message);
      return;
    }
    setMessage("Password reset email sent. Check your inbox.");
  }

  async function updatePassword(event) {
    event.preventDefault();
    setBusy(true);
    setError("");
    setMessage("");
    const { data, error } = await supabase.auth.updateUser({ password: newPassword });
    setBusy(false);
    if (error) {
      setError(error.message);
      return;
    }
    setMessage("Password updated. You are signed in.");
    onResetComplete();
    onLoggedIn(data.user ? (await supabase.auth.getSession()).data.session : null);
  }

  return (
    <section className="login-shell">
      <div className="panel login-panel">
        <div className="section-heading">
          <div>
            <p className="eyebrow">Admin Access</p>
            <h2>Operations Login</h2>
          </div>
        </div>
        <form className="form-grid compact" onSubmit={resetMode ? updatePassword : submit}>
          {!resetMode && <label>Email<input type="email" value={email} onChange={(e) => setEmail(e.target.value)} required /></label>}
          {resetMode ? (
            <label>New Password<input type="password" value={newPassword} onChange={(e) => setNewPassword(e.target.value)} minLength="6" required /></label>
          ) : (
            <label>Password<input type="password" value={password} onChange={(e) => setPassword(e.target.value)} required /></label>
          )}
          {error && <p className="login-error">{error}</p>}
          {message && <p className="login-success">{message}</p>}
          <button className="primary-button" disabled={busy}>{busy ? "Working..." : resetMode ? "Update Password" : "Login"}</button>
          {!resetMode && <button className="ghost-button" type="button" onClick={sendResetEmail} disabled={busy}>Forgot password?</button>}
        </form>
      </div>
    </section>
  );
}

function UserProfile({ session, localDisplayName, onSessionChange, onLocalDisplayNameChange }) {
  const currentName = hasSupabaseConfig ? displayNameForSession(session) : localDisplayName;
  const [displayName, setDisplayName] = useState(currentName);
  const [password, setPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    setDisplayName(currentName);
  }, [currentName]);

  async function saveDisplayName(event) {
    event.preventDefault();
    const name = displayName.trim();
    if (!name) {
      setError("Display name is required.");
      return;
    }
    setBusy(true);
    setError("");
    setMessage("");
    if (!hasSupabaseConfig) {
      onLocalDisplayNameChange(name);
      setMessage("Display name saved.");
      setBusy(false);
      return;
    }
    const { error: updateError } = await supabase.auth.updateUser({ data: { display_name: name } });
    if (updateError) {
      setError(updateError.message);
      setBusy(false);
      return;
    }
    const { data } = await supabase.auth.getSession();
    onSessionChange(data.session);
    setMessage("Profile updated.");
    setBusy(false);
  }

  async function changePassword(event) {
    event.preventDefault();
    setError("");
    setMessage("");
    if (!hasSupabaseConfig) {
      setError("Password changes require Supabase authentication.");
      return;
    }
    if (password.length < 6) {
      setError("Password must be at least 6 characters.");
      return;
    }
    if (password !== confirmPassword) {
      setError("Passwords do not match.");
      return;
    }
    setBusy(true);
    const { error: updateError } = await supabase.auth.updateUser({ password });
    if (updateError) {
      setError(updateError.message);
      setBusy(false);
      return;
    }
    setPassword("");
    setConfirmPassword("");
    setMessage("Password updated.");
    setBusy(false);
  }

  return (
    <section className="profile-layout">
      <article className="panel profile-card">
        <div className="profile-avatar">{currentName.slice(0, 1).toUpperCase()}</div>
        <div>
          <p className="eyebrow">Account</p>
          <h2>{currentName}</h2>
          <span>{session?.user?.email ?? "Local admin profile"}</span>
        </div>
      </article>

      <section className="panel">
        <div className="section-heading"><div><p className="eyebrow">Profile Settings</p><h2>Display Name</h2></div></div>
        <form className="form-grid compact" onSubmit={saveDisplayName}>
          <label>Display Name<input value={displayName} onChange={(e) => setDisplayName(e.target.value)} placeholder="Operations Manager" required /></label>
          <button className="primary-button" disabled={busy}>{busy ? "Saving..." : "Save Display Name"}</button>
        </form>
      </section>

      <section className="panel">
        <div className="section-heading"><div><p className="eyebrow">Security</p><h2>Change Password</h2></div></div>
        <form className="form-grid compact" onSubmit={changePassword}>
          <label>New Password<input type="password" value={password} onChange={(e) => setPassword(e.target.value)} minLength="6" autoComplete="new-password" /></label>
          <label>Confirm Password<input type="password" value={confirmPassword} onChange={(e) => setConfirmPassword(e.target.value)} minLength="6" autoComplete="new-password" /></label>
          <button className="primary-button" disabled={busy || !hasSupabaseConfig}>{busy ? "Updating..." : "Update Password"}</button>
        </form>
      </section>

      {(error || message) && <section className="profile-feedback">{error ? <p className="login-error">{error}</p> : <p className="login-success">{message}</p>}</section>}
    </section>
  );
}

function OperatorForm({ contractors, onSubmit, calloutLink, showShareLink }) {
  const [form, setForm] = useState({ site: defaultSite, name: "", date: today, shift: "AM Shift", reason: "Sick", notes: "" });
  const [sent, setSent] = useState(false);
  const [copied, setCopied] = useState(false);
  const activeNames = contractors.filter((c) => isOperationsContractor(c) && siteFor(c) === form.site).map((c) => c.name);

  function update(field, value) {
    setForm((current) => ({ ...current, [field]: value }));
    setSent(false);
  }

  function handleSubmit(event) {
    event.preventDefault();
    if (!form.name.trim() || !form.reason.trim()) return;
    onSubmit(form);
    setForm({ site: form.site, name: "", date: today, shift: "AM Shift", reason: "Sick", notes: "" });
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
          <label>Site<select value={form.site} onChange={(event) => update("site", event.target.value)}>{sites.map((site) => <option key={site}>{site}</option>)}</select></label>
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

function AdminCommandCenter({ state, actions, dbStatus, session, localDisplayName, onSessionChange, onLocalDisplayNameChange }) {
  const [page, setPage] = useState("Daily Ops");
  const [selectedSite, setSelectedSite] = useState(defaultSite);
  const [selectedOpsShift, setSelectedOpsShift] = useState(operationShifts[0].value);
  const [currentDate, setCurrentDate] = useState(localDateString());
  const currentDay = days[new Date(`${currentDate}T00:00:00`).getDay()];
  const pages = ["Daily Ops", "Weekly Attendance", "Calendar", "Contractors", "Callouts", "Vehicles", "Performance", "Reports", "Notes", "Notifications", "AI Assistant", "Executive", "Profile"];
  const pageLabels = { "Daily Ops": "Overview", "Weekly Attendance": "Attendance", Notifications: "Alerts", "AI Assistant": "AI", Contractors: "Staffing" };
  const pageIcons = { "Daily Ops": "▦", "Weekly Attendance": "◴", Calendar: "□", Contractors: "♙", Callouts: "☎", Vehicles: "▰", Performance: "↗", Reports: "▧", Notes: "✎", Notifications: "△", "AI Assistant": "✦", Executive: "◆", Profile: "⚙" };
  const siteState = useMemo(() => Object.fromEntries(Object.entries(state).map(([key, value]) => [key, Array.isArray(value) ? value.filter((item) => siteFor(item) === selectedSite) : value])), [state, selectedSite]);
  const shiftState = useMemo(() => {
    const shiftContractors = siteState.contractors.filter((contractor) => normalizeShift(contractor.shift) === selectedOpsShift).map((contractor) => resetDailyStatusForDate(contractor, currentDate));
    const shiftContractorIds = new Set(shiftContractors.map((contractor) => contractor.id));
    const shiftContractorNames = new Set(shiftContractors.map((contractor) => contractor.name.toLowerCase()));
    return {
      ...siteState,
      contractors: shiftContractors,
      callouts: siteState.callouts.filter((callout) => normalizeShift(callout.shift) === selectedOpsShift || shiftContractorNames.has(callout.name?.toLowerCase())),
      tickets: siteState.tickets.filter((ticket) => normalizeShift(ticket.shift) === selectedOpsShift || shiftContractorNames.has(ticket.contractorName?.toLowerCase())),
      attendanceRecords: siteState.attendanceRecords.filter((record) => normalizeShift(record.shift) === selectedOpsShift),
      laptops: siteState.laptops.filter((laptop) => !laptop.contractorId || shiftContractorIds.has(laptop.contractorId)),
      notifications: siteState.notifications.filter((notification) => !notification.message || notification.message.includes(selectedOpsShift) || !shifts.some((shift) => notification.message.includes(shift))),
      notes: siteState.notes.filter((note) => !note.text || note.text.includes(selectedOpsShift) || !shifts.some((shift) => note.text.includes(shift)))
    };
  }, [siteState, selectedOpsShift, currentDate]);
  const siteActions = useMemo(() => {
    const scopedSave = (key, saveFn) => (next) => saveFn([
      ...state[key].filter((item) => siteFor(item) !== selectedSite),
      ...next.map((item) => ({ ...item, site: selectedSite }))
    ]);
    return {
      ...actions,
      saveContractors: scopedSave("contractors", actions.saveContractors),
      saveCallouts: scopedSave("callouts", actions.saveCallouts),
      saveTickets: scopedSave("tickets", actions.saveTickets),
      saveLaptops: scopedSave("laptops", actions.saveLaptops),
      saveVehicles: scopedSave("vehicles", actions.saveVehicles),
      saveAttendanceRecords: scopedSave("attendanceRecords", actions.saveAttendanceRecords),
      saveNotes: scopedSave("notes", actions.saveNotes),
      saveNotifications: scopedSave("notifications", actions.saveNotifications),
      pushNotification: (message) => actions.pushNotification(message, selectedSite),
      submitCallout: (callout) => actions.submitCallout({ ...callout, site: selectedSite })
    };
  }, [actions, state, selectedSite]);
  const shiftActions = useMemo(() => {
    const scopedSave = (key, saveFn, filterFn) => (next) => saveFn([
      ...siteState[key].filter((item) => !filterFn(item)),
      ...next.map((item) => ({ ...item, site: selectedSite }))
    ]);
    const inShiftByName = (name) => shiftState.contractors.some((contractor) => contractor.name.toLowerCase() === name?.toLowerCase());
    const inShiftById = (id) => shiftState.contractors.some((contractor) => contractor.id === id);
    return {
      ...siteActions,
      saveContractors: scopedSave("contractors", siteActions.saveContractors, (item) => normalizeShift(item.shift) === selectedOpsShift),
      saveCallouts: scopedSave("callouts", siteActions.saveCallouts, (item) => normalizeShift(item.shift) === selectedOpsShift || inShiftByName(item.name)),
      saveTickets: scopedSave("tickets", siteActions.saveTickets, (item) => normalizeShift(item.shift) === selectedOpsShift || inShiftByName(item.contractorName)),
      saveAttendanceRecords: scopedSave("attendanceRecords", siteActions.saveAttendanceRecords, (item) => normalizeShift(item.shift) === selectedOpsShift),
      saveLaptops: scopedSave("laptops", siteActions.saveLaptops, (item) => !item.contractorId || inShiftById(item.contractorId)),
      saveNotes: scopedSave("notes", siteActions.saveNotes, (item) => !item.text || item.text.includes(selectedOpsShift) || !shifts.some((shift) => item.text.includes(shift))),
      saveNotifications: scopedSave("notifications", siteActions.saveNotifications, (item) => !item.message || item.message.includes(selectedOpsShift) || !shifts.some((shift) => item.message.includes(shift))),
      pushNotification: (message) => siteActions.pushNotification(`${selectedOpsShift}: ${message}`),
      submitCallout: (callout) => siteActions.submitCallout({ ...callout, shift: selectedOpsShift }),
      opsShift: selectedOpsShift
    };
  }, [siteActions, siteState, shiftState.contractors, selectedSite, selectedOpsShift]);
  useEffect(() => {
    const timer = window.setInterval(() => setCurrentDate(localDateString()), 60000);
    return () => window.clearInterval(timer);
  }, []);

  const metrics = useMemo(() => buildMetrics(shiftState, currentDate, currentDay), [shiftState, currentDate, currentDay]);
  const selectedOpsLabel = operationShifts.find((shift) => shift.value === selectedOpsShift)?.label ?? selectedOpsShift;
  const operationsState = useMemo(() => {
    const operationsContractors = shiftState.contractors.filter(isOperationsContractor);
    const operationIds = new Set(operationsContractors.map((contractor) => contractor.id));
    const operationNames = new Set(operationsContractors.map((contractor) => contractor.name.toLowerCase()));
    return {
      ...shiftState,
      contractors: operationsContractors,
      callouts: shiftState.callouts.filter((callout) => operationNames.has(callout.name?.toLowerCase())),
      tickets: shiftState.tickets.filter((ticket) => operationNames.has(ticket.contractorName?.toLowerCase())),
      attendanceRecords: shiftState.attendanceRecords.filter((record) => operationIds.has(record.contractorId) || operationNames.has(record.contractorName?.toLowerCase())),
      vehicles: shiftState.vehicles.filter((vehicle) => (!vehicle.assignedContractorId && !vehicle.assignedContractorName) || operationIds.has(vehicle.assignedContractorId) || operationNames.has(vehicle.assignedContractorName?.toLowerCase()))
    };
  }, [shiftState]);
  const operationsMetrics = useMemo(() => buildMetrics(operationsState, currentDate, currentDay), [operationsState, currentDate, currentDay]);
  const operationsActions = useMemo(() => {
    const operationIds = new Set(operationsState.contractors.map((contractor) => contractor.id));
    const operationNames = new Set(operationsState.contractors.map((contractor) => contractor.name.toLowerCase()));
    const hiddenContractors = shiftState.contractors.filter((contractor) => !operationIds.has(contractor.id));
    return {
      ...shiftActions,
      saveContractors: (next) => shiftActions.saveContractors([...hiddenContractors, ...next]),
      saveCallouts: (next) => shiftActions.saveCallouts([...shiftState.callouts.filter((callout) => !operationNames.has(callout.name?.toLowerCase())), ...next]),
      saveTickets: (next) => shiftActions.saveTickets([...shiftState.tickets.filter((ticket) => !operationNames.has(ticket.contractorName?.toLowerCase())), ...next]),
      saveAttendanceRecords: (next) => shiftActions.saveAttendanceRecords([...shiftState.attendanceRecords.filter((record) => !operationIds.has(record.contractorId) && !operationNames.has(record.contractorName?.toLowerCase())), ...next]),
      saveVehicles: (next) => shiftActions.saveVehicles([...shiftState.vehicles.filter((vehicle) => (vehicle.assignedContractorId && !operationIds.has(vehicle.assignedContractorId)) || (vehicle.assignedContractorName && !operationNames.has(vehicle.assignedContractorName.toLowerCase()))), ...next])
    };
  }, [shiftActions, shiftState, operationsState]);

  return (
    <div className="dashboard">
      <div className="dashboard-shell">
        <nav className="ops-nav" aria-label="Operations views">
          {pages.map((item) => <button key={item} className={page === item ? "active" : ""} onClick={() => setPage(item)} title={item}><span className="nav-icon">{pageIcons[item] ?? "•"}</span><span className="nav-label">{pageLabels[item] ?? item}</span></button>)}
        </nav>
        <section className="dashboard-content">
          <section className="site-banner">
            <div>
              <p className="eyebrow">Site Dashboard</p>
              <h2>{selectedSite} / {selectedOpsLabel}</h2>
            </div>
            <div className="dashboard-switches">
              <div className="site-switch" aria-label="Site switcher">
                {sites.map((site) => <button key={site} className={selectedSite === site ? "active" : ""} onClick={() => setSelectedSite(site)}>{site}</button>)}
              </div>
              <div className="site-switch shift-switch" aria-label="Operations shift switcher">
                {operationShifts.map((shift) => <button key={shift.value} className={selectedOpsShift === shift.value ? "active" : ""} onClick={() => setSelectedOpsShift(shift.value)}>{shift.label}</button>)}
              </div>
            </div>
          </section>
          <p className={`db-status ${dbStatusClass(dbStatus)}`}>{dbStatus}</p>
          {page === "Daily Ops" && <DailyOperations state={operationsState} actions={operationsActions} metrics={operationsMetrics} currentDate={currentDate} />}
          {page === "Weekly Attendance" && <WeeklyAttendanceTracker state={operationsState} actions={operationsActions} opsShift={selectedOpsShift} />}
      {page === "Calendar" && <WeeklyCalendarView state={operationsState} actions={operationsActions} currentDate={currentDate} />}
          {page === "Callouts" && <CalloutCenter state={operationsState} actions={operationsActions} />}
          {page === "Vehicles" && <VehicleModule vehicles={operationsState.vehicles} contractors={operationsState.contractors} saveVehicles={operationsActions.saveVehicles} />}
          {page === "Performance" && <PerformanceScorecard contractors={operationsState.contractors} callouts={operationsState.callouts} attendanceRecords={operationsState.attendanceRecords} />}
          {page === "Reports" && <ReportsGenerator state={operationsState} actions={operationsActions} />}
          {page === "Notes" && <OperationsNotes notes={shiftState.notes} saveNotes={shiftActions.saveNotes} opsShift={selectedOpsShift} />}
          {page === "Notifications" && <NotificationsFeed notifications={shiftState.notifications} saveNotifications={shiftActions.saveNotifications} />}
          {page === "AI Assistant" && <AiAssistant state={operationsState} metrics={operationsMetrics} />}
          {page === "Executive" && <ExecutiveDashboard state={operationsState} metrics={operationsMetrics} />}
          {page === "Profile" && <UserProfile session={session} localDisplayName={localDisplayName} onSessionChange={onSessionChange} onLocalDisplayNameChange={onLocalDisplayNameChange} />}
          {page === "Contractors" && <ContractorAdmin contractors={shiftState.contractors} callouts={shiftState.callouts} onSaveContractors={shiftActions.saveContractors} defaultShift={selectedOpsShift} />}
        </section>
      </div>
    </div>
  );
}

function buildMetrics({ contractors, callouts, tickets, laptops, attendanceRecords = [] }, currentDate = today, currentDay = todayDay) {
  const active = contractors.filter(isOperationsContractor);
  const scheduled = active.filter((c) => c.scheduledDays.includes(currentDay));
  const todayCallouts = callouts.filter((c) => dateOnly(c.date) === currentDate);
  const openTickets = tickets.filter((t) => ["Pending", "Coverage Needed"].includes(t.status));
  const coveredTickets = tickets.filter((t) => t.status === "Covered" || t.status === "Closed");
  const overdue = laptops.filter((l) => l.status === "Overdue").length;
  const returned = laptops.filter((l) => l.status === "Returned").length;
  const assigned = laptops.filter((l) => l.status === "Assigned" || l.status === "Overdue").length;
  const available = Math.max(0, scheduled.length - todayCallouts.length);
  const coveragePct = scheduled.length ? Math.round((available / scheduled.length) * 100) : 100;
  const attendance = active.length ? Math.round(active.reduce((sum, c) => sum + attendancePercent(c, callouts, attendanceRecords), 0) / active.length) : 100;
  const laptopCompliance = assigned + returned ? Math.round((returned / (assigned + returned)) * 100) : 100;
  const readiness = Math.max(0, Math.round((coveragePct * 0.45) + (attendance * 0.3) + (laptopCompliance * 0.2) - (openTickets.length * 2)));
  const avgCoverage = coveredTickets.length ? Math.round(coveredTickets.reduce((sum, t) => sum + (Number(t.responseTimeMinutes) || 0), 0) / coveredTickets.length) : 0;
  return { active, scheduled, todayCallouts, openTickets, coveredTickets, overdue, returned, assigned, available, coveragePct, attendance, laptopCompliance, readiness, avgCoverage };
}

function Metric({ label, value }) {
  return <article className="metric-card"><p>{label}</p><strong>{value}</strong></article>;
}

function DailyOperations({ state, actions, metrics, currentDate = today }) {
  const level = metrics.readiness >= 90 ? "Healthy" : metrics.readiness >= 75 ? "Watch" : "Critical";
  const [showReadinessBreakdown, setShowReadinessBreakdown] = useState(false);
  const coverageScore = Math.round(metrics.coveragePct * 0.45);
  const attendanceScore = Math.round(metrics.attendance * 0.3);
  const laptopScore = Math.round(metrics.laptopCompliance * 0.2);
  const coveragePenalty = metrics.openTickets.length * 2;
  return (
    <>
      <section className="metric-grid">
        <Metric label="Total Scheduled" value={metrics.scheduled.length} />
        <Metric label="Total Available" value={metrics.available} />
        <Metric label="Total Callouts" value={metrics.todayCallouts.length} />
        <Metric label="Open Coverage Gaps" value={metrics.openTickets.length} />
        <Metric label="Contractors In Training" value={state.contractors.filter((c) => isOperationsContractor(c) && c.status === "Training").length} />
        <Metric label="Laptop Issues" value={metrics.overdue} />
      </section>
      <button className="panel readiness-panel readiness-button" type="button" onClick={() => setShowReadinessBreakdown((current) => !current)} aria-expanded={showReadinessBreakdown}>
        <div><p className="eyebrow">Operational Readiness / {currentDate}</p><h2>{metrics.readiness}%</h2></div>
        <div className="readiness-side"><span className={`readiness ${level.toLowerCase()}`}>{level}</span><small>{showReadinessBreakdown ? "Hide calculation" : "View calculation"}</small></div>
      </button>
      {showReadinessBreakdown && (
        <section className="panel readiness-breakdown">
          <div className="section-heading"><div><p className="eyebrow">Calculation</p><h2>How readiness is calculated</h2></div></div>
          <div className="readiness-formula">
            <article><span>Coverage</span><strong>{metrics.coveragePct}% x 45% = {coverageScore}</strong><p>{metrics.available} available / {metrics.scheduled.length} scheduled</p></article>
            <article><span>Attendance</span><strong>{metrics.attendance}% x 30% = {attendanceScore}</strong><p>Average attendance across {metrics.active.length} active contractors</p></article>
            <article><span>Laptop Compliance</span><strong>{metrics.laptopCompliance}% x 20% = {laptopScore}</strong><p>{metrics.returned} returned / {metrics.assigned + metrics.returned} tracked laptop records</p></article>
            <article><span>Coverage Gap Penalty</span><strong>-{coveragePenalty}</strong><p>{metrics.openTickets.length} open ticket(s) x 2 points</p></article>
          </div>
          <p className="readiness-equation">Readiness = {coverageScore} + {attendanceScore} + {laptopScore} - {coveragePenalty} = {metrics.readiness}%</p>
        </section>
      )}
      <section className="daily-command-grid">
        <AvailabilityBoard contractors={state.contractors} callouts={state.callouts} currentDate={currentDate} currentDay={days[new Date(`${currentDate}T00:00:00`).getDay()]} />
        <LiveCalloutsPanel callouts={state.callouts} currentDate={currentDate} />
      </section>
      <AttendanceActionPanel state={state} actions={actions} />
    </>
  );
}

function LiveCalloutsPanel({ callouts, currentDate = today }) {
  const todaysCallouts = callouts.filter((callout) => dateOnly(callout.date) === currentDate);
  return (
    <section className="panel live-callouts-panel">
      <div className="section-heading"><div><p className="eyebrow">Live Feed</p><h2>Live Callouts</h2></div><span className="date-pill">{todaysCallouts.length} Today</span></div>
      <div className="live-callout-list">
        {todaysCallouts.length ? todaysCallouts.slice(0, 8).map((callout) => (
          <article key={callout.id}>
            <div><strong>{callout.name}</strong><p>{normalizeShift(callout.shift)} / {callout.company || "No company"} / {callout.reason}</p></div>
            <span>{callout.submittedTime || "Logged"}</span>
          </article>
        )) : <p>No callouts logged today.</p>}
      </div>
    </section>
  );
}

function AttendanceActionPanel({ state, actions }) {
  const activeContractors = state.contractors.filter((c) => c.active);
  const currentTime = new Date().toTimeString().slice(0, 5);
  const [form, setForm] = useState({ contractorId: activeContractors[0]?.id ?? "", date: today, time: currentTime, action: "Scheduled", notes: "" });

  function submit(event) {
    event.preventDefault();
    const contractor = state.contractors.find((c) => c.id === form.contractorId);
    if (!contractor) return;

    const nextContractors = state.contractors.map((item) => {
      if (item.id !== contractor.id) return item;
      return {
        ...item,
        status: form.action,
        statusDate: dailyResetStatuses.includes(form.action) ? form.date : "",
        lateArrivals: form.action === "Late" ? (Number(item.lateArrivals) || 0) + 1 : item.lateArrivals,
        noShows: form.action === "No Show" ? (Number(item.noShows) || 0) + 1 : item.noShows
      };
    });
    const noteText = `${actions.opsShift ?? normalizeShift(contractor.shift)}: Attendance log - ${contractor.name}: ${form.action} on ${form.date} at ${form.time}${form.notes ? ` - ${form.notes}` : ""}`;
    actions.saveContractors(nextContractors);
    actions.saveNotes([{ id: crypto.randomUUID(), date: form.date, text: noteText, createdBy: "Ops Manager" }, ...state.notes]);
    actions.pushNotification(noteText);
    setForm({ contractorId: activeContractors[0]?.id ?? "", date: today, time: currentTime, action: "Scheduled", notes: "" });
  }

  return (
    <section className="panel">
      <div className="section-heading"><div><p className="eyebrow">Attendance Actions</p><h2>Daily Operator Status</h2></div></div>
      <form className="form-grid" onSubmit={submit}>
        <label>Contractor<select value={form.contractorId} onChange={(e) => setForm({ ...form, contractorId: e.target.value })} required>{activeContractors.map((c) => <option value={c.id} key={c.id}>{c.name}</option>)}</select></label>
        <label>Date<input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></label>
        <label>Time<input type="time" value={form.time} onChange={(e) => setForm({ ...form, time: e.target.value })} /></label>
        <label>Status<select value={form.action} onChange={(e) => setForm({ ...form, action: e.target.value })}>{attendanceActions.map((action) => <option key={action}>{action}</option>)}</select></label>
        <div className="quick-actions wide">{attendanceActions.map((action) => <button className={form.action === action ? "active" : ""} type="button" key={action} onClick={() => setForm({ ...form, action })}>{action}</button>)}</div>
        <label className="wide">Notes<textarea value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} placeholder="Reason, time, or manager follow-up" /></label>
        <button className="primary-button">Log Status</button>
      </form>
    </section>
  );
}

function mondayFor(dateText) {
  const date = new Date(`${dateText}T00:00:00`);
  const offset = (date.getDay() + 6) % 7;
  date.setDate(date.getDate() - offset);
  return date.toISOString().slice(0, 10);
}

function WeeklyAttendanceTracker({ state, actions, opsShift }) {
  const [weekStart, setWeekStart] = useState(mondayFor(today));
  const [sortBy, setSortBy] = useState("name");
  const [selectedIds, setSelectedIds] = useState([]);
  const [bulkDay, setBulkDay] = useState("monday");
  const [bulkStatus, setBulkStatus] = useState("P");
  const [operatorToAdd, setOperatorToAdd] = useState("");
  const weekRecords = state.attendanceRecords.filter((record) => record.weekStart === weekStart);
  const sortedWeekRecords = [...weekRecords].sort((a, b) => {
    const primaryKey = sortBy === "agency" ? "company" : "contractorName";
    const secondaryKey = sortBy === "agency" ? "contractorName" : "company";
    const primary = compareText(a[primaryKey], b[primaryKey]);
    if (primary) return primary;
    return compareText(a[secondaryKey], b[secondaryKey]);
  });
  const contractorsWithoutRows = state.contractors.filter((contractor) => contractor.active && !weekRecords.some((record) => record.contractorId === contractor.id));
  const totals = attendanceStatuses.reduce((acc, status) => ({ ...acc, [status]: weekRecords.reduce((sum, record) => sum + [record.monday, record.tuesday, record.wednesday, record.thursday, record.friday].filter((value) => value === status).length, 0) }), {});
  const dayKeys = ["monday", "tuesday", "wednesday", "thursday", "friday"];
  const weekEnd = new Date(`${weekStart}T00:00:00`);
  weekEnd.setDate(weekEnd.getDate() + 4);
  const totalMarks = totals.P + totals.A + totals.T + totals.WO;
  const presentRate = totalMarks ? Math.round((totals.P / totalMarks) * 100) : 0;
  const completionRate = weekRecords.length ? Math.round((totalMarks / (weekRecords.length * dayKeys.length)) * 100) : 0;

  useEffect(() => {
    if (!contractorsWithoutRows.some((contractor) => contractor.id === operatorToAdd)) {
      setOperatorToAdd(contractorsWithoutRows[0]?.id ?? "");
    }
  }, [contractorsWithoutRows, operatorToAdd]);

  function rowForContractor(contractor) {
    return {
      id: crypto.randomUUID(),
      site: contractor.site ?? defaultSite,
      shift: opsShift,
      weekStart,
      contractorId: contractor.id,
      contractorName: contractor.name,
      company: contractor.company,
      monday: "",
      tuesday: "",
      wednesday: "",
      thursday: "",
      friday: "",
      comments: ""
    };
  }

  function addMissingRows() {
    const rows = contractorsWithoutRows.map(rowForContractor);
    if (rows.length) actions.saveAttendanceRecords([...rows, ...state.attendanceRecords]);
  }

  function addSpecificOperator(event) {
    event.preventDefault();
    const contractor = contractorsWithoutRows.find((item) => item.id === operatorToAdd);
    if (!contractor) return;
    actions.saveAttendanceRecords([rowForContractor(contractor), ...state.attendanceRecords]);
  }

  function updateRecord(id, patch) {
    actions.saveAttendanceRecords(state.attendanceRecords.map((record) => record.id === id ? { ...record, ...patch } : record));
  }

  function toggleSelected(id) {
    setSelectedIds((current) => current.includes(id) ? current.filter((item) => item !== id) : [...current, id]);
  }

  function toggleAllVisible() {
    const visibleIds = sortedWeekRecords.map((record) => record.id);
    setSelectedIds((current) => visibleIds.every((id) => current.includes(id)) ? current.filter((id) => !visibleIds.includes(id)) : [...new Set([...current, ...visibleIds])]);
  }

  function applyBulkStatus() {
    if (!selectedIds.length) return;
    actions.saveAttendanceRecords(state.attendanceRecords.map((record) => selectedIds.includes(record.id) ? { ...record, [bulkDay]: bulkStatus } : record));
  }

  async function removeRecord(id) {
    if (!(await confirmDelete("this attendance row"))) return;
    actions.saveAttendanceRecords(state.attendanceRecords.filter((record) => record.id !== id));
  }

  function exportWeek() {
    downloadCsv(`weekly-attendance-${opsShift}-${weekStart}.csv`, ["Week Start", "Name", "Agency", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "P%", "A%", "Tardy %", "Comments"], sortedWeekRecords.map((record) => [record.weekStart, record.contractorName, record.company, record.monday, record.tuesday, record.wednesday, record.thursday, record.friday, `${weeklyAttendancePercent(record, "P")}%`, `${weeklyAttendancePercent(record, "A")}%`, `${weeklyAttendancePercent(record, "T")}%`, record.comments]));
  }

  return (
    <section className="panel attendance-panel">
      <div className="attendance-header">
        <div>
          <p className="eyebrow">Weekly Attendance Tracker</p>
          <h2>{opsShift} Attendance</h2>
          <span>{weekStart} through {weekEnd.toISOString().slice(0, 10)}</span>
        </div>
        <div className="attendance-actions">
          <button className="ghost-button" onClick={exportWeek}>Export CSV</button>
          <button className="primary-button" onClick={addMissingRows}>Add Active Contractors</button>
        </div>
      </div>

      <div className="attendance-toolbar">
        <label>Week Starting<input type="date" value={weekStart} onChange={(e) => setWeekStart(mondayFor(e.target.value))} /></label>
        <label>Sort By<select value={sortBy} onChange={(e) => setSortBy(e.target.value)}><option value="name">Name A-Z</option><option value="agency">Agency A-Z</option></select></label>
        <article><span>Rows</span><strong>{weekRecords.length}</strong></article>
        <article><span>Complete</span><strong>{completionRate}%</strong></article>
        <article><span>Present Rate</span><strong>{presentRate}%</strong></article>
        <article><span>Need Rows</span><strong>{contractorsWithoutRows.length}</strong></article>
      </div>

      <form className="single-operator-add" onSubmit={addSpecificOperator}>
        <label>Add Specific Operator<select value={operatorToAdd} onChange={(event) => setOperatorToAdd(event.target.value)} disabled={!contractorsWithoutRows.length}><option value="">{contractorsWithoutRows.length ? "Select operator" : "All active operators added"}</option>{contractorsWithoutRows.map((contractor) => <option value={contractor.id} key={contractor.id}>{contractor.name} / {contractor.company || "No agency"}</option>)}</select></label>
        <button className="primary-button" type="submit" disabled={!operatorToAdd}>Add Operator Row</button>
      </form>

      <div className="attendance-legend">
        <span className="present"><b>P</b> Present {totals.P}</span>
        <span className="absent"><b>A</b> Absent {totals.A}</span>
        <span className="training"><b>T</b> Tardy {totals.T}</span>
        <span className="week-off"><b>WO</b> Week Off {totals.WO}</span>
      </div>

      <div className="bulk-actions">
        <button className="ghost-button" type="button" onClick={toggleAllVisible}>{selectedIds.length ? "Clear Selection" : "Select All Visible"}</button>
        <label>Day<select value={bulkDay} onChange={(e) => setBulkDay(e.target.value)}>{dayKeys.map((day) => <option value={day} key={day}>{day[0].toUpperCase() + day.slice(1)}</option>)}</select></label>
        <label>Status<select value={bulkStatus} onChange={(e) => setBulkStatus(e.target.value)}><option value="P">Present</option><option value="T">Tardy</option><option value="A">Absent</option><option value="WO">Week Off</option></select></label>
        <button className="primary-button" type="button" onClick={applyBulkStatus}>Apply to {selectedIds.length} Selected</button>
      </div>

      <div className="table-wrap attendance-wrap">
        <table>
          <thead><tr><th>Select</th><th><button className="sort-button" type="button" onClick={() => setSortBy("name")}>Employee Name{sortBy === "name" ? " ↑" : ""}</button></th><th><button className="sort-button" type="button" onClick={() => setSortBy("agency")}>Agency{sortBy === "agency" ? " ↑" : ""}</button></th><th>Monday</th><th>Tuesday</th><th>Wednesday</th><th>Thursday</th><th>Friday</th><th>P%</th><th>A%</th><th>Tardy %</th><th>Comments</th><th>Actions</th></tr></thead>
          <tbody>{sortedWeekRecords.length ? sortedWeekRecords.map((record) => (
            <tr key={record.id}>
              <td><input type="checkbox" checked={selectedIds.includes(record.id)} onChange={() => toggleSelected(record.id)} aria-label={`Select ${record.contractorName}`} /></td>
              <td><strong>{record.contractorName}</strong></td>
              <td>{record.company}</td>
              {dayKeys.map((day) => <td key={day}><div className="attendance-status">{attendanceStatuses.map((status) => <button type="button" data-status={status.toLowerCase()} className={record[day] === status ? "active" : ""} key={status} onClick={() => updateRecord(record.id, { [day]: record[day] === status ? "" : status })}>{status}</button>)}</div></td>)}
              <td className="percent-cell present">{weeklyAttendancePercent(record, "P")}%</td>
              <td className="percent-cell absent">{weeklyAttendancePercent(record, "A")}%</td>
              <td className="percent-cell training">{weeklyAttendancePercent(record, "T")}%</td>
              <td><input value={record.comments} onChange={(e) => updateRecord(record.id, { comments: e.target.value })} placeholder="Notes" /></td>
              <td><div className="row-actions"><button className="danger" onClick={() => removeRecord(record.id)}>Remove</button></div></td>
            </tr>
          )) : <tr><td colSpan="13">No attendance rows for this week. Add active contractors to start tracking.</td></tr>}</tbody>
        </table>
      </div>
    </section>
  );
}

function AvailabilityBoard({ contractors, callouts, currentDate = today, currentDay = todayDay }) {
  const calledOut = new Set(callouts.filter((c) => dateOnly(c.date) === currentDate).map((c) => c.name.toLowerCase()));
  function statusFor(contractor) {
    if (!contractor.active) return "Off Shift";
    if (calledOut.has(contractor.name.toLowerCase())) return "Called Out";
    if (!contractor.scheduledDays.includes(currentDay)) return "Off Shift";
    return contractor.status ?? "Scheduled";
  }
  return (
    <section className="panel">
      <div className="section-heading"><div><p className="eyebrow">Real-Time Staffing</p><h2>Contractor Availability Board</h2></div><span className="date-pill">{currentDay} {currentDate}</span></div>
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

function WeeklyCalendarView({ state, actions, currentDate = today }) {
  const monthStart = currentDate.slice(0, 7);
  const [visibleMonth, setVisibleMonth] = useState(monthStart);
  const [selectedDate, setSelectedDate] = useState(currentDate);
  const activeContractors = state.contractors.filter((c) => c.active);
  const [form, setForm] = useState({ contractorId: activeContractors[0]?.id ?? "", status: "Late", notes: "" });
  const firstOfMonth = new Date(`${visibleMonth}-01T00:00:00`);
  const gridStart = addDays(localDateString(firstOfMonth), -firstOfMonth.getDay());
  const monthDays = Array.from({ length: 42 }, (_, index) => addDays(gridStart, index));

  useEffect(() => {
    if (!activeContractors.some((contractor) => contractor.id === form.contractorId)) {
      setForm((current) => ({ ...current, contractorId: activeContractors[0]?.id ?? "" }));
    }
  }, [activeContractors, form.contractorId]);

  function shiftMonth(amount) {
    const date = new Date(`${visibleMonth}-01T00:00:00`);
    date.setMonth(date.getMonth() + amount);
    const nextMonth = localDateString(date).slice(0, 7);
    setVisibleMonth(nextMonth);
    setSelectedDate(`${nextMonth}-01`);
  }

  function dayKeyFor(date) {
    return ["", "monday", "tuesday", "wednesday", "thursday", "friday", ""][new Date(`${date}T00:00:00`).getDay()];
  }

  function attendanceLabel(status) {
    return { P: "Present", A: "Absent", T: "Tardy", WO: "Time Off" }[status] ?? status;
  }

  function statusToWeeklyCode(status) {
    if (status === "Late") return "T";
    if (status === "No Show" || status === "Absent") return "A";
    if (status === "Day Off Requested" || status === "Time Off") return "WO";
    return "P";
  }

  function itemsFor(date) {
    const dayName = days[new Date(`${date}T00:00:00`).getDay()];
    const scheduled = state.contractors.filter((contractor) => contractor.active && contractor.scheduledDays.includes(dayName));
    const callouts = state.callouts.filter((callout) => dateOnly(callout.date) === date);
    const attendanceWeek = state.attendanceRecords.filter((record) => record.weekStart === mondayFor(date));
    const dayKey = dayKeyFor(date);
    const attendanceMarks = dayKey ? attendanceWeek.filter((record) => record[dayKey]).map((record) => ({ ...record, mark: record[dayKey] })) : [];
    const training = scheduled.filter((contractor) => contractor.status === "Training");
    const dayOff = scheduled.filter((contractor) => contractor.status === "Day Off Requested");
    const notes = state.notes.filter((note) => dateOnly(note.date) === date || note.text?.includes(date)).filter((note) => /Attendance log|Late|No Show|Day Off Requested|Leaving Early|Scheduled/i.test(note.text ?? ""));
    const vehicles = state.vehicles.filter((vehicle) => vehicle.assignedContractorName);
    return { scheduled, callouts, training, dayOff, attendanceMarks, notes, vehicles };
  }

  function selectedSummary() {
    return itemsFor(selectedDate);
  }

  function submitStatus(event) {
    event.preventDefault();
    const contractor = state.contractors.find((item) => item.id === form.contractorId);
    const dayKey = dayKeyFor(selectedDate);
    if (!contractor || !dayKey) return;

    const weekStart = mondayFor(selectedDate);
    const existing = state.attendanceRecords.find((record) => record.weekStart === weekStart && record.contractorId === contractor.id);
    const code = statusToWeeklyCode(form.status);
    const noteText = `${actions.opsShift ?? normalizeShift(contractor.shift)}: Calendar log - ${contractor.name}: ${form.status} on ${selectedDate}${form.notes ? ` - ${form.notes}` : ""}`;
    const nextRecord = existing ? {
      ...existing,
      [dayKey]: code,
      comments: [existing.comments, `${selectedDate}: ${form.status}${form.notes ? ` - ${form.notes}` : ""}`].filter(Boolean).join("; ")
    } : {
      id: crypto.randomUUID(),
      site: contractor.site ?? defaultSite,
      shift: actions.opsShift ?? normalizeShift(contractor.shift),
      weekStart,
      contractorId: contractor.id,
      contractorName: contractor.name,
      company: contractor.company,
      monday: "",
      tuesday: "",
      wednesday: "",
      thursday: "",
      friday: "",
      [dayKey]: code,
      comments: `${selectedDate}: ${form.status}${form.notes ? ` - ${form.notes}` : ""}`
    };
    actions.saveAttendanceRecords(existing ? state.attendanceRecords.map((record) => record.id === existing.id ? nextRecord : record) : [nextRecord, ...state.attendanceRecords]);
    actions.saveNotes([{ id: crypto.randomUUID(), date: selectedDate, text: noteText, createdBy: "Ops Manager" }, ...state.notes]);
    actions.pushNotification(noteText);
    setForm({ contractorId: contractor.id, status: "Late", notes: "" });
  }

  const selectedData = selectedSummary();

  return (
    <section className="panel calendar-panel full-calendar-panel">
      <div className="section-heading calendar-heading">
        <div><p className="eyebrow">Operations Calendar</p><h2>{firstOfMonth.toLocaleString("default", { month: "long", year: "numeric" })}</h2></div>
        <div className="button-row"><button className="ghost-button" type="button" onClick={() => shiftMonth(-1)}>Previous</button><button className="ghost-button" type="button" onClick={() => { setVisibleMonth(monthStart); setSelectedDate(currentDate); }}>Today</button><button className="ghost-button" type="button" onClick={() => shiftMonth(1)}>Next</button></div>
      </div>
      <div className="calendar-layout">
        <div className="calendar-month">
          {days.map((day) => <div className="calendar-weekday" key={day}>{day}</div>)}
          {monthDays.map((date) => {
          const data = itemsFor(date);
          const isCurrentMonth = date.slice(0, 7) === visibleMonth;
          const isSelected = date === selectedDate;
          return (
            <button className={`calendar-cell ${isCurrentMonth ? "" : "muted"} ${isSelected ? "selected" : ""}`} type="button" key={date} onClick={() => setSelectedDate(date)}>
              <span className="calendar-date">{new Date(`${date}T00:00:00`).getDate()}</span>
              <small>{data.scheduled.length} scheduled</small>
              {data.callouts.slice(0, 2).map((callout) => <b className="calendar-event callout" key={callout.id}>{callout.name} callout</b>)}
              {data.attendanceMarks.slice(0, 3).map((record) => <b className={`calendar-event ${record.mark.toLowerCase()}`} key={`${record.id}-${record.mark}`}>{record.contractorName}: {attendanceLabel(record.mark)}</b>)}
              {[...data.training.map((c) => `${c.name}: Training`), ...data.dayOff.map((c) => `${c.name}: Time Off`)].slice(0, 2).map((item) => <b className="calendar-event info" key={item}>{item}</b>)}
              {data.notes.slice(0, 1).map((note) => <b className="calendar-event note" key={note.id}>Log: {note.text.split(" - ")[0]}</b>)}
            </button>
          );
        })}
        </div>
        <aside className="calendar-detail">
          <div><p className="eyebrow">Selected Date</p><h2>{selectedDate}</h2></div>
          <form className="calendar-status-form" onSubmit={submitStatus}>
            <label>Contractor<select value={form.contractorId} onChange={(e) => setForm({ ...form, contractorId: e.target.value })} required>{activeContractors.map((contractor) => <option value={contractor.id} key={contractor.id}>{contractor.name}</option>)}</select></label>
            <label>Status<select value={form.status} onChange={(e) => setForm({ ...form, status: e.target.value })}><option>Late</option><option>Absent</option><option>No Show</option><option>Leaving Early</option><option>Time Off</option><option>Day Off Requested</option><option>Scheduled</option><option>Present</option></select></label>
            <label>Notes<textarea value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} placeholder="Reason, approved time off, or manager follow-up" /></label>
            <button className="primary-button">Log on Calendar</button>
          </form>
          <div className="calendar-detail-list"><strong>Scheduled</strong>{selectedData.scheduled.length ? selectedData.scheduled.slice(0, 8).map((contractor) => <p key={contractor.id}>{contractor.name} / {contractor.company}</p>) : <p>None</p>}</div>
          <div className="calendar-detail-list"><strong>Callouts</strong>{selectedData.callouts.length ? selectedData.callouts.map((callout) => <p key={callout.id}>{callout.name} / {callout.reason}</p>) : <p>None</p>}</div>
          <div className="calendar-detail-list"><strong>Attendance Logs</strong>{selectedData.attendanceMarks.length ? selectedData.attendanceMarks.map((record) => <p key={`${record.id}-detail`}>{record.contractorName}: {attendanceLabel(record.mark)}</p>) : <p>None</p>}</div>
          <div className="calendar-detail-list"><strong>Vehicles</strong>{selectedData.vehicles.length ? selectedData.vehicles.slice(0, 6).map((vehicle) => <p key={vehicle.id}>{vehicle.name} - {vehicle.assignedContractorName}</p>) : <p>No pairings</p>}</div>
        </aside>
      </div>
    </section>
  );
}

function GlobalSearch({ state }) {
  const [query, setQuery] = useState("");
  const search = query.trim().toLowerCase();
  const results = useMemo(() => {
    const rows = [
      ...state.contractors.map((item) => ({ type: "Contractor", title: item.name, detail: `${item.company} / ${normalizeShift(item.shift)} / ${item.status}` })),
      ...state.vehicles.map((item) => ({ type: "Vehicle", title: item.name, detail: `${item.vin} / ${item.licensePlate} / ${item.assignedContractorName || "Unpaired"}` })),
      ...state.laptops.map((item) => ({ type: "Laptop", title: item.asset, detail: `${item.contractor || "Unassigned"} / ${item.status}` })),
      ...state.callouts.map((item) => ({ type: "Callout", title: item.name, detail: `${item.date} / ${item.reason} / ${item.status}` })),
      ...state.tickets.map((item) => ({ type: "Coverage", title: item.contractorName, detail: `${item.date} / ${item.status} / ${item.replacementAssigned || "Unassigned"}` })),
      ...state.notes.map((item) => ({ type: "Note", title: item.date, detail: item.text })),
      ...state.attendanceRecords.map((item) => ({ type: "Attendance", title: item.contractorName, detail: `${item.weekStart} / ${item.company} / P:${weeklyAttendancePercent(item, "P")}% A:${weeklyAttendancePercent(item, "A")}% Tardy:${weeklyAttendancePercent(item, "T")}%` })),
      ...reportTypes.map((item) => ({ type: "Report", title: item, detail: "Available in Reports Generator" }))
    ];
    if (!search) return rows.slice(0, 20);
    return rows.filter((row) => [row.type, row.title, row.detail].some((value) => String(value).toLowerCase().includes(search))).slice(0, 80);
  }, [state, search]);

  return (
    <section className="panel search-panel">
      <div className="section-heading"><div><p className="eyebrow">Global Search</p><h2>Search Everything</h2></div><span className="date-pill">{results.length} Results</span></div>
      <label>Search<input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Contractor, vehicle, laptop, callout, note, report..." autoFocus /></label>
      <div className="search-results">
        {results.map((result, index) => <article key={`${result.type}-${result.title}-${index}`}><span>{result.type}</span><strong>{result.title}</strong><p>{result.detail}</p></article>)}
      </div>
    </section>
  );
}

function CalloutCenter({ state, actions }) {
  const [filters, setFilters] = useState({ date: "", company: "", shift: "", reason: "" });
  const [sortBy, setSortBy] = useState("date");
  const companies = [...new Set(state.contractors.map((c) => c.company).filter(Boolean))];
  const filtered = state.callouts.filter((c) =>
    (!filters.date || c.date === filters.date) &&
    (!filters.company || c.company === filters.company) &&
    (!filters.shift || normalizeShift(c.shift) === filters.shift) &&
    (!filters.reason || c.reason === filters.reason)
  );
  const sortedCallouts = [...filtered].sort((a, b) => {
    const primaryKey = sortBy === "company" ? "company" : sortBy === "name" ? "name" : "date";
    const secondaryKey = primaryKey === "name" ? "company" : "name";
    const primary = primaryKey === "date" ? compareDate(b.date, a.date) : compareText(a[primaryKey], b[primaryKey]);
    if (primary) return primary;
    return compareText(a[secondaryKey], b[secondaryKey]) || compareDate(b.date, a.date);
  });
  function exportFiltered() {
    downloadCsv(`callouts-${today}.csv`, ["Date", "Time", "Name", "Company", "Shift", "Reason", "Notes", "Hours Before Shift", "Coverage Status", "Replacement", "Status"], sortedCallouts.map((c) => [c.date, c.submittedTime, c.name, c.company, normalizeShift(c.shift), c.reason, c.notes || "", c.hoursBeforeShiftStart, c.coverageStatus, c.replacementAssigned, c.status]));
  }
  async function removeCallout(id) {
    if (!(await confirmDelete("this callout"))) return;
    actions.saveCallouts(state.callouts.filter((callout) => callout.id !== id));
  }
  return (
    <section className="panel">
      <div className="section-heading"><div><p className="eyebrow">Callout Management</p><h2>Callout Center</h2></div><button className="primary-button" onClick={exportFiltered}>Export CSV</button></div>
      <ManualCalloutEntry contractors={state.contractors} opsShift={actions.opsShift} onSubmit={actions.submitCallout} />
      <div className="filters-grid">
        <label>Date<input type="date" value={filters.date} onChange={(e) => setFilters({ ...filters, date: e.target.value })} /></label>
        <label>Company<select value={filters.company} onChange={(e) => setFilters({ ...filters, company: e.target.value })}><option value="">All</option>{companies.map((c) => <option key={c}>{c}</option>)}</select></label>
        <label>Shift<select value={filters.shift} onChange={(e) => setFilters({ ...filters, shift: e.target.value })}><option value="">All</option>{shifts.map((s) => <option key={s}>{s}</option>)}</select></label>
        <label>Reason<select value={filters.reason} onChange={(e) => setFilters({ ...filters, reason: e.target.value })}><option value="">All</option>{reasons.map((r) => <option key={r}>{r}</option>)}</select></label>
      </div>
      <CalloutTable callouts={sortedCallouts} sortBy={sortBy} setSortBy={setSortBy} onRemove={removeCallout} />
    </section>
  );
}

function ManualCalloutEntry({ contractors, opsShift = "AM Shift", onSubmit }) {
  const defaultForm = { name: "", date: today, shift: normalizeShift(opsShift), reason: "Sick", notes: "" };
  const [form, setForm] = useState(defaultForm);
  const [saved, setSaved] = useState(false);
  const contractorNames = [...contractors].sort((a, b) => compareText(a.name, b.name)).map((contractor) => contractor.name);
  const selectedContractor = contractors.find((contractor) => contractor.name.toLowerCase() === form.name.toLowerCase());

  useEffect(() => {
    setForm((current) => ({ ...current, shift: normalizeShift(opsShift) }));
  }, [opsShift]);

  function update(field, value) {
    setSaved(false);
    setForm((current) => ({ ...current, [field]: value }));
  }

  function submit(event) {
    event.preventDefault();
    onSubmit({
      name: form.name.trim(),
      date: form.date,
      shift: normalizeShift(form.shift),
      reason: form.reason,
      notes: form.notes
    });
    setForm({ ...defaultForm, shift: normalizeShift(opsShift) });
    setSaved(true);
  }

  return (
    <form className="manual-callout-panel" onSubmit={submit}>
      <div>
        <p className="eyebrow">Admin Entry</p>
        <h3>Manual Callout Entry</h3>
      </div>
      <div className="form-grid manual-callout-grid">
        <label>Contractor Name<input list="manual-callout-contractors" value={form.name} onChange={(event) => update("name", event.target.value)} placeholder="Start typing a contractor name" required /></label>
        <datalist id="manual-callout-contractors">{contractorNames.map((name) => <option value={name} key={name} />)}</datalist>
        <label>Date<input type="date" value={form.date} onChange={(event) => update("date", event.target.value)} required /></label>
        <label>Shift<select value={form.shift} onChange={(event) => update("shift", event.target.value)}>{shifts.map((shift) => <option key={shift}>{shift}</option>)}</select></label>
        <label>Reason<select value={form.reason} onChange={(event) => update("reason", event.target.value)}>{reasons.map((reason) => <option key={reason}>{reason}</option>)}</select></label>
        <label className="wide">Notes<textarea value={form.notes} onChange={(event) => update("notes", event.target.value)} placeholder="Optional manager note" /></label>
      </div>
      <div className="manual-callout-footer">
        <span>{selectedContractor ? `${selectedContractor.company || "No company"} / ${normalizeShift(selectedContractor.shift)}` : "Coverage ticket will be created automatically."}</span>
        <div className="row-actions">
          {saved && <span className="success-pill">Callout added</span>}
          <button className="primary-button" type="submit">Add Callout</button>
        </div>
      </div>
    </form>
  );
}

function CalloutTable({ callouts, sortBy, setSortBy, onRemove }) {
  return (
    <div className="table-wrap">
      <table>
        <thead><tr><th><button className="sort-button" type="button" onClick={() => setSortBy("date")}>Date{sortBy === "date" ? " ↑" : ""}</button></th><th>Time</th><th><button className="sort-button" type="button" onClick={() => setSortBy("name")}>Name{sortBy === "name" ? " ↑" : ""}</button></th><th><button className="sort-button" type="button" onClick={() => setSortBy("company")}>Company{sortBy === "company" ? " ↑" : ""}</button></th><th>Shift</th><th>Reason</th><th>Notes</th><th>Hours Before</th><th>Coverage</th><th>Replacement</th><th>Status</th><th>Actions</th></tr></thead>
        <tbody>{callouts.length ? callouts.map((c) => <tr key={c.id}><td>{c.date}</td><td>{c.submittedTime}</td><td>{c.name}</td><td>{c.company}</td><td>{normalizeShift(c.shift)}</td><td>{c.reason}</td><td className="notes-cell" title={c.notes || ""}>{c.notes || "-"}</td><td>{c.hoursBeforeShiftStart}</td><td>{c.coverageStatus}</td><td>{c.replacementAssigned || "Unassigned"}</td><td><span className="status called-out">{c.status}</span></td><td><div className="row-actions"><button className="danger" onClick={() => onRemove(c.id)}>Remove</button></div></td></tr>) : <tr><td colSpan="12">No callouts match filters.</td></tr>}</tbody>
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
    actions.saveTickets(next);
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
  const blank = { asset: "", contractorId: "", contractor: "", status: "Assigned", assignedBy: "Ops Admin", assignedDate: today, returnedBy: "", returnedDate: "", dueDate: today, history: "" };
  const [editingId, setEditingId] = useState(null);
  const [showForm, setShowForm] = useState(false);
  const [draft, setDraft] = useState(blank);
  const assigned = state.laptops.filter((l) => l.status === "Assigned" || l.status === "Overdue").length;
  const returned = state.laptops.filter((l) => l.status === "Returned").length;
  const overdue = state.laptops.filter((l) => l.status === "Overdue").length;

  function contractorName(laptop) {
    return laptop.contractor || state.contractors.find((c) => c.id === laptop.contractorId)?.name || "Unassigned";
  }

  function updateLaptop(id, patch) {
    actions.saveLaptops(state.laptops.map((l) => l.id === id ? { ...l, ...patch, history: `${l.history}; ${patch.status ?? "Updated"} ${new Date().toLocaleTimeString()}` } : l));
    if (patch.status === "Overdue") actions.pushNotification("Laptop became overdue.");
  }

  function updateDraft(patch) {
    const next = { ...draft, ...patch };
    if (patch.contractorId !== undefined) {
      const contractor = state.contractors.find((c) => c.id === patch.contractorId);
      next.contractor = contractor?.name ?? "";
    }
    if (patch.status === "Returned" && !draft.returnedDate) next.returnedDate = today;
    setDraft(next);
  }

  function startEdit(laptop) {
    setEditingId(laptop.id);
    setDraft({ ...blank, ...laptop, contractor: contractorName(laptop) === "Unassigned" ? "" : contractorName(laptop) });
    setShowForm(true);
  }

  function resetForm() {
    setEditingId(null);
    setDraft(blank);
    setShowForm(false);
  }

  function saveDraft(event) {
    event.preventDefault();
    if (!draft.asset.trim()) return;
    const historyEntry = `${editingId ? "Updated" : "Added"} ${draft.asset} ${new Date().toLocaleString()}`;
    const laptop = {
      ...draft,
      id: editingId ?? crypto.randomUUID(),
      history: draft.history ? `${draft.history}; ${historyEntry}` : historyEntry
    };
    const next = editingId ? state.laptops.map((item) => item.id === editingId ? laptop : item) : [laptop, ...state.laptops];
    actions.saveLaptops(next);
    if (laptop.contractor) actions.pushNotification(`Laptop ${laptop.asset} assigned to ${laptop.contractor}.`);
    resetForm();
  }

  async function removeLaptop(id) {
    if (!(await confirmDelete("this laptop"))) return;
    actions.saveLaptops(state.laptops.filter((laptop) => laptop.id !== id));
  }

  function statusClass(status) {
    return String(status ?? "Assigned").toLowerCase().replaceAll(" ", "-");
  }

  return (
    <>
      <section className="metric-grid laptop-grid">
        <Metric label="Assigned" value={assigned} /><Metric label="Returned" value={returned} /><Metric label="Overdue" value={overdue} />
      </section>
      <section className="panel">
          <div className="section-heading"><div><p className="eyebrow">Laptop Accountability</p><h2>Inventory & Assignments</h2></div><button className="primary-button" onClick={() => setShowForm(true)}>Add Laptop</button></div>
          <div className="laptop-list">
            {state.laptops.length ? state.laptops.map((laptop) => (
              <article className="laptop-card" key={laptop.id}>
                <div className="laptop-card-main">
                  <div>
                    <strong>{laptop.asset}</strong>
                    <span>{contractorName(laptop)}</span>
                  </div>
                  <span className={`status-pill ${statusClass(laptop.status)}`}>{laptop.status}</span>
                </div>
                <div className="laptop-status-control" aria-label={`Status for ${laptop.asset}`}>
                  {laptopStatuses.map((status) => <button className={laptop.status === status ? "active" : ""} type="button" key={status} onClick={() => updateLaptop(laptop.id, { status })}>{status}</button>)}
                </div>
                <div className="laptop-meta">
                  <span><b>Assigned</b>{laptop.assignedBy || "-"} / {laptop.assignedDate || "-"}</span>
                  <span><b>Returned</b>{laptop.returnedBy || "-"} / {laptop.returnedDate || "-"}</span>
                  <span><b>Audit</b>{laptop.history || "-"}</span>
                </div>
                <div className="row-actions">
                  <button onClick={() => startEdit(laptop)}>Edit</button>
                  <button className="danger" onClick={() => removeLaptop(laptop.id)}>Remove</button>
                </div>
              </article>
            )) : <p>No laptops added yet.</p>}
          </div>
          {showForm && (
            <div className="modal-backdrop" role="dialog" aria-modal="true" aria-label="Laptop form">
              <div className="panel modal-panel">
                <div className="section-heading"><div><p className="eyebrow">{editingId ? "Edit" : "Add"}</p><h2>Laptop</h2></div><button className="ghost-button" type="button" onClick={resetForm}>Close</button></div>
                <form className="form-grid" onSubmit={saveDraft}>
                  <label>Asset Tag<input value={draft.asset} onChange={(e) => updateDraft({ asset: e.target.value })} placeholder="RIV-LT-104" required /></label>
                  <label>Assign To<select value={draft.contractorId} onChange={(e) => updateDraft({ contractorId: e.target.value })}><option value="">Unassigned</option>{state.contractors.filter((c) => c.active).map((c) => <option value={c.id} key={c.id}>{c.name}</option>)}</select></label>
                  <label>Status<select value={draft.status} onChange={(e) => updateDraft({ status: e.target.value })}>{laptopStatuses.map((status) => <option key={status}>{status}</option>)}</select></label>
                  <div className="laptop-status-control wide">{laptopStatuses.map((status) => <button className={draft.status === status ? "active" : ""} type="button" key={status} onClick={() => updateDraft({ status })}>{status}</button>)}</div>
                  <label>Assigned By<input value={draft.assignedBy} onChange={(e) => updateDraft({ assignedBy: e.target.value })} /></label>
                  <label>Assigned Date<input type="date" value={draft.assignedDate} onChange={(e) => updateDraft({ assignedDate: e.target.value })} /></label>
                  <label>Returned By<input value={draft.returnedBy} onChange={(e) => updateDraft({ returnedBy: e.target.value })} /></label>
                  <label>Returned Date<input type="date" value={draft.returnedDate} onChange={(e) => updateDraft({ returnedDate: e.target.value })} /></label>
                  <label className="wide">Audit Notes<textarea value={draft.history} onChange={(e) => updateDraft({ history: e.target.value })} placeholder="Assignment notes, condition, return details" /></label>
                  <div className="button-row wide"><button className="primary-button">{editingId ? "Save" : "Add Laptop"}</button><button className="ghost-button" type="button" onClick={resetForm}>Cancel</button></div>
                </form>
              </div>
            </div>
          )}
      </section>
    </>
  );
}

function VehicleModule({ vehicles, contractors, saveVehicles }) {
  const blank = { name: "", vin: "", status: "Available", licensePlate: "", model: "R1T", year: String(new Date().getFullYear()), location: "", assignedContractorId: "", assignedContractorName: "" };
  const [editingId, setEditingId] = useState(null);
  const [showForm, setShowForm] = useState(false);
  const [draft, setDraft] = useState(blank);
  const [importMessage, setImportMessage] = useState("");
  const [filters, setFilters] = useState({ search: "", status: "", model: "", location: "" });
  const [sort, setSort] = useState({ key: "name", direction: "asc" });
  const activeContractors = contractors.filter((contractor) => contractor.active);
  const pairedCount = vehicles.filter((vehicle) => vehicle.assignedContractorId || vehicle.assignedContractorName).length;
  const locationOptions = useMemo(() => [...new Set(vehicles.map((vehicle) => vehicle.location).filter(Boolean))].sort(), [vehicles]);
  const filteredVehicles = useMemo(() => {
    const search = filters.search.trim().toLowerCase();
    const matchesSearch = (vehicle) => !search || [vehicle.name, vehicle.vin, vehicle.licensePlate, vehicle.location, vehicle.status, vehicle.model, vehicle.year, vehicle.assignedContractorName].some((value) => String(value ?? "").toLowerCase().includes(search));
    const matchesFilter = (vehicle) => (!filters.status || vehicle.status === filters.status) && (!filters.model || vehicle.model === filters.model) && (!filters.location || vehicle.location === filters.location);
    return vehicles
      .filter((vehicle) => matchesSearch(vehicle) && matchesFilter(vehicle))
      .sort((a, b) => {
        const result = sort.key === "year" ? Number(a[sort.key] || 0) - Number(b[sort.key] || 0) : compareText(a[sort.key], b[sort.key]);
        return sort.direction === "asc" ? result : -result;
      });
  }, [vehicles, filters, sort]);

  function updateFilter(key, value) {
    setFilters((current) => ({ ...current, [key]: value }));
  }

  function toggleSort(key) {
    setSort((current) => ({ key, direction: current.key === key && current.direction === "asc" ? "desc" : "asc" }));
  }

  function sortLabel(key) {
    if (sort.key !== key) return "";
    return sort.direction === "asc" ? " ↑" : " ↓";
  }

  function startEdit(vehicle) {
    setEditingId(vehicle.id);
    setDraft({ ...blank, ...vehicle });
    setShowForm(true);
  }

  function resetForm() {
    setEditingId(null);
    setDraft(blank);
    setShowForm(false);
  }

  function saveDraft(event) {
    event.preventDefault();
    if (!draft.name.trim() || !draft.vin.trim()) return;
    const vehicle = { ...draft, id: editingId ?? crypto.randomUUID() };
    const next = editingId ? vehicles.map((item) => item.id === editingId ? vehicle : item) : [vehicle, ...vehicles];
    saveVehicles(next);
    resetForm();
  }

  async function removeVehicle(id) {
    if (!(await confirmDelete("this vehicle"))) return;
    saveVehicles(vehicles.filter((vehicle) => vehicle.id !== id));
  }

  function assignVehicle(vehicleId, contractorId) {
    const contractor = activeContractors.find((item) => item.id === contractorId);
    saveVehicles(vehicles.map((vehicle) => vehicle.id === vehicleId ? { ...vehicle, assignedContractorId: contractor?.id ?? "", assignedContractorName: contractor?.name ?? "", status: contractor ? "Assigned" : vehicle.status } : vehicle));
  }

  async function importVehicleFile(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    const isSpreadsheet = /\.(xlsx|xls)$/i.test(file.name);
    let rows;
    if (isSpreadsheet) {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
      rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1, defval: "" });
    } else {
      rows = parseCsv(await file.text());
    }
    const [headers = [], ...dataRows] = rows;
    const headerMap = Object.fromEntries(headers.map((header, index) => [normalizeHeader(header), index]));
    const value = (row, names) => {
      const index = names.map(normalizeHeader).map((name) => headerMap[name]).find((item) => item !== undefined);
      return index === undefined ? "" : String(row[index] ?? "").trim();
    };
    const imported = dataRows.map((row) => ({
      id: crypto.randomUUID(),
      name: value(row, ["name", "vehicle", "vehicle name"]),
      vin: value(row, ["vin"]),
      status: value(row, ["status"]) || "Available",
      licensePlate: value(row, ["license_plate", "license plate", "plate"]),
      model: vehicleModels.includes(value(row, ["model"]).toUpperCase()) ? value(row, ["model"]).toUpperCase() : "R1T",
      year: value(row, ["year"]),
      location: value(row, ["location"]),
      assignedContractorId: "",
      assignedContractorName: value(row, ["assigned_contractor", "assigned contractor", "contractor"])
    })).filter((vehicle) => vehicle.name && vehicle.vin);

    saveVehicles([...imported, ...vehicles]);
    setImportMessage(`${imported.length} vehicles imported`);
    event.target.value = "";
  }

  return (
    <section className="panel">
        <div className="section-heading"><div><p className="eyebrow">Fleet</p><h2>Vehicle Tracking</h2></div><div className="button-row"><span className="date-pill">{vehicles.length} Vehicles</span><button className="primary-button" onClick={() => setShowForm(true)}>Add Vehicle</button></div></div>
        <section className="pairing-summary">
          <article><span>Paired Vehicles</span><strong>{pairedCount}</strong></article>
          <article><span>Unpaired Vehicles</span><strong>{Math.max(0, vehicles.length - pairedCount)}</strong></article>
          <article><span>Active Contractors</span><strong>{activeContractors.length}</strong></article>
        </section>
        <section className="pairing-panel">
          <div className="section-heading"><div><p className="eyebrow">Pairing</p><h2>Contractor Vehicle Pairing</h2></div></div>
          <div className="pairing-list">
            {vehicles.length ? vehicles.map((vehicle) => (
              <article className="pairing-card" key={vehicle.id}>
                <div>
                  <strong>{vehicle.name}</strong>
                  <span>{vehicle.model} / {vehicle.licensePlate || "No plate"} / {vehicle.location || "No location"}</span>
                </div>
                <label>Assigned Contractor<select value={vehicle.assignedContractorId || ""} onChange={(e) => assignVehicle(vehicle.id, e.target.value)}><option value="">Unpaired</option>{activeContractors.map((contractor) => <option value={contractor.id} key={contractor.id}>{contractor.name}</option>)}</select></label>
                <button className="ghost-button" type="button" onClick={() => assignVehicle(vehicle.id, "")}>Unpair</button>
              </article>
            )) : <p>No vehicles available to pair.</p>}
          </div>
        </section>
        <div className="vehicle-toolbar">
          <label className="search-field">Search Vehicles<input value={filters.search} onChange={(e) => updateFilter("search", e.target.value)} placeholder="Name, VIN, plate, location" /></label>
          <label>Status<select value={filters.status} onChange={(e) => updateFilter("status", e.target.value)}><option value="">All Statuses</option>{vehicleStatuses.map((status) => <option key={status}>{status}</option>)}</select></label>
          <label>Model<select value={filters.model} onChange={(e) => updateFilter("model", e.target.value)}><option value="">All Models</option>{vehicleModels.map((model) => <option key={model}>{model}</option>)}</select></label>
          <label>Location<select value={filters.location} onChange={(e) => updateFilter("location", e.target.value)}><option value="">All Locations</option>{locationOptions.map((location) => <option key={location}>{location}</option>)}</select></label>
          <button className="ghost-button" type="button" onClick={() => setFilters({ search: "", status: "", model: "", location: "" })}>Clear Filters</button>
        </div>
        <div className="import-row">
          <label className="ghost-button">Import CSV/XLSX<input type="file" accept=".csv,text/csv,.xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel" onChange={importVehicleFile} /></label>
          <span>{filteredVehicles.length} shown. Headers: name, vin, status, license_plate, model, year, location, assigned_contractor</span>
          {importMessage && <strong>{importMessage}</strong>}
        </div>
        <div className="table-wrap"><table><thead><tr>
          <th><button className="sort-button" type="button" onClick={() => toggleSort("name")}>Name{sortLabel("name")}</button></th>
          <th><button className="sort-button" type="button" onClick={() => toggleSort("vin")}>VIN{sortLabel("vin")}</button></th>
          <th><button className="sort-button" type="button" onClick={() => toggleSort("status")}>Status{sortLabel("status")}</button></th>
          <th><button className="sort-button" type="button" onClick={() => toggleSort("licensePlate")}>Plate{sortLabel("licensePlate")}</button></th>
          <th><button className="sort-button" type="button" onClick={() => toggleSort("model")}>Model{sortLabel("model")}</button></th>
          <th><button className="sort-button" type="button" onClick={() => toggleSort("year")}>Year{sortLabel("year")}</button></th>
          <th><button className="sort-button" type="button" onClick={() => toggleSort("location")}>Location{sortLabel("location")}</button></th>
          <th><button className="sort-button" type="button" onClick={() => toggleSort("assignedContractorName")}>Paired To{sortLabel("assignedContractorName")}</button></th>
          <th>Actions</th>
        </tr></thead><tbody>
          {filteredVehicles.length ? filteredVehicles.map((vehicle) => <tr key={vehicle.id}><td><strong>{vehicle.name}</strong></td><td>{vehicle.vin}</td><td><span className={`status ${(vehicle.status ?? "available").toLowerCase().replaceAll(" ", "-")}`}>{vehicle.status}</span></td><td>{vehicle.licensePlate}</td><td>{vehicle.model}</td><td>{vehicle.year}</td><td>{vehicle.location}</td><td>{vehicle.assignedContractorName || "Unpaired"}</td><td><div className="row-actions"><button onClick={() => startEdit(vehicle)}>Edit</button><button className="danger" onClick={() => removeVehicle(vehicle.id)}>Remove</button></div></td></tr>) : <tr><td colSpan="9">{vehicles.length ? "No vehicles match those filters." : "No vehicles added yet."}</td></tr>}
        </tbody></table></div>
        {showForm && (
          <div className="modal-backdrop" role="dialog" aria-modal="true" aria-label="Vehicle form">
            <div className="panel modal-panel">
              <div className="section-heading"><div><p className="eyebrow">{editingId ? "Edit" : "Add"}</p><h2>Vehicle</h2></div><button className="ghost-button" type="button" onClick={resetForm}>Close</button></div>
              <form className="form-grid" onSubmit={saveDraft}>
                <label>Name<input value={draft.name} onChange={(e) => setDraft({ ...draft, name: e.target.value })} placeholder="Fleet R1T 03" required /></label>
                <label>VIN<input value={draft.vin} onChange={(e) => setDraft({ ...draft, vin: e.target.value.toUpperCase() })} required /></label>
                <label>Status<select value={draft.status} onChange={(e) => setDraft({ ...draft, status: e.target.value })}>{vehicleStatuses.map((status) => <option key={status}>{status}</option>)}</select></label>
                <label>License Plate<input value={draft.licensePlate} onChange={(e) => setDraft({ ...draft, licensePlate: e.target.value.toUpperCase() })} /></label>
                <label>Model<select value={draft.model} onChange={(e) => setDraft({ ...draft, model: e.target.value })}>{vehicleModels.map((model) => <option key={model}>{model}</option>)}</select></label>
                <label>Year<input type="number" min="2020" max="2035" value={draft.year} onChange={(e) => setDraft({ ...draft, year: e.target.value })} /></label>
                <label>Assigned Contractor<select value={draft.assignedContractorId} onChange={(e) => { const contractor = activeContractors.find((item) => item.id === e.target.value); setDraft({ ...draft, assignedContractorId: contractor?.id ?? "", assignedContractorName: contractor?.name ?? "" }); }}><option value="">Unpaired</option>{activeContractors.map((contractor) => <option value={contractor.id} key={contractor.id}>{contractor.name}</option>)}</select></label>
                <label className="wide">Location<input value={draft.location} onChange={(e) => setDraft({ ...draft, location: e.target.value })} placeholder="Main Lot, Bay 3, Service Center" /></label>
                <div className="button-row wide"><button className="primary-button">{editingId ? "Save" : "Add Vehicle"}</button><button className="ghost-button" type="button" onClick={resetForm}>Cancel</button></div>
              </form>
            </div>
          </div>
        )}
    </section>
  );
}

function PerformanceScorecard({ contractors, callouts, attendanceRecords = [] }) {
  const [sortBy, setSortBy] = useState("name");
  const [filters, setFilters] = useState({
    name: "",
    company: "",
    startFrom: "",
    startTo: "",
    daysMin: "",
    attendanceMin: "",
    attendanceMax: "",
    calloutsMin: "",
    noShowsMax: "",
    lateMax: "",
    consecutiveMin: "",
    trainingMin: "",
    badge: ""
  });
  function daysActive(contractor) {
    if (!contractor.startDate) return 0;
    return Math.max(0, Math.floor((Date.now() - new Date(`${contractor.startDate}T00:00:00`).getTime()) / 86400000));
  }
  function badge(c) {
    const stats = attendanceStatsForContractor(c, callouts, attendanceRecords);
    const attendance = stats.attendance;
    if (attendance >= 96 && stats.absences === 0) return "Excellent";
    if (attendance >= 90) return "Good";
    if (stats.absences > 1 || attendance < 85) return "Needs Attention";
    return "Watchlist";
  }
  function updateFilter(key, value) {
    setFilters((current) => ({ ...current, [key]: value }));
  }
  function passesMin(value, filter) {
    return filter === "" || Number(value) >= Number(filter);
  }
  function passesMax(value, filter) {
    return filter === "" || Number(value) <= Number(filter);
  }
  const companies = [...new Set(contractors.map((c) => c.company).filter(Boolean))].sort(compareText);
  const scoreRows = contractors.map((contractor) => {
    const stats = attendanceStatsForContractor(contractor, callouts, attendanceRecords);
    return {
      contractor,
      daysActive: daysActive(contractor),
      attendance: stats.attendance,
      trackedDays: stats.trackedDays,
      callouts: calloutsForContractor(contractor, callouts).length,
      noShows: stats.absences,
      late: stats.tardy,
      consecutive: contractor.attendanceDays ?? 0,
      training: contractor.trainingCompletion ?? 100,
      badge: badge(contractor)
    };
  });
  const visibleRows = scoreRows.filter(({ contractor, ...row }) =>
    (!filters.name || contractor.name.toLowerCase().includes(filters.name.toLowerCase())) &&
    (!filters.company || contractor.company === filters.company) &&
    (!filters.startFrom || contractor.startDate >= filters.startFrom) &&
    (!filters.startTo || contractor.startDate <= filters.startTo) &&
    passesMin(row.daysActive, filters.daysMin) &&
    passesMin(row.attendance, filters.attendanceMin) &&
    passesMax(row.attendance, filters.attendanceMax) &&
    passesMin(row.callouts, filters.calloutsMin) &&
    passesMax(row.noShows, filters.noShowsMax) &&
    passesMax(row.late, filters.lateMax) &&
    passesMin(row.consecutive, filters.consecutiveMin) &&
    passesMin(row.training, filters.trainingMin) &&
    (!filters.badge || row.badge === filters.badge)
  ).sort((a, b) => {
    const valueFor = (row) => {
      if (sortBy === "company") return row.contractor.company ?? "";
      if (sortBy === "startDate") return row.contractor.startDate ?? "";
      if (sortBy === "daysActive") return row.daysActive;
      if (sortBy === "attendance") return row.attendance;
      if (sortBy === "callouts") return row.callouts;
      if (sortBy === "noShows") return row.noShows;
      if (sortBy === "late") return row.late;
      if (sortBy === "consecutive") return row.consecutive;
      if (sortBy === "training") return row.training;
      if (sortBy === "badge") return row.badge;
      return row.contractor.name ?? "";
    };
    const aValue = valueFor(a);
    const bValue = valueFor(b);
    if (typeof aValue === "number" && typeof bValue === "number") return bValue - aValue;
    const primary = sortBy === "startDate" ? compareDate(aValue, bValue) : compareText(aValue, bValue);
    if (primary) return primary;
    return compareText(a.contractor.name, b.contractor.name);
  });
  const filterCount = Object.values(filters).filter(Boolean).length;
  const sortHeader = (key, label) => <button className="sort-button" type="button" onClick={() => setSortBy(key)}>{label}{sortBy === key ? " ↑" : ""}</button>;
  return (
    <section className="panel">
      <div className="section-heading"><div><p className="eyebrow">Contractor Performance</p><h2>Scorecard</h2></div><div className="button-row"><span className="date-pill">{visibleRows.length} Showing</span><button className="ghost-button" type="button" onClick={() => setFilters({ name: "", company: "", startFrom: "", startTo: "", daysMin: "", attendanceMin: "", attendanceMax: "", calloutsMin: "", noShowsMax: "", lateMax: "", consecutiveMin: "", trainingMin: "", badge: "" })}>Clear Filters{filterCount ? ` (${filterCount})` : ""}</button></div></div>
      <div className="filters-grid">
        <label>Name<input value={filters.name} onChange={(e) => updateFilter("name", e.target.value)} placeholder="Search name" /></label>
        <label>Agency<select value={filters.company} onChange={(e) => updateFilter("company", e.target.value)}><option value="">All</option>{companies.map((company) => <option key={company}>{company}</option>)}</select></label>
        <label>Start From<input type="date" value={filters.startFrom} onChange={(e) => updateFilter("startFrom", e.target.value)} /></label>
        <label>Start To<input type="date" value={filters.startTo} onChange={(e) => updateFilter("startTo", e.target.value)} /></label>
        <label>Min Days Active<input type="number" min="0" value={filters.daysMin} onChange={(e) => updateFilter("daysMin", e.target.value)} /></label>
        <label>Min Attendance %<input type="number" min="0" max="100" value={filters.attendanceMin} onChange={(e) => updateFilter("attendanceMin", e.target.value)} /></label>
        <label>Max Attendance %<input type="number" min="0" max="100" value={filters.attendanceMax} onChange={(e) => updateFilter("attendanceMax", e.target.value)} /></label>
        <label>Min Callouts<input type="number" min="0" value={filters.calloutsMin} onChange={(e) => updateFilter("calloutsMin", e.target.value)} /></label>
        <label>Max No Shows<input type="number" min="0" value={filters.noShowsMax} onChange={(e) => updateFilter("noShowsMax", e.target.value)} /></label>
        <label>Max Late<input type="number" min="0" value={filters.lateMax} onChange={(e) => updateFilter("lateMax", e.target.value)} /></label>
        <label>Min Consecutive Days<input type="number" min="0" value={filters.consecutiveMin} onChange={(e) => updateFilter("consecutiveMin", e.target.value)} /></label>
        <label>Min Training %<input type="number" min="0" max="100" value={filters.trainingMin} onChange={(e) => updateFilter("trainingMin", e.target.value)} /></label>
        <label>Badge<select value={filters.badge} onChange={(e) => updateFilter("badge", e.target.value)}><option value="">All</option><option>Excellent</option><option>Good</option><option>Watchlist</option><option>Needs Attention</option></select></label>
      </div>
      <div className="table-wrap"><table><thead><tr><th>{sortHeader("name", "Name")}</th><th>{sortHeader("company", "Agency")}</th><th>{sortHeader("startDate", "Start Date")}</th><th>{sortHeader("daysActive", "Days Active")}</th><th>{sortHeader("attendance", "Attendance")}</th><th>{sortHeader("callouts", "Total Callouts")}</th><th>{sortHeader("noShows", "No Shows")}</th><th>{sortHeader("late", "Late")}</th><th>{sortHeader("consecutive", "Consecutive Days")}</th><th>{sortHeader("training", "Training")}</th><th>{sortHeader("badge", "Badge")}</th></tr></thead><tbody>
        {visibleRows.length ? visibleRows.map(({ contractor: c, ...row }) => <tr key={c.id}><td>{c.name}</td><td>{c.company || "-"}</td><td>{c.startDate || "-"}</td><td>{row.daysActive}</td><td>{row.attendance}%</td><td>{row.callouts}</td><td>{row.noShows}</td><td>{row.late}</td><td>{row.consecutive}</td><td>{row.training}%</td><td><span className={`status ${row.badge.toLowerCase().replaceAll(" ", "-")}`}>{row.badge}</span></td></tr>) : <tr><td colSpan="11">No contractors match these filters.</td></tr>}
      </tbody></table></div>
    </section>
  );
}

function WorkforceAnalytics({ state }) {
  const contractorCompanyByName = Object.fromEntries(state.contractors.map((contractor) => [contractor.name, contractor.company]));
  const byCompany = groupCompanyCallouts(state.callouts, contractorCompanyByName);
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

function normalizeCompany(company) {
  const value = String(company || "").trim().toLowerCase();
  if (value.includes("telus")) return "TELUS";
  if (value.includes("dekra")) return "DEKRA";
  if (value.includes("kett")) return "KETT";
  return value ? "Other" : "Unknown";
}

function groupCompanyCallouts(callouts, contractorCompanyByName = {}) {
  const companyCounts = { TELUS: 0, DEKRA: 0, KETT: 0 };
  callouts.forEach((callout) => {
    const company = normalizeCompany(callout.company || contractorCompanyByName[callout.name]);
    companyCounts[company] = (companyCounts[company] || 0) + 1;
  });
  return companyCounts;
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

function inDateRange(date, start, end) {
  const value = dateOnly(date);
  return (!start || value >= start) && (!end || value <= end);
}

function attendanceMarks(record) {
  return [record.monday, record.tuesday, record.wednesday, record.thursday, record.friday].filter(Boolean);
}

function reportDataFor(type, state, startDate, endDate) {
  const attendance = state.attendanceRecords.filter((record) => inDateRange(record.weekStart, startDate, endDate));
  const callouts = state.callouts.filter((callout) => inDateRange(callout.date, startDate, endDate));

  if (type === "Daily attendance") {
    const dayNames = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
    const dayKeys = ["monday", "tuesday", "wednesday", "thursday", "friday"];
    return {
      headers: ["Week Start", "Day", "Present", "Absent", "Tardy", "Week Off"],
      rows: dayKeys.map((key, index) => [attendance[0]?.weekStart || startDate || today, dayNames[index], ...attendanceStatuses.map((status) => attendance.filter((record) => record[key] === status).length)])
    };
  }

  if (type === "Weekly attendance") {
    return {
      headers: ["Week Start", "Name", "Company", "P%", "A%", "Tardy %", "Comments"],
      rows: attendance.map((record) => [record.weekStart, record.contractorName, record.company, `${weeklyAttendancePercent(record, "P")}%`, `${weeklyAttendancePercent(record, "A")}%`, `${weeklyAttendancePercent(record, "T")}%`, record.comments])
    };
  }

  if (type === "Monthly attendance") {
    const grouped = attendance.reduce((acc, record) => {
      const month = dateOnly(record.weekStart).slice(0, 7) || "Unknown";
      const marks = attendanceMarks(record);
      acc[month] ??= { P: 0, A: 0, T: 0, WO: 0 };
      attendanceStatuses.forEach((status) => acc[month][status] += marks.filter((mark) => mark === status).length);
      return acc;
    }, {});
    return { headers: ["Month", "Present", "Absent", "Tardy", "Week Off"], rows: Object.entries(grouped).map(([month, counts]) => [month, counts.P, counts.A, counts.T, counts.WO]) };
  }

  if (type === "Vehicle utilization") {
    return {
      headers: ["Vehicle", "Model", "Plate", "Status", "Location", "Assigned Contractor"],
      rows: state.vehicles.map((vehicle) => [vehicle.name, vehicle.model, vehicle.licensePlate, vehicle.status, vehicle.location, vehicle.assignedContractorName || "Unpaired"])
    };
  }

  if (type === "Operator utilization") {
    return {
      headers: ["Name", "Company", "Status", "Attendance %", "Callouts", "Late", "Vehicle"],
      rows: state.contractors.map((contractor) => {
        const vehicle = state.vehicles.find((item) => item.assignedContractorId === contractor.id || item.assignedContractorName === contractor.name);
        const stats = attendanceStatsForContractor(contractor, state.callouts, state.attendanceRecords);
        return [contractor.name, contractor.company, contractor.status, `${stats.attendance}%`, calloutsForContractor(contractor, state.callouts).length, stats.tardy, vehicle?.name || "Unpaired"];
      })
    };
  }

  if (type === "Callout trends") {
    const grouped = callouts.reduce((acc, callout) => {
      const key = `${dateOnly(callout.date)} / ${callout.reason || "Other"}`;
      acc[key] = (acc[key] || 0) + 1;
      return acc;
    }, {});
    return { headers: ["Date / Reason", "Callouts"], rows: Object.entries(grouped).map(([label, count]) => [label, count]) };
  }

  const lateNotes = state.notes.filter((note) => inDateRange(note.date, startDate, endDate) && note.text.toLowerCase().includes("late"));
  return {
    headers: ["Name / Log", "Late Count"],
    rows: [...state.contractors.map((contractor) => [contractor.name, contractor.lateArrivals ?? 0]), ...lateNotes.map((note) => [note.text, 1])]
  };
}

function ReportsGenerator({ state, actions }) {
  const [type, setType] = useState("Daily attendance");
  const [startDate, setStartDate] = useState(today);
  const [endDate, setEndDate] = useState(today);
  const [restoreStatus, setRestoreStatus] = useState("");
  const report = useMemo(() => reportDataFor(type, state, startDate, endDate), [type, state, startDate, endDate]);
  const totalRows = report.rows.length;
  const populatedRows = report.rows.filter((row) => row.some((cell) => cell !== "" && cell !== 0 && cell !== "0%" && cell !== "Unpaired")).length;

  function exportReport() {
    downloadCsv(`${type.toLowerCase().replaceAll(" ", "-")}-${startDate || "all"}-${endDate || "all"}.csv`, report.headers, report.rows);
  }

  function exportBackup() {
    downloadJson(`fleetops-backup-${today}.json`, {
      exportedAt: new Date().toISOString(),
      version: 1,
      data: state
    });
    setRestoreStatus("Backup downloaded.");
  }

  async function restoreBackup(event) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    if (!(await confirmDelete("current dashboard data before restoring backup"))) return;
    try {
      const backup = JSON.parse(await file.text());
      const data = backup.data ?? backup;
      const keys = ["contractors", "callouts", "tickets", "laptops", "vehicles", "attendanceRecords", "notes", "notifications"];
      if (!keys.every((key) => Array.isArray(data[key]))) throw new Error("Missing FleetOps data tables.");
      actions.saveContractors(data.contractors);
      actions.saveCallouts(data.callouts);
      actions.saveTickets(data.tickets);
      actions.saveLaptops(data.laptops);
      actions.saveVehicles(data.vehicles);
      actions.saveAttendanceRecords(data.attendanceRecords);
      actions.saveNotes(data.notes);
      actions.saveNotifications(data.notifications);
      setRestoreStatus("Backup restored.");
    } catch (error) {
      setRestoreStatus(`Restore failed: ${error.message}`);
    }
  }

  return (
    <>
      <section className="panel reports-panel">
        <div className="section-heading"><div><p className="eyebrow">Reports</p><h2>Reports Generator</h2></div><button className="primary-button" onClick={exportReport}>Export CSV</button></div>
        <div className="reports-controls">
          <label>Report Type<select value={type} onChange={(e) => setType(e.target.value)}>{reportTypes.map((item) => <option key={item}>{item}</option>)}</select></label>
          <label>Start Date<input type="date" value={startDate} onChange={(e) => setStartDate(e.target.value)} /></label>
          <label>End Date<input type="date" value={endDate} onChange={(e) => setEndDate(e.target.value)} /></label>
        </div>
        <section className="pairing-summary reports-summary">
          <article><span>Report Rows</span><strong>{totalRows}</strong></article>
          <article><span>Populated Rows</span><strong>{populatedRows}</strong></article>
          <article><span>Date Range</span><strong>{startDate || "All"} - {endDate || "All"}</strong></article>
        </section>
        <div className="table-wrap">
          <table>
            <thead><tr>{report.headers.map((header) => <th key={header}>{header}</th>)}</tr></thead>
            <tbody>{report.rows.length ? report.rows.map((row, index) => <tr key={`${type}-${index}`}>{row.map((cell, cellIndex) => <td key={cellIndex}>{cell}</td>)}</tr>) : <tr><td colSpan={report.headers.length}>No report data for this selection.</td></tr>}</tbody>
          </table>
        </div>
      </section>
      <section className="panel backup-panel">
        <div className="section-heading"><div><p className="eyebrow">Backup / Restore</p><h2>Data Protection</h2></div><span className="date-pill">{state.contractors.length + state.callouts.length + state.vehicles.length} Records</span></div>
        <div className="backup-actions">
          <button className="primary-button" type="button" onClick={exportBackup}>Download Backup</button>
          <label className="ghost-button restore-upload">Restore Backup<input type="file" accept="application/json,.json" onChange={restoreBackup} /></label>
          {restoreStatus && <span>{restoreStatus}</span>}
        </div>
      </section>
    </>
  );
}

function OperationsNotes({ notes, saveNotes, opsShift = "" }) {
  const [draft, setDraft] = useState("");
  const [date, setDate] = useState(today);
  const [search, setSearch] = useState("");
  const filtered = notes.filter((n) => (!search || n.text.toLowerCase().includes(search.toLowerCase())) && (!date || n.date === date));
  function addNote(event) {
    event.preventDefault();
    if (!draft.trim()) return;
    saveNotes([{ id: crypto.randomUUID(), date, text: `${opsShift ? `${opsShift}: ` : ""}${draft}` }, ...notes]);
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

function NotificationsFeed({ notifications, saveNotifications }) {
  async function removeNotification(id) {
    if (!(await confirmDelete("this notification"))) return;
    saveNotifications(notifications.filter((notification) => notification.id !== id));
  }
  async function clearNotifications() {
    if (!(await confirmDelete("all notifications"))) return;
    saveNotifications([]);
  }

  return (
    <section className="panel callout-log">
      <div className="section-heading"><div><p className="eyebrow">Slack-Ready Feed</p><h2>Operations Notifications</h2></div>{notifications.length ? <button className="ghost-button" onClick={clearNotifications}>Clear All</button> : null}</div>
      <div className="notes-list">{notifications.length ? notifications.map((n) => <article className="notification-item" key={n.id}><div><strong>{n.time}</strong><p>{n.message}</p></div><button className="danger ghost-button" onClick={() => removeNotification(n.id)}>Delete</button></article>) : <p>No notifications yet.</p>}</div>
    </section>
  );
}

function AiAssistant({ state, metrics }) {
  const [prompt, setPrompt] = useState("Generate daily staffing recommendations.");
  const answer = useMemo(() => generateAiAnswer(prompt, state, metrics), [prompt, state, metrics]);
  const insights = useMemo(() => generateAiInsights(state, metrics), [state, metrics]);
  const promptGroups = [
    { title: "Staffing", prompts: ["Generate daily staffing recommendations.", "Predict staffing shortages before shift starts.", "Which operators are available after 2 PM?"] },
    { title: "Vehicles", prompts: ["Recommend optimal vehicle pairings.", "Who can cover Vehicle RS1-10234?", "Suggest vehicle rotations."] },
    { title: "Risk", prompts: ["Forecast attendance trends and callout risk.", "Detect operators approaching overtime limits.", "Which contractor companies have the most callouts?"] },
    { title: "Summary", prompts: ["Generate end-of-shift summary.", "Which contractors have the highest attendance?"] }
  ];
  return (
    <>
      <section className="ai-hero panel">
        <div>
          <p className="eyebrow">AI Workforce Assistant</p>
          <h2>Operations Intelligence</h2>
          <span>Uses live staffing, attendance, callout, coverage, and vehicle pairing data from this dashboard.</span>
        </div>
        <strong>{metrics.readiness}% Ready</strong>
      </section>
      <section className="analytics-grid ai-insights-grid">
        {insights.map((insight) => <article className="panel ai-insight-card" key={insight.title}><p className="eyebrow">{insight.type}</p><h2>{insight.title}</h2><strong>{insight.value}</strong><span>{insight.detail}</span></article>)}
      </section>
      <section className="panel ai-workspace">
        <div className="section-heading"><div><p className="eyebrow">Ask / Generate</p><h2>Planning Recommendations</h2></div></div>
        <div className="assistant-modern-layout">
          <aside className="prompt-library">
            {promptGroups.map((group) => (
              <section key={group.title}>
                <h3>{group.title}</h3>
                {group.prompts.map((item) => <button className={prompt === item ? "active" : ""} key={item} onClick={() => setPrompt(item)}>{item}</button>)}
              </section>
            ))}
          </aside>
          <section className="assistant-compose">
            <label>Ask the assistant<textarea value={prompt} onChange={(e) => setPrompt(e.target.value)} /></label>
            <article className="assistant-answer modern"><span>Recommendation</span><p>{answer}</p></article>
          </section>
        </div>
      </section>
    </>
  );
}

function availableOperators(state) {
  return state.contractors.filter((c) => isOperationsContractor(c) && !["Called Out", "No Show", "Off Shift", "Inactive", "Offboarding"].includes(c.status));
}

function highRiskOperators(state) {
  return state.contractors
    .map((contractor) => {
      const stats = attendanceStatsForContractor(contractor, state.callouts, state.attendanceRecords);
      return { contractor, callouts: calloutsForContractor(contractor, state.callouts).length, attendance: stats.attendance, tardy: stats.tardy, absences: stats.absences };
    })
    .filter((item) => item.callouts >= 2 || item.attendance < 90 || item.tardy >= 2 || item.absences >= 2)
    .sort((a, b) => a.attendance - b.attendance || b.callouts - a.callouts);
}

function generateAiInsights(state, metrics) {
  const available = availableOperators(state);
  const openVehicles = state.vehicles.filter((vehicle) => !vehicle.assignedContractorId && !vehicle.assignedContractorName);
  const pairedVehicles = state.vehicles.filter((vehicle) => vehicle.assignedContractorId || vehicle.assignedContractorName);
  const risk = highRiskOperators(state);
  const overtimeWatch = state.contractors.filter((contractor) => (Number(contractor.attendanceDays) || 0) >= 5 || (Number(contractor.lateArrivals) || 0) >= 3);
  return [
    { type: "Pairing", title: "Optimal Pairing Pool", value: Math.min(available.length, openVehicles.length), detail: `${available.length} available operators and ${openVehicles.length} unpaired vehicles.` },
    { type: "Staffing", title: "Shortage Risk", value: metrics.openTickets.length || Math.max(0, metrics.scheduled.length - metrics.available), detail: metrics.available < metrics.scheduled.length ? "Coverage is below scheduled demand." : "No immediate shortage detected." },
    { type: "Attendance", title: "Callout Risk", value: risk.length, detail: risk.slice(0, 3).map((item) => item.contractor.name).join(", ") || "No high-risk operators flagged." },
    { type: "Utilization", title: "Overtime Watch", value: overtimeWatch.length, detail: overtimeWatch.slice(0, 3).map((contractor) => contractor.name).join(", ") || "No overtime risk indicators." },
    { type: "Vehicles", title: "Rotation Candidates", value: pairedVehicles.length, detail: pairedVehicles.slice(0, 3).map((vehicle) => vehicle.name).join(", ") || "No paired vehicles yet." },
    { type: "Summary", title: "End-of-Shift Ready", value: `${metrics.readiness}%`, detail: `${metrics.scheduled.length} scheduled, ${metrics.todayCallouts.length} callouts, ${metrics.openTickets.length} open gaps.` }
  ];
}

function generateAiAnswer(prompt, state, metrics) {
  const lower = prompt.toLowerCase();
  const available = availableOperators(state);
  const openVehicles = state.vehicles.filter((vehicle) => !vehicle.assignedContractorId && !vehicle.assignedContractorName);
  if (lower.includes("pair")) return openVehicles.slice(0, 5).map((vehicle, index) => `${vehicle.name}: ${available[index]?.name || "No available operator"}${available[index] ? ` (${available[index].company})` : ""}`).join("; ") || "No unpaired vehicles or available operators found.";
  if (lower.includes("shortage")) return metrics.available < metrics.scheduled.length ? `Shortage risk: ${metrics.scheduled.length - metrics.available} operator gap. Open coverage tickets: ${metrics.openTickets.length}.` : `No shortage predicted. ${metrics.available} available for ${metrics.scheduled.length} scheduled.`;
  if (lower.includes("risk") || lower.includes("forecast")) return highRiskOperators(state).slice(0, 5).map((item) => `${item.contractor.name}: ${item.attendance}% attendance, ${item.callouts} callouts, ${item.tardy} tardy, ${item.absences} absent`).join("; ") || "No attendance or callout risk flags found.";
  if (lower.includes("overtime")) return state.contractors.filter((c) => (Number(c.attendanceDays) || 0) >= 5 || (Number(c.lateArrivals) || 0) >= 3).map((c) => `${c.name}: ${c.attendanceDays ?? 0} consecutive attendance days`).join("; ") || "No operators approaching overtime indicators.";
  if (lower.includes("rotation")) return state.vehicles.filter((v) => v.assignedContractorName).map((v) => `${v.name}: currently paired to ${v.assignedContractorName}; consider rotating if utilization is high.`).join("; ") || "No vehicle pairings available for rotation suggestions.";
  if (lower.includes("end-of-shift") || lower.includes("summary")) return `End-of-shift summary: ${metrics.scheduled.length} scheduled, ${metrics.available} available, ${metrics.todayCallouts.length} callouts, ${metrics.openTickets.length} open gaps, readiness ${metrics.readiness}%.`;
  if (lower.includes("staffing recommendations")) return `Staffing recommendation: keep ${Math.max(0, metrics.scheduled.length - metrics.available)} backup operators identified, close ${metrics.openTickets.length} coverage tickets, and prioritize ${available.slice(0, 3).map((c) => c.name).join(", ") || "available operators"} for coverage.`;
  if (lower.includes("vehicle")) {
    const vehicleQuery = state.vehicles.find((vehicle) => lower.includes(vehicle.name.toLowerCase()) || lower.includes((vehicle.licensePlate || "").toLowerCase()));
    if (vehicleQuery) return `${vehicleQuery.name} can be covered by ${available.slice(0, 5).map((c) => `${c.name} (${c.company})`).join(", ") || "no currently available operators"}. Current pairing: ${vehicleQuery.assignedContractorName || "Unpaired"}.`;
  }
  if (lower.includes("after 2") || lower.includes("after 2 pm") || lower.includes("available")) return available.map((c) => `${c.name} (${normalizeShift(c.shift)})`).join(", ") || "No available contractors found.";
  if (lower.includes("highest attendance")) return [...state.contractors].sort((a, b) => attendancePercent(b, state.callouts, state.attendanceRecords) - attendancePercent(a, state.callouts, state.attendanceRecords)).slice(0, 3).map((c) => `${c.name}: ${attendancePercent(c, state.callouts, state.attendanceRecords)}%`).join(", ");
  if (lower.includes("companies")) {
    const contractorCompanyByName = Object.fromEntries(state.contractors.map((contractor) => [contractor.name, contractor.company]));
    return Object.entries(groupCompanyCallouts(state.callouts, contractorCompanyByName)).sort((a, b) => b[1] - a[1]).map(([company, count]) => `${company}: ${count}`).join(", ") || "No callouts by company yet.";
  }
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

function ContractorAdmin({ contractors, callouts, onSaveContractors, defaultShift = "AM Shift" }) {
  const [editingId, setEditingId] = useState(null);
  const [showForm, setShowForm] = useState(false);
  const blank = { name: "", email: "", phone: "", company: "", startDate: today, shift: defaultShift, scheduledDays: ["Mon", "Tue", "Wed", "Thu", "Fri"], active: true, status: "Scheduled", statusDate: "", attendance: 95, noShows: 0, lateArrivals: 0, attendanceDays: 0, trainingCompletion: 100 };
  const [draft, setDraft] = useState(blank);
  useEffect(() => {
    if (!editingId) setDraft((current) => ({ ...current, shift: defaultShift }));
  }, [defaultShift, editingId]);
  function startEdit(contractor) {
    setEditingId(contractor.id);
    setDraft({ ...blank, ...contractor, shift: normalizeShift(contractor.shift) });
    setShowForm(true);
  }
  function resetForm() {
    setEditingId(null);
    setDraft(blank);
    setShowForm(false);
  }
  function saveDraft(event) {
    event.preventDefault();
    if (!draft.name.trim()) return;
    const normalizedDraft = {
      ...draft,
      active: draft.status === "Offboarding" ? false : draft.active,
      statusDate: dailyResetStatuses.includes(draft.status) ? (draft.statusDate || today) : ""
    };
    const next = editingId ? contractors.map((c) => c.id === editingId ? normalizedDraft : c) : [{ ...normalizedDraft, id: crypto.randomUUID() }, ...contractors];
    onSaveContractors(next);
    resetForm();
  }
  async function removeContractor(id) {
    if (!(await confirmDelete("this contractor"))) return;
    onSaveContractors(contractors.filter((c) => c.id !== id));
  }
  function toggleDay(day) {
    setDraft({ ...draft, scheduledDays: draft.scheduledDays.includes(day) ? draft.scheduledDays.filter((d) => d !== day) : [...draft.scheduledDays, day] });
  }
  return (
    <section className="panel">
      <div className="section-heading">
        <div><p className="eyebrow">Admin</p><h2>Contractors</h2></div>
        <div className="button-row"><span className="date-pill">{contractors.length} Contractors</span><span className="date-pill">{todayDay} {today}</span><button className="primary-button" onClick={() => setShowForm(true)}>Add Contractor</button></div>
      </div>
      <ContractorTable contractors={contractors} callouts={callouts} onEdit={startEdit} onRemove={removeContractor} />
      {showForm && (
        <div className="modal-backdrop" role="dialog" aria-modal="true" aria-label="Contractor form">
          <div className="panel modal-panel">
            <div className="section-heading"><div><p className="eyebrow">{editingId ? "Edit" : "Add"}</p><h2>Contractor</h2></div><button className="ghost-button" type="button" onClick={resetForm}>Close</button></div>
            <form className="form-grid" onSubmit={saveDraft}>
              <label>Name<input value={draft.name} onChange={(e) => setDraft({ ...draft, name: e.target.value })} required /></label>
              <label>Email<input type="email" value={draft.email} onChange={(e) => setDraft({ ...draft, email: e.target.value })} /></label>
              <label>Phone<input value={draft.phone} onChange={(e) => setDraft({ ...draft, phone: e.target.value })} /></label>
              <label>Company<input value={draft.company} onChange={(e) => setDraft({ ...draft, company: e.target.value })} /></label>
              <label>Start Date<input type="date" value={draft.startDate} onChange={(e) => setDraft({ ...draft, startDate: e.target.value })} /></label>
              <label>Shift<select value={draft.shift} onChange={(e) => setDraft({ ...draft, shift: e.target.value })}>{shifts.map((s) => <option key={s}>{s}</option>)}</select></label>
              <label>Status<select value={draft.status} onChange={(e) => setDraft({ ...draft, status: e.target.value, active: e.target.value === "Offboarding" ? false : draft.active })}>{workforceStatuses.map((s) => <option key={s}>{s}</option>)}</select></label>
              <fieldset><legend>Scheduled Days</legend><div className="day-grid">{days.map((day) => <label className="check-label" key={day}><input type="checkbox" checked={draft.scheduledDays.includes(day)} onChange={() => toggleDay(day)} />{day}</label>)}</div></fieldset>
              <label className="check-label inline"><input type="checkbox" checked={draft.active} onChange={(e) => setDraft({ ...draft, active: e.target.checked })} />Active</label>
              <div className="button-row wide"><button className="primary-button">{editingId ? "Save" : "Add"}</button><button className="ghost-button" type="button" onClick={resetForm}>Cancel</button></div>
            </form>
          </div>
        </div>
      )}
    </section>
  );
}

function ContractorTable({ contractors, callouts, onEdit, onRemove }) {
  const [sortBy, setSortBy] = useState("name");
  const sortedContractors = [...contractors].sort((a, b) => {
    const primaryKey = sortBy === "company" ? "company" : sortBy === "startDate" ? "startDate" : "name";
    const secondaryKey = sortBy === "company" ? "name" : "company";
    const primary = primaryKey === "startDate" ? compareDate(a.startDate, b.startDate) : compareText(a[primaryKey], b[primaryKey]);
    if (primary) return primary;
    return compareText(a[secondaryKey], b[secondaryKey]);
  });
  function lastCalloutFor(name) {
    return callouts.find((callout) => callout.name.toLowerCase() === name.toLowerCase())?.date ?? "None";
  }
  return (
    <div className="table-wrap"><table><thead><tr><th><button className="sort-button" type="button" onClick={() => setSortBy("name")}>Name{sortBy === "name" ? " ↑" : ""}</button></th><th><button className="sort-button" type="button" onClick={() => setSortBy("company")}>Company{sortBy === "company" ? " ↑" : ""}</button></th><th><button className="sort-button" type="button" onClick={() => setSortBy("startDate")}>Start Date{sortBy === "startDate" ? " ↑" : ""}</button></th><th>Shift</th><th>Scheduled Today</th><th>Status</th><th>Last Callout</th><th>Actions</th></tr></thead><tbody>
      {sortedContractors.map((c) => {
        const status = c.status === "Offboarding" ? "Offboarding" : c.active ? c.status : "Inactive";
        return <tr key={c.id}><td><strong>{c.name}</strong></td><td>{c.company}</td><td>{c.startDate || "-"}</td><td>{normalizeShift(c.shift)}</td><td>{isOperationsContractor(c) && c.scheduledDays.includes(todayDay) ? "Yes" : "No"}</td><td><span className={`status ${(status ?? "scheduled").toLowerCase().replaceAll(" ", "-")}`}>{status}</span></td><td>{lastCalloutFor(c.name)}</td><td><div className="row-actions"><button onClick={() => onEdit(c)}>Edit</button><button className="danger" onClick={() => onRemove(c.id)}>Remove</button></div></td></tr>;
      })}
    </tbody></table></div>
  );
}

createRoot(document.getElementById("root")).render(<App />);
