import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { Upload, BarChart3, Users, Target, MessageSquare, Eye, Timer, Search, AlertTriangle, Settings, ClipboardCheck, TrendingUp, CalendarDays, Moon, Sun, Info, X, ChevronDown } from "lucide-react";
import { ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, LineChart, Line, PieChart, Pie, Cell } from "recharts";

const COLORS = ["#2563eb", "#16a34a", "#f59e0b", "#dc2626", "#7c3aed", "#0891b2", "#ea580c", "#64748b"];
const DEFAULT_GOALS = { recordingCompliance: 0.8, scriptCompliance: 0.8, talkRatioMax: 0.6, minimumViewsRatio: 0.35 };
const DEFAULT_TEAM_VOLUME_GOALS = { WMASS: 12000000, EMASS: 12000000, Albany: 8000000, CT: 8000000, Exteriors: 2000000 };
const DEMO_WORKBOOK_NAME = "Rilla-Usage-2025-06-23-to-2026-06-27.xlsx";
const DEMO_WORKBOOK_URL = `/${DEMO_WORKBOOK_NAME}`;
const SHEET_ALIASES = {
  userManagement: ["user_management", "user management", "users"],
  userUsage: ["user_usage", "user usage", "usage"],
  teamUsage: ["team_usage", "team usage", "teams"],
  conversations: ["conversations", "conversation", "calls", "recordings"],
};

const cx = (...parts) => parts.filter(Boolean).join(" ");
const cleanSheetName = (name) => String(name || "").trim().toLowerCase();
const cleanKey = (key) => String(key || "").toLowerCase().replace(/[^a-z0-9]/g, "");

function getValue(row, names) {
  if (!row) return null;
  const keys = Object.keys(row);
  const index = new Map(keys.map((key) => [cleanKey(key), key]));
  for (const name of names) {
    const key = index.get(cleanKey(name));
    if (key !== undefined) return row[key];
  }
  return null;
}

function numberValue(value, fallback = 0) {
  if (value === null || value === undefined || value === "") return fallback;
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "boolean") return value ? 1 : 0;
  const parsed = Number(String(value).replace(/[,$% ]/g, ""));
  return Number.isFinite(parsed) ? parsed : fallback;
}

function percentValue(value, fallback = null) {
  const n = numberValue(value, NaN);
  if (!Number.isFinite(n)) return fallback;
  return n > 1 ? n / 100 : n;
}

function boolValue(value) {
  if (typeof value === "boolean") return value;
  return ["true", "yes", "1"].includes(String(value || "").trim().toLowerCase());
}

function parseDate(value) {
  if (!value) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number") {
    const date = new Date(Math.floor(value - 25569) * 86400 * 1000);
    return Number.isNaN(date.getTime()) ? null : date;
  }
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function fmtInt(value) { return Math.round(numberValue(value, 0)).toLocaleString(); }
function fmtPct(value, digits = 0) { return value === null || value === undefined || Number.isNaN(value) ? "—" : `${(value * 100).toFixed(digits)}%`; }
function fmtMoney(value) { return numberValue(value, 0).toLocaleString(undefined, { style: "currency", currency: "USD", maximumFractionDigits: 0 }); }
function fmtDuration(seconds) {
  const total = numberValue(seconds, 0);
  if (!total) return "—";
  const minutes = Math.round(total / 60);
  const hours = Math.floor(minutes / 60);
  const mins = minutes % 60;
  return hours ? `${hours}h ${mins}m` : `${mins}m`;
}
function shortDate(date) { return date ? date.toLocaleDateString(undefined, { month: "short", day: "numeric" }) : "—"; }
function isoDate(date) {
  if (!date) return "";
  const copy = new Date(date);
  copy.setMinutes(copy.getMinutes() - copy.getTimezoneOffset());
  return copy.toISOString().slice(0, 10);
}

function startOfDay(date) { const d = new Date(date); d.setHours(0, 0, 0, 0); return d; }
function endOfDay(date) { const d = new Date(date); d.setHours(23, 59, 59, 999); return d; }
function startOfMonth(date) { return new Date(date.getFullYear(), date.getMonth(), 1); }
function startOfQuarter(date) { return new Date(date.getFullYear(), Math.floor(date.getMonth() / 3) * 3, 1); }
function startOfYear(date) { return new Date(date.getFullYear(), 0, 1); }
function subDays(date, days) { const d = new Date(date); d.setDate(d.getDate() - days); return d; }
function getWeekStart(date) {
  const d = new Date(date);
  const day = d.getDay();
  d.setDate(d.getDate() - day + (day === 0 ? -6 : 1));
  d.setHours(0, 0, 0, 0);
  return d;
}

function getDateWindow(mode, maxDate, customStart, customEnd) {
  if (!maxDate || mode === "all") return { start: null, end: null };
  const end = endOfDay(maxDate);
  if (mode === "custom") return { start: customStart ? startOfDay(new Date(`${customStart}T00:00:00`)) : null, end: customEnd ? endOfDay(new Date(`${customEnd}T00:00:00`)) : null };
  if (mode === "mtd") return { start: startOfMonth(maxDate), end };
  if (mode === "qtd") return { start: startOfQuarter(maxDate), end };
  if (mode === "ytd") return { start: startOfYear(maxDate), end };
  const days = Number(mode);
  return Number.isFinite(days) ? { start: startOfDay(subDays(maxDate, days - 1)), end } : { start: null, end: null };
}

function dateRangeLabel(mode, start, end) {
  if (mode === "all") return "All conversation rows";
  if (mode === "mtd") return "Month to date";
  if (mode === "qtd") return "Quarter to date";
  if (mode === "ytd") return "Year to date";
  if (["7", "30", "60", "90"].includes(String(mode))) return `Last ${mode} days in export`;
  return `${start ? shortDate(start) : "Start"} – ${end ? shortDate(end) : "End"}`;
}

function parseJsonArray(value) {
  if (!value || typeof value !== "string") return [];
  try { const parsed = JSON.parse(value); return Array.isArray(parsed) ? parsed : []; } catch { return []; }
}

function teamNamesFromRow(row) {
  const raw = getValue(row, ["Teams", "Team", "Name"]);
  const parsed = parseJsonArray(raw);
  if (parsed.length) return parsed.map((item) => item?.name).filter(Boolean);
  if (!raw) return [];
  return String(raw).split(/[;,]/).map((item) => item.trim()).filter(Boolean);
}

function scriptScoreFromScripts(raw) {
  const scripts = parseJsonArray(raw);
  let weighted = 0;
  let weight = 0;
  for (const script of scripts) {
    const score = percentValue(script?.average_script_score, null);
    const count = numberValue(script?.number_of_conversations, 0);
    if (score !== null && count > 0) { weighted += score * count; weight += count; }
  }
  return weight ? weighted / weight : null;
}

function checklistScoreFromRow(row) {
  const checks = parseJsonArray(getValue(row, ["Checklists"]));
  let total = 0;
  let count = 0;
  for (const check of checks) {
    const score = numberValue(check?.checklist_score, NaN);
    const max = numberValue(check?.checklist_denominator || check?.checklist_total || check?.total || check?.max_score, NaN);
    if (Number.isFinite(score) && Number.isFinite(max) && max > 0) { total += score / max; count += 1; }
  }
  return count ? total / count : null;
}

function checklistEntriesFromRow(row) {
  return parseJsonArray(getValue(row, ["Checklists"])).map((check) => {
    const score = numberValue(check?.checklist_score, NaN);
    const denominator = numberValue(check?.checklist_denominator || check?.checklist_total || check?.total || check?.max_score, NaN);
    const trackers = Array.isArray(check?.tracker_data) ? check.tracker_data : [];
    const missedTrackers = trackers.filter((tracker) => !boolValue(tracker?.tracker_is_hit));
    return {
      name: check?.checklist_name || "Untitled Script",
      score,
      denominator,
      pct: Number.isFinite(score) && Number.isFinite(denominator) && denominator > 0 ? score / denominator : null,
      trackers,
      missedTrackers,
    };
  });
}

function sum(rows, names) { return rows.reduce((acc, row) => acc + numberValue(getValue(row, names), 0), 0); }
function weightedAverage(rows, valueNames, weightNames) {
  let total = 0;
  let weight = 0;
  for (const row of rows) {
    const value = percentValue(getValue(row, valueNames), null) ?? numberValue(getValue(row, valueNames), NaN);
    const w = numberValue(getValue(row, weightNames), 0);
    if (Number.isFinite(value) && w > 0) { total += value * w; weight += w; }
  }
  return weight ? total / weight : null;
}

function downloadCsv(rows, filename) {
  if (!rows?.length) return;
  const headers = Object.keys(rows[0]);
  const newline = String.fromCharCode(10);
  const escape = (value) => {
    const text = value === null || value === undefined ? "" : String(value);
    const quoted = text.includes('"') || text.includes(",") || text.includes(newline);
    return quoted ? `"${text.replace(/"/g, '""')}"` : text;
  };
  const csv = [headers.join(","), ...rows.map((row) => headers.map((h) => escape(row[h])).join(","))].join(newline);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function parseRillaWorkbook(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const sheetMap = new Map(workbook.SheetNames.map((name) => [cleanSheetName(name), name]));
  const findSheet = (aliases) => aliases.map(cleanSheetName).map((alias) => sheetMap.get(alias)).find(Boolean) || workbook.SheetNames.find((name) => aliases.some((alias) => cleanSheetName(name).includes(cleanSheetName(alias))));
  const readRows = (aliases) => {
    const sheetName = findSheet(aliases);
    if (!sheetName) return [];
    return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: null, raw: true }).filter((row) => Object.values(row).some((value) => value !== null && value !== ""));
  };
  return {
    userManagement: readRows(SHEET_ALIASES.userManagement),
    userUsage: readRows(SHEET_ALIASES.userUsage),
    teamUsage: readRows(SHEET_ALIASES.teamUsage),
    conversations: readRows(SHEET_ALIASES.conversations),
  };
}

function metricStatus(value, goal, direction = "above") {
  if (value === null || value === undefined || Number.isNaN(value)) return "neutral";
  if (direction === "above") return value >= goal ? "good" : value >= goal * 0.85 ? "warning" : "bad";
  return value <= goal ? "good" : value <= goal * 1.15 ? "warning" : "bad";
}
function statusClass(status) {
  if (status === "good") return "border-emerald-500/40 bg-emerald-500/10 text-emerald-200";
  if (status === "warning") return "border-amber-500/40 bg-amber-500/10 text-amber-100";
  if (status === "bad") return "border-red-500/40 bg-red-500/10 text-red-100";
  return "border-slate-700 bg-slate-900/70 text-slate-100";
}

function ThemeOverrideStyle() {
  return (
    <style>{`
      .rilla-app[data-theme="light"] [class*="bg-slate-950"] { background-color: rgba(255,255,255,0.96) !important; }
      .rilla-app[data-theme="light"] [class*="bg-slate-900"] { background-color: rgba(248,250,252,0.96) !important; }
      .rilla-app[data-theme="light"] [class*="border-slate-800"], .rilla-app[data-theme="light"] [class*="border-slate-700"] { border-color: #cbd5e1 !important; }
      .rilla-app[data-theme="light"] .text-white, .rilla-app[data-theme="light"] [class*="text-slate-100"], .rilla-app[data-theme="light"] [class*="text-slate-200"] { color: #0f172a !important; }
      .rilla-app[data-theme="light"] [class*="text-slate-300"] { color: #334155 !important; }
      .rilla-app[data-theme="light"] [class*="text-slate-400"], .rilla-app[data-theme="light"] [class*="text-slate-500"] { color: #64748b !important; }
      .rilla-app[data-theme="light"] input, .rilla-app[data-theme="light"] select { background-color: #ffffff !important; color: #0f172a !important; }
      .rilla-app[data-theme="light"] option { background-color: #ffffff; color: #0f172a; }
      .rilla-app[data-theme="light"] tbody tr:hover { background-color: #f1f5f9 !important; }
    `}</style>
  );
}

function KpiCard({ icon: Icon, label, value, sublabel, status = "neutral" }) {
  return (
    <div className={`flex min-h-[118px] rounded-2xl border p-4 shadow-xl shadow-black/10 ${statusClass(status)}`}>
      <div className="flex w-full items-start justify-between gap-3">
        <div>
          <div className="text-xs uppercase tracking-[0.18em] text-slate-400">{label}</div>
          <div className="mt-2 text-2xl font-semibold text-white">{value}</div>
          {sublabel && <div className="mt-1 text-xs text-slate-400">{sublabel}</div>}
        </div>
        {Icon && <Icon className="h-5 w-5 text-slate-400" />}
      </div>
    </div>
  );
}
function Section({ title, subtitle, right, children }) {
  return (
    <section className="rounded-3xl border border-slate-800 bg-slate-950/80 p-5 shadow-2xl shadow-black/20">
      <div className="mb-4 flex flex-col gap-2 md:flex-row md:items-start md:justify-between">
        <div><h2 className="text-lg font-semibold text-white">{title}</h2>{subtitle && <p className="mt-1 text-sm text-slate-400">{subtitle}</p>}</div>
        {right}
      </div>
      {children}
    </section>
  );
}
function EmptyState({ onFile }) {
  return (
    <div className="mx-auto max-w-4xl rounded-3xl border border-dashed border-slate-700 bg-slate-950/80 p-10 text-center shadow-2xl shadow-black/20">
      <div className="mx-auto flex h-16 w-16 items-center justify-center rounded-2xl bg-blue-500/10 text-blue-300"><Upload className="h-8 w-8" /></div>
      <h2 className="mt-6 text-2xl font-semibold text-white">Upload a raw Rilla Excel export</h2>
      <p className="mx-auto mt-3 max-w-2xl text-sm leading-6 text-slate-400">The parser expects Rilla-style sheets named <span className="text-slate-200">user_usage</span>, <span className="text-slate-200">team_usage</span>, <span className="text-slate-200">user_management</span>, and <span className="text-slate-200">conversations</span>.</p>
      <label className="mt-7 inline-flex cursor-pointer items-center gap-2 rounded-2xl bg-blue-600 px-5 py-3 text-sm font-semibold text-white shadow-lg shadow-blue-950/40 transition hover:bg-blue-500"><Upload className="h-4 w-4" />Choose Excel File<input type="file" accept=".xlsx,.xls" className="hidden" onChange={onFile} /></label>
    </div>
  );
}
function DataTable({ rows, columns, maxHeight = "max-h-[460px]", wrapCells = false, sortKey = null, sortDirection = "desc", onSort = null }) {
  const isSortable = typeof onSort === "function";
  return (
    <div className={`overflow-auto rounded-2xl border border-slate-800 ${maxHeight}`}>
      <table className="min-w-full divide-y divide-slate-800 text-sm">
        <thead className="sticky top-0 z-10 bg-slate-900 text-left text-xs uppercase tracking-[0.14em] text-slate-400">
          <tr>
            {columns.map((col) => {
              const activeSort = sortKey === col.key;
              const sortableColumn = isSortable && col.sortable !== false;
              return (
                <th key={col.key} className="whitespace-nowrap px-4 py-3 font-medium">
                  {sortableColumn ? (
                    <button onClick={() => onSort(col.key)} className="inline-flex items-center gap-2 text-left text-inherit hover:text-slate-200">
                      <span>{col.label}</span>
                      <span className={cx("text-[10px]", activeSort ? "text-blue-300" : "text-slate-500")}>
                        {activeSort ? (sortDirection === "asc" ? "▲" : "▼") : "↕"}
                      </span>
                    </button>
                  ) : (
                    col.label
                  )}
                </th>
              );
            })}
          </tr>
        </thead>
        <tbody className="divide-y divide-slate-800 bg-slate-950/70">
          {rows.length ? rows.map((row, index) => <tr key={row.id || row.name || index} className="hover:bg-slate-900/80">{columns.map((col) => <td key={col.key} className={cx("px-4 py-3 text-slate-200", wrapCells ? "whitespace-normal align-top" : "whitespace-nowrap")}>{col.render ? col.render(row) : row[col.key] ?? "—"}</td>)}</tr>) : <tr><td colSpan={columns.length} className="px-4 py-8 text-center text-slate-500">No matching rows.</td></tr>}
        </tbody>
      </table>
    </div>
  );
}

function InfoPill({ active, children, onClick }) {
  return (
    <button onClick={onClick} className={cx("rounded-xl px-3 py-2 text-xs font-semibold transition", active ? "bg-blue-600 text-white" : "bg-slate-900 text-slate-300 hover:bg-slate-800")}>
      {children}
    </button>
  );
}

function MetricDefinition({ label, children }) {
  return (
    <div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-4">
      <div className="text-xs font-semibold uppercase tracking-[0.16em] text-blue-300">{label}</div>
      <p className="mt-2 text-sm leading-6 text-slate-300">{children}</p>
    </div>
  );
}

export default function RillaUsageDashboardApp() {
  const [fileName, setFileName] = useState(DEMO_WORKBOOK_NAME);
  const [raw, setRaw] = useState(null);
  const [error, setError] = useState("");
  const [isLoadingDemo, setIsLoadingDemo] = useState(true);
  const [selectedTeam, setSelectedTeam] = useState("All Teams");
  const [selectedRep, setSelectedRep] = useState("All Reps");
  const [dateMode, setDateMode] = useState("all");
  const [customStart, setCustomStart] = useState("");
  const [customEnd, setCustomEnd] = useState("");
  const [granularity, setGranularity] = useState("day");
  const [search, setSearch] = useState("");
  const [goals, setGoals] = useState(DEFAULT_GOALS);
  const [teamVolumeGoals] = useState(DEFAULT_TEAM_VOLUME_GOALS);
  const [trendLines, setTrendLines] = useState({ conversations: true, scriptCompliance: true, talkRatio: true });
  const [headerCollapsed, setHeaderCollapsed] = useState(true);
  const [isDark, setIsDark] = useState(false);
  const [showInfoPanel, setShowInfoPanel] = useState(false);
  const [infoTab, setInfoTab] = useState("kpis");
  const [repSortKey, setRepSortKey] = useState("conversationsRecorded");
  const [repSortDirection, setRepSortDirection] = useState("desc");

  const timeFrameOptions = [
    { value: "all", label: "All" }, { value: "ytd", label: "YTD" }, { value: "qtd", label: "QTD" }, { value: "mtd", label: "MTD" },
    { value: "7", label: "7D" }, { value: "30", label: "30D" }, { value: "60", label: "60D" }, { value: "90", label: "90D" },
  ];
  const actionButtonClass = "inline-flex min-h-10 items-center justify-center gap-2 rounded-2xl border border-slate-700 bg-slate-900 px-4 py-2 text-center text-sm font-medium text-slate-100 hover:bg-slate-800";

  function resetFilters() {
    setSelectedTeam("All Teams");
    setSelectedRep("All Reps");
    setDateMode("all");
    setCustomStart("");
    setCustomEnd("");
    setSearch("");
    setRepSortKey("conversationsRecorded");
    setRepSortDirection("desc");
  }

  useEffect(() => {
    let cancelled = false;
    async function loadDemoWorkbook() {
      setIsLoadingDemo(true);
      setError("");
      try {
        const response = await fetch(DEMO_WORKBOOK_URL);
        if (!response.ok) throw new Error(`Demo workbook request failed with ${response.status}`);
        const arrayBuffer = await response.arrayBuffer();
        const parsed = parseRillaWorkbook(arrayBuffer);
        if (cancelled) return;
        setRaw(parsed);
        setFileName(DEMO_WORKBOOK_NAME);
        resetFilters();
      } catch (err) {
        console.error(err);
        if (!cancelled) setError("The demo workbook could not be loaded. You can still upload a Rilla .xlsx export manually.");
      } finally {
        if (!cancelled) setIsLoadingDemo(false);
      }
    }
    loadDemoWorkbook();
    return () => {
      cancelled = true;
    };
  }, []);

  async function handleFile(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    setError("");
    setFileName(file.name);
    try {
      const arrayBuffer = await file.arrayBuffer();
      setRaw(parseRillaWorkbook(arrayBuffer));
      resetFilters();
    } catch (err) {
      console.error(err);
      setError("The file could not be parsed. Confirm this is a valid Rilla .xlsx export.");
    }
  }

  const model = useMemo(() => {
    if (!raw) return null;
    const userUsage = raw.userUsage || [];
    const teamUsage = raw.teamUsage || [];
    const userManagement = raw.userManagement || [];
    const conversationsRaw = raw.conversations || [];
    const userByEmail = new Map();
    const userById = new Map();

    userUsage.forEach((row) => {
      const email = String(getValue(row, ["Email", "User Email"]) || "").toLowerCase();
      const id = String(getValue(row, ["User Id", "User ID"]) || "");
      const profile = { row, teams: teamNamesFromRow(row) };
      if (email) userByEmail.set(email, profile);
      if (id) userById.set(id, profile);
    });

    const conversations = conversationsRaw.map((row) => {
      const date = parseDate(getValue(row, ["Date", "Processed Date"]));
      const email = String(getValue(row, ["User Email", "Email"]) || "").toLowerCase();
      const userId = String(getValue(row, ["User Id", "User ID"]) || "");
      const profile = userByEmail.get(email) || userById.get(userId);
      return {
        row,
        date,
        repName: getValue(row, ["User Name", "Name", "Rep"]) || getValue(profile?.row, ["Name"]) || "Unknown",
        email,
        userId,
        teams: profile?.teams || [],
        duration: numberValue(getValue(row, ["Duration In Seconds"]), 0),
        sold: numberValue(getValue(row, ["Total Sold"]), 0),
        outcome: getValue(row, ["Outcome"]) || "Unknown",
        talkRatio: percentValue(getValue(row, ["Rep Talk Ratio", "Talk Ratio"]), null),
        checklistScore: checklistScoreFromRow(row),
        checklists: checklistEntriesFromRow(row),
        rillaUrl: getValue(row, ["Rilla Url", "Rilla URL"]),
      };
    }).filter((row) => row.date);

    const dateValues = conversations.map((c) => c.date).sort((a, b) => a - b);
    const totalAppointments = sum(userUsage, ["Total Appointments"]);
    const appointmentsRecorded = sum(userUsage, ["Appointments Recorded"]);
    const recordedConversations = sum(userUsage, ["Conversations Recorded"]);
    const conversationsViewed = sum(userUsage, ["Conversations Viewed"]);
    const commentsReceived = sum(userUsage, ["Comments Received"]);
    const commentsGiven = sum(userUsage, ["Comments Given"]);
    const activeReps = userUsage.filter((row) => numberValue(getValue(row, ["Conversations Recorded"]), 0) > 0).length;
    const activeUsers = userManagement.filter((row) => !boolValue(getValue(row, ["Is Removed"])) && boolValue(getValue(row, ["Account Set Up"]))).length;

    const exportTotals = {
      totalAppointments,
      appointmentsRecorded,
      recordedConversations,
      detailConversationRows: conversationsRaw.length,
      validDatedConversations: conversations.length,
      conversationsViewed,
      commentsReceived,
      commentsGiven,
      scorecardsReceived: sum(userUsage, ["Scorecards Received"]),
      ridealongsReceived: sum(userUsage, ["Ridealongs Received"]),
      activeReps,
      totalUsers: activeUsers || userUsage.length,
      recordingCompliance: totalAppointments ? appointmentsRecorded / totalAppointments : weightedAverage(userUsage, ["Recording Compliance"], ["Total Appointments"]),
      avgScriptCompliance: weightedAverage(userUsage, ["Average Script Compliance"], ["Conversations Recorded"]),
      avgTalkRatio: weightedAverage(userUsage, ["Talk Ratio Average"], ["Conversations Recorded"]),
      avgDuration: weightedAverage(userUsage, ["Average Conversation Duration"], ["Conversations Recorded"]),
      viewedRecordedRatio: recordedConversations ? conversationsViewed / recordedConversations : weightedAverage(userUsage, ["Viewed Recorded Ratio"], ["Conversations Recorded"]),
      dateStart: dateValues[0] || null,
      dateEnd: dateValues[dateValues.length - 1] || null,
    };

    const teamSummary = teamUsage.map((row) => {
      const name = getValue(row, ["Name", "Team"]);
      const appointments = numberValue(getValue(row, ["Total Appointments"]), 0);
      const recordedAppointments = numberValue(getValue(row, ["Appointments Recorded"]), 0);
      return {
        id: name,
        name,
        totalUsers: numberValue(getValue(row, ["Total Users"]), 0),
        usersWhoRecorded: numberValue(getValue(row, ["Total Users Who Recorded"]), 0),
        appointments,
        recordedAppointments,
        conversationsRecorded: numberValue(getValue(row, ["Conversations Recorded"]), 0),
        recordingCompliance: appointments ? recordedAppointments / appointments : percentValue(getValue(row, ["Recording Compliance"]), null),
        scriptCompliance: scriptScoreFromScripts(getValue(row, ["Scripts"])),
        talkRatio: percentValue(getValue(row, ["Talk Ratio Average"]), null),
        commentsReceived: numberValue(getValue(row, ["Comments Received"]), 0),
        commentsGiven: numberValue(getValue(row, ["Comments Given"]), 0),
      };
    });

    const repSummary = userUsage.map((row) => {
      const name = getValue(row, ["Name"]);
      const email = getValue(row, ["Email"]);
      const teams = teamNamesFromRow(row);
      const appointments = numberValue(getValue(row, ["Total Appointments"]), 0);
      const recordedAppointments = numberValue(getValue(row, ["Appointments Recorded"]), 0);
      const conversationsRecorded = numberValue(getValue(row, ["Conversations Recorded"]), 0);
      const conversationsViewedRep = numberValue(getValue(row, ["Conversations Viewed"]), 0);
      const commentsReceivedRep = numberValue(getValue(row, ["Comments Received"]), 0);
      const compliance = appointments ? recordedAppointments / appointments : percentValue(getValue(row, ["Recording Compliance"]), null);
      const script = percentValue(getValue(row, ["Average Script Compliance"]), null);
      const talkRatio = percentValue(getValue(row, ["Talk Ratio Average"]), null);
      const viewedRatio = conversationsRecorded ? conversationsViewedRep / conversationsRecorded : percentValue(getValue(row, ["Viewed Recorded Ratio"]), null);
      const isRemoved = boolValue(getValue(row, ["Is Removed"]));
      const reasons = [];
      let risk = 0;
      if (appointments > 0 && recordedAppointments === 0) { risk += 4; reasons.push("No recorded appointments"); }
      if (compliance !== null && compliance < goals.recordingCompliance) { risk += compliance < goals.recordingCompliance * 0.65 ? 3 : 2; reasons.push("Low recording compliance"); }
      if (script !== null && script < goals.scriptCompliance) { risk += script < goals.scriptCompliance * 0.75 ? 2 : 1; reasons.push("Low script compliance"); }
      if (talkRatio !== null && talkRatio > goals.talkRatioMax) { risk += talkRatio > goals.talkRatioMax * 1.15 ? 2 : 1; reasons.push("High rep talk ratio"); }
      if (conversationsRecorded > 0 && commentsReceivedRep === 0) { risk += 1; reasons.push("No comments received"); }
      if (conversationsRecorded > 0 && viewedRatio !== null && viewedRatio < goals.minimumViewsRatio) { risk += 1; reasons.push("Low review activity"); }
      return {
        id: email || name,
        name,
        email,
        teams,
        teamLabel: teams.join(", ") || "Unassigned",
        isRemoved,
        appointments,
        recordedAppointments,
        conversationsRecorded,
        recordingCompliance: compliance,
        scriptCompliance: script,
        talkRatio,
        conversationsViewed: conversationsViewedRep,
        commentsReceived: commentsReceivedRep,
        commentsGiven: numberValue(getValue(row, ["Comments Given"]), 0),
        viewedRatio,
        risk,
        reasons: reasons.join("; ") || "No immediate flags",
      };
    });

    const teamOptions = ["All Teams", ...Array.from(new Set([...teamSummary.map((t) => t.name), ...repSummary.flatMap((r) => r.teams)].filter(Boolean))).sort()];
    const repOptions = ["All Reps", ...repSummary.map((r) => r.name).filter(Boolean).sort()];
    return { exportTotals, conversations, teamSummary, repSummary, teamOptions, repOptions, sheetCounts: { userManagement: userManagement.length, userUsage: userUsage.length, teamUsage: teamUsage.length, conversations: conversationsRaw.length } };
  }, [raw, goals]);

  const filtered = useMemo(() => {
    if (!model) return null;
    const { start, end } = getDateWindow(dateMode, model.exportTotals.dateEnd, customStart, customEnd);
    const teamPass = (teams) => selectedTeam === "All Teams" || teams.includes(selectedTeam);
    const repPass = (name) => selectedRep === "All Reps" || name === selectedRep;
    const conversations = model.conversations.filter((conversation) => teamPass(conversation.teams) && repPass(conversation.repName) && (!start || conversation.date >= start) && (!end || conversation.date <= end));
    const repMetadata = new Map(model.repSummary.map((row) => [row.name, row]));
    const repMap = new Map();
    const teamMap = new Map();

    for (const c of conversations) {
      const repName = c.repName || "Unknown";
      const teamLabel = c.teams.join(", ") || "Unassigned";
      const checklist = Number.isFinite(c.checklistScore) ? c.checklistScore : null;
      const talk = Number.isFinite(c.talkRatio) ? c.talkRatio : null;
      const sold = numberValue(c.sold, 0);
      const rep = repMap.get(repName) || { id: repName, name: repName, teamLabel, teams: c.teams, conversationsRecorded: 0, soldCount: 0, totalSold: 0, talkSum: 0, talkCount: 0, scriptSum: 0, scriptCount: 0, durationTotal: 0 };
      rep.conversationsRecorded += 1;
      rep.soldCount += sold > 0 || /sold/i.test(String(c.outcome || "")) ? 1 : 0;
      rep.totalSold += sold;
      rep.durationTotal += numberValue(c.duration, 0);
      if (talk !== null) { rep.talkSum += talk; rep.talkCount += 1; }
      if (checklist !== null) { rep.scriptSum += checklist; rep.scriptCount += 1; }
      repMap.set(repName, rep);

      const teams = c.teams.length ? c.teams : ["Unassigned"];
      for (const teamName of teams) {
        const team = teamMap.get(teamName) || { id: teamName, name: teamName, repSet: new Set(), conversationsRecorded: 0, soldCount: 0, totalSold: 0, talkSum: 0, talkCount: 0, scriptSum: 0, scriptCount: 0, durationTotal: 0 };
        team.repSet.add(repName);
        team.conversationsRecorded += 1;
        team.soldCount += sold > 0 || /sold/i.test(String(c.outcome || "")) ? 1 : 0;
        team.totalSold += sold;
        team.durationTotal += numberValue(c.duration, 0);
        if (talk !== null) { team.talkSum += talk; team.talkCount += 1; }
        if (checklist !== null) { team.scriptSum += checklist; team.scriptCount += 1; }
        teamMap.set(teamName, team);
      }
    }

    const repRows = Array.from(repMap.values())
      .map((row) => {
        const meta = repMetadata.get(row.name);
        const scriptCompliance = row.scriptCount ? row.scriptSum / row.scriptCount : null;
        const talkRatio = row.talkCount ? row.talkSum / row.talkCount : null;
        const avgDuration = row.conversationsRecorded ? row.durationTotal / row.conversationsRecorded : null;
        const reasons = [];
        let risk = 0;
        if (row.conversationsRecorded >= 3 && scriptCompliance !== null && scriptCompliance < goals.scriptCompliance) { risk += scriptCompliance < goals.scriptCompliance * 0.75 ? 2 : 1; reasons.push("Low script compliance"); }
        if (row.conversationsRecorded >= 3 && talkRatio !== null && talkRatio > goals.talkRatioMax) { risk += talkRatio > goals.talkRatioMax * 1.15 ? 2 : 1; reasons.push("High rep talk ratio"); }
        if (row.conversationsRecorded >= 5 && row.soldCount === 0) { risk += 1; reasons.push("No sold outcomes in period"); }
        if (row.conversationsRecorded < 3) { risk += 1; reasons.push("Low conversation volume"); }
        return {
          id: row.id,
          name: row.name,
          email: meta?.email || "",
          teamLabel: row.teamLabel,
          teams: row.teams,
          isRemoved: meta?.isRemoved || false,
          conversationsRecorded: row.conversationsRecorded,
          soldCount: row.soldCount,
          totalSold: row.totalSold,
          recordingCompliance: null,
          scriptCompliance,
          talkRatio,
          conversationsViewed: null,
          commentsReceived: null,
          avgDuration,
          risk,
          reasons: reasons.join("; ") || "No immediate flags",
        };
      })
      .filter((row) => !search.trim() || `${row.name} ${row.email} ${row.teamLabel}`.toLowerCase().includes(search.toLowerCase()))
      .sort((a, b) => b.conversationsRecorded - a.conversationsRecorded);

    const teamRows = Array.from(teamMap.values())
      .map((row) => ({
        id: row.id,
        name: row.name,
        usersWhoRecorded: row.repSet.size,
        totalUsers: row.repSet.size,
        conversationsRecorded: row.conversationsRecorded,
        soldCount: row.soldCount,
        totalSold: row.totalSold,
        recordingCompliance: null,
        scriptCompliance: row.scriptCount ? row.scriptSum / row.scriptCount : null,
        talkRatio: row.talkCount ? row.talkSum / row.talkCount : null,
        avgDuration: row.conversationsRecorded ? row.durationTotal / row.conversationsRecorded : null,
      }))
      .sort((a, b) => b.conversationsRecorded - a.conversationsRecorded);

    return { conversations, repRows, teamRows, start, end };
  }, [model, selectedTeam, selectedRep, dateMode, customStart, customEnd, search, goals.scriptCompliance, goals.talkRatioMax]);

  const selectedStats = useMemo(() => {
    if (!filtered) return null;
    const conversations = filtered.conversations;
    const soldRows = conversations.filter((c) => c.sold > 0 || /sold/i.test(String(c.outcome || "")));
    const talkValues = conversations.map((c) => c.talkRatio).filter((x) => Number.isFinite(x));
    const checklistValues = conversations.map((c) => c.checklistScore).filter((x) => Number.isFinite(x));
    const activeRepCount = new Set(conversations.map((c) => c.repName).filter(Boolean)).size;
    const avgTalk = talkValues.length ? talkValues.reduce((a, b) => a + b, 0) / talkValues.length : null;
    const avgScriptCompliance = checklistValues.length ? checklistValues.reduce((a, b) => a + b, 0) / checklistValues.length : null;
    const avgDuration = conversations.length ? conversations.reduce((sumValue, c) => sumValue + c.duration, 0) / conversations.length : null;
    const totalSold = conversations.reduce((sumValue, c) => sumValue + c.sold, 0);
    const soldRate = conversations.length ? soldRows.length / conversations.length : null;

    const groupedMap = new Map();
    for (const c of conversations) {
      let bucketDate = c.date;
      if (granularity === "week") bucketDate = getWeekStart(c.date);
      if (granularity === "month") bucketDate = new Date(c.date.getFullYear(), c.date.getMonth(), 1);
      const key = isoDate(bucketDate);
      const existing = groupedMap.get(key) || { key, label: granularity === "day" ? shortDate(bucketDate) : granularity === "month" ? bucketDate.toLocaleDateString(undefined, { month: "short" }) : `Wk ${shortDate(bucketDate)}`, conversations: 0, sold: 0, totalDuration: 0, talkSum: 0, talkCount: 0, scriptSum: 0, scriptCount: 0 };
      existing.conversations += 1;
      existing.sold += c.sold || 0;
      existing.totalDuration += c.duration || 0;
      if (c.talkRatio !== null) { existing.talkSum += c.talkRatio; existing.talkCount += 1; }
      if (Number.isFinite(c.checklistScore)) { existing.scriptSum += c.checklistScore; existing.scriptCount += 1; }
      groupedMap.set(key, existing);
    }
    const trend = Array.from(groupedMap.values()).sort((a, b) => new Date(a.key) - new Date(b.key)).map((item) => ({ ...item, avgTalkRatio: item.talkCount ? item.talkSum / item.talkCount : null, avgScriptCompliance: item.scriptCount ? item.scriptSum / item.scriptCount : null, avgDurationMinutes: item.conversations ? Math.round(item.totalDuration / item.conversations / 60) : 0 }));
    const outcomeMap = new Map();
    conversations.forEach((c) => outcomeMap.set(c.outcome, (outcomeMap.get(c.outcome) || 0) + 1));
    const outcomes = Array.from(outcomeMap.entries()).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
    const repVolumeMap = new Map();
    conversations.forEach((c) => repVolumeMap.set(c.repName, (repVolumeMap.get(c.repName) || 0) + 1));
    const repVolume = Array.from(repVolumeMap.entries()).map(([name, conversations]) => ({ name, conversations })).sort((a, b) => b.conversations - a.conversations).slice(0, 12);
    return { conversations: conversations.length, soldCount: soldRows.length, soldRate, totalSold, activeRepCount, avgTalk, avgScriptCompliance, avgDuration, trend, outcomes, repVolume };
  }, [filtered, granularity]);

  const scriptComplianceReport = useMemo(() => {
    if (!filtered) return null;
    const scriptMap = new Map();
    const trackerMap = new Map();
    const repMap = new Map();
    const lowScoreRows = [];
    let totalInstances = 0;
    let scoreSum = 0;
    let scoredInstances = 0;
    let belowGoal = 0;
    let totalHits = 0;
    let totalOpportunities = 0;

    for (const conversation of filtered.conversations) {
      for (const checklist of conversation.checklists || []) {
        if (checklist.pct === null) continue;
        totalInstances += 1;
        scoreSum += checklist.pct;
        scoredInstances += 1;
        if (checklist.pct < goals.scriptCompliance) belowGoal += 1;
        if (Number.isFinite(checklist.score)) totalHits += checklist.score;
        if (Number.isFinite(checklist.denominator)) totalOpportunities += checklist.denominator;

        const script = scriptMap.get(checklist.name) || { id: checklist.name, name: checklist.name, instances: 0, scoreSum: 0, belowGoal: 0, hits: 0, opportunities: 0, misses: new Map() };
        script.instances += 1;
        script.scoreSum += checklist.pct;
        script.belowGoal += checklist.pct < goals.scriptCompliance ? 1 : 0;
        script.hits += Number.isFinite(checklist.score) ? checklist.score : 0;
        script.opportunities += Number.isFinite(checklist.denominator) ? checklist.denominator : 0;
        checklist.missedTrackers.forEach((tracker) => script.misses.set(tracker?.tracker_name || "Unnamed Step", (script.misses.get(tracker?.tracker_name || "Unnamed Step") || 0) + 1));
        scriptMap.set(checklist.name, script);

        const rep = repMap.get(conversation.repName) || { id: conversation.repName, name: conversation.repName, teamLabel: conversation.teams.join(", ") || "Unassigned", scripts: 0, conversations: new Set(), scoreSum: 0, belowGoal: 0, hits: 0, opportunities: 0, misses: new Map() };
        rep.scripts += 1;
        rep.conversations.add(conversation.row["Conversation Id"] || `${conversation.repName}-${conversation.date?.toISOString?.()}`);
        rep.scoreSum += checklist.pct;
        rep.belowGoal += checklist.pct < goals.scriptCompliance ? 1 : 0;
        rep.hits += Number.isFinite(checklist.score) ? checklist.score : 0;
        rep.opportunities += Number.isFinite(checklist.denominator) ? checklist.denominator : 0;
        checklist.missedTrackers.forEach((tracker) => rep.misses.set(tracker?.tracker_name || "Unnamed Step", (rep.misses.get(tracker?.tracker_name || "Unnamed Step") || 0) + 1));
        repMap.set(conversation.repName, rep);

        checklist.trackers.forEach((tracker) => {
          const name = tracker?.tracker_name || "Unnamed Step";
          const trackerRow = trackerMap.get(name) || { id: name, name, hits: 0, opportunities: 0, aiScoreSum: 0, aiScoreCount: 0 };
          trackerRow.opportunities += 1;
          if (boolValue(tracker?.tracker_is_hit)) trackerRow.hits += 1;
          const aiScore = numberValue(tracker?.tracker_ai_score, NaN);
          if (Number.isFinite(aiScore)) { trackerRow.aiScoreSum += aiScore; trackerRow.aiScoreCount += 1; }
          trackerMap.set(name, trackerRow);
        });

        if (checklist.pct > 0 && checklist.pct < goals.scriptCompliance) {
          lowScoreRows.push({
            id: `${conversation.row["Conversation Id"] || conversation.repName}-${checklist.name}`,
            date: conversation.date,
            repName: conversation.repName,
            scriptName: checklist.name,
            score: checklist.pct,
            missed: checklist.missedTrackers.map((tracker) => tracker?.tracker_name).filter(Boolean).slice(0, 5).join(", ") || "No missed steps listed",
            url: conversation.rillaUrl,
          });
        }
      }
    }

    const formatMisses = (misses) => Array.from(misses.entries()).sort((a, b) => b[1] - a[1]).slice(0, 3).map(([name, count]) => `${name} (${count})`).join(", ") || "No repeated misses";
    const scriptRows = Array.from(scriptMap.values()).map((row) => ({
      ...row,
      avgScore: row.instances ? row.scoreSum / row.instances : null,
      goalHitRate: row.instances ? (row.instances - row.belowGoal) / row.instances : null,
      missedSteps: formatMisses(row.misses),
    })).sort((a, b) => a.avgScore - b.avgScore || b.instances - a.instances);
    const repRows = Array.from(repMap.values()).map((row) => ({
      ...row,
      conversationCount: row.conversations.size,
      avgScore: row.scripts ? row.scoreSum / row.scripts : null,
      missedSteps: formatMisses(row.misses),
    })).sort((a, b) => a.avgScore - b.avgScore || b.scripts - a.scripts);
    const trackerRows = Array.from(trackerMap.values()).map((row) => ({
      ...row,
      hitRate: row.opportunities ? row.hits / row.opportunities : null,
      avgAiScore: row.aiScoreCount ? row.aiScoreSum / row.aiScoreCount : null,
    })).sort((a, b) => a.hitRate - b.hitRate || b.opportunities - a.opportunities);

    return {
      totalInstances,
      avgScore: scoredInstances ? scoreSum / scoredInstances : null,
      weightedScore: totalOpportunities ? totalHits / totalOpportunities : null,
      belowGoal,
      goalHitRate: totalInstances ? (totalInstances - belowGoal) / totalInstances : null,
      scriptRows,
      repRows,
      trackerRows,
      lowScoreRows: lowScoreRows.sort((a, b) => a.score - b.score || b.date - a.date).slice(0, 25),
    };
  }, [filtered, goals.scriptCompliance]);

  const coachingPriorities = useMemo(() => {
    if (!filtered) return [];
    return [...filtered.repRows].filter((row) => !row.isRemoved && (row.risk > 0 || row.conversationsRecorded > 0)).sort((a, b) => b.risk - a.risk || b.conversationsRecorded - a.conversationsRecorded).slice(0, 15);
  }, [filtered]);

  const t = model?.exportTotals;
  const periodLabel = filtered ? dateRangeLabel(dateMode, filtered.start, filtered.end) : "All conversation rows";
  const displayedStartValue = dateMode === "custom" ? customStart : isoDate(filtered?.start);
  const displayedEndValue = dateMode === "custom" ? customEnd : isoDate(filtered?.end);
  const reconciliationDelta = t ? t.recordedConversations - t.detailConversationRows : 0;
  const chartGrid = isDark ? "#1e293b" : "#e2e8f0";
  const chartAxis = isDark ? "#94a3b8" : "#475569";
  const tooltipBg = isDark ? "#020617" : "#ffffff";
  const tooltipBorder = isDark ? "#334155" : "#cbd5e1";
  const tooltipLabel = isDark ? "#e2e8f0" : "#0f172a";

  const teamColumns = [
    { key: "name", label: "Team" },
    { key: "usersWhoRecorded", label: "Active Reps", render: (r) => fmtInt(r.usersWhoRecorded) },
    { key: "conversationsRecorded", label: "Conversations", render: (r) => fmtInt(r.conversationsRecorded) },
    { key: "soldCount", label: "Sold Count", render: (r) => fmtInt(r.soldCount) },
    { key: "totalSold", label: "Net Volume", render: (r) => {
      const goal = teamVolumeGoals[r.name];
      const pct = goal ? r.totalSold / goal : null;
      const color = pct === null ? "text-slate-300" : pct >= 1 ? "text-emerald-300" : pct >= 0.75 ? "text-amber-300" : "text-red-300";
      return <span className={color}>{fmtMoney(r.totalSold)}{goal ? <span className="ml-1 text-xs text-slate-500">/ {fmtMoney(goal)}</span> : null}</span>;
    }},
    { key: "scriptCompliance", label: "Script", render: (r) => fmtPct(r.scriptCompliance) },
    { key: "talkRatio", label: "Talk Ratio", render: (r) => fmtPct(r.talkRatio) },
    { key: "avgDuration", label: "Avg Duration", render: (r) => fmtDuration(r.avgDuration) },
  ];
  const repColumns = [
    { key: "name", label: "Rep" },
    { key: "teamLabel", label: "Team" },
    { key: "conversationsRecorded", label: "Convos", render: (r) => fmtInt(r.conversationsRecorded) },
    { key: "soldCount", label: "Sold", render: (r) => fmtInt(r.soldCount) },
    { key: "scriptCompliance", label: "Script", render: (r) => fmtPct(r.scriptCompliance) },
    { key: "talkRatio", label: "Talk Ratio", render: (r) => fmtPct(r.talkRatio) },
    { key: "avgDuration", label: "Avg Duration", render: (r) => fmtDuration(r.avgDuration) },
    { key: "risk", label: "Risk", render: (r) => <span className={r.risk >= 5 ? "text-red-300" : r.risk >= 3 ? "text-amber-300" : "text-slate-300"}>{fmtInt(r.risk)}</span> },
  ];

  const sortedRepRows = useMemo(() => {
    if (!filtered) return [];
    const rows = [...filtered.repRows];
    const key = repSortKey;
    const direction = repSortDirection === "asc" ? 1 : -1;
    rows.sort((a, b) => {
      const av = a?.[key];
      const bv = b?.[key];
      const aMissing = av === null || av === undefined;
      const bMissing = bv === null || bv === undefined;
      if (aMissing && bMissing) return 0;
      if (aMissing) return 1;
      if (bMissing) return -1;
      if (typeof av === "number" && typeof bv === "number") return (av - bv) * direction;
      return String(av).localeCompare(String(bv)) * direction;
    });
    return rows;
  }, [filtered, repSortKey, repSortDirection]);

  function handleRepSort(nextKey) {
    if (repSortKey === nextKey) {
      setRepSortDirection((prev) => prev === "asc" ? "desc" : "asc");
      return;
    }
    setRepSortKey(nextKey);
    setRepSortDirection("desc");
  }
  const scriptRepColumns = [
    { key: "name", label: "Rep" },
    { key: "teamLabel", label: "Team" },
    { key: "conversationCount", label: "Calls", render: (r) => fmtInt(r.conversationCount) },
    { key: "scripts", label: "Script Checks", render: (r) => fmtInt(r.scripts) },
    { key: "avgScore", label: "Avg Script", render: (r) => <span className={metricStatus(r.avgScore, goals.scriptCompliance) === "good" ? "text-emerald-300" : "text-amber-300"}>{fmtPct(r.avgScore)}</span> },
    { key: "belowGoal", label: "Below Goal", render: (r) => fmtInt(r.belowGoal) },
    { key: "missedSteps", label: "Most Missed Steps" },
  ];
  const scriptBreakdownColumns = [
    { key: "name", label: "Script" },
    { key: "instances", label: "Checks", render: (r) => fmtInt(r.instances) },
    { key: "avgScore", label: "Avg Score", render: (r) => fmtPct(r.avgScore) },
    { key: "goalHitRate", label: "Goal Hit Rate", render: (r) => fmtPct(r.goalHitRate) },
    { key: "missedSteps", label: "Most Missed Steps" },
  ];
  const trackerColumns = [
    { key: "name", label: "Script Step" },
    { key: "hitRate", label: "Hit Rate", render: (r) => fmtPct(r.hitRate) },
    { key: "hits", label: "Hits / Opps", render: (r) => `${fmtInt(r.hits)} / ${fmtInt(r.opportunities)}` },
    { key: "avgAiScore", label: "Avg AI Score", render: (r) => r.avgAiScore === null ? "—" : r.avgAiScore.toFixed(1) },
  ];
  const lowScoreColumns = [
    { key: "date", label: "Date", render: (r) => shortDate(r.date) },
    { key: "repName", label: "Rep" },
    { key: "scriptName", label: "Script" },
    { key: "score", label: "Score", render: (r) => fmtPct(r.score) },
    { key: "missed", label: "Missed Steps" },
    { key: "url", label: "Recording", render: (r) => r.url ? <a href={r.url} target="_blank" rel="noreferrer" className="text-blue-300 hover:text-blue-200">Open</a> : "—" },
  ];

  if (!raw || !model || !filtered || !selectedStats || !t) {
    return (
      <div data-theme={isDark ? "dark" : "light"} className={cx("rilla-app min-h-screen px-6 py-8", isDark ? "bg-slate-950 text-slate-100" : "bg-slate-50 text-slate-900")}>
        <ThemeOverrideStyle />
        <div className="pointer-events-none fixed inset-0 bg-[radial-gradient(circle_at_top_left,_rgba(37,99,235,0.18),_transparent_35%),radial-gradient(circle_at_bottom_right,_rgba(22,163,74,0.12),_transparent_35%)]" />
        <div className="relative mx-auto max-w-7xl">
          <div className="mb-8 flex items-center justify-between gap-3">
            <div className="flex items-center gap-3"><div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-blue-600 text-white"><BarChart3 className="h-6 w-6" /></div><div><h1 className="text-2xl font-bold tracking-tight text-white">Rilla Usage Analytics</h1><p className="text-sm text-slate-400">Interactive dashboard generated from raw Rilla export data.</p></div></div>
            <button onClick={() => setIsDark((prev) => !prev)} className="inline-flex items-center gap-2 rounded-2xl border border-slate-700 bg-slate-900 px-4 py-2 text-sm font-medium text-slate-100 hover:bg-slate-800">{isDark ? <Sun className="h-4 w-4" /> : <Moon className="h-4 w-4" />}{isDark ? "Light" : "Dark"} Mode</button>
          </div>
          {isLoadingDemo ? (
            <div className="mx-auto max-w-4xl rounded-3xl border border-slate-800 bg-slate-950/80 p-10 text-center shadow-2xl shadow-black/20">
              <div className="mx-auto flex h-16 w-16 items-center justify-center rounded-2xl bg-blue-500/10 text-blue-300"><BarChart3 className="h-8 w-8 animate-pulse" /></div>
              <h2 className="mt-6 text-2xl font-semibold text-white">Loading demo Rilla workbook</h2>
              <p className="mx-auto mt-3 max-w-2xl text-sm leading-6 text-slate-400">Opening with {DEMO_WORKBOOK_NAME} so the dashboard is populated immediately.</p>
            </div>
          ) : (
            <EmptyState onFile={handleFile} />
          )}
          {error && <div className="mx-auto mt-4 max-w-4xl rounded-2xl border border-red-500/40 bg-red-500/10 p-4 text-sm text-red-100">{error}</div>}
        </div>
      </div>
    );
  }

  return (
    <div data-theme={isDark ? "dark" : "light"} className={cx("rilla-app min-h-screen px-4 py-5 md:px-6 lg:px-8", isDark ? "bg-slate-950 text-slate-100" : "bg-slate-50 text-slate-900")}>
      <ThemeOverrideStyle />
      <div className="pointer-events-none fixed inset-0 bg-[radial-gradient(circle_at_top_left,_rgba(37,99,235,0.20),_transparent_36%),radial-gradient(circle_at_bottom_right,_rgba(16,185,129,0.10),_transparent_34%)]" />
      <div className="relative mx-auto max-w-[1800px] space-y-5">
        <header className="z-30 rounded-3xl border border-slate-800 bg-slate-950/90 p-4 shadow-2xl shadow-black/30 backdrop-blur lg:sticky lg:top-3">
          <div className="flex flex-col gap-4 xl:flex-row xl:items-center xl:justify-between">
            <div className="flex items-center gap-3"><div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-blue-600 text-white"><BarChart3 className="h-6 w-6" /></div><div><h1 className="text-2xl font-bold tracking-tight text-white">Rilla Usage Analytics</h1><p className="text-sm text-slate-400">{fileName || "Raw Rilla workbook"} · Export dates {shortDate(t.dateStart)} – {shortDate(t.dateEnd)} · View: {periodLabel}</p></div></div>
            <div className="flex items-center gap-2">
              <div className="grid w-full grid-cols-2 gap-2 sm:grid-cols-3 xl:w-auto xl:grid-cols-none xl:flex xl:flex-wrap xl:justify-end">
                <a href="https://app.rillavoice.com/settings/export" target="_blank" rel="noreferrer" className={actionButtonClass}>Data Source</a>
                <label className={cx(actionButtonClass, "cursor-pointer")}><Upload className="h-4 w-4" />Upload New File<input type="file" accept=".xlsx,.xls" className="hidden" onChange={handleFile} /></label>

                <button onClick={() => setIsDark((prev) => !prev)} className={actionButtonClass}>{isDark ? <Sun className="h-4 w-4" /> : <Moon className="h-4 w-4" />}{isDark ? "Light" : "Dark"} Mode</button>
              </div>
              <button onClick={() => setHeaderCollapsed((prev) => !prev)} className={cx(actionButtonClass, "shrink-0")} title={headerCollapsed ? "Expand filters" : "Collapse filters"}><ChevronDown className={cx("h-4 w-4 transition-transform duration-200", headerCollapsed ? "" : "rotate-180")} /></button>
            </div>
          </div>
          {!headerCollapsed && <>
            <div className="mt-4 grid grid-cols-1 gap-3 md:grid-cols-2 xl:grid-cols-5">
              <select value={selectedTeam} onChange={(e) => setSelectedTeam(e.target.value)} className="rounded-2xl border border-slate-700 bg-slate-900 px-3 py-2 text-sm text-white outline-none focus:border-blue-500">{model.teamOptions.map((team) => <option key={team}>{team}</option>)}</select>
              <select value={selectedRep} onChange={(e) => setSelectedRep(e.target.value)} className="rounded-2xl border border-slate-700 bg-slate-900 px-3 py-2 text-sm text-white outline-none focus:border-blue-500">{model.repOptions.map((rep) => <option key={rep}>{rep}</option>)}</select>
              <select value={granularity} onChange={(e) => setGranularity(e.target.value)} className="rounded-2xl border border-slate-700 bg-slate-900 px-3 py-2 text-sm text-white outline-none focus:border-blue-500"><option value="day">Daily Trend</option><option value="week">Weekly Trend</option><option value="month">Monthly Trend</option></select>
              <div className="relative xl:col-span-2"><Search className="pointer-events-none absolute left-3 top-2.5 h-4 w-4 text-slate-500" /><input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search reps, teams, emails..." className="w-full rounded-2xl border border-slate-700 bg-slate-900 py-2 pl-9 pr-3 text-sm text-white outline-none placeholder:text-slate-500 focus:border-blue-500" /></div>
            </div>
            <div className="mt-3 grid gap-3 xl:grid-cols-[1fr_370px]">
              <div className="flex flex-wrap gap-2 rounded-2xl border border-slate-800 bg-slate-900/70 p-2">{timeFrameOptions.map((option) => <button key={option.value} onClick={() => { setDateMode(option.value); setCustomStart(""); setCustomEnd(""); }} className={cx("rounded-xl px-3 py-2 text-xs font-semibold transition", dateMode === option.value ? "bg-blue-600 text-white" : "bg-slate-950 text-slate-300 hover:bg-slate-800")}>{option.label}</button>)}<button onClick={() => setDateMode("custom")} className={cx("rounded-xl px-3 py-2 text-xs font-semibold transition", dateMode === "custom" ? "bg-blue-600 text-white" : "bg-slate-950 text-slate-300 hover:bg-slate-800")}>Custom</button></div>
              <div className="grid grid-cols-2 gap-2 rounded-2xl border border-slate-800 bg-slate-900/70 p-2"><label className="flex items-center gap-2 rounded-xl bg-slate-950 px-3 py-2 text-xs text-slate-400"><CalendarDays className="h-4 w-4" /><input type="date" value={displayedStartValue} min={isoDate(t.dateStart)} max={isoDate(t.dateEnd)} onChange={(e) => { setDateMode("custom"); setCustomStart(e.target.value); }} className="w-full bg-transparent text-sm text-white outline-none" /></label><label className="flex items-center gap-2 rounded-xl bg-slate-950 px-3 py-2 text-xs text-slate-400"><CalendarDays className="h-4 w-4" /><input type="date" value={displayedEndValue} min={isoDate(t.dateStart)} max={isoDate(t.dateEnd)} onChange={(e) => { setDateMode("custom"); setCustomEnd(e.target.value); }} className="w-full bg-transparent text-sm text-white outline-none" /></label></div>
            </div>
          </>}
        </header>

        <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 xl:grid-cols-5">
          <KpiCard icon={ClipboardCheck} label="Period Recordings" value={fmtInt(selectedStats.conversations)} sublabel={periodLabel} />
          <KpiCard icon={Users} label="Active Reps" value={fmtInt(selectedStats.activeRepCount)} sublabel="With calls in period" />
          <KpiCard icon={MessageSquare} label="Script Compliance" value={fmtPct(selectedStats.avgScriptCompliance)} sublabel={`Goal ${fmtPct(goals.scriptCompliance)}`} status={metricStatus(selectedStats.avgScriptCompliance, goals.scriptCompliance)} />
          <KpiCard icon={Timer} label="Avg Talk Ratio" value={fmtPct(selectedStats.avgTalk)} sublabel={`Max ${fmtPct(goals.talkRatioMax)}`} status={metricStatus(selectedStats.avgTalk, goals.talkRatioMax, "below")} />
          <KpiCard icon={Eye} label="Avg Duration" value={fmtDuration(selectedStats.avgDuration)} sublabel="Per selected conversation" />
        </div>

        {showInfoPanel && (
          <Section
            title="KPI Info"
            subtitle="How to read the top cards and coaching risk score."
            right={<button onClick={() => setShowInfoPanel(false)} className="rounded-xl border border-slate-700 bg-slate-900 p-2 text-slate-300 hover:bg-slate-800"><X className="h-4 w-4" /></button>}
          >
            <div className="mb-4 flex flex-wrap gap-2 rounded-2xl border border-slate-800 bg-slate-950/70 p-2">
              <InfoPill active={infoTab === "kpis"} onClick={() => setInfoTab("kpis")}>KPI Definitions</InfoPill>
              <InfoPill active={infoTab === "risk"} onClick={() => setInfoTab("risk")}>Risk Score</InfoPill>
              <InfoPill active={infoTab === "filters"} onClick={() => setInfoTab("filters")}>Filters & Data</InfoPill>
            </div>

            {infoTab === "kpis" && (
              <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-4">
                <MetricDefinition label="Period Recordings">Count of dated conversation rows after the current date, team, and rep filters.</MetricDefinition>
                <MetricDefinition label="Sold Count">Conversations marked as sold because Total Sold is above zero or the outcome text includes sold.</MetricDefinition>
                <MetricDefinition label="Total Sold">Sum of Total Sold from the selected conversation rows. Blank values count as zero.</MetricDefinition>
                <MetricDefinition label="Active Reps">Unique reps with at least one conversation in the selected period and filters.</MetricDefinition>
                <MetricDefinition label="Script Compliance">Average checklist completion from conversation checklists, calculated as checklist score divided by checklist denominator.</MetricDefinition>
                <MetricDefinition label="Avg Talk Ratio">Average rep talk ratio for the selected conversations. Lower is better against the Talk Ratio Max goal.</MetricDefinition>
                <MetricDefinition label="Avg Duration">Average conversation duration for selected conversations, converted from seconds into minutes or hours.</MetricDefinition>
              </div>
            )}

            {infoTab === "risk" && (
              <div className="grid gap-4 xl:grid-cols-[1fr_1fr]">
                <div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-5">
                  <h3 className="text-base font-semibold text-white">How the risk score is built</h3>
                  <p className="mt-2 text-sm leading-6 text-slate-300">Risk is a coaching-priority score from export-level rep usage data. Higher scores mean a rep has more coaching signals to review. It is not a performance grade; it is a triage helper.</p>
                  <div className="mt-4 grid gap-2 text-sm text-slate-300">
                    <div className="flex justify-between gap-4 rounded-xl bg-slate-950/70 px-3 py-2"><span>No recorded appointments despite appointments</span><span className="font-semibold text-red-300">+4</span></div>
                    <div className="flex justify-between gap-4 rounded-xl bg-slate-950/70 px-3 py-2"><span>Recording compliance below goal</span><span className="font-semibold text-amber-300">+2 to +3</span></div>
                    <div className="flex justify-between gap-4 rounded-xl bg-slate-950/70 px-3 py-2"><span>Script compliance below goal</span><span className="font-semibold text-amber-300">+1 to +2</span></div>
                    <div className="flex justify-between gap-4 rounded-xl bg-slate-950/70 px-3 py-2"><span>Talk ratio above max goal</span><span className="font-semibold text-amber-300">+1 to +2</span></div>
                    <div className="flex justify-between gap-4 rounded-xl bg-slate-950/70 px-3 py-2"><span>No comments received on recorded conversations</span><span className="font-semibold text-slate-200">+1</span></div>
                    <div className="flex justify-between gap-4 rounded-xl bg-slate-950/70 px-3 py-2"><span>Low conversation review activity</span><span className="font-semibold text-slate-200">+1</span></div>
                  </div>
                </div>
                <div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-5">
                  <h3 className="text-base font-semibold text-white">How to interpret it</h3>
                  <div className="mt-4 space-y-3 text-sm leading-6 text-slate-300">
                    <p><span className="font-semibold text-slate-100">0 to 2:</span> Low immediate signal. The rep may still be worth reviewing, but the export does not show many automatic flags.</p>
                    <p><span className="font-semibold text-amber-200">3 to 4:</span> Moderate coaching signal. Review the flags column to see whether compliance, script, talk ratio, or engagement is driving it.</p>
                    <p><span className="font-semibold text-red-200">5+:</span> High-priority coaching queue. These reps have multiple flags or a severe recording-compliance issue.</p>
                    <p className="rounded-2xl border border-blue-500/30 bg-blue-500/10 p-3 text-blue-100">The goal settings panel changes the thresholds used for risk scoring and KPI colors.</p>
                  </div>
                </div>
              </div>
            )}

            {infoTab === "filters" && (
              <div className="grid gap-3 md:grid-cols-3">
                <MetricDefinition label="Timeframe">The top KPI cards, trend chart, selected period detail, and rep volume chart use the selected date range from the conversation sheet.</MetricDefinition>
                <MetricDefinition label="Team & Rep">Team and rep filters narrow all cards and tables on this page using conversation-level rows.</MetricDefinition>
                <MetricDefinition label="Timeframe Coverage">Team Summary, Rep Summary, and Coaching Priorities now recalculate from the selected date range, team, and rep filters.</MetricDefinition>
              </div>
            )}
          </Section>
        )}

        <div className="grid grid-cols-1 gap-5 xl:grid-cols-[1.65fr_1fr]">
          <Section title="Conversation Trend" subtitle={`${periodLabel}. Date filters use the conversation-detail sheet.`} right={<div className="text-sm text-slate-400">{fmtInt(selectedStats.conversations)} conversations · {fmtMoney(selectedStats.totalSold)} sold</div>}>
            <div className="mb-3 flex flex-wrap gap-2">
              {[{ key: "conversations", label: "Conversations", color: "#10b981" }, { key: "scriptCompliance", label: "Script Compliance", color: "#2563eb" }, { key: "talkRatio", label: "Avg Talk Ratio", color: "#f59e0b" }].map(({ key, label, color }) => (
                <button key={key} onClick={() => setTrendLines((prev) => ({ ...prev, [key]: !prev[key] }))} className={cx("flex items-center gap-1.5 rounded-xl border px-3 py-1.5 text-xs font-semibold transition", trendLines[key] ? "border-transparent text-white" : "border-slate-700 bg-slate-900 text-slate-400")} style={trendLines[key] ? { backgroundColor: color } : {}}>
                  <span className="inline-block h-2 w-2 rounded-full" style={{ backgroundColor: trendLines[key] ? "white" : color }} />{label}
                </button>
              ))}
            </div>
            <div className="h-[300px]"><ResponsiveContainer width="100%" height="100%"><LineChart data={selectedStats.trend} margin={{ top: 8, right: 20, left: 0, bottom: 0 }}><CartesianGrid strokeDasharray="3 3" stroke={chartGrid} /><XAxis dataKey="label" stroke={chartAxis} tick={{ fontSize: 12 }} />{trendLines.conversations && <YAxis yAxisId="left" stroke={chartAxis} tick={{ fontSize: 12 }} />}{(trendLines.scriptCompliance || trendLines.talkRatio) && <YAxis yAxisId="right" orientation="right" tickFormatter={(v) => `${Math.round(v * 100)}%`} stroke={chartAxis} tick={{ fontSize: 12 }} />}<Tooltip contentStyle={{ background: tooltipBg, border: `1px solid ${tooltipBorder}`, borderRadius: 14 }} labelStyle={{ color: tooltipLabel }} formatter={(value, name) => (name === "Conversations" ? value?.toLocaleString() : fmtPct(value, 1))} />{trendLines.conversations && <Line yAxisId="left" type="monotone" dataKey="conversations" name="Conversations" stroke="#10b981" strokeWidth={3} dot={false} />}{trendLines.scriptCompliance && <Line yAxisId="right" type="monotone" dataKey="avgScriptCompliance" name="Script Compliance" stroke="#2563eb" strokeWidth={3} dot={false} />}{trendLines.talkRatio && <Line yAxisId="right" type="monotone" dataKey="avgTalkRatio" name="Avg Talk Ratio" stroke="#f59e0b" strokeWidth={3} dot={false} />}</LineChart></ResponsiveContainer></div>
          </Section>
          <Section title="Selected Period Detail" subtitle="Calculated directly from conversation rows only."><div className="grid grid-cols-2 gap-3"><div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-4"><div className="text-xs text-slate-400">Conversations</div><div className="mt-1 text-2xl font-semibold text-white">{fmtInt(selectedStats.conversations)}</div></div><div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-4"><div className="text-xs text-slate-400">Sold Count</div><div className="mt-1 text-2xl font-semibold text-white">{fmtInt(selectedStats.soldCount)}</div></div><div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-4"><div className="text-xs text-slate-400">Avg Talk</div><div className="mt-1 text-2xl font-semibold text-white">{fmtPct(selectedStats.avgTalk)}</div></div><div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-4"><div className="text-xs text-slate-400">Avg Duration</div><div className="mt-1 text-2xl font-semibold text-white">{fmtDuration(selectedStats.avgDuration)}</div></div></div><div className="mt-4 h-[220px]"><ResponsiveContainer width="100%" height="100%"><PieChart><Pie data={selectedStats.outcomes.slice(0, 7)} dataKey="value" nameKey="name" innerRadius={55} outerRadius={90} paddingAngle={3}>{selectedStats.outcomes.slice(0, 7).map((entry, index) => <Cell key={entry.name} fill={COLORS[index % COLORS.length]} />)}</Pie><Tooltip contentStyle={{ background: tooltipBg, border: `1px solid ${tooltipBorder}`, borderRadius: 14 }} /><Legend wrapperStyle={{ fontSize: 12 }} /></PieChart></ResponsiveContainer></div></Section>
        </div>

        <Section title="Team Summary" subtitle="Conversation-level team metrics for the selected timeframe."><DataTable rows={filtered.teamRows} columns={teamColumns} maxHeight="max-h-[420px]" /></Section>
        <Section title="Conversation Volume by Rep" subtitle="Uses selected date/team/rep filters from conversation rows."><div className="h-[420px]"><ResponsiveContainer width="100%" height="100%"><BarChart data={selectedStats.repVolume} layout="vertical" margin={{ top: 5, right: 20, left: 40, bottom: 5 }}><CartesianGrid strokeDasharray="3 3" stroke={chartGrid} /><XAxis type="number" stroke={chartAxis} tick={{ fontSize: 12 }} /><YAxis type="category" dataKey="name" width={110} stroke={chartAxis} tick={{ fontSize: 12 }} /><Tooltip contentStyle={{ background: tooltipBg, border: `1px solid ${tooltipBorder}`, borderRadius: 14 }} /><Bar dataKey="conversations" name="Conversations" fill="#2563eb" radius={[0, 8, 8, 0]} /></BarChart></ResponsiveContainer></div></Section>
        <Section title="Coaching Priorities" subtitle="Ranked by script, talk-ratio, sold outcomes, and volume within the selected timeframe." right={<AlertTriangle className="h-5 w-5 text-amber-300" />}><DataTable rows={coachingPriorities} columns={[{ key: "name", label: "Rep" }, { key: "teamLabel", label: "Team" }, { key: "risk", label: "Risk", render: (r) => <span className={r.risk >= 5 ? "text-red-300" : r.risk >= 3 ? "text-amber-300" : "text-slate-300"}>{fmtInt(r.risk)}</span> }, { key: "conversationsRecorded", label: "Convos", render: (r) => fmtInt(r.conversationsRecorded) }, { key: "scriptCompliance", label: "Script", render: (r) => fmtPct(r.scriptCompliance) }, { key: "talkRatio", label: "Talk", render: (r) => fmtPct(r.talkRatio) }, { key: "reasons", label: "Flags" }]} maxHeight="max-h-[420px]" wrapCells /></Section>

        <div className="grid grid-cols-1 gap-5 2xl:grid-cols-[1.2fr_0.8fr]">
          <Section title="Rep Summary" subtitle="Conversation-level rep metrics in the selected timeframe. Click any column header to sort."><DataTable rows={sortedRepRows} columns={repColumns} maxHeight="max-h-[720px]" sortKey={repSortKey} sortDirection={repSortDirection} onSort={handleRepSort} /></Section>
          <Section title="Goal Settings" subtitle="Adjust thresholds used for KPI color and coaching risk."><div className="space-y-3">{[["recordingCompliance", "Recording Compliance Goal"], ["scriptCompliance", "Script Compliance Goal"], ["talkRatioMax", "Talk Ratio Max"], ["minimumViewsRatio", "Minimum Review Ratio"]].map(([key, label]) => <label key={key} className="grid grid-cols-[1fr_95px] items-center gap-3 text-sm text-slate-300"><span>{label}</span><input type="number" min="0" max="100" value={Math.round(goals[key] * 100)} onChange={(e) => setGoals((prev) => ({ ...prev, [key]: numberValue(e.target.value, 0) / 100 }))} className="rounded-xl border border-slate-700 bg-slate-900 px-3 py-2 text-right text-white outline-none focus:border-blue-500" /></label>)}<button onClick={() => setGoals(DEFAULT_GOALS)} className="inline-flex items-center gap-2 rounded-2xl border border-slate-700 bg-slate-900 px-4 py-2 text-sm font-medium text-slate-100 hover:bg-slate-800"><Settings className="h-4 w-4" /> Reset Goals</button></div></Section>
        </div>
        <Section
          title="Detailed Script Compliance"
          subtitle={`Conversation-level checklist reporting for ${periodLabel}. Team, rep, and date filters are applied.`}
          right={<div className="rounded-2xl border border-blue-500/30 bg-blue-500/10 px-4 py-2 text-sm font-semibold text-blue-100">Goal {fmtPct(goals.scriptCompliance)}</div>}
        >
          <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-3 2xl:grid-cols-5">
            <div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-4"><div className="text-xs uppercase tracking-[0.16em] text-slate-400">Avg Script</div><div className="mt-2 text-2xl font-semibold text-white">{fmtPct(scriptComplianceReport?.avgScore)}</div><div className="mt-1 text-xs text-slate-400">Average across checklist instances</div></div>
            <div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-4"><div className="text-xs uppercase tracking-[0.16em] text-slate-400">Weighted Score</div><div className="mt-2 text-2xl font-semibold text-white">{fmtPct(scriptComplianceReport?.weightedScore)}</div><div className="mt-1 text-xs text-slate-400">Hit steps divided by available steps</div></div>
            <div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-4"><div className="text-xs uppercase tracking-[0.16em] text-slate-400">Goal Hit Rate</div><div className="mt-2 text-2xl font-semibold text-white">{fmtPct(scriptComplianceReport?.goalHitRate)}</div><div className="mt-1 text-xs text-slate-400">Checklist scores at or above goal</div></div>
            <div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-4"><div className="text-xs uppercase tracking-[0.16em] text-slate-400">Below Goal</div><div className="mt-2 text-2xl font-semibold text-white">{fmtInt(scriptComplianceReport?.belowGoal)}</div><div className="mt-1 text-xs text-slate-400">Script checks needing review</div></div>
            <div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-4"><div className="text-xs uppercase tracking-[0.16em] text-slate-400">Script Checks</div><div className="mt-2 text-2xl font-semibold text-white">{fmtInt(scriptComplianceReport?.totalInstances)}</div><div className="mt-1 text-xs text-slate-400">Parsed from conversation checklists</div></div>
          </div>

          <div className="mt-5 grid grid-cols-1 gap-5 2xl:grid-cols-[1.1fr_0.9fr]">
            <div className="space-y-5">
              <div>
                <h3 className="mb-2 text-sm font-semibold uppercase tracking-[0.16em] text-slate-400">Rep Script Ranking</h3>
                <DataTable rows={scriptComplianceReport?.repRows || []} columns={scriptRepColumns} maxHeight="max-h-[360px]" wrapCells />
              </div>
              <div>
                <h3 className="mb-2 text-sm font-semibold uppercase tracking-[0.16em] text-slate-400">Lowest Scoring Conversations</h3>
                <DataTable rows={scriptComplianceReport?.lowScoreRows || []} columns={lowScoreColumns} maxHeight="max-h-[360px]" wrapCells />
              </div>
            </div>
            <div className="space-y-5">
              <div>
                <h3 className="mb-2 text-sm font-semibold uppercase tracking-[0.16em] text-slate-400">Script Breakdown</h3>
                <DataTable rows={scriptComplianceReport?.scriptRows || []} columns={scriptBreakdownColumns} maxHeight="max-h-[360px]" wrapCells />
              </div>
              <div>
                <h3 className="mb-2 text-sm font-semibold uppercase tracking-[0.16em] text-slate-400">Script Step Hit Rates</h3>
                <DataTable rows={(scriptComplianceReport?.trackerRows || []).slice(0, 20)} columns={trackerColumns} maxHeight="max-h-[360px]" wrapCells />
              </div>
            </div>
          </div>
        </Section>
      </div>
    </div>
  );
}
