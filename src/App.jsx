import React, { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import {
  Phone, Mail, CalendarCheck, TrendingUp, Trophy, Plus, Target,
  ChevronDown, ChevronLeft, ChevronRight, X, Users, BarChart3, Trash2, Shield, LogOut,
  UserPlus, Pencil, Eye, EyeOff, Briefcase, DollarSign, Kanban,
  Table2, ArrowRight, Building2, Percent, CheckCircle2, Download, FileSpreadsheet, UserCheck, Clock, AlertTriangle
} from "lucide-react";
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid,
  Tooltip, ResponsiveContainer, Legend
} from "recharts";
import * as api from "./api";
import { supabase } from "./supabaseClient";


// Tellemica brand palette — sampled from the portal design
const INK = "#0E1D32";      // near-black navy — card headers, chrome, primary text
const NAVY2 = "#102B5C";    // mid navy for gradients
const PAPER = "#F1F5F9";    // cool off-white app canvas
const CARD = "#FFFFFF";
const CYAN = "#0E9EE1";     // signature cyan accent (tabs, underlines, highlights)
const BTN_A = "#1854BB";    // button gradient start (blue)
const BTN_B = "#0F96DC";    // button gradient end (cyan)
const CALL = "#0B7285";     // teal — calls
const EMAIL = "#1854BB";    // portal blue — emails
const APPT = "#0E9EE1";     // cyan — appointments (the goal keeps the accent)
const LINE_C = "#DCE4EC";

const TODAY = () => new Date().toISOString().slice(0, 10);
// Format a phone as (111) 111-1111 as the user types, keeping only digits.
// Drops a leading US country code "1" if present, so 16105550100 -> (610) 555-0100.
const formatPhone = (val) => {
  let d = (val || "").replace(/\D/g, "");
  if (d.length === 11 && d[0] === "1") d = d.slice(1);
  d = d.slice(0, 10);
  if (d.length === 0) return "";
  if (d.length < 4) return `(${d}`;
  if (d.length < 7) return `(${d.slice(0, 3)}) ${d.slice(3)}`;
  return `(${d.slice(0, 3)}) ${d.slice(3, 6)}-${d.slice(6)}`;
};
// MM-DD-YYYY for the bulk-upload template (display format).
const TODAY_US = () => { const d = new Date(); return `${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}-${d.getFullYear()}`; };
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);

const ROLES = {
  admin: { label: "Admin", rank: 4, color: "#8E3B46" },
  management: { label: "Management", rank: 3, color: "#5B4B8A" },
  sales: { label: "Sales Rep", rank: 2, color: EMAIL },
  bdr: { label: "BDR", rank: 1, color: CALL },
};

const DEFAULT_GOALS = { calls: 1000, emails: 650, appts: 65 };

// ---- Month reporting helpers ----
// Portal starts fresh in June 2026; build the selectable month list from there to now.
const FIRST_MONTH = "2026-06"; // earliest month with data
const monthKey = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
const CURRENT_MONTH = monthKey(new Date());
// Returns list of "YYYY-MM" from FIRST_MONTH through the current month (inclusive), newest first.
const buildMonthOptions = () => {
  const out = [];
  const [fy, fm] = FIRST_MONTH.split("-").map(Number);
  const start = new Date(fy, fm - 1, 1);
  const now = new Date();
  const cur = new Date(now.getFullYear(), now.getMonth(), 1);
  let d = new Date(start);
  while (d <= cur) { out.push(monthKey(d)); d.setMonth(d.getMonth() + 1); }
  return out.reverse();
};
const monthLabel = (mk) => {
  const [y, m] = mk.split("-").map(Number);
  return new Date(y, m - 1, 1).toLocaleDateString("en-US", { month: "long", year: "numeric" });
};
// True if an entry/deal date (YYYY-MM-DD) falls in the given month key.
const inMonth = (dateStr, mk) => (dateStr || "").slice(0, 7) === mk;

// Pipeline stages: ordered flow with per-stage color and a default win-probability
// used for the weighted forecast. Closed Won/Lost are terminal.
const STAGES = [
  { id: "new", label: "New", color: "#64748B", prob: 10 },
  { id: "contacted", label: "Contacted", color: "#2563A8", prob: 25 },
  { id: "appointment", label: "Appointment Set", color: "#00A9E0", prob: 50 },
  { id: "proposal", label: "Proposal", color: "#0B7285", prob: 75 },
  { id: "won", label: "Closed Won", color: "#16A34A", prob: 100 },
  { id: "lost", label: "Closed Lost", color: "#B4453F", prob: 0 },
];
const STAGE = Object.fromEntries(STAGES.map((s) => [s.id, s]));
const OPEN_STAGES = STAGES.filter((s) => s.id !== "won" && s.id !== "lost");
// Reasons captured when a deal is marked Closed Lost — drives win/loss reporting.
const LOST_REASONS = ["Price / budget", "Went with competitor", "Timing / no decision", "No response / went dark", "Not a fit", "Other"];

// --- Follow-up & aging helpers ---
const DAY_MS = 86400000;
const startOfToday = () => { const d = new Date(); d.setHours(0, 0, 0, 0); return d; };
// Days between a yyyy-mm-dd string and today (negative = in the past / overdue).
const daysUntil = (dateStr) => { if (!dateStr) return null; const t = new Date(dateStr + "T00:00"); return Math.round((t - startOfToday()) / DAY_MS); };
const daysSince = (isoStr) => { if (!isoStr) return null; return Math.round((startOfToday() - new Date(isoStr)) / DAY_MS); };
// A follow-up is "due" if its date is today or earlier; label describes urgency.
const followUpState = (dateStr) => {
  const d = daysUntil(dateStr);
  if (d === null) return null;
  if (d < 0) return { label: `${Math.abs(d)}d overdue`, color: "#B4453F", due: true };
  if (d === 0) return { label: "Due today", color: "#C2410C", due: true };
  if (d <= 3) return { label: `Due in ${d}d`, color: "#B7791F", due: false };
  return { label: `Due ${new Date(dateStr + "T00:00").toLocaleDateString("en-US", { month: "short", day: "numeric" })}`, color: "#5A6B7B", due: false };
};
// Stale = open deal untouched (by created date proxy) longer than the stage threshold.
const STAGE_STALE_DAYS = { new: 14, contacted: 21, appointment: 30, proposal: 30 };
const dealAge = (deal) => {
  if (deal.stage === "won" || deal.stage === "lost") return null;
  // Prefer time in current stage; fall back to creation date for legacy deals.
  const age = daysSince(deal.stageChangedAt || deal.createdAt);
  if (age === null) return null;
  const threshold = STAGE_STALE_DAYS[deal.stage] || 30;
  return { age, stale: age >= threshold, threshold };
};

// --- Calendar invite (.ics) ---
// Builds and downloads a pre-filled calendar file for an appointment. Opens in
// Outlook / Google / Apple Calendar with the subject, agenda body, time, and
// location already set — the rep just adds the attendee and sends.
const APPT_TEMPLATE = {
  subject: "AT&T Account Review: with {company}",
  body: [
    "AT&T Account Review Meeting",
    "Date: {date}",
    "Time: {time}",
    "",
    "Agenda:",
    "- Introductions and business overview",
    "- Review of current AT&T services and account setup",
    "- Discussion of current usage, challenges, and business needs",
    "- Review of available plan updates and service enhancements",
    "- Identification of potential cost-saving opportunities",
    "- Questions and next steps",
  ].join("\n"),
};

const icsEscape = (s) => String(s || "").replace(/\\/g, "\\\\").replace(/;/g, "\\;").replace(/,/g, "\\,").replace(/\n/g, "\\n");
const icsStamp = (d) => d.toISOString().replace(/[-:]/g, "").split(".")[0] + "Z";
// Convert a stored ISO timestamp to the value a <input type="datetime-local"> expects (local time, no zone).
const toLocalInput = (iso) => { const d = new Date(iso); const off = d.getTimezoneOffset() * 60000; return new Date(d - off).toISOString().slice(0, 16); };

// US timezones reps pick from. `zone` is an IANA name used to compute the correct UTC instant.
// Convert stored "HH:MM" (24h) to { h12, min, ap } for the dropdowns, and back.
const parseTime12 = (hhmm) => {
  if (!hhmm) return { h12: "", min: "", ap: "" };
  const [h, m] = hhmm.split(":").map(Number);
  const ap = h >= 12 ? "PM" : "AM";
  let h12 = h % 12; if (h12 === 0) h12 = 12;
  return { h12: String(h12), min: String(m).padStart(2, "0"), ap };
};
const buildTime24 = ({ h12, min, ap }) => {
  if (!h12 || !min || !ap) return "";
  let h = Number(h12) % 12;
  if (ap === "PM") h += 12;
  return `${String(h).padStart(2, "0")}:${min}`;
};

const US_TIMEZONES = [
  { id: "America/New_York", label: "Eastern (ET)" },
  { id: "America/Chicago", label: "Central (CT)" },
  { id: "America/Denver", label: "Mountain (MT)" },
  { id: "America/Phoenix", label: "Arizona (no DST)" },
  { id: "America/Los_Angeles", label: "Pacific (PT)" },
  { id: "America/Anchorage", label: "Alaska (AKT)" },
  { id: "Pacific/Honolulu", label: "Hawaii (HT)" },
];
// Given a date ("YYYY-MM-DD"), a time ("HH:MM"), and an IANA zone, return the
// absolute UTC Date for that wall-clock time in that zone. Works by measuring the
// zone's offset at that moment and adjusting.
function zonedToUTC(dateStr, timeStr, zone) {
  if (!dateStr || !timeStr) return null;
  const [y, m, d] = dateStr.split("-").map(Number);
  const [hh, mm] = timeStr.split(":").map(Number);
  // Start with the naive UTC guess, then correct by the zone's offset at that time.
  const asUTC = Date.UTC(y, m - 1, d, hh, mm);
  const guess = new Date(asUTC);
  // What wall-clock time does `guess` show in the target zone?
  const fmt = new Intl.DateTimeFormat("en-US", { timeZone: zone, hour12: false, year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit" });
  const parts = Object.fromEntries(fmt.formatToParts(guess).map((p) => [p.type, p.value]));
  const shown = Date.UTC(+parts.year, +parts.month - 1, +parts.day, +parts.hour % 24, +parts.minute);
  // Difference between what we wanted and what the zone showed = the offset to remove.
  return new Date(asUTC - (shown - asUTC));
}

// Build an Outlook web deeplink that opens a pre-filled calendar event (no file).
function outlookDeeplink(deal, rep, manager, extraEmails, start) {
  const repName = rep?.name || "";
  const end = new Date(start.getTime() + 30 * 60000);
  const fill = (t) => t
    .replace(/{company}/g, deal.company || "")
    .replace(/{contact}/g, deal.contact || "")
    .replace(/{value}/g, deal.value ? fmtMoney(deal.value) : "")
    .replace(/{rep}/g, repName)
    .replace(/{date}/g, start.toLocaleDateString("en-US", { weekday: "long", month: "long", day: "numeric", year: "numeric" }))
    .replace(/{time}/g, start.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" }));
  const to = [deal.contactEmail, manager?.email, ...(extraEmails || "").split(/[,;\s]+/).map((s) => s.trim())]
    .filter((e) => e && e.includes("@"));
  const params = new URLSearchParams({
    path: "/calendar/action/compose",
    rru: "addevent",
    subject: fill(APPT_TEMPLATE.subject),
    body: fill(APPT_TEMPLATE.body),
    startdt: start.toISOString(),
    enddt: end.toISOString(),
    location: deal.company || "",
  });
  if (to.length) params.set("to", to.join(","));
  return `https://outlook.office.com/calendar/0/deeplink/compose?${params.toString()}`;
}

function downloadAppointmentICS(deal, rep, manager, extraEmails) {
  if (!deal.apptAt) return;
  const start = new Date(deal.apptAt);
  const end = new Date(start.getTime() + 30 * 60000); // 30-minute default
  const repName = rep?.name || "";
  const fill = (t) => t
    .replace(/{company}/g, deal.company || "")
    .replace(/{contact}/g, deal.contact || "")
    .replace(/{value}/g, deal.value ? fmtMoney(deal.value) : "")
    .replace(/{rep}/g, repName)
    .replace(/{date}/g, start.toLocaleDateString("en-US", { weekday: "long", month: "long", day: "numeric", year: "numeric" }))
    .replace(/{time}/g, start.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" }));

  const lines = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//Tellemica//Sales Command Center//EN",
    "METHOD:REQUEST",
    "BEGIN:VEVENT",
    `UID:${deal.id}-${start.getTime()}@tellemica`,
    `DTSTAMP:${icsStamp(new Date())}`,
    `DTSTART:${icsStamp(start)}`,
    `DTEND:${icsStamp(end)}`,
    `SUMMARY:${icsEscape(fill(APPT_TEMPLATE.subject))}`,
    `DESCRIPTION:${icsEscape(fill(APPT_TEMPLATE.body))}`,
    `LOCATION:${icsEscape(deal.company || "")}`,
  ];
  // Organizer = rep; required attendee = customer; optional = manager + extras.
  if (rep?.email) lines.push(`ORGANIZER;CN=${icsEscape(repName)}:mailto:${rep.email}`);
  if (deal.contactEmail) lines.push(`ATTENDEE;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE;CN=${icsEscape(deal.contact || deal.contactEmail)}:mailto:${deal.contactEmail}`);
  if (manager?.email) lines.push(`ATTENDEE;ROLE=OPT-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE;CN=${icsEscape(manager.name || manager.email)}:mailto:${manager.email}`);
  // Any additional invitees the rep listed (comma/semicolon/space separated).
  (extraEmails || "").split(/[,;\s]+/).map((s) => s.trim()).filter((s) => s.includes("@")).forEach((em) => {
    lines.push(`ATTENDEE;ROLE=OPT-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE;CN=${icsEscape(em)}:mailto:${em}`);
  });
  lines.push("END:VEVENT", "END:VCALENDAR");

  const blob = new Blob([lines.join("\r\n")], { type: "text/calendar;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `appointment-${(deal.company || "meeting").replace(/[^a-z0-9]+/gi, "-").toLowerCase()}.ics`;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
const fmtMoney = (n) => "$" + (n || 0).toLocaleString(undefined, { maximumFractionDigits: 0 });

// Tellemica logo lockup. `wordmark` renders the letter-spaced uppercase treatment
// from the portal design; default renders the signal-tower mark + wordmark.
function TellemicaLogo({ height = 30, light = false, wordmark = false }) {
  const wordColor = light ? "#FFFFFF" : INK;
  const dotColor = CYAN;
  if (wordmark) {
    return (
      <span style={{ fontFamily: "'Inter', system-ui, sans-serif", fontWeight: 800, fontSize: height, color: wordColor, letterSpacing: height * 0.18, textTransform: "uppercase", lineHeight: 1 }}>
        Tellemica
      </span>
    );
  }
  return (
    <span style={{ display: "inline-flex", alignItems: "center", gap: 10 }}>
      <svg width={height} height={height} viewBox="0 0 48 48" fill="none" aria-label="Tellemica">
        {/* signal tower: rising bars + broadcast arcs */}
        <rect x="21" y="26" width="6" height="16" rx="2" fill={dotColor} />
        <rect x="11" y="32" width="6" height="10" rx="2" fill={wordColor} opacity="0.85" />
        <rect x="31" y="32" width="6" height="10" rx="2" fill={wordColor} opacity="0.85" />
        <path d="M14 18 A16 16 0 0 1 34 18" stroke={dotColor} strokeWidth="3.2" strokeLinecap="round" fill="none" />
        <path d="M18.5 12.5 A9.5 9.5 0 0 1 29.5 12.5" stroke={wordColor} strokeWidth="3" strokeLinecap="round" fill="none" opacity="0.9" />
        <circle cx="24" cy="7.5" r="3" fill={dotColor} />
      </svg>
      <span style={{ fontFamily: "'Fraunces', serif", fontWeight: 600, fontSize: height * 0.72, color: wordColor, letterSpacing: 0.2 }}>
        Tellemica
      </span>
    </span>
  );
}

export default function App() {
  const [users, setUsers] = useState([]);
  const [entries, setEntries] = useState([]);
  const [deals, setDeals] = useState([]);
  const [goals, setGoals] = useState(DEFAULT_GOALS);
  const [userGoals, setUserGoals] = useState({});   // per-person overrides { userId: {calls,emails,appts} }
  const [liveUser, setLiveUser] = useState(null);   // the signed-in user's profile
  const [viewAsId, setViewAsId] = useState("");      // admin "view as" impersonation (empty = self)
  const [view, setView] = useState("dashboard");
  const [nudgeDismissed, setNudgeDismissed] = useState(false);
  const [search, setSearch] = useState("");
  const [companies, setCompanies] = useState([]);
  const [leads, setLeads] = useState([]);
  const [selectedCompanyId, setSelectedCompanyId] = useState(null); // when set, show company detail
  const [loaded, setLoaded] = useState(false);       // finished checking session
  const [authed, setAuthed] = useState(false);       // has a valid session

  // Pull all data the signed-in user is allowed to see. RLS filters server-side.
  const refetch = async () => {
    const [us, es, ds, g, ug, co, ld] = await Promise.all([
      api.listProfiles(), api.listEntries(), api.listDeals(), api.getGoals(), api.listUserGoals(), api.listCompanies(), api.listLeads(),
    ]);
    setUsers(us); setEntries(es); setDeals(ds); setUserGoals(ug); setCompanies(co); setLeads(ld);
    setGoals({ calls: g.calls, emails: g.emails, appts: g.appts });
  };

  // Navigate to a company's detail page by name (used by all the "linked" company names).
  const openCompanyByName = (name) => {
    const key = (name || "").trim().toLowerCase();
    const c = companies.find((x) => x.nameKey === key);
    if (c) { setSelectedCompanyId(c.id); setView("companies"); }
  };
  const openCompanyById = (id) => { setSelectedCompanyId(id); setView("companies"); };

  // On load and whenever auth state changes, resolve session -> profile -> data.
  useEffect(() => {
    let active = true;
    const boot = async () => {
      const session = await api.getSession();
      if (!session) { if (active) { setAuthed(false); setLiveUser(null); setLoaded(true); } return; }
      try {
        const me = await api.getMyProfile();
        if (!active) return;
        setLiveUser(me);
        setAuthed(true);
        await refetch();
      } catch (e) {
        console.error(e);
      } finally {
        if (active) setLoaded(true);
      }
    };
    boot();
    const { data: sub } = supabase.auth.onAuthStateChange((_event, session) => {
      if (!session) { setAuthed(false); setLiveUser(null); }
      else boot();
    });
    return () => { active = false; sub.subscription.unsubscribe(); };
  }, []);

  const login = async () => { setView("dashboard"); /* onAuthStateChange triggers boot() */ };
  const logout = async () => { await api.signOut(); setAuthed(false); setLiveUser(null); };

  // Save handlers: perform the mutation, then refetch so every view stays truthful.
  const saveEntries = async (mutate) => { await mutate(); await refetch(); };
  const saveDeals = async (mutate) => { await mutate(); await refetch(); };
  const saveUsers = async (mutate) => { await mutate(); await refetch(); };
  const saveGoals = async (g) => { await api.saveGoals(g); setGoals(g); };
  const saveUserGoals = async (mutate) => { await mutate(); await refetch(); };

  if (!loaded) {
    return <div style={{ minHeight: "100vh", background: PAPER, display: "grid", placeItems: "center" }}>
      <div style={{ color: INK, fontFamily: "system-ui", opacity: 0.5 }}>Loading…</div>
    </div>;
  }

  if (!authed || !liveUser) return <Login onLogin={login} />;

  const realIsAdmin = liveUser.role === "admin";
  // When an admin uses "View as", the app renders as that person. Their real
  // identity stays in liveUser (for the banner + exit); everything else uses effectiveUser.
  const impersonating = realIsAdmin && viewAsId && users.some((u) => u.id === viewAsId);
  const effectiveUser = impersonating ? users.find((u) => u.id === viewAsId) : liveUser;

  const role = effectiveUser.role;
  const canLog = true; // everyone can log activity (BDRs still tag "working for")
  const isAdmin = role === "admin";
  const canManageGoals = role === "admin" || role === "management";

  // Which PEOPLE this user can see (drives rep pickers, leaderboard rows, etc).
  // Note: a Sales Rep's data visibility is per-row (own + tagged to them), handled
  // below — but for building rep lists we include reps who have tagged them.
  const visibleUserIds = (() => {
    if (role === "admin" || role === "management") return users.map((u) => u.id);
    if (role === "sales") {
      // themselves + any BDR who has tagged them on an entry or deal
      const taggedBy = new Set([effectiveUser.id]);
      entries.forEach((e) => { if (e.taggedRepId === effectiveUser.id) taggedBy.add(e.userId); });
      deals.forEach((d) => { if (d.taggedRepId === effectiveUser.id) taggedBy.add(d.ownerId); });
      return [...taggedBy];
    }
    return [effectiveUser.id];
  })();

  // Row-level visibility: own rows, or (for a sales rep) rows tagged to them.
  // Admin/management see everything. This mirrors the server-side RLS.
  const canSeeEntry = (e) =>
    role === "admin" || role === "management" ||
    e.userId === effectiveUser.id || e.taggedRepId === effectiveUser.id;
  const canSeeDeal = (d) =>
    role === "admin" || role === "management" ||
    d.ownerId === effectiveUser.id || d.taggedRepId === effectiveUser.id;

  const visibleEntries = entries.filter(canSeeEntry);
  const visibleDeals = deals.filter(canSeeDeal);

  // Count overdue/due-today follow-ups across visible deals + scoped leads, for the nav badge & login nudge.
  const scopedLeads = leads.filter((l) => role === "admin" || role === "management" || l.assignedTo === effectiveUser.id || l.createdBy === effectiveUser.id);
  const dueCount = (() => {
    let n = 0;
    visibleDeals.forEach((d) => { if (d.nextActionDate && d.stage !== "won" && d.stage !== "lost" && daysUntil(d.nextActionDate) <= 0) n++; });
    scopedLeads.forEach((l) => { if (l.nextActionDate && !["Won", "Lost", "Dead"].includes(l.status) && daysUntil(l.nextActionDate) <= 0) n++; });
    return n;
  })();

  // Global search across companies, deals, and leads (scoped to what the user can see).
  const searchResults = (() => {
    const q = search.trim().toLowerCase();
    if (q.length < 2) return null;
    const hit = (s) => (s || "").toLowerCase().includes(q);
    const cos = companies.filter((c) => hit(c.name) || hit(c.ban) || hit(c.fan)).slice(0, 6)
      .map((c) => ({ type: "Company", id: c.id, title: c.name, sub: c.ban || "", onGo: () => { openCompanyById(c.id); setSearch(""); } }));
    const dls = visibleDeals.filter((d) => hit(d.company) || hit(d.contact)).slice(0, 6)
      .map((d) => ({ type: "Deal", id: d.id, title: d.company, sub: `${STAGE[d.stage]?.label || d.stage}${d.value ? ` · ${fmtMoney(d.value)}` : ""}`, onGo: () => { setView("pipeline"); setSearch(""); } }));
    const lds = scopedLeads.filter((l) => hit(l.company) || hit(l.contact) || hit(l.email)).slice(0, 6)
      .map((l) => ({ type: "Lead", id: l.id, title: l.company, sub: l.contact || l.status || "", onGo: () => { setView("leads"); setSearch(""); } }));
    return [...cos, ...dls, ...lds];
  })();

  const nav = [["dashboard", "Dashboard", BarChart3]];
  if (canLog) nav.push(["log", "Log Activity", Plus]);
  nav.push(["companies", "Companies", Building2]);
  nav.push(["leads", "Leads", UserCheck]);
  nav.push(["activity", "Activity Log", Table2]);
  nav.push(["pipeline", "Pipeline", Briefcase]);
  nav.push(["followups", "Follow-ups", Clock]);
  if (role !== "bdr") nav.push(["reports", "Reports", TrendingUp]);
  if (canManageGoals) nav.push(["goals", "Goals", Target]);
  if (isAdmin) nav.push(["admin", "Admin Portal", Shield]);

  return (
    <div style={{ minHeight: "100vh", background: PAPER, fontFamily: "'Inter', system-ui, sans-serif", color: INK, overflowX: "hidden" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Fraunces:opsz,wght@9..144,500;9..144,600&display=swap');
        * { box-sizing: border-box; }
        .tap { cursor: pointer; transition: all .15s ease; }
        .tap:hover { transform: translateY(-1px); }
        @keyframes addPulse {
          0%, 100% { box-shadow: 0 8px 26px rgba(15,150,220,.50), 0 0 0 1px rgba(255,255,255,.14) inset; }
          50% { box-shadow: 0 10px 34px rgba(15,150,220,.82), 0 0 0 1px rgba(255,255,255,.20) inset; }
        }
        .fab-add { animation: addPulse 2.4s ease-in-out infinite; }
        .fab-add:hover { transform: translateY(-2px) scale(1.04); animation: none; box-shadow: 0 12px 36px rgba(15,150,220,.85); }
        @media (max-width: 720px){ .fab-add{ right: 16px !important; bottom: 16px !important; padding: 14px 20px !important; } }
        input, select, textarea { font-family: inherit; }
        @media (max-width: 720px){ .grid-4{grid-template-columns:1fr 1fr!important} .charts{grid-template-columns:1fr!important} .logwrap{grid-template-columns:1fr!important} }
        @media (max-width: 720px){
          .hdr-row{ flex-direction: column !important; align-items: stretch !important; gap: 10px !important; }
          .hdr-right{ width: 100%; justify-content: flex-start !important; gap: 10px !important; }
          .hdr-nav{ width: 100%; justify-content: flex-start; }
          .hdr-nav button{ padding: 7px 10px !important; font-size: 12.5px !important; }
          .hdr-user{ width: 100%; justify-content: space-between !important; }
          header{ padding: 12px 16px !important; }
        }
      `}</style>

      <header style={{ background: `linear-gradient(100deg, ${INK}, ${NAVY2})`, color: PAPER, padding: "16px 24px", borderBottom: `3px solid ${CYAN}` }}>
        <div className="hdr-row" style={{ maxWidth: 1120, margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
            <TellemicaLogo height={19} light wordmark />
            <span style={{ fontSize: 12, opacity: 0.7, letterSpacing: 1, textTransform: "uppercase", borderLeft: "1px solid rgba(255,255,255,.25)", paddingLeft: 14 }}>Command Center</span>
          </div>
          <div className="hdr-right" style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap", minWidth: 0, maxWidth: "100%" }}>
            <nav className="hdr-nav" style={{ display: "flex", flexWrap: "wrap", gap: 6, background: "rgba(255,255,255,.08)", padding: 4, borderRadius: 10, maxWidth: "100%" }}>
              {nav.map(([id, label, Icon]) => (
                <button key={id} onClick={() => { setView(id); if (id !== "companies") setSelectedCompanyId(null); }} className="tap"
                  style={{ display: "flex", alignItems: "center", gap: 6, border: "none", borderRadius: 7, padding: "8px 13px", fontSize: 13.5, fontWeight: 500, whiteSpace: "nowrap",
                    background: view === id ? PAPER : "transparent", color: view === id ? INK : PAPER }}>
                  <Icon size={15} /> {label}
                  {id === "followups" && dueCount > 0 && (
                    <span style={{ background: "#B4453F", color: "#fff", fontSize: 10.5, fontWeight: 700, borderRadius: 99, padding: "1px 6px", minWidth: 17, textAlign: "center", lineHeight: 1.5 }}>{dueCount}</span>
                  )}
                </button>
              ))}
            </nav>
            <div className="hdr-user" style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", minWidth: 0 }}>
              {realIsAdmin && (
                <div style={{ position: "relative" }}>
                  <select value={viewAsId} onChange={(e) => { setViewAsId(e.target.value); setView("dashboard"); }}
                    title="View the app as another role/person"
                    style={{ appearance: "none", background: impersonating ? CYAN : "rgba(255,255,255,.12)", color: impersonating ? INK : PAPER, border: impersonating ? "none" : "1px solid rgba(255,255,255,.25)", borderRadius: 8, padding: "8px 30px 8px 12px", fontSize: 13, fontWeight: 600, cursor: "pointer", maxWidth: "100%", textOverflow: "ellipsis" }}>
                    <option value="">View as… (Admin)</option>
                    {["management", "sales", "bdr"].map((r) => {
                      const group = users.filter((u) => u.role === r).sort((a, b) => a.name.localeCompare(b.name));
                      if (group.length === 0) return null;
                      return (
                        <optgroup key={r} label={ROLES[r].label}>
                          {group.map((u) => <option key={u.id} value={u.id}>{u.name}</option>)}
                        </optgroup>
                      );
                    })}
                  </select>
                  <ChevronDown size={14} style={{ position: "absolute", right: 10, top: 10, pointerEvents: "none", color: impersonating ? INK : PAPER, opacity: 0.7 }} />
                </div>
              )}
              <div style={{ textAlign: "right", lineHeight: 1.3 }}>
                <div style={{ fontSize: 13, fontWeight: 500 }}>{impersonating ? effectiveUser.name : liveUser.name}</div>
                <div style={{ fontSize: 11, opacity: 0.6 }}>{ROLES[role].label}{impersonating ? " · viewing" : ""}</div>
              </div>
              <button onClick={logout} className="tap" title="Sign out"
                style={{ background: "rgba(255,255,255,.1)", border: "none", borderRadius: 8, padding: 8, color: PAPER, cursor: "pointer" }}>
                <LogOut size={16} />
              </button>
            </div>
          </div>
        </div>
      </header>

      <div style={{ background: CARD, borderBottom: `1px solid ${LINE_C}`, padding: "10px 24px" }}>
        <div style={{ maxWidth: 1120, margin: "0 auto", position: "relative" }}>
          <input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Search companies, deals, leads…"
            style={{ width: "100%", maxWidth: 420, border: `1px solid ${LINE_C}`, borderRadius: 9, padding: "9px 13px", fontSize: 13.5, outline: "none", background: "#F8FAFC" }}
          />
          {searchResults && (
            <div style={{ position: "absolute", top: "calc(100% + 4px)", left: 0, width: "100%", maxWidth: 420, background: CARD, border: `1px solid ${LINE_C}`, borderRadius: 11, boxShadow: "0 12px 32px rgba(11,42,74,.14)", zIndex: 40, overflow: "hidden" }}>
              {searchResults.length === 0 ? (
                <div style={{ padding: 16, fontSize: 13, opacity: 0.55, textAlign: "center" }}>No matches for "{search}"</div>
              ) : searchResults.map((r, i) => (
                <button key={r.type + r.id} onClick={r.onGo} className="tap"
                  style={{ display: "flex", alignItems: "center", gap: 10, width: "100%", textAlign: "left", background: "transparent", border: "none", borderTop: i ? `1px solid ${LINE_C}` : "none", padding: "10px 14px", cursor: "pointer" }}>
                  <span style={{ fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.4, color: r.type === "Company" ? INK : r.type === "Deal" ? EMAIL : CALL, background: (r.type === "Company" ? INK : r.type === "Deal" ? EMAIL : CALL) + "14", borderRadius: 5, padding: "3px 7px", flexShrink: 0 }}>{r.type}</span>
                  <span style={{ minWidth: 0 }}>
                    <span style={{ display: "block", fontSize: 13.5, fontWeight: 600, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{r.title}</span>
                    {r.sub && <span style={{ display: "block", fontSize: 12, opacity: 0.55 }}>{r.sub}</span>}
                  </span>
                </button>
              ))}
            </div>
          )}
        </div>
      </div>

      {impersonating && (
        <div style={{ background: CYAN, color: INK, padding: "10px 24px" }}>
          <div style={{ maxWidth: 1120, margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 10 }}>
            <div style={{ fontSize: 13.5, fontWeight: 600, display: "flex", alignItems: "center", gap: 8 }}>
              <Eye size={16} /> Viewing as <b>{effectiveUser.name}</b> ({ROLES[effectiveUser.role].label}) — this is a preview of what they see. Changes you make are still made as {liveUser.name}.
            </div>
            <button onClick={() => { setViewAsId(""); setView("dashboard"); }} className="tap"
              style={{ background: INK, color: PAPER, border: "none", borderRadius: 8, padding: "7px 14px", fontSize: 13, fontWeight: 600, cursor: "pointer", whiteSpace: "nowrap" }}>
              Exit preview
            </button>
          </div>
        </div>
      )}
      <main style={{ maxWidth: 1120, margin: "0 auto", padding: "26px 24px 60px" }}>
        {dueCount > 0 && !nudgeDismissed && view !== "followups" && (
          <div style={{ display: "flex", alignItems: "center", gap: 12, background: "#FFF4E5", border: "1px solid #F3D9B0", borderRadius: 12, padding: "12px 16px", marginBottom: 18 }}>
            <Clock size={18} color="#B7791F" style={{ flexShrink: 0 }} />
            <div style={{ flex: 1, fontSize: 13.5, color: "#8A6D3B" }}>
              You have <b>{dueCount}</b> follow-up{dueCount > 1 ? "s" : ""} due or overdue.
            </div>
            <button onClick={() => setView("followups")} className="tap" style={{ background: INK, color: PAPER, border: "none", borderRadius: 8, padding: "7px 14px", fontSize: 13, fontWeight: 600, cursor: "pointer", whiteSpace: "nowrap" }}>Review</button>
            <button onClick={() => setNudgeDismissed(true)} className="tap" title="Dismiss" style={{ background: "transparent", border: "none", color: "#8A6D3B", cursor: "pointer", display: "flex", alignItems: "center", padding: 4 }}><X size={16} /></button>
          </div>
        )}
        {view === "dashboard" && (
          <Dashboard entries={visibleEntries} deals={visibleDeals} users={users} goals={goals} saveGoals={saveGoals}
            leads={leads.filter((l) => role === "admin" || role === "management" || l.assignedTo === effectiveUser.id || l.createdBy === effectiveUser.id)}
            userGoals={userGoals} liveUser={effectiveUser} visibleUserIds={visibleUserIds} setView={setView} canLog={canLog}
            onOpenCompany={(id, name) => id ? openCompanyById(id) : openCompanyByName(name)} />
        )}
        {view === "log" && canLog && (
          <LogView liveUser={effectiveUser} entries={entries} saveEntries={saveEntries} users={users} allEntries={visibleEntries} visibleUserIds={visibleUserIds} />
        )}
        {view === "pipeline" && (
          <Pipeline deals={visibleDeals} allDeals={deals} saveDeals={saveDeals}
            liveUser={effectiveUser} users={users} visibleUserIds={visibleUserIds}
            entries={entries} saveEntries={saveEntries} />
        )}
        {view === "followups" && (
          <FollowUpsView deals={visibleDeals} leads={leads.filter((l) => role === "admin" || role === "management" || l.assignedTo === effectiveUser.id || l.createdBy === effectiveUser.id)}
            users={users} liveUser={effectiveUser} setView={setView} />
        )}
        {view === "reports" && role !== "bdr" && (
          <ReportsView entries={visibleEntries} deals={visibleDeals} users={users} liveUser={effectiveUser} visibleUserIds={visibleUserIds} />
        )}
        {view === "companies" && (
          selectedCompanyId ? (
            <CompanyDetail
              companyId={selectedCompanyId}
              companies={companies}
              entries={visibleEntries}
              deals={visibleDeals}
              users={users}
              effectiveUser={effectiveUser}
              saveDeals={saveDeals}
              visibleUserIds={visibleUserIds}
              onBack={() => setSelectedCompanyId(null)}
              onOpenCompany={openCompanyById}
              refetch={refetch}
            />
          ) : (
            <CompaniesList
              companies={companies}
              entries={visibleEntries}
              deals={visibleDeals}
              onOpen={openCompanyById}
            />
          )
        )}
        {view === "leads" && (
          <LeadsView
            leads={leads}
            users={users}
            effectiveUser={effectiveUser}
            visibleUserIds={visibleUserIds}
            refetch={refetch}
            onOpenCompany={openCompanyByName}
          />
        )}
        {view === "activity" && (
          <>
            <div style={{ marginBottom: 20 }}>
              <h1 style={{ fontFamily: "'Fraunces', serif", fontSize: 29, fontWeight: 600, margin: 0 }}>Activity Log</h1>
              <p style={{ margin: "4px 0 0", opacity: 0.55, fontSize: 14 }}>
                Every activity record in your scope. Click a column to sort, or export to Excel.
              </p>
            </div>
            <ActivityTable entries={visibleEntries} users={users} liveUser={effectiveUser}
              onOpenCompany={(id, name) => id ? openCompanyById(id) : openCompanyByName(name)} />
          </>
        )}
        {view === "goals" && canManageGoals && (
          <GoalsManager users={users} visibleUserIds={visibleUserIds} liveUser={effectiveUser}
            goals={goals} saveGoals={saveGoals} userGoals={userGoals} saveUserGoals={saveUserGoals} />
        )}
        {view === "admin" && isAdmin && (
          <AdminPortal users={users} saveUsers={saveUsers} entries={entries} saveEntries={saveEntries} deals={deals} saveDeals={saveDeals} />
        )}
      </main>

      {canLog && view !== "log" && (
        <button onClick={() => setView("log")} className="fab-add" aria-label="Add Activity"
          style={{ position: "fixed", right: 24, bottom: 24, zIndex: 60, display: "flex", alignItems: "center", gap: 9,
            background: `linear-gradient(90deg, ${BTN_A}, ${BTN_B})`, color: "#fff", border: "none", borderRadius: 999,
            padding: "15px 24px", fontSize: 15.5, fontWeight: 700, letterSpacing: 0.2, cursor: "pointer",
            boxShadow: "0 8px 26px rgba(15,150,220,.55), 0 0 0 1px rgba(255,255,255,.14) inset" }}>
          <Plus size={20} strokeWidth={2.7} /> Add Activity
        </button>
      )}
    </div>
  );
}

function Login({ onLogin }) {
  const [tab, setTab] = useState("signin"); // signin | request
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);
  const [show, setShow] = useState(false);
  // request-access form
  const [reqName, setReqName] = useState("");
  const [reqEmail, setReqEmail] = useState("");
  const [reqSent, setReqSent] = useState(false);

  const submit = async () => {
    if (busy) return;
    setBusy(true); setErr("");
    try {
      await api.signIn(email.trim(), password);
      await onLogin(); // App's auth listener refetches everything
    } catch (e) {
      setErr(e?.message === "Invalid login credentials" ? "Email or password is incorrect." : (e?.message || "Sign-in failed."));
    } finally {
      setBusy(false);
    }
  };

  const inp = {
    width: "100%", padding: "12px 14px", border: `1px solid ${LINE_C}`, borderRadius: 10,
    fontSize: 14, color: INK, background: "#fff", outline: "none",
  };
  const lbl = { fontSize: 11.5, fontWeight: 700, letterSpacing: 0.8, textTransform: "uppercase", color: "#5B6B7F", marginBottom: 7 };

  return (
    <div style={{ minHeight: "100vh", position: "relative", background: `linear-gradient(135deg, #0E1C32 0%, #102B5C 55%, #2C4A86 100%)`, display: "grid", placeItems: "center", padding: 20, fontFamily: "'Inter', system-ui, sans-serif" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
        .portal-input:focus { border-color: ${CYAN} !important; box-shadow: 0 0 0 3px ${CYAN}22 !important; }
      `}</style>

      <div style={{ width: "100%", maxWidth: 420, background: "#fff", borderRadius: 16, overflow: "hidden", boxShadow: "0 24px 60px rgba(8,20,40,.45)" }}>
        {/* Header block */}
        <div style={{ background: INK, padding: "34px 26px 26px", textAlign: "center" }}>
          <TellemicaLogo height={26} light wordmark />
          <div style={{ fontSize: 13, fontWeight: 700, letterSpacing: 2, textTransform: "uppercase", color: "#fff", marginTop: 12 }}>Sales Command Center</div>
        </div>

        {/* Tabs */}
        <div style={{ display: "flex", borderBottom: `1px solid ${LINE_C}` }}>
          {[["signin", "Sign In"], ["request", "Request Access"]].map(([id, label]) => (
            <button key={id} onClick={() => { setTab(id); setErr(""); }} className="tap"
              style={{ flex: 1, background: "transparent", border: "none", cursor: "pointer", padding: "16px 0", fontSize: 13, fontWeight: 700, letterSpacing: 1, textTransform: "uppercase",
                color: tab === id ? CYAN : "#8494A6", borderBottom: tab === id ? `3px solid ${CYAN}` : "3px solid transparent", marginBottom: -1 }}>
              {label}
            </button>
          ))}
        </div>

        {/* Body */}
        <div style={{ padding: "26px" }}>
          {tab === "signin" ? (
            <>
              <div style={{ marginBottom: 16 }}>
                <div style={lbl}>Email</div>
                <input className="portal-input" value={email} onChange={(e) => { setEmail(e.target.value); setErr(""); }} onKeyDown={(e) => e.key === "Enter" && submit()}
                  placeholder="you@tellemica.com" style={inp} autoFocus />
              </div>
              <div style={{ marginBottom: 8 }}>
                <div style={lbl}>Password</div>
                <div style={{ position: "relative" }}>
                  <input className="portal-input" type={show ? "text" : "password"} value={password} onChange={(e) => { setPassword(e.target.value); setErr(""); }} onKeyDown={(e) => e.key === "Enter" && submit()}
                    placeholder="••••••••" style={inp} />
                  <button onClick={() => setShow(!show)} className="tap" style={{ position: "absolute", right: 12, top: 11, background: "none", border: "none", cursor: "pointer", opacity: 0.45 }}>
                    {show ? <EyeOff size={16} /> : <Eye size={16} />}
                  </button>
                </div>
              </div>
              {err && <div style={{ color: "#B4453F", fontSize: 13, margin: "6px 0 0" }}>{err}</div>}
              <button onClick={submit} disabled={busy} className="tap"
                style={{ width: "100%", marginTop: 20, background: `linear-gradient(90deg, ${BTN_A}, ${BTN_B})`, color: "#fff", border: "none", borderRadius: 10, padding: 14, fontSize: 15, fontWeight: 700, cursor: busy ? "default" : "pointer", opacity: busy ? 0.7 : 1, boxShadow: "0 6px 16px rgba(24,84,187,.3)" }}>
                {busy ? "Signing in…" : "Sign into CRM"}
              </button>
              <div style={{ textAlign: "center", marginTop: 18, fontSize: 13.5, color: "#5B6B7F" }}>
                No account?{" "}
                <button onClick={() => { setTab("request"); setErr(""); }} className="tap" style={{ background: "none", border: "none", color: BTN_A, fontWeight: 700, cursor: "pointer", fontSize: 13.5, padding: 0 }}>Request access</button>
              </div>
            </>
          ) : reqSent ? (
            <div style={{ textAlign: "center", padding: "12px 0" }}>
              <div style={{ width: 48, height: 48, borderRadius: "50%", background: CYAN + "1a", display: "grid", placeItems: "center", margin: "0 auto 14px" }}>
                <CheckCircle2 size={26} color={CYAN} />
              </div>
              <div style={{ fontSize: 16, fontWeight: 700, color: INK, marginBottom: 6 }}>Request submitted</div>
              <div style={{ fontSize: 13.5, color: "#5B6B7F", lineHeight: 1.5 }}>An admin will review your request and set up your account. You'll get an email when access is granted.</div>
              <button onClick={() => { setTab("signin"); setReqSent(false); setReqName(""); setReqEmail(""); }} className="tap"
                style={{ marginTop: 18, background: "transparent", border: `1px solid ${LINE_C}`, borderRadius: 10, padding: "11px 20px", fontSize: 14, fontWeight: 600, color: INK, cursor: "pointer" }}>
                Back to sign in
              </button>
            </div>
          ) : (
            <>
              <div style={{ fontSize: 13.5, color: "#5B6B7F", marginBottom: 18, lineHeight: 1.5 }}>
                Request access to the Tellemica team portal. An admin will approve and assign your role.
              </div>
              <div style={{ marginBottom: 16 }}>
                <div style={lbl}>Full name</div>
                <input className="portal-input" value={reqName} onChange={(e) => setReqName(e.target.value)} placeholder="Jane Smith" style={inp} autoFocus />
              </div>
              <div style={{ marginBottom: 8 }}>
                <div style={lbl}>Work email</div>
                <input className="portal-input" value={reqEmail} onChange={(e) => setReqEmail(e.target.value)} placeholder="you@tellemica.com" style={inp} />
              </div>
              <button onClick={() => { if (reqName.trim() && reqEmail.trim()) setReqSent(true); }} className="tap"
                style={{ width: "100%", marginTop: 20, background: `linear-gradient(90deg, ${BTN_A}, ${BTN_B})`, color: "#fff", border: "none", borderRadius: 10, padding: 14, fontSize: 15, fontWeight: 700, cursor: "pointer", boxShadow: "0 6px 16px rgba(24,84,187,.3)" }}>
                Submit Request
              </button>
              <div style={{ textAlign: "center", marginTop: 18, fontSize: 13.5, color: "#5B6B7F" }}>
                Already have access?{" "}
                <button onClick={() => setTab("signin")} className="tap" style={{ background: "none", border: "none", color: BTN_A, fontWeight: 700, cursor: "pointer", fontSize: 13.5, padding: 0 }}>Sign in</button>
              </div>
            </>
          )}
        </div>
      </div>

      {/* LIVE status pill */}
      <div style={{ position: "fixed", left: 18, bottom: 18, display: "flex", alignItems: "center", gap: 6, background: "#189B72", color: "#fff", borderRadius: 20, padding: "5px 12px", fontSize: 11, fontWeight: 700, letterSpacing: 1, textTransform: "uppercase", boxShadow: "0 4px 12px rgba(0,0,0,.2)" }}>
        <span style={{ width: 7, height: 7, borderRadius: "50%", background: "#fff" }} /> Live
      </div>
    </div>
  );
}

// ---- Leads ----
const LEAD_STATUSES = ["new", "contacted", "qualified", "converted", "dead"];
const LEAD_STATUS_META = {
  new:       { label: "New",       color: "#2C6FB5", bg: "#E8F1FB" },
  contacted: { label: "Contacted", color: "#B5852C", bg: "#FBF3E1" },
  qualified: { label: "Qualified", color: "#2C8C6F", bg: "#E3F5EE" },
  converted: { label: "Converted", color: "#189B72", bg: "#DDF3EA" },
  dead:      { label: "Dead",      color: "#8494A6", bg: "#EEF1F4" },
};

function LeadsView({ leads, users, effectiveUser, visibleUserIds, refetch, onOpenCompany }) {
  const role = effectiveUser.role;
  const canManageAll = role === "admin" || role === "management";
  const [q, setQ] = useState("");
  const [statusFilter, setStatusFilter] = useState("all");
  const [assigneeFilter, setAssigneeFilter] = useState("all");
  const [showUpload, setShowUpload] = useState(false);
  const [busy, setBusy] = useState("");

  // People a lead can be assigned to = BDRs + Sales Reps in scope.
  const assignable = users.filter((u) => (u.role === "bdr" || u.role === "sales") && (canManageAll || visibleUserIds.includes(u.id)));
  const nameOf = (id) => { const u = users.find((x) => x.id === id); return u ? u.name : ""; };

  const filtered = leads.filter((l) => {
    if (q && !(`${l.company} ${l.contact} ${l.email}`.toLowerCase().includes(q.trim().toLowerCase()))) return false;
    if (statusFilter !== "all" && l.status !== statusFilter) return false;
    if (assigneeFilter !== "all") {
      if (assigneeFilter === "unassigned" && l.assignedTo) return false;
      if (assigneeFilter !== "unassigned" && l.assignedTo !== assigneeFilter) return false;
    }
    return true;
  });

  const setStatus = async (lead, status) => { setBusy("Saving…"); try { await api.updateLead(lead.id, { status }); await refetch(); } finally { setBusy(""); } };
  const setAssignee = async (lead, assignedTo) => { setBusy("Saving…"); try { await api.updateLead(lead.id, { assignedTo }); await refetch(); } finally { setBusy(""); } };
  const del = async (lead) => { if (confirm(`Delete lead "${lead.company}"?`)) { await api.deleteLead(lead.id); await refetch(); } };

  // status counts for the summary chips
  const counts = LEAD_STATUSES.reduce((a, s) => ({ ...a, [s]: leads.filter((l) => l.status === s).length }), {});

  return (
    <>
      <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", flexWrap: "wrap", gap: 12, marginBottom: 18 }}>
        <div>
          <h1 style={{ fontFamily: "'Fraunces', serif", fontSize: 29, fontWeight: 600, margin: 0 }}>Leads</h1>
          <p style={{ margin: "4px 0 0", opacity: 0.55, fontSize: 14 }}>
            {filtered.length} {filtered.length === 1 ? "lead" : "leads"}{canManageAll ? "" : " assigned to you"}. {busy && <b style={{ color: CALL }}>{busy}</b>}
          </p>
        </div>
        <button onClick={() => setShowUpload(true)} className="tap"
          style={{ display: "flex", alignItems: "center", gap: 7, background: `linear-gradient(90deg, ${BTN_A}, ${BTN_B})`, color: "#fff", border: "none", borderRadius: 9, padding: "10px 16px", fontSize: 14, fontWeight: 600, cursor: "pointer" }}>
          <FileSpreadsheet size={15} /> Upload leads
        </button>
      </div>

      {/* Status summary chips */}
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 16 }}>
        {LEAD_STATUSES.map((s) => (
          <button key={s} onClick={() => setStatusFilter(statusFilter === s ? "all" : s)} className="tap"
            style={{ display: "flex", alignItems: "center", gap: 6, border: statusFilter === s ? `1px solid ${LEAD_STATUS_META[s].color}` : `1px solid ${LINE_C}`, background: statusFilter === s ? LEAD_STATUS_META[s].bg : "#fff", borderRadius: 20, padding: "6px 12px", fontSize: 12.5, fontWeight: 600, cursor: "pointer", color: LEAD_STATUS_META[s].color }}>
            {LEAD_STATUS_META[s].label} <span style={{ opacity: 0.7 }}>{counts[s]}</span>
          </button>
        ))}
      </div>

      {/* Filters */}
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 14 }}>
        <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search company, contact, email…" style={{ ...inputStyle, width: 280, marginBottom: 0 }} />
        {canManageAll && (
          <div style={{ position: "relative" }}>
            <select value={assigneeFilter} onChange={(e) => setAssigneeFilter(e.target.value)}
              style={{ appearance: "none", background: CARD, border: `1px solid ${LINE_C}`, borderRadius: 9, padding: "10px 34px 10px 14px", fontSize: 14, cursor: "pointer" }}>
              <option value="all">All assignees</option>
              <option value="unassigned">Unassigned</option>
              {assignable.map((u) => <option key={u.id} value={u.id}>{u.name}</option>)}
            </select>
            <ChevronDown size={15} style={{ position: "absolute", right: 12, top: 12, pointerEvents: "none", opacity: 0.5 }} />
          </div>
        )}
      </div>

      <div style={{ background: CARD, border: `1px solid ${LINE_C}`, borderRadius: 14, overflow: "hidden" }}>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, minWidth: 900 }}>
            <thead>
              <tr style={{ background: "#F1F5F9" }}>
                {["Company", "Contact", "Phone", "Email", "Status", "Assigned to", ""].map((h) => (
                  <th key={h} style={{ textAlign: "left", padding: "11px 14px", fontSize: 11.5, fontWeight: 700, textTransform: "uppercase", color: "#5A6B7B", whiteSpace: "nowrap" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (
                <tr><td colSpan={7} style={{ padding: 28, textAlign: "center", opacity: 0.5 }}>No leads{leads.length ? " match your filters" : " yet — upload a lead list to get started"}.</td></tr>
              ) : filtered.map((l) => (
                <tr key={l.id} style={{ borderTop: `1px solid ${LINE_C}` }}>
                  <td style={{ padding: "10px 14px", fontWeight: 600, whiteSpace: "nowrap" }}>
                    <button onClick={() => onOpenCompany(l.company)} className="tap" style={{ background: "transparent", border: "none", color: EMAIL, fontWeight: 600, fontSize: 13, cursor: "pointer", padding: 0 }}>{l.company}</button>
                  </td>
                  <td style={{ padding: "10px 14px", whiteSpace: "nowrap" }}>{l.contact || "—"}</td>
                  <td style={{ padding: "10px 14px", whiteSpace: "nowrap" }}>{l.phone || "—"}</td>
                  <td style={{ padding: "10px 14px", whiteSpace: "nowrap" }}>{l.email || "—"}</td>
                  <td style={{ padding: "10px 14px" }}>
                    <select value={l.status} onChange={(e) => setStatus(l, e.target.value)}
                      style={{ appearance: "none", border: "none", borderRadius: 20, padding: "5px 12px", fontSize: 12, fontWeight: 700, cursor: "pointer", color: LEAD_STATUS_META[l.status].color, background: LEAD_STATUS_META[l.status].bg }}>
                      {LEAD_STATUSES.map((s) => <option key={s} value={s}>{LEAD_STATUS_META[s].label}</option>)}
                    </select>
                  </td>
                  <td style={{ padding: "10px 14px" }}>
                    {canManageAll ? (
                      <div style={{ position: "relative" }}>
                        <select value={l.assignedTo || ""} onChange={(e) => setAssignee(l, e.target.value)}
                          style={{ appearance: "none", background: "#fff", border: `1px solid ${LINE_C}`, borderRadius: 8, padding: "6px 26px 6px 10px", fontSize: 13, cursor: "pointer" }}>
                          <option value="">Unassigned</option>
                          {assignable.map((u) => <option key={u.id} value={u.id}>{u.name}</option>)}
                        </select>
                        <ChevronDown size={13} style={{ position: "absolute", right: 8, top: 9, pointerEvents: "none", opacity: 0.5 }} />
                      </div>
                    ) : (
                      <span style={{ fontSize: 13 }}>{nameOf(l.assignedTo) || "—"}</span>
                    )}
                  </td>
                  <td style={{ padding: "10px 14px", textAlign: "right" }}>
                    {(canManageAll || l.createdBy === effectiveUser.id) && (
                      <button onClick={() => del(l)} className="tap" style={{ background: "transparent", border: "none", cursor: "pointer", opacity: 0.4 }}><Trash2 size={14} /></button>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {showUpload && <LeadUpload users={users} assignable={assignable} refetch={refetch} onClose={() => setShowUpload(false)} />}
    </>
  );
}

function LeadUpload({ users, assignable, refetch, onClose }) {
  const [rows, setRows] = useState(null);
  const [fileName, setFileName] = useState("");
  const [busy, setBusy] = useState(false);
  const [done, setDone] = useState("");
  const fileRef = React.useRef(null);
  const COLS = ["Company", "Contact", "Phone", "Email", "BAN", "FAN", "Assign To", "Status", "Notes"];

  const matchUser = (name) => { const n = (name || "").trim().toLowerCase(); if (!n) return null; return assignable.find((u) => u.name.trim().toLowerCase() === n) || null; };

  const downloadTemplate = async () => {
    const example = { Company: "Acme Corp", Contact: "Jane Smith", Phone: "(610) 555-0100", Email: "jane@acme.com", BAN: "123456789", FAN: "987654321", "Assign To": (assignable[0] && assignable[0].name) || "", Status: "new", Notes: "Referred by partner" };
    const blank = Object.fromEntries(COLS.map((c) => [c, ""]));
    const ws = XLSX.utils.json_to_sheet([example, blank], { header: COLS });
    ws["!cols"] = [{ wch: 22 }, { wch: 18 }, { wch: 16 }, { wch: 24 }, { wch: 14 }, { wch: 14 }, { wch: 18 }, { wch: 12 }, { wch: 34 }];
    // Lists sheet for dropdowns
    const assignOpts = assignable.map((u) => u.name);
    const maxLen = Math.max(assignOpts.length, LEAD_STATUSES.length);
    const listAoa = [["AssignTo", "Status"]];
    for (let i = 0; i < maxLen; i++) listAoa.push([assignOpts[i] || "", LEAD_STATUSES[i] || ""]);
    const wsList = XLSX.utils.aoa_to_sheet(listAoa);
    const ref = [{ Column: "Assign To", "Accepted values": "(leave blank = unassigned)" }, ...assignable.map((u) => ({ Column: "Assign To", "Accepted values": `${u.name} (${ROLES[u.role].label})` })), ...LEAD_STATUSES.map((s) => ({ Column: "Status", "Accepted values": s }))];
    const wsRef = XLSX.utils.json_to_sheet(ref); wsRef["!cols"] = [{ wch: 16 }, { wch: 40 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Leads");
    XLSX.utils.book_append_sheet(wb, wsRef, "Reference");
    XLSX.utils.book_append_sheet(wb, wsList, "Lists");
    wb.Workbook = { Sheets: wb.SheetNames.map((name) => ({ name, Hidden: name === "Leads" ? 0 : 1 })) };
    const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });
    try {
      const zip = await JSZip.loadAsync(buf);
      let xml = await zip.file("xl/worksheets/sheet1.xml").async("string");
      const N = 2000;
      const dvs = [
        { sqref: `G2:G${N + 1}`, f: `Lists!$A$2:$A$${assignOpts.length + 1}` }, // Assign To
        { sqref: `H2:H${N + 1}`, f: `Lists!$B$2:$B$${LEAD_STATUSES.length + 1}` }, // Status
      ];
      const dvXml = `<dataValidations count="${dvs.length}">` + dvs.map((d) => `<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="${d.sqref}"><formula1>${d.f}</formula1></dataValidation>`).join("") + `</dataValidations>`;
      xml = xml.includes("</sheetData>") ? xml.replace("</sheetData>", "</sheetData>" + dvXml) : xml.replace("</worksheet>", dvXml + "</worksheet>");
      zip.file("xl/worksheets/sheet1.xml", xml);
      const blob = await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob); const a = document.createElement("a"); a.href = url; a.download = "tellemica-leads-template.xlsx"; document.body.appendChild(a); a.click(); document.body.removeChild(a); setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch { XLSX.writeFile(wb, "tellemica-leads-template.xlsx"); }
  };

  const onFile = async (file) => {
    setDone(""); setRows(null); setFileName(file.name);
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const raw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "", raw: false });
    const parsed = raw.map((r, i) => {
      const errors = [];
      const company = String(r["Company"] || "").trim();
      if (!company) errors.push("Company is required");
      let assignedTo = null;
      const a = String(r["Assign To"] || "").trim();
      if (a) { const u = matchUser(a); if (!u) errors.push(`"Assign To" — no BDR/Sales Rep named "${a}"`); else assignedTo = u.id; }
      let status = String(r["Status"] || "new").trim().toLowerCase();
      if (!LEAD_STATUSES.includes(status)) { if (status) errors.push(`"Status" — "${status}" isn't valid`); status = "new"; }
      return {
        _row: i + 2, errors, company, contact: String(r["Contact"] || "").trim(), phone: String(r["Phone"] || "").trim(),
        email: String(r["Email"] || "").trim(), ban: String(r["BAN"] || "").trim(), fan: String(r["FAN"] || "").trim(),
        notes: String(r["Notes"] || "").trim(), assignedTo, status,
        assigneeName: assignedTo ? (assignable.find((u) => u.id === assignedTo) || {}).name : "Unassigned",
      };
    }).filter((r) => r.company || r.errors.length);
    setRows(parsed);
  };

  const valid = (rows || []).filter((r) => !r.errors.length);
  const invalid = (rows || []).filter((r) => r.errors.length);

  const doImport = async () => {
    if (!valid.length) return; setBusy(true);
    try {
      await api.addLeadsBulk(valid.map((r) => ({ company: r.company, contact: r.contact, phone: r.phone, email: r.email, ban: r.ban, fan: r.fan, notes: r.notes, status: r.status, assignedTo: r.assignedTo })));
      await refetch(); setDone(`Imported ${valid.length} ${valid.length === 1 ? "lead" : "leads"}.`); setRows(null); setFileName(""); if (fileRef.current) fileRef.current.value = "";
    } catch (e) { setDone("Import failed: " + (e.message || "")); } finally { setBusy(false); }
  };

  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(11,42,74,.5)", display: "grid", placeItems: "center", padding: 20, zIndex: 50 }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: PAPER, borderRadius: 16, padding: 24, width: "100%", maxWidth: 680, maxHeight: "88vh", overflowY: "auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
          <h3 style={{ fontFamily: "'Fraunces', serif", fontSize: 21, fontWeight: 600, margin: 0 }}>Upload leads</h3>
          <button onClick={onClose} className="tap" style={{ background: "transparent", border: "none", cursor: "pointer" }}><X size={20} /></button>
        </div>
        <p style={{ margin: "0 0 16px", fontSize: 13.5, opacity: 0.6, lineHeight: 1.5 }}>Download the template, fill in one row per lead, then upload. You can assign during upload or from the Leads list afterward.</p>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 8 }}>
          <button onClick={downloadTemplate} className="tap" style={{ display: "flex", alignItems: "center", gap: 7, background: "#fff", color: INK, border: `1px solid ${LINE_C}`, borderRadius: 9, padding: "10px 16px", fontSize: 14, fontWeight: 600, cursor: "pointer" }}><Download size={15} /> Download template</button>
          <button onClick={() => fileRef.current && fileRef.current.click()} className="tap" style={{ display: "flex", alignItems: "center", gap: 7, background: `linear-gradient(90deg, ${BTN_A}, ${BTN_B})`, color: "#fff", border: "none", borderRadius: 9, padding: "10px 16px", fontSize: 14, fontWeight: 600, cursor: "pointer" }}><FileSpreadsheet size={15} /> Choose Excel file</button>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={(e) => { const f = e.target.files && e.target.files[0]; if (f) onFile(f); }} />
        </div>
        {fileName && <div style={{ fontSize: 12.5, opacity: 0.6, marginBottom: 8 }}>Loaded: <b>{fileName}</b></div>}
        {done && <div style={{ marginTop: 10, background: CALL + "18", color: CALL, borderRadius: 8, padding: "10px 12px", fontSize: 13.5, fontWeight: 500 }}>{done}</div>}

        {rows && (
          <div style={{ marginTop: 14 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap", marginBottom: 12 }}>
              <span style={{ fontSize: 14, fontWeight: 600 }}>{valid.length} ready</span>
              {invalid.length > 0 && <span style={{ fontSize: 14, fontWeight: 600, color: "#B4453F" }}>{invalid.length} with issues</span>}
              <button onClick={doImport} disabled={busy || !valid.length} className="tap" style={{ marginLeft: "auto", background: valid.length && !busy ? INK : LINE_C, color: valid.length && !busy ? PAPER : "#8494A6", border: "none", borderRadius: 9, padding: "10px 18px", fontSize: 14, fontWeight: 600, cursor: valid.length && !busy ? "pointer" : "default" }}>{busy ? "Importing…" : `Import ${valid.length}`}</button>
            </div>
            {invalid.length > 0 && (
              <div style={{ background: "#FBECEB", border: "1px solid #E6C9C7", borderRadius: 10, padding: "12px 14px", marginBottom: 12 }}>
                {invalid.slice(0, 6).map((r) => <div key={r._row} style={{ fontSize: 12.5, color: "#8A3B36", marginBottom: 3 }}>Row {r._row}{r.company ? ` (${r.company})` : ""}: {r.errors.join("; ")}</div>)}
                {invalid.length > 6 && <div style={{ fontSize: 12.5, color: "#8A3B36" }}>…and {invalid.length - 6} more.</div>}
              </div>
            )}
            <div style={{ overflowX: "auto", border: `1px solid ${LINE_C}`, borderRadius: 10 }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5, minWidth: 560 }}>
                <thead><tr style={{ background: "#F1F5F9" }}>{["", "Company", "Contact", "Status", "Assign to"].map((h) => <th key={h} style={{ textAlign: "left", padding: "8px 10px", fontSize: 11, fontWeight: 700, textTransform: "uppercase", color: "#5A6B7B" }}>{h}</th>)}</tr></thead>
                <tbody>
                  {rows.slice(0, 50).map((r) => (
                    <tr key={r._row} style={{ borderTop: `1px solid ${LINE_C}`, background: r.errors.length ? "#FDF4F3" : "transparent" }}>
                      <td style={{ padding: "7px 10px" }}>{r.errors.length ? <X size={14} color="#B4453F" /> : <CheckCircle2 size={14} color="#189B72" />}</td>
                      <td style={{ padding: "7px 10px", fontWeight: 600 }}>{r.company || "—"}</td>
                      <td style={{ padding: "7px 10px" }}>{r.contact || "—"}</td>
                      <td style={{ padding: "7px 10px", textTransform: "capitalize" }}>{r.status}</td>
                      <td style={{ padding: "7px 10px" }}>{r.assigneeName}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ---- CRM: Companies list ----
function CompaniesList({ companies, entries, deals, onOpen }) {
  const [q, setQ] = useState("");
  // Only show companies the user actually has visible activity/deals for (matches scope),
  // plus any company with no rows yet that they can still see (RLS already filtered `companies`).
  const stats = useMemo(() => {
    const m = {};
    companies.forEach((c) => { m[c.id] = { calls: 0, emails: 0, appts: 0, activities: 0, deals: 0, pipeline: 0, last: "" }; });
    entries.forEach((e) => {
      const s = m[e.companyId]; if (!s) return;
      s.calls += e.calls || 0; s.emails += e.emails || 0; s.appts += e.appts || 0; s.activities += 1;
      if (!s.last || e.date > s.last) s.last = e.date;
    });
    deals.forEach((d) => {
      const s = m[d.companyId]; if (!s) return;
      s.deals += 1; if (d.stage !== "lost") s.pipeline += Number(d.value) || 0;
    });
    return m;
  }, [companies, entries, deals]);

  const rows = companies
    .filter((c) => c.name.toLowerCase().includes(q.trim().toLowerCase()))
    .sort((a, b) => a.name.localeCompare(b.name));

  return (
    <>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12, marginBottom: 20 }}>
        <div>
          <h1 style={{ fontFamily: "'Fraunces', serif", fontSize: 29, fontWeight: 600, margin: 0 }}>Companies</h1>
          <p style={{ margin: "4px 0 0", opacity: 0.55, fontSize: 14 }}>{rows.length} {rows.length === 1 ? "company" : "companies"} in your scope</p>
        </div>
        <div style={{ position: "relative" }}>
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search companies…"
            style={{ ...inputStyle, width: 260, marginBottom: 0 }} />
        </div>
      </div>

      <div style={{ background: CARD, border: `1px solid ${LINE_C}`, borderRadius: 14, overflow: "hidden" }}>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13.5, minWidth: 760 }}>
            <thead>
              <tr style={{ background: "#F1F5F9" }}>
                {["Company", "Activities", "Calls", "Emails", "Appts", "Deals", "Pipeline", "Last activity"].map((h) => (
                  <th key={h} style={{ textAlign: h === "Company" ? "left" : "center", padding: "11px 14px", fontSize: 11.5, fontWeight: 700, textTransform: "uppercase", color: "#5A6B7B", whiteSpace: "nowrap" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.length === 0 ? (
                <tr><td colSpan={8} style={{ padding: 28, textAlign: "center", opacity: 0.5 }}>No companies yet. They appear here as activity is logged.</td></tr>
              ) : rows.map((c) => {
                const s = stats[c.id] || {};
                return (
                  <tr key={c.id} className="tap" onClick={() => onOpen(c.id)}
                    style={{ borderTop: `1px solid ${LINE_C}`, cursor: "pointer" }}>
                    <td style={{ padding: "11px 14px", fontWeight: 600, color: EMAIL, whiteSpace: "nowrap" }}>{c.name}</td>
                    <td style={{ padding: "11px 14px", textAlign: "center" }}>{s.activities || 0}</td>
                    <td style={{ padding: "11px 14px", textAlign: "center" }}>{s.calls || 0}</td>
                    <td style={{ padding: "11px 14px", textAlign: "center" }}>{s.emails || 0}</td>
                    <td style={{ padding: "11px 14px", textAlign: "center" }}>{s.appts || 0}</td>
                    <td style={{ padding: "11px 14px", textAlign: "center" }}>{s.deals || 0}</td>
                    <td style={{ padding: "11px 14px", textAlign: "center" }}>{s.pipeline ? "$" + s.pipeline.toLocaleString() : "—"}</td>
                    <td style={{ padding: "11px 14px", textAlign: "center", whiteSpace: "nowrap", opacity: 0.7 }}>{s.last || "—"}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </>
  );
}

// ---- CRM: Company detail page (Overview, Activities, Deals, Contacts, Notes, Attachments) ----
function CompanyDetail({ companyId, companies, entries, deals, users, effectiveUser, saveDeals, visibleUserIds, onBack, onOpenCompany, refetch }) {
  const company = companies.find((c) => c.id === companyId);
  const [tab, setTab] = useState("overview");
  const [contacts, setContacts] = useState([]);
  const [notes, setNotes] = useState([]);
  const [attachments, setAttachments] = useState([]);
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(company || {});
  const [busy, setBusy] = useState("");
  const [dealModal, setDealModal] = useState(false);
  const canEditDeals = ["bdr", "sales", "admin", "management"].includes(effectiveUser.role);
  const repUsers = users.filter((u) => (!visibleUserIds || visibleUserIds.includes(u.id)) && (u.role === "bdr" || u.role === "sales"));

  const saveNewDeal = (data) => saveDeals(async () => {
    // Force this deal onto the current company, in the appointment-free default flow.
    await api.upsertDeal({ ...data, companyId, company: company.name, ownerId: data.ownerId || effectiveUser.id, stageChangedAt: new Date().toISOString() });
  }).then(() => { setDealModal(false); });

  const companyEntries = entries.filter((e) => e.companyId === companyId).sort((a, b) => (b.date < a.date ? -1 : 1));
  const companyDeals = deals.filter((d) => d.companyId === companyId);
  const nameOf = (id) => { const u = users.find((x) => x.id === id); return u ? u.name : ""; };

  const loadSub = async () => {
    try {
      const [cs, ns, as] = await Promise.all([api.listContacts(companyId), api.listCompanyNotes(companyId), api.listAttachments(companyId)]);
      setContacts(cs); setNotes(ns); setAttachments(as);
    } catch (e) { /* surfaced per-action */ }
  };
  useEffect(() => { setDraft(company || {}); loadSub(); /* eslint-disable-next-line */ }, [companyId]);

  if (!company) return <Empty msg="Company not found in your scope." />;

  const totals = companyEntries.reduce((a, e) => ({ calls: a.calls + (e.calls || 0), emails: a.emails + (e.emails || 0), appts: a.appts + (e.appts || 0) }), { calls: 0, emails: 0, appts: 0 });
  const pipeline = companyDeals.filter((d) => d.stage !== "lost").reduce((s, d) => s + (Number(d.value) || 0), 0);

  const saveInfo = async () => {
    setBusy("Saving…");
    try { await api.updateCompany(companyId, { name: draft.name, website: draft.website, phone: draft.phone, address: draft.address, ban: draft.ban, fan: draft.fan, notes: draft.notes, ownerId: draft.ownerId, secondaryOwnerId: draft.secondaryOwnerId }); await refetch(); setEditing(false); }
    finally { setBusy(""); }
  };

  const subTabs = [
    ["overview", "Overview", Building2],
    ["activities", `Activities (${companyEntries.length})`, TrendingUp],
    ["deals", `Deals (${companyDeals.length})`, Briefcase],
    ["contacts", `Contacts (${contacts.length})`, Users],
    ["notes", `Notes (${notes.length})`, Pencil],
    ["attachments", `Files (${attachments.length})`, FileSpreadsheet],
  ];

  return (
    <>
      <button onClick={onBack} className="tap" style={{ display: "flex", alignItems: "center", gap: 6, background: "transparent", border: "none", color: EMAIL, fontSize: 13.5, fontWeight: 600, cursor: "pointer", marginBottom: 12, padding: 0 }}>
        <ChevronLeft size={16} /> All companies
      </button>

      <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", flexWrap: "wrap", gap: 12, marginBottom: 18 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ width: 46, height: 46, borderRadius: 12, background: `linear-gradient(135deg, ${BTN_A}, ${BTN_B})`, display: "grid", placeItems: "center", color: "#fff", fontFamily: "'Fraunces', serif", fontSize: 22, fontWeight: 600 }}>{(company.name[0] || "?").toUpperCase()}</div>
          <div>
            <h1 style={{ fontFamily: "'Fraunces', serif", fontSize: 28, fontWeight: 600, margin: 0 }}>{company.name}</h1>
            <div style={{ fontSize: 13, opacity: 0.6 }}>{company.ban ? `BAN ${company.ban}` : "No BAN set"}</div>
          </div>
        </div>
        <div style={{ display: "flex", gap: 18, flexWrap: "wrap" }}>
          <MiniStat label="Calls" value={totals.calls} color={CALL} />
          <MiniStat label="Emails" value={totals.emails} color={EMAIL} />
          <MiniStat label="Appts" value={totals.appts} color={APPT} />
          <MiniStat label="Pipeline" value={pipeline ? "$" + pipeline.toLocaleString() : "$0"} color={INK} />
        </div>
      </div>

      {/* Sub-tabs */}
      <div style={{ display: "flex", gap: 6, flexWrap: "wrap", borderBottom: `1px solid ${LINE_C}`, marginBottom: 18 }}>
        {subTabs.map(([id, label, Icon]) => (
          <button key={id} onClick={() => setTab(id)} className="tap"
            style={{ display: "flex", alignItems: "center", gap: 6, border: "none", background: "transparent", cursor: "pointer", padding: "9px 12px", fontSize: 13.5, fontWeight: 600,
              color: tab === id ? INK : "#8494A6", borderBottom: tab === id ? `2px solid ${CYAN}` : "2px solid transparent", marginBottom: -1 }}>
            <Icon size={15} /> {label}
          </button>
        ))}
      </div>

      {busy && <div style={{ fontSize: 13, color: CALL, marginBottom: 10 }}>{busy}</div>}

      {/* OVERVIEW */}
      {tab === "overview" && (
        <Panel title="Company information" icon={Building2}>
          {editing ? (
            <div style={{ display: "grid", gap: 12 }}>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                <Field label="Company name"><input value={draft.name || ""} onChange={(e) => setDraft({ ...draft, name: e.target.value })} style={inputStyle} /></Field>
                <Field label="Website"><input value={draft.website || ""} onChange={(e) => setDraft({ ...draft, website: e.target.value })} style={inputStyle} /></Field>
                <Field label="Phone"><input value={draft.phone || ""} onChange={(e) => setDraft({ ...draft, phone: formatPhone(e.target.value) })} style={inputStyle} /></Field>
                <Field label="BAN"><input value={draft.ban || ""} onChange={(e) => setDraft({ ...draft, ban: e.target.value })} style={inputStyle} /></Field>
                <Field label="FAN"><input value={draft.fan || ""} onChange={(e) => setDraft({ ...draft, fan: e.target.value })} style={inputStyle} /></Field>
                <Field label="Address"><input value={draft.address || ""} onChange={(e) => setDraft({ ...draft, address: e.target.value })} style={inputStyle} /></Field>
              </div>
              {(() => {
                const isMgmt = effectiveUser.role === "admin" || effectiveUser.role === "management";
                // Permission gate: once an owner is set, only admin/management may change it.
                const ownerLocked = !!company.ownerId && !isMgmt;
                const ownerPool = users.filter((u) => u.role === "bdr" || u.role === "sales" || u.role === "management" || u.role === "admin");
                return (
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                    <Field label="Account owner">
                      <div style={{ position: "relative" }}>
                        <select value={draft.ownerId || ""} disabled={ownerLocked}
                          onChange={(e) => setDraft({ ...draft, ownerId: e.target.value })}
                          style={{ ...inputStyle, appearance: "none", cursor: ownerLocked ? "not-allowed" : "pointer", opacity: ownerLocked ? 0.6 : 1 }}>
                          <option value="">Unassigned</option>
                          {ownerPool.map((u) => <option key={u.id} value={u.id}>{u.name}</option>)}
                        </select>
                        <ChevronDown size={15} style={{ position: "absolute", right: 12, top: 12, pointerEvents: "none", opacity: 0.5 }} />
                      </div>
                    </Field>
                    <Field label="Shared with (optional)">
                      <div style={{ position: "relative" }}>
                        <select value={draft.secondaryOwnerId || ""} disabled={ownerLocked}
                          onChange={(e) => setDraft({ ...draft, secondaryOwnerId: e.target.value })}
                          style={{ ...inputStyle, appearance: "none", cursor: ownerLocked ? "not-allowed" : "pointer", opacity: ownerLocked ? 0.6 : 1 }}>
                          <option value="">No one</option>
                          {ownerPool.filter((u) => u.id !== draft.ownerId).map((u) => <option key={u.id} value={u.id}>{u.name}</option>)}
                        </select>
                        <ChevronDown size={15} style={{ position: "absolute", right: 12, top: 12, pointerEvents: "none", opacity: 0.5 }} />
                      </div>
                    </Field>
                    {ownerLocked && <p style={{ gridColumn: "1 / -1", fontSize: 12, opacity: 0.55, margin: 0 }}>This account already has an owner. Ask a manager or admin to reassign it.</p>}
                  </div>
                );
              })()}
              <Field label="Description / overview"><textarea rows={3} value={draft.notes || ""} onChange={(e) => setDraft({ ...draft, notes: e.target.value })} style={{ ...inputStyle, resize: "vertical" }} /></Field>
              <div style={{ display: "flex", gap: 8 }}>
                <button onClick={saveInfo} className="tap" style={{ background: INK, color: PAPER, border: "none", borderRadius: 8, padding: "9px 18px", fontSize: 14, fontWeight: 500, cursor: "pointer" }}>Save</button>
                <button onClick={() => { setDraft(company); setEditing(false); }} className="tap" style={{ background: "transparent", border: `1px solid ${LINE_C}`, borderRadius: 8, padding: "9px 14px", fontSize: 14, cursor: "pointer" }}>Cancel</button>
              </div>
            </div>
          ) : (
            <div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "14px 24px", marginBottom: 14 }}>
                <InfoRow label="Website" value={company.website} link />
                <InfoRow label="Phone" value={company.phone} />
                <InfoRow label="BAN" value={company.ban} />
                <InfoRow label="FAN" value={company.fan} />
                <InfoRow label="Address" value={company.address} />
                <InfoRow label="Account owner" value={company.ownerId ? nameOf(company.ownerId) : ""} />
                {company.secondaryOwnerId && <InfoRow label="Shared with" value={nameOf(company.secondaryOwnerId)} />}
              </div>
              {company.notes && <p style={{ fontSize: 14, lineHeight: 1.6, margin: "0 0 14px", opacity: 0.85 }}>{company.notes}</p>}
              <button onClick={() => { setDraft(company); setEditing(true); }} className="tap" style={{ background: "transparent", border: "none", color: EMAIL, fontSize: 13.5, fontWeight: 600, cursor: "pointer", padding: 0 }}>Edit company info</button>
            </div>
          )}
        </Panel>
      )}

      {tab === "overview" && (() => {
        // Interleave activities and deals into one chronological account story.
        const events = [];
        companyEntries.forEach((e) => {
          const bits = [];
          if (e.calls) bits.push(`${e.calls} call${e.calls > 1 ? "s" : ""}`);
          if (e.emails) bits.push(`${e.emails} email${e.emails > 1 ? "s" : ""}`);
          if (e.appts) bits.push(`${e.appts} appt${e.appts > 1 ? "s" : ""}`);
          events.push({ date: e.date, kind: "activity", who: nameOf(e.userId), summary: bits.join(" · ") || "Activity logged", detail: e.notes });
        });
        companyDeals.forEach((d) => {
          events.push({ date: (d.createdAt || "").slice(0, 10), kind: "deal", who: nameOf(d.ownerId), summary: `Deal · ${STAGE[d.stage]?.label || d.stage}${d.value ? ` · ${fmtMoney(d.value)}` : ""}`, detail: d.stage === "lost" && d.lostReason ? `Lost: ${d.lostReason}` : d.notes });
        });
        const sorted = events.filter((e) => e.date).sort((a, b) => b.date.localeCompare(a.date)).slice(0, 20);
        if (sorted.length === 0) return null;
        return (
          <Panel title="Account timeline" icon={Clock} style={{ marginTop: 16 }}>
            <div style={{ display: "grid", gap: 2 }}>
              {sorted.map((ev, i) => (
                <div key={i} style={{ display: "flex", gap: 12, padding: "10px 0", borderTop: i ? `1px solid ${LINE_C}` : "none" }}>
                  <span style={{ width: 8, height: 8, borderRadius: "50%", marginTop: 5, flexShrink: 0, background: ev.kind === "deal" ? EMAIL : CALL }} />
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontSize: 13.5, fontWeight: 600 }}>{ev.summary}</div>
                    {ev.detail && <div style={{ fontSize: 12.5, opacity: 0.65, marginTop: 2 }}>{ev.detail}</div>}
                    <div style={{ fontSize: 11.5, opacity: 0.5, marginTop: 2 }}>{new Date(ev.date + "T00:00").toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" })} · {ev.who}</div>
                  </div>
                </div>
              ))}
            </div>
          </Panel>
        );
      })()}
      {tab === "activities" && <ActivityTable entries={companyEntries} users={users} liveUser={effectiveUser} />}

      {/* DEALS */}
      {tab === "deals" && (
        <Panel title="Deals" icon={Briefcase} action={canEditDeals ? (
          <button onClick={() => setDealModal(true)} className="tap" style={{ background: INK, color: PAPER, border: "none", borderRadius: 8, padding: "8px 14px", fontSize: 13, fontWeight: 600, cursor: "pointer", display: "flex", alignItems: "center", gap: 6 }}>
            <Plus size={14} /> Add deal
          </button>
        ) : null}>
          {companyDeals.length === 0 ? <Empty msg="No deals for this company yet." /> : (
            <div style={{ display: "grid", gap: 8 }}>
              {companyDeals.map((d) => (
                <div key={d.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: "#fff", border: `1px solid ${LINE_C}`, borderRadius: 10, padding: "12px 14px" }}>
                  <div>
                    <div style={{ fontWeight: 600, fontSize: 14 }}>{d.contact || company.name}</div>
                    <div style={{ fontSize: 12.5, opacity: 0.6, textTransform: "capitalize" }}>{d.stage} · {nameOf(d.ownerId)}</div>
                  </div>
                  <div style={{ fontFamily: "'Fraunces', serif", fontSize: 18, fontWeight: 600 }}>${(Number(d.value) || 0).toLocaleString()}</div>
                </div>
              ))}
            </div>
          )}
        </Panel>
      )}
      {dealModal && (
        <DealModal deal={{ company: company.name, companyId }} onSave={saveNewDeal} onClose={() => setDealModal(false)}
          liveUser={effectiveUser} salesReps={users.filter((u) => u.role === "sales")} assignableOwners={repUsers} dealUsers={users} />
      )}

      {/* CONTACTS */}
      {tab === "contacts" && <ContactsSection companyId={companyId} contacts={contacts} reload={loadSub} users={users} effectiveUser={effectiveUser} />}

      {/* NOTES */}
      {tab === "notes" && <NotesSection companyId={companyId} notes={notes} nameOf={nameOf} effectiveUser={effectiveUser} reload={loadSub} />}

      {/* ATTACHMENTS */}
      {tab === "attachments" && <AttachmentsSection companyId={companyId} attachments={attachments} nameOf={nameOf} reload={loadSub} />}
    </>
  );
}

function MiniStat({ label, value, color }) {
  return (
    <div style={{ textAlign: "center" }}>
      <div style={{ fontFamily: "'Fraunces', serif", fontSize: 22, fontWeight: 600, color }}>{value}</div>
      <div style={{ fontSize: 11, opacity: 0.55, textTransform: "uppercase", letterSpacing: 0.5 }}>{label}</div>
    </div>
  );
}
function InfoRow({ label, value, link }) {
  return (
    <div>
      <div style={{ fontSize: 11.5, fontWeight: 700, textTransform: "uppercase", color: "#8494A6", marginBottom: 2 }}>{label}</div>
      {value ? (link ? <a href={value.startsWith("http") ? value : `https://${value}`} target="_blank" rel="noreferrer" style={{ color: EMAIL, fontSize: 14 }}>{value}</a> : <div style={{ fontSize: 14 }}>{value}</div>) : <div style={{ fontSize: 14, opacity: 0.4 }}>—</div>}
    </div>
  );
}

function ContactsSection({ companyId, contacts, reload, users, effectiveUser }) {
  const [modal, setModal] = useState(null); // contact being edited, or {} for new
  const [openNotes, setOpenNotes] = useState(() => new Set());
  const save = async (c) => { await api.saveContact(companyId, c); setModal(null); await reload(); };
  const del = async (id) => { if (confirm("Delete this contact?")) { await api.deleteContact(id); await reload(); } };
  const toggleNotes = (id) => setOpenNotes((p) => { const n = new Set(p); n.has(id) ? n.delete(id) : n.add(id); return n; });
  return (
    <Panel title="Contacts" icon={Users} action={<button onClick={() => setModal({})} className="tap" style={{ background: INK, color: PAPER, border: "none", borderRadius: 8, padding: "8px 14px", fontSize: 13, fontWeight: 600, cursor: "pointer", display: "flex", alignItems: "center", gap: 6 }}><UserPlus size={14} /> Add contact</button>}>
      {contacts.length === 0 ? <Empty msg="No contacts yet." /> : (
        <div style={{ display: "grid", gap: 8 }}>
          {contacts.map((c) => (
            <div key={c.id} style={{ background: "#fff", border: `1px solid ${LINE_C}`, borderRadius: 10, padding: "12px 14px" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div style={{ minWidth: 0 }}>
                  <div style={{ fontWeight: 600, fontSize: 14 }}>{c.name}{c.title ? <span style={{ fontWeight: 400, opacity: 0.6 }}> · {c.title}</span> : null}</div>
                  <div style={{ fontSize: 12.5, opacity: 0.65 }}>{[c.email, c.phone].filter(Boolean).join(" · ") || "No contact details"}</div>
                </div>
                <div style={{ display: "flex", gap: 4 }}>
                  <button onClick={() => toggleNotes(c.id)} className="tap" style={{ ...iconBtn, color: openNotes.has(c.id) ? EMAIL : INK }} title="Notes / history"><FileSpreadsheet size={14} /></button>
                  <button onClick={() => setModal(c)} className="tap" style={iconBtn} title="Edit"><Pencil size={14} /></button>
                  <button onClick={() => del(c.id)} className="tap" style={iconBtn} title="Delete"><Trash2 size={14} /></button>
                </div>
              </div>
              {openNotes.has(c.id) && (
                <div style={{ marginTop: 12, borderTop: `1px solid ${LINE_C}`, paddingTop: 12 }}>
                  <EntityNotes entityType="contact" entityId={c.id} users={users} effectiveUser={effectiveUser} compact />
                </div>
              )}
            </div>
          ))}
        </div>
      )}
      {modal && <ContactModal contact={modal.id ? modal : null} onSave={save} onClose={() => setModal(null)} />}
    </Panel>
  );
}
function ContactModal({ contact, onSave, onClose }) {
  const [f, setF] = useState(contact || { name: "", title: "", phone: "", email: "", notes: "" });
  const [err, setErr] = useState("");
  const submit = () => { if (!f.name.trim()) { setErr("Name is required."); return; } onSave(contact ? { ...f, id: contact.id } : f); };
  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(11,42,74,.5)", display: "grid", placeItems: "center", padding: 20, zIndex: 50 }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: PAPER, borderRadius: 16, padding: 24, width: "100%", maxWidth: 440 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <h3 style={{ fontFamily: "'Fraunces', serif", fontSize: 20, fontWeight: 600, margin: 0 }}>{contact ? "Edit contact" : "New contact"}</h3>
          <button onClick={onClose} className="tap" style={{ background: "transparent", border: "none", cursor: "pointer" }}><X size={20} /></button>
        </div>
        <Field label="Name"><input value={f.name} onChange={(e) => setF({ ...f, name: e.target.value })} style={inputStyle} autoFocus /></Field>
        <Field label="Title"><input value={f.title} onChange={(e) => setF({ ...f, title: e.target.value })} style={inputStyle} /></Field>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          <Field label="Email"><input value={f.email} onChange={(e) => setF({ ...f, email: e.target.value })} style={inputStyle} /></Field>
          <Field label="Phone"><input value={f.phone} onChange={(e) => setF({ ...f, phone: formatPhone(e.target.value) })} style={inputStyle} /></Field>
        </div>
        {err && <div style={{ color: "#B4453F", fontSize: 13, marginBottom: 10 }}>{err}</div>}
        <button onClick={submit} className="tap" style={{ width: "100%", background: INK, color: PAPER, border: "none", borderRadius: 8, padding: 12, fontSize: 14, fontWeight: 600, cursor: "pointer" }}>Save contact</button>
      </div>
    </div>
  );
}

function NotesSection({ companyId, notes, nameOf, effectiveUser, reload }) {
  const [text, setText] = useState("");
  const [staged, setStaged] = useState([]); // files chosen before the note is saved
  const [busy, setBusy] = useState(false);
  const composeFileRef = React.useRef(null);

  const add = async () => {
    if (!text.trim() && staged.length === 0) return;
    setBusy(true);
    try {
      // Create the note first so we have an id, then upload any staged files onto it.
      const note = await api.addCompanyNote(companyId, text.trim() || "(file attached)");
      for (const f of staged) await api.uploadNoteAttachment(companyId, note.id, f);
      setText(""); setStaged([]);
      await reload();
    } finally { setBusy(false); }
  };
  const del = async (id) => { if (confirm("Delete this note and its attachments?")) { await api.deleteCompanyNote(id); await reload(); } };

  return (
    <Panel title="Notes" icon={Pencil}>
      <div style={{ marginBottom: 16 }}>
        <div style={{ display: "flex", gap: 8 }}>
          <textarea rows={2} value={text} onChange={(e) => setText(e.target.value)} placeholder="Add a note…" style={{ ...inputStyle, marginBottom: 0, resize: "vertical" }} />
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            <button onClick={add} disabled={busy || (!text.trim() && staged.length === 0)} className="tap" style={{ background: (text.trim() || staged.length) && !busy ? INK : LINE_C, color: (text.trim() || staged.length) && !busy ? PAPER : "#8494A6", border: "none", borderRadius: 8, padding: "10px 18px", fontSize: 14, fontWeight: 600, cursor: (text.trim() || staged.length) ? "pointer" : "default", whiteSpace: "nowrap" }}>{busy ? "Saving…" : "Add"}</button>
            <input ref={composeFileRef} type="file" style={{ display: "none" }} onChange={(e) => { if (e.target.files[0]) setStaged((s) => [...s, e.target.files[0]]); e.target.value = ""; }} />
            <button onClick={() => composeFileRef.current?.click()} className="tap" title="Attach a file to this note" style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 5, background: "transparent", border: `1px solid ${LINE_C}`, color: EMAIL, borderRadius: 8, padding: "8px 12px", fontSize: 12.5, fontWeight: 600, cursor: "pointer", whiteSpace: "nowrap" }}><Plus size={13} /> Attach</button>
          </div>
        </div>
        {staged.length > 0 && (
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginTop: 8 }}>
            {staged.map((f, i) => (
              <span key={i} style={{ display: "flex", alignItems: "center", gap: 6, background: "#F1F5F9", border: `1px solid ${LINE_C}`, borderRadius: 7, padding: "4px 9px", fontSize: 12.5 }}>
                <FileSpreadsheet size={12} color={EMAIL} /> {f.name}
                <button onClick={() => setStaged((s) => s.filter((_, j) => j !== i))} className="tap" style={{ background: "transparent", border: "none", cursor: "pointer", opacity: 0.5, padding: 0, display: "flex" }}><X size={12} /></button>
              </span>
            ))}
          </div>
        )}
      </div>
      {notes.length === 0 ? <Empty msg="No notes yet." /> : (
        <div style={{ display: "grid", gap: 8 }}>
          {notes.map((n) => (
            <NoteRow key={n.id} note={n} companyId={companyId} nameOf={nameOf} effectiveUser={effectiveUser} onDelete={del} />
          ))}
        </div>
      )}
    </Panel>
  );
}

function NoteRow({ note, companyId, nameOf, effectiveUser, onDelete }) {
  const [files, setFiles] = useState(null);
  const [busy, setBusy] = useState("");
  const fileRef = React.useRef(null);
  const load = async () => { try { setFiles(await api.listNoteAttachments(note.id)); } catch { setFiles([]); } };
  useEffect(() => { load(); /* eslint-disable-next-line */ }, [note.id]);

  const onFile = async (file) => {
    if (!file) return;
    setBusy("Uploading…");
    try { await api.uploadNoteAttachment(companyId, note.id, file); await load(); }
    catch (e) { setBusy("Upload failed"); setTimeout(() => setBusy(""), 2000); return; }
    setBusy("");
  };
  const open = async (att) => { try { const url = await api.attachmentUrl(att.storagePath); window.open(url, "_blank"); } catch {} };
  const removeFile = async (att) => { if (confirm(`Delete ${att.fileName}?`)) { await api.deleteAttachment(att); await load(); } };

  return (
    <div style={{ background: "#fff", border: `1px solid ${LINE_C}`, borderRadius: 10, padding: "12px 14px" }}>
      <div style={{ fontSize: 14, lineHeight: 1.5, whiteSpace: "pre-wrap" }}>{note.body}</div>
      {files && files.length > 0 && (
        <div style={{ display: "grid", gap: 5, marginTop: 10 }}>
          {files.map((a) => (
            <div key={a.id} style={{ display: "flex", alignItems: "center", gap: 8, background: "#F1F5F9", borderRadius: 7, padding: "6px 10px" }}>
              <FileSpreadsheet size={13} color={EMAIL} />
              <button onClick={() => open(a)} className="tap" style={{ flex: 1, textAlign: "left", background: "transparent", border: "none", color: EMAIL, fontSize: 12.5, fontWeight: 500, cursor: "pointer", padding: 0, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{a.fileName}</button>
              {a.uploaderId === effectiveUser.id && <button onClick={() => removeFile(a)} className="tap" style={{ background: "transparent", border: "none", cursor: "pointer", opacity: 0.4 }}><Trash2 size={12} /></button>}
            </div>
          ))}
        </div>
      )}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 10 }}>
        <span style={{ fontSize: 12, opacity: 0.5 }}>{nameOf(note.authorId) || "Someone"} · {new Date(note.createdAt).toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" })}</span>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <input ref={fileRef} type="file" style={{ display: "none" }} onChange={(e) => { onFile(e.target.files[0]); e.target.value = ""; }} />
          <button onClick={() => fileRef.current?.click()} className="tap" style={{ display: "flex", alignItems: "center", gap: 5, background: "transparent", border: "none", color: EMAIL, fontSize: 12, fontWeight: 600, cursor: "pointer", opacity: busy ? 0.6 : 1 }}>
            <Plus size={12} /> {busy || "Attach file"}
          </button>
          {note.authorId === effectiveUser.id && <button onClick={() => onDelete(note.id)} className="tap" style={{ background: "transparent", border: "none", cursor: "pointer", opacity: 0.4 }}><Trash2 size={13} /></button>}
        </div>
      </div>
    </div>
  );
}

// Timestamped running notes for a contact or a deal. Loads its own data.
function EntityNotes({ entityType, entityId, users, effectiveUser, compact }) {
  const [notes, setNotes] = useState(null);
  const [text, setText] = useState("");
  const [busy, setBusy] = useState(false);
  const nameOf = (id) => { const u = (users || []).find((x) => x.id === id); return u ? u.name : "Someone"; };
  const load = async () => { try { setNotes(await api.listEntityNotes(entityType, entityId)); } catch { setNotes([]); } };
  useEffect(() => { load(); /* eslint-disable-next-line */ }, [entityType, entityId]);
  const add = async () => { if (!text.trim()) return; setBusy(true); try { await api.addEntityNote(entityType, entityId, text.trim()); setText(""); await load(); } finally { setBusy(false); } };
  const del = async (id) => { if (confirm("Delete this note?")) { await api.deleteEntityNote(id); await load(); } };
  return (
    <div style={compact ? { marginTop: 8 } : {}}>
      <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
        <textarea rows={2} value={text} onChange={(e) => setText(e.target.value)} placeholder="Add a note…" style={{ ...inputStyle, marginBottom: 0, resize: "vertical" }} />
        <button onClick={add} disabled={busy || !text.trim()} className="tap" style={{ background: text.trim() && !busy ? INK : LINE_C, color: text.trim() && !busy ? PAPER : "#8494A6", border: "none", borderRadius: 8, padding: "0 18px", fontSize: 14, fontWeight: 600, cursor: text.trim() ? "pointer" : "default", whiteSpace: "nowrap" }}>Add</button>
      </div>
      {notes === null ? <div style={{ fontSize: 13, opacity: 0.4, padding: 8 }}>Loading…</div>
        : notes.length === 0 ? <Empty msg="No notes yet." /> : (
        <div style={{ display: "grid", gap: 8 }}>
          {notes.map((n) => (
            <div key={n.id} style={{ background: "#fff", border: `1px solid ${LINE_C}`, borderRadius: 10, padding: "12px 14px" }}>
              <div style={{ fontSize: 14, lineHeight: 1.5, whiteSpace: "pre-wrap" }}>{n.body}</div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 8 }}>
                <span style={{ fontSize: 12, opacity: 0.5 }}>{nameOf(n.authorId)} · {new Date(n.createdAt).toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" })}</span>
                {n.authorId === effectiveUser.id && <button onClick={() => del(n.id)} className="tap" style={{ background: "transparent", border: "none", cursor: "pointer", opacity: 0.4 }}><Trash2 size={13} /></button>}
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function AttachmentsSection({ companyId, attachments, nameOf, reload }) {
  const [busy, setBusy] = useState("");
  const fileRef = React.useRef(null);
  const onFile = async (file) => { if (!file) return; setBusy("Uploading…"); try { await api.uploadAttachment(companyId, file); await reload(); } catch (e) { setBusy("Upload failed: " + (e.message || "")); return; } setBusy(""); };
  const open = async (a) => { try { const url = await api.attachmentUrl(a.storagePath); window.open(url, "_blank"); } catch (e) { alert("Couldn't open file."); } };
  const del = async (a) => { if (confirm("Delete this file?")) { await api.deleteAttachment(a); await reload(); } };
  const fmtSize = (b) => b > 1e6 ? (b / 1e6).toFixed(1) + " MB" : b > 1e3 ? Math.round(b / 1e3) + " KB" : b + " B";
  return (
    <Panel title="Attachments" icon={FileSpreadsheet} action={
      <>
        <button onClick={() => fileRef.current && fileRef.current.click()} className="tap" style={{ background: INK, color: PAPER, border: "none", borderRadius: 8, padding: "8px 14px", fontSize: 13, fontWeight: 600, cursor: "pointer", display: "flex", alignItems: "center", gap: 6 }}><Plus size={14} /> Upload file</button>
        <input ref={fileRef} type="file" style={{ display: "none" }} onChange={(e) => { const f = e.target.files && e.target.files[0]; if (f) onFile(f); e.target.value = ""; }} />
      </>
    }>
      {busy && <div style={{ fontSize: 13, color: CALL, marginBottom: 10 }}>{busy}</div>}
      {attachments.length === 0 ? <Empty msg="No files uploaded yet." /> : (
        <div style={{ display: "grid", gap: 8 }}>
          {attachments.map((a) => (
            <div key={a.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: "#fff", border: `1px solid ${LINE_C}`, borderRadius: 10, padding: "12px 14px" }}>
              <div style={{ minWidth: 0 }}>
                <button onClick={() => open(a)} className="tap" style={{ background: "transparent", border: "none", color: EMAIL, fontWeight: 600, fontSize: 14, cursor: "pointer", padding: 0, textAlign: "left" }}>{a.fileName}</button>
                <div style={{ fontSize: 12, opacity: 0.55 }}>{fmtSize(a.sizeBytes)} · {nameOf(a.uploaderId) || "Someone"} · {new Date(a.createdAt).toLocaleDateString("en-US", { month: "short", day: "numeric" })}</div>
              </div>
              <div style={{ display: "flex", gap: 4 }}>
                <button onClick={() => open(a)} className="tap" style={iconBtn} title="Download"><Download size={14} /></button>
                <button onClick={() => del(a)} className="tap" style={iconBtn} title="Delete"><Trash2 size={14} /></button>
              </div>
            </div>
          ))}
        </div>
      )}
    </Panel>
  );
}

// Spreadsheet-style table of activity records, with Excel export.
function ActivityTable({ entries, users, liveUser, compact, onOpenCompany }) {
  const [sortKey, setSortKey] = useState("date");
  const [sortDir, setSortDir] = useState("desc");
  const nameOf = (id) => { const u = (users || []).find((x) => x.id === id); return u ? u.name : ""; };
  const repOf = (e) => e.taggedRepId ? nameOf(e.taggedRepId) : "Self-generated";

  const rows = useMemo(() => {
    const r = [...entries];
    r.sort((a, b) => {
      let av, bv;
      switch (sortKey) {
        case "company": av = (a.company || "").toLowerCase(); bv = (b.company || "").toLowerCase(); break;
        case "rep": av = repOf(a).toLowerCase(); bv = repOf(b).toLowerCase(); break;
        case "logged": av = nameOf(a.userId).toLowerCase(); bv = nameOf(b.userId).toLowerCase(); break;
        case "carrier": av = (a.carrierRep || "").toLowerCase(); bv = (b.carrierRep || "").toLowerCase(); break;
        case "calls": av = a.calls || 0; bv = b.calls || 0; break;
        case "emails": av = a.emails || 0; bv = b.emails || 0; break;
        case "appts": av = a.appts || 0; bv = b.appts || 0; break;
        default: av = a.date || ""; bv = b.date || "";
      }
      if (av < bv) return sortDir === "asc" ? -1 : 1;
      if (av > bv) return sortDir === "asc" ? 1 : -1;
      return (a.id < b.id ? 1 : -1);
    });
    return r;
  }, [entries, sortKey, sortDir, users]);

  const setSort = (k) => { if (sortKey === k) setSortDir((d) => (d === "asc" ? "desc" : "asc")); else { setSortKey(k); setSortDir(k === "date" ? "desc" : "asc"); } };

  const exportXlsx = () => {
    const data = rows.map((e) => ({
      Date: e.date,
      Company: e.company || "",
      BAN: e.ban || "",
      FAN: e.fan || "",
      Contact: e.contact || "",
      Phone: e.phone || "",
      Email: e.email || "",
      Calls: e.calls || 0,
      Emails: e.emails || 0,
      Appointments: e.appts || 0,
      "Logged by": nameOf(e.userId),
      "Tellemica Sales Rep": repOf(e),
      "Carrier Rep": e.carrierRep || "",
      Notes: e.notes || "",
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    ws["!cols"] = [
      { wch: 11 }, { wch: 24 }, { wch: 14 }, { wch: 14 }, { wch: 20 }, { wch: 16 }, { wch: 24 },
      { wch: 7 }, { wch: 8 }, { wch: 13 }, { wch: 18 }, { wch: 18 }, { wch: 18 }, { wch: 40 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Activity");
    const stamp = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `tellemica-activity-${stamp}.xlsx`);
  };

  // Columns always visible in the collapsed row. Everything else lives in the
  // expand panel so nothing gets pushed off-screen. "rep" hides on narrow screens.
  const headCols = [
    ["date", "Date", false],
    ["company", "Company", false],
    ["calls", "Calls", true],
    ["emails", "Emails", true],
    ["appts", "Appts", true],
    ["logged", "Logged by", false],
    ["rep", "Tellemica Sales Rep", false],
  ];

  const [open, setOpen] = useState(() => new Set());
  const toggle = (id) => setOpen((prev) => { const n = new Set(prev); n.has(id) ? n.delete(id) : n.add(id); return n; });

  const cellPad = "11px 14px";
  const kvLabel = { fontSize: 10.5, fontWeight: 700, letterSpacing: 0.4, textTransform: "uppercase", color: "#8494A6" };
  const kvValue = { fontSize: 13.5 };

  return (
    <div style={{ background: CARD, border: `1px solid ${LINE_C}`, borderRadius: 14, overflow: "hidden" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "16px 18px", borderBottom: `1px solid ${LINE_C}`, flexWrap: "wrap", gap: 10 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <Table2 size={17} color={INK} />
          <span style={{ fontSize: 15, fontWeight: 600 }}>Activity records</span>
          <span style={{ fontSize: 12.5, opacity: 0.5 }}>({rows.length})</span>
        </div>
        <button onClick={exportXlsx} disabled={rows.length === 0} className="tap"
          style={{ display: "flex", alignItems: "center", gap: 7, background: rows.length ? INK : LINE_C, color: rows.length ? PAPER : "#8494A6", border: "none", borderRadius: 9, padding: "9px 14px", fontSize: 13.5, fontWeight: 600, cursor: rows.length ? "pointer" : "default" }}>
          <Download size={15} /> Export to Excel
        </button>
      </div>
      <div style={{ overflowX: "auto", maxHeight: compact ? 420 : "none" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F1F5F9", position: "sticky", top: 0, zIndex: 1 }}>
              <th style={{ width: 34, borderBottom: `1px solid ${LINE_C}` }} />
              {headCols.map(([k, label, isNum]) => (
                <th key={k} onClick={() => setSort(k)}
                  style={{ textAlign: isNum ? "center" : "left", padding: "10px 12px", fontWeight: 700, fontSize: 11.5, letterSpacing: 0.4, textTransform: "uppercase", color: "#5A6B7B", whiteSpace: "nowrap", cursor: "pointer", borderBottom: `1px solid ${LINE_C}` }}>
                  {label}{sortKey === k ? (sortDir === "asc" ? " ▲" : " ▼") : ""}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 ? (
              <tr><td colSpan={headCols.length + 1} style={{ padding: 28, textAlign: "center", opacity: 0.5 }}>No activity records yet.</td></tr>
            ) : rows.map((e) => {
              const isOpen = open.has(e.id);
              return (
                <React.Fragment key={e.id}>
                  <tr onClick={() => toggle(e.id)} className="tap"
                    style={{ borderBottom: isOpen ? "none" : `1px solid ${LINE_C}`, cursor: "pointer", background: isOpen ? "#F8FAFC" : "transparent" }}>
                    <td style={{ padding: cellPad, textAlign: "center" }}>
                      <ChevronRight size={15} style={{ opacity: 0.5, transition: "transform .15s ease", transform: isOpen ? "rotate(90deg)" : "none" }} />
                    </td>
                    <td style={{ padding: cellPad, whiteSpace: "nowrap" }}>{e.date}</td>
                    <td style={{ padding: cellPad, fontWeight: 600, whiteSpace: "nowrap" }}>
                      {e.company ? (onOpenCompany ? <button onClick={(ev) => { ev.stopPropagation(); onOpenCompany(e.companyId, e.company); }} className="tap" style={{ background: "transparent", border: "none", color: EMAIL, fontWeight: 600, fontSize: 13, cursor: "pointer", padding: 0 }}>{e.company}</button> : e.company) : "—"}
                    </td>
                    <td style={{ padding: cellPad, textAlign: "center" }}>{e.calls || 0}</td>
                    <td style={{ padding: cellPad, textAlign: "center" }}>{e.emails || 0}</td>
                    <td style={{ padding: cellPad, textAlign: "center" }}>{e.appts || 0}</td>
                    <td style={{ padding: cellPad, whiteSpace: "nowrap" }}>{nameOf(e.userId) || "—"}</td>
                    <td style={{ padding: cellPad, whiteSpace: "nowrap" }}>{repOf(e)}</td>
                  </tr>
                  {isOpen && (
                    <tr style={{ borderBottom: `1px solid ${LINE_C}`, background: "#F8FAFC" }}>
                      <td colSpan={headCols.length + 1} style={{ padding: 0 }}>
                        <div style={{ padding: "6px 22px 18px 48px", display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))", gap: "14px 28px" }}>
                          <div><div style={kvLabel}>BAN</div><div style={kvValue}>{e.ban || "—"}</div></div>
                          <div><div style={kvLabel}>FAN</div><div style={kvValue}>{e.fan || "—"}</div></div>
                          <div><div style={kvLabel}>Contact</div><div style={kvValue}>{e.contact || "—"}</div></div>
                          <div><div style={kvLabel}>Phone</div><div style={kvValue}>{e.phone || "—"}</div></div>
                          <div><div style={kvLabel}>Email</div><div style={kvValue}>{e.email || "—"}</div></div>
                          <div>
                            <div style={kvLabel}>Carrier Rep</div>
                            <div style={kvValue}>{e.carrierRep ? <span style={{ display: "inline-block", padding: "2px 10px", borderRadius: 20, fontSize: 12.5, fontWeight: 600, background: "#E7F0FF", color: EMAIL }}>{e.carrierRep}</span> : "—"}</div>
                          </div>
                          <div style={{ gridColumn: "1 / -1" }}><div style={kvLabel}>Notes</div><div style={kvValue}>{e.notes || "—"}</div></div>
                        </div>
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function ReportsView({ entries, deals, users, liveUser, visibleUserIds }) {
  const nameOf = (id) => { const u = users.find((x) => x.id === id); return u ? u.name : "Unknown"; };
  const months = buildMonthOptions().slice().reverse(); // oldest -> newest for trend charts

  // --- Activity trends by month ---
  const activityByMonth = months.map((mk) => {
    const rows = entries.filter((e) => inMonth(e.date, mk));
    return {
      month: monthLabel(mk).replace(/ \d{4}$/, ""), // short label
      Calls: rows.reduce((a, e) => a + (e.calls || 0), 0),
      Emails: rows.reduce((a, e) => a + (e.emails || 0), 0),
      Appts: rows.reduce((a, e) => a + (e.appts || 0), 0),
    };
  });

  // --- Deal / win-rate trends by month (based on close date for won/lost) ---
  const dealByMonth = months.map((mk) => {
    const closedThis = deals.filter((d) => (d.stage === "won" || d.stage === "lost") && inMonth(d.closeDate, mk));
    const won = closedThis.filter((d) => d.stage === "won");
    const lost = closedThis.filter((d) => d.stage === "lost");
    const created = deals.filter((d) => inMonth((d.createdAt || "").slice(0, 10), mk));
    const rate = closedThis.length ? Math.round((won.length / closedThis.length) * 100) : 0;
    return {
      month: monthLabel(mk).replace(/ \d{4}$/, ""),
      "New deals": created.length,
      Won: won.length,
      Lost: lost.length,
      "Win %": rate,
      "Won $": won.reduce((a, d) => a + (d.value || 0), 0),
    };
  });

  // --- Per-rep comparison (all-time, within visible scope) ---
  const repRows = users
    .filter((u) => (u.role === "bdr" || u.role === "sales") && (!visibleUserIds || visibleUserIds.includes(u.id)))
    .map((u) => {
      const myEntries = entries.filter((e) => e.userId === u.id);
      const myDeals = deals.filter((d) => d.ownerId === u.id || d.taggedRepId === u.id);
      const won = myDeals.filter((d) => d.stage === "won");
      const closed = myDeals.filter((d) => d.stage === "won" || d.stage === "lost");
      return {
        id: u.id, name: u.name, role: ROLES[u.role].label,
        calls: myEntries.reduce((a, e) => a + (e.calls || 0), 0),
        emails: myEntries.reduce((a, e) => a + (e.emails || 0), 0),
        appts: myEntries.reduce((a, e) => a + (e.appts || 0), 0),
        won: won.length,
        wonValue: won.reduce((a, d) => a + (d.value || 0), 0),
        winRate: closed.length ? Math.round((won.length / closed.length) * 100) : null,
      };
    })
    .sort((a, b) => b.wonValue - a.wonValue);

  const hasActivity = activityByMonth.some((m) => m.Calls || m.Emails || m.Appts);
  const hasDeals = dealByMonth.some((m) => m["New deals"] || m.Won || m.Lost);

  const axisStyle = { fontSize: 11, fill: "#8494A6" };

  return (
    <div>
      <div style={{ marginBottom: 20 }}>
        <h1 style={{ fontFamily: "'Fraunces', serif", fontSize: 29, fontWeight: 600, margin: 0 }}>Reports</h1>
        <p style={{ fontSize: 14, opacity: 0.6, margin: "4px 0 0" }}>Trends over time and rep comparison, across everything you can see.</p>
      </div>

      <Panel title="Activity over time" icon={TrendingUp}>
        {!hasActivity ? <Empty msg="No activity recorded yet." /> : (
          <ResponsiveContainer width="100%" height={280}>
            <LineChart data={activityByMonth} margin={{ top: 8, right: 12, bottom: 4, left: -8 }}>
              <CartesianGrid strokeDasharray="3 3" stroke={LINE_C} />
              <XAxis dataKey="month" tick={axisStyle} />
              <YAxis tick={axisStyle} />
              <Tooltip />
              <Legend />
              <Line type="monotone" dataKey="Calls" stroke={CALL} strokeWidth={2} dot={false} />
              <Line type="monotone" dataKey="Emails" stroke={EMAIL} strokeWidth={2} dot={false} />
              <Line type="monotone" dataKey="Appts" stroke={APPT} strokeWidth={2} dot={false} />
            </LineChart>
          </ResponsiveContainer>
        )}
      </Panel>

      <div className="charts" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginTop: 16 }}>
        <Panel title="Deals closed per month" icon={Briefcase}>
          {!hasDeals ? <Empty msg="No closed deals yet." /> : (
            <ResponsiveContainer width="100%" height={240}>
              <BarChart data={dealByMonth} margin={{ top: 8, right: 8, bottom: 4, left: -12 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={LINE_C} />
                <XAxis dataKey="month" tick={axisStyle} />
                <YAxis tick={axisStyle} />
                <Tooltip />
                <Legend />
                <Bar dataKey="Won" fill="#16A34A" radius={[3, 3, 0, 0]} />
                <Bar dataKey="Lost" fill="#B4453F" radius={[3, 3, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          )}
        </Panel>
        <Panel title="Win rate trend" icon={Percent}>
          {!hasDeals ? <Empty msg="No closed deals yet." /> : (
            <ResponsiveContainer width="100%" height={240}>
              <LineChart data={dealByMonth} margin={{ top: 8, right: 12, bottom: 4, left: -12 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={LINE_C} />
                <XAxis dataKey="month" tick={axisStyle} />
                <YAxis tick={axisStyle} domain={[0, 100]} unit="%" />
                <Tooltip />
                <Line type="monotone" dataKey="Win %" stroke={CYAN} strokeWidth={2} dot={{ r: 3 }} />
              </LineChart>
            </ResponsiveContainer>
          )}
        </Panel>
      </div>

      <Panel title="Rep comparison" icon={Trophy} style={{ marginTop: 16 }}>
        {repRows.length === 0 ? <Empty msg="No reps in scope." /> : (
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead>
                <tr style={{ background: "#F1F5F9" }}>
                  {["Rep", "Calls", "Emails", "Appts", "Won", "Won $", "Win %"].map((h, i) => (
                    <th key={h} style={{ textAlign: i === 0 ? "left" : "right", padding: "10px 12px", fontWeight: 700, fontSize: 11.5, letterSpacing: 0.4, textTransform: "uppercase", color: "#5A6B7B", whiteSpace: "nowrap" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {repRows.map((r) => (
                  <tr key={r.id} style={{ borderTop: `1px solid ${LINE_C}` }}>
                    <td style={{ padding: "10px 12px" }}><div style={{ fontWeight: 600 }}>{r.name}</div><div style={{ fontSize: 11.5, opacity: 0.55 }}>{r.role}</div></td>
                    <td style={{ padding: "10px 12px", textAlign: "right" }}>{r.calls.toLocaleString()}</td>
                    <td style={{ padding: "10px 12px", textAlign: "right" }}>{r.emails.toLocaleString()}</td>
                    <td style={{ padding: "10px 12px", textAlign: "right" }}>{r.appts.toLocaleString()}</td>
                    <td style={{ padding: "10px 12px", textAlign: "right" }}>{r.won}</td>
                    <td style={{ padding: "10px 12px", textAlign: "right" }}>{fmtMoney(r.wonValue)}</td>
                    <td style={{ padding: "10px 12px", textAlign: "right" }}>{r.winRate == null ? "—" : `${r.winRate}%`}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Panel>
    </div>
  );
}

function FollowUpsView({ deals, leads, users, liveUser, setView }) {
  const nameOf = (id) => { const u = users.find((x) => x.id === id); return u ? u.name : ""; };

  // Build a unified list of follow-up items from deals and leads that have a next-action date.
  const items = [];
  (deals || []).forEach((d) => {
    if (d.nextActionDate && d.stage !== "won" && d.stage !== "lost") {
      items.push({ kind: "Deal", id: d.id, company: d.company, who: nameOf(d.ownerId), date: d.nextActionDate, sub: STAGE[d.stage]?.label || d.stage, value: d.value });
    }
  });
  (leads || []).forEach((l) => {
    if (l.nextActionDate && l.status !== "Won" && l.status !== "Lost" && l.status !== "Dead") {
      items.push({ kind: "Lead", id: l.id, company: l.company, who: nameOf(l.assignedTo) || "Unassigned", date: l.nextActionDate, sub: l.status });
    }
  });
  items.sort((a, b) => a.date.localeCompare(b.date));

  const overdue = items.filter((i) => daysUntil(i.date) < 0);
  const today = items.filter((i) => daysUntil(i.date) === 0);
  const soon = items.filter((i) => { const d = daysUntil(i.date); return d > 0 && d <= 7; });
  const later = items.filter((i) => daysUntil(i.date) > 7);

  const Section = ({ title, list, accent }) => list.length === 0 ? null : (
    <div style={{ marginBottom: 22 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
        <span style={{ width: 9, height: 9, borderRadius: "50%", background: accent }} />
        <h3 style={{ fontSize: 14, fontWeight: 700, margin: 0, color: INK }}>{title}</h3>
        <span style={{ fontSize: 12.5, opacity: 0.5 }}>({list.length})</span>
      </div>
      <div style={{ background: CARD, border: `1px solid ${LINE_C}`, borderRadius: 12, overflow: "hidden" }}>
        {list.map((i, idx) => {
          const fu = followUpState(i.date);
          return (
            <div key={i.kind + i.id} style={{ display: "flex", alignItems: "center", gap: 12, padding: "12px 16px", borderTop: idx ? `1px solid ${LINE_C}` : "none" }}>
              <span style={{ fontSize: 10.5, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.4, color: i.kind === "Deal" ? EMAIL : CALL, background: (i.kind === "Deal" ? EMAIL : CALL) + "14", borderRadius: 5, padding: "3px 7px", flexShrink: 0 }}>{i.kind}</span>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontSize: 14, fontWeight: 600, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{i.company}</div>
                <div style={{ fontSize: 12, opacity: 0.55 }}>{i.sub} · {i.who}{i.value ? ` · ${fmtMoney(i.value)}` : ""}</div>
              </div>
              <span style={{ fontSize: 12, fontWeight: 600, color: fu.color, whiteSpace: "nowrap", flexShrink: 0 }}>{fu.label}</span>
            </div>
          );
        })}
      </div>
    </div>
  );

  return (
    <div>
      <div style={{ marginBottom: 20 }}>
        <h1 style={{ fontFamily: "'Fraunces', serif", fontSize: 29, fontWeight: 600, margin: 0 }}>Follow-ups</h1>
        <p style={{ fontSize: 14, opacity: 0.6, margin: "4px 0 0" }}>Deals and leads with a next-action date, sorted by urgency.</p>
      </div>
      {items.length === 0 ? (
        <div style={{ background: CARD, border: `1px solid ${LINE_C}`, borderRadius: 14, padding: 40, textAlign: "center" }}>
          <Clock size={30} color="#B7C2CE" style={{ marginBottom: 10 }} />
          <div style={{ fontSize: 15, fontWeight: 600, marginBottom: 4 }}>Nothing scheduled</div>
          <p style={{ fontSize: 13.5, opacity: 0.6, margin: "0 0 16px" }}>Add a "next action" date to a deal or lead and it'll show up here.</p>
          <button onClick={() => setView("pipeline")} className="tap" style={{ background: INK, color: PAPER, border: "none", borderRadius: 9, padding: "9px 16px", fontSize: 13.5, fontWeight: 600, cursor: "pointer" }}>Go to Pipeline</button>
        </div>
      ) : (
        <>
          <Section title="Overdue" list={overdue} accent="#B4453F" />
          <Section title="Due today" list={today} accent="#C2410C" />
          <Section title="Next 7 days" list={soon} accent="#B7791F" />
          <Section title="Later" list={later} accent="#5A6B7B" />
        </>
      )}
    </div>
  );
}

function Dashboard({ entries, deals, leads, users, goals, saveGoals, userGoals, liveUser, visibleUserIds, setView, canLog, onOpenCompany }) {
  const role = liveUser.role;
  const scopeUsers = users.filter((u) => visibleUserIds.includes(u.id));
  const repUsers = scopeUsers.filter((u) => u.role === "bdr" || u.role === "sales");

  const [repFilter, setRepFilter] = useState("all");
  const [month, setMonth] = useState(CURRENT_MONTH);
  const monthOptions = buildMonthOptions();
  const isCurrentMonth = month === CURRENT_MONTH;

  // Month-scope first: everything on the dashboard reflects the selected month.
  const monthEntries = entries.filter((e) => inMonth(e.date, month));
  const monthDeals = deals.filter((d) => inMonth(d.closeDate, month) || inMonth(d.createdAt, month));

  const scoped = repFilter === "all" ? monthEntries : monthEntries.filter((e) => e.userId === repFilter);
  const scopedDeals = repFilter === "all" ? monthDeals : monthDeals.filter((d) => d.ownerId === repFilter);

  // Pipeline rollups for the dashboard
  const openDeals = scopedDeals.filter((d) => d.stage !== "won" && d.stage !== "lost");
  const wonDeals = scopedDeals.filter((d) => d.stage === "won");
  const lostDeals = scopedDeals.filter((d) => d.stage === "lost");
  const pipelineValue = openDeals.reduce((s, d) => s + (d.value || 0), 0);
  const weightedValue = openDeals.reduce((s, d) => s + (d.value || 0) * (STAGE[d.stage]?.prob || 0) / 100, 0);
  const wonValue = wonDeals.reduce((s, d) => s + (d.value || 0), 0);
  const closedCount = wonDeals.length + lostDeals.length;
  const winRate = closedCount ? ((wonDeals.length / closedCount) * 100).toFixed(0) : "0";

  const totals = scoped.reduce((a, e) => {
    a.calls += e.calls || 0; a.emails += e.emails || 0; a.appts += e.appts || 0; return a;
  }, { calls: 0, emails: 0, appts: 0 });

  const outreach = totals.calls + totals.emails;
  const conv = outreach ? ((totals.appts / outreach) * 100).toFixed(1) : "0.0";
  const callConv = totals.calls ? ((totals.appts / totals.calls) * 100).toFixed(1) : "0.0";
  const emailConv = totals.emails ? ((totals.appts / totals.emails) * 100).toFixed(1) : "0.0";

  const trend = (() => {
    const [y, m] = month.split("-").map(Number);
    const daysInMonth = new Date(y, m, 0).getDate();
    // Current month: 1st → today. Past month: full 1st → last day.
    const lastDay = isCurrentMonth ? new Date().getDate() : daysInMonth;
    const days = [];
    for (let day = 1; day <= lastDay; day++) {
      const key = `${month}-${String(day).padStart(2, "0")}`;
      const de = scoped.filter((e) => e.date === key);
      days.push({
        date: new Date(y, m - 1, day).toLocaleDateString("en-US", { month: "short", day: "numeric" }),
        Calls: de.reduce((s, e) => s + (e.calls || 0), 0),
        Emails: de.reduce((s, e) => s + (e.emails || 0), 0),
        Appts: de.reduce((s, e) => s + (e.appts || 0), 0),
      });
    }
    return days;
  })();

  const leaderboard = (() => {
    const map = {};
    repUsers.forEach((u) => (map[u.id] = { rep: u.name, calls: 0, emails: 0, appts: 0 }));
    monthEntries.forEach((e) => {
      if (!map[e.userId]) return;
      map[e.userId].calls += e.calls || 0;
      map[e.userId].emails += e.emails || 0;
      map[e.userId].appts += e.appts || 0;
    });
    return Object.values(map).sort((a, b) => b.appts - a.appts);
  })();

  const scopeLabel = role === "bdr" ? "Your activity"
    : role === "sales" ? "You + activity tagged to you"
    : "All activity";

  const showLeaderboard = role !== "bdr";
  // Effective goal per rep = their override if set, else the team default.
  const effGoal = (uid) => (userGoals && userGoals[uid]) ? userGoals[uid] : goals;
  const targets = (() => {
    if (repFilter !== "all") {
      // Single rep selected: measure against that rep's own goal.
      return { ...effGoal(repFilter) };
    }
    // "All in scope" (or a BDR viewing only themselves): sum each rep's effective goal.
    return repUsers.reduce((acc, u) => {
      const g = effGoal(u.id);
      acc.calls += g.calls || 0; acc.emails += g.emails || 0; acc.appts += g.appts || 0;
      return acc;
    }, { calls: 0, emails: 0, appts: 0 });
  })();

  return (
    <>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12, marginBottom: 20 }}>
        <div>
          <h1 style={{ fontFamily: "'Fraunces', serif", fontSize: 29, fontWeight: 600, margin: 0 }}>
            {role === "bdr" ? "My Dashboard" : "Team Performance"}
          </h1>
          <p style={{ margin: "4px 0 0", opacity: 0.55, fontSize: 14 }}>
            {scopeLabel} · {monthLabel(month)}{isCurrentMonth ? " (month-to-date)" : ""} · {scoped.length} logged sessions
          </p>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
          {/* Month picker: arrows + dropdown */}
          <div style={{ display: "flex", alignItems: "center", gap: 4, background: CARD, border: `1px solid ${LINE_C}`, borderRadius: 9, padding: 3 }}>
            <button
              onClick={() => { const i = monthOptions.indexOf(month); if (i < monthOptions.length - 1) setMonth(monthOptions[i + 1]); }}
              disabled={monthOptions.indexOf(month) >= monthOptions.length - 1}
              className="tap" title="Previous month"
              style={{ background: "transparent", border: "none", borderRadius: 6, padding: "6px 8px", cursor: "pointer", color: INK, opacity: monthOptions.indexOf(month) >= monthOptions.length - 1 ? 0.3 : 1, display: "grid", placeItems: "center" }}>
              <ChevronLeft size={16} />
            </button>
            <div style={{ position: "relative" }}>
              <select value={month} onChange={(e) => setMonth(e.target.value)}
                style={{ appearance: "none", background: "transparent", border: "none", padding: "4px 26px 4px 8px", fontSize: 14, fontWeight: 600, color: INK, cursor: "pointer", minWidth: 130, textAlign: "center" }}>
                {monthOptions.map((mk) => (
                  <option key={mk} value={mk}>{monthLabel(mk)}{mk === CURRENT_MONTH ? " (MTD)" : ""}</option>
                ))}
              </select>
              <ChevronDown size={14} style={{ position: "absolute", right: 8, top: 9, pointerEvents: "none", opacity: 0.5 }} />
            </div>
            <button
              onClick={() => { const i = monthOptions.indexOf(month); if (i > 0) setMonth(monthOptions[i - 1]); }}
              disabled={monthOptions.indexOf(month) <= 0}
              className="tap" title="Next month"
              style={{ background: "transparent", border: "none", borderRadius: 6, padding: "6px 8px", cursor: "pointer", color: INK, opacity: monthOptions.indexOf(month) <= 0 ? 0.3 : 1, display: "grid", placeItems: "center" }}>
              <ChevronRight size={16} />
            </button>
          </div>
          {role !== "bdr" && repUsers.length > 0 && (
            <div style={{ position: "relative" }}>
              <select value={repFilter} onChange={(e) => setRepFilter(e.target.value)}
                style={{ appearance: "none", background: CARD, border: `1px solid ${LINE_C}`, borderRadius: 9, padding: "9px 34px 9px 14px", fontSize: 14, fontWeight: 500, color: INK, cursor: "pointer" }}>
                <option value="all">All in scope</option>
                {repUsers.map((u) => <option key={u.id} value={u.id}>{u.name}</option>)}
              </select>
              <ChevronDown size={15} style={{ position: "absolute", right: 12, top: 11, pointerEvents: "none", opacity: 0.5 }} />
            </div>
          )}
        </div>
      </div>

      <div className="grid-4" style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 14, marginBottom: 16 }}>
        <StatCard icon={Phone} color={CALL} label="Calls" value={totals.calls} sub={`${callConv}% → appt`} />
        <StatCard icon={Mail} color={EMAIL} label="Emails" value={totals.emails} sub={`${emailConv}% → appt`} />
        <StatCard icon={CalendarCheck} color={APPT} label="Appointments Set" value={totals.appts} sub={`${conv}% overall conv.`} />
        <StatCard icon={TrendingUp} color={INK} label="Total Outreach" value={outreach} sub={`${scoped.length ? (outreach / scoped.length).toFixed(0) : 0} avg/session`} />
      </div>

      <GoalBars totals={totals} targets={targets} isTeam={repFilter === "all" && role !== "bdr"}
        pace={isCurrentMonth ? (() => { const now = new Date(); const dim = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate(); return now.getDate() / dim; })() : null} />

      <div className="charts" style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr", gap: 16, marginTop: 16 }}>
        <Panel title={`Activity — ${monthLabel(month)}${isCurrentMonth ? " (MTD)" : ""}`}>
          <ResponsiveContainer width="100%" height={260}>
            <LineChart data={trend} margin={{ top: 8, right: 8, left: -18, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke={LINE_C} vertical={false} />
              <XAxis dataKey="date" tick={{ fontSize: 11, fill: INK }} interval={trend.length > 16 ? 2 : 1} axisLine={{ stroke: LINE_C }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: INK }} axisLine={false} tickLine={false} />
              <Tooltip contentStyle={{ borderRadius: 10, border: `1px solid ${LINE_C}`, fontSize: 13 }} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Line type="monotone" dataKey="Calls" stroke={CALL} strokeWidth={2} dot={false} />
              <Line type="monotone" dataKey="Emails" stroke={EMAIL} strokeWidth={2} dot={false} />
              <Line type="monotone" dataKey="Appts" stroke={APPT} strokeWidth={2.5} dot={false} />
            </LineChart>
          </ResponsiveContainer>
        </Panel>
        <Panel title="Conversion funnel">
          <Funnel calls={totals.calls} emails={totals.emails} appts={totals.appts} />
        </Panel>
      </div>

      {/* Pipeline snapshot */}
      <div style={{ marginTop: 22, marginBottom: 8, display: "flex", alignItems: "center", gap: 8 }}>
        <Briefcase size={17} color={INK} />
        <h2 style={{ fontFamily: "'Fraunces', serif", fontSize: 20, fontWeight: 600, margin: 0 }}>Sales Pipeline</h2>
      </div>
      <div className="grid-4" style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 14, marginBottom: 16 }}>
        <StatCard icon={Briefcase} color={EMAIL} label="Open Pipeline" value={fmtMoney(pipelineValue)} sub={`${openDeals.length} open deals`} raw />
        <StatCard icon={TrendingUp} color={CYAN} label="Weighted Forecast" value={fmtMoney(Math.round(weightedValue))} sub="by stage probability" raw />
        <StatCard icon={DollarSign} color="#16A34A" label="Closed Won" value={fmtMoney(wonValue)} sub={`${wonDeals.length} won deals`} raw />
        <StatCard icon={Percent} color={APPT} label="Win Rate" value={`${winRate}%`} sub={`${closedCount} closed total`} raw />
      </div>
      {(() => {
        // Assemble due/overdue follow-ups from scoped deals + leads.
        const fuItems = [];
        (deals || []).forEach((d) => { if (d.nextActionDate && d.stage !== "won" && d.stage !== "lost") fuItems.push({ kind: "Deal", company: d.company, date: d.nextActionDate }); });
        (leads || []).forEach((l) => { if (l.nextActionDate && !["Won", "Lost", "Dead"].includes(l.status)) fuItems.push({ kind: "Lead", company: l.company, date: l.nextActionDate }); });
        const due = fuItems.filter((i) => daysUntil(i.date) <= 0).sort((a, b) => a.date.localeCompare(b.date));
        const staleDeals = (deals || []).filter((d) => { const a = dealAge(d); return a && a.stale; });
        if (due.length === 0 && staleDeals.length === 0) return null;
        return (
          <div style={{ background: "#FFF9F0", border: "1px solid #F0DFC0", borderRadius: 14, padding: "14px 18px", marginBottom: 16 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: due.length ? 10 : 0 }}>
              <AlertTriangle size={16} color="#B7791F" />
              <h3 style={{ fontSize: 14, fontWeight: 700, margin: 0, color: "#8A6D3B" }}>Needs attention</h3>
              <button onClick={() => setView("followups")} className="tap" style={{ marginLeft: "auto", background: "transparent", border: "none", color: EMAIL, fontSize: 12.5, fontWeight: 600, cursor: "pointer" }}>View all →</button>
            </div>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
              {due.slice(0, 6).map((i, idx) => {
                const fu = followUpState(i.date);
                return <span key={idx} style={{ fontSize: 12, background: CARD, border: `1px solid ${LINE_C}`, borderRadius: 7, padding: "5px 9px" }}><b style={{ color: fu.color }}>{fu.label}</b> · {i.company}</span>;
              })}
              {staleDeals.length > 0 && <span style={{ fontSize: 12, background: CARD, border: `1px solid ${LINE_C}`, borderRadius: 7, padding: "5px 9px", color: "#8A6D3B" }}>⚠ {staleDeals.length} stale deal{staleDeals.length > 1 ? "s" : ""}</span>}
            </div>
          </div>
        );
      })()}
      <Panel title="Deals by stage">
        {openDeals.length + wonDeals.length + lostDeals.length === 0 ? (
          <Empty msg="No deals yet. Reps can add opportunities from the Pipeline tab." />
        ) : (
          <div style={{ display: "grid", gap: 10 }}>
            {STAGES.map((s) => {
              const inStage = scopedDeals.filter((d) => d.stage === s.id);
              const val = inStage.reduce((sum, d) => sum + (d.value || 0), 0);
              const maxVal = Math.max(1, ...STAGES.map((st) => scopedDeals.filter((d) => d.stage === st.id).reduce((sum, d) => sum + (d.value || 0), 0)));
              return (
                <div key={s.id} style={{ display: "flex", alignItems: "center", gap: 12 }}>
                  <span style={{ width: 120, fontSize: 13, fontWeight: 500, display: "flex", alignItems: "center", gap: 7 }}>
                    <span style={{ width: 9, height: 9, borderRadius: "50%", background: s.color, flexShrink: 0 }} />{s.label}
                  </span>
                  <div style={{ flex: 1, height: 22, background: LINE_C, borderRadius: 6, overflow: "hidden" }}>
                    <div style={{ width: `${(val / maxVal) * 100}%`, height: "100%", background: s.color, borderRadius: 6, minWidth: val > 0 ? 3 : 0, transition: "width .4s ease" }} />
                  </div>
                  <span style={{ width: 130, textAlign: "right", fontSize: 13, opacity: 0.75 }}>{fmtMoney(val)} · {inStage.length}</span>
                </div>
              );
            })}
          </div>
        )}
      </Panel>

      {lostDeals.some((d) => d.lostReason) && (
        <Panel title="Why deals were lost" style={{ marginTop: 16 }}>
          <div style={{ display: "grid", gap: 10 }}>
            {(() => {
              const counts = {};
              lostDeals.forEach((d) => { const r = d.lostReason || "Unspecified"; counts[r] = (counts[r] || 0) + 1; });
              const ordered = Object.entries(counts).sort((a, b) => b[1] - a[1]);
              const maxC = Math.max(1, ...ordered.map(([, c]) => c));
              return ordered.map(([reason, count]) => (
                <div key={reason} style={{ display: "flex", alignItems: "center", gap: 12 }}>
                  <span style={{ width: 150, fontSize: 13, fontWeight: 500 }}>{reason}</span>
                  <div style={{ flex: 1, height: 20, background: LINE_C, borderRadius: 6, overflow: "hidden" }}>
                    <div style={{ width: `${(count / maxC) * 100}%`, height: "100%", background: "#B4453F", borderRadius: 6, minWidth: 3, transition: "width .4s ease" }} />
                  </div>
                  <span style={{ width: 40, textAlign: "right", fontSize: 13, opacity: 0.75 }}>{count}</span>
                </div>
              ));
            })()}
          </div>
        </Panel>
      )}

      {showLeaderboard && (
        <Panel title="Rep leaderboard" style={{ marginTop: 16 }} icon={Trophy}>
          {leaderboard.every((r) => !r.calls && !r.emails && !r.appts) ? (
            <Empty msg="No activity logged yet in your scope." />
          ) : (
            <>
              <ResponsiveContainer width="100%" height={Math.max(140, leaderboard.length * 46)}>
                <BarChart data={leaderboard} layout="vertical" margin={{ left: 8, right: 24, top: 4, bottom: 4 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={LINE_C} horizontal={false} />
                  <XAxis type="number" tick={{ fontSize: 11, fill: INK }} axisLine={false} tickLine={false} />
                  <YAxis type="category" dataKey="rep" width={110} tick={{ fontSize: 12, fill: INK }} axisLine={false} tickLine={false} />
                  <Tooltip contentStyle={{ borderRadius: 10, border: `1px solid ${LINE_C}`, fontSize: 13 }} cursor={{ fill: "rgba(0,0,0,.03)" }} />
                  <Legend wrapperStyle={{ fontSize: 12 }} />
                  <Bar dataKey="calls" name="Calls" fill={CALL} radius={[0, 4, 4, 0]} />
                  <Bar dataKey="emails" name="Emails" fill={EMAIL} radius={[0, 4, 4, 0]} />
                  <Bar dataKey="appts" name="Appts" fill={APPT} radius={[0, 4, 4, 0]} />
                </BarChart>
              </ResponsiveContainer>
              <div style={{ marginTop: 14, borderTop: `1px solid ${LINE_C}`, paddingTop: 10 }}>
                {leaderboard.map((r, i) => {
                  const c = (r.calls + r.emails) ? ((r.appts / (r.calls + r.emails)) * 100).toFixed(1) : "0.0";
                  return (
                    <div key={r.rep} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "8px 0", borderBottom: i < leaderboard.length - 1 ? `1px solid ${LINE_C}` : "none" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                        <span style={{ width: 22, height: 22, borderRadius: "50%", background: i === 0 ? APPT : LINE_C, color: i === 0 ? "#fff" : INK, display: "grid", placeItems: "center", fontSize: 12, fontWeight: 700 }}>{i + 1}</span>
                        <span style={{ fontWeight: 500, fontSize: 14 }}>{r.rep}</span>
                      </div>
                      <span style={{ fontSize: 13, opacity: 0.7 }}><b style={{ color: APPT }}>{r.appts}</b> appts · {c}% conv</span>
                    </div>
                  );
                })}
              </div>
            </>
          )}
        </Panel>
      )}

      {showLeaderboard && <WeeklyDigest entries={entries} deals={deals} users={users} />}

      <div style={{ marginTop: 16 }}>
        <ActivityTable entries={scoped} users={users} liveUser={liveUser} compact onOpenCompany={onOpenCompany} />
      </div>
    </>
  );
}

function WeeklyDigest({ entries, deals, users }) {
  const [copied, setCopied] = useState(false);
  const nameOf = (id) => { const u = users.find((x) => x.id === id); return u ? u.name : "Unknown"; };

  // Last 7 days window.
  const since = new Date(); since.setDate(since.getDate() - 7); since.setHours(0, 0, 0, 0);
  const sinceStr = since.toISOString().slice(0, 10);

  const weekEntries = (entries || []).filter((e) => e.date >= sinceStr);
  const t = weekEntries.reduce((a, e) => ({ calls: a.calls + (e.calls || 0), emails: a.emails + (e.emails || 0), appts: a.appts + (e.appts || 0) }), { calls: 0, emails: 0, appts: 0 });

  const weekDeals = (deals || []).filter((d) => (d.createdAt || "").slice(0, 10) >= sinceStr);
  const wonThisWeek = (deals || []).filter((d) => d.stage === "won" && (d.closeDate || "").slice(0, 10) >= sinceStr);
  const lostThisWeek = (deals || []).filter((d) => d.stage === "lost" && (d.closeDate || "").slice(0, 10) >= sinceStr);

  // Per-rep activity rollup.
  const byRep = {};
  weekEntries.forEach((e) => { const n = nameOf(e.userId); byRep[n] = byRep[n] || { calls: 0, emails: 0, appts: 0 }; byRep[n].calls += e.calls || 0; byRep[n].emails += e.emails || 0; byRep[n].appts += e.appts || 0; });

  const buildText = () => {
    const lines = [];
    lines.push(`Tellemica — Weekly Summary (last 7 days)`);
    lines.push(``);
    lines.push(`Activity: ${t.calls} calls · ${t.emails} emails · ${t.appts} appointments`);
    lines.push(`New deals: ${weekDeals.length} · Won: ${wonThisWeek.length} · Lost: ${lostThisWeek.length}`);
    if (wonThisWeek.length) lines.push(`Wins: ${wonThisWeek.map((d) => `${d.company} (${fmtMoney(d.value)})`).join(", ")}`);
    if (lostThisWeek.length) lines.push(`Losses: ${lostThisWeek.map((d) => `${d.company}${d.lostReason ? ` — ${d.lostReason}` : ""}`).join(", ")}`);
    lines.push(``);
    lines.push(`By rep:`);
    Object.entries(byRep).sort((a, b) => b[1].appts - a[1].appts).forEach(([n, r]) => lines.push(`  ${n}: ${r.calls} calls, ${r.emails} emails, ${r.appts} appts`));
    return lines.join("\n");
  };

  const copy = async () => {
    try { await navigator.clipboard.writeText(buildText()); setCopied(true); setTimeout(() => setCopied(false), 1800); }
    catch { /* clipboard blocked; user can still read the panel */ }
  };

  const hasData = weekEntries.length > 0 || weekDeals.length > 0 || wonThisWeek.length > 0 || lostThisWeek.length > 0;

  return (
    <Panel title="This week's summary" style={{ marginTop: 16 }} icon={FileSpreadsheet}>
      {!hasData ? (
        <Empty msg="No activity or deal changes in the last 7 days." />
      ) : (
        <>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 18, marginBottom: 14 }}>
            <MiniStat label="Calls" value={t.calls} color={CALL} />
            <MiniStat label="Emails" value={t.emails} color={EMAIL} />
            <MiniStat label="Appts" value={t.appts} color={APPT} />
            <MiniStat label="New deals" value={weekDeals.length} color={EMAIL} />
            <MiniStat label="Won" value={wonThisWeek.length} color="#16A34A" />
            <MiniStat label="Lost" value={lostThisWeek.length} color="#B4453F" />
          </div>
          <pre style={{ background: "#F1F5F9", border: `1px solid ${LINE_C}`, borderRadius: 10, padding: 14, fontSize: 12.5, lineHeight: 1.6, whiteSpace: "pre-wrap", fontFamily: "ui-monospace, SFMono-Regular, Menlo, monospace", margin: "0 0 12px" }}>{buildText()}</pre>
          <button onClick={copy} className="tap" style={{ background: INK, color: PAPER, border: "none", borderRadius: 9, padding: "9px 16px", fontSize: 13.5, fontWeight: 600, cursor: "pointer" }}>
            {copied ? "Copied!" : "Copy summary"}
          </button>
        </>
      )}
    </Panel>
  );
}

function StatCard({ icon: Icon, color, label, value, sub, raw }) {
  return (
    <div style={{ background: CARD, border: `1px solid ${LINE_C}`, borderRadius: 14, padding: 18 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14 }}>
        <span style={{ width: 32, height: 32, borderRadius: 9, background: color + "18", display: "grid", placeItems: "center" }}>
          <Icon size={17} color={color} />
        </span>
        <span style={{ fontSize: 13, fontWeight: 500, opacity: 0.65 }}>{label}</span>
      </div>
      <div style={{ fontFamily: "'Fraunces', serif", fontSize: 34, fontWeight: 600, lineHeight: 1 }}>{raw ? value : value.toLocaleString()}</div>
      <div style={{ fontSize: 12.5, opacity: 0.55, marginTop: 6 }}>{sub}</div>
    </div>
  );
}

function GoalsManager({ users, visibleUserIds, liveUser, goals, saveGoals, userGoals, saveUserGoals }) {
  const [draftDefault, setDraftDefault] = useState(goals);
  const [editingDefault, setEditingDefault] = useState(false);
  const [editRow, setEditRow] = useState(null); // userId being edited
  const [rowDraft, setRowDraft] = useState({ calls: 0, emails: 0, appts: 0 });
  const [savingMsg, setSavingMsg] = useState("");

  // Reps this manager/admin can set goals for = those in their scope who log activity.
  const reps = users.filter((u) => visibleUserIds.includes(u.id) && (u.role === "bdr" || u.role === "sales"));

  const effGoal = (uid) => userGoals[uid] || goals;
  const hasOverride = (uid) => !!userGoals[uid];

  const startEditRow = (u) => { setEditRow(u.id); setRowDraft({ ...effGoal(u.id) }); };

  const saveRow = async (uid) => {
    setSavingMsg("Saving…");
    await saveUserGoals(() => api.setUserGoal(uid, {
      calls: +rowDraft.calls || 0, emails: +rowDraft.emails || 0, appts: +rowDraft.appts || 0,
    }));
    setEditRow(null); setSavingMsg("");
  };
  const resetRow = async (uid) => {
    if (!confirm("Reset this person to the team default goal?")) return;
    setSavingMsg("Resetting…");
    await saveUserGoals(() => api.clearUserGoal(uid));
    setSavingMsg("");
  };
  const saveDefault = async () => {
    await saveGoals({ calls: +draftDefault.calls || 0, emails: +draftDefault.emails || 0, appts: +draftDefault.appts || 0 });
    setEditingDefault(false);
  };

  const num = (v, on) => (
    <input type="number" min="0" value={v} onChange={on}
      style={{ width: 82, padding: "8px 10px", border: `1px solid ${LINE_C}`, borderRadius: 8, fontSize: 14 }} />
  );

  return (
    <>
      <div style={{ marginBottom: 20 }}>
        <h1 style={{ fontFamily: "'Fraunces', serif", fontSize: 29, fontWeight: 600, margin: 0 }}>Goals</h1>
        <p style={{ margin: "4px 0 0", opacity: 0.55, fontSize: 14 }}>
          Set the team default and give individual reps their own monthly targets. {savingMsg && <b style={{ color: CALL }}>{savingMsg}</b>}
        </p>
      </div>

      {/* Team default */}
      <Panel title="Team default (monthly)" icon={Target} style={{ marginBottom: 16 }}>
        {editingDefault ? (
          <div style={{ display: "flex", gap: 12, flexWrap: "wrap", alignItems: "flex-end" }}>
            {["calls", "emails", "appts"].map((k) => (
              <label key={k} style={{ fontSize: 12, fontWeight: 500, opacity: 0.7 }}>
                <div style={{ marginBottom: 4, textTransform: "capitalize" }}>{k}/mo</div>
                {num(draftDefault[k], (e) => setDraftDefault({ ...draftDefault, [k]: e.target.value }))}
              </label>
            ))}
            <button onClick={saveDefault} className="tap" style={{ background: INK, color: PAPER, border: "none", borderRadius: 8, padding: "9px 18px", fontSize: 14, fontWeight: 500, cursor: "pointer" }}>Save</button>
            <button onClick={() => setEditingDefault(false)} className="tap" style={{ background: "transparent", color: INK, border: `1px solid ${LINE_C}`, borderRadius: 8, padding: "9px 14px", fontSize: 14, cursor: "pointer" }}>Cancel</button>
          </div>
        ) : (
          <div style={{ display: "flex", alignItems: "center", gap: 24, flexWrap: "wrap" }}>
            <GoalStat label="Calls" value={goals.calls} color={CALL} />
            <GoalStat label="Emails" value={goals.emails} color={EMAIL} />
            <GoalStat label="Appointments" value={goals.appts} color={APPT} />
            <button onClick={() => { setDraftDefault(goals); setEditingDefault(true); }} className="tap"
              style={{ marginLeft: "auto", background: "transparent", border: "none", color: EMAIL, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>Edit default</button>
          </div>
        )}
        <p style={{ fontSize: 12.5, opacity: 0.5, margin: "12px 0 0" }}>Anyone without a custom goal below uses these targets.</p>
      </Panel>

      {/* Per-person goals */}
      <Panel title="Individual goals" icon={Users}>
        {reps.length === 0 ? (
          <Empty msg="No reps in your scope yet." />
        ) : (
          <div style={{ display: "grid", gap: 8 }}>
            <div style={{ display: "grid", gridTemplateColumns: "1.6fr repeat(3, 1fr) auto", gap: 12, padding: "0 4px 6px", fontSize: 11.5, fontWeight: 700, letterSpacing: 0.5, textTransform: "uppercase", color: "#8494A6" }}>
              <span>Rep</span><span>Calls</span><span>Emails</span><span>Appts</span><span></span>
            </div>
            {reps.map((u) => {
              const g = effGoal(u.id);
              const editing = editRow === u.id;
              return (
                <div key={u.id} style={{ display: "grid", gridTemplateColumns: "1.6fr repeat(3, 1fr) auto", gap: 12, alignItems: "center", background: "#fff", border: `1px solid ${LINE_C}`, borderRadius: 11, padding: "12px 14px" }}>
                  <div style={{ minWidth: 0 }}>
                    <div style={{ fontWeight: 600, fontSize: 14 }}>{u.name}</div>
                    <div style={{ fontSize: 12, opacity: 0.5 }}>
                      {ROLES[u.role].label}{hasOverride(u.id) ? " · custom goal" : " · default"}
                    </div>
                  </div>
                  {editing ? (
                    <>
                      {num(rowDraft.calls, (e) => setRowDraft({ ...rowDraft, calls: e.target.value }))}
                      {num(rowDraft.emails, (e) => setRowDraft({ ...rowDraft, emails: e.target.value }))}
                      {num(rowDraft.appts, (e) => setRowDraft({ ...rowDraft, appts: e.target.value }))}
                      <div style={{ display: "flex", gap: 6 }}>
                        <button onClick={() => saveRow(u.id)} className="tap" style={{ background: INK, color: PAPER, border: "none", borderRadius: 8, padding: "8px 14px", fontSize: 13, fontWeight: 500, cursor: "pointer" }}>Save</button>
                        <button onClick={() => setEditRow(null)} className="tap" style={{ background: "transparent", border: `1px solid ${LINE_C}`, borderRadius: 8, padding: "8px 10px", fontSize: 13, cursor: "pointer" }}>Cancel</button>
                      </div>
                    </>
                  ) : (
                    <>
                      <span style={{ fontSize: 14, fontWeight: hasOverride(u.id) ? 600 : 400, opacity: hasOverride(u.id) ? 1 : 0.6 }}>{g.calls.toLocaleString()}</span>
                      <span style={{ fontSize: 14, fontWeight: hasOverride(u.id) ? 600 : 400, opacity: hasOverride(u.id) ? 1 : 0.6 }}>{g.emails.toLocaleString()}</span>
                      <span style={{ fontSize: 14, fontWeight: hasOverride(u.id) ? 600 : 400, opacity: hasOverride(u.id) ? 1 : 0.6 }}>{g.appts.toLocaleString()}</span>
                      <div style={{ display: "flex", gap: 4, justifyContent: "flex-end" }}>
                        <button onClick={() => startEditRow(u)} className="tap" style={iconBtn} title="Set custom goal"><Pencil size={14} /></button>
                        {hasOverride(u.id) && (
                          <button onClick={() => resetRow(u.id)} className="tap" style={iconBtn} title="Reset to default"><X size={14} /></button>
                        )}
                      </div>
                    </>
                  )}
                </div>
              );
            })}
          </div>
        )}
      </Panel>
    </>
  );
}

function GoalStat({ label, value, color }) {
  return (
    <div>
      <div style={{ fontSize: 12, fontWeight: 500, opacity: 0.6, marginBottom: 2 }}>{label}</div>
      <div style={{ fontFamily: "'Fraunces', serif", fontSize: 24, fontWeight: 600, color }}>{value.toLocaleString()}</div>
    </div>
  );
}

function GoalBars({ totals, targets, isTeam, pace }) {
  const rows = [
    ["Calls", totals.calls, targets.calls, CALL],
    ["Emails", totals.emails, targets.emails, EMAIL],
    ["Appointments", totals.appts, targets.appts, APPT],
  ];
  return (
    <div style={{ background: CARD, border: `1px solid ${LINE_C}`, borderRadius: 14, padding: 18 }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <Target size={16} color={INK} />
          <span style={{ fontSize: 14, fontWeight: 600 }}>{isTeam ? "Team monthly goals" : "Monthly goals"}</span>
        </div>
        {pace != null && <span style={{ fontSize: 11.5, opacity: 0.5 }}>{Math.round(pace * 100)}% through the month</span>}
      </div>
      <div style={{ display: "grid", gap: 12 }}>
        {rows.map(([label, val, target, color]) => {
          const pct = Math.min(100, target ? (val / target) * 100 : 0);
          // Expected progress by now if pacing evenly across the month.
          const expected = pace != null && target ? target * pace : null;
          const onPace = expected != null ? val >= expected : null;
          const gap = expected != null ? Math.round(expected - val) : null;
          return (
            <div key={label}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, marginBottom: 5 }}>
                <span style={{ fontWeight: 500 }}>{label}</span>
                <span style={{ opacity: 0.6 }}>{val.toLocaleString()} / {target.toLocaleString()}</span>
              </div>
              <div style={{ height: 8, background: LINE_C, borderRadius: 99, overflow: "hidden", position: "relative" }}>
                <div style={{ width: `${pct}%`, height: "100%", background: color, borderRadius: 99, transition: "width .4s ease" }} />
                {pace != null && <div style={{ position: "absolute", top: -2, bottom: -2, left: `${Math.min(100, pace * 100)}%`, width: 2, background: INK, opacity: 0.35 }} title="Expected pace" />}
              </div>
              {onPace != null && (
                <div style={{ fontSize: 11.5, marginTop: 4, color: onPace ? "#16794C" : "#B4453F" }}>
                  {onPace ? "On pace" : `${gap > 0 ? gap.toLocaleString() : 0} behind pace`}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function Funnel({ calls, emails, appts }) {
  const outreach = calls + emails;
  const stages = [["Total outreach", outreach, INK], ["Calls placed", calls, CALL], ["Emails sent", emails, EMAIL], ["Appointments set", appts, APPT]];
  const max = Math.max(outreach, 1);
  return (
    <div style={{ display: "grid", gap: 12, paddingTop: 6 }}>
      {stages.map(([label, val, color]) => (
        <div key={label}>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, marginBottom: 5 }}>
            <span style={{ fontWeight: 500 }}>{label}</span><span style={{ opacity: 0.6 }}>{val.toLocaleString()}</span>
          </div>
          <div style={{ height: 26, background: LINE_C, borderRadius: 7, overflow: "hidden" }}>
            <div style={{ width: `${(val / max) * 100}%`, height: "100%", background: color, borderRadius: 7, minWidth: val > 0 ? 4 : 0, transition: "width .4s ease" }} />
          </div>
        </div>
      ))}
    </div>
  );
}

// Bulk upload activity from an Excel file: download template, parse, validate, preview, insert.
function BulkUpload({ liveUser, users, saveEntries, visibleUserIds }) {
  const role = liveUser.role;
  const isPrivileged = role === "admin" || role === "management";
  const salesReps = (users || []).filter((u) => u.role === "sales");
  // People this uploader may attribute rows to (BDR/Sale Rep column). Everyone can log as themselves.
  const assignable = (users || []).filter((u) => visibleUserIds.includes(u.id) && (u.role === "bdr" || u.role === "sales" || u.id === liveUser.id));

  const [rows, setRows] = useState(null);     // validated preview rows
  const [fileName, setFileName] = useState("");
  const [busy, setBusy] = useState(false);
  const [done, setDone] = useState("");
  const fileRef = React.useRef(null);

  const TEMPLATE_COLS = ["Date", "Company", "BAN", "FAN", "Contact", "Phone", "Email", "Calls", "Emails", "Appointments", "Prospecting For", "BDR/Sale Rep [You]", "Notes", "Carrier Rep"];

  const downloadTemplate = async () => {
    const N = 2000; // rows the dropdowns cover
    const example = {
      Date: TODAY_US(), Company: "Acme Corp", BAN: "123456789", FAN: "987654321", Contact: "Jane Smith",
      Phone: "(610) 555-0100", Email: "jane@acme.com", Calls: 0, Emails: 0, Appointments: 0,
      "Prospecting For": "Self-generated", "BDR/Sale Rep [You]": liveUser.name, Notes: "Intro call, follow up next week",
      "Carrier Rep": "John Doe (AT&T)",
    };
    const blank = Object.fromEntries(TEMPLATE_COLS.map((c) => [c, ""]));
    const ws = XLSX.utils.json_to_sheet([example, blank], { header: TEMPLATE_COLS });
    ws["!cols"] = [{ wch: 12 }, { wch: 22 }, { wch: 14 }, { wch: 14 }, { wch: 18 }, { wch: 16 }, { wch: 22 }, { wch: 7 }, { wch: 8 }, { wch: 13 }, { wch: 18 }, { wch: 18 }, { wch: 34 }, { wch: 20 }];
    // Force the Date column (A) to display MM-DD-YYYY for any date the user types.
    for (let row = 2; row <= 2001; row++) {
      const addr = `A${row}`;
      if (!ws[addr]) ws[addr] = { t: "s", v: "" };
      ws[addr].z = "mm-dd-yyyy";
    }

    // The name options for the two person-columns.
    const workingForOpts = ["Self-generated", ...salesReps.map((s) => s.name)];
    const repOwnerOpts = assignable.map((u) => u.name);

    // Hidden "Lists" sheet holds the option values; dropdowns reference these ranges
    // (avoids Excel's 255-char inline-list limit and any comma-in-name issues).
    const maxLen = Math.max(workingForOpts.length, repOwnerOpts.length);
    const listAoa = [["WorkingFor", "RepOwner"]];
    for (let i = 0; i < maxLen; i++) listAoa.push([workingForOpts[i] || "", repOwnerOpts[i] || ""]);
    const wsList = XLSX.utils.aoa_to_sheet(listAoa);

    // Human-readable reference sheet (visible) so users know what's valid.
    const ref = [
      { Column: "Prospecting For", "Accepted values": "Self-generated" },
      ...salesReps.map((s) => ({ Column: "Prospecting For", "Accepted values": s.name })),
      ...assignable.map((u) => ({ Column: "BDR/Sale Rep [You]", "Accepted values": `${u.name} (${ROLES[u.role].label})` })),
    ];
    const wsRef = XLSX.utils.json_to_sheet(ref);
    wsRef["!cols"] = [{ wch: 16 }, { wch: 40 }];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Activity");
    XLSX.utils.book_append_sheet(wb, wsRef, "Reference");
    XLSX.utils.book_append_sheet(wb, wsList, "Lists");

    // Hide the helper tabs so users only see the Activity sheet.
    // (They still exist so the dropdowns keep working.) hidden:1 = hidden, 2 = very hidden.
    wb.Workbook = wb.Workbook || {};
    wb.Workbook.Sheets = wb.SheetNames.map((name) => ({
      name, Hidden: name === "Activity" ? 0 : 1,
    }));

    const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });

    // Inject data-validation dropdowns into the Activity sheet (sheet1) via zip patching.
    try {
      const zip = await JSZip.loadAsync(buf);
      const path = "xl/worksheets/sheet1.xml";
      let xml = await zip.file(path).async("string");
      const wfEnd = workingForOpts.length + 1;   // Lists!$A$2:$A$n
      const roEnd = repOwnerOpts.length + 1;      // Lists!$B$2:$B$n
      const dvs = [
        { sqref: `H2:H${N + 1}`, f: '"0,1"' },      // Calls
        { sqref: `I2:I${N + 1}`, f: '"0,1"' },      // Emails
        { sqref: `J2:J${N + 1}`, f: '"0,1"' },      // Appointments
        { sqref: `K2:K${N + 1}`, f: `Lists!$A$2:$A$${wfEnd}` },  // Prospecting For
        { sqref: `L2:L${N + 1}`, f: `Lists!$B$2:$B$${roEnd}` }, // BDR/Sale Rep [You]
      ];
      const dvXml = `<dataValidations count="${dvs.length}">` +
        dvs.map((d) => `<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="${d.sqref}"><formula1>${d.f}</formula1></dataValidation>`).join("") +
        `</dataValidations>`;
      // dataValidations must appear right after </sheetData> and before any
      // ignoredErrors/pageMargins elements, or Excel flags the file as corrupt.
      if (xml.includes("</sheetData>")) {
        xml = xml.replace("</sheetData>", "</sheetData>" + dvXml);
      } else {
        xml = xml.replace("</worksheet>", dvXml + "</worksheet>");
      }
      zip.file(path, xml);
      const outBlob = await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(outBlob);
      const a = document.createElement("a");
      a.href = url; a.download = "tellemica-activity-template.xlsx";
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (e) {
      // Fallback: if patching fails for any reason, still give them the plain template.
      XLSX.writeFile(wb, "tellemica-activity-template.xlsx");
    }
  };

  // Match a typed name to a user in a candidate list (case-insensitive, trimmed).
  const matchUser = (name, list) => {
    const n = (name || "").trim().toLowerCase();
    if (!n) return null;
    return list.find((u) => u.name.trim().toLowerCase() === n) || null;
  };

  const onFile = async (file) => {
    setDone(""); setRows(null); setFileName(file.name);
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(ws, { defval: "", raw: false });
    const parsed = raw.map((r, idx) => {
      const errors = [];
      const company = String(r["Company"] || "").trim();
      if (!company) errors.push("Company is required");

      // Date: accept blank (default today) or MM-DD-YYYY (also tolerant of / and 2-digit year).
      let date = String(r["Date"] || "").trim();
      if (!date) date = TODAY();
      else {
        // Excel may hand us a serial number, an ISO string, or the MM-DD-YYYY text we ask for.
        const iso = (() => {
          // Already ISO (YYYY-MM-DD)?
          const isoM = date.match(/^(\d{4})-(\d{2})-(\d{2})$/);
          if (isoM) return date;
          // MM-DD-YYYY or MM/DD/YYYY (2- or 4-digit year)
          const m = date.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})$/);
          if (m) {
            let [, mm, dd, yy] = m;
            if (yy.length === 2) yy = "20" + yy;
            mm = mm.padStart(2, "0"); dd = dd.padStart(2, "0");
            const test = new Date(`${yy}-${mm}-${dd}T00:00`);
            if (!isNaN(test.getTime())) return `${yy}-${mm}-${dd}`;
          }
          // Fallback: let Date try (handles some Excel formats)
          const d = new Date(date);
          return isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10);
        })();
        if (!iso) errors.push(`Date "${date}" isn't in MM-DD-YYYY format`);
        else date = iso;
      }

      const numOf = (v) => { const n = parseInt(String(v).replace(/[^0-9-]/g, ""), 10); return isNaN(n) ? 0 : Math.max(0, n); };
      const calls = numOf(r["Calls"]); const emails = numOf(r["Emails"]); const appts = numOf(r["Appointments"]);

      // Prospecting For -> tagged_rep_id (a sales rep) or null for self-generated.
      let taggedRepId = null;
      const wf = String(r["Prospecting For"] || "").trim();
      if (wf && wf.toLowerCase() !== "self-generated" && wf.toLowerCase() !== "self") {
        const rep = matchUser(wf, salesReps);
        if (!rep) errors.push(`"Prospecting For" — no Sales Rep named "${wf}"`);
        else taggedRepId = rep.id;
      }

      // BDR/Sale Rep -> whose activity this row is. Default = uploader.
      let userId = liveUser.id;
      const owner = String(r["BDR/Sale Rep [You]"] || "").trim();
      if (owner) {
        const u = matchUser(owner, assignable);
        if (!u) errors.push(`"BDR/Sale Rep [You]" — "${owner}" isn't someone you can assign`);
        else if (u.id !== liveUser.id && !isPrivileged) errors.push(`Only managers/admins can assign rows to others`);
        else userId = u.id;
      }

      return {
        _row: idx + 2, // account for header row (Excel row number)
        errors,
        userId, taggedRepId, date, company,
        ban: String(r["BAN"] || "").trim(), fan: String(r["FAN"] || "").trim(), contact: String(r["Contact"] || "").trim(),
        phone: String(r["Phone"] || "").trim(), email: String(r["Email"] || "").trim(),
        calls, emails, appts, notes: String(r["Notes"] || "").trim(),
        carrierRep: String(r["Carrier Rep"] || "").trim(),
        ownerName: (assignable.find((u) => u.id === userId) || {}).name || liveUser.name,
        repName: taggedRepId ? (salesReps.find((s) => s.id === taggedRepId) || {}).name : "Self-generated",
      };
    }).filter((r) => r.company || r.errors.length); // drop fully-empty rows
    setRows(parsed);
  };

  const valid = (rows || []).filter((r) => r.errors.length === 0);
  const invalid = (rows || []).filter((r) => r.errors.length > 0);

  const doImport = async () => {
    if (valid.length === 0) return;
    setBusy(true);
    try {
      await saveEntries(() => api.addEntriesBulk(valid.map((r) => ({
        userId: r.userId, date: r.date, company: r.company, ban: r.ban, fan: r.fan, contact: r.contact,
        phone: r.phone, email: r.email, calls: r.calls, emails: r.emails, appts: r.appts,
        notes: r.notes, taggedRepId: r.taggedRepId, carrierRep: r.carrierRep,
      }))));
      setDone(`Imported ${valid.length} ${valid.length === 1 ? "record" : "records"}.`);
      setRows(null); setFileName("");
      if (fileRef.current) fileRef.current.value = "";
    } catch (e) {
      setDone("Import failed: " + (e.message || "unknown error"));
    } finally { setBusy(false); }
  };

  return (
    <div style={{ background: CARD, border: `1px solid ${LINE_C}`, borderRadius: 14, padding: 22 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
        <FileSpreadsheet size={18} color={INK} />
        <h2 style={{ fontFamily: "'Fraunces', serif", fontSize: 20, fontWeight: 600, margin: 0 }}>Bulk upload</h2>
      </div>
      <p style={{ margin: "0 0 16px", fontSize: 13.5, opacity: 0.6, lineHeight: 1.5 }}>
        Download the template, fill in one row per activity, then upload it. Every row is checked before anything is saved.
      </p>

      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 8 }}>
        <button onClick={downloadTemplate} className="tap"
          style={{ display: "flex", alignItems: "center", gap: 7, background: "#fff", color: INK, border: `1px solid ${LINE_C}`, borderRadius: 9, padding: "10px 16px", fontSize: 14, fontWeight: 600, cursor: "pointer" }}>
          <Download size={15} /> Download template
        </button>
        <button onClick={() => fileRef.current && fileRef.current.click()} className="tap"
          style={{ display: "flex", alignItems: "center", gap: 7, background: `linear-gradient(90deg, ${BTN_A}, ${BTN_B})`, color: "#fff", border: "none", borderRadius: 9, padding: "10px 16px", fontSize: 14, fontWeight: 600, cursor: "pointer" }}>
          <FileSpreadsheet size={15} /> Choose Excel file
        </button>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }}
          onChange={(e) => { const f = e.target.files && e.target.files[0]; if (f) onFile(f); }} />
      </div>
      {fileName && <div style={{ fontSize: 12.5, opacity: 0.6, marginBottom: 8 }}>Loaded: <b>{fileName}</b></div>}
      {done && <div style={{ marginTop: 10, background: CALL + "18", color: CALL, borderRadius: 8, padding: "10px 12px", fontSize: 13.5, fontWeight: 500 }}>{done}</div>}

      {rows && (
        <div style={{ marginTop: 14 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap", marginBottom: 12 }}>
            <span style={{ fontSize: 14, fontWeight: 600 }}>{valid.length} ready</span>
            {invalid.length > 0 && <span style={{ fontSize: 14, fontWeight: 600, color: "#B4453F" }}>{invalid.length} with issues</span>}
            <button onClick={doImport} disabled={busy || valid.length === 0} className="tap"
              style={{ marginLeft: "auto", background: valid.length && !busy ? INK : LINE_C, color: valid.length && !busy ? PAPER : "#8494A6", border: "none", borderRadius: 9, padding: "10px 18px", fontSize: 14, fontWeight: 600, cursor: valid.length && !busy ? "pointer" : "default" }}>
              {busy ? "Importing…" : `Import ${valid.length} ${valid.length === 1 ? "row" : "rows"}`}
            </button>
          </div>

          {invalid.length > 0 && (
            <div style={{ background: "#FBECEB", border: "1px solid #E6C9C7", borderRadius: 10, padding: "12px 14px", marginBottom: 12 }}>
              <div style={{ fontSize: 13, fontWeight: 600, color: "#B4453F", marginBottom: 6 }}>These rows won't be imported until fixed:</div>
              {invalid.slice(0, 8).map((r) => (
                <div key={r._row} style={{ fontSize: 12.5, color: "#8A3B36", marginBottom: 3 }}>
                  Row {r._row}{r.company ? ` (${r.company})` : ""}: {r.errors.join("; ")}
                </div>
              ))}
              {invalid.length > 8 && <div style={{ fontSize: 12.5, color: "#8A3B36", marginTop: 4 }}>…and {invalid.length - 8} more.</div>}
            </div>
          )}

          <div style={{ overflowX: "auto", border: `1px solid ${LINE_C}`, borderRadius: 10 }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5, minWidth: 760 }}>
              <thead>
                <tr style={{ background: "#F1F5F9" }}>
                  {["", "Date", "Company", "Calls", "Emails", "Appts", "Tellemica Sales Rep", "Owner"].map((h) => (
                    <th key={h} style={{ textAlign: "left", padding: "8px 10px", fontSize: 11, fontWeight: 700, textTransform: "uppercase", color: "#5A6B7B", whiteSpace: "nowrap" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.slice(0, 50).map((r) => (
                  <tr key={r._row} style={{ borderTop: `1px solid ${LINE_C}`, background: r.errors.length ? "#FDF4F3" : "transparent" }}>
                    <td style={{ padding: "7px 10px" }}>{r.errors.length ? <X size={14} color="#B4453F" /> : <CheckCircle2 size={14} color="#189B72" />}</td>
                    <td style={{ padding: "7px 10px", whiteSpace: "nowrap" }}>{r.date}</td>
                    <td style={{ padding: "7px 10px", fontWeight: 600 }}>{r.company || "—"}</td>
                    <td style={{ padding: "7px 10px" }}>{r.calls}</td>
                    <td style={{ padding: "7px 10px" }}>{r.emails}</td>
                    <td style={{ padding: "7px 10px" }}>{r.appts}</td>
                    <td style={{ padding: "7px 10px", whiteSpace: "nowrap" }}>{r.repName}</td>
                    <td style={{ padding: "7px 10px", whiteSpace: "nowrap" }}>{r.ownerName}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {rows.length > 50 && <div style={{ fontSize: 12, opacity: 0.5, marginTop: 6 }}>Showing first 50 of {rows.length} rows.</div>}
        </div>
      )}
    </div>
  );
}

function LogView({ liveUser, entries, saveEntries, users, allEntries, visibleUserIds }) {
  const isBDR = liveUser.role === "bdr";
  const salesReps = (users || []).filter((u) => u.role === "sales");
  // "self" sentinel = self-generated (no rep tagged). BDRs must pick; others default to self.
  const [form, setForm] = useState({ date: TODAY(), company: "", ban: "", fan: "", contact: "", phone: "", email: "", calls: "", emails: "", appts: "", notes: "", carrierRep: "", apptDate: "", apptTime: "", apptTz: liveUser.timezone || "America/New_York", apptEmail: "", workingFor: isBDR ? "" : "self" });
  const [toast, setToast] = useState(false);
  const [err, setErr] = useState("");
  const [showSuggest, setShowSuggest] = useState(false);
  // (invite is downloaded via the button in the appointment box)

  // Unique company names from everything the user can see, for autocomplete.
  const companyIndex = useMemo(() => {
    const map = new Map();
    (allEntries || entries).forEach((e) => {
      const c = (e.company || "").trim();
      if (c && !map.has(c.toLowerCase())) map.set(c.toLowerCase(), { company: c, ban: e.ban || "", fan: e.fan || "", contact: e.contact || "", phone: e.phone || "", email: e.email || "" });
    });
    return [...map.values()].sort((a, b) => a.company.localeCompare(b.company));
  }, [allEntries, entries]);

  // Remembered carrier rep names from past activity, for autocomplete.
  const carrierRepIndex = useMemo(() => {
    const seen = new Map();
    (allEntries || entries).forEach((e) => {
      const r = (e.carrierRep || "").trim();
      if (r && !seen.has(r.toLowerCase())) seen.set(r.toLowerCase(), r);
    });
    return [...seen.values()].sort((a, b) => a.localeCompare(b));
  }, [allEntries, entries]);
  const [showCarrier, setShowCarrier] = useState(false);
  const [extraContacts, setExtraContacts] = useState([]); // additional company contacts to save on submit
  const carrierSuggestions = form.carrierRep.trim().length >= 1
    ? carrierRepIndex.filter((r) => r.toLowerCase().includes(form.carrierRep.trim().toLowerCase()) && r.toLowerCase() !== form.carrierRep.trim().toLowerCase()).slice(0, 6)
    : [];

  // When an appointment is first indicated, seed the date/time with real
  // committed values (today + next half hour) so the fields show black,
  // selected values — not the browser's faded placeholder — and the invite
  // button is usable immediately without tapping each field.
  useEffect(() => {
    if ((+form.appts || 0) >= 1 && !form.apptDate && !form.apptTime) {
      const now = new Date();
      now.setMinutes(now.getMinutes() > 30 ? 60 : 30, 0, 0); // round up to next :30 or :00
      const pad = (n) => String(n).padStart(2, "0");
      setForm((f) => ({
        ...f,
        apptDate: `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`,
        apptTime: `${pad(now.getHours())}:${pad(now.getMinutes())}`,
      }));
    }
  }, [form.appts]); // eslint-disable-line

  const suggestions = form.company.trim().length >= 1
    ? companyIndex.filter((c) => c.company.toLowerCase().includes(form.company.trim().toLowerCase())).slice(0, 6)
    : [];

  const pickCompany = (c) => {
    // Reuse known details for this company, but let the user override any field.
    setForm((f) => ({ ...f, company: c.company, ban: f.ban || c.ban, fan: f.fan || c.fan, contact: f.contact || c.contact, phone: f.phone || c.phone, email: f.email || c.email }));
    setShowSuggest(false);
  };

  const submit = async () => {
    if (!form.company.trim()) { setErr("Company name is required."); return; }
    if (isBDR && !form.workingFor) { setErr("Please choose who you're working for."); return; }
    setErr("");
    const taggedRepId = form.workingFor && form.workingFor !== "self" ? form.workingFor : null;
    await saveEntries(() => api.addEntry({
      userId: liveUser.id, date: form.date,
      company: form.company.trim(), ban: form.ban.trim(), fan: form.fan.trim(), contact: form.contact.trim(),
      phone: form.phone.trim(), email: form.email.trim(),
      calls: +form.calls || 0, emails: +form.emails || 0, appts: +form.appts || 0, notes: form.notes.trim(),
      carrierRep: form.carrierRep.trim(),
      taggedRepId,
    }));
    // If an appointment was set with a date/time, ensure a deal exists for it.
    // (The invite is downloaded via the button in the appointment box.)
    if ((+form.appts || 0) >= 1 && form.apptDate && form.apptTime) {
      try {
        const start = zonedToUTC(form.apptDate, form.apptTime, form.apptTz);
        await api.upsertAppointmentDeal({
          company: form.company.trim(), contact: form.contact.trim(),
          contactEmail: form.email.trim(), apptAt: start ? start.toISOString() : null,
          ownerId: taggedRepId || liveUser.id, taggedRepId,
        });
      } catch (e) { /* deal creation is best-effort; activity already saved */ }
    }
    // Save any additional company contacts entered on this activity.
    const validExtra = extraContacts.filter((c) => c.name.trim() || c.email.trim() || c.phone.trim());
    if (validExtra.length) {
      try {
        const companyId = await api.findOrCreateCompany(form.company.trim());
        for (const c of validExtra) {
          await api.saveContact(companyId, { name: c.name.trim() || c.email.trim(), phone: c.phone.trim(), email: c.email.trim() });
        }
      } catch (e) { /* best-effort; activity already saved */ }
    }
    setForm({ ...form, company: "", ban: "", fan: "", contact: "", phone: "", email: "", calls: "", emails: "", appts: "", notes: "", carrierRep: "", apptDate: "", apptTime: "", apptEmail: "" });
    setExtraContacts([]);
    setToast(true); setTimeout(() => setToast(false), 1800);
  };

  const mine = [...entries].filter((e) => e.userId === liveUser.id)
    .sort((a, b) => (b.date < a.date ? -1 : b.date > a.date ? 1 : (b.id < a.id ? -1 : 1))).slice(0, 12);

  const del = (id) => saveEntries(() => api.deleteEntry(id));
  const repName = (id) => { const u = (users || []).find((x) => x.id === id); return u ? u.name : null; };

  return (
    <>
    <div className="logwrap" style={{ display: "grid", gridTemplateColumns: "minmax(0,1fr) minmax(0,1.2fr)", gap: 20, alignItems: "start" }}>
      <div style={{ background: CARD, border: `1px solid ${LINE_C}`, borderRadius: 14, padding: 22 }}>
        <h2 style={{ fontFamily: "'Fraunces', serif", fontSize: 22, fontWeight: 600, margin: "0 0 4px" }}>Log a session</h2>
        <p style={{ margin: "0 0 20px", fontSize: 13.5, opacity: 0.55 }}>Logging as <b>{liveUser.name}</b> · {ROLES[liveUser.role].label}</p>
        <div style={{ marginBottom: 18 }}>
          <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: 0.4, textTransform: "uppercase", color: "#8494A6", marginBottom: 8 }}>Quick fill</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 7 }}>
            {[
              { label: "+1 Call", apply: (f) => ({ ...f, calls: String((+f.calls || 0) + 1) }) },
              { label: "+1 Email", apply: (f) => ({ ...f, emails: String((+f.emails || 0) + 1) }) },
              { label: "+1 Appt", apply: (f) => ({ ...f, appts: String((+f.appts || 0) + 1) }) },
              { label: "Left voicemail", apply: (f) => ({ ...f, calls: String((+f.calls || 0) + 1), notes: f.notes ? f.notes : "Left voicemail" }) },
              { label: "Sent follow-up", apply: (f) => ({ ...f, emails: String((+f.emails || 0) + 1), notes: f.notes ? f.notes : "Sent follow-up email" }) },
              { label: "Booked appt", apply: (f) => ({ ...f, appts: String((+f.appts || 0) + 1), notes: f.notes ? f.notes : "Booked appointment" }) },
            ].map((p) => (
              <button key={p.label} type="button" onClick={() => setForm((f) => p.apply(f))} className="tap"
                style={{ background: "#F1F5F9", border: `1px solid ${LINE_C}`, borderRadius: 8, padding: "6px 11px", fontSize: 12.5, fontWeight: 600, color: INK, cursor: "pointer" }}>
                {p.label}
              </button>
            ))}
          </div>
        </div>
        {isBDR && (
          <Field label="Tellemica Sales Rep">
            <div style={{ position: "relative" }}>
              <select value={form.workingFor} onChange={(e) => { setForm({ ...form, workingFor: e.target.value }); setErr(""); }}
                style={{ ...inputStyle, appearance: "none", cursor: "pointer" }}>
                <option value="" disabled>Choose a Sales Rep…</option>
                <option value="self">Self-generated</option>
                {salesReps.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
              <ChevronDown size={15} style={{ position: "absolute", right: 12, top: 12, pointerEvents: "none", opacity: 0.5 }} />
            </div>
          </Field>
        )}
        <Field label="Company name *">
          <div style={{ position: "relative" }}>
            <input value={form.company} autoComplete="off"
              onChange={(e) => { setForm({ ...form, company: e.target.value }); setShowSuggest(true); setErr(""); }}
              onFocus={() => setShowSuggest(true)}
              onBlur={() => setTimeout(() => setShowSuggest(false), 150)}
              style={inputStyle} placeholder="Start typing to search existing companies…" />
            {showSuggest && suggestions.length > 0 && (
              <div style={{ position: "absolute", top: "100%", left: 0, right: 0, zIndex: 20, background: "#fff", border: `1px solid ${LINE_C}`, borderRadius: 10, marginTop: 4, boxShadow: "0 10px 30px rgba(8,20,40,.15)", overflow: "hidden" }}>
                {suggestions.map((c) => (
                  <button key={c.company} type="button" onMouseDown={(ev) => { ev.preventDefault(); pickCompany(c); }} className="tap"
                    style={{ display: "block", width: "100%", textAlign: "left", padding: "10px 12px", background: "transparent", border: "none", borderBottom: `1px solid ${LINE_C}`, cursor: "pointer", fontSize: 14 }}>
                    <span style={{ fontWeight: 600 }}>{c.company}</span>
                    {c.contact ? <span style={{ opacity: 0.55, fontSize: 12.5 }}> · {c.contact}</span> : null}
                  </button>
                ))}
              </div>
            )}
          </div>
        </Field>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          <Field label="BAN"><input value={form.ban} onChange={(e) => setForm({ ...form, ban: e.target.value })} style={inputStyle} placeholder="Billing account #" /></Field>
          <Field label="FAN"><input value={form.fan} onChange={(e) => setForm({ ...form, fan: e.target.value })} style={inputStyle} placeholder="Foundation account #" /></Field>
        </div>
        <Field label="Contact"><input value={form.contact} onChange={(e) => setForm({ ...form, contact: e.target.value })} style={inputStyle} placeholder="Name / title" /></Field>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          <Field label="Phone"><input value={form.phone} onChange={(e) => setForm({ ...form, phone: formatPhone(e.target.value) })} style={inputStyle} placeholder="(610) 555-0100" /></Field>
          <Field label="Email"><input value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })} style={inputStyle} placeholder="name@company.com" /></Field>
        </div>
        <div style={{ marginBottom: 14 }}>
          {extraContacts.map((c, i) => (
            <div key={i} style={{ background: "#F8FAFC", border: `1px solid ${LINE_C}`, borderRadius: 9, padding: 10, marginBottom: 8 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                <span style={{ fontSize: 11.5, fontWeight: 700, letterSpacing: 0.4, textTransform: "uppercase", color: "#8494A6" }}>Additional contact {i + 1}</span>
                <button type="button" onClick={() => setExtraContacts((arr) => arr.filter((_, j) => j !== i))} className="tap" style={{ background: "transparent", border: "none", cursor: "pointer", opacity: 0.5, display: "flex" }}><X size={14} /></button>
              </div>
              <input value={c.name} onChange={(e) => setExtraContacts((arr) => arr.map((x, j) => j === i ? { ...x, name: e.target.value } : x))} style={{ ...inputStyle, marginBottom: 8 }} placeholder="Name / title" />
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                <input value={c.phone} onChange={(e) => setExtraContacts((arr) => arr.map((x, j) => j === i ? { ...x, phone: formatPhone(e.target.value) } : x))} style={{ ...inputStyle, marginBottom: 0 }} placeholder="Phone" />
                <input value={c.email} onChange={(e) => setExtraContacts((arr) => arr.map((x, j) => j === i ? { ...x, email: e.target.value } : x))} style={{ ...inputStyle, marginBottom: 0 }} placeholder="Email" />
              </div>
            </div>
          ))}
          <button type="button" onClick={() => setExtraContacts((arr) => [...arr, { name: "", phone: "", email: "" }])} className="tap"
            style={{ display: "flex", alignItems: "center", gap: 6, background: "transparent", border: `1px dashed ${LINE_C}`, borderRadius: 8, padding: "8px 12px", fontSize: 12.5, fontWeight: 600, color: EMAIL, cursor: "pointer", width: "100%", justifyContent: "center" }}>
            <Plus size={13} /> Add another contact
          </button>
        </div>
        <Field label="Carrier Rep (AT&T, VZW, TMo)">
          <div style={{ position: "relative" }}>
            <input value={form.carrierRep}
              onChange={(e) => { setForm({ ...form, carrierRep: e.target.value }); setShowCarrier(true); }}
              onFocus={() => setShowCarrier(true)}
              onBlur={() => setTimeout(() => setShowCarrier(false), 150)}
              style={inputStyle} placeholder="Name of carrier rep" />
            {showCarrier && carrierSuggestions.length > 0 && (
              <div style={{ position: "absolute", top: "calc(100% + 2px)", left: 0, right: 0, background: CARD, border: `1px solid ${LINE_C}`, borderRadius: 9, boxShadow: "0 10px 28px rgba(11,42,74,.12)", zIndex: 30, overflow: "hidden" }}>
                {carrierSuggestions.map((r) => (
                  <button key={r} type="button" onMouseDown={(ev) => { ev.preventDefault(); setForm((f) => ({ ...f, carrierRep: r })); setShowCarrier(false); }} className="tap"
                    style={{ display: "block", width: "100%", textAlign: "left", background: "transparent", border: "none", padding: "9px 13px", fontSize: 13.5, cursor: "pointer" }}>
                    {r}
                  </button>
                ))}
              </div>
            )}
          </div>
        </Field>
        <Field label="Date"><input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} style={inputStyle} /></Field>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 }}>
          <Field label="Calls"><input type="number" min="0" placeholder="0" value={form.calls} onChange={(e) => setForm({ ...form, calls: e.target.value })} style={inputStyle} /></Field>
          <Field label="Emails"><input type="number" min="0" placeholder="0" value={form.emails} onChange={(e) => setForm({ ...form, emails: e.target.value })} style={inputStyle} /></Field>
          <Field label="Appts set"><input type="number" min="0" placeholder="0" value={form.appts} onChange={(e) => setForm({ ...form, appts: e.target.value })} style={inputStyle} /></Field>
        </div>
        {(+form.appts || 0) >= 1 && (
          <div style={{ background: APPT + "10", border: `1px solid ${APPT}33`, borderRadius: 10, padding: 14, marginBottom: 14 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 10, fontSize: 13, fontWeight: 600, color: "#7A5C1E" }}>
              <CalendarCheck size={15} /> Appointment details
            </div>
            <p style={{ fontSize: 12, opacity: 0.6, margin: "0 0 10px" }}>Set a date, time, and timezone, then download the Outlook invite below. The invite goes to the <b>Email</b> above{form.email.trim() ? <> (<span style={{ color: EMAIL }}>{form.email.trim()}</span>)</> : <span style={{ color: "#B4453F" }}> — none entered yet, add one above</span>}, plus you and your manager.</p>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
              <Field label="Appointment date"><input type="date" value={form.apptDate} onChange={(e) => setForm({ ...form, apptDate: e.target.value })} style={inputStyle} /></Field>
              <Field label="Time">
                {(() => {
                  const t = parseTime12(form.apptTime);
                  const setPart = (part, val) => setForm((f) => ({ ...f, apptTime: buildTime24({ ...parseTime12(f.apptTime), [part]: val }) }));
                  const selStyle = { ...inputStyle, marginBottom: 0, appearance: "none", cursor: "pointer", paddingRight: 26 };
                  const wrap = { position: "relative", flex: 1 };
                  const chev = { position: "absolute", right: 8, top: 12, pointerEvents: "none", opacity: 0.5 };
                  return (
                    <div style={{ display: "flex", gap: 6 }}>
                      <div style={wrap}>
                        <select value={t.h12} onChange={(e) => setPart("h12", e.target.value)} style={selStyle}>
                          <option value="" disabled>Hr</option>
                          {Array.from({ length: 12 }, (_, i) => String(i + 1)).map((h) => <option key={h} value={h}>{h}</option>)}
                        </select>
                        <ChevronDown size={13} style={chev} />
                      </div>
                      <div style={wrap}>
                        <select value={t.min} onChange={(e) => setPart("min", e.target.value)} style={selStyle}>
                          <option value="" disabled>Min</option>
                          <option value="00">00</option>
                          <option value="30">30</option>
                        </select>
                        <ChevronDown size={13} style={chev} />
                      </div>
                      <div style={wrap}>
                        <select value={t.ap} onChange={(e) => setPart("ap", e.target.value)} style={selStyle}>
                          <option value="" disabled>AM/PM</option>
                          <option value="AM">AM</option>
                          <option value="PM">PM</option>
                        </select>
                        <ChevronDown size={13} style={chev} />
                      </div>
                    </div>
                  );
                })()}
              </Field>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
              <Field label="Timezone">
                <div style={{ position: "relative" }}>
                  <select value={form.apptTz} onChange={(e) => { const tz = e.target.value; setForm({ ...form, apptTz: tz }); if (tz !== liveUser.timezone) { api.updateProfile(liveUser.id, { timezone: tz }).catch(() => {}); liveUser.timezone = tz; } }}
                    style={{ ...inputStyle, appearance: "none", cursor: "pointer" }}>
                    {US_TIMEZONES.map((z) => <option key={z.id} value={z.id}>{z.label}</option>)}
                  </select>
                  <ChevronDown size={15} style={{ position: "absolute", right: 12, top: 12, pointerEvents: "none", opacity: 0.5 }} />
                </div>
              </Field>
              <Field label="Additional invitees (optional)"><input type="text" value={form.apptEmail} onChange={(e) => setForm({ ...form, apptEmail: e.target.value })} style={inputStyle} placeholder="other@company.com, ..." /></Field>
            </div>
            {(() => {
              const ready = form.apptDate && form.apptTime && form.email.trim();
              const buildStart = () => zonedToUTC(form.apptDate, form.apptTime, form.apptTz);
              const rep = () => (users || []).find((u) => u.id === (form.workingFor && form.workingFor !== "self" ? form.workingFor : liveUser.id)) || liveUser;
              const mgr = (r) => (users || []).find((u) => u.id === r.managerId);
              const dealObj = (startISO) => ({ id: "appt", company: form.company.trim(), contact: form.contact.trim(), contactEmail: form.email.trim(), value: 0, apptAt: startISO });
              return (
                <>
                  <button type="button" disabled={!ready}
                    onClick={() => { const start = buildStart(); if (!start) return; const r = rep(); window.open(outlookDeeplink(dealObj(start.toISOString()), r, mgr(r), form.apptEmail.trim(), start), "_blank"); }}
                    className="tap"
                    style={{ width: "100%", marginTop: 4, display: "flex", alignItems: "center", justifyContent: "center", gap: 8,
                      background: ready ? INK : LINE_C, color: ready ? PAPER : "#8494A6", border: "none", borderRadius: 9, padding: "10px 14px", fontSize: 13.5, fontWeight: 600, cursor: ready ? "pointer" : "default" }}>
                    <CalendarCheck size={15} /> Open invite in Outlook
                  </button>
                  <button type="button" disabled={!ready}
                    onClick={() => { const start = buildStart(); if (!start) return; const r = rep(); downloadAppointmentICS(dealObj(start.toISOString()), r, mgr(r), form.apptEmail.trim()); }}
                    className="tap"
                    style={{ width: "100%", marginTop: 8, background: "transparent", color: ready ? EMAIL : "#8494A6", border: `1px solid ${ready ? EMAIL + "55" : LINE_C}`, borderRadius: 9, padding: "8px 14px", fontSize: 12.5, fontWeight: 600, cursor: ready ? "pointer" : "default" }}>
                    or download .ics file
                  </button>
                </>
              );
            })()}
            {(!form.apptDate || !form.apptTime || !form.email.trim()) && <p style={{ fontSize: 11.5, opacity: 0.5, margin: "6px 0 0", textAlign: "center" }}>Add a date, time, and the customer's Email above to enable.</p>}
          </div>
        )}
        <Field label="Notes (optional)"><textarea rows={2} placeholder="Prospect, company, follow-up…" value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} style={{ ...inputStyle, resize: "vertical" }} /></Field>
        {err && <div style={{ color: "#B4453F", fontSize: 13, marginBottom: 10 }}>{err}</div>}
        <button onClick={submit} className="tap"
          style={{ width: "100%", marginTop: 8, background: `linear-gradient(90deg, ${BTN_A}, ${BTN_B})`, color: "#fff", border: "none", borderRadius: 10, padding: 13, fontSize: 15, fontWeight: 600, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}>
          <Plus size={17} /> Save session
        </button>
        {toast && <div style={{ marginTop: 12, background: CALL + "18", color: CALL, borderRadius: 8, padding: "10px 12px", fontSize: 13.5, fontWeight: 500, textAlign: "center" }}>Saved — dashboard updated.</div>}
      </div>
      <div style={{ background: CARD, border: `1px solid ${LINE_C}`, borderRadius: 14, padding: 22 }}>
        <h2 style={{ fontFamily: "'Fraunces', serif", fontSize: 20, fontWeight: 600, margin: "0 0 16px" }}>My recent activity</h2>
        {mine.length === 0 ? <Empty msg="Nothing logged yet. Your saved sessions appear here." /> : (
          <div>
            {mine.map((e) => (
              <div key={e.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "11px 0", borderBottom: `1px solid ${LINE_C}` }}>
                <div style={{ minWidth: 0 }}>
                  <div style={{ fontSize: 13, fontWeight: 500 }}>{new Date(e.date + "T00:00").toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" })}{isBDR ? ` · ${e.taggedRepId ? repName(e.taggedRepId) || "Rep" : "Self-generated"}` : ""}</div>
                  {e.notes && <div style={{ fontSize: 12, opacity: 0.5 }}>{e.notes}</div>}
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
                  <Pill icon={Phone} color={CALL} n={e.calls} />
                  <Pill icon={Mail} color={EMAIL} n={e.emails} />
                  <Pill icon={CalendarCheck} color={APPT} n={e.appts} />
                  <button onClick={() => del(e.id)} className="tap" style={{ background: "transparent", border: "none", cursor: "pointer", opacity: 0.4 }}><Trash2 size={15} /></button>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
    <div style={{ marginTop: 20 }}>
      <BulkUpload liveUser={liveUser} users={users} saveEntries={saveEntries} visibleUserIds={visibleUserIds} />
    </div>
    </>
  );
}

function Pill({ icon: Icon, color, n }) {
  return <span style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 13, fontWeight: 600, color, minWidth: 34 }}><Icon size={14} /> {n || 0}</span>;
}

// ---- Pipeline (CRM) ----
function Pipeline({ deals, allDeals, saveDeals, liveUser, users, visibleUserIds, entries, saveEntries }) {
  const role = liveUser.role;
  const canEdit = role === "bdr" || role === "sales" || role === "admin" || role === "management"; // reps own deals; admins/managers can also manage
  const [mode, setMode] = useState("board"); // board | table
  const [modal, setModal] = useState(null); // {deal} or {deal:null}
  const [ownerFilter, setOwnerFilter] = useState("all");
  const [sortKey, setSortKey] = useState("value");

  const ownerName = (id) => { const u = users.find((x) => x.id === id); return u ? u.name : "—"; };
  const repUsers = users.filter((u) => visibleUserIds.includes(u.id) && (u.role === "bdr" || u.role === "sales"));

  const shown = ownerFilter === "all" ? deals : deals.filter((d) => d.ownerId === ownerFilter);

  // When a deal first reaches "Appointment Set", credit one appointment to the owner's
  // activity (once per deal). Returns extra fields to persist on the deal.
  const maybeCreditAppt = async (deal, prevStage) => {
    if (deal.stage === "appointment" && !deal.apptCredited && prevStage !== "appointment") {
      await api.addEntry({
        userId: deal.ownerId, date: TODAY(),
        calls: 0, emails: 0, appts: 1,
        notes: `Auto: appointment set — ${deal.company}`, fromDeal: deal.id || null,
      });
      return true;
    }
    return deal.apptCredited || false;
  };

  const upsertDeal = (data) => saveDeals(async () => {
    if (data.id) {
      const prev = allDeals.find((d) => d.id === data.id);
      const stageChanged = prev && data.stage && data.stage !== prev.stage;
      const merged = { ...prev, ...data, ...(stageChanged ? { stageChangedAt: new Date().toISOString() } : {}) };
      // Save the deal first so we have an id, then handle credit + credited flag.
      const saved = await api.upsertDeal(merged);
      const credited = await maybeCreditAppt(saved, prev?.stage);
      if (credited && !saved.apptCredited) await api.updateDeal(saved.id, { apptCredited: true });
      // Best-effort history log — never let it block the save.
      if (stageChanged) { try { await api.logStageChange(saved.id, prev.stage, saved.stage); } catch { /* ignore */ } }
    } else {
      const created = await api.upsertDeal({ ...data, ownerId: data.ownerId || liveUser.id, stageChangedAt: new Date().toISOString() });
      const credited = await maybeCreditAppt(created, null);
      if (credited) await api.updateDeal(created.id, { apptCredited: true });
      try { await api.logStageChange(created.id, null, created.stage); } catch { /* ignore */ }
    }
  }).then(() => setModal(null));

  const moveDeal = (dealId, newStage) => {
    const prev = allDeals.find((d) => d.id === dealId);
    if (!prev || prev.stage === newStage) return;
    return saveDeals(async () => {
      const saved = await api.updateDeal(dealId, { ...prev, stage: newStage, stageChangedAt: new Date().toISOString() });
      const credited = await maybeCreditAppt(saved, prev.stage);
      if (credited && !saved.apptCredited) await api.updateDeal(dealId, { apptCredited: true });
      try { await api.logStageChange(dealId, prev.stage, newStage); } catch { /* ignore */ }
    });
  };

  const removeDeal = (id) => {
    if (!confirm("Delete this deal?")) return;
    return saveDeals(() => api.deleteDeal(id));
  };

  // drag + drop
  const [dragId, setDragId] = useState(null);

  const sortedTable = [...shown].sort((a, b) => {
    if (sortKey === "value") return (b.value || 0) - (a.value || 0);
    if (sortKey === "close") return (a.closeDate || "9999").localeCompare(b.closeDate || "9999");
    if (sortKey === "stage") return STAGES.findIndex((s) => s.id === a.stage) - STAGES.findIndex((s) => s.id === b.stage);
    return (a.company || "").localeCompare(b.company || "");
  });

  return (
    <>
      <style>{`@media (max-width: 720px){ .board{grid-auto-columns:80% !important} }`}</style>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12, marginBottom: 20 }}>
        <div>
          <h1 style={{ fontFamily: "'Fraunces', serif", fontSize: 29, fontWeight: 600, margin: 0 }}>Sales Pipeline</h1>
          <p style={{ margin: "4px 0 0", opacity: 0.55, fontSize: 14 }}>
            {role === "bdr" ? "Your deals" : role === "sales" ? "You + your BDRs" : "All deals"} · {shown.length} in view
          </p>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          {repUsers.length > 1 && (
            <div style={{ position: "relative" }}>
              <select value={ownerFilter} onChange={(e) => setOwnerFilter(e.target.value)}
                style={{ appearance: "none", background: CARD, border: `1px solid ${LINE_C}`, borderRadius: 9, padding: "9px 32px 9px 12px", fontSize: 14, fontWeight: 500, color: INK, cursor: "pointer" }}>
                <option value="all">All owners</option>
                {repUsers.map((u) => <option key={u.id} value={u.id}>{u.name}</option>)}
              </select>
              <ChevronDown size={15} style={{ position: "absolute", right: 10, top: 11, pointerEvents: "none", opacity: 0.5 }} />
            </div>
          )}
          <div style={{ display: "flex", background: CARD, border: `1px solid ${LINE_C}`, borderRadius: 9, padding: 3 }}>
            <button onClick={() => setMode("board")} className="tap" style={{ display: "flex", alignItems: "center", gap: 5, border: "none", borderRadius: 6, padding: "7px 11px", fontSize: 13, fontWeight: 500, cursor: "pointer", background: mode === "board" ? INK : "transparent", color: mode === "board" ? PAPER : INK }}><Kanban size={14} /> Board</button>
            <button onClick={() => setMode("table")} className="tap" style={{ display: "flex", alignItems: "center", gap: 5, border: "none", borderRadius: 6, padding: "7px 11px", fontSize: 13, fontWeight: 500, cursor: "pointer", background: mode === "table" ? INK : "transparent", color: mode === "table" ? PAPER : INK }}><Table2 size={14} /> Table</button>
          </div>
          {canEdit && (
            <button onClick={() => setModal({ deal: null })} className="tap"
              style={{ display: "flex", alignItems: "center", gap: 6, background: `linear-gradient(90deg, ${BTN_A}, ${BTN_B})`, color: "#fff", border: "none", borderRadius: 9, padding: "10px 15px", fontSize: 14, fontWeight: 600, cursor: "pointer" }}>
              <Plus size={16} /> New deal
            </button>
          )}
        </div>
      </div>

      {shown.length === 0 ? (
        <Panel title="No deals yet">
          <Empty msg={canEdit ? "Click 'New deal' to add your first opportunity." : "No deals in your scope yet."} />
        </Panel>
      ) : mode === "board" ? (
        <div className="board" style={{ display: "grid", gridAutoFlow: "column", gridAutoColumns: "minmax(230px, 1fr)", gap: 12, overflowX: "auto", paddingBottom: 8 }}>
          {STAGES.map((s) => {
            const col = shown.filter((d) => d.stage === s.id);
            const colVal = col.reduce((sum, d) => sum + (d.value || 0), 0);
            return (
              <div key={s.id}
                onDragOver={(e) => { if (canEdit) e.preventDefault(); }}
                onDrop={() => { if (canEdit && dragId) { moveDeal(dragId, s.id); setDragId(null); } }}
                style={{ background: "rgba(255,255,255,.5)", border: `1px solid ${LINE_C}`, borderRadius: 12, padding: 10, minHeight: 120 }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10, paddingBottom: 8, borderBottom: `2px solid ${s.color}` }}>
                  <span style={{ fontSize: 13, fontWeight: 600, display: "flex", alignItems: "center", gap: 6 }}>
                    <span style={{ width: 8, height: 8, borderRadius: "50%", background: s.color }} />{s.label}
                  </span>
                  <span style={{ fontSize: 11, opacity: 0.55 }}>{col.length}</span>
                </div>
                <div style={{ fontSize: 11.5, opacity: 0.6, marginBottom: 8 }}>{fmtMoney(colVal)}</div>
                <div style={{ display: "grid", gap: 8 }}>
                  {col.map((d) => (
                    <div key={d.id} draggable={canEdit}
                      onDragStart={() => setDragId(d.id)}
                      onClick={() => canEdit && setModal({ deal: d })}
                      className="tap"
                      style={{ background: CARD, border: `1px solid ${LINE_C}`, borderRadius: 9, padding: 11, cursor: canEdit ? "grab" : "default" }}>
                      <div style={{ fontSize: 13.5, fontWeight: 600, marginBottom: 3 }}>{d.company}</div>
                      {d.contact && <div style={{ fontSize: 12, opacity: 0.55, marginBottom: 6 }}>{d.contact}</div>}
                      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                        <span style={{ fontSize: 13, fontWeight: 700, color: s.color }}>{fmtMoney(d.value)}</span>
                        {ownerFilter === "all" && repUsers.length > 1 && <span style={{ fontSize: 10.5, opacity: 0.5 }}>{ownerName(d.ownerId).split(" ")[0]}</span>}
                      </div>
                      {d.closeDate && <div style={{ fontSize: 11, opacity: 0.5, marginTop: 5 }}>Close: {new Date(d.closeDate + "T00:00").toLocaleDateString("en-US", { month: "short", day: "numeric" })}</div>}
                      {(() => {
                        const fu = followUpState(d.nextActionDate);
                        const ag = dealAge(d);
                        return (
                          <div style={{ display: "flex", flexWrap: "wrap", gap: 5, marginTop: 6 }}>
                            {fu && <span style={{ fontSize: 10.5, fontWeight: 600, color: fu.color, background: fu.color + "16", borderRadius: 5, padding: "2px 6px" }}>⏱ {fu.label}</span>}
                            {ag && ag.stale && <span style={{ fontSize: 10.5, fontWeight: 600, color: "#8A6D3B", background: "#FBF3E2", borderRadius: 5, padding: "2px 6px" }}>⚠ {ag.age}d in stage</span>}
                          </div>
                        );
                      })()}
                      {d.stage === "appointment" && d.apptAt && (
                        <button onClick={(ev) => { ev.stopPropagation(); const rep = users.find((u) => u.id === d.ownerId); const mgr = users.find((u) => u.id === rep?.managerId); downloadAppointmentICS(d, rep, mgr); }} className="tap"
                          style={{ display: "flex", alignItems: "center", gap: 5, marginTop: 7, background: CYAN + "18", color: "#0B6A8C", border: "none", borderRadius: 6, padding: "4px 8px", fontSize: 11, fontWeight: 600, cursor: "pointer" }}>
                          <CalendarCheck size={12} /> Add to calendar
                        </button>
                      )}
                    </div>
                  ))}
                  {col.length === 0 && <div style={{ fontSize: 12, opacity: 0.3, textAlign: "center", padding: "12px 0" }}>—</div>}
                </div>
              </div>
            );
          })}
        </div>
      ) : (
        <Panel title="" style={{ padding: 0, overflow: "hidden" }}>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13.5 }}>
              <thead>
                <tr style={{ background: "rgba(11,42,74,.04)" }}>
                  {[["company", "Company"], ["stage", "Stage"], ["value", "Value"], ["close", "Close date"]].map(([k, lbl]) => (
                    <th key={k} onClick={() => setSortKey(k)} className="tap"
                      style={{ textAlign: "left", padding: "12px 16px", fontWeight: 600, cursor: "pointer", whiteSpace: "nowrap", color: sortKey === k ? CYAN : INK }}>
                      {lbl}{sortKey === k ? " ↓" : ""}
                    </th>
                  ))}
                  {ownerFilter === "all" && <th style={{ textAlign: "left", padding: "12px 16px", fontWeight: 600 }}>Owner</th>}
                  {canEdit && <th style={{ padding: "12px 16px" }}></th>}
                </tr>
              </thead>
              <tbody>
                {sortedTable.map((d) => {
                  const s = STAGE[d.stage];
                  return (
                    <tr key={d.id} style={{ borderTop: `1px solid ${LINE_C}` }}>
                      <td style={{ padding: "12px 16px" }}>
                        <div style={{ fontWeight: 600 }}>{d.company}</div>
                        {d.contact && <div style={{ fontSize: 12, opacity: 0.5 }}>{d.contact}</div>}
                      </td>
                      <td style={{ padding: "12px 16px" }}>
                        <span style={{ display: "inline-flex", alignItems: "center", gap: 6, background: s.color + "18", color: s.color, borderRadius: 20, padding: "4px 10px", fontSize: 12, fontWeight: 600, whiteSpace: "nowrap" }}>
                          <span style={{ width: 7, height: 7, borderRadius: "50%", background: s.color }} />{s.label}
                        </span>
                        {d.stage === "lost" && d.lostReason && <div style={{ fontSize: 11.5, opacity: 0.6, marginTop: 4 }}>{d.lostReason}</div>}
                      </td>
                      <td style={{ padding: "12px 16px", fontWeight: 600, whiteSpace: "nowrap" }}>{fmtMoney(d.value)}</td>
                      <td style={{ padding: "12px 16px", opacity: 0.7, whiteSpace: "nowrap" }}>{d.closeDate ? new Date(d.closeDate + "T00:00").toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" }) : "—"}</td>
                      {ownerFilter === "all" && <td style={{ padding: "12px 16px", opacity: 0.7 }}>{ownerName(d.ownerId)}</td>}
                      {canEdit && (
                        <td style={{ padding: "12px 16px", whiteSpace: "nowrap", textAlign: "right" }}>
                          <button onClick={() => setModal({ deal: d })} className="tap" style={iconBtn}><Pencil size={14} /></button>
                          <button onClick={() => removeDeal(d.id)} className="tap" style={{ ...iconBtn, marginLeft: 4 }}><Trash2 size={14} /></button>
                        </td>
                      )}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </Panel>
      )}

      {canEdit && mode === "board" && (
        <p style={{ fontSize: 12.5, opacity: 0.5, marginTop: 12, textAlign: "center" }}>Drag a card between columns to move a deal. Click a card to edit it.</p>
      )}

      {modal && (
        <DealModal deal={modal.deal} onSave={upsertDeal} onDelete={removeDeal} onClose={() => setModal(null)}
          liveUser={liveUser} salesReps={users.filter((u) => u.role === "sales")}
          assignableOwners={repUsers} dealUsers={users} />
      )}
    </>
  );
}

function DealModal({ deal, onSave, onDelete, onClose, liveUser, salesReps, assignableOwners, dealUsers }) {
  const isBDR = liveUser && liveUser.role === "bdr";
  const isAdminMgr = liveUser && (liveUser.role === "admin" || liveUser.role === "management");
  const [f, setF] = useState({ company: "", contact: "", contactEmail: "", value: "", stage: "new", closeDate: "", notes: "", lostReason: "", nextActionDate: "", ownerId: "", taggedRepId: isBDR ? "" : "self", ...(deal || {}) });
  const [err, setErr] = useState("");
  const owners = assignableOwners || [];
  const [history, setHistory] = useState(null);
  useEffect(() => {
    let alive = true;
    if (deal?.id) {
      api.listStageHistory(deal.id).then((h) => { if (alive) setHistory(h); }).catch(() => { if (alive) setHistory([]); });
    }
    return () => { alive = false; };
  }, [deal?.id]);
  const submit = () => {
    if (!f.company.trim()) { setErr("Company / prospect name is required."); return; }
    if (isBDR && !f.taggedRepId) { setErr("Please choose who you're working for."); return; }
    // Admins/managers must pick an owner when creating a brand-new deal.
    if (isAdminMgr && !deal?.id && !f.ownerId) { setErr("Please choose which rep owns this deal."); return; }
    if (f.stage === "lost" && !String(f.lostReason || "").trim()) { setErr("Please choose a reason for the loss."); return; }
    const taggedRepId = f.taggedRepId && f.taggedRepId !== "self" ? f.taggedRepId : null;
    // Clear the lost reason if the deal isn't actually lost.
    const lostReason = f.stage === "lost" ? String(f.lostReason || "").trim() : "";
    onSave({ ...f, company: f.company.trim(), contact: f.contact.trim(), contactEmail: (f.contactEmail || "").trim(), value: +f.value || 0, taggedRepId, lostReason, ...(deal?.id ? { id: deal.id } : {}) });
  };
  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(11,42,74,.5)", display: "grid", placeItems: "center", padding: 20, zIndex: 50 }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: PAPER, borderRadius: 16, padding: 24, width: "100%", maxWidth: 460, maxHeight: "90vh", overflowY: "auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18 }}>
          <h3 style={{ fontFamily: "'Fraunces', serif", fontSize: 20, fontWeight: 600, margin: 0 }}>{deal ? "Edit deal" : "New deal"}</h3>
          <button onClick={onClose} className="tap" style={{ background: "transparent", border: "none", cursor: "pointer" }}><X size={20} /></button>
        </div>
        <Field label="Company / prospect"><input value={f.company} onChange={(e) => setF({ ...f, company: e.target.value })} style={inputStyle} placeholder="Acme Corp" autoFocus /></Field>
        {isAdminMgr && (
          <Field label="Deal owner">
            <div style={{ position: "relative" }}>
              <select value={f.ownerId || ""} onChange={(e) => { setF({ ...f, ownerId: e.target.value }); setErr(""); }}
                style={{ ...inputStyle, appearance: "none", cursor: "pointer" }}>
                <option value="" disabled>Choose the rep who owns this…</option>
                {owners.map((u) => <option key={u.id} value={u.id}>{u.name}</option>)}
              </select>
              <ChevronDown size={15} style={{ position: "absolute", right: 12, top: 12, pointerEvents: "none", opacity: 0.5 }} />
            </div>
          </Field>
        )}
        {isBDR && (
          <Field label="Tellemica Sales Rep">
            <div style={{ position: "relative" }}>
              <select value={f.taggedRepId} onChange={(e) => { setF({ ...f, taggedRepId: e.target.value }); setErr(""); }}
                style={{ ...inputStyle, appearance: "none", cursor: "pointer" }}>
                <option value="" disabled>Choose a Sales Rep…</option>
                <option value="self">Self-generated</option>
                {(salesReps || []).map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
              <ChevronDown size={15} style={{ position: "absolute", right: 12, top: 12, pointerEvents: "none", opacity: 0.5 }} />
            </div>
          </Field>
        )}
        <Field label="Contact (optional)"><input value={f.contact} onChange={(e) => setF({ ...f, contact: e.target.value })} style={inputStyle} placeholder="Name, title, phone/email" /></Field>
        <Field label="Contact email (for calendar invites)"><input type="email" value={f.contactEmail || ""} onChange={(e) => setF({ ...f, contactEmail: e.target.value })} style={inputStyle} placeholder="contact@company.com" /></Field>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          <Field label="Deal value ($)"><input type="number" min="0" value={f.value} onChange={(e) => setF({ ...f, value: e.target.value })} style={inputStyle} placeholder="0" /></Field>
          <Field label="Expected close"><input type="date" value={f.closeDate} onChange={(e) => setF({ ...f, closeDate: e.target.value })} style={inputStyle} /></Field>
        </div>
        <Field label="Next action / follow-up date"><input type="date" value={f.nextActionDate || ""} onChange={(e) => setF({ ...f, nextActionDate: e.target.value })} style={inputStyle} /></Field>
        <Field label="Appointment date & time">
          <input type="datetime-local" value={f.apptAt ? toLocalInput(f.apptAt) : ""} onChange={(e) => setF({ ...f, apptAt: e.target.value ? new Date(e.target.value).toISOString() : "" })} style={inputStyle} />
        </Field>
        {f.apptAt && (
          <button type="button" onClick={() => { const rep = (dealUsers || []).find((u) => u.id === (f.ownerId || liveUser.id)) || liveUser; const mgr = (dealUsers || []).find((u) => u.id === rep.managerId); downloadAppointmentICS({ ...f, id: deal?.id || "new" }, rep, mgr); }}
            className="tap" style={{ display: "flex", alignItems: "center", gap: 8, width: "100%", justifyContent: "center", background: CYAN + "18", color: "#0B6A8C", border: `1px solid ${CYAN}55`, borderRadius: 9, padding: "10px 14px", fontSize: 13.5, fontWeight: 600, cursor: "pointer", marginBottom: 14 }}>
            <CalendarCheck size={15} /> Add to calendar (.ics)
          </button>
        )}
        <Field label="Stage">
          <div style={{ position: "relative" }}>
            <select value={f.stage} onChange={(e) => setF({ ...f, stage: e.target.value })} style={{ ...inputStyle, appearance: "none", cursor: "pointer" }}>
              {STAGES.map((s) => <option key={s.id} value={s.id}>{s.label}</option>)}
            </select>
            <ChevronDown size={15} style={{ position: "absolute", right: 12, top: 12, pointerEvents: "none", opacity: 0.5 }} />
          </div>
        </Field>
        {f.stage === "appointment" && !deal?.apptCredited && (
          <div style={{ background: CYAN + "14", color: "#0B6A8C", borderRadius: 8, padding: "9px 12px", fontSize: 12.5, marginBottom: 12 }}>
            Setting this to "Appointment Set" will credit 1 appointment to the owner's activity stats.
          </div>
        )}
        {f.stage === "lost" && (
          <Field label="Reason for loss">
            <div style={{ position: "relative" }}>
              <select value={f.lostReason || ""} onChange={(e) => { setF({ ...f, lostReason: e.target.value }); setErr(""); }}
                style={{ ...inputStyle, appearance: "none", cursor: "pointer" }}>
                <option value="" disabled>Choose a reason…</option>
                {LOST_REASONS.map((r) => <option key={r} value={r}>{r}</option>)}
              </select>
              <ChevronDown size={15} style={{ position: "absolute", right: 12, top: 12, pointerEvents: "none", opacity: 0.5 }} />
            </div>
          </Field>
        )}
        <Field label="Notes (optional)"><textarea rows={3} value={f.notes} onChange={(e) => setF({ ...f, notes: e.target.value })} style={{ ...inputStyle, resize: "vertical" }} placeholder="Next steps, context, objections…" /></Field>
        {deal && history && history.length > 0 && (
          <div style={{ marginBottom: 16 }}>
            <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: 0.4, textTransform: "uppercase", color: "#8494A6", marginBottom: 8 }}>Stage history</div>
            <div style={{ display: "grid", gap: 8 }}>
              {history.map((h) => (
                <div key={h.id} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5 }}>
                  <span style={{ width: 6, height: 6, borderRadius: "50%", background: STAGE[h.toStage]?.color || "#8494A6", flexShrink: 0 }} />
                  <span style={{ fontWeight: 600 }}>{h.fromStage ? `${STAGE[h.fromStage]?.label || h.fromStage} → ` : "Created as "}{STAGE[h.toStage]?.label || h.toStage}</span>
                  <span style={{ opacity: 0.5, marginLeft: "auto" }}>{new Date(h.changedAt).toLocaleDateString("en-US", { month: "short", day: "numeric" })}</span>
                </div>
              ))}
            </div>
          </div>
        )}
        {deal?.id && (
          <div style={{ marginBottom: 16 }}>
            <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: 0.4, textTransform: "uppercase", color: "#8494A6", marginBottom: 8 }}>Notes &amp; history</div>
            <EntityNotes entityType="deal" entityId={deal.id} users={dealUsers} effectiveUser={liveUser} compact />
          </div>
        )}
        {err && <div style={{ color: "#B4453F", fontSize: 13, marginBottom: 10 }}>{err}</div>}
        <div style={{ display: "flex", gap: 10 }}>
          <button onClick={submit} className="tap"
            style={{ flex: 1, background: `linear-gradient(90deg, ${BTN_A}, ${BTN_B})`, color: "#fff", border: "none", borderRadius: 10, padding: 13, fontSize: 15, fontWeight: 600, cursor: "pointer" }}>
            {deal?.id ? "Save changes" : "Create deal"}
          </button>
          {deal?.id && (
            <button onClick={() => { onDelete(deal.id); onClose(); }} className="tap"
              style={{ background: "transparent", color: "#B4453F", border: `1px solid ${LINE_C}`, borderRadius: 10, padding: "13px 16px", fontSize: 14, fontWeight: 500, cursor: "pointer" }}>
              <Trash2 size={16} />
            </button>
          )}
        </div>
      </div>
    </div>
  );
}


function AdminPortal({ users, saveUsers, entries, saveEntries, deals, saveDeals }) {
  const [modal, setModal] = useState(null);
  const [busyMsg, setBusyMsg] = useState("");
  const salesReps = users.filter((u) => u.role === "sales");

  const removeUser = (id) => {
    if (!confirm("Remove this user? Their account, logged activity, and deals will also be deleted.")) return;
    // Deleting the auth user cascades to profile/entries/deals via the DB foreign keys.
    saveUsers(() => api.deleteProfile(id));
  };

  const saveUser = async (data) => {
    try {
      setBusyMsg(data.id ? "Saving…" : "Creating account…");
      if (data.id) {
        await saveUsers(() => api.updateProfile(data.id, {
          name: data.name, email: data.email, role: data.role,
        }));
      } else {
        // Creates the auth account (via Edge Function) + profile in one step.
        await saveUsers(() => api.adminCreateUser({
          name: data.name, email: data.email, password: data.password,
          role: data.role,
        }));
      }
      setModal(null);
    } catch (e) {
      alert("Could not save user: " + (e?.message || e));
    } finally {
      setBusyMsg("");
    }
  };

  const byRole = { admin: [], management: [], sales: [], bdr: [] };
  users.forEach((u) => byRole[u.role] && byRole[u.role].push(u));
  const mgrName = (id) => { const m = users.find((u) => u.id === id); return m ? m.name : "—"; };

  return (
    <>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12, marginBottom: 20 }}>
        <div>
          <h1 style={{ fontFamily: "'Fraunces', serif", fontSize: 29, fontWeight: 600, margin: 0 }}>Admin Portal</h1>
          <p style={{ margin: "4px 0 0", opacity: 0.55, fontSize: 14 }}>{users.length} users · manage people, roles, and reporting lines</p>
        </div>
        <button onClick={() => setModal({ user: null })} className="tap"
          style={{ display: "flex", alignItems: "center", gap: 7, background: INK, color: PAPER, border: "none", borderRadius: 10, padding: "11px 18px", fontSize: 14, fontWeight: 600, cursor: "pointer" }}>
          <UserPlus size={16} /> Add user
        </button>
      </div>

      {["admin", "management", "sales", "bdr"].map((r) => (
        <div key={r} style={{ marginBottom: 18 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
            <span style={{ width: 9, height: 9, borderRadius: "50%", background: ROLES[r].color }} />
            <h3 style={{ fontSize: 14, fontWeight: 600, margin: 0 }}>{ROLES[r].label}</h3>
            <span style={{ fontSize: 12.5, opacity: 0.45 }}>({byRole[r].length})</span>
          </div>
          {byRole[r].length === 0 ? (
            <div style={{ fontSize: 13, opacity: 0.4, padding: "6px 2px" }}>No one in this role yet.</div>
          ) : (
            <div style={{ display: "grid", gap: 8 }}>
              {byRole[r].map((u) => (
                <div key={u.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", background: CARD, border: `1px solid ${LINE_C}`, borderRadius: 11, padding: "12px 16px" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 12, minWidth: 0 }}>
                    <span style={{ width: 36, height: 36, borderRadius: "50%", background: ROLES[r].color + "1e", color: ROLES[r].color, display: "grid", placeItems: "center", fontWeight: 700, fontSize: 14, flexShrink: 0 }}>
                      {u.name.split(" ").map((w) => w[0]).slice(0, 2).join("")}
                    </span>
                    <div style={{ minWidth: 0 }}>
                      <div style={{ fontWeight: 500, fontSize: 14 }}>{u.name}</div>
                      <div style={{ fontSize: 12.5, opacity: 0.5 }}>
                        {u.email}
                      </div>
                    </div>
                  </div>
                  <div style={{ display: "flex", gap: 4 }}>
                    <button onClick={() => setModal({ user: u })} className="tap" style={iconBtn}><Pencil size={15} /></button>
                    <button onClick={() => removeUser(u.id)} className="tap" style={iconBtn}><Trash2 size={15} /></button>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      ))}

      {modal && <UserModal user={modal.user} salesReps={salesReps} onSave={saveUser} onClose={() => setModal(null)} />}

      <div style={{ marginTop: 28 }}>
        <CarrierRepManager saveEntries={saveEntries} />
      </div>
    </>
  );
}

function CarrierRepManager({ saveEntries }) {
  const [reps, setReps] = useState(null);
  const [selected, setSelected] = useState(() => new Set());
  const [renaming, setRenaming] = useState(null); // { name, value }
  const [busy, setBusy] = useState(false);

  const load = async () => { try { setReps(await api.listCarrierReps()); } catch { setReps([]); } };
  useEffect(() => { load(); }, []);

  const toggle = (name) => setSelected((p) => { const n = new Set(p); n.has(name) ? n.delete(name) : n.add(name); return n; });

  const doRename = async () => {
    if (!renaming.value.trim() || renaming.value.trim() === renaming.name) { setRenaming(null); return; }
    setBusy(true);
    try { await saveEntries(() => api.renameCarrierRep(renaming.name, renaming.value)); await load(); setRenaming(null); }
    finally { setBusy(false); }
  };

  const doMerge = async () => {
    const names = [...selected];
    if (names.length < 2) return;
    const into = prompt(`Merge these ${names.length} names into one. Type the name to keep exactly:\n\n${names.join("\n")}`, names[0]);
    if (!into || !into.trim()) return;
    const from = names.filter((n) => n !== into.trim());
    setBusy(true);
    try { await saveEntries(() => api.mergeCarrierReps(from, into.trim())); setSelected(new Set()); await load(); }
    finally { setBusy(false); }
  };

  const doDelete = async (name) => {
    if (!confirm(`Remove "${name}" from all activities? This clears the carrier rep on those records (the activities themselves stay).`)) return;
    setBusy(true);
    try { await saveEntries(() => api.deleteCarrierRep(name)); await load(); }
    finally { setBusy(false); }
  };

  return (
    <Panel title="Carrier rep roster" icon={UserCheck} action={selected.size >= 2 ? (
      <button onClick={doMerge} disabled={busy} className="tap" style={{ background: INK, color: PAPER, border: "none", borderRadius: 8, padding: "8px 14px", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
        Merge {selected.size} selected
      </button>
    ) : null}>
      <p style={{ fontSize: 12.5, opacity: 0.6, margin: "0 0 14px" }}>Names typed on activities. Rename to fix typos, select 2+ and merge to combine duplicates, or delete to clear one everywhere.</p>
      {reps === null ? <div style={{ opacity: 0.4, fontSize: 13, padding: 8 }}>Loading…</div>
        : reps.length === 0 ? <Empty msg="No carrier reps recorded yet." /> : (
        <div style={{ display: "grid", gap: 6 }}>
          {reps.map((r) => (
            <div key={r.name} style={{ display: "flex", alignItems: "center", gap: 10, background: "#fff", border: `1px solid ${selected.has(r.name) ? EMAIL : LINE_C}`, borderRadius: 9, padding: "9px 12px" }}>
              <input type="checkbox" checked={selected.has(r.name)} onChange={() => toggle(r.name)} style={{ cursor: "pointer" }} />
              {renaming?.name === r.name ? (
                <>
                  <input autoFocus value={renaming.value} onChange={(e) => setRenaming({ ...renaming, value: e.target.value })}
                    onKeyDown={(e) => { if (e.key === "Enter") doRename(); if (e.key === "Escape") setRenaming(null); }}
                    style={{ ...inputStyle, marginBottom: 0, flex: 1 }} />
                  <button onClick={doRename} disabled={busy} className="tap" style={{ background: INK, color: PAPER, border: "none", borderRadius: 7, padding: "6px 12px", fontSize: 12.5, fontWeight: 600, cursor: "pointer" }}>Save</button>
                  <button onClick={() => setRenaming(null)} className="tap" style={{ background: "transparent", border: `1px solid ${LINE_C}`, borderRadius: 7, padding: "6px 10px", fontSize: 12.5, cursor: "pointer" }}>Cancel</button>
                </>
              ) : (
                <>
                  <span style={{ flex: 1, fontSize: 14, fontWeight: 500 }}>{r.name}</span>
                  <span style={{ fontSize: 12, opacity: 0.5 }}>{r.count} {r.count === 1 ? "activity" : "activities"}</span>
                  <button onClick={() => setRenaming({ name: r.name, value: r.name })} className="tap" style={iconBtn} title="Rename"><Pencil size={14} /></button>
                  <button onClick={() => doDelete(r.name)} className="tap" style={iconBtn} title="Delete"><Trash2 size={14} /></button>
                </>
              )}
            </div>
          ))}
        </div>
      )}
    </Panel>
  );
}

function UserModal({ user, salesReps, onSave, onClose }) {
  const [f, setF] = useState(user || { name: "", email: "", password: "", role: "bdr" });
  const [err, setErr] = useState("");
  // Password-reset state (only for editing an existing non-admin user).
  const [showReset, setShowReset] = useState(false);
  const [newPw, setNewPw] = useState("");
  const [pwMsg, setPwMsg] = useState("");
  const [pwBusy, setPwBusy] = useState(false);
  const [pwVisible, setPwVisible] = useState(true);

  const canReset = user && user.role !== "admin";

  const genPassword = () => {
    // Readable-ish random temp password.
    const chars = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789";
    let out = "";
    const arr = new Uint32Array(12);
    crypto.getRandomValues(arr);
    for (let i = 0; i < 12; i++) out += chars[arr[i] % chars.length];
    setNewPw(out); setPwVisible(true); setPwMsg("");
  };

  const doReset = async () => {
    if (newPw.trim().length < 8) { setPwMsg("Password must be at least 8 characters."); return; }
    setPwBusy(true); setPwMsg("");
    try {
      await api.adminSetPassword(user.id, newPw.trim());
      setPwMsg(`Password updated. Share it with ${user.name.split(" ")[0]}: ${newPw.trim()}`);
    } catch (e) {
      setPwMsg("Couldn't update: " + (e?.message || e));
    } finally { setPwBusy(false); }
  };

  const submit = () => {
    if (!f.name.trim() || !f.email.trim()) { setErr("Name and email are required."); return; }
    if (!user && !f.password.trim()) { setErr("A password is required for a new user."); return; }
    const data = { ...f, name: f.name.trim(), email: f.email.trim() };
    onSave(user ? { ...data, id: user.id } : data);
  };

  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(18,33,30,.5)", display: "grid", placeItems: "center", padding: 20, zIndex: 50 }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: PAPER, borderRadius: 16, padding: 24, width: "100%", maxWidth: 440, maxHeight: "90vh", overflowY: "auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18 }}>
          <h3 style={{ fontFamily: "'Fraunces', serif", fontSize: 20, fontWeight: 600, margin: 0 }}>{user ? "Edit user" : "Add user"}</h3>
          <button onClick={onClose} className="tap" style={{ background: "transparent", border: "none", cursor: "pointer" }}><X size={20} /></button>
        </div>
        <Field label="Full name"><input value={f.name} onChange={(e) => setF({ ...f, name: e.target.value })} style={inputStyle} placeholder="Jane Smith" /></Field>
        <Field label="Email"><input value={f.email} onChange={(e) => setF({ ...f, email: e.target.value })} style={inputStyle} placeholder="jane@tellemica.com" disabled={!!user} /></Field>
        {!user && <Field label="Password"><input value={f.password} onChange={(e) => setF({ ...f, password: e.target.value })} style={inputStyle} placeholder="Set a temporary password" /></Field>}
        <Field label="Role">
          <div style={{ position: "relative" }}>
            <select value={f.role} onChange={(e) => setF({ ...f, role: e.target.value })} style={{ ...inputStyle, appearance: "none", cursor: "pointer" }}>
              <option value="bdr">BDR</option>
              <option value="sales">Sales Rep</option>
              <option value="management">Management</option>
              <option value="admin">Admin</option>
            </select>
            <ChevronDown size={15} style={{ position: "absolute", right: 12, top: 12, pointerEvents: "none", opacity: 0.5 }} />
          </div>
        </Field>
        {f.role === "bdr" && (
          <p style={{ fontSize: 12.5, opacity: 0.55, margin: "-2px 0 14px", lineHeight: 1.5 }}>
            BDRs choose which Sales Rep they're working for each time they log activity or add a deal — no fixed assignment needed.
          </p>
        )}

        {/* Password reset — only for existing non-admin users */}
        {canReset && (
          <div style={{ background: "#fff", border: `1px solid ${LINE_C}`, borderRadius: 10, padding: 14, marginBottom: 14 }}>
            {!showReset ? (
              <button onClick={() => { setShowReset(true); setPwMsg(""); setNewPw(""); }} className="tap"
                style={{ background: "transparent", border: "none", color: EMAIL, fontSize: 13.5, fontWeight: 600, cursor: "pointer", padding: 0, display: "flex", alignItems: "center", gap: 7 }}>
                <Shield size={14} /> Reset this user's password
              </button>
            ) : (
              <>
                <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 8, display: "flex", alignItems: "center", gap: 7 }}><Shield size={14} color={INK} /> Set a new password for {user.name.split(" ")[0]}</div>
                <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
                  <div style={{ position: "relative", flex: 1 }}>
                    <input type={pwVisible ? "text" : "password"} value={newPw} onChange={(e) => { setNewPw(e.target.value); setPwMsg(""); }}
                      placeholder="Type or generate a password" style={{ ...inputStyle, marginBottom: 0, paddingRight: 38 }} />
                    <button onClick={() => setPwVisible((v) => !v)} className="tap" style={{ position: "absolute", right: 10, top: 10, background: "none", border: "none", cursor: "pointer", opacity: 0.5 }}>
                      {pwVisible ? <EyeOff size={16} /> : <Eye size={16} />}
                    </button>
                  </div>
                  <button onClick={genPassword} className="tap" style={{ background: "#F1F5F9", border: `1px solid ${LINE_C}`, borderRadius: 8, padding: "0 14px", fontSize: 13, fontWeight: 600, cursor: "pointer", whiteSpace: "nowrap" }}>Generate</button>
                </div>
                <div style={{ display: "flex", gap: 8 }}>
                  <button onClick={doReset} disabled={pwBusy || newPw.trim().length < 8} className="tap"
                    style={{ background: newPw.trim().length >= 8 && !pwBusy ? INK : LINE_C, color: newPw.trim().length >= 8 && !pwBusy ? PAPER : "#8494A6", border: "none", borderRadius: 8, padding: "9px 16px", fontSize: 13.5, fontWeight: 600, cursor: newPw.trim().length >= 8 && !pwBusy ? "pointer" : "default" }}>
                    {pwBusy ? "Updating…" : "Update password"}
                  </button>
                  <button onClick={() => { setShowReset(false); setNewPw(""); setPwMsg(""); }} className="tap" style={{ background: "transparent", border: `1px solid ${LINE_C}`, borderRadius: 8, padding: "9px 12px", fontSize: 13, cursor: "pointer" }}>Cancel</button>
                </div>
                {pwMsg && <div style={{ marginTop: 10, background: pwMsg.startsWith("Couldn't") || pwMsg.startsWith("Password must") ? "#FBECEB" : CALL + "18", color: pwMsg.startsWith("Couldn't") || pwMsg.startsWith("Password must") ? "#B4453F" : CALL, borderRadius: 8, padding: "10px 12px", fontSize: 13, fontWeight: 500, wordBreak: "break-all" }}>{pwMsg}</div>}
              </>
            )}
          </div>
        )}

        {err && <div style={{ color: "#B4453F", fontSize: 13, marginBottom: 10 }}>{err}</div>}
        <button onClick={submit} className="tap"
          style={{ width: "100%", marginTop: 6, background: `linear-gradient(90deg, ${BTN_A}, ${BTN_B})`, color: "#fff", border: "none", borderRadius: 10, padding: 13, fontSize: 15, fontWeight: 600, cursor: "pointer" }}>
          {user ? "Save changes" : "Create user"}
        </button>
      </div>
    </div>
  );
}

function Panel({ title, children, style, icon: Icon, action }) {
  return (
    <div style={{ background: CARD, border: `1px solid ${LINE_C}`, borderRadius: 14, padding: 18, ...style }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 7, marginBottom: 14 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 7 }}>
          {Icon && <Icon size={16} color={INK} />}
          <span style={{ fontSize: 14, fontWeight: 600 }}>{title}</span>
        </div>
        {action && <div style={{ display: "flex", gap: 6, alignItems: "center" }}>{action}</div>}
      </div>
      {children}
    </div>
  );
}
function Field({ label, children }) {
  return <label style={{ display: "block", marginBottom: 14 }}><div style={lblStyle}>{label}</div>{children}</label>;
}
function Empty({ msg }) {
  return <div style={{ textAlign: "center", padding: "30px 10px", opacity: 0.4, fontSize: 14 }}>{msg}</div>;
}

const inputStyle = { width: "100%", padding: "10px 12px", border: `1px solid ${LINE_C}`, borderRadius: 9, fontSize: 14, color: INK, background: CARD, outline: "none" };
const lblStyle = { fontSize: 12.5, fontWeight: 500, opacity: 0.65, marginBottom: 5 };
const iconBtn = { background: "transparent", border: `1px solid ${LINE_C}`, borderRadius: 8, padding: 8, cursor: "pointer", color: INK, display: "grid", placeItems: "center" };
