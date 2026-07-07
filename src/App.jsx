import React, { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import {
  Phone, Mail, CalendarCheck, TrendingUp, Trophy, Plus, Target,
  ChevronDown, ChevronLeft, ChevronRight, X, Users, BarChart3, Trash2, Shield, LogOut,
  UserPlus, Pencil, Eye, EyeOff, Briefcase, DollarSign, Kanban,
  Table2, ArrowRight, Building2, Percent, CheckCircle2, Download, FileSpreadsheet
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
  const [loaded, setLoaded] = useState(false);       // finished checking session
  const [authed, setAuthed] = useState(false);       // has a valid session

  // Pull all data the signed-in user is allowed to see. RLS filters server-side.
  const refetch = async () => {
    const [us, es, ds, g, ug] = await Promise.all([
      api.listProfiles(), api.listEntries(), api.listDeals(), api.getGoals(), api.listUserGoals(),
    ]);
    setUsers(us); setEntries(es); setDeals(ds); setUserGoals(ug);
    setGoals({ calls: g.calls, emails: g.emails, appts: g.appts });
  };

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

  const nav = [["dashboard", "Dashboard", BarChart3]];
  if (canLog) nav.push(["log", "Log Activity", Plus]);
  nav.push(["activity", "Activity Log", Table2]);
  nav.push(["pipeline", "Pipeline", Briefcase]);
  if (canManageGoals) nav.push(["goals", "Goals", Target]);
  if (isAdmin) nav.push(["admin", "Admin Portal", Shield]);

  return (
    <div style={{ minHeight: "100vh", background: PAPER, fontFamily: "'Inter', system-ui, sans-serif", color: INK }}>
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
      `}</style>

      <header style={{ background: `linear-gradient(100deg, ${INK}, ${NAVY2})`, color: PAPER, padding: "16px 24px", borderBottom: `3px solid ${CYAN}` }}>
        <div style={{ maxWidth: 1120, margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
            <TellemicaLogo height={19} light wordmark />
            <span style={{ fontSize: 12, opacity: 0.7, letterSpacing: 1, textTransform: "uppercase", borderLeft: "1px solid rgba(255,255,255,.25)", paddingLeft: 14 }}>Command Center</span>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
            <nav style={{ display: "flex", gap: 6, background: "rgba(255,255,255,.08)", padding: 4, borderRadius: 10 }}>
              {nav.map(([id, label, Icon]) => (
                <button key={id} onClick={() => setView(id)} className="tap"
                  style={{ display: "flex", alignItems: "center", gap: 6, border: "none", borderRadius: 7, padding: "8px 13px", fontSize: 13.5, fontWeight: 500,
                    background: view === id ? PAPER : "transparent", color: view === id ? INK : PAPER }}>
                  <Icon size={15} /> {label}
                </button>
              ))}
            </nav>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              {realIsAdmin && (
                <div style={{ position: "relative" }}>
                  <select value={viewAsId} onChange={(e) => { setViewAsId(e.target.value); setView("dashboard"); }}
                    title="View the app as another role/person"
                    style={{ appearance: "none", background: impersonating ? CYAN : "rgba(255,255,255,.12)", color: impersonating ? INK : PAPER, border: impersonating ? "none" : "1px solid rgba(255,255,255,.25)", borderRadius: 8, padding: "8px 30px 8px 12px", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
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
        {view === "dashboard" && (
          <Dashboard entries={visibleEntries} deals={visibleDeals} users={users} goals={goals} saveGoals={saveGoals}
            userGoals={userGoals} liveUser={effectiveUser} visibleUserIds={visibleUserIds} setView={setView} canLog={canLog} />
        )}
        {view === "log" && canLog && (
          <LogView liveUser={effectiveUser} entries={entries} saveEntries={saveEntries} users={users} allEntries={visibleEntries} visibleUserIds={visibleUserIds} />
        )}
        {view === "pipeline" && (
          <Pipeline deals={visibleDeals} allDeals={deals} saveDeals={saveDeals}
            liveUser={effectiveUser} users={users} visibleUserIds={visibleUserIds}
            entries={entries} saveEntries={saveEntries} />
        )}
        {view === "activity" && (
          <>
            <div style={{ marginBottom: 20 }}>
              <h1 style={{ fontFamily: "'Fraunces', serif", fontSize: 29, fontWeight: 600, margin: 0 }}>Activity Log</h1>
              <p style={{ margin: "4px 0 0", opacity: 0.55, fontSize: 14 }}>
                Every activity record in your scope. Click a column to sort, or export to Excel.
              </p>
            </div>
            <ActivityTable entries={visibleEntries} users={users} liveUser={effectiveUser} />
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

// Spreadsheet-style table of activity records, with Excel export.
function ActivityTable({ entries, users, liveUser, compact }) {
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
      Contact: e.contact || "",
      Phone: e.phone || "",
      Email: e.email || "",
      Calls: e.calls || 0,
      Emails: e.emails || 0,
      Appointments: e.appts || 0,
      "Logged by": nameOf(e.userId),
      "Working for": repOf(e),
      Notes: e.notes || "",
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    ws["!cols"] = [
      { wch: 11 }, { wch: 24 }, { wch: 14 }, { wch: 20 }, { wch: 16 }, { wch: 24 },
      { wch: 7 }, { wch: 8 }, { wch: 13 }, { wch: 18 }, { wch: 18 }, { wch: 40 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Activity");
    const stamp = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `tellemica-activity-${stamp}.xlsx`);
  };

  const cols = [
    ["date", "Date"], ["company", "Company"], ["ban", "BAN"], ["contact", "Contact"],
    ["phone", "Phone"], ["email", "Email"], ["calls", "Calls"], ["emails", "Emails"],
    ["appts", "Appts"], ["logged", "Logged by"], ["rep", "Working for"], ["notes", "Notes"],
  ];
  const sortable = new Set(["date", "company", "calls", "emails", "appts", "logged", "rep"]);

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
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, minWidth: 900 }}>
          <thead>
            <tr style={{ background: "#F1F5F9", position: "sticky", top: 0, zIndex: 1 }}>
              {cols.map(([k, label]) => (
                <th key={k} onClick={() => sortable.has(k) && setSort(k)}
                  style={{ textAlign: (["calls", "emails", "appts"].includes(k)) ? "center" : "left", padding: "10px 12px", fontWeight: 700, fontSize: 11.5, letterSpacing: 0.4, textTransform: "uppercase", color: "#5A6B7B", whiteSpace: "nowrap", cursor: sortable.has(k) ? "pointer" : "default", borderBottom: `1px solid ${LINE_C}` }}>
                  {label}{sortKey === k ? (sortDir === "asc" ? " ▲" : " ▼") : ""}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 ? (
              <tr><td colSpan={cols.length} style={{ padding: 28, textAlign: "center", opacity: 0.5 }}>No activity records yet.</td></tr>
            ) : rows.map((e) => (
              <tr key={e.id} style={{ borderBottom: `1px solid ${LINE_C}` }}>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{e.date}</td>
                <td style={{ padding: "9px 12px", fontWeight: 600, whiteSpace: "nowrap" }}>{e.company || "—"}</td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{e.ban || "—"}</td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{e.contact || "—"}</td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{e.phone || "—"}</td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{e.email || "—"}</td>
                <td style={{ padding: "9px 12px", textAlign: "center" }}>{e.calls || 0}</td>
                <td style={{ padding: "9px 12px", textAlign: "center" }}>{e.emails || 0}</td>
                <td style={{ padding: "9px 12px", textAlign: "center" }}>{e.appts || 0}</td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{nameOf(e.userId)}</td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{repOf(e)}</td>
                <td style={{ padding: "9px 12px", maxWidth: 280, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={e.notes || ""}>{e.notes || "—"}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function Dashboard({ entries, deals, users, goals, saveGoals, userGoals, liveUser, visibleUserIds, setView, canLog }) {
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

      <GoalBars totals={totals} targets={targets} isTeam={repFilter === "all" && role !== "bdr"} />

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

      <div style={{ marginTop: 16 }}>
        <ActivityTable entries={scoped} users={users} liveUser={liveUser} compact />
      </div>
    </>
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

function GoalBars({ totals, targets, isTeam }) {
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
      </div>
      <div style={{ display: "grid", gap: 12 }}>
        {rows.map(([label, val, target, color]) => {
          const pct = Math.min(100, target ? (val / target) * 100 : 0);
          return (
            <div key={label}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, marginBottom: 5 }}>
                <span style={{ fontWeight: 500 }}>{label}</span>
                <span style={{ opacity: 0.6 }}>{val.toLocaleString()} / {target.toLocaleString()}</span>
              </div>
              <div style={{ height: 8, background: LINE_C, borderRadius: 99, overflow: "hidden" }}>
                <div style={{ width: `${pct}%`, height: "100%", background: color, borderRadius: 99, transition: "width .4s ease" }} />
              </div>
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

  const TEMPLATE_COLS = ["Date", "Company", "BAN", "Contact", "Phone", "Email", "Calls", "Emails", "Appointments", "Prospecting For", "BDR/Sale Rep [You]", "Notes"];

  const downloadTemplate = async () => {
    const N = 2000; // rows the dropdowns cover
    const example = {
      Date: TODAY_US(), Company: "Acme Corp", BAN: "123456789", Contact: "Jane Smith",
      Phone: "(610) 555-0100", Email: "jane@acme.com", Calls: 0, Emails: 0, Appointments: 0,
      "Prospecting For": "Self-generated", "BDR/Sale Rep [You]": liveUser.name, Notes: "Intro call, follow up next week",
    };
    const blank = Object.fromEntries(TEMPLATE_COLS.map((c) => [c, ""]));
    const ws = XLSX.utils.json_to_sheet([example, blank], { header: TEMPLATE_COLS });
    ws["!cols"] = [{ wch: 12 }, { wch: 22 }, { wch: 14 }, { wch: 18 }, { wch: 16 }, { wch: 22 }, { wch: 7 }, { wch: 8 }, { wch: 13 }, { wch: 18 }, { wch: 18 }, { wch: 34 }];
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
        { sqref: `G2:G${N + 1}`, f: '"0,1"' },      // Calls
        { sqref: `H2:H${N + 1}`, f: '"0,1"' },      // Emails
        { sqref: `I2:I${N + 1}`, f: '"0,1"' },      // Appointments
        { sqref: `J2:J${N + 1}`, f: `Lists!$A$2:$A$${wfEnd}` },  // Prospecting For
        { sqref: `K2:K${N + 1}`, f: `Lists!$B$2:$B$${roEnd}` }, // BDR/Sale Rep [You]
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
        ban: String(r["BAN"] || "").trim(), contact: String(r["Contact"] || "").trim(),
        phone: String(r["Phone"] || "").trim(), email: String(r["Email"] || "").trim(),
        calls, emails, appts, notes: String(r["Notes"] || "").trim(),
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
        userId: r.userId, date: r.date, company: r.company, ban: r.ban, contact: r.contact,
        phone: r.phone, email: r.email, calls: r.calls, emails: r.emails, appts: r.appts,
        notes: r.notes, taggedRepId: r.taggedRepId,
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
                  {["", "Date", "Company", "Calls", "Emails", "Appts", "Working for", "Owner"].map((h) => (
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
  const [form, setForm] = useState({ date: TODAY(), company: "", ban: "", contact: "", phone: "", email: "", calls: "", emails: "", appts: "", notes: "", workingFor: isBDR ? "" : "self" });
  const [toast, setToast] = useState(false);
  const [err, setErr] = useState("");
  const [showSuggest, setShowSuggest] = useState(false);

  // Unique company names from everything the user can see, for autocomplete.
  const companyIndex = useMemo(() => {
    const map = new Map();
    (allEntries || entries).forEach((e) => {
      const c = (e.company || "").trim();
      if (c && !map.has(c.toLowerCase())) map.set(c.toLowerCase(), { company: c, ban: e.ban || "", contact: e.contact || "", phone: e.phone || "", email: e.email || "" });
    });
    return [...map.values()].sort((a, b) => a.company.localeCompare(b.company));
  }, [allEntries, entries]);

  const suggestions = form.company.trim().length >= 1
    ? companyIndex.filter((c) => c.company.toLowerCase().includes(form.company.trim().toLowerCase())).slice(0, 6)
    : [];

  const pickCompany = (c) => {
    // Reuse known details for this company, but let the user override any field.
    setForm((f) => ({ ...f, company: c.company, ban: f.ban || c.ban, contact: f.contact || c.contact, phone: f.phone || c.phone, email: f.email || c.email }));
    setShowSuggest(false);
  };

  const submit = async () => {
    if (!form.company.trim()) { setErr("Company name is required."); return; }
    if (isBDR && !form.workingFor) { setErr("Please choose who you're working for."); return; }
    setErr("");
    const taggedRepId = form.workingFor && form.workingFor !== "self" ? form.workingFor : null;
    await saveEntries(() => api.addEntry({
      userId: liveUser.id, date: form.date,
      company: form.company.trim(), ban: form.ban.trim(), contact: form.contact.trim(),
      phone: form.phone.trim(), email: form.email.trim(),
      calls: +form.calls || 0, emails: +form.emails || 0, appts: +form.appts || 0, notes: form.notes.trim(),
      taggedRepId,
    }));
    setForm({ ...form, company: "", ban: "", contact: "", phone: "", email: "", calls: "", emails: "", appts: "", notes: "" });
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
        {isBDR && (
          <Field label="Working for">
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
          <Field label="Contact"><input value={form.contact} onChange={(e) => setForm({ ...form, contact: e.target.value })} style={inputStyle} placeholder="Name / title" /></Field>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          <Field label="Phone"><input value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} style={inputStyle} placeholder="(610) 555-0100" /></Field>
          <Field label="Email"><input value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })} style={inputStyle} placeholder="name@company.com" /></Field>
        </div>
        <Field label="Date"><input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} style={inputStyle} /></Field>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 }}>
          <Field label="Calls"><input type="number" min="0" placeholder="0" value={form.calls} onChange={(e) => setForm({ ...form, calls: e.target.value })} style={inputStyle} /></Field>
          <Field label="Emails"><input type="number" min="0" placeholder="0" value={form.emails} onChange={(e) => setForm({ ...form, emails: e.target.value })} style={inputStyle} /></Field>
          <Field label="Appts set"><input type="number" min="0" placeholder="0" value={form.appts} onChange={(e) => setForm({ ...form, appts: e.target.value })} style={inputStyle} /></Field>
        </div>
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
  const canEdit = role === "bdr" || role === "sales"; // only reps own/edit deals
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
      const merged = { ...prev, ...data };
      // Save the deal first so we have an id, then handle credit + credited flag.
      const saved = await api.upsertDeal(merged);
      const credited = await maybeCreditAppt(saved, prev?.stage);
      if (credited && !saved.apptCredited) await api.updateDeal(saved.id, { apptCredited: true });
    } else {
      const created = await api.upsertDeal({ ...data, ownerId: liveUser.id });
      const credited = await maybeCreditAppt(created, null);
      if (credited) await api.updateDeal(created.id, { apptCredited: true });
    }
  }).then(() => setModal(null));

  const moveDeal = (dealId, newStage) => {
    const prev = allDeals.find((d) => d.id === dealId);
    if (!prev || prev.stage === newStage) return;
    return saveDeals(async () => {
      const saved = await api.updateDeal(dealId, { ...prev, stage: newStage });
      const credited = await maybeCreditAppt(saved, prev.stage);
      if (credited && !saved.apptCredited) await api.updateDeal(dealId, { apptCredited: true });
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
          liveUser={liveUser} salesReps={users.filter((u) => u.role === "sales")} />
      )}
    </>
  );
}

function DealModal({ deal, onSave, onDelete, onClose, liveUser, salesReps }) {
  const isBDR = liveUser && liveUser.role === "bdr";
  const [f, setF] = useState(deal || { company: "", contact: "", value: "", stage: "new", closeDate: "", notes: "", taggedRepId: isBDR ? "" : "self" });
  const [err, setErr] = useState("");
  const submit = () => {
    if (!f.company.trim()) { setErr("Company / prospect name is required."); return; }
    if (isBDR && !f.taggedRepId) { setErr("Please choose who you're working for."); return; }
    const taggedRepId = f.taggedRepId && f.taggedRepId !== "self" ? f.taggedRepId : null;
    onSave({ ...f, company: f.company.trim(), contact: f.contact.trim(), value: +f.value || 0, taggedRepId, ...(deal ? { id: deal.id } : {}) });
  };
  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(11,42,74,.5)", display: "grid", placeItems: "center", padding: 20, zIndex: 50 }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: PAPER, borderRadius: 16, padding: 24, width: "100%", maxWidth: 460, maxHeight: "90vh", overflowY: "auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18 }}>
          <h3 style={{ fontFamily: "'Fraunces', serif", fontSize: 20, fontWeight: 600, margin: 0 }}>{deal ? "Edit deal" : "New deal"}</h3>
          <button onClick={onClose} className="tap" style={{ background: "transparent", border: "none", cursor: "pointer" }}><X size={20} /></button>
        </div>
        <Field label="Company / prospect"><input value={f.company} onChange={(e) => setF({ ...f, company: e.target.value })} style={inputStyle} placeholder="Acme Corp" autoFocus /></Field>
        {isBDR && (
          <Field label="Working for">
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
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          <Field label="Deal value ($)"><input type="number" min="0" value={f.value} onChange={(e) => setF({ ...f, value: e.target.value })} style={inputStyle} placeholder="0" /></Field>
          <Field label="Expected close"><input type="date" value={f.closeDate} onChange={(e) => setF({ ...f, closeDate: e.target.value })} style={inputStyle} /></Field>
        </div>
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
        <Field label="Notes (optional)"><textarea rows={3} value={f.notes} onChange={(e) => setF({ ...f, notes: e.target.value })} style={{ ...inputStyle, resize: "vertical" }} placeholder="Next steps, context, objections…" /></Field>
        {err && <div style={{ color: "#B4453F", fontSize: 13, marginBottom: 10 }}>{err}</div>}
        <div style={{ display: "flex", gap: 10 }}>
          <button onClick={submit} className="tap"
            style={{ flex: 1, background: `linear-gradient(90deg, ${BTN_A}, ${BTN_B})`, color: "#fff", border: "none", borderRadius: 10, padding: 13, fontSize: 15, fontWeight: 600, cursor: "pointer" }}>
            {deal ? "Save changes" : "Create deal"}
          </button>
          {deal && (
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
    </>
  );
}

function UserModal({ user, salesReps, onSave, onClose }) {
  const [f, setF] = useState(user || { name: "", email: "", password: "", role: "bdr" });
  const [err, setErr] = useState("");

  const submit = () => {
    if (!f.name.trim() || !f.email.trim()) { setErr("Name and email are required."); return; }
    if (!user && !f.password.trim()) { setErr("A password is required for a new user."); return; }
    const data = { ...f, name: f.name.trim(), email: f.email.trim() };
    onSave(user ? { ...data, id: user.id } : data);
  };

  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(18,33,30,.5)", display: "grid", placeItems: "center", padding: 20, zIndex: 50 }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: PAPER, borderRadius: 16, padding: 24, width: "100%", maxWidth: 440 }}>
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
        {err && <div style={{ color: "#B4453F", fontSize: 13, marginBottom: 10 }}>{err}</div>}
        <button onClick={submit} className="tap"
          style={{ width: "100%", marginTop: 6, background: `linear-gradient(90deg, ${BTN_A}, ${BTN_B})`, color: "#fff", border: "none", borderRadius: 10, padding: 13, fontSize: 15, fontWeight: 600, cursor: "pointer" }}>
          {user ? "Save changes" : "Create user"}
        </button>
      </div>
    </div>
  );
}

function Panel({ title, children, style, icon: Icon }) {
  return (
    <div style={{ background: CARD, border: `1px solid ${LINE_C}`, borderRadius: 14, padding: 18, ...style }}>
      <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 14 }}>
        {Icon && <Icon size={16} color={INK} />}
        <span style={{ fontSize: 14, fontWeight: 600 }}>{title}</span>
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
