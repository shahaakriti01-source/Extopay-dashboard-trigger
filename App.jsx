import { useState, useMemo, useEffect, useCallback, useRef } from "react";
import { MsalProvider, useMsal, useIsAuthenticated } from "@azure/msal-react";
import { LineChart, Line, BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, Legend } from "recharts";
import { msalInstance, loginRequest } from "./authConfig";
import { fetchLivePilotData } from "./graphService";
import { EXTO_BACKEND_NAME, ALWAYS_COUNT_TYPES, PAIR_TYPES, auditTransactions } from "./transactionLogic";
import { dedupeOnboardingEvents, getAllMerchantNames, getConsumersForMerchant, SPECIAL_ACCOUNT_NAMES, isSpecialAccountName } from "./onboardingLogic";
import { calculateMerchantIncentive } from "./incentiveEngine";
import {
  loadIncentiveConfigs, saveIncentiveConfigs,
  loadDashboardTargets, saveDashboardTargets,
  loadPaidTriggers, savePaidTriggers,
  DEFAULT_INCENTIVE_CONFIG,
} from "./phaseConfig";

const AUTO_REFRESH_MS = 60 * 1000;
const ADMIN_PASSWORD = "exto@pilot2026"; // change this before sharing the deployed link

// ─── FALLBACK SAMPLE DATA (used until live OneDrive data loads, or if it fails) ──
const FALLBACK_ONBOARDING_RAW = [];
const FALLBACK_PAYMENTS = [];

// ─── DESIGN TOKENS (Exto Pay brand) ──────────────────────────────────────────
const C = {
  bgDark: "#0d1812", panelDark: "#142621", panelDark2: "#1a2f29", forest: "#1f3d35",
  forestLight: "#2d5a4d", gold: "#d4a531", goldLight: "#e8c468", cream: "#f4f1ea",
  textMuted: "#9bb3ab", textFaint: "#5c7068", border: "#243d36",
  red: "#e0664f", amber: "#d4a531", green: "#5fae8c", blue: "#5b9bd1",
};

// ─── SMALL UI PRIMITIVES ─────────────────────────────────────────────────────
function Badge({ color, children }) {
  return (
    <span style={{ background: color + "22", color, border: `1px solid ${color}55`, borderRadius: 6, padding: "2px 9px", fontSize: 11, fontWeight: 600, letterSpacing: 0.2 }}>{children}</span>
  );
}

function KPICard({ label, value, sub, color = C.gold, icon }) {
  return (
    <div style={{ background: C.panelDark, border: `1px solid ${C.border}`, borderRadius: 14, padding: "20px 22px", flex: 1, minWidth: 170 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
        <span style={{ fontSize: 18 }}>{icon}</span>
        <span style={{ color: C.textFaint, fontSize: 11, fontWeight: 600, textTransform: "uppercase", letterSpacing: 1 }}>{label}</span>
      </div>
      <div style={{ color, fontSize: 27, fontWeight: 700, letterSpacing: -0.5, fontFamily: "'Space Grotesk', sans-serif" }}>{value}</div>
      {sub && <div style={{ color: C.textFaint, fontSize: 11.5, marginTop: 4 }}>{sub}</div>}
    </div>
  );
}

function NumberInput({ label, value, onChange, width = 75, suffix = "" }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
      <span style={{ color: C.textFaint, fontSize: 10.5 }}>{label}</span>
      <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
        <input type="number" value={value} onChange={e => onChange(parseFloat(e.target.value) || 0)}
          style={{ width, padding: "6px 8px", background: C.bgDark, border: `1px solid ${C.border}`, borderRadius: 6, color: C.cream, fontSize: 13 }} />
        {suffix && <span style={{ color: C.textFaint, fontSize: 11 }}>{suffix}</span>}
      </div>
    </div>
  );
}

// Reusable phase + date-range filter bar, used independently on every tab
// (Dashboard, Incentives, Triggers, Merchants, Internal Transfers, Audit).
function PhaseDateFilterBar({ phase, setPhase, dateFrom, setDateFrom, dateTo, setDateTo, note }) {
  return (
    <div style={{ background: C.panelDark, border: `1px solid ${C.border}`, borderRadius: 14, padding: "14px 20px", marginBottom: 20, display: "flex", gap: 20, flexWrap: "wrap", alignItems: "center" }}>
      <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
        <span style={{ color: C.textFaint, fontSize: 11, fontWeight: 600 }}>PHASE</span>
        {["0", "1", "2", "3"].map(p => (
          <button key={p} onClick={() => setPhase(p)} style={{ background: phase === p ? C.gold : "transparent", color: phase === p ? "#1a1a1a" : C.textMuted, border: `1px solid ${phase === p ? C.gold : C.border}`, borderRadius: 7, padding: "5px 12px", fontSize: 12, cursor: "pointer" }}>Phase {p}</button>
        ))}
      </div>
      <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
        <span style={{ color: C.textFaint, fontSize: 11, fontWeight: 600 }}>DATE RANGE</span>
        <input type="date" value={dateFrom} onChange={e => setDateFrom(e.target.value)} style={{ background: C.bgDark, border: `1px solid ${C.border}`, borderRadius: 6, padding: "5px 8px", color: C.cream, fontSize: 12 }} />
        <span style={{ color: C.textFaint }}>→</span>
        <input type="date" value={dateTo} onChange={e => setDateTo(e.target.value)} style={{ background: C.bgDark, border: `1px solid ${C.border}`, borderRadius: 6, padding: "5px 8px", color: C.cream, fontSize: 12 }} />
      </div>
      {note && <span style={{ color: C.textFaint, fontSize: 11, fontStyle: "italic" }}>{note}</span>}
    </div>
  );
}

function TargetBar({ label, actual, target, suffix = "" }) {
  const pct = target > 0 ? Math.min((actual / target) * 100, 100) : 0;
  const met = actual >= target;
  return (
    <div style={{ marginBottom: 16 }}>
      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
        <span style={{ color: C.textMuted, fontSize: 12.5 }}>{label}</span>
        <span style={{ fontSize: 12.5 }}>
          <span style={{ color: met ? C.green : C.gold, fontWeight: 700 }}>{actual.toLocaleString()}</span>
          <span style={{ color: C.textFaint }}> / {target.toLocaleString()}{suffix} target</span>
        </span>
      </div>
      <div style={{ background: C.border, borderRadius: 8, height: 9, overflow: "hidden" }}>
        <div style={{ width: `${pct}%`, background: met ? C.green : C.gold, height: "100%", borderRadius: 8, transition: "width 0.6s ease" }} />
      </div>
    </div>
  );
}

// ─── AUTH: PASSWORD GATE ─────────────────────────────────────────────────────
function LoginScreen({ onLogin }) {
  const [pw, setPw] = useState("");
  const [error, setError] = useState(false);
  const [shake, setShake] = useState(false);

  const handleSubmit = () => {
    if (pw === ADMIN_PASSWORD) onLogin();
    else { setError(true); setShake(true); setTimeout(() => setShake(false), 400); }
  };

  return (
    <div style={{ minHeight: "100vh", background: `radial-gradient(circle at 30% 20%, #1a3329 0%, ${C.bgDark} 55%)`, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Inter', sans-serif" }}>
      <div style={{ background: C.panelDark, border: `1px solid ${C.border}`, borderRadius: 20, padding: "48px 40px", width: 380, animation: shake ? "shake 0.4s ease" : "none", boxShadow: "0 20px 60px rgba(0,0,0,0.4)" }}>
        <div style={{ textAlign: "center", marginBottom: 32 }}>
          <div style={{ width: 52, height: 52, background: C.gold, borderRadius: 12, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 16px", fontSize: 26, fontWeight: 900, color: "#1a1a1a", fontFamily: "'Space Grotesk', sans-serif" }}>E</div>
          <div style={{ color: C.cream, fontSize: 22, fontWeight: 700, letterSpacing: -0.5, fontFamily: "'Space Grotesk', sans-serif" }}>Exto Pay</div>
          <div style={{ color: C.textMuted, fontSize: 13, marginTop: 4 }}>Project PulaConnect · Botswana Pilot</div>
        </div>
        <div style={{ marginBottom: 16 }}>
          <label style={{ color: C.textFaint, fontSize: 11.5, fontWeight: 600, textTransform: "uppercase", letterSpacing: 1, display: "block", marginBottom: 8 }}>Access Password</label>
          <input type="password" value={pw} onChange={e => { setPw(e.target.value); setError(false); }} onKeyDown={e => e.key === "Enter" && handleSubmit()} placeholder="Enter team password"
            style={{ width: "100%", padding: "12px 16px", background: C.bgDark, border: error ? `1px solid ${C.red}` : `1px solid ${C.border}`, borderRadius: 10, color: C.cream, fontSize: 15, outline: "none", boxSizing: "border-box" }} />
          {error && <div style={{ color: C.red, fontSize: 12, marginTop: 6 }}>Incorrect password. Please try again.</div>}
        </div>
        <button onClick={handleSubmit} style={{ width: "100%", padding: "13px", background: C.gold, border: "none", borderRadius: 10, color: "#1a1a1a", fontSize: 15, fontWeight: 700, cursor: "pointer", letterSpacing: 0.2 }}>
          Access Dashboard →
        </button>
        <div style={{ textAlign: "center", color: C.textFaint, fontSize: 11.5, marginTop: 20 }}>The Last Mile Solution For Digitizing Cash</div>
      </div>
      <style>{`@keyframes shake { 0%,100%{transform:translateX(0)} 25%{transform:translateX(-8px)} 75%{transform:translateX(8px)} }`}</style>
    </div>
  );
}

// ─── AUTH: MICROSOFT SIGN-IN ─────────────────────────────────────────────────
function MicrosoftSignIn({ onUseSampleData }) {
  const { instance } = useMsal();
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);

  const handleSignIn = () => {
    if (loading) return;
    // Call loginPopup() FIRST, synchronously, with no awaited work or state
    // updates beforehand — browsers only allow window.open() to produce a real
    // popup (rather than silently falling back to a same-tab navigation) when
    // it happens directly inside the user's click, with nothing async in between.
    const popupPromise = instance.loginPopup(loginRequest);
    setLoading(true);
    setError(null);
    popupPromise
      .then(() => new Promise(r => setTimeout(r, 300)))
      .catch((err) => setError(err.message || "Sign-in failed. Please try again."))
      .finally(() => setLoading(false));
  };

  return (
    <div style={{ minHeight: "100vh", background: `radial-gradient(circle at 30% 20%, #1a3329 0%, ${C.bgDark} 55%)`, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Inter', sans-serif" }}>
      <div style={{ background: C.panelDark, border: `1px solid ${C.border}`, borderRadius: 20, padding: "48px 40px", width: 420, textAlign: "center", boxShadow: "0 20px 60px rgba(0,0,0,0.4)" }}>
        <div style={{ width: 52, height: 52, background: C.gold, borderRadius: 12, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 16px", fontSize: 26, fontWeight: 900, color: "#1a1a1a", fontFamily: "'Space Grotesk', sans-serif" }}>E</div>
        <div style={{ color: C.cream, fontSize: 19, fontWeight: 700, marginBottom: 8, fontFamily: "'Space Grotesk', sans-serif" }}>Connect to Live Data</div>
        <div style={{ color: C.textMuted, fontSize: 13, marginBottom: 28, lineHeight: 1.5 }}>Sign in with your Microsoft work account to pull live data from the pilot Excel file on OneDrive.</div>
        <button onClick={handleSignIn} disabled={loading} style={{ width: "100%", padding: "13px", background: C.gold, border: "none", borderRadius: 10, color: "#1a1a1a", fontSize: 14, fontWeight: 700, cursor: "pointer", marginBottom: 12, opacity: loading ? 0.7 : 1 }}>
          {loading ? "Signing in..." : "🔑  Sign in with Microsoft"}
        </button>
        <button onClick={onUseSampleData} style={{ width: "100%", padding: "11px", background: "transparent", border: `1px solid ${C.border}`, borderRadius: 10, color: C.textMuted, fontSize: 13, cursor: "pointer" }}>
          Continue with sample data instead
        </button>
        {error && <div style={{ color: C.red, fontSize: 12, marginTop: 14 }}>{error}</div>}
      </div>
    </div>
  );
}

// ─── SURVEY PDF PARSING (lightweight, generic question-block detection) ─────
// Parses the text extracted from an Odoo survey "Results / Statistics" export.
// Handles the three real question patterns observed in these exports:
//   1. Choice:  "Answer User Choice" header, then "Label  XX.X %  N Votes" rows
//   2. Numeric: a "Max N Min N Avg N" stats line
//   3. Text:    "# User Responses" header, then "1 <response>" numbered rows
function parseSurveyText(text) {
  const rawLines = text.split("\n").map(l => l.trim()).filter(Boolean);
  const questions = [];
  let current = null;
  const pushCurrent = () => {
    if (current && (current.options.length || current.responses.length || current.stats)) {
      questions.push(current);
    }
    current = null;
  };

  for (let idx = 0; idx < rawLines.length; idx++) {
    const line = rawLines[idx];

    // Skip noise / structural markers
    if (/^\d+\s*\/\s*\d+\s+Responded/i.test(line)) {
      // A "X / Y Responded" line may ALSO carry a stats tail, e.g.
      // "1 / 1 Responded Max 29.0 Min 29.0 Avg 29.0"
      const statsMatch = line.match(/Max\s+([\d.]+)\s+Min\s+([\d.]+)\s+Avg\s+([\d.]+)/i);
      if (statsMatch && current) {
        current.type = "numeric";
        current.stats = { max: parseFloat(statsMatch[1]), min: parseFloat(statsMatch[2]), avg: parseFloat(statsMatch[3]) };
      }
      continue;
    }
    if (/^Answer\s+User Choice$/i.test(line)) { if (current) current.type = "choice"; continue; }
    if (/^#\s+User Responses$/i.test(line)) { if (current) current.type = current.type || "text"; continue; }
    if (/^Top User Responses\s+Occurrence$/i.test(line)) { if (current) current.skipOccurrence = true; continue; }
    if (/^No answers yet!?$/i.test(line)) { if (current) current.empty = true; continue; }
    if (/^Powered by/i.test(line)) continue;

    // Standalone stats line
    const standaloneStats = line.match(/^Max\s+([\d.]+)\s+Min\s+([\d.]+)\s+Avg\s+([\d.]+)$/i);
    if (standaloneStats && current) {
      current.type = "numeric";
      current.stats = { max: parseFloat(standaloneStats[1]), min: parseFloat(standaloneStats[2]), avg: parseFloat(standaloneStats[3]) };
      continue;
    }

    // Choice row: "Label   12.5 %   3 Votes"  (also handles "0.0% 0 Votes")
    const choiceMatch = line.match(/^(.+?)\s+(\d+(?:\.\d+)?)\s*%\s+(\d+)\s*Votes?$/i);
    if (choiceMatch && current) {
      current.type = "choice";
      current.options.push({ label: choiceMatch[1].trim(), pct: parseFloat(choiceMatch[2]), votes: parseInt(choiceMatch[3], 10) });
      continue;
    }

    // Numbered response row: "1 Oratile Tshepo Masala" or "1 29"
    const respMatch = line.match(/^(\d+)\s+(.+)$/);
    if (respMatch && current && (current.type === "text" || current.type === "numeric")) {
      // For numeric "Top User Responses / Occurrence" section, skip the value-count pairs
      if (!current.skipOccurrence || current.type === "text") {
        current.responses.push(respMatch[2].trim());
      }
      continue;
    }

    // Otherwise: this is either a new question title, OR a continuation of the
    // previous line if that question hasn't started collecting any data yet
    // (covers titles that wrap onto two lines in the PDF, e.g. "How many
    // customer transactions does the business" / "process per day").
    if (line.length <= 140 && !/^\d+$/.test(line)) {
      const looksLikeContinuation =
        current && current.type === null && current.options.length === 0 &&
        current.responses.length === 0 && !current.stats && !current.empty &&
        current.title.length < 100 && !/[.?!:]$/.test(current.title);
      if (looksLikeContinuation) {
        current.title = `${current.title} ${line}`;
        continue;
      }
      pushCurrent();
      current = { title: line, type: null, options: [], responses: [], stats: null, skipOccurrence: false, empty: false };
    }
  }
  pushCurrent();

  // Keep only questions that actually carry data
  return questions.filter(q => q.options.length || q.responses.length || q.stats);
}

// Extract raw text from a survey file depending on its type (PDF / XLSX / CSV).
async function extractSurveyText(file) {
  const name = (file.name || "").toLowerCase();

  if (name.endsWith(".pdf")) {
    if (!window.pdfjsLib) throw new Error("PDF reader is still loading — please try again in a moment.");
    const buf = await file.arrayBuffer();
    const pdf = await window.pdfjsLib.getDocument({ data: buf }).promise;
    let allText = "";
    const Y_TOLERANCE = 3; // points — groups text items into the same line even with small sub-pixel y differences
    for (let p = 1; p <= pdf.numPages; p++) {
      const page = await pdf.getPage(p);
      const content = await page.getTextContent();
      // Reconstruct lines by clustering text items whose y-position is within
      // Y_TOLERANCE of each other, rather than requiring an exact rounded match
      // (which can incorrectly split or merge visually-identical lines due to
      // tiny sub-pixel rendering differences between text runs).
      const items = content.items
        .filter(item => item.str !== undefined)
        .map(item => ({ x: item.transform[4], y: item.transform[5], s: item.str }))
        .sort((a, b) => b.y - a.y || a.x - b.x);

      const lines = [];
      items.forEach(item => {
        let line = lines.find(l => Math.abs(l.y - item.y) <= Y_TOLERANCE);
        if (!line) {
          line = { y: item.y, items: [] };
          lines.push(line);
        }
        line.items.push(item);
      });

      lines.forEach(line => {
        const lineStr = line.items.sort((a, b) => a.x - b.x).map(i => i.s).join(" ").replace(/\s+/g, " ").trim();
        if (lineStr) allText += lineStr + "\n";
      });
    }
    return allText;
  }

  if (name.endsWith(".xlsx") || name.endsWith(".xls") || name.endsWith(".csv")) {
    if (!window.XLSX) throw new Error("Spreadsheet reader is still loading — please try again in a moment.");
    const buf = await file.arrayBuffer();
    const wb = window.XLSX.read(buf, { type: "array" });
    // Flatten every sheet into tab-separated lines so the same line-based parser works.
    let allText = "";
    wb.SheetNames.forEach(sheetName => {
      const sheet = wb.Sheets[sheetName];
      const rows = window.XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
      rows.forEach(r => {
        const lineStr = (r || []).map(c => (c === null || c === undefined ? "" : String(c))).join(" ").replace(/\s+/g, " ").trim();
        if (lineStr) allText += lineStr + "\n";
      });
    });
    return allText;
  }

  // Fallback: plain text
  return await file.text();
}

const SURVEY_COLORS = [C.gold, C.green, C.blue, C.amber, C.red, C.forestLight, C.goldLight];

function SurveyInsights() {
  const [questions, setQuestions] = useState(null);
  const [fileName, setFileName] = useState(null);
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);
  const [hiddenIdx, setHiddenIdx] = useState(new Set()); // session-level hidden questions
  const [showPicker, setShowPicker] = useState(false);

  const handleFile = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setError(null);
    setLoading(true);
    setFileName(file.name);
    setHiddenIdx(new Set()); // reset visibility choices on each new upload
    try {
      const text = await extractSurveyText(file);
      const parsed = parseSurveyText(text);
      if (parsed.length === 0) {
        setError("Could not detect any survey questions in this file. Make sure it's an Odoo survey results export (PDF, Excel, or CSV).");
        setQuestions(null);
      } else {
        setQuestions(parsed);
      }
    } catch (err) {
      setError("Failed to read this file. " + (err.message || err));
      setQuestions(null);
    } finally {
      setLoading(false);
    }
  };

  const toggleHidden = (i) => {
    setHiddenIdx(prev => {
      const next = new Set(prev);
      if (next.has(i)) next.delete(i); else next.add(i);
      return next;
    });
  };

  const totalResponses = (q) => q.options.reduce((s, o) => s + o.votes, 0);
  const visibleQuestions = questions ? questions.filter((_, i) => !hiddenIdx.has(i)) : [];

  return (
    <div style={{ background: C.panelDark, border: `1px solid ${C.border}`, borderRadius: 14, padding: 20 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
        <div style={{ fontWeight: 700 }}>📋 Survey Insights</div>
        <div style={{ display: "flex", gap: 8 }}>
          {questions && questions.length > 0 && (
            <button onClick={() => setShowPicker(s => !s)} style={{ background: "transparent", color: C.textMuted, border: `1px solid ${C.border}`, borderRadius: 8, padding: "7px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
              {showPicker ? "Done" : `Choose questions (${visibleQuestions.length}/${questions.length})`}
            </button>
          )}
          <label style={{ background: C.gold, color: "#1a1a1a", borderRadius: 8, padding: "7px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
            {loading ? "Reading..." : "Upload Survey File"}
            <input type="file" accept=".pdf,.xlsx,.xls,.csv" onChange={handleFile} style={{ display: "none" }} disabled={loading} />
          </label>
        </div>
      </div>
      {!questions && !loading && <div style={{ color: C.textFaint, fontSize: 13 }}>Upload your Odoo survey results export — PDF, Excel, or CSV. New uploads replace the previous snapshot.</div>}
      {loading && <div style={{ color: C.textMuted, fontSize: 13 }}>Reading and parsing your survey file…</div>}
      {fileName && !loading && <div style={{ color: C.textFaint, fontSize: 11, marginBottom: 10 }}>Showing: {fileName}</div>}
      {error && <div style={{ color: C.red, fontSize: 12, marginBottom: 10 }}>{error}</div>}

      {/* Question picker — checkboxes to show/hide each question */}
      {showPicker && questions && (
        <div style={{ background: C.bgDark, border: `1px solid ${C.border}`, borderRadius: 10, padding: 14, marginBottom: 14 }}>
          <div style={{ color: C.textFaint, fontSize: 11, fontWeight: 600, marginBottom: 10, textTransform: "uppercase", letterSpacing: 1 }}>Tick the questions to display</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))", gap: 6 }}>
            {questions.map((q, i) => (
              <label key={i} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, color: C.textMuted, cursor: "pointer", padding: "4px 0" }}>
                <input type="checkbox" checked={!hiddenIdx.has(i)} onChange={() => toggleHidden(i)} style={{ accentColor: C.gold, width: 15, height: 15 }} />
                <span style={{ color: hiddenIdx.has(i) ? C.textFaint : C.cream }}>{q.title}</span>
              </label>
            ))}
          </div>
        </div>
      )}

      {questions && questions.length > 0 && (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: 14 }}>
          {questions.map((q, i) => {
            if (hiddenIdx.has(i)) return null;
            return (
            <div key={i} style={{ background: C.bgDark, border: `1px solid ${C.border}`, borderRadius: 10, padding: 16 }}>
              <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 12, color: C.cream }}>{q.title}</div>

              {/* Choice questions → pie chart + legend table */}
              {q.type === "choice" && q.options.length > 0 && (
                <div>
                  {totalResponses(q) > 0 && (
                    <ResponsiveContainer width="100%" height={200}>
                      <PieChart>
                        <Pie data={q.options.filter(o => o.votes > 0)} dataKey="votes" nameKey="label" cx="50%" cy="50%" outerRadius={70} label={(e) => `${e.pct}%`}>
                          {q.options.filter(o => o.votes > 0).map((o, j) => <Cell key={j} fill={SURVEY_COLORS[j % SURVEY_COLORS.length]} />)}
                        </Pie>
                        <Tooltip contentStyle={{ background: C.panelDark2, border: `1px solid ${C.border}`, color: C.cream, fontSize: 12 }} formatter={(v, n) => [`${v} votes`, n]} />
                      </PieChart>
                    </ResponsiveContainer>
                  )}
                  <div style={{ display: "flex", flexDirection: "column", gap: 5, marginTop: 8 }}>
                    {q.options.map((o, j) => (
                      <div key={j} style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <span style={{ width: 8, height: 8, borderRadius: 4, background: SURVEY_COLORS[j % SURVEY_COLORS.length], flexShrink: 0, opacity: o.votes > 0 ? 1 : 0.3 }} />
                        <span style={{ flex: 1, fontSize: 12, color: o.votes > 0 ? C.textMuted : C.textFaint }}>{o.label}</span>
                        <span style={{ fontSize: 12, color: o.votes > 0 ? C.cream : C.textFaint, fontWeight: 600 }}>{o.pct}% ({o.votes})</span>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {/* Numeric questions → stat cards */}
              {q.type === "numeric" && q.stats && (
                <div style={{ display: "flex", gap: 8 }}>
                  {[{ k: "Min", v: q.stats.min }, { k: "Avg", v: q.stats.avg }, { k: "Max", v: q.stats.max }].map(s => (
                    <div key={s.k} style={{ flex: 1, background: C.panelDark, border: `1px solid ${C.border}`, borderRadius: 8, padding: "10px 8px", textAlign: "center" }}>
                      <div style={{ color: C.textFaint, fontSize: 10.5, marginBottom: 3 }}>{s.k}</div>
                      <div style={{ color: C.gold, fontSize: 18, fontWeight: 700 }}>{Number.isFinite(s.v) ? (Math.round(s.v * 10) / 10) : "—"}</div>
                    </div>
                  ))}
                </div>
              )}

              {/* Text questions → response chips */}
              {q.type === "text" && q.responses.length > 0 && (
                <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                  {q.responses.slice(0, 20).map((r, j) => (
                    <span key={j} style={{ background: C.forest, color: C.textMuted, borderRadius: 6, padding: "3px 9px", fontSize: 11.5 }}>{r}</span>
                  ))}
                  {q.responses.length > 20 && <span style={{ color: C.textFaint, fontSize: 11.5 }}>+{q.responses.length - 20} more</span>}
                </div>
              )}
            </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ─── MAIN DASHBOARD ──────────────────────────────────────────────────────────
function Dashboard({ onboardingEvents, payments, dataSource, lastUpdated, onRefresh, refreshing }) {
  const [tab, setTab] = useState("dashboard");

  const [dashPhase, setDashPhase] = useState("0");
  const [incPhase, setIncPhase] = useState("0");
  const [incDateFrom, setIncDateFrom] = useState("");
  const [incDateTo, setIncDateTo] = useState("");
  const [dashDateFrom, setDashDateFrom] = useState("");
  const [dashDateTo, setDashDateTo] = useState("");
  const [triggersPhase, setTriggersPhase] = useState("0");
  const [triggersDateFrom, setTriggersDateFrom] = useState("");
  const [triggersDateTo, setTriggersDateTo] = useState("");
  const [merchantsPhase, setMerchantsPhase] = useState("0");
  const [merchantsDateFrom, setMerchantsDateFrom] = useState("");
  const [merchantsDateTo, setMerchantsDateTo] = useState("");
  const [internalPhase, setInternalPhase] = useState("0");
  const [internalDateFrom, setInternalDateFrom] = useState("");
  const [internalDateTo, setInternalDateTo] = useState("");
  const [auditPhase, setAuditPhase] = useState("0");
  const [auditDateFrom, setAuditDateFrom] = useState("");
  const [auditDateTo, setAuditDateTo] = useState("");
  const [expandedTxnIncentive, setExpandedTxnIncentive] = useState(null);

  const [configs, setConfigs] = useState(() => loadIncentiveConfigs());
  const [dashTargets, setDashTargets] = useState(() => loadDashboardTargets());
  const [paidTriggers, setPaidTriggers] = useState(() => loadPaidTriggers());
  const [triggerFilter, setTriggerFilter] = useState("all");
  const [search, setSearch] = useState("");
  const [expandedMerchant, setExpandedMerchant] = useState(null);
  const [expandedCoordinator, setExpandedCoordinator] = useState(null);
  const [expandedCounterparty, setExpandedCounterparty] = useState(null);

  useEffect(() => { saveIncentiveConfigs(configs); }, [configs]);
  useEffect(() => { saveDashboardTargets(dashTargets); }, [dashTargets]);
  useEffect(() => { savePaidTriggers(paidTriggers); }, [paidTriggers]);

  const config = configs[incPhase] || DEFAULT_INCENTIVE_CONFIG;
  const updateConfig = (key, value) => setConfigs(prev => ({ ...prev, [incPhase]: { ...prev[incPhase], [key]: value } }));
  const updateMilestone = (idx, field, value) => {
    const newMilestones = [...config.milestones];
    newMilestones[idx] = { ...newMilestones[idx], [field]: value };
    updateConfig("milestones", newMilestones);
  };
  const addMilestone = () => updateConfig("milestones", [...config.milestones, { count: 0, bonus: 0 }]);
  const removeMilestone = (idx) => updateConfig("milestones", config.milestones.filter((_, i) => i !== idx));

  const updateTarget = (key, value) => setDashTargets(prev => ({ ...prev, [dashPhase]: { ...prev[dashPhase], [key]: value } }));

  const inRange = (dateStr, from, to) => {
    if (!from && !to) return true;
    if (!dateStr) return false;
    const parts = dateStr.split(",")[0].trim().split("/");
    if (parts.length !== 3) return false;
    const [day, month, year] = parts.map(Number);
    const d = new Date(year, month - 1, day);
    if (from && d < new Date(from)) return false;
    if (to && d > new Date(to)) return false;
    return true;
  };

  const dedupedOnboarding = useMemo(() => dedupeOnboardingEvents(onboardingEvents), [onboardingEvents]);

  const incOnboarding = useMemo(() => dedupedOnboarding.filter(e => inRange(e.dlt_close, incDateFrom, incDateTo)), [dedupedOnboarding, incDateFrom, incDateTo]);
  const incPayments = useMemo(() => payments.filter(p => inRange(p.dlt_close, incDateFrom, incDateTo)), [payments, incDateFrom, incDateTo]);
  const merchantNames = useMemo(() => getAllMerchantNames(incOnboarding, incPayments), [incOnboarding, incPayments]);
  const incentiveResults = useMemo(() =>
    merchantNames.map(name => calculateMerchantIncentive(name, incOnboarding, incPayments, config)),
    [merchantNames, incOnboarding, incPayments, config]
  );
  const filteredResults = incentiveResults.filter(r => r.merchantName.toLowerCase().includes(search.toLowerCase()));

  // ── Internal Transfers tab data: every transaction touching any of the 5
  // special accounts (Exto + 4 field coordinators), grouped by coordinator
  // then by counterparty. Uses the full, unfiltered payments list — this is
  // an always-on operational ledger, independent of any phase/date selector.
  const dashOnboarding = useMemo(() => dedupedOnboarding.filter(e => inRange(e.dlt_close, dashDateFrom, dashDateTo)), [dedupedOnboarding, dashDateFrom, dashDateTo]);
  const dashPayments = useMemo(() => payments.filter(p => inRange(p.dlt_close, dashDateFrom, dashDateTo)), [payments, dashDateFrom, dashDateTo]);
  const dashMerchantNames = useMemo(() => getAllMerchantNames(dashOnboarding, dashPayments), [dashOnboarding, dashPayments]);

  // ── Filtered data + per-merchant results for Triggers / Merchants / Internal Transfers / Audit ──
  // Each tab has its OWN independent phase + date range, exactly like Dashboard and Incentives.
  const triggersConfig = configs[triggersPhase] || DEFAULT_INCENTIVE_CONFIG;
  const triggersOnboarding = useMemo(() => dedupedOnboarding.filter(e => inRange(e.dlt_close, triggersDateFrom, triggersDateTo)), [dedupedOnboarding, triggersDateFrom, triggersDateTo]);
  const triggersPayments = useMemo(() => payments.filter(p => inRange(p.dlt_close, triggersDateFrom, triggersDateTo)), [payments, triggersDateFrom, triggersDateTo]);
  const triggersMerchantNames = useMemo(() => getAllMerchantNames(triggersOnboarding, triggersPayments), [triggersOnboarding, triggersPayments]);
  const triggersResults = useMemo(() =>
    triggersMerchantNames.map(name => calculateMerchantIncentive(name, triggersOnboarding, triggersPayments, triggersConfig)),
    [triggersMerchantNames, triggersOnboarding, triggersPayments, triggersConfig]
  );

  const merchantsConfig = configs[merchantsPhase] || DEFAULT_INCENTIVE_CONFIG;
  const merchantsOnboarding = useMemo(() => dedupedOnboarding.filter(e => inRange(e.dlt_close, merchantsDateFrom, merchantsDateTo)), [dedupedOnboarding, merchantsDateFrom, merchantsDateTo]);
  const merchantsPayments = useMemo(() => payments.filter(p => inRange(p.dlt_close, merchantsDateFrom, merchantsDateTo)), [payments, merchantsDateFrom, merchantsDateTo]);
  const merchantsMerchantNames = useMemo(() => getAllMerchantNames(merchantsOnboarding, merchantsPayments), [merchantsOnboarding, merchantsPayments]);
  const merchantsResults = useMemo(() =>
    merchantsMerchantNames.map(name => calculateMerchantIncentive(name, merchantsOnboarding, merchantsPayments, merchantsConfig)),
    [merchantsMerchantNames, merchantsOnboarding, merchantsPayments, merchantsConfig]
  );
  const merchantsFiltered = merchantsResults.filter(r => r.merchantName.toLowerCase().includes(search.toLowerCase()));

  const internalPayments = useMemo(() => payments.filter(p => inRange(p.dlt_close, internalDateFrom, internalDateTo)), [payments, internalDateFrom, internalDateTo]);

  const internalTransfersByCoordinator = useMemo(() => {
    const byCoordinator = {};
    internalPayments.forEach(p => {
      const senderIsSpecial = isSpecialAccountName(p.sender);
      const receiverIsSpecial = isSpecialAccountName(p.receiver);
      if (!senderIsSpecial && !receiverIsSpecial) return;

      // A transaction can have a special account on one or both sides.
      // Record it once per special-account party involved.
      [
        { coordinator: p.sender, counterparty: p.receiver, direction: "sent", isSpecial: senderIsSpecial },
        { coordinator: p.receiver, counterparty: p.sender, direction: "received", isSpecial: receiverIsSpecial },
      ].forEach(({ coordinator, counterparty, direction, isSpecial }) => {
        if (!isSpecial || !coordinator) return;
        byCoordinator[coordinator] = byCoordinator[coordinator] || { totalSent: 0, totalReceived: 0, counterparties: {} };
        byCoordinator[coordinator].counterparties[counterparty] = byCoordinator[coordinator].counterparties[counterparty] || { sent: 0, received: 0, txns: [] };
        if (direction === "sent") {
          byCoordinator[coordinator].totalSent += p.amount;
          byCoordinator[coordinator].counterparties[counterparty].sent += p.amount;
        } else {
          byCoordinator[coordinator].totalReceived += p.amount;
          byCoordinator[coordinator].counterparties[counterparty].received += p.amount;
        }
        byCoordinator[coordinator].counterparties[counterparty].txns.push(p);
      });
    });
    return byCoordinator;
  }, [internalPayments]);

  const auditOnboarding = useMemo(() => dedupedOnboarding.filter(e => inRange(e.dlt_close, auditDateFrom, auditDateTo)), [dedupedOnboarding, auditDateFrom, auditDateTo]);
  const auditPayments = useMemo(() => payments.filter(p => inRange(p.dlt_close, auditDateFrom, auditDateTo)), [payments, auditDateFrom, auditDateTo]);
  const auditMerchantNamesForTab = useMemo(() => getAllMerchantNames(auditOnboarding, auditPayments), [auditOnboarding, auditPayments]);

  // ── Transaction Audit: every payment tagged counted/excluded with a reason ──
  // Uses its OWN independent phase + date filter, like every other tab.
  const auditedTransactions = useMemo(() => {
    const merchantNameSet = new Set(auditMerchantNamesForTab);
    return auditTransactions(auditPayments, merchantNameSet);
  }, [auditPayments, auditMerchantNamesForTab]);
  const [auditFilter, setAuditFilter] = useState("all");
  const auditFiltered = auditedTransactions.filter(t =>
    auditFilter === "all" ? true : auditFilter === "counted" ? t.counted : !t.counted
  );

  const dashStats = useMemo(() => {
    const totalMerchants = dashMerchantNames.length;
    const consumerEvents = dashOnboarding.filter(e => !e.isMerchantOnboarding);
    const uniqueConsumers = new Set(consumerEvents.map(e => e.user_name)).size;
    let totalQualifying = 0;
    dashMerchantNames.forEach(m => {
      const r = calculateMerchantIncentive(m, dashOnboarding, dashPayments, config);
      totalQualifying += r.netCount;
    });
    const totalVolume = dashPayments
      .filter(p => p.sender !== EXTO_BACKEND_NAME && p.receiver !== EXTO_BACKEND_NAME && (ALWAYS_COUNT_TYPES.has(p.txn_type) || PAIR_TYPES[p.txn_type]))
      .reduce((s, p) => s + p.amount, 0);
    return { totalMerchants, uniqueConsumers, totalQualifying, totalVolume, consumerEvents };
  }, [dashMerchantNames, dashOnboarding, dashPayments, config]);

  const onboardTrend = useMemo(() => {
    const byMonth = {};
    dashOnboarding.forEach(e => {
      if (!e.dlt_close) return;
      const datePart = e.dlt_close.split(",")[0].trim();
      const parts = datePart.split("/");
      if (parts.length !== 3) return;
      const key = `${parts[2]}-${parts[1].padStart(2, "0")}`;
      byMonth[key] = byMonth[key] || { merchants: new Set(), consumers: new Set() };
      if (e.isMerchantOnboarding) byMonth[key].merchants.add(e.user_name);
      else byMonth[key].consumers.add(e.user_name);
    });
    return Object.entries(byMonth).sort().map(([k, v]) => ({ month: k, merchants: v.merchants.size, consumers: v.consumers.size }));
  }, [dashOnboarding]);

  const triggers = useMemo(() => {
    const list = [];
    triggersResults.forEach(r => {
      if (r.activationEligible) list.push({ id: `${r.merchantName}-activation`, merchant: r.merchantName, type: "Activation Bonus", message: `Reached ${triggersConfig.activationThreshold}+ qualifying transactions`, amount: triggersConfig.activationBonus, color: C.green });
      r.milestonesHit.forEach(ms => list.push({ id: `${r.merchantName}-ms-${ms.count}`, merchant: r.merchantName, type: "Milestone", message: `${ms.count}-transaction milestone reached`, amount: ms.bonus, color: C.amber }));
      if (r.consumerCapHit) list.push({ id: `${r.merchantName}-cap`, merchant: r.merchantName, type: "Cap Reached", message: `Consumer onboarding cap (${triggersConfig.consumerCap}) reached`, amount: 0, color: C.red });
      if (r.onboardingIncentive > 0) list.push({ id: `${r.merchantName}-onb`, merchant: r.merchantName, type: "Onboarding", message: `${r.cappedConsumers} consumers onboarded`, amount: r.onboardingIncentive, color: C.blue });
      // Transaction Incentive: ONE running trigger per merchant, summing fixed + variable
      // across all currently-qualifying transactions (within this tab's phase/date filter).
      // Resets to zero once marked paid, then accumulates again from new transactions.
      const txnIncentiveTotal = Math.round((r.fixedIncentive + r.variableIncentive) * 100) / 100;
      if (txnIncentiveTotal > 0) {
        list.push({
          id: `${r.merchantName}-txnIncentive`, merchant: r.merchantName, type: "Transaction Incentive",
          message: `${r.netCount} qualifying transactions (fixed BWP ${r.fixedIncentive} + variable BWP ${r.variableIncentive})`,
          amount: txnIncentiveTotal, color: C.blue, expandable: true,
          qualifyingTransactions: r.qualifyingTransactions, fixedRate: triggersConfig.fixedRate, variableRate: triggersConfig.variableRate, variableCap: triggersConfig.variableCap,
        });
      }
    });
    return list;
  }, [triggersResults, triggersConfig]);

  const pendingTriggers = triggers.filter(t => !paidTriggers.has(t.id) && (triggerFilter === "all" || t.type === triggerFilter));
  const paidList = triggers.filter(t => paidTriggers.has(t.id));
  const totalDue = pendingTriggers.reduce((s, t) => s + t.amount, 0);
  const groupedByMerchant = {};
  pendingTriggers.forEach(t => { groupedByMerchant[t.merchant] = groupedByMerchant[t.merchant] || []; groupedByMerchant[t.merchant].push(t); });

  const markPaid = (id) => setPaidTriggers(prev => new Set([...prev, id]));

  const tabs = [
    { id: "dashboard", label: "📊 Dashboard" },
    { id: "incentives", label: "💰 Incentives" },
    { id: "triggers", label: `🔔 Triggers ${pendingTriggers.length ? `(${pendingTriggers.length})` : ""}` },
    { id: "merchants", label: "🏪 Merchants" },
    { id: "internal", label: "🔄 Internal Transfers" },
    { id: "audit", label: "🔍 Transaction Audit" },
  ];

  const currentTarget = dashTargets[dashPhase];

  return (
    <div style={{ minHeight: "100vh", background: C.bgDark, fontFamily: "'Inter', system-ui, sans-serif", color: C.cream }}>
      <div style={{ background: C.panelDark, borderBottom: `1px solid ${C.border}`, padding: "0 32px", display: "flex", alignItems: "center", height: 64, flexWrap: "wrap" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, flex: 1 }}>
          <div style={{ width: 30, height: 30, background: C.gold, borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 900, fontSize: 15, color: "#1a1a1a", fontFamily: "'Space Grotesk', sans-serif" }}>E</div>
          <div>
            <span style={{ fontWeight: 700, fontSize: 16, fontFamily: "'Space Grotesk', sans-serif" }}>Exto Pay</span>
            <span style={{ color: C.textFaint, fontSize: 13, marginLeft: 8 }}>PulaConnect · Botswana</span>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 14, marginRight: 16 }}>
          <Badge color={dataSource === "live" ? C.green : C.amber}>{dataSource === "live" ? "🟢 Live OneDrive Data" : "🟡 Sample Data"}</Badge>
          {lastUpdated && <span style={{ color: C.textFaint, fontSize: 12 }}>Updated {lastUpdated}</span>}
          <button onClick={onRefresh} disabled={refreshing} style={{ background: C.bgDark, border: `1px solid ${C.border}`, borderRadius: 8, padding: "6px 12px", color: C.cream, fontSize: 12, fontWeight: 500, cursor: refreshing ? "default" : "pointer", display: "flex", alignItems: "center", gap: 6, opacity: refreshing ? 0.6 : 1 }}>
            <span style={{ display: "inline-block", animation: refreshing ? "spin 1s linear infinite" : "none" }}>🔄</span>
            {refreshing ? "Refreshing..." : "Refresh"}
          </button>
        </div>
        <div style={{ display: "flex", gap: 6 }}>
          {tabs.map(t => (
            <button key={t.id} onClick={() => setTab(t.id)} style={{ background: tab === t.id ? C.gold : "transparent", color: tab === t.id ? "#1a1a1a" : C.textMuted, border: "none", borderRadius: 8, padding: "7px 14px", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>{t.label}</button>
          ))}
        </div>
      </div>
      <style>{`@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>

      <div style={{ padding: "24px 32px", maxWidth: 1440, margin: "0 auto" }}>

        {tab === "dashboard" && (
          <div>
            <div style={{ background: C.panelDark, border: `1px solid ${C.border}`, borderRadius: 14, padding: "14px 20px", marginBottom: 20, display: "flex", gap: 20, flexWrap: "wrap", alignItems: "center" }}>
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <span style={{ color: C.textFaint, fontSize: 11, fontWeight: 600 }}>PHASE</span>
                {["0", "1", "2", "3"].map(p => (
                  <button key={p} onClick={() => setDashPhase(p)} style={{ background: dashPhase === p ? C.gold : "transparent", color: dashPhase === p ? "#1a1a1a" : C.textMuted, border: `1px solid ${dashPhase === p ? C.gold : C.border}`, borderRadius: 7, padding: "5px 12px", fontSize: 12, cursor: "pointer" }}>Phase {p}</button>
                ))}
              </div>
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <span style={{ color: C.textFaint, fontSize: 11, fontWeight: 600 }}>DATE RANGE</span>
                <input type="date" value={dashDateFrom} onChange={e => setDashDateFrom(e.target.value)} style={{ background: C.bgDark, border: `1px solid ${C.border}`, borderRadius: 6, padding: "5px 8px", color: C.cream, fontSize: 12 }} />
                <span style={{ color: C.textFaint }}>→</span>
                <input type="date" value={dashDateTo} onChange={e => setDashDateTo(e.target.value)} style={{ background: C.bgDark, border: `1px solid ${C.border}`, borderRadius: 6, padding: "5px 8px", color: C.cream, fontSize: 12 }} />
              </div>
              <span style={{ color: C.textFaint, fontSize: 11, fontStyle: "italic" }}>(independent of the Incentives tab's selectors)</span>
            </div>

            <div style={{ display: "flex", gap: 14, marginBottom: 20, flexWrap: "wrap" }}>
              <KPICard icon="🏪" label="Merchants" value={dashStats.totalMerchants} color={C.gold} />
              <KPICard icon="👥" label="Consumers" value={dashStats.uniqueConsumers} color={C.green} />
              <KPICard icon="⚡" label="Qualifying Txns" value={dashStats.totalQualifying} color={C.blue} />
              <KPICard icon="💰" label="Volume" value={`BWP ${Math.round(dashStats.totalVolume).toLocaleString()}`} color={C.gold} />
            </div>

            <div style={{ background: C.panelDark, border: `1px solid ${C.border}`, borderRadius: 14, padding: 20, marginBottom: 20 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
                <div style={{ fontWeight: 700 }}>Phase {dashPhase} Targets</div>
                <div style={{ display: "flex", gap: 10 }}>
                  <NumberInput label="Merchants target" value={currentTarget.merchants} onChange={v => updateTarget("merchants", v)} />
                  <NumberInput label="Consumers target" value={currentTarget.consumers} onChange={v => updateTarget("consumers", v)} />
                  <NumberInput label="Txns target" value={currentTarget.p2mTransactions} onChange={v => updateTarget("p2mTransactions", v)} />
                </div>
              </div>
              <TargetBar label="Merchants onboarded" actual={dashStats.totalMerchants} target={currentTarget.merchants} />
              <TargetBar label="Consumers onboarded" actual={dashStats.uniqueConsumers} target={currentTarget.consumers} />
              <TargetBar label="Qualifying transactions" actual={dashStats.totalQualifying} target={currentTarget.p2mTransactions} suffix="+" />
            </div>

            <div style={{ background: C.panelDark, border: `1px solid ${C.border}`, borderRadius: 14, padding: 20, marginBottom: 20 }}>
              <div style={{ fontWeight: 700, marginBottom: 14 }}>Onboarding Trend (Merchants + Consumers)</div>
              <ResponsiveContainer width="100%" height={220}>
                <LineChart data={onboardTrend}>
                  <XAxis dataKey="month" tick={{ fill: C.textFaint, fontSize: 11 }} />
                  <YAxis tick={{ fill: C.textFaint, fontSize: 11 }} />
                  <Tooltip contentStyle={{ background: C.panelDark2, border: `1px solid ${C.border}`, color: C.cream }} />
                  <Line type="monotone" dataKey="merchants" stroke={C.gold} strokeWidth={2} />
                  <Line type="monotone" dataKey="consumers" stroke={C.green} strokeWidth={2} />
                </LineChart>
              </ResponsiveContainer>
            </div>

            <SurveyInsights />
          </div>
        )}

        {tab === "incentives" && (
          <div>
            <div style={{ display: "flex", gap: 16, marginBottom: 16, alignItems: "center", flexWrap: "wrap" }}>
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <span style={{ color: C.textFaint, fontSize: 11, fontWeight: 600 }}>PHASE</span>
                {["0", "1", "2", "3"].map(p => (
                  <button key={p} onClick={() => setIncPhase(p)} style={{ background: incPhase === p ? C.gold : "transparent", color: incPhase === p ? "#1a1a1a" : C.textMuted, border: `1px solid ${incPhase === p ? C.gold : C.border}`, borderRadius: 7, padding: "5px 12px", fontSize: 12, cursor: "pointer" }}>Phase {p}</button>
                ))}
              </div>
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <span style={{ color: C.textFaint, fontSize: 11, fontWeight: 600 }}>DATE RANGE</span>
                <input type="date" value={incDateFrom} onChange={e => setIncDateFrom(e.target.value)} style={{ background: C.bgDark, border: `1px solid ${C.border}`, borderRadius: 6, padding: "5px 8px", color: C.cream, fontSize: 12 }} />
                <span style={{ color: C.textFaint }}>→</span>
                <input type="date" value={incDateTo} onChange={e => setIncDateTo(e.target.value)} style={{ background: C.bgDark, border: `1px solid ${C.border}`, borderRadius: 6, padding: "5px 8px", color: C.cream, fontSize: 12 }} />
              </div>
            </div>

            <div style={{ background: C.panelDark, border: `1px solid ${C.border}`, borderRadius: 14, padding: 20, marginBottom: 16 }}>
              <div style={{ fontWeight: 700, marginBottom: 14 }}>Phase {incPhase} Incentive Variables</div>
              <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 16 }}>
                <NumberInput label="Onboarding rate (BWP)" value={config.consumerRate} onChange={v => updateConfig("consumerRate", v)} />
                <NumberInput label="Onboarding cap" value={config.consumerCap} onChange={v => updateConfig("consumerCap", v)} />
                <NumberInput label="Activation bonus (BWP)" value={config.activationBonus} onChange={v => updateConfig("activationBonus", v)} />
                <NumberInput label="Activation threshold" value={config.activationThreshold} onChange={v => updateConfig("activationThreshold", v)} suffix="txns" />
                <NumberInput label="Fixed rate (BWP)" value={config.fixedRate} onChange={v => updateConfig("fixedRate", v)} />
                <NumberInput label="Variable rate" value={config.variableRate} onChange={v => updateConfig("variableRate", v)} suffix="%" />
                <NumberInput label="Variable cap (BWP)" value={config.variableCap} onChange={v => updateConfig("variableCap", v)} />
                <NumberInput label="Txn cap" value={config.txnCap} onChange={v => updateConfig("txnCap", v)} />
              </div>
              <div style={{ color: C.textFaint, fontSize: 11, fontWeight: 600, marginBottom: 8 }}>MILESTONE TIERS</div>
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "flex-end" }}>
                {config.milestones.map((ms, i) => (
                  <div key={i} style={{ display: "flex", gap: 6, alignItems: "flex-end", background: C.bgDark, padding: 8, borderRadius: 8, border: `1px solid ${C.border}` }}>
                    <NumberInput label="At txns" value={ms.count} onChange={v => updateMilestone(i, "count", v)} width={55} />
                    <NumberInput label="Bonus BWP" value={ms.bonus} onChange={v => updateMilestone(i, "bonus", v)} width={55} />
                    <button onClick={() => removeMilestone(i)} style={{ background: C.red + "22", color: C.red, border: "none", borderRadius: 6, padding: "5px 9px", cursor: "pointer", fontSize: 12 }}>✕</button>
                  </div>
                ))}
                <button onClick={addMilestone} style={{ background: C.green + "22", color: C.green, border: `1px solid ${C.green}55`, borderRadius: 6, padding: "8px 12px", cursor: "pointer", fontSize: 12 }}>+ Add Tier</button>
              </div>
            </div>

            <input placeholder="🔍 Search merchant..." value={search} onChange={e => setSearch(e.target.value)} style={{ background: C.panelDark, border: `1px solid ${C.border}`, borderRadius: 10, padding: "9px 16px", color: C.cream, fontSize: 13, outline: "none", width: 260, marginBottom: 14 }} />

            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {filteredResults.map(r => (
                <div key={r.merchantName} style={{ background: C.panelDark, border: `1px solid ${C.border}`, borderRadius: 12, padding: 16 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 10 }}>
                    <span style={{ fontWeight: 700 }}>{r.merchantName}</span>
                    <Badge color={C.gold}>Total: BWP {r.totalIncentive.toLocaleString()}</Badge>
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: 8, fontSize: 12 }}>
                    <div><div style={{ color: C.textFaint }}>Onboarding ({r.cappedConsumers}/{r.consumerCount})</div><div style={{ color: C.green, fontWeight: 700 }}>BWP {r.onboardingIncentive}</div></div>
                    <div><div style={{ color: C.textFaint }}>Activation</div><div style={{ color: r.activationEligible ? C.green : C.textFaint, fontWeight: 700 }}>{r.activationEligible ? `BWP ${config.activationBonus}` : "Pending"}</div></div>
                    <div><div style={{ color: C.textFaint }}>Fixed ({r.cappedCount} txns)</div><div style={{ color: C.blue, fontWeight: 700 }}>BWP {r.fixedIncentive}</div></div>
                    <div><div style={{ color: C.textFaint }}>Variable</div><div style={{ color: C.gold, fontWeight: 700 }}>BWP {r.variableIncentive}</div></div>
                    <div><div style={{ color: C.textFaint }}>Milestones</div><div style={{ color: C.amber, fontWeight: 700 }}>BWP {r.milestoneIncentive}</div></div>
                  </div>
                </div>
              ))}
              {filteredResults.length === 0 && <div style={{ color: C.textFaint, padding: 20, textAlign: "center" }}>No merchants found for this date range / search.</div>}
            </div>
          </div>
        )}

        {tab === "triggers" && (
          <div>
            <PhaseDateFilterBar phase={triggersPhase} setPhase={setTriggersPhase} dateFrom={triggersDateFrom} setDateFrom={setTriggersDateFrom} dateTo={triggersDateTo} setDateTo={setTriggersDateTo} />
            <div style={{ display: "flex", gap: 16, marginBottom: 20, alignItems: "center", flexWrap: "wrap" }}>
              <div style={{ background: C.panelDark, border: `1px solid ${C.gold}55`, borderRadius: 12, padding: "16px 24px" }}>
                <div style={{ color: C.textFaint, fontSize: 11 }}>TOTAL BWP DUE</div>
                <div style={{ color: C.gold, fontSize: 28, fontWeight: 700 }}>BWP {totalDue.toLocaleString()}</div>
              </div>
              <select value={triggerFilter} onChange={e => setTriggerFilter(e.target.value)} style={{ background: C.panelDark, border: `1px solid ${C.border}`, borderRadius: 8, padding: "8px 12px", color: C.cream, fontSize: 13 }}>
                <option value="all">All trigger types</option>
                <option value="Activation Bonus">Activation Bonus</option>
                <option value="Milestone">Milestone</option>
                <option value="Onboarding">Onboarding</option>
                <option value="Cap Reached">Cap Reached</option>
                <option value="Transaction Incentive">Transaction Incentive</option>
              </select>
            </div>

            {Object.entries(groupedByMerchant).map(([merchant, items]) => (
              <div key={merchant} style={{ background: C.panelDark, border: `1px solid ${C.border}`, borderRadius: 12, padding: 16, marginBottom: 10 }}>
                <div style={{ fontWeight: 700, marginBottom: 10 }}>{merchant}</div>
                {items.map(t => (
                  <div key={t.id} style={{ borderTop: `1px solid ${C.border}` }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 0" }}>
                      <Badge color={t.color}>{t.type}</Badge>
                      <span style={{ flex: 1, fontSize: 13, color: C.textMuted }}>{t.message}</span>
                      <span style={{ color: t.color, fontWeight: 700, fontSize: 13 }}>BWP {t.amount}</span>
                      {t.expandable && (
                        <button onClick={() => setExpandedTxnIncentive(expandedTxnIncentive === t.id ? null : t.id)} style={{ background: "transparent", color: C.textFaint, border: `1px solid ${C.border}`, borderRadius: 6, padding: "4px 8px", fontSize: 11, cursor: "pointer" }}>
                          {expandedTxnIncentive === t.id ? "▲ Hide" : "▼ Details"}
                        </button>
                      )}
                      <button onClick={() => markPaid(t.id)} style={{ background: C.green + "22", color: C.green, border: `1px solid ${C.green}55`, borderRadius: 6, padding: "4px 10px", fontSize: 11, cursor: "pointer" }}>Mark as Paid</button>
                    </div>
                    {t.expandable && expandedTxnIncentive === t.id && (
                      <div style={{ background: C.bgDark, borderRadius: 8, padding: 12, marginBottom: 10, overflowX: "auto" }}>
                        <table style={{ width: "100%", fontSize: 12, borderCollapse: "collapse" }}>
                          <thead><tr>{["Type", "Sender", "Receiver", "Amount", "Fixed", "Variable"].map(h => <th key={h} style={{ textAlign: "left", color: C.textFaint, padding: "4px 8px", borderBottom: `1px solid ${C.border}` }}>{h}</th>)}</tr></thead>
                          <tbody>
                            {t.qualifyingTransactions.map((p, i) => {
                              const variable = Math.min(p.amount * (t.variableRate / 100), t.variableCap);
                              return (
                                <tr key={i}>
                                  <td style={{ padding: "4px 8px", color: C.textMuted }}>{p.txn_type}</td>
                                  <td style={{ padding: "4px 8px" }}>{p.sender}</td>
                                  <td style={{ padding: "4px 8px" }}>{p.receiver}</td>
                                  <td style={{ padding: "4px 8px", color: C.gold }}>{p.amount}</td>
                                  <td style={{ padding: "4px 8px", color: C.blue }}>BWP {t.fixedRate}</td>
                                  <td style={{ padding: "4px 8px", color: C.amber }}>BWP {Math.round(variable * 100) / 100}</td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                ))}
              </div>
            ))}
            {pendingTriggers.length === 0 && <div style={{ color: C.textFaint, padding: 20 }}>No pending triggers for the current phase/date range.</div>}

            {paidList.length > 0 && (
              <div style={{ marginTop: 20 }}>
                <div style={{ fontWeight: 700, marginBottom: 10, color: C.textMuted }}>📜 Paid History ({paidList.length})</div>
                {paidList.map(t => (
                  <div key={t.id} style={{ display: "flex", gap: 10, padding: "8px 12px", color: C.textFaint, fontSize: 12, borderBottom: `1px solid ${C.border}` }}>
                    <span>{t.merchant}</span><span>—</span><span>{t.message}</span><span style={{ marginLeft: "auto" }}>BWP {t.amount}</span>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        {tab === "merchants" && (
          <div>
            <PhaseDateFilterBar phase={merchantsPhase} setPhase={setMerchantsPhase} dateFrom={merchantsDateFrom} setDateFrom={setMerchantsDateFrom} dateTo={merchantsDateTo} setDateTo={setMerchantsDateTo} />
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16 }}>
              <div style={{ fontWeight: 700, fontSize: 16 }}>All Merchants ({merchantsMerchantNames.length})</div>
              <input placeholder="🔍 Search..." value={search} onChange={e => setSearch(e.target.value)} style={{ background: C.panelDark, border: `1px solid ${C.border}`, borderRadius: 10, padding: "8px 14px", color: C.cream, fontSize: 13, outline: "none", width: 220 }} />
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {merchantsFiltered.map(r => (
                <div key={r.merchantName} style={{ border: `1px solid ${expandedMerchant === r.merchantName ? C.gold + "55" : C.border}`, borderRadius: 12, overflow: "hidden" }}>
                  <div onClick={() => setExpandedMerchant(expandedMerchant === r.merchantName ? null : r.merchantName)} style={{ display: "flex", alignItems: "center", padding: "12px 18px", cursor: "pointer", background: expandedMerchant === r.merchantName ? C.gold + "0c" : C.panelDark, gap: 14 }}>
                    <span style={{ flex: 1, fontWeight: 600 }}>{r.merchantName}</span>
                    <span style={{ color: C.textFaint, fontSize: 12 }}>{r.consumerCount} consumers · {r.netCount} qualifying txns</span>
                    <Badge color={C.gold}>BWP {r.totalIncentive.toLocaleString()}</Badge>
                    <span style={{ color: C.textFaint }}>{expandedMerchant === r.merchantName ? "▲" : "▼"}</span>
                  </div>
                  {expandedMerchant === r.merchantName && (
                    <div style={{ background: C.bgDark, padding: 16, borderTop: `1px solid ${C.border}` }}>
                      <div style={{ color: C.textFaint, fontSize: 11, marginBottom: 8 }}>CONSUMERS ONBOARDED</div>
                      <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 14 }}>
                        {r.consumers.map((c, i) => <span key={i} style={{ background: C.green + "14", color: C.green, borderRadius: 6, padding: "3px 9px", fontSize: 12 }}>{c.user_name}</span>)}
                        {r.consumers.length === 0 && <span style={{ color: C.textFaint, fontSize: 12 }}>None yet</span>}
                      </div>
                      <div style={{ color: C.textFaint, fontSize: 11, marginBottom: 8 }}>QUALIFYING TRANSACTIONS ({r.qualifyingTransactions.length})</div>
                      <div style={{ overflowX: "auto" }}>
                        <table style={{ width: "100%", fontSize: 12, borderCollapse: "collapse" }}>
                          <thead><tr>{["Type", "Sender", "Receiver", "Amount"].map(h => <th key={h} style={{ textAlign: "left", color: C.textFaint, padding: "4px 8px", borderBottom: `1px solid ${C.border}` }}>{h}</th>)}</tr></thead>
                          <tbody>
                            {r.qualifyingTransactions.slice(0, 15).map((p, i) => (
                              <tr key={i}><td style={{ padding: "4px 8px", color: C.textMuted }}>{p.txn_type}</td><td style={{ padding: "4px 8px" }}>{p.sender}</td><td style={{ padding: "4px 8px" }}>{p.receiver}</td><td style={{ padding: "4px 8px", color: C.gold }}>{p.amount}</td></tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}
                </div>
              ))}
            </div>
          </div>
        )}

        {tab === "internal" && (
          <div>
            <PhaseDateFilterBar phase={internalPhase} setPhase={setInternalPhase} dateFrom={internalDateFrom} setDateFrom={setInternalDateFrom} dateTo={internalDateTo} setDateTo={setInternalDateTo} />
            <div style={{ fontWeight: 700, fontSize: 16, marginBottom: 6 }}>Internal Transfers</div>
            <div style={{ color: C.textFaint, fontSize: 12.5, marginBottom: 16 }}>
              Every transaction touching Exto Backend or a field coordinator (Sandeep, Cindy, Segomotso, Brian) — for visibility only. None of these generate incentive payouts.
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {Object.entries(internalTransfersByCoordinator).map(([coordinator, data]) => (
                <div key={coordinator} style={{ border: `1px solid ${C.border}`, borderRadius: 12, overflow: "hidden" }}>
                  <div onClick={() => setExpandedCoordinator(expandedCoordinator === coordinator ? null : coordinator)} style={{ display: "flex", alignItems: "center", padding: "12px 18px", cursor: "pointer", background: expandedCoordinator === coordinator ? C.gold + "0c" : C.panelDark, gap: 14 }}>
                    <span style={{ flex: 1, fontWeight: 600 }}>{coordinator}</span>
                    <Badge color={C.blue}>Sent: BWP {Math.round(data.totalSent).toLocaleString()}</Badge>
                    <Badge color={C.green}>Received: BWP {Math.round(data.totalReceived).toLocaleString()}</Badge>
                    <span style={{ color: C.textFaint }}>{expandedCoordinator === coordinator ? "▲" : "▼"}</span>
                  </div>
                  {expandedCoordinator === coordinator && (
                    <div style={{ background: C.bgDark, padding: 16, borderTop: `1px solid ${C.border}` }}>
                      {Object.entries(data.counterparties).map(([counterparty, cp]) => (
                        <div key={counterparty} style={{ marginBottom: 10, border: `1px solid ${C.border}`, borderRadius: 8, overflow: "hidden" }}>
                          <div onClick={() => setExpandedCounterparty(expandedCounterparty === `${coordinator}|${counterparty}` ? null : `${coordinator}|${counterparty}`)} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 14px", cursor: "pointer", background: C.panelDark }}>
                            <span style={{ flex: 1, fontSize: 13 }}>{counterparty || "(unknown)"}</span>
                            {cp.sent > 0 && <Badge color={C.blue}>Sent BWP {Math.round(cp.sent).toLocaleString()}</Badge>}
                            {cp.received > 0 && <Badge color={C.green}>Received BWP {Math.round(cp.received).toLocaleString()}</Badge>}
                            <span style={{ color: C.textFaint, fontSize: 11 }}>{expandedCounterparty === `${coordinator}|${counterparty}` ? "▲" : "▼"}</span>
                          </div>
                          {expandedCounterparty === `${coordinator}|${counterparty}` && (
                            <table style={{ width: "100%", fontSize: 12, borderCollapse: "collapse" }}>
                              <thead><tr>{["Date", "Type", "Direction", "Amount"].map(h => <th key={h} style={{ textAlign: "left", color: C.textFaint, padding: "4px 10px", borderBottom: `1px solid ${C.border}` }}>{h}</th>)}</tr></thead>
                              <tbody>
                                {cp.txns.map((t, i) => (
                                  <tr key={i}><td style={{ padding: "4px 10px", color: C.textMuted }}>{t.dlt_close?.split(",")[0]}</td><td style={{ padding: "4px 10px", color: C.textMuted }}>{t.txn_type}</td><td style={{ padding: "4px 10px" }}>{t.sender === coordinator ? "Sent" : "Received"}</td><td style={{ padding: "4px 10px", color: C.gold }}>{t.amount}</td></tr>
                                ))}
                              </tbody>
                            </table>
                          )}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              ))}
              {Object.keys(internalTransfersByCoordinator).length === 0 && <div style={{ color: C.textFaint, padding: 20 }}>No internal transfers found.</div>}
            </div>
          </div>
        )}

        {tab === "audit" && (
          <div>
            <PhaseDateFilterBar phase={auditPhase} setPhase={setAuditPhase} dateFrom={auditDateFrom} setDateFrom={setAuditDateFrom} dateTo={auditDateTo} setDateTo={setAuditDateTo} />
            <div style={{ fontWeight: 700, fontSize: 16, marginBottom: 6 }}>Transaction Audit</div>
            <div style={{ color: C.textFaint, fontSize: 12.5, marginBottom: 16 }}>
              Every payment transaction in this tab's phase/date range, showing whether it counts toward incentive and why.
            </div>
            <div style={{ display: "flex", gap: 10, marginBottom: 14, alignItems: "center" }}>
              {[
                { id: "all", label: `All (${auditedTransactions.length})` },
                { id: "counted", label: `✅ Counted (${auditedTransactions.filter(t => t.counted).length})` },
                { id: "excluded", label: `❌ Excluded (${auditedTransactions.filter(t => !t.counted).length})` },
              ].map(f => (
                <button key={f.id} onClick={() => setAuditFilter(f.id)} style={{ background: auditFilter === f.id ? C.gold : "transparent", color: auditFilter === f.id ? "#1a1a1a" : C.textMuted, border: `1px solid ${auditFilter === f.id ? C.gold : C.border}`, borderRadius: 7, padding: "6px 14px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>{f.label}</button>
              ))}
            </div>
            <div style={{ overflowX: "auto", background: C.panelDark, border: `1px solid ${C.border}`, borderRadius: 12 }}>
              <table style={{ width: "100%", fontSize: 12.5, borderCollapse: "collapse" }}>
                <thead>
                  <tr style={{ background: C.panelDark2 }}>
                    {["", "Type", "Sender", "Receiver", "Amount", "Date", "Reason"].map(h => (
                      <th key={h} style={{ textAlign: "left", color: C.textFaint, padding: "10px 12px", borderBottom: `1px solid ${C.border}`, fontWeight: 600, whiteSpace: "nowrap" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {auditFiltered.map((t, i) => (
                    <tr key={i} style={{ borderBottom: `1px solid ${C.border}` }}>
                      <td style={{ padding: "9px 12px", fontSize: 15 }}>{t.counted ? "✅" : "❌"}</td>
                      <td style={{ padding: "9px 12px", color: C.cream, whiteSpace: "nowrap" }}>{t.txn_type}</td>
                      <td style={{ padding: "9px 12px", color: C.textMuted }}>{t.sender || "—"}</td>
                      <td style={{ padding: "9px 12px", color: C.textMuted }}>{t.receiver || "—"}</td>
                      <td style={{ padding: "9px 12px", color: C.gold, whiteSpace: "nowrap" }}>BWP {t.amount}</td>
                      <td style={{ padding: "9px 12px", color: C.textFaint, whiteSpace: "nowrap" }}>{t.dlt_close?.split(",")[0]}</td>
                      <td style={{ padding: "9px 12px", color: t.counted ? C.green : C.textFaint }}>{t.reason}</td>
                    </tr>
                  ))}
                  {auditFiltered.length === 0 && (
                    <tr><td colSpan={7} style={{ padding: 20, color: C.textFaint, textAlign: "center" }}>No transactions in this view.</td></tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ─── LIVE DATA ORCHESTRATOR ──────────────────────────────────────────────────
function LiveDataApp() {
  const isAuthenticated = useIsAuthenticated();
  const [useSample, setUseSample] = useState(false);
  const [onboardingEvents, setOnboardingEvents] = useState(FALLBACK_ONBOARDING_RAW);
  const [payments, setPayments] = useState(FALLBACK_PAYMENTS);
  const [dataSource, setDataSource] = useState("sample");
  const [lastUpdated, setLastUpdated] = useState(null);
  const [refreshing, setRefreshing] = useState(false);
  const intervalRef = useRef(null);

  const loadLiveData = useCallback(async () => {
    setRefreshing(true);
    try {
      const result = await fetchLivePilotData();
      setOnboardingEvents(result.onboardingEvents);
      setPayments(result.payments);
      setDataSource("live");
      setLastUpdated(new Date(result.fetchedAt).toLocaleTimeString());
    } catch (err) {
      console.error("Live data fetch failed:", err);
    } finally {
      setRefreshing(false);
    }
  }, []);

  useEffect(() => {
    if (!isAuthenticated) return;
    // Wait a moment after authentication settles before the first fetch —
    // avoids racing MSAL's own popup-close/cache-write right after sign-in.
    const initialLoad = setTimeout(loadLiveData, 800);
    intervalRef.current = setInterval(loadLiveData, AUTO_REFRESH_MS);
    return () => {
      clearTimeout(initialLoad);
      clearInterval(intervalRef.current);
    };
  }, [isAuthenticated, loadLiveData]);

  if (!isAuthenticated && !useSample) {
    return <MicrosoftSignIn onUseSampleData={() => setUseSample(true)} />;
  }

  return (
    <Dashboard onboardingEvents={onboardingEvents} payments={payments} dataSource={dataSource} lastUpdated={lastUpdated} onRefresh={loadLiveData} refreshing={refreshing} />
  );
}

// ─── APP ROOT ────────────────────────────────────────────────────────────────
export default function App() {
  const [loggedIn, setLoggedIn] = useState(false);
  if (!loggedIn) return <LoginScreen onLogin={() => setLoggedIn(true)} />;
  return (
    <MsalProvider instance={msalInstance}>
      <LiveDataApp />
    </MsalProvider>
  );
}
