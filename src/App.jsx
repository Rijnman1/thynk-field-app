import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import RouteTrace from "./RouteTrace.jsx";
import RouteReview from "./RouteReview.jsx";
import {
  Droplet, Camera, MapPin, Check, X, Flag, Edit3, ArrowRight,
  ArrowLeft, ChevronRight, Loader2, CheckCircle2, AlertTriangle,
  Lock, ClipboardList, LayoutGrid, RotateCcw, FileSpreadsheet, Printer,
  History, CloudOff, Cloud, CloudUpload, Send, Radio, Activity, Trash2
} from "lucide-react";

/* ---------- design tokens ---------- */
const C = {
  primary: "#0D86F3",
  primaryDeep: "#0A5FB0",
  charcoal: "#2B2F33",
  charcoalSoft: "#5B6570",
  paper: "#F4F7F9",
  paperDeep: "#E8EDF1",
  line: "#DCE3E8",
  approve: "#1B9C6E",
  approveSoft: "#E4F5EE",
  review: "#D98A22",
  reviewSoft: "#FBF0DE",
  flag: "#D6485A",
  flagSoft: "#FBE6E9",
};

const FONTS = `
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@500;600&display=swap');
`;

/* ---------- helpers ---------- */

/* Real device GPS — captured at the moment a photo is taken/uploaded. Resolves to
   {lat, lng} strings, or null if location isn't available/permitted (caller falls back to "Unknown"). */
function getRealGPS(timeoutMs = 8000) {
  return new Promise((resolve) => {
    if (!navigator.geolocation) { resolve(null); return; }
    let done = false;
    const finish = (val) => { if (!done) { done = true; resolve(val); } };
    navigator.geolocation.getCurrentPosition(
      (pos) => finish({ lat: pos.coords.latitude.toFixed(6), lng: pos.coords.longitude.toFixed(6) }),
      () => finish(null),
      { enableHighAccuracy: true, timeout: timeoutMs, maximumAge: 0 }
    );
    setTimeout(() => finish(null), timeoutMs + 500);
  });
}

const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const nowStamp = () => {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  let h = d.getHours();
  const ampm = h >= 12 ? "PM" : "AM";
  h = h % 12 || 12;
  return {
    date: `${pad(d.getDate())} ${MONTHS[d.getMonth()]} ${d.getFullYear()}`,
    time: `${h}:${pad(d.getMinutes())} ${ampm}`,
  };
};

/* ---------- persistence (shared via your Cloudflare worker, so field phones and the office see the same surveys) ---------- */
const WORKER_URL = "https://thynk-vision.geyserr.workers.dev";
const SURVEY_PREFIX = "survey:";
const genSurveyId = () => `${SURVEY_PREFIX}${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;

/* Team access key — entered once, then remembered on this device for a working day
   so a download, a reload or the app being backgrounded doesn't force a fresh login. */
let APP_KEY = "";
const SESSION_STORE = "thynk_session";
const SESSION_HOURS = 12;

const setAppKey = (k) => { APP_KEY = k; };
const keyHeaders = (extra = {}) => ({ "X-App-Key": APP_KEY, ...extra });

function saveSession(key, role, username) {
  try {
    localStorage.setItem(SESSION_STORE, JSON.stringify({
      key, role, username, expires: Date.now() + SESSION_HOURS * 60 * 60 * 1000,
    }));
  } catch { /* storage unavailable — session simply won't persist */ }
}

function loadSession() {
  try {
    const raw = localStorage.getItem(SESSION_STORE);
    if (!raw) return null;
    const s = JSON.parse(raw);
    if (!s.key || !s.role || !s.expires || Date.now() > s.expires) {
      localStorage.removeItem(SESSION_STORE);
      return null;
    }
    return s;
  } catch {
    return null;
  }
}

function clearSession() {
  APP_KEY = "";
  try { localStorage.removeItem(SESSION_STORE); } catch { /* ignore */ }
}

/* Returns {role, username} if the key is valid, or null. */
async function verifyAppKey(candidate) {
  try {
    const res = await fetch(`${WORKER_URL}/whoami`, { headers: { "X-App-Key": candidate } });
    if (!res.ok) return null;
    const data = await res.json();
    return data.role ? { role: data.role, username: data.username || "Unknown" } : null;
  } catch {
    throw new Error("network");
  }
}

async function saveSurveyRecord(id, survey, captures) {
  if (!id) return false;
  const res = await fetch(`${WORKER_URL}/surveys/${encodeURIComponent(id)}`, {
    method: "PUT",
    headers: keyHeaders({ "Content-Type": "application/json" }),
    body: JSON.stringify({ survey, captures, updatedAt: Date.now() }),
  });
  if (!res.ok) throw new Error(`Save failed (${res.status})`);
  return true;
}

/* Returns lightweight summaries only (no photos) so the list loads fast. */
async function listSurveyRecords() {
  const res = await fetch(`${WORKER_URL}/surveys`, { headers: keyHeaders() });
  if (!res.ok) return [];
  const data = await res.json();
  return data.surveys || [];
}

/* Fetches one full survey (all captures and photos) for resuming/reviewing. */
async function fetchSurveyRecord(key) {
  const res = await fetch(`${WORKER_URL}/surveys/${encodeURIComponent(key)}`, { headers: keyHeaders() });
  if (!res.ok) throw new Error(`Load failed (${res.status})`);
  return await res.json();
}

async function deleteSurveyRecord(key) {
  const res = await fetch(`${WORKER_URL}/surveys/${encodeURIComponent(key)}`, { method: "DELETE", headers: keyHeaders() });
  return res.ok;
}


async function loadDownscaledPhoto(file, maxW = 640) {
  const dataUrl = await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
  const img = await new Promise((resolve, reject) => {
    const el = new Image();
    el.onload = () => resolve(el);
    el.onerror = reject;
    el.src = dataUrl;
  });
  const scale = Math.min(1, maxW / img.width);
  const canvas = document.createElement("canvas");
  canvas.width = Math.round(img.width * scale);
  canvas.height = Math.round(img.height * scale);
  const ctx = canvas.getContext("2d");
  ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
  return canvas.toDataURL("image/jpeg", 0.85);
}

/* ---------- AI vision extraction (real Claude API call, used for sensor screenshots) ---------- */
/* ---------- AI vision extraction (via your Cloudflare worker) ---------- */
async function extractWithVision(photoDataUrl, instructions) {
  const base64 = photoDataUrl.split(",")[1];
  const response = await fetch(`${WORKER_URL}/ai`, {
    method: "POST",
    headers: keyHeaders({ "Content-Type": "application/json" }),
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 500,
      messages: [{
        role: "user",
        content: [
          { type: "image", source: { type: "base64", media_type: "image/jpeg", data: base64 } },
          { type: "text", text: instructions },
        ],
      }],
    }),
  });
  if (!response.ok) throw new Error(`Vision request failed (${response.status})`);
  const data = await response.json();
  const textBlock = (data.content || []).find((b) => b.type === "text");
  if (!textBlock) throw new Error("No response from vision model");
  const cleaned = textBlock.text.replace(/```json|```/g, "").trim();
  return JSON.parse(cleaned);
}

const METER_PROMPT = `This is a photograph of a water meter taken in the field. Read the meter register carefully. Respond with ONLY a JSON object, no other text, in this exact shape:
{"reading": string|null, "serial": string|null, "meterType": string|null, "confidence": number}
Rules:
- "reading": the meter reading as displayed, in m³ (cubic meters). Include decimals if red/decimal dials or digits are visible. Use a plain number string like "104.21". If the register is not clearly legible, use null.
- "serial": the meter's serial number if visible on the meter body or register plate, else null.
- "meterType": the manufacturer/model if identifiable (e.g. "Kamstrup flowIQ 2200", "Elster Kent"), else null.
- "confidence": your honest 0-100 confidence in the reading value specifically. If reading is null, use 0.
- Do not guess digits you cannot see. A wrong reading is worse than no reading.`;

const SENSOR_PROMPT = `This is a screenshot from a water/sensor logger app showing a deployed sensor's status. Extract only what is clearly legible. Respond with ONLY a JSON object, no other text, in this exact shape:
{"sessionId": string|null, "deviceId": string|null, "firmwareVersion": string|null, "signalStrength": string|null, "batteryVoltage": string|null}
Use null for any field that isn't clearly visible. Do not guess or estimate values.`;

const CONSUMPTION_PROMPT = `This is a screenshot of a water consumption/flow profile graph from a sensor app. Extract only what is clearly legible or directly readable from labels/axes. Respond with ONLY a JSON object, no other text, in this exact shape:
{"dateRange": string|null, "minFlow": string|null, "maxFlow": string|null, "avgFlow": string|null, "notes": string|null}
"notes" should be a short (max 20 words) factual observation about the visible trend (e.g. "Continuous minimum night flow suggests possible leak"), only if clearly supported by the graph. Use null for anything not clearly visible. Do not guess or estimate exact values.`;

const FIDO_SESSION_TYPES = [
  "Bug Acoustic Session",
  "Bug Top Sounding",
  "Bug Consumption Profile",
  "FIDO Correlation Results",
];

/* ---------- FIDO 2 ---------- */

/* The sensor register. Add new serials here as the fleet grows. */
const BUG_REGISTER = [
  "B01:41805", "B01:41810", "B01:41811", "B01:41847",
  "B01:42253", "B01:42254", "B01:42258", "B01:42259",
  "B01:42260", "B01:42263", "B01:42264", "B01:42265",
  "B01:42269", "B01:42270", "B01:42274", "B01:42275",
];

/* What a bug can physically be deployed on. Coupling quality affects the reading —
   direct metal contact transmits best, a chamber wall worst. */
const FIDO_DEPLOY_ASSETS = [
  { v: "PIPE (DIRECT)", coupling: 2, note: "Best coupling \u2014 straight onto the pipe wall" },
  { v: "VALVE", coupling: 2, note: "Direct metal path to the pipe" },
  { v: "FIRE HYDRANT", coupling: 2, note: "Direct metal path to the pipe" },
  { v: "WATER METER", coupling: 1, note: "Good, but the meter body can damp the signal" },
  { v: "INDIVIDUAL UNIT METER", coupling: 1, note: "Good on the service, limited reach to the main" },
  { v: "PUMP", coupling: 0, note: "Pump noise can mask a leak" },
  { v: "OTHER", coupling: 0, note: "" },
];

/* Session modes offered by the FIDO app at deployment. */
const FIDO_MODES = [
  "START SESSION",
  "SOUNDING LITE",
  "CONSUMPTION PROFILING",
];

/* Pipe material acoustics. Figures are from published leak-detection research and are
   material characteristics, not device thresholds — the dB level that constitutes a leak
   depends on the logger's own scale, mount, pressure, diameter and soil. */
const PIPE_MATERIALS = {
  POLYETHYLENE: {
    label: "Polyethylene / HDPE",
    freq: "100 \u2013 400 Hz",
    normal: "250 Hz",
    attenuation: "~5.5 dB/m",
    spacing: "Up to ~16 m",
    survey: "Point to point. Sound every fitting \u2014 signal fades fast.",
    tone: "#D6485A",
  },
  PVC: {
    label: "PVC / uPVC",
    freq: "200 \u2013 500 Hz",
    normal: "300 Hz",
    attenuation: "~5.5 dB/m",
    spacing: "Up to ~16 m",
    survey: "Point to point. Every appurtenance should be sounded.",
    tone: "#D6485A",
  },
  LEAD: {
    label: "Lead",
    freq: "200 \u2013 700 Hz",
    normal: "400 Hz",
    attenuation: "moderate",
    spacing: "Short runs \u2014 usually services",
    survey: "Typically found on old service connections.",
    tone: "#D98A22",
  },
  AC: {
    label: "AC",
    freq: "300 \u2013 800 Hz",
    normal: "500 Hz",
    attenuation: "between metal and plastic",
    spacing: "Up to ~60 m",
    survey: "Sound travels reasonably at good pressure. Max 60\u201390 m.",
    tone: "#D98A22",
  },
  IRON: {
    label: "Iron",
    freq: "300 \u2013 1200 Hz",
    normal: "700 Hz",
    attenuation: "~0.9\u20132.4 dB/m",
    spacing: "Up to ~38\u201399 m",
    survey: "Sound valves and hydrants. Shorten spacing below 40 psi.",
    tone: "#1B9C6E",
  },
  STEEL: {
    label: "Steel",
    freq: "400 \u2013 1500 Hz",
    normal: "800 Hz",
    attenuation: "~0.9 dB/m",
    spacing: "Up to ~99 m",
    survey: "Good conductor. Pinhole leaks can be very loud.",
    tone: "#1B9C6E",
  },
  COPPER: {
    label: "Copper",
    freq: "700 \u2013 2500 Hz",
    normal: "1800 Hz",
    attenuation: "~0.9 dB/m",
    spacing: "Up to ~99 m",
    survey: "Highest frequency of all. Common on short service runs.",
    tone: "#1B9C6E",
  },
  UNKNOWN: {
    label: "Unknown",
    freq: "\u2014",
    normal: "\u2014",
    attenuation: "\u2014",
    spacing: "Treat as plastic",
    survey: "Assume the worst case \u2014 short spacing, point to point.",
    tone: "#5B6570",
  },
};

/* Conditions that affect whether a leak can be heard at all.
   Source: Gutermann leak detection theory \u2014 all factors at a minimum except pressure at maximum. */
const DEPLOY_CONDITIONS = {
  diameter: {
    label: "PIPE DIAMETER",
    hint: "Small is good",
    options: [
      { v: "15-25mm", good: 2 },
      { v: "25-65mm", good: 1 },
      { v: "65-125mm", good: -1 },
      { v: ">125mm", good: -2 },
    ],
  },
  pressure: {
    label: "PRESSURE",
    hint: "High is good",
    options: [
      { v: "HIGH", good: 2 },
      { v: "NORMAL", good: 1 },
      { v: "LOW", good: -2 },
    ],
  },
  backfill: {
    label: "BACKFILL",
    hint: "Hard is good, soft is poor",
    options: [
      { v: "HARD", good: 2 },
      { v: "MIXED", good: 0 },
      { v: "SOFT / SANDY", good: -2 },
      { v: "N/A \u2014 EXPOSED", good: 0, na: true },
    ],
  },
  pipecondition: {
    label: "PIPE CONDITION",
    hint: "Clean is good, encrusted or lined is poor",
    options: [
      { v: "CLEAN", good: 2 },
      { v: "SOME SCALE", good: 0 },
      { v: "ENCRUSTED / LINED", good: -2 },
    ],
  },
  background: {
    label: "BACKGROUND NOISE",
    hint: "PRVs and throttled valves mask leaks",
    options: [
      { v: "QUIET", good: 2 },
      { v: "SOME", good: 0 },
      { v: "PRV / VALVE NEARBY", good: -2 },
    ],
  },
  consumption: {
    label: "CONSUMPTION",
    hint: "High demand drowns out leak noise",
    options: [
      { v: "OVERNIGHT / LOW", good: 2 },
      { v: "MIXED", good: 0 },
      { v: "DAYTIME / HIGH", good: -2 },
    ],
  },
};

/* Pipe material feeds the same judgement — metallic carries leak noise well, plastic poorly.
   It is already recorded at deployment, so it scores without asking again. */
const MATERIAL_ACOUSTIC_SCORE = {
  COPPER: 2, STEEL: 2, IRON: 2,
  AC: 0, LEAD: 0,
  PVC: -2, POLYETHYLENE: -2,
  UNKNOWN: -1,
};

/* Rates how favourable the deployment conditions were, so an inconclusive
   result can be explained rather than just recorded.
   Source: Gutermann leak detection theory. */
function conditionScore(c) {
  const keys = Object.keys(DEPLOY_CONDITIONS);
  const set = keys.filter((k) => c[`cond_${k}`]);
  const hasMaterial = c.pipeMaterial && MATERIAL_ACOUSTIC_SCORE[c.pipeMaterial] !== undefined;
  if (!set.length && !hasMaterial) return null;

  let score = 0;
  let counted = 0;
  set.forEach((k) => {
    const opt = DEPLOY_CONDITIONS[k].options.find((o) => o.v === c[`cond_${k}`]);
    // A factor marked not applicable is excluded rather than scored as mediocre
    if (opt && !opt.na) { score += opt.good; counted += 1; }
  });
  if (hasMaterial) { score += MATERIAL_ACOUSTIC_SCORE[c.pipeMaterial]; counted += 1; }
  const mount = FIDO_DEPLOY_ASSETS.find((a) => a.v === c.deployAsset);
  if (mount && mount.v !== "OTHER") { score += mount.coupling; counted += 1; }
  if (!counted) return null;

  const factors = counted;
  const pct = Math.round(((score + factors * 2) / (factors * 4)) * 100);

  let label, colour, soft, advice;
  if (pct >= 70) {
    label = "GOOD LISTENING CONDITIONS"; colour = "#1B9C6E"; soft = "#E4F5EE";
    advice = "Conditions favour detection. A quiet result is likely to be a genuine no-leak.";
  } else if (pct >= 40) {
    label = "FAIR LISTENING CONDITIONS"; colour = "#D98A22"; soft = "#FBF0DE";
    advice = "Workable, but a quiet result is less conclusive here.";
  } else {
    label = "POOR LISTENING CONDITIONS"; colour = "#D6485A"; soft = "#FBE6E9";
    advice = "A leak may be masked. Consider redeploying overnight or at a quieter point.";
  }
  return { pct, label, colour, soft, advice, recorded: factors, total: keys.length + 1 };
}

/* Investigation methods used within a waypoint. */
const WAYPOINT_TESTS = {
  CORRELATION: { label: "CORRELATION", needsDistance: true, icon: "Activity" },
  "TOP SOUNDING": { label: "TOP SOUNDING", needsDistance: false },
  "SOUNDING LITE": { label: "SOUNDING LITE", needsDistance: false },
  "GROUND MIC": { label: "GROUND MIC", needsDistance: false },
  HYDROPHONE: { label: "HYDROPHONE", needsDistance: false },
  OTHER: { label: "OTHER", needsDistance: false },
};

/* How an investigation ended. */
const WAYPOINT_OUTCOMES = {
  "LEAK CONFIRMED": { colour: "#D6485A", soft: "#FBE6E9" },
  "LEAK SUSPECTED": { colour: "#D98A22", soft: "#FBF0DE" },
  "NO LEAK FOUND": { colour: "#1B9C6E", soft: "#E4F5EE" },
  "ONGOING": { colour: "#0D86F3", soft: "#E7F2FE" },
};

const CORRELATION_PROMPT = `This is a screenshot from a leak noise correlator. Extract only what is clearly legible. Respond with ONLY a JSON object, no other text, in this exact shape:
{"distance": string|null, "sensorSpacing": string|null, "confidence": string|null, "notes": string|null}
Rules:
- "distance": the located leak distance from a sensor, as displayed, including units.
- "sensorSpacing": the total pipe length between sensors if shown.
- "confidence": any correlation quality, coherence or confidence figure shown.
- "notes": a short factual description of the correlation peak, maximum 15 words.
Use null for anything not clearly visible. Do not guess values.`;

/* Waypoints are numbered per survey so a client report can refer to them plainly. */
function nextWaypointRef(captures) {
  const used = captures
    .filter((c) => c.type === "fido2_waypoint")
    .map((c) => parseInt(String(c.waypointRef || "").replace(/\D/g, ""), 10))
    .filter((n) => !isNaN(n));
  const next = used.length ? Math.max(...used) + 1 : 1;
  return `WP-${String(next).padStart(3, "0")}`;
}

/* Findings recorded at retrieval. */
const FIDO_OUTCOMES = {
  "NO LEAK": { colour: "#1B9C6E", soft: "#E4F5EE" },
  "SUSPECTED LEAK": { colour: "#D98A22", soft: "#FBF0DE" },
  "CONFIRMED LEAK": { colour: "#D6485A", soft: "#FBE6E9" },
  "INCONCLUSIVE": { colour: "#5B6570", soft: "#E8EDF1" },
};

const FIDO_DEPLOY_PROMPT = `This is a screenshot from the FIDO leak detection app showing a sensor session that has just been started. Extract only what is clearly legible. Respond with ONLY a JSON object, no other text, in this exact shape:
{"sessionId": string|null, "bugSerial": string|null, "signalStrength": string|null, "batteryVoltage": string|null, "startedAt": string|null}
Rules:
- "sessionId": the session identifier shown (often a short code such as "7kvd").
- "bugSerial": the sensor serial, usually in the form B01:XXXXX.
- "signalStrength" and "batteryVoltage": as displayed, including units.
- "startedAt": any date or time shown for the session start.
Use null for anything not clearly visible. Do not guess.`;

const FIDO_RESULT_PROMPT = `This is a screenshot of a FIDO leak detection results graph from a completed sensor session. Extract only what is clearly legible. Respond with ONLY a JSON object, no other text, in this exact shape:
{"sessionId": string|null, "bugSerial": string|null, "dateRange": string|null, "minLevel": string|null, "notes": string|null}
Rules:
- "sessionId": the session identifier shown on the graph or its heading.
- "bugSerial": the sensor serial if shown, usually B01:XXXXX.
- "dateRange": the period the graph covers.
- "minLevel": any minimum night level or lowest value labelled on the graph.
- "notes": a short factual description of the visible trend, maximum 20 words. Describe only what the graph shows. Do not conclude whether there is a leak.
Use null for anything not clearly visible. Do not guess values.`;

/* Builds the current state of every sensor from the captures in a survey. */
function buildBugStatus(captures) {
  const status = {};
  BUG_REGISTER.forEach((serial) => {
    status[serial] = { serial, state: "AVAILABLE", deployment: null, retrieval: null };
  });
  // deployments first, most recent last so the latest wins
  const sorted = [...captures].sort((a, b) => (a.id || 0) - (b.id || 0));
  sorted.forEach((c) => {
    if (c.type === "fido2_deploy" && c.bugSerial) {
      if (!status[c.bugSerial]) status[c.bugSerial] = { serial: c.bugSerial, state: "AVAILABLE", deployment: null, retrieval: null };
      status[c.bugSerial].state = "DEPLOYED";
      status[c.bugSerial].deployment = c;
      status[c.bugSerial].retrieval = null;
    }
  });
  sorted.forEach((c) => {
    if (c.type === "fido2_retrieve" && c.bugSerial && status[c.bugSerial]) {
      status[c.bugSerial].state = "RETURNED";
      status[c.bugSerial].retrieval = c;
    }
  });
  return status;
}

const daysSince = (id) => {
  if (!id) return 0;
  return Math.max(0, Math.floor((Date.now() - id) / (1000 * 60 * 60 * 24)));
};

/* ---------- field task definitions (the hub) ---------- */
const TASKS = {
  meterwork: {
    label: "Meter Work",
    blurb: "Read existing meters, or replace and install new ones.",
    icon: Camera,
  },
  fido: {
    label: "FIDO Leak Analysis",
    blurb: "Ultrasonic AI leak detection — sensor deployments and FIDO session feedback.",
    icon: Radio,
  },
  fido2: {
    label: "FIDO 2",
    blurb: "Deploy and retrieve leak sensors, with outcomes linked by bug serial.",
    icon: Radio,
  },
  amr: {
    label: "AMR Survey",
    blurb: "Map MUCs and repeaters, and record which meters each one can see.",
    icon: Radio,
  },
  routetrace: {
    label: "Route Trace",
    blurb: "Walk and map buried services — pipes, cables, valves and fittings as positioned routes.",
    icon: MapPin,
  },
  assets: {
    label: "Asset Mapping",
    blurb: "Map estate infrastructure — water, electrical, AMR and general assets with coordinates.",
    icon: LayoutGrid,
  },
};

const METER_MODES = {
  read: {
    label: "READ METERS",
    blurb: "Photograph each meter and capture its reading and serial.",
    icon: Camera,
  },
  replace: {
    label: "REPLACE METERS",
    blurb: "Photograph the old meter out and the new meter in at each position.",
    icon: RotateCcw,
  },
};

/* What kind of meter this is — recorded at every reading */
const METER_TYPES = [
  "WATER METER",
  "BULK METER",
  "ZONE METER",
  "SECTIONAL TITLE BULK",
  "COMMON PROPERTY",
];

/* Why a meter could not be read — drives the revisit list */
const EXCEPTION_REASONS = [
  "NO ACCESS — GATE LOCKED",
  "NO ACCESS — DOG",
  "METER BURIED / COVERED",
  "METER CHAMBER FLOODED",
  "GLASS FOGGED / UNREADABLE",
  "METER MISSING",
  "OCCUPANT REFUSED",
  "OTHER",
];

/* Condition problems worth flagging to the estate */
const CONDITION_FLAGS = [
  "DAMAGED",
  "LEAKING",
  "NEEDS REPAIR",
  "CHAMBER DAMAGED",
  "LID MISSING",
  "HARD TO ACCESS",
];

/* Parses an imported previous-readings CSV into a lookup keyed by position and serial. */
function parsePreviousReadings(text) {
  const lines = text.split(/\r?\n/).filter((l) => l.trim().length);
  if (lines.length < 2) return {};
  const splitRow = (line) => {
    const out = [];
    let cur = "";
    let inQ = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        if (inQ && line[i + 1] === '"') { cur += '"'; i++; }
        else inQ = !inQ;
      } else if (ch === "," && !inQ) { out.push(cur); cur = ""; }
      else cur += ch;
    }
    out.push(cur);
    return out.map((s) => s.trim());
  };
  const headers = splitRow(lines[0]).map((h) => h.toLowerCase().replace(/[^a-z0-9]/g, ""));
  const findCol = (...names) => {
    for (const n of names) {
      const i = headers.indexOf(n);
      if (i !== -1) return i;
    }
    return -1;
  };
  const posCol = findCol("position", "positionname", "unit", "unitnumber", "erf", "address", "assetid");
  const readCol = findCol("reading", "previousreading", "lastreading", "meterreading", "meterreadingm");
  const serialCol = findCol("serial", "serialnumber", "meterserial");
  const dateCol = findCol("date", "readingdate", "lastreaddate", "previousdate");
  const typeCol = findCol("type", "metertype");
  if (posCol === -1 && serialCol === -1) return {};

  const map = {};
  for (let i = 1; i < lines.length; i++) {
    const cols = splitRow(lines[i]);
    const entry = {
      reading: readCol !== -1 ? cols[readCol] || "" : "",
      date: dateCol !== -1 ? cols[dateCol] || "" : "",
      serial: serialCol !== -1 ? cols[serialCol] || "" : "",
      meterType: typeCol !== -1 ? (cols[typeCol] || "").toUpperCase() : "",
    };
    if (!entry.reading) continue;
    if (posCol !== -1 && cols[posCol]) map[`P:${cols[posCol].trim().toUpperCase()}`] = entry;
    if (serialCol !== -1 && cols[serialCol]) map[`S:${cols[serialCol].trim().toUpperCase()}`] = entry;
  }
  return map;
}

/* Parses an AMR register CSV. Two shapes are accepted:
   - an asset register: one row per MUC or repeater, with its location and coordinates
   - a meter allocation: one row per meter, with a column naming its parent asset */
function parseAmrRegister(text) {
  const clean = text.replace(/^\uFEFF/, "");
  const lines = clean.split(/\r?\n/).filter((l) => l.trim().length);
  if (lines.length < 2) return null;

  const splitRow = (line) => {
    const out = []; let cur = ""; let inQ = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        if (inQ && line[i + 1] === '"') { cur += '"'; i++; } else inQ = !inQ;
      } else if (ch === "," && !inQ) { out.push(cur); cur = ""; }
      else cur += ch;
    }
    out.push(cur);
    return out.map((s) => s.trim());
  };

  const headers = splitRow(lines[0]).map((h) => h.toLowerCase().replace(/[^a-z0-9]/g, ""));
  const findCol = (...names) => {
    for (const n of names) { const i = headers.indexOf(n); if (i !== -1) return i; }
    return -1;
  };

  const assetCol = findCol("repeaterserialno", "assetserial", "mucserial", "muc", "repeater",
    "concentrator", "gateway", "deviceserial", "device", "serial", "serialnumber", "parent");
  const meterCol = findCol("meterserial", "meter");
  const descCol = findCol("description", "location", "name", "position", "zone", "street");
  const typeCol = findCol("assettype", "type");
  const latCol = findCol("latitude", "lat");
  const lngCol = findCol("longitude", "lng", "long");
  const notesCol = findCol("accessnotes", "notes", "access", "comment", "comments", "status");

  if (assetCol === -1) return null;

  // Meter allocation: a separate meter column exists
  if (meterCol !== -1 && meterCol !== assetCol) {
    const byAsset = {};
    let count = 0;
    for (let i = 1; i < lines.length; i++) {
      const cols = splitRow(lines[i]);
      const asset = (cols[assetCol] || "").trim().toUpperCase();
      const meter = (cols[meterCol] || "").trim().toUpperCase();
      if (!asset || !meter) continue;
      if (!byAsset[asset]) {
        byAsset[asset] = {
          serial: cols[assetCol].trim(),
          description: descCol !== -1 ? (cols[descCol] || "") : "",
          assetType: typeCol !== -1 ? (cols[typeCol] || "").toUpperCase() : "",
          gps: null, notes: "", meters: [],
        };
      }
      if (!byAsset[asset].meters.includes(meter)) { byAsset[asset].meters.push(meter); count += 1; }
    }
    if (!count) return null;
    return { mode: "allocation", byAsset, meterCount: count, assetCount: Object.keys(byAsset).length };
  }

  // Asset register: one row per MUC or repeater
  const byAsset = {};
  for (let i = 1; i < lines.length; i++) {
    const cols = splitRow(lines[i]);
    const serial = (cols[assetCol] || "").trim();
    if (!serial) continue;
    const lat = latCol !== -1 ? (cols[latCol] || "").trim() : "";
    const lng = lngCol !== -1 ? (cols[lngCol] || "").trim() : "";
    byAsset[serial.toUpperCase()] = {
      serial,
      description: descCol !== -1 ? (cols[descCol] || "").trim() : "",
      assetType: typeCol !== -1 ? (cols[typeCol] || "").trim().toUpperCase() : "",
      gps: lat && lng ? { lat, lng } : null,
      notes: notesCol !== -1 ? (cols[notesCol] || "").trim() : "",
      meters: [],
    };
  }
  const assetCount = Object.keys(byAsset).length;
  if (!assetCount) return null;
  return { mode: "assets", byAsset, meterCount: 0, assetCount };
}

/* Compares what a survey found on an asset against what the register expects.
   Only meaningful when a meter allocation was imported. */
function compareToRegister(register, assetSerial, foundMeters) {
  if (!register || register.mode !== "allocation" || !assetSerial) return null;
  const key = assetSerial.trim().toUpperCase();
  const entry = register.byAsset[key];
  if (!entry) return { unknownAsset: true, expected: [], missing: [], extra: [], matched: [] };
  const found = foundMeters.map((m) => (m.serial || "").trim().toUpperCase()).filter(Boolean);
  const expected = entry.meters;
  const matched = expected.filter((e) => found.includes(e));
  const missing = expected.filter((e) => !found.includes(e));
  const extra = found.filter((f) => !expected.includes(f));
  return { unknownAsset: false, expected, matched, missing, extra };
}

/* Which registered assets have been surveyed and which are still outstanding. */
function registerProgress(register, captures) {
  if (!register) return null;
  const done = new Set(
    captures.filter((c) => c.type === "amr" && c.amrSerial)
      .map((c) => c.amrSerial.trim().toUpperCase())
  );
  const all = Object.values(register.byAsset);
  return {
    total: all.length,
    done: all.filter((a) => done.has(a.serial.toUpperCase())).length,
    outstanding: all.filter((a) => !done.has(a.serial.toUpperCase())),
    isDone: (serial) => done.has(String(serial || "").trim().toUpperCase()),
  };
}

/* Finds the previous reading for a capture, by position first then serial. */
function lookupPrevious(previous, position, serial) {
  if (!previous) return null;
  const byPos = position ? previous[`P:${position.trim().toUpperCase()}`] : null;
  if (byPos) return byPos;
  const bySerial = serial ? previous[`S:${serial.trim().toUpperCase()}`] : null;
  return bySerial || null;
}

/* Signal quality bands for AMR meter readings.
   Stronger than -85 dBm = GOOD, -85 to -95 = FAIR, weaker than -95 = WEAK. */
function signalBand(signal) {
  if (!signal) return null;
  const m = String(signal).match(/-?\d+/);
  if (!m) return null;
  const dbm = parseInt(m[0], 10);
  if (isNaN(dbm)) return null;
  if (dbm > -85) return { label: "GOOD", colour: "#1B9C6E", soft: "#E4F5EE", dbm };
  if (dbm >= -95) return { label: "FAIR", colour: "#D98A22", soft: "#FBF0DE", dbm };
  return { label: "WEAK", colour: "#D6485A", soft: "#FBE6E9", dbm };
}

/* Finds meters seen by more than one AMR asset across a survey.
   Returns a map of meter serial -> { sightings, best } so the stronger owner can be identified. */
function findDuplicateMeters(captures) {
  const seen = {};
  captures.filter((c) => c.type === "amr").forEach((c) => {
    (c.amrShots || []).forEach((shot) => {
      (shot.meters || []).forEach((m) => {
        const serial = (m.serial || "").trim().toUpperCase();
        if (!serial) return;
        const band = signalBand(m.signal);
        (seen[serial] = seen[serial] || []).push({
          captureId: c.id,
          position: c.position,
          assetType: c.amrAssetType || "AMR",
          assetSerial: c.amrSerial || "",
          signal: m.signal || "",
          dbm: band ? band.dbm : null,
        });
      });
    });
  });
  const dupes = {};
  Object.entries(seen).forEach(([serial, sightings]) => {
    // only a duplicate if seen by two different assets, not twice in one asset's screenshots
    const distinct = new Set(sightings.map((s) => s.captureId));
    if (distinct.size < 2) return;
    const withSignal = sightings.filter((s) => s.dbm !== null);
    const best = withSignal.length
      ? withSignal.reduce((a, b) => (b.dbm > a.dbm ? b : a))
      : null;
    dupes[serial] = { sightings, best };
  });
  return dupes;
}

/* Asset taxonomy — matches the estate asset intake template exactly */
const ASSET_CATEGORIES = {
  WATER: [
    "BULK WATER METER",
    "ZONE METER",
    "SECTIONAL TITLE BULK METER",
    "WATER METER",
    "COMMON PROPERTY METER",
    "SECTIONAL TITLE INTERNAL METER",
    "ISOLATION VALVE",
    "PRESSURE CONTROL VALVE",
    "BOREHOLE",
    "WATER TANK",
  ],
  ELECTRICAL: [
    "ELECTRICAL METER SINGLE PHASE",
    "ELECTRICAL METER THREE PHASE",
    "CT METER",
    "SECTIONAL TITLE BULK ELECTRICAL METER",
    "ELECTRICAL KIOSK",
    "DB BOARD",
    "POLE",
    "MINI SUBSTATION",
  ],
  "AMR EQUIPMENT": [
    "MUC",
    "REPEATER",
    "CONCENTRATOR",
    "GATEWAY",
    "KAMSTRUP GATEWAY",
    "MINI KAMSTRUP GATEWAY",
    "LORA GATEWAY",
    "4G BRIDGE",
  ],
  FILTRATION: [
    "FILTRATION SYSTEM",
    "CARTRIDGE FILTER HOUSING",
    "SAND FILTER",
    "UV STERILISER",
    "WATER SOFTENER",
    "REVERSE OSMOSIS UNIT",
    "DOSING UNIT / CHLORINATOR",
  ],
  FIRE: [
    "FIRE EXTINGUISHER",
    "FIRE HOSE REEL",
    "FIRE HYDRANT",
    "FIRE WATER TANK",
    "FIRE BOOSTER PUMP",
    "SPRINKLER CONTROL VALVE",
    "FIRE ALARM PANEL",
    "FIRE BLANKET",
  ],
  GENERAL: [
    "MANHOLE",
    "DRAIN",
    "FIBRE CHAMBER",
    "STREET LIGHT",
    "IRRIGATION POINT",
    "SIGNAGE",
    "OTHER",
  ],
};

/* Types that carry a service record. Fire equipment is the compliance-critical set,
   but pumps, PRVs and boreholes are serviced too. */
const SERVICEABLE_TYPES = new Set([
  "FIRE EXTINGUISHER",
  "FIRE HOSE REEL",
  "FIRE HYDRANT",
  "FIRE WATER TANK",
  "FIRE BOOSTER PUMP",
  "SPRINKLER CONTROL VALVE",
  "FIRE ALARM PANEL",
  "FIRE BLANKET",
  "PRESSURE CONTROL VALVE",
  "BOREHOLE",
  "WATER TANK",
]);

/* Filtration assets carry two independent service cycles rather than one. */
const FILTRATION_TYPES = new Set([
  "FILTRATION SYSTEM",
  "CARTRIDGE FILTER HOUSING",
  "SAND FILTER",
  "UV STERILISER",
  "WATER SOFTENER",
  "REVERSE OSMOSIS UNIT",
  "DOSING UNIT / CHLORINATOR",
]);

const CARTRIDGE_INTERVAL_MONTHS = 6;
const FLUSH_INTERVAL_MONTHS = 3;

const CARTRIDGE_CONDITIONS = ["GOOD", "DISCOLOURED", "FOULED", "NEEDS REPLACEMENT"];

/* Months between services, by type. Used to flag an overdue service. */
const SERVICE_INTERVAL_MONTHS = {
  "FIRE EXTINGUISHER": 12,
  "FIRE HOSE REEL": 12,
  "FIRE HYDRANT": 12,
  "FIRE WATER TANK": 12,
  "FIRE BOOSTER PUMP": 12,
  "SPRINKLER CONTROL VALVE": 12,
  "FIRE ALARM PANEL": 12,
  "FIRE BLANKET": 12,
  "PRESSURE CONTROL VALVE": 12,
  "BOREHOLE": 12,
  "WATER TANK": 24,
};

/* Returns { dueDate, overdue, monthsOverdue } for any date and interval, or null. */
function dueStatus(lastISO, months) {
  if (!lastISO) return null;
  const last = new Date(lastISO);
  if (isNaN(last.getTime())) return null;
  const due = new Date(last);
  due.setMonth(due.getMonth() + months);
  const now = new Date();
  const overdue = due < now;
  const monthsOverdue = overdue
    ? Math.max(1, Math.round((now - due) / (1000 * 60 * 60 * 24 * 30.44)))
    : 0;
  const pad = (n) => String(n).padStart(2, "0");
  return {
    dueDate: `${pad(due.getDate())} ${MONTHS[due.getMonth()]} ${due.getFullYear()}`,
    overdue, monthsOverdue, intervalMonths: months,
  };
}

/* Returns { dueDate, overdue, monthsOverdue } or null when no service date is recorded. */
function serviceStatus(assetType, lastServiceISO) {
  if (!lastServiceISO) return null;
  const months = SERVICE_INTERVAL_MONTHS[assetType] || 12;
  const last = new Date(lastServiceISO);
  if (isNaN(last.getTime())) return null;
  const due = new Date(last);
  due.setMonth(due.getMonth() + months);
  const now = new Date();
  const overdue = due < now;
  const monthsOverdue = overdue
    ? Math.max(1, Math.round((now - due) / (1000 * 60 * 60 * 24 * 30.44)))
    : 0;
  const pad = (n) => String(n).padStart(2, "0");
  const dueDate = `${pad(due.getDate())} ${MONTHS[due.getMonth()]} ${due.getFullYear()}`;
  return { dueDate, overdue, monthsOverdue, intervalMonths: months };
}

const CATEGORY_COLOURS = {
  WATER: "#0D86F3",
  ELECTRICAL: "#D98A22",
  "AMR EQUIPMENT": "#7B5BD6",
  FILTRATION: "#0E9AA7",
  FIRE: "#D6485A",
  GENERAL: "#5B6570",
};

const AMR_ASSET_TYPES = ["MUC", "Repeater"];

const AMR_SCREENSHOT_PROMPT = `This is a screenshot from a Kamstrup handheld reader showing a list of water meters detected by a MUC, concentrator, or repeater. Read EVERY meter row visible in the list, top to bottom, including partially visible rows at the edges if legible. Respond with ONLY a JSON object, no other text, in this exact shape:
{"concentratorId": string|null, "readingDate": string|null, "meters": [{"serial": string, "model": string|null, "signal": string|null}]}
Rules:
- "serial": the full meter serial exactly as shown including any prefix (e.g. "KAM 24336621", "KAW 54316243").
- "model": the model name on the line beneath the serial (e.g. "MULTICAL 21", "KWMx230"), else null.
- "signal": the signal strength exactly as shown including units (e.g. "-40 dBm"), else null.
- "concentratorId": any concentrator, gateway, or device ID shown at the top of the screen, else null.
- "readingDate": any date or timestamp visible on the screen, else null.
- Return every row — a typical screen shows 8 to 10 meters. Do not stop early, do not summarise, do not invent rows.`;

/* ---------- meter dial (signature element) ---------- */
function MeterDial({ value, size = 108, stroke = 9, color = C.primary, animate = true }) {
  const [display, setDisplay] = useState(animate ? 0 : value);
  useEffect(() => {
    if (!animate) return;
    let raf;
    const start = performance.now();
    const dur = 900;
    const tick = (t) => {
      const p = Math.min(1, (t - start) / dur);
      setDisplay(Math.round(value * (1 - Math.pow(1 - p, 3))));
      if (p < 1) raf = requestAnimationFrame(tick);
    };
    raf = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(raf);
    // eslint-disable-next-line
  }, [value]);

  const r = (size - stroke) / 2;
  const c = 2 * Math.PI * r;
  const offset = c * (1 - display / 100);
  const center = size / 2;
  const ticks = Array.from({ length: 24 });

  return (
    <div style={{ position: "relative", width: size, height: size }}>
      <svg width={size} height={size} style={{ position: "absolute", inset: 0 }}>
        {ticks.map((_, i) => {
          const angle = (i / 24) * 360;
          const isMajor = i % 6 === 0;
          const len = isMajor ? 6 : 3;
          const r1 = r - stroke / 2 - 2;
          const r2 = r1 - len;
          const rad = (angle * Math.PI) / 180;
          const x1 = center + r1 * Math.cos(rad);
          const y1 = center + r1 * Math.sin(rad);
          const x2 = center + r2 * Math.cos(rad);
          const y2 = center + r2 * Math.sin(rad);
          return (
            <line key={i} x1={x1} y1={y1} x2={x2} y2={y2}
              stroke={isMajor ? C.charcoalSoft : C.line} strokeWidth={isMajor ? 1.4 : 1} />
          );
        })}
        <circle cx={center} cy={center} r={r} fill="none" stroke={C.paperDeep} strokeWidth={stroke} />
        <circle
          cx={center} cy={center} r={r} fill="none" stroke={color} strokeWidth={stroke}
          strokeDasharray={c} strokeDashoffset={offset} strokeLinecap="round"
          transform={`rotate(-90 ${center} ${center})`}
          style={{ transition: "stroke 0.3s" }}
        />
      </svg>
      <div style={{
        position: "absolute", inset: 0, display: "flex", flexDirection: "column",
        alignItems: "center", justifyContent: "center"
      }}>
        <span style={{ fontFamily: "'Space Grotesk',sans-serif", fontWeight: 700, fontSize: size * 0.24, color: C.charcoal }}>
          {display}%
        </span>
        <span style={{ fontFamily: "'Inter',sans-serif", fontSize: size * 0.075, color: C.charcoalSoft, letterSpacing: 0.4 }}>
          COMPLETE
        </span>
      </div>
    </div>
  );
}

function MiniGauge({ value, color }) {
  const size = 34, stroke = 4;
  const r = (size - stroke) / 2, c = 2 * Math.PI * r, center = size / 2;
  const offset = c * (1 - value / 100);
  return (
    <svg width={size} height={size}>
      <circle cx={center} cy={center} r={r} fill="none" stroke={C.paperDeep} strokeWidth={stroke} />
      <circle cx={center} cy={center} r={r} fill="none" stroke={color} strokeWidth={stroke}
        strokeDasharray={c} strokeDashoffset={offset} strokeLinecap="round"
        transform={`rotate(-90 ${center} ${center})`} />
      <text x={center} y={center + 3} textAnchor="middle" fontSize="10" fontWeight="600"
        fontFamily="'IBM Plex Mono',monospace" fill={C.charcoal}>{value}</text>
    </svg>
  );
}

/* Shows what a leak should sound like on a given pipe material. */
function PipeAcousticGuide({ material, compact }) {
  const m = PIPE_MATERIALS[material];
  if (!m) return null;
  return (
    <div style={{
      padding: compact ? "10px 12px" : 12, borderRadius: 9,
      background: `${m.tone}12`, border: `1px solid ${m.tone}55`
    }}>
      <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 9 }}>
        <Activity size={12} color={m.tone} />
        <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, fontWeight: 700, color: m.tone, letterSpacing: 0.3 }}>
          LEAK NOISE ON {m.label.toUpperCase()}
        </span>
      </div>

      <div style={{
        display: "flex", alignItems: "baseline", justifyContent: "center", gap: 8,
        padding: "10px 12px", borderRadius: 8, background: "#fff", marginBottom: 9
      }}>
        <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 9, fontWeight: 700, color: C.charcoalSoft }}>
          EXPECT AROUND
        </span>
        <span style={{ fontFamily: "'Space Grotesk',sans-serif", fontSize: 22, fontWeight: 700, color: m.tone }}>
          {m.normal}
        </span>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
        <div>
          <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 8.5, fontWeight: 600, color: C.charcoalSoft }}>FREQUENCY RANGE</div>
          <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 12, fontWeight: 700, color: C.charcoal }}>{m.freq}</div>
        </div>
        <div>
          <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 8.5, fontWeight: 600, color: C.charcoalSoft }}>MAX SENSOR SPACING</div>
          <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 12, fontWeight: 700, color: C.charcoal }}>{m.spacing}</div>
        </div>
        <div style={{ gridColumn: "1 / -1" }}>
          <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 8.5, fontWeight: 600, color: C.charcoalSoft }}>SIGNAL LOSS</div>
          <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 11, color: C.charcoal }}>{m.attenuation}</div>
        </div>
      </div>

      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, color: C.charcoalSoft, marginTop: 8, lineHeight: 1.4 }}>
        {m.survey}
      </div>
      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 8.5, color: C.charcoalSoft, marginTop: 6, opacity: 0.75 }}>
        Frequency figures: Gutermann leak detection theory
      </div>
    </div>
  );
}

/* ---------- optional evidence slot (sensor deployment / consumption profile) ---------- */
function EvidenceSlot({ title, promptText, value, onChange, fieldsConfig, icon: SlotIcon }) {
  const [status, setStatus] = useState("idle"); // idle | scanning | error
  const [error, setError] = useState("");
  const fileRef = useRef(null);

  const handleFile = async (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!file) return;
    setStatus("scanning");
    setError("");
    try {
      const [photo, gps] = await Promise.all([loadDownscaledPhoto(file, 1000), getRealGPS()]);
      try {
        const extracted = await extractWithVision(photo, promptText);
        onChange({ photo, gps, ...extracted });
      } catch (visionErr) {
        onChange({ photo, gps });
        setError("Photo saved, but AI couldn't read the details — you can enter them manually below.");
      }
      setStatus("idle");
    } catch (err) {
      setError("Couldn't read that photo. Please try again.");
      setStatus("idle");
    }
  };

  const updateField = (key, val) => onChange({ ...value, [key]: val });

  return (
    <div style={{ marginTop: 12, padding: 11, borderRadius: 12, border: `1.5px dashed ${C.line}`, background: C.paper }}>
      <input ref={fileRef} type="file" accept="image/*" onChange={handleFile} style={{ display: "none" }} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <span style={{ display: "flex", alignItems: "center", gap: 6, fontFamily: "'Inter',sans-serif", fontWeight: 600, fontSize: 12.5, color: C.charcoal }}>
          <SlotIcon size={13} color={C.charcoalSoft} /> {title}
        </span>
        <button
          onClick={() => fileRef.current?.click()}
          disabled={status === "scanning"}
          style={{
            border: "none", background: "none", color: C.primary, fontFamily: "'Inter',sans-serif",
            fontWeight: 600, fontSize: 11.5, cursor: "pointer", display: "flex", alignItems: "center", gap: 4
          }}>
          {status === "scanning" ? <Loader2 size={13} className="spin" /> : <Camera size={13} />}
          {status === "scanning" ? "Reading…" : value?.photo ? "Retake" : "Add"}
        </button>
      </div>

      {value?.photo && (
        <div style={{ display: "flex", gap: 8, marginTop: 9 }}>
          <img src={value.photo} alt={title} style={{ width: 58, height: 44, objectFit: "cover", borderRadius: 6, flexShrink: 0 }} />
          <div style={{ flex: 1 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
              <MapPin size={10} color={value.gps ? C.approve : C.charcoalSoft} />
              <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 10, color: value.gps ? C.charcoalSoft : C.review }}>
                {value.gps ? `${value.gps.lat}, ${value.gps.lng}` : "GPS unavailable"}
              </span>
            </div>
            <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 8.5, color: "#A6AFB6", marginBottom: 6 }}>
              from your phone at upload — not the image file
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}>
              {fieldsConfig.map((f) => (
                <div key={f.key}>
                  <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9, color: C.charcoalSoft }}>{f.label}</div>
                  <input
                    value={value[f.key] || ""}
                    onChange={(e) => updateField(f.key, e.target.value)}
                    placeholder="—"
                    style={{
                      width: "100%", boxSizing: "border-box", padding: "4px 6px", borderRadius: 5,
                      border: `1px solid ${C.line}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 11, color: C.charcoal
                    }}
                  />
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
      {error && (
        <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, color: C.review, marginTop: 6 }}>{error}</div>
      )}
    </div>
  );
}

/* ---------- shared bits ---------- */
function Logo({ dark }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
      <div style={{
        width: 26, height: 26, borderRadius: 7, background: `linear-gradient(135deg, ${C.primary}, ${C.primaryDeep})`,
        display: "flex", alignItems: "center", justifyContent: "center"
      }}>
        <Droplet size={15} color="#fff" fill="#fff" />
      </div>
      <span style={{
        fontFamily: "'Space Grotesk',sans-serif", fontWeight: 700, fontSize: 16, letterSpacing: -0.2,
        color: dark ? "#fff" : C.charcoal
      }}>
        THYNK<span style={{ color: C.primary }}>-H2O</span>
      </span>
    </div>
  );
}

function NavPills({ screen, setScreen, captures, role }) {
  const items = [
    { id: "setup", label: "Setup" },
    { id: "capture", label: "Field Capture" },
    ...(role !== "field" ? [{ id: "office", label: `Office Review${captures.length ? ` (${captures.length})` : ""}` }] : []),
  ];
  return (
    <div style={{ display: "flex", gap: 6, background: C.paperDeep, padding: 4, borderRadius: 999 }}>
      {items.map((it) => (
        <button key={it.id} onClick={() => setScreen(it.id)}
          style={{
            border: "none", cursor: "pointer", padding: "7px 14px", borderRadius: 999,
            fontFamily: "'Inter',sans-serif", fontSize: 12.5, fontWeight: 600,
            background: screen === it.id ? C.charcoal : "transparent",
            color: screen === it.id ? "#fff" : C.charcoalSoft,
            transition: "all 0.15s"
          }}>
          {it.label}
        </button>
      ))}
    </div>
  );
}

const STATUS_META = {
  captured: { label: "Awaiting Review", color: C.charcoalSoft, bg: C.paperDeep, icon: ClipboardList },
  needs_review: { label: "Needs Review", color: C.review, bg: C.reviewSoft, icon: AlertTriangle },
  approved: { label: "Approved", color: C.approve, bg: C.approveSoft, icon: CheckCircle2 },
  flagged: { label: "Flagged", color: C.flag, bg: C.flagSoft, icon: Flag },
};

const STATUS_EXPORT_LABEL = {
  captured: "Awaiting Review",
  needs_review: "Needs Review",
  approved: "Approved",
  flagged: "Flagged",
};

function safeFilename(name) {
  return (name || "Survey").replace(/[^a-z0-9-_]+/gi, "_");
}

function exportExcel(survey, captures, reviewer) {
  // Older surveys used separate "meter" and "replacement" task types
  const rawTask = survey.taskType || "meterwork";
  const task = rawTask === "meter" || rawTask === "replacement" ? "meterwork" : rawTask;
  const meterMode = survey.meterMode || (rawTask === "replacement" ? "replace" : "read");
  const wb = XLSX.utils.book_new();

  // Fields every task shares
  const common = (c) => ({
    Site: survey.siteName || "",
    Survey: survey.surveyName || "",
    "Position Name": c.position,
  });
  const trailing = (c) => ({
    Date: c.timestamp.date,
    Time: c.timestamp.time,
    GPS: c.gps ? `${c.gps.lat}, ${c.gps.lng}` : "",
    Technician: c.tech || "",
    Reviewer: c.status === "approved" ? (reviewer || "") : "",
    Status: STATUS_EXPORT_LABEL[c.status] || c.status,
  });

  let rows = [];
  let widths = [];
  let sheetName = "Captures";
  let fileSuffix = "report";

  if (task === "amr") {
    sheetName = "AMR Assets";
    fileSuffix = "amr_survey";
    rows = captures.map((c) => ({
      ...common(c),
      "Asset Type": c.amrAssetType || "",
      "Asset Serial": c.amrSerial || "",
      "Meters Linked": (c.amrShots || []).reduce((n, s) => n + (s.meters || []).length, 0),
      "Good Signal": (c.amrShots || []).flatMap((s) => s.meters || []).filter((m) => { const b = signalBand(m.signal); return b && b.label === "GOOD"; }).length,
      "Fair Signal": (c.amrShots || []).flatMap((s) => s.meters || []).filter((m) => { const b = signalBand(m.signal); return b && b.label === "FAIR"; }).length,
      "Weak Signal": (c.amrShots || []).flatMap((s) => s.meters || []).filter((m) => { const b = signalBand(m.signal); return b && b.label === "WEAK"; }).length,
      "No Meters Allocated": c.noMetersAllocated ? "YES" : "",
      "No Meters Reason": c.noMetersNote || "",
      "Expected On Register": c.expectedCount ?? "",
      "Found": c.matchedCount ?? "",
      "Missing": (c.missingMeters || []).length || "",
      "Missing Serials": (c.missingMeters || []).join("; "),
      "Not On Register": (c.extraMeters || []).join("; "),
      ...trailing(c),
    }));
    widths = [{ wch: 18 }, { wch: 20 }, { wch: 16 }, { wch: 12 }, { wch: 18 }, { wch: 13 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 16 }, { wch: 30 }, { wch: 18 }, { wch: 10 }, { wch: 10 }, { wch: 40 }, { wch: 30 }, { wch: 12 }, { wch: 9 }, { wch: 24 }, { wch: 14 }, { wch: 14 }, { wch: 15 }];
  } else if (task === "assets") {
    sheetName = "Asset Register";
    fileSuffix = "asset_register";
    rows = captures.map((c) => ({
      asset_id: c.position || "",
      category: c.assetCategory || "",
      asset_type: c.assetType || "",
      description: c.assetDescription || "",
      serial: c.assetSerial || "",
      latitude: c.gps?.lat && c.gps.lat !== "Unknown" ? c.gps.lat : "",
      longitude: c.gps?.lng && c.gps.lng !== "Unknown" ? c.gps.lng : "",
      zone_or_street: c.assetZone || "",
      erf_or_unit: c.assetErf || "",
      access_notes: c.assetAccessNotes || "",
      last_service_date: c.assetLastService || "",
      next_service_due: (() => { const s = serviceStatus(c.assetType, c.assetLastService); return s ? s.dueDate : ""; })(),
      service_status: (() => { const s = serviceStatus(c.assetType, c.assetLastService); return s ? (s.overdue ? "OVERDUE" : "IN DATE") : (SERVICEABLE_TYPES.has(c.assetType) ? "NO RECORD" : ""); })(),
      last_cartridge_replacement: c.assetLastCartridge || "",
      cartridge_due: (() => { const s = dueStatus(c.assetLastCartridge, CARTRIDGE_INTERVAL_MONTHS); return s ? `${s.dueDate}${s.overdue ? " (OVERDUE)" : ""}` : ""; })(),
      last_flush: c.assetLastFlush || "",
      flush_due: (() => { const s = dueStatus(c.assetLastFlush, FLUSH_INTERVAL_MONTHS); return s ? `${s.dueDate}${s.overdue ? " (OVERDUE)" : ""}` : ""; })(),
      cartridge_condition: c.assetCartridgeCondition || "",
      filter_spec: c.assetFilterSpec || "",
      date_captured: c.timestamp.date,
      captured_by: c.tech || "",
      verified_by: c.status === "approved" ? (reviewer || "") : "",
      status: STATUS_EXPORT_LABEL[c.status] || c.status,
    }));
    widths = [{ wch: 16 }, { wch: 15 }, { wch: 30 }, { wch: 46 }, { wch: 16 }, { wch: 13 }, { wch: 13 }, { wch: 20 }, { wch: 14 }, { wch: 34 }, { wch: 16 }, { wch: 16 }, { wch: 14 }, { wch: 22 }, { wch: 18 }, { wch: 14 }, { wch: 18 }, { wch: 20 }, { wch: 18 }, { wch: 13 }, { wch: 14 }, { wch: 14 }, { wch: 15 }];
  } else if (task === "meterwork" && meterMode === "replace") {
    sheetName = "Meter Replacements";
    fileSuffix = "meter_replacements";
    rows = captures.map((c) => ({
      ...common(c),
      "Meter Type": c.meterType || "",
      "Old Meter Serial": c.oldMeter?.serial || "",
      "Old Meter Reading (m\u00b3)": c.oldMeter?.reading || "",
      "New Meter Serial": c.newMeter?.serial || "",
      "New Meter Reading (m\u00b3)": c.newMeter?.reading || "",
      "Could Not Read": c.type === "exception" ? "YES" : "",
      "Reason": c.exceptionReason || "",
      "Problem Reported": (c.conditionFlags || []).join("; "),
      "Problem Detail": c.conditionNote || "",
      ...trailing(c),
    }));
    widths = [{ wch: 18 }, { wch: 20 }, { wch: 16 }, { wch: 20 }, { wch: 18 }, { wch: 20 }, { wch: 18 }, { wch: 20 }, { wch: 13 }, { wch: 26 }, { wch: 30 }, { wch: 30 }, { wch: 12 }, { wch: 9 }, { wch: 24 }, { wch: 14 }, { wch: 14 }, { wch: 15 }];
  } else if (task === "fido2") {
    sheetName = "FIDO Sensor Log";
    fileSuffix = "fido_leak_analysis";
    const byBug = buildBugStatus(captures);
    rows = captures
      .filter((c) => c.type === "fido2_deploy" || c.type === "fido2_retrieve" || c.type === "fido2_waypoint")
      .map((c) => {
        const b = byBug[c.bugSerial];
        const dep = b?.deployment;
        const ret = b?.retrieval;
        const a = (c.sessionId || "").trim().toLowerCase();
        const g = (c.resultSessionId || "").trim().toLowerCase();
        return {
          ...common(c),
          "Stage": c.type === "fido2_deploy" ? "DEPLOYED" : c.type === "fido2_waypoint" ? "INVESTIGATED" : "RETRIEVED",
          "Waypoint": c.waypointRef || "",
          "Tests Carried Out": (c.tests || []).map((t) => t.method + (t.distance ? ` @ ${t.distance}` : "")).join("; "),
          "Pinpoint GPS": c.pinpointGps ? `${c.pinpointGps.lat}, ${c.pinpointGps.lng}` : "",
          "Bug Serial": c.bugSerial || c.linkedBug || "",
          "Deployed On": c.deployAsset || "",
          "Pipe Material": PIPE_MATERIALS[c.pipeMaterial]?.label || "",
          "Coupling": (() => { const a = FIDO_DEPLOY_ASSETS.find((x) => x.v === c.deployAsset); return a ? (a.coupling > 1 ? "DIRECT" : "INDIRECT") : ""; })(),
          "Pipe Diameter": c.cond_diameter || "",
          "Pressure": c.cond_pressure || "",
          "Backfill": c.cond_backfill || "",
          "Pipe Condition": c.cond_pipecondition || "",
          "Background Noise": c.cond_background || "",
          "Consumption": c.cond_consumption || "",
          "Listening Conditions": (() => { const s = conditionScore(c); return s ? `${s.label} (${s.pct}%)` : ""; })(),
          "Normal Leak Frequency": PIPE_MATERIALS[c.pipeMaterial]?.normal || "",
          "Frequency Range": PIPE_MATERIALS[c.pipeMaterial]?.freq || "",
          "Session Mode": c.fidoMode || "",
          "Session ID": c.sessionId || "",
          "Session ID On Graph": c.resultSessionId || "",
          "ID Match": c.type === "fido2_retrieve" ? (a && g ? (a === g ? "MATCH" : "MISMATCH") : "NOT CONFIRMED") : "",
          "Days Deployed": c.type === "fido2_retrieve" ? (c.deployedDays ?? "") : (dep && !ret ? daysSince(dep.id) : ""),
          "Finding": c.outcome || "",
          "Note": c.outcomeNote || "",
          ...trailing(c),
        };
      });
    widths = [{ wch: 18 }, { wch: 20 }, { wch: 16 }, { wch: 13 }, { wch: 10 }, { wch: 34 }, { wch: 22 }, { wch: 14 }, { wch: 20 }, { wch: 15 }, { wch: 11 }, { wch: 22 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 18 }, { wch: 26 }, { wch: 21 }, { wch: 13 }, { wch: 19 }, { wch: 14 }, { wch: 14 }, { wch: 18 }, { wch: 34 }, { wch: 12 }, { wch: 9 }, { wch: 24 }, { wch: 14 }, { wch: 14 }, { wch: 15 }];
  } else if (task === "fido") {
    sheetName = "FIDO Deployments";
    fileSuffix = "fido_leak_analysis";
    rows = captures.map((c) => ({
      ...common(c),
      "Capture Type": c.type === "sensor" ? "Sensor Deployment" : "FIDO Feedback",
      "Session Type": c.fido?.sessionType || "",
      "Bug / Sensor Serial": c.sensor?.deviceId || "",
      "Session ID": c.sensor?.sessionId || "",
      "Signal": c.sensor?.signalStrength || "",
      "Battery": c.sensor?.batteryVoltage || "",
      ...trailing(c),
    }));
    widths = [{ wch: 18 }, { wch: 20 }, { wch: 16 }, { wch: 18 }, { wch: 24 }, { wch: 20 }, { wch: 16 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 9 }, { wch: 24 }, { wch: 14 }, { wch: 14 }, { wch: 15 }];
  } else {
    sheetName = "Meter Readings";
    fileSuffix = "meter_survey";
    rows = captures.map((c) => {
      const prev = lookupPrevious(survey.previousReadings, c.position, c.serial);
      const used = prev && !isNaN(parseFloat(c.reading)) && !isNaN(parseFloat(prev.reading))
        ? (parseFloat(c.reading) - parseFloat(prev.reading)).toFixed(2) : "";
      return {
        ...common(c),
        "Meter Type": c.meterType || "",
        "Serial Number": c.serial || "",
        "Meter Reading (m\u00b3)": c.type === "exception" ? "" : (c.reading || ""),
        "Previous Reading (m\u00b3)": prev ? prev.reading : "",
        "Consumption (m\u00b3)": used,
        "AI Confidence %": c.type === "exception" ? "" : c.confidence,
        "Could Not Read": c.type === "exception" ? "YES" : "",
        "Reason": c.exceptionReason || "",
        "Needs Revisit": c.needsRevisit ? "YES" : "",
        "Problem Reported": (c.conditionFlags || []).join("; "),
        "Problem Detail": c.conditionNote || "",
        ...trailing(c),
      };
    });
    widths = [{ wch: 18 }, { wch: 20 }, { wch: 16 }, { wch: 20 }, { wch: 16 }, { wch: 18 }, { wch: 19 }, { wch: 16 }, { wch: 15 }, { wch: 13 }, { wch: 26 }, { wch: 13 }, { wch: 30 }, { wch: 30 }, { wch: 12 }, { wch: 9 }, { wch: 24 }, { wch: 14 }, { wch: 14 }, { wch: 15 }];
  }

  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = widths;
  XLSX.utils.book_append_sheet(wb, ws, sheetName);

  // FIDO 2 gets a findings sheet: one row per sensor, deployment paired with retrieval
  if (task === "fido2") {
    const status = buildBugStatus(captures);
    const findingRows = Object.values(status)
      .filter((b) => b.deployment)
      .map((b) => {
        const d = b.deployment;
        const r = b.retrieval;
        const a = (d.sessionId || "").trim().toLowerCase();
        const g = (r?.resultSessionId || "").trim().toLowerCase();
        return {
          Site: survey.siteName || "",
          "Bug Serial": b.serial,
          "Deployed On": d.deployAsset || "",
          Position: d.position || "",
          Latitude: d.gps?.lat && d.gps.lat !== "Unknown" ? d.gps.lat : "",
          Longitude: d.gps?.lng && d.gps.lng !== "Unknown" ? d.gps.lng : "",
          "Pipe Material": PIPE_MATERIALS[d.pipeMaterial]?.label || "",
          "Listening Conditions": (() => { const s = conditionScore(d); return s ? `${s.label} (${s.pct}%)` : ""; })(),
          "Normal Frequency": PIPE_MATERIALS[d.pipeMaterial]?.normal || "",
          "Frequency Range": PIPE_MATERIALS[d.pipeMaterial]?.freq || "",
          "Session Mode": d.fidoMode || "",
          "Session ID": d.sessionId || "",
          "Deployed": d.timestamp?.date || "",
          "Retrieved": r?.timestamp?.date || "",
          "Days Out": r ? (r.deployedDays ?? "") : daysSince(d.id),
          "ID Match": r ? (a && g ? (a === g ? "MATCH" : "MISMATCH") : "NOT CONFIRMED") : "",
          Finding: r?.outcome || (r ? "" : "STILL DEPLOYED"),
          Note: r?.outcomeNote || "",
          Technician: d.tech || "",
        };
      });
    if (findingRows.length) {
      const fWs = XLSX.utils.json_to_sheet(findingRows);
      fWs["!cols"] = [
        { wch: 18 }, { wch: 14 }, { wch: 20 }, { wch: 18 }, { wch: 13 }, { wch: 13 },
        { wch: 15 }, { wch: 20 }, { wch: 21 }, { wch: 14 }, { wch: 13 }, { wch: 13 },
        { wch: 11 }, { wch: 15 }, { wch: 18 }, { wch: 34 }, { wch: 14 },
      ];
      XLSX.utils.book_append_sheet(wb, fWs, "Leak Findings");
    }
  }

  // AMR surveys get a second sheet: one row per linked meter
  if (task === "amr") {
    const dupes = findDuplicateMeters(captures);
    const amrRows = [];
    captures.forEach((c) => {
      (c.amrShots || []).forEach((shot) => {
        (shot.meters || []).forEach((m) => {
          amrRows.push({
            Site: survey.siteName || "",
            Position: c.position,
            "Asset Type": c.amrAssetType || "",
            "Asset Serial": c.amrSerial || "",
            "Concentrator ID": shot.concentratorId || "",
            "Meter Serial": m.serial || "",
            "Meter Model": m.model || "",
            "Signal Strength": m.signal || "",
            "Signal Band": (() => { const b = signalBand(m.signal); return b ? b.label : ""; })(),
            "Duplicate": (() => {
              const key = (m.serial || "").trim().toUpperCase();
              const d = dupes[key];
              if (!d) return "";
              return d.best && d.best.captureId === c.id ? "YES - STRONGEST" : "YES - WEAKER";
            })(),
            "Also Seen By": (() => {
              const key = (m.serial || "").trim().toUpperCase();
              const d = dupes[key];
              if (!d) return "";
              return d.sightings.filter((s) => s.captureId !== c.id)
                .map((s) => `${s.assetType} ${s.position} (${s.signal || "no signal"})`).join("; ");
            })(),
            "Reading Date": shot.readingDate || c.timestamp.date,
            GPS: c.gps ? `${c.gps.lat}, ${c.gps.lng}` : "",
            Technician: c.tech || "",
          });
        });
      });
    });
    if (amrRows.length) {
      const amrWs = XLSX.utils.json_to_sheet(amrRows);
      amrWs["!cols"] = [
        { wch: 18 }, { wch: 14 }, { wch: 12 }, { wch: 16 }, { wch: 16 },
        { wch: 16 }, { wch: 16 }, { wch: 14 }, { wch: 12 }, { wch: 16 }, { wch: 38 },
        { wch: 14 }, { wch: 24 }, { wch: 14 },
      ];
      XLSX.utils.book_append_sheet(wb, amrWs, "Linked Meters");
    }

    // Meters expected on the register that did not report
    const missingRows = [];
    captures.filter((c) => c.type === "amr").forEach((c) => {
      (c.missingMeters || []).forEach((m) => {
        missingRows.push({
          Site: survey.siteName || "",
          "Expected On": c.amrSerial || "",
          "Asset Type": c.amrAssetType || "",
          Position: c.position || "",
          "Meter Serial": m,
          "Survey Date": c.timestamp?.date || "",
          Technician: c.tech || "",
        });
      });
    });
    if (missingRows.length) {
      const mWs = XLSX.utils.json_to_sheet(missingRows);
      mWs["!cols"] = [{ wch: 18 }, { wch: 16 }, { wch: 12 }, { wch: 16 }, { wch: 16 }, { wch: 13 }, { wch: 14 }];
      XLSX.utils.book_append_sheet(wb, mWs, "Missing Meters");
    }
  }

  XLSX.writeFile(wb, `${safeFilename(survey.surveyName)}_${fileSuffix}.xlsx`);
}

/* CSV export in the estate asset intake template format */
function exportAssetCSV(survey, captures) {
  const headers = [
    "asset_id", "category", "asset_type", "description", "serial",
    "latitude", "longitude", "zone_or_street", "erf_or_unit", "access_notes",
    "last_service_date", "next_service_due", "service_status",
    "last_cartridge_replacement", "cartridge_due", "last_flush", "flush_due",
    "cartridge_condition", "filter_spec",
    "date_captured", "captured_by",
  ];
  const esc = (v) => {
    const s = (v === null || v === undefined) ? "" : String(v);
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const lines = [headers.join(",")];
  captures.forEach((c) => {
    lines.push([
      c.position || "",
      c.assetCategory || "",
      c.assetType || "",
      c.assetDescription || "",
      c.assetSerial || "",
      c.gps?.lat && c.gps.lat !== "Unknown" ? c.gps.lat : "",
      c.gps?.lng && c.gps.lng !== "Unknown" ? c.gps.lng : "",
      c.assetZone || "",
      c.assetErf || "",
      c.assetAccessNotes || "",
      c.assetLastService || "",
      (() => { const s = serviceStatus(c.assetType, c.assetLastService); return s ? s.dueDate : ""; })(),
      (() => { const s = serviceStatus(c.assetType, c.assetLastService); return s ? (s.overdue ? "OVERDUE" : "IN DATE") : (SERVICEABLE_TYPES.has(c.assetType) ? "NO RECORD" : ""); })(),
      c.assetLastCartridge || "",
      (() => { const s = dueStatus(c.assetLastCartridge, CARTRIDGE_INTERVAL_MONTHS); return s ? `${s.dueDate}${s.overdue ? " (OVERDUE)" : ""}` : ""; })(),
      c.assetLastFlush || "",
      (() => { const s = dueStatus(c.assetLastFlush, FLUSH_INTERVAL_MONTHS); return s ? `${s.dueDate}${s.overdue ? " (OVERDUE)" : ""}` : ""; })(),
      c.assetCartridgeCondition || "",
      c.assetFilterSpec || "",
      c.timestamp?.date || "",
      c.tech || "",
    ].map(esc).join(","));
  });
  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${safeFilename(survey.surveyName)}_asset_register.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/* FIDO leak analysis report — written for the client.
   Tells the story: what was surveyed, what was flagged, what was investigated, what was found.
   Internal detail (acoustics, listening conditions, coupling) stays in the app and the spreadsheet. */
function printFidoReport(survey, captures, reviewer) {
  const win = window.open("", "_blank", "width=1000,height=1100");
  if (!win) return;

  const status = buildBugStatus(captures);
  const deployments = Object.values(status).filter((b) => b.deployment);
  const waypoints = captures.filter((c) => c.type === "fido2_waypoint");

  const outcomeColour = {
    "NO LEAK": "#1B9C6E",
    "SUSPECTED LEAK": "#D98A22",
    "CONFIRMED LEAK": "#D6485A",
    "INCONCLUSIVE": "#5B6570",
    "STILL DEPLOYED": "#0D86F3",
    "LEAK CONFIRMED": "#D6485A",
    "LEAK SUSPECTED": "#D98A22",
    "NO LEAK FOUND": "#1B9C6E",
    "ONGOING": "#0D86F3",
  };

  const counts = {};
  deployments.forEach((b) => {
    const k = b.retrieval ? (b.retrieval.outcome || "INCONCLUSIVE") : "STILL DEPLOYED";
    counts[k] = (counts[k] || 0) + 1;
  });

  const flagged = deployments.filter((b) =>
    b.retrieval && (b.retrieval.outcome === "SUSPECTED LEAK" || b.retrieval.outcome === "CONFIRMED LEAK"));
  const confirmed = waypoints.filter((w) => w.outcome === "LEAK CONFIRMED");

  /* Stage 1 — coverage */
  const coverageRows = deployments.map((b) => {
    const d = b.deployment, r = b.retrieval;
    const outcome = r ? (r.outcome || "INCONCLUSIVE") : "STILL DEPLOYED";
    return `<tr>
      <td class="mono">${b.serial}</td>
      <td>${d.position || "\u2014"}</td>
      <td>${d.deployAsset || "\u2014"}</td>
      <td>${d.timestamp?.date || "\u2014"}</td>
      <td>${r ? `${r.deployedDays ?? "\u2014"} days` : "out"}</td>
      <td style="color:${outcomeColour[outcome]};font-weight:700;">${outcome}</td>
    </tr>`;
  }).join("");

  /* Stage 2 — flagged points, with the session graph as evidence */
  const flaggedSections = flagged.map((b) => {
    const d = b.deployment, r = b.retrieval;
    const col = outcomeColour[r.outcome];
    const wp = waypoints.find((w) => w.linkedBug === b.serial);
    return `
      <div class="block">
        <div class="blockhead" style="border-left-color:${col};">
          <div>
            <div class="btitle">${d.position || b.serial}</div>
            <div class="bsub">Sensor ${b.serial} on ${(d.deployAsset || "").toLowerCase()} \u00b7 ${r.deployedDays ?? "\u2014"} days${wp ? ` \u00b7 investigated as ${wp.waypointRef}` : ""}</div>
          </div>
          <div class="pill" style="background:${col};">${r.outcome}</div>
        </div>
        ${r.resultShot?.photo ? `
          <div class="graphwrap">
            <div class="graphlabel">SENSOR SESSION${r.resultShot.dateRange ? ` \u00b7 ${r.resultShot.dateRange}` : ""}</div>
            <img src="${r.resultShot.photo}" class="graph" />
          </div>` : ""}
        ${r.outcomeNote ? `<div class="note"><b>Technician note.</b> ${r.outcomeNote}</div>` : ""}
      </div>`;
  }).join("");

  /* Stage 3 & 4 — what was done at each waypoint and what was found */
  const waypointSections = waypoints.map((w) => {
    const col = outcomeColour[w.outcome] || "#5B6570";
    const tests = w.tests || [];
    return `
      <div class="block">
        <div class="blockhead" style="border-left-color:${col};">
          <div>
            <div class="btitle">${w.waypointRef} \u00b7 ${w.position || ""}</div>
            <div class="bsub">${w.linkedBug ? `Raised by sensor ${w.linkedBug}` : "Investigated directly"} \u00b7 ${tests.length} test${tests.length === 1 ? "" : "s"} carried out</div>
          </div>
          <div class="pill" style="background:${col};">${w.outcome}</div>
        </div>

        ${tests.length ? tests.map((t) => `
          <div class="graphwrap">
            <div class="graphlabel">${t.method}${t.distance ? ` \u00b7 LEAK AT ${t.distance}` : ""}</div>
            <img src="${t.photo}" class="graph" />
            ${t.note ? `<div class="graphnote">${t.note}</div>` : ""}
          </div>`).join("") : `<div class="nograph">No test evidence attached</div>`}

        ${w.pinpointPhoto ? `
          <div class="pinpoint">
            <img src="${w.pinpointPhoto}" />
            <div>
              <div class="plabel">LEAK POSITION MARKED</div>
              <div class="mono pcoord">${w.pinpointGps ? `${w.pinpointGps.lat}, ${w.pinpointGps.lng}` : "\u2014"}</div>
              <div class="pnote">Beacon placed on site at the located position.</div>
            </div>
          </div>` : ""}

        ${w.outcomeNote ? `<div class="note"><b>Conclusion.</b> ${w.outcomeNote}</div>` : ""}
      </div>`;
  }).join("");

  win.document.write(`
    <!doctype html><html><head><title>${survey.siteName || "Estate"} Leak Detection Report</title>
    <style>
      * { box-sizing: border-box; }
      body { font-family: Arial, Helvetica, sans-serif; color:#2B2F33; padding:28px 32px; margin:0; }
      .brand { display:flex; align-items:center; gap:8px; margin-bottom:4px; }
      .brand .box { width:20px; height:20px; border-radius:5px; background:#0D86F3; }
      .brand span { font-weight:700; font-size:16px; }
      h1 { font-size:21px; margin:14px 0 2px; }
      h2 { font-size:15px; margin:30px 0 4px; padding-bottom:5px; border-bottom:2px solid #DCE3E8; }
      .lead { font-size:12px; color:#5B6570; margin:0 0 14px; line-height:1.5; }
      .meta { color:#5B6570; font-size:12.5px; margin-bottom:18px; }

      .summary { display:flex; gap:10px; flex-wrap:wrap; margin-bottom:8px; }
      .stat { flex:1; min-width:105px; padding:12px 14px; border-radius:10px; text-align:center; }
      .stat b { display:block; font-size:24px; line-height:1.1; }
      .stat span { font-size:9px; letter-spacing:0.4px; font-weight:700; }

      table.cov { width:100%; border-collapse:collapse; font-size:11px; margin-top:8px; }
      table.cov th { background:#F4F7F9; color:#5B6570; font-size:9px; letter-spacing:0.3px;
        text-align:left; padding:7px 9px; border:1px solid #DCE3E8; }
      table.cov td { padding:6px 9px; border:1px solid #DCE3E8; }

      .block { border:1px solid #DCE3E8; border-radius:10px; margin-bottom:18px;
        page-break-inside:avoid; overflow:hidden; padding-bottom:14px; }
      .blockhead { display:flex; justify-content:space-between; align-items:center;
        padding:12px 16px; background:#F4F7F9; border-left:5px solid #5B6570; }
      .btitle { font-size:14px; font-weight:700; }
      .bsub { font-size:11px; color:#5B6570; margin-top:2px; }
      .pill { color:#fff; font-size:10px; font-weight:700; padding:5px 13px; border-radius:999px; }

      .graphwrap { margin:13px 16px 0; }
      .graphlabel { font-size:9px; font-weight:700; color:#5B6570; letter-spacing:0.4px; margin-bottom:5px; }
      .graph { width:100%; height:auto; max-height:430px; object-fit:contain;
        border:1px solid #DCE3E8; border-radius:8px; background:#fff; display:block; }
      .graphnote { font-size:11px; color:#5B6570; margin-top:6px; font-style:italic; }
      .nograph { margin:13px 16px 0; padding:18px; text-align:center; border:1px dashed #DCE3E8;
        border-radius:8px; color:#A6AFB6; font-size:11px; }

      .pinpoint { display:flex; gap:14px; margin:13px 16px 0; padding:12px;
        background:#FBE6E9; border-radius:9px; align-items:center; }
      .pinpoint img { width:150px; height:112px; object-fit:cover; border-radius:7px; flex-shrink:0; }
      .plabel { font-size:9px; font-weight:700; color:#D6485A; letter-spacing:0.4px; }
      .pcoord { font-size:14px; font-weight:700; margin-top:3px; }
      .pnote { font-size:10.5px; color:#5B6570; margin-top:4px; }

      .note { margin:12px 16px 0; padding:10px 13px; background:#F4F7F9; border-radius:7px;
        font-size:11.5px; line-height:1.5; }
      .note b { color:#2B2F33; }
      .mono { font-family:'Courier New', monospace; }
      .footer { margin-top:26px; padding-top:12px; border-top:1px solid #DCE3E8;
        font-size:10.5px; color:#5B6570; line-height:1.5; }
      @media print { .no-print { display:none; } .block { page-break-inside:avoid; } }
    </style></head><body>
      <div class="brand"><div class="box"></div><span>THYNK-H2O</span></div>
      <h1>${survey.siteName || "Untitled Site"} \u2014 Leak Detection Report</h1>
      <div class="meta">
        ${survey.address ? `${survey.address} &nbsp;\u00b7&nbsp;` : ""}
        Survey: ${survey.surveyName || "\u2014"} &nbsp;\u00b7&nbsp;
        Technician: ${survey.tech || "\u2014"} &nbsp;\u00b7&nbsp;
        Report date: ${new Date().toLocaleDateString()}
      </div>

      <h2>Summary</h2>
      <p class="lead">
        ${deployments.length} acoustic sensor${deployments.length === 1 ? " was" : "s were"} deployed across the estate.
        ${flagged.length ? `${flagged.length} point${flagged.length === 1 ? "" : "s"} returned a signal warranting further investigation.` : "No points returned a signal warranting further investigation."}
        ${waypoints.length ? ` ${waypoints.length} waypoint${waypoints.length === 1 ? " was" : "s were"} opened and investigated on site.` : ""}
        ${confirmed.length ? ` ${confirmed.length} leak${confirmed.length === 1 ? " was" : "s were"} confirmed and marked for repair.` : ""}
      </p>
      <div class="summary">
        <div class="stat" style="background:#F4F7F9;">
          <b>${deployments.length}</b><span style="color:#5B6570;">SENSORS DEPLOYED</span>
        </div>
        ${Object.entries(counts).map(([k, n]) => `
          <div class="stat" style="background:${outcomeColour[k]}18;">
            <b style="color:${outcomeColour[k]};">${n}</b>
            <span style="color:${outcomeColour[k]};">${k}</span>
          </div>`).join("")}
        ${confirmed.length ? `<div class="stat" style="background:#FBE6E9;">
          <b style="color:#D6485A;">${confirmed.length}</b><span style="color:#D6485A;">LEAKS CONFIRMED</span>
        </div>` : ""}
      </div>

      <h2>1. Survey coverage</h2>
      <p class="lead">Every sensor deployed during this survey, where it was placed and what it returned.</p>
      <table class="cov">
        <thead><tr><th>Sensor</th><th>Location</th><th>Deployed on</th><th>Date</th><th>Duration</th><th>Result</th></tr></thead>
        <tbody>${coverageRows || `<tr><td colspan="6">No deployments recorded.</td></tr>`}</tbody>
      </table>

      ${flagged.length ? `
        <h2>2. Points flagged for investigation</h2>
        <p class="lead">These sensors returned a signal consistent with a possible leak. The session data is shown below.</p>
        ${flaggedSections}` : ""}

      ${waypoints.length ? `
        <h2>${flagged.length ? "3" : "2"}. Investigations carried out</h2>
        <p class="lead">Each flagged area was opened as a numbered waypoint and investigated on site using acoustic correlation, sounding and ground microphone as appropriate.</p>
        ${waypointSections}` : ""}

      <div class="footer">
        Acoustic leak detection identifies leak noise transmitted through the pipe network. Detection is
        influenced by pipe material, diameter, pressure, backfill, background noise and consumption at the
        time of the survey; these were recorded for every deployment and are available on request.
        Confirmed leak positions were marked on site and their coordinates recorded.
        <br/><br/>
        Prepared by ${survey.tech || "\u2014"}${reviewer ? `, verified by ${reviewer}` : ""} for ${survey.siteName || "the estate"}.
      </div>
      <div class="no-print" style="margin-top:24px;">
        <button onclick="window.print()" style="background:#0D86F3;color:#fff;border:none;padding:10px 18px;border-radius:8px;font-weight:600;cursor:pointer;">Print / Save as PDF</button>
      </div>
    </body></html>
  `);
  win.document.close();
}

/* Asset register PDF — one card per asset so a person can go and find it */
function printAssetRegister(survey, captures, reviewer) {
  const win = window.open("", "_blank", "width=900,height=1000");
  if (!win) return;

  const byCategory = {};
  captures.forEach((c) => {
    const cat = c.assetCategory || "UNCATEGORISED";
    (byCategory[cat] = byCategory[cat] || []).push(c);
  });

  const catColour = { WATER: "#0D86F3", ELECTRICAL: "#D98A22", "AMR EQUIPMENT": "#7B5BD6", FILTRATION: "#0E9AA7", FIRE: "#D6485A", GENERAL: "#5B6570", UNCATEGORISED: "#5B6570" };

  const serviceable = captures.filter((c) => SERVICEABLE_TYPES.has(c.assetType) || FILTRATION_TYPES.has(c.assetType));
  const isOverdue = (c) => {
    if (FILTRATION_TYPES.has(c.assetType)) {
      const cart = dueStatus(c.assetLastCartridge, CARTRIDGE_INTERVAL_MONTHS);
      const flush = dueStatus(c.assetLastFlush, FLUSH_INTERVAL_MONTHS);
      return (cart && cart.overdue) || (flush && flush.overdue);
    }
    const s = serviceStatus(c.assetType, c.assetLastService);
    return s && s.overdue;
  };
  const hasRecord = (c) => FILTRATION_TYPES.has(c.assetType)
    ? !!(c.assetLastCartridge || c.assetLastFlush)
    : !!serviceStatus(c.assetType, c.assetLastService);
  const overdueItems = serviceable.filter(isOverdue);
  const noRecordItems = serviceable.filter((c) => !hasRecord(c));

  const serviceBanner = serviceable.length ? `
      <div class="servicebox">
        <div class="servicehead">SERVICE COMPLIANCE \u2014 ${serviceable.length} SERVICEABLE ASSETS</div>
        <div class="servicerow">
          <span class="pill ok">${serviceable.length - overdueItems.length - noRecordItems.length} IN DATE</span>
          <span class="pill bad">${overdueItems.length} OVERDUE</span>
          <span class="pill warn">${noRecordItems.length} NO RECORD</span>
        </div>
        ${overdueItems.length ? `<div class="servicelist"><b>Overdue:</b> ${overdueItems.map((c) => `${c.assetType} \u00b7 ${c.assetDescription || c.position}`).join("; ")}</div>` : ""}
        ${noRecordItems.length ? `<div class="servicelist"><b>No service date recorded:</b> ${noRecordItems.map((c) => `${c.assetType} \u00b7 ${c.assetDescription || c.position}`).join("; ")}</div>` : ""}
      </div>` : "";

  const sections = Object.entries(byCategory).map(([cat, items]) => `
    <h2 style="font-size:15px; margin:26px 0 10px; color:${catColour[cat] || "#2B2F33"}; border-bottom:2px solid ${catColour[cat] || "#DCE3E8"}; padding-bottom:4px;">
      ${cat} <span style="font-weight:normal;color:#5B6570;font-size:12px;">(${items.length})</span>
    </h2>
    ${items.map((c) => `
      <div class="card">
        ${c.photo ? `<img src="${c.photo}" class="cardphoto" />` : `<div class="cardphoto placeholder">NO PHOTO</div>`}
        <div class="cardbody">
          <div class="cardtype">${c.assetType || "—"}</div>
          <div class="carddesc">${c.assetDescription || "NO DESCRIPTION RECORDED"}</div>
          <table class="cardmeta">
            <tr><td>COORDINATES</td><td class="mono">${c.gps && c.gps.lat !== "Unknown" ? `${c.gps.lat}, ${c.gps.lng}` : "NOT CAPTURED"}</td></tr>
            ${c.assetSerial ? `<tr><td>SERIAL</td><td class="mono">${c.assetSerial}</td></tr>` : ""}
            ${c.assetZone ? `<tr><td>ZONE / STREET</td><td>${c.assetZone}</td></tr>` : ""}
            ${c.assetErf ? `<tr><td>ERF / UNIT</td><td class="mono">${c.assetErf}</td></tr>` : ""}
            ${c.assetAccessNotes ? `<tr><td>ACCESS</td><td>${c.assetAccessNotes}</td></tr>` : ""}
            ${(() => {
              if (FILTRATION_TYPES.has(c.assetType)) {
                const cart = dueStatus(c.assetLastCartridge, CARTRIDGE_INTERVAL_MONTHS);
                const flush = dueStatus(c.assetLastFlush, FLUSH_INTERVAL_MONTHS);
                const badCond = c.assetCartridgeCondition === "FOULED" || c.assetCartridgeCondition === "NEEDS REPLACEMENT";
                return `
                  ${c.assetFilterSpec ? `<tr><td>FILTER SPEC</td><td class="mono">${c.assetFilterSpec}</td></tr>` : ""}
                  <tr><td>CARTRIDGE</td><td>${c.assetLastCartridge
                    ? `<span class="mono">${c.assetLastCartridge}</span> \u2014 <span class="${cart.overdue ? "overdue" : "indate"}">due ${cart.dueDate}${cart.overdue ? ` (OVERDUE ${cart.monthsOverdue}M)` : ""}</span>`
                    : `<span class="overdue">NO RECORD</span>`}</td></tr>
                  <tr><td>FLUSH</td><td>${c.assetLastFlush
                    ? `<span class="mono">${c.assetLastFlush}</span> \u2014 <span class="${flush.overdue ? "overdue" : "indate"}">due ${flush.dueDate}${flush.overdue ? ` (OVERDUE ${flush.monthsOverdue}M)` : ""}</span>`
                    : `<span class="overdue">NO RECORD</span>`}</td></tr>
                  ${c.assetCartridgeCondition ? `<tr><td>CONDITION</td><td class="${badCond ? "overdue" : "indate"}">${c.assetCartridgeCondition}</td></tr>` : ""}`;
              }
              const s = serviceStatus(c.assetType, c.assetLastService);
              if (!s && !SERVICEABLE_TYPES.has(c.assetType)) return "";
              if (!s) return `<tr><td>SERVICE</td><td class="overdue">NO SERVICE DATE RECORDED</td></tr>`;
              return `<tr><td>LAST SERVICE</td><td class="mono">${c.assetLastService}</td></tr>
                      <tr><td>NEXT DUE</td><td class="${s.overdue ? "overdue" : "indate"}">${s.dueDate}${s.overdue ? ` \u2014 OVERDUE BY ${s.monthsOverdue} MONTH${s.monthsOverdue === 1 ? "" : "S"}` : ""}</td></tr>`;
            })()}
            <tr><td>CAPTURED</td><td>${c.timestamp.date} · ${c.tech || "—"}</td></tr>
          </table>
        </div>
      </div>`).join("")}
  `).join("");

  win.document.write(`
    <!doctype html><html><head><title>${survey.siteName || "Estate"} Asset Register</title>
    <style>
      * { box-sizing: border-box; }
      body { font-family: Arial, Helvetica, sans-serif; color:#2B2F33; padding:32px; margin:0; }
      .brand { display:flex; align-items:center; gap:8px; margin-bottom:4px; }
      .brand .box { width:20px; height:20px; border-radius:5px; background:#0D86F3; }
      .brand span { font-weight:700; font-size:16px; }
      h1 { font-size:20px; margin:16px 0 2px; }
      .meta { color:#5B6570; font-size:12.5px; margin-bottom:14px; }
      .summary { display:flex; gap:22px; padding:12px 16px; background:#F4F7F9; border-radius:10px; margin-bottom:8px; }
      .summary b { display:block; font-size:19px; }
      .summary span { font-size:10.5px; color:#5B6570; letter-spacing:0.3px; }
      .card { display:flex; gap:14px; border:1px solid #DCE3E8; border-radius:10px; padding:12px; margin-bottom:10px; page-break-inside:avoid; }
      .cardphoto { width:150px; height:112px; object-fit:cover; border-radius:7px; flex-shrink:0; }
      .cardphoto.placeholder { display:flex; align-items:center; justify-content:center; background:#F4F7F9; color:#A6AFB6; font-size:10px; }
      .cardbody { flex:1; }
      .cardtype { font-size:13px; font-weight:700; margin-bottom:3px; }
      .carddesc { font-size:12px; color:#2B2F33; margin-bottom:8px; }
      .cardmeta { width:100%; border-collapse:collapse; font-size:11px; }
      .cardmeta td { padding:2px 0; vertical-align:top; }
      .cardmeta td:first-child { color:#5B6570; width:120px; font-size:9.5px; letter-spacing:0.3px; padding-top:3px; }
      .mono { font-family:'Courier New', monospace; }
      .overdue { color:#D6485A; font-weight:700; }
      .servicebox { border:1px solid #DCE3E8; border-left:4px solid #D6485A; border-radius:8px; padding:12px 16px; margin-bottom:14px; }
      .servicehead { font-size:11px; font-weight:700; color:#2B2F33; letter-spacing:0.4px; margin-bottom:8px; }
      .servicerow { display:flex; gap:8px; margin-bottom:6px; }
      .pill { font-size:10.5px; font-weight:700; padding:3px 10px; border-radius:999px; }
      .pill.ok { background:#E4F5EE; color:#1B9C6E; }
      .pill.bad { background:#FBE6E9; color:#D6485A; }
      .pill.warn { background:#FBF0DE; color:#D98A22; }
      .servicelist { font-size:10.5px; color:#5B6570; margin-top:5px; line-height:1.5; }
      .indate { color:#1B9C6E; font-weight:700; }
      .footer { margin-top:24px; font-size:11px; color:#5B6570; }
      @media print { .no-print { display:none; } }
    </style></head><body>
      <div class="brand"><div class="box"></div><span>THYNK-H2O</span></div>
      <h1>${survey.siteName || "Untitled Site"} — Asset Register</h1>
      <div class="meta">
        Survey: ${survey.surveyName || "—"} &nbsp;·&nbsp; Address: ${survey.address || "—"} &nbsp;·&nbsp;
        Captured by: ${survey.tech || "—"} &nbsp;·&nbsp; Verified by: ${reviewer || "—"} &nbsp;·&nbsp;
        Generated: ${new Date().toLocaleDateString()}
      </div>
      <div class="summary">
        <div><b>${captures.length}</b><span>ASSETS MAPPED</span></div>
        ${Object.entries(byCategory).map(([cat, items]) => `<div><b>${items.length}</b><span>${cat}</span></div>`).join("")}
      </div>
      ${serviceBanner}
      ${sections}
      <div class="footer">Coordinates are WGS84 decimal degrees. Generated by THYNK-H2O Field Capture.</div>
      <div class="no-print" style="margin-top:24px;">
        <button onclick="window.print()" style="background:#0D86F3;color:#fff;border:none;padding:10px 18px;border-radius:8px;font-weight:600;cursor:pointer;">Print / Save as PDF</button>
      </div>
    </body></html>
  `);
  win.document.close();
}

function printReport(survey, captures, reviewer, counts, pct) {
  const win = window.open("", "_blank", "width=900,height=1000");
  if (!win) return;
  const TYPE_SHORT = { meter: "Meter", sensor: "Sensor", fido: "FIDO", replacement: "Replacement", amr: "AMR", asset: "Asset", consumption: "Profile" };
  const rows = captures.map((c) => {
    if (c.type === "amr") {
      const count = (c.amrShots || []).reduce((n, s) => n + (s.meters || []).length, 0);
      return `
    <tr>
      <td>${c.photo ? `<img src="${c.photo}" class="thumb" />` : "—"}</td>
      <td>${c.position}</td>
      <td>${c.amrAssetType || "AMR"}</td>
      <td>${count} meters linked</td>
      <td>${c.amrSerial || "—"}</td>
      <td>${c.timestamp.date}</td>
      <td>${c.timestamp.time}</td>
      <td>${c.gps ? `${c.gps.lat}, ${c.gps.lng}` : "—"}</td>
      <td>${STATUS_EXPORT_LABEL[c.status] || c.status}</td>
      <td>—</td>
    </tr>`;
    }
    if (c.type === "replacement") {
      return `
    <tr>
      <td>
        ${c.oldMeter?.photo ? `<img src="${c.oldMeter.photo}" class="thumb" title="Old" />` : "—"}
        ${c.newMeter?.photo ? `<img src="${c.newMeter.photo}" class="thumb" title="New" />` : ""}
      </td>
      <td>${c.position}</td>
      <td>Replacement</td>
      <td>OLD: ${c.oldMeter?.reading ? `${c.oldMeter.reading} m³` : "—"}<br/>NEW: ${c.newMeter?.reading ? `${c.newMeter.reading} m³` : "—"}</td>
      <td>OLD: ${c.oldMeter?.serial || "—"}<br/>NEW: ${c.newMeter?.serial || "—"}</td>
      <td>${c.timestamp.date}</td>
      <td>${c.timestamp.time}</td>
      <td>${c.gps ? `${c.gps.lat}, ${c.gps.lng}` : "—"}</td>
      <td>${STATUS_EXPORT_LABEL[c.status] || c.status}</td>
      <td>—</td>
    </tr>`;
    }
    return `
    <tr>
      <td>${(c.photo || c.fido?.photo || c.consumption?.photo) ? `<img src="${c.photo || c.fido?.photo || c.consumption?.photo}" class="thumb" />` : "—"}</td>
      <td>${c.position}</td>
      <td>${c.type === "fido" && c.fido?.sessionType ? c.fido.sessionType : (TYPE_SHORT[c.type] || "Meter")}</td>
      <td>${c.reading ? `${c.reading} m³` : "—"}</td>
      <td>${c.serial || "—"}</td>
      <td>${c.timestamp.date}</td>
      <td>${c.timestamp.time}</td>
      <td>${c.gps ? `${c.gps.lat}, ${c.gps.lng}` : "—"}</td>
      <td>${STATUS_EXPORT_LABEL[c.status] || c.status}</td>
      <td>${c.sensor?.sessionId || "—"}</td>
    </tr>`;
  }).join("");

  const amrDupes = findDuplicateMeters(captures);
  const amrAssets = captures.filter((c) => c.type === "amr" && (c.amrShots || []).length);
  const amrAppendix = amrAssets.length ? `
      <h2 style="font-size:15px; margin:30px 0 6px;">Appendix — AMR Linked Meters</h2>
      ${amrAssets.map((c) => {
        const meters = (c.amrShots || []).flatMap((s) => s.meters || []);
        return `
        <h3 style="font-size:12.5px; margin:16px 0 4px;">${c.amrAssetType || "AMR"} · ${c.position}${c.amrSerial ? ` · ${c.amrSerial}` : ""} <span style="font-weight:normal;color:#5B6570;">(${meters.length} meters)</span></h3>
        <table>
          <thead><tr><th>Meter Serial</th><th>Model</th><th>Signal</th><th>Band</th><th>Note</th></tr></thead>
          <tbody>
            ${meters.map((m) => {
              const b = signalBand(m.signal);
              const key = (m.serial || "").trim().toUpperCase();
              const d = amrDupes[key];
              const dupNote = d
                ? (d.best && d.best.captureId === c.id
                    ? "Strongest — should own this meter"
                    : `Also on ${d.best ? `${d.best.assetType} ${d.best.position}` : "another asset"} (stronger)`)
                : "";
              return `<tr><td>${m.serial || "—"}</td><td>${m.model || "—"}</td><td>${m.signal || "—"}</td><td class="band ${b ? b.label.toLowerCase() : ""}">${b ? b.label : "—"}</td><td class="dupnote">${dupNote}</td></tr>`;
            }).join("")}
          </tbody>
        </table>`;
      }).join("")}
  ` : "";

  win.document.write(`
    <!doctype html><html><head><title>${survey.surveyName || "Meter Survey"} Report</title>
    <style>
      * { box-sizing: border-box; }
      body { font-family: Arial, Helvetica, sans-serif; color: #2B2F33; padding: 32px; margin:0; }
      .brand { display:flex; align-items:center; gap:8px; margin-bottom: 4px; }
      .brand .box { width:20px; height:20px; border-radius:5px; background:#0D86F3; }
      .brand span { font-weight:700; font-size:16px; }
      h1 { font-size: 19px; margin: 18px 0 2px; }
      .meta { color:#5B6570; font-size: 12.5px; margin-bottom: 18px; }
      .stats { display:flex; gap:24px; margin: 18px 0 22px; padding: 14px 18px; background:#F4F7F9; border-radius:10px; }
      .stat b { display:block; font-size:20px; }
      .stat span { font-size:11px; color:#5B6570; }
      table { width:100%; border-collapse: collapse; font-size:12px; }
      th, td { border: 1px solid #DCE3E8; padding: 6px 8px; text-align:left; vertical-align: middle; }
      th { background: #F4F7F9; }
      .thumb { width: 56px; height: 42px; object-fit: cover; border-radius: 4px; display:inline-block; margin: 1px; }
      .footer { margin-top: 22px; font-size: 11px; color:#5B6570; }
      .band { font-weight:700; text-align:center; }
      .band.good { color:#1B9C6E; }
      .band.fair { color:#D98A22; }
      .band.weak { color:#D6485A; }
      .dupnote { font-size:10px; color:#D98A22; }
      @media print { .no-print { display:none; } }
    </style></head><body>
      <div class="brand"><div class="box"></div><span>THYNK-H2O</span></div>
      <h1>${survey.siteName || "Untitled Site"} — Meter Capture Report</h1>
      <div class="meta">
        Survey: ${survey.surveyName || "—"} &nbsp;·&nbsp; Address: ${survey.address || "—"} &nbsp;·&nbsp;
        Technician: ${survey.tech || "—"} &nbsp;·&nbsp; Reviewed by: ${reviewer || "—"} &nbsp;·&nbsp;
        Generated: ${new Date().toLocaleDateString()}
      </div>
      <div class="stats">
        <div class="stat"><b>${counts.total}</b><span>CAPTURED</span></div>
        <div class="stat"><b>${counts.approved}</b><span>APPROVED</span></div>
        <div class="stat"><b>${counts.review}</b><span>NEEDS REVIEW</span></div>
        <div class="stat"><b>${counts.flagged}</b><span>FLAGGED</span></div>
        <div class="stat"><b>${pct}%</b><span>COMPLETE</span></div>
      </div>
      <table>
        <thead><tr><th>Photo</th><th>Position</th><th>Type</th><th>Reading</th><th>Serial</th><th>Date</th><th>Time</th><th>GPS</th><th>Status</th><th>Session ID</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
      ${amrAppendix}
      <div class="footer">Generated by THYNK-H2O Meter Capture & Audit.</div>
      <div class="no-print" style="margin-top:24px;">
        <button onclick="window.print()" style="background:#0D86F3;color:#fff;border:none;padding:10px 18px;border-radius:8px;font-weight:600;cursor:pointer;">Print / Save as PDF</button>
      </div>
    </body></html>
  `);
  win.document.close();
}

function PhotoWell({ position, timestamp, gps, size = "large", photo }) {
  const h = size === "large" ? 190 : 74;
  return (
    <div style={{
      position: "relative", width: "100%", height: h, borderRadius: 12, overflow: "hidden",
      background: photo ? "#000" : `linear-gradient(155deg, ${C.primaryDeep} 0%, ${C.primary} 55%, #57AEFF 100%)`
    }}>
      {photo ? (
        <img src={photo} alt={position || "meter photo"} style={{ width: "100%", height: "100%", objectFit: "cover", display: "block" }} />
      ) : (
        <div style={{
          position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center", opacity: 0.25
        }}>
          <Droplet size={size === "large" ? 64 : 26} color="#fff" fill="#fff" />
        </div>
      )}
      {size === "large" && (
        <div style={{
          position: "absolute", left: 10, bottom: 8, right: 10, display: "flex",
          justifyContent: "space-between", alignItems: "flex-end"
        }}>
          <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 10.5, color: "#fff", textShadow: "0 1px 3px rgba(0,0,0,0.8)" }}>
            {position} · {timestamp?.date} {timestamp?.time}
          </span>
          <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 10, color: "rgba(255,255,255,0.9)", textShadow: "0 1px 3px rgba(0,0,0,0.8)" }}>
            {gps?.lat}, {gps?.lng}
          </span>
        </div>
      )}
    </div>
  );
}

/* ---------- Screen: Setup ---------- */
function SetupScreen({ survey, setSurvey, onStart, onResume, resuming, role, task, username }) {
  const [gpsLocking, setGpsLocking] = useState(false);
  const [previous, setPrevious] = useState([]);
  const [loadingPrev, setLoadingPrev] = useState(true);
  const [confirmDeleteKey, setConfirmDeleteKey] = useState(null);
  const csvInputRef = useRef(null);
  const [csvError, setCsvError] = useState("");
  const amrCsvRef = useRef(null);
  const [amrCsvError, setAmrCsvError] = useState("");

  const handleAmrCSV = (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!file) return;
    setAmrCsvError("");
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const reg = parseAmrRegister(String(reader.result));
        if (!reg) {
          setAmrCsvError("Couldn't read that file. It needs a column with the MUC or repeater serial number.");
          return;
        }
        setSurvey((s) => ({ ...s, amrRegister: reg }));
      } catch (err) {
        setAmrCsvError("Couldn't read that file. Please check it's a CSV.");
      }
    };
    reader.onerror = () => setAmrCsvError("Couldn't read that file.");
    reader.readAsText(file);
  };

  const handleCSV = (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!file) return;
    setCsvError("");
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const map = parsePreviousReadings(String(reader.result));
        const count = Object.keys(map).length;
        if (!count) {
          setCsvError("No readings found. The file needs a position (or serial) column and a reading column.");
          return;
        }
        setSurvey((s) => ({ ...s, previousReadings: map }));
      } catch (err) {
        setCsvError("Couldn't read that file. Please check it's a CSV.");
      }
    };
    reader.onerror = () => setCsvError("Couldn't read that file.");
    reader.readAsText(file);
  };

  // Auto-fill the technician name from the logged-in user
  useEffect(() => {
    if (username && !survey.tech) {
      setSurvey((s) => ({ ...s, tech: username }));
    }
  }, [username]);

  const removeSite = async (e, key) => {
    e.stopPropagation();
    if (confirmDeleteKey !== key) {
      setConfirmDeleteKey(key);
      setTimeout(() => setConfirmDeleteKey((k) => (k === key ? null : k)), 3500);
      return;
    }
    setConfirmDeleteKey(null);
    try {
      await deleteSurveyRecord(key);
      setPrevious((prev) => prev.filter((r) => r.key !== key));
    } catch {
      // leave it in the list if delete failed
    }
  };

  useEffect(() => {
    let cancelled = false;
    listSurveyRecords()
      .then((recs) => {
        if (!cancelled) {
          // Older records used separate "meter" and "replacement" task types
          const normalise = (t) => (!t || t === "meter" || t === "replacement" ? "meterwork" : t);
          setPrevious(recs.filter((r) => normalise(r.taskType) === task));
        }
      })
      .catch(() => {})
      .finally(() => { if (!cancelled) setLoadingPrev(false); });
    return () => { cancelled = true; };
  }, [task]);

  const field = (label, key, placeholder) => (
    <div style={{ marginBottom: 16 }}>
      <label style={{ fontFamily: "'Inter',sans-serif", fontSize: 12, fontWeight: 600, color: C.charcoalSoft, display: "block", marginBottom: 6 }}>
        {label}
      </label>
      <input
        value={survey[key]}
        onChange={(e) => setSurvey((s) => ({ ...s, [key]: e.target.value }))}
        placeholder={placeholder}
        style={{
          width: "100%", boxSizing: "border-box", padding: "11px 13px", borderRadius: 9,
          border: `1.5px solid ${C.line}`, fontFamily: "'Inter',sans-serif", fontSize: 14.5,
          color: C.charcoal, outline: "none"
        }}
        onFocus={(e) => (e.target.style.borderColor = C.primary)}
        onBlur={(e) => (e.target.style.borderColor = C.line)}
      />
    </div>
  );

  const lockGPS = async () => {
    setGpsLocking(true);
    const gps = await getRealGPS();
    setSurvey((s) => ({ ...s, gps: gps || { lat: "Unknown", lng: "Unknown" } }));
    setGpsLocking(false);
  };

  const missing = [
    !survey.siteName && "Estate / Site Name",
    !survey.surveyName && "Survey Name",
    !survey.tech && "Technician Name",
    task === "meterwork" && !survey.meterMode && "Read or Replace",
  ].filter(Boolean);
  const canStart = missing.length === 0;

  return (
    <div style={{ maxWidth: 440, margin: "0 auto", padding: "28px 4px" }}>
      {!loadingPrev && previous.length > 0 && (
        <div style={{ marginBottom: 28 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 10 }}>
            <History size={14} color={C.charcoalSoft} />
            <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 12, fontWeight: 600, color: C.charcoalSoft }}>
              Continue a site you've already started
            </span>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {previous.map((rec) => {
              const unsent = rec.unsentCount || 0;
              const confirming = confirmDeleteKey === rec.key;
              return (
                <div key={rec.key} onClick={() => onResume(rec)} style={{
                  display: "flex", alignItems: "center", justifyContent: "space-between",
                  padding: "11px 13px", borderRadius: 10, border: `1.5px solid ${C.line}`, background: "#fff",
                  cursor: "pointer", textAlign: "left", gap: 10
                }}>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontFamily: "'Inter',sans-serif", fontWeight: 600, fontSize: 13.5, color: C.charcoal }}>
                      {rec.siteName || "Untitled Site"}
                    </div>
                    <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 11, color: C.charcoalSoft, marginTop: 2 }}>
                      {rec.captureCount || 0} meters captured{unsent ? ` · ${unsent} not yet sent` : ""}{rec.tech ? ` · ${rec.tech}` : ""}
                    </div>
                  </div>
                  {role === "office" && (
                    <button
                      onClick={(e) => removeSite(e, rec.key)}
                      style={{
                        border: confirming ? `1.5px solid ${C.flag}` : "none",
                        background: confirming ? C.flagSoft : "none",
                        borderRadius: 8, padding: confirming ? "6px 10px" : 6, cursor: "pointer",
                        display: "flex", alignItems: "center", gap: 5, flexShrink: 0
                      }}>
                      <Trash2 size={14} color={confirming ? C.flag : C.charcoalSoft} />
                      {confirming && (
                        <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 11, fontWeight: 600, color: C.flag }}>
                          Confirm
                        </span>
                      )}
                    </button>
                  )}
                  <ChevronRight size={16} color={C.charcoalSoft} className={resuming ? "spin" : ""} />
                </div>
              );
            })}
          </div>
        </div>
      )}

      <h1 style={{ fontFamily: "'Space Grotesk',sans-serif", fontSize: 22, fontWeight: 700, color: C.charcoal, margin: "0 0 4px" }}>
        New {TASKS[task]?.label || "survey"}
      </h1>
      <p style={{ fontFamily: "'Inter',sans-serif", fontSize: 13.5, color: C.charcoalSoft, margin: "0 0 22px" }}>
        Set up the site once. Everything after this is a photo and a tap.
      </p>
      {field("Estate / Site Name", "siteName", "e.g. Fairview Estate")}
      {field("Site Address", "address", "e.g. 14 Hillcrest Road, Kloof")}
      {field("Survey Name", "surveyName", "e.g. FE-2026-METER-AUDIT")}
      {field("Technician Name", "tech", "e.g. Thula Biyela")}

      <div style={{ marginBottom: 22 }}>
        <label style={{ fontFamily: "'Inter',sans-serif", fontSize: 12, fontWeight: 600, color: C.charcoalSoft, display: "block", marginBottom: 6 }}>
          Site GPS Coordinates
        </label>
        <button onClick={lockGPS} style={{
          width: "100%", display: "flex", alignItems: "center", justifyContent: "space-between",
          padding: "11px 13px", borderRadius: 9, border: `1.5px dashed ${C.line}`, background: C.paper,
          cursor: "pointer", fontFamily: "'IBM Plex Mono',monospace", fontSize: 13, color: C.charcoal
        }}>
          <span style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <MapPin size={15} color={survey.gps ? C.approve : C.charcoalSoft} />
            {gpsLocking ? "Locking signal…" : survey.gps ? `${survey.gps.lat}, ${survey.gps.lng}` : "Tap to capture GPS"}
          </span>
          {gpsLocking && <Loader2 size={15} className="spin" color={C.primary} />}
        </button>
      </div>

      {task === "meterwork" && (
        <div style={{ marginBottom: 22 }}>
          <label style={{ fontFamily: "'Inter',sans-serif", fontSize: 12, fontWeight: 600, color: C.charcoalSoft, display: "block", marginBottom: 8 }}>
            What kind of work is this?
          </label>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {Object.entries(METER_MODES).map(([id, m]) => {
              const active = survey.meterMode === id;
              const Icon = m.icon;
              return (
                <button key={id}
                  onClick={() => setSurvey((s) => ({ ...s, meterMode: id }))}
                  style={{
                    display: "flex", alignItems: "center", gap: 12, textAlign: "left",
                    padding: "13px 14px", borderRadius: 11, cursor: "pointer",
                    border: `1.5px solid ${active ? C.primary : C.line}`,
                    background: active ? "#E7F2FE" : "#fff"
                  }}>
                  <div style={{
                    width: 34, height: 34, borderRadius: 9, flexShrink: 0,
                    background: active ? C.primary : C.paperDeep,
                    display: "flex", alignItems: "center", justifyContent: "center"
                  }}>
                    <Icon size={16} color={active ? "#fff" : C.charcoalSoft} />
                  </div>
                  <div>
                    <div style={{ fontFamily: "'Inter',sans-serif", fontWeight: 700, fontSize: 12.5, color: active ? C.primary : C.charcoal }}>
                      {m.label}
                    </div>
                    <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 11, color: C.charcoalSoft, marginTop: 1 }}>
                      {m.blurb}
                    </div>
                  </div>
                </button>
              );
            })}
          </div>
        </div>
      )}

      {task === "amr" && (
        <div style={{ marginBottom: 22 }}>
          <label style={{ fontFamily: "'Inter',sans-serif", fontSize: 12, fontWeight: 600, color: C.charcoalSoft, display: "block", marginBottom: 6 }}>
            AMR Register (optional)
          </label>
          <input ref={amrCsvRef} type="file" accept=".csv,text/csv" onChange={handleAmrCSV} style={{ display: "none" }} />
          {(() => {
            const reg = survey.amrRegister;
            return (
              <button
                onClick={() => amrCsvRef.current?.click()}
                style={{
                  width: "100%", display: "flex", alignItems: "center", justifyContent: "space-between",
                  padding: "11px 13px", borderRadius: 9,
                  border: `1.5px dashed ${reg ? C.approve : C.line}`,
                  background: reg ? C.approveSoft : C.paper, cursor: "pointer"
                }}>
                <span style={{ display: "flex", alignItems: "center", gap: 8, fontFamily: "'Inter',sans-serif", fontSize: 12.5, fontWeight: 600, color: reg ? C.approve : C.charcoal }}>
                  <FileSpreadsheet size={15} color={reg ? C.approve : C.charcoalSoft} />
                  {reg
                    ? (reg.mode === "assets"
                        ? `${reg.assetCount} assets loaded`
                        : `${reg.assetCount} assets, ${reg.meterCount} meters loaded`)
                    : "Import asset register or meter allocation"}
                </span>
                {reg && <Check size={15} color={C.approve} />}
              </button>
            );
          })()}
          {amrCsvError ? (
            <p style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, color: C.flag, marginTop: 6, marginBottom: 0 }}>{amrCsvError}</p>
          ) : (
            <p style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, color: C.charcoalSoft, marginTop: 6, marginBottom: 0 }}>
              Either a list of MUCs and repeaters to survey \u2014 with description and coordinates \u2014 or a
              meter allocation with one row per meter. Loading a register gives the technician a worklist.
            </p>
          )}
        </div>
      )}

      {task === "meterwork" && survey.meterMode === "read" && (
        <div style={{ marginBottom: 22 }}>
          <label style={{ fontFamily: "'Inter',sans-serif", fontSize: 12, fontWeight: 600, color: C.charcoalSoft, display: "block", marginBottom: 6 }}>
            Previous Readings (optional)
          </label>
          <input
            ref={csvInputRef}
            type="file"
            accept=".csv,text/csv"
            onChange={handleCSV}
            style={{ display: "none" }}
          />
          {(() => {
            const count = Object.keys(survey.previousReadings || {}).length;
            return (
              <button
                onClick={() => csvInputRef.current?.click()}
                style={{
                  width: "100%", display: "flex", alignItems: "center", justifyContent: "space-between",
                  padding: "11px 13px", borderRadius: 9,
                  border: `1.5px dashed ${count ? C.approve : C.line}`,
                  background: count ? C.approveSoft : C.paper, cursor: "pointer"
                }}>
                <span style={{ display: "flex", alignItems: "center", gap: 8, fontFamily: "'Inter',sans-serif", fontSize: 12.5, fontWeight: 600, color: count ? C.approve : C.charcoal }}>
                  <FileSpreadsheet size={15} color={count ? C.approve : C.charcoalSoft} />
                  {count ? `${count} previous readings loaded` : "Import previous readings CSV"}
                </span>
                {count > 0 && <Check size={15} color={C.approve} />}
              </button>
            );
          })()}
          {csvError ? (
            <p style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, color: C.flag, marginTop: 6, marginBottom: 0 }}>{csvError}</p>
          ) : (
            <p style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, color: C.charcoalSoft, marginTop: 6, marginBottom: 0 }}>
              A CSV with columns for position (or serial) and reading. The technician will see the last reading at each meter.
            </p>
          )}
        </div>
      )}

      <button
        disabled={!canStart}
        onClick={onStart}
        style={{
          width: "100%", padding: "13px", borderRadius: 10, border: "none", cursor: canStart ? "pointer" : "not-allowed",
          background: canStart ? C.primary : C.line, color: "#fff",
          fontFamily: "'Inter',sans-serif", fontWeight: 600, fontSize: 15,
          display: "flex", alignItems: "center", justifyContent: "center", gap: 8
        }}>
        Start Survey <ArrowRight size={16} />
      </button>
      {!canStart && (
        <p style={{
          fontFamily: "'Inter',sans-serif", fontSize: 11.5, color: C.review,
          textAlign: "center", marginTop: 9, marginBottom: 0
        }}>
          Still needed: {missing.join(", ")}
        </p>
      )}
    </div>
  );
}

/* ---------- Screen: Field Capture ---------- */
function CaptureScreen({ survey, captures, setCaptures, setScreen, task }) {
  const [position, setPosition] = useState("");
  const [stage, setStage] = useState("idle"); // idle | scanning | result
  const [pending, setPending] = useState(null);
  const [shake, setShake] = useState(false);
  const [captureError, setCaptureError] = useState("");
  const [confirmDiscard, setConfirmDiscard] = useState(false);
  const [justSaved, setJustSaved] = useState(false);
  const [showException, setShowException] = useState(false);
  const [fidoMode, setFidoMode] = useState("deploy");
  const [pendingBug, setPendingBug] = useState("");
  const [zoomShot, setZoomShot] = useState(null);
  const [pendingTestMethod, setPendingTestMethod] = useState("");
  const [pendingRegAsset, setPendingRegAsset] = useState(null);

  const previousForPosition = position.trim()
    ? lookupPrevious(survey.previousReadings, position, "")
    : null;

  const openException = () => {
    if (!position.trim()) {
      setShake(true);
      inputRef.current?.focus();
      setTimeout(() => setShake(false), 420);
      return;
    }
    setCaptureError("");
    setShowException(true);
  };

  const saveException = async (reason) => {
    setShowException(false);
    const gps = await getRealGPS();
    setCaptures((c) => [{
      id: Date.now(),
      type: "exception",
      position: position.trim(),
      reading: "",
      serial: "",
      confidence: 0,
      gps: gps || survey.gps || { lat: "Unknown", lng: "Unknown" },
      timestamp: nowStamp(),
      status: "flagged",
      tech: survey.tech,
      sentAt: null,
      photo: null,
      exceptionReason: reason,
      needsRevisit: true,
    }, ...c]);
    setPosition("");
    setJustSaved(true);
    setTimeout(() => setJustSaved(false), 2200);
  };
  const inputRef = useRef(null);
  const fileInputRef = useRef(null); // camera-first (meter, sensor install)
  const galleryInputRef = useRef(null); // gallery-first (consumption screenshots)

  const pickTypeRef = useRef("meter"); // which kind of capture the next photo is for
  const [chooseFor, setChooseFor] = useState(null); // when set, show take-vs-upload chooser for this type

  const openPicker = (type) => {
    if (stage === "scanning") return;
    if (!pending && !position.trim()) {
      setShake(true);
      inputRef.current?.focus();
      setTimeout(() => setShake(false), 420);
      return;
    }
    setCaptureError("");
    // Screenshots come from the gallery.
    if (type === "fido" || type === "amrShot" || type === "fido2Session" || type === "fido2Result" || type === "waypointTest") {
      pickTypeRef.current = type;
      galleryInputRef.current?.click();
      return;
    }
    // Meters, sensor installs and asset photos: choose camera or gallery.
    setChooseFor(type);
  };

  const chooseSource = (source) => {
    if (!chooseFor) return;
    pickTypeRef.current = chooseFor;
    setChooseFor(null);
    if (source === "camera") {
      fileInputRef.current?.click();
    } else {
      galleryInputRef.current?.click();
    }
  };

  const handleFileSelected = async (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = ""; // allow selecting the same photo again next time
    if (!file) return; // user cancelled the camera / file picker
    const type = pickTypeRef.current;
    setStage("scanning");
    try {
      if (type === "meter") {
        // Higher-res copy for AI reading accuracy; GPS captured at this exact moment
        const [photo, gps] = await Promise.all([loadDownscaledPhoto(file, 1100), getRealGPS()]);
        let reading = "";
        let serial = "";
        let confidence = 0;
        let aiFailed = false;
        try {
          const extracted = await extractWithVision(photo, METER_PROMPT);
          reading = extracted.reading || "";
          serial = extracted.serial || "";
          confidence = Math.max(0, Math.min(100, Number(extracted.confidence) || 0));
        } catch (visionErr) {
          aiFailed = true;
        }
        setPending({
          id: Date.now(),
          type: "meter",
          position: position.trim(),
          reading,
          serial,
          confidence,
          gps: gps || survey.gps || { lat: "Unknown", lng: "Unknown" },
          timestamp: nowStamp(),
          status: confidence < 85 ? "needs_review" : "captured",
          tech: survey.tech,
          sentAt: null,
          photo,
          sensor: null,
          consumption: null,
        });
        if (aiFailed) {
          setCaptureError("AI couldn't read this meter — please type the reading and serial manually.");
        } else if (!reading) {
          setCaptureError("The register wasn't clearly legible — please check the photo or type the reading manually.");
        }
      } else if (type === "fido2Deploy") {
        // FIDO 2 deployment photo — where the bug is physically placed
        const [photo, gps] = await Promise.all([loadDownscaledPhoto(file, 1100), getRealGPS()]);
        setPending((p) => ({
          ...(p || {}),
          id: p?.id || Date.now(),
          type: "fido2_deploy",
          position: position.trim(),
          reading: "",
          serial: "",
          confidence: 0,
          gps: gps || survey.gps || { lat: "Unknown", lng: "Unknown" },
          timestamp: p?.timestamp || nowStamp(),
          status: "needs_review",
          tech: survey.tech,
          sentAt: null,
          photo,
          bugSerial: p?.bugSerial || "",
          deployAsset: p?.deployAsset || "",
          fidoMode: p?.fidoMode || "",
          sessionId: p?.sessionId || "",
          sessionShot: p?.sessionShot || null,
          pipeMaterial: p?.pipeMaterial || "",
        }));
      } else if (type === "fido2Session") {
        // FIDO 2 session screenshot — AI reads the session ID
        const photo = await loadDownscaledPhoto(file, 1100);
        let ex = {};
        let aiFailed = false;
        try {
          ex = await extractWithVision(photo, FIDO_DEPLOY_PROMPT);
        } catch (e) { aiFailed = true; }
        setPending((p) => p ? {
          ...p,
          sessionShot: { photo, ...ex },
          sessionId: p.sessionId || ex.sessionId || "",
          bugSerial: p.bugSerial || ex.bugSerial || "",
        } : p);
        if (aiFailed) setCaptureError("Couldn't read that screenshot — enter the session ID manually.");
      } else if (type === "waypointTest") {
        // A test carried out within a waypoint
        const photo = await loadDownscaledPhoto(file, 1200);
        let ex = {};
        if (pendingTestMethod === "CORRELATION") {
          try { ex = await extractWithVision(photo, CORRELATION_PROMPT); } catch (e) { /* manual entry */ }
        }
        setPending((p) => p ? {
          ...p,
          tests: [...(p.tests || []), {
            method: pendingTestMethod,
            photo,
            distance: ex.distance || "",
            sensorSpacing: ex.sensorSpacing || "",
            note: ex.notes || "",
            at: nowStamp(),
          }],
        } : p);
      } else if (type === "waypointPinpoint") {
        // Beacon placed at the suspected leak position
        const [photo, gps] = await Promise.all([loadDownscaledPhoto(file, 1100), getRealGPS()]);
        setPending((p) => p ? {
          ...p,
          pinpointPhoto: photo,
          pinpointGps: gps || p.gps,
        } : p);
      } else if (type === "fido2Result") {
        // FIDO 2 results graph at retrieval — AI reads the session ID for cross-check
        const [photo, gps] = await Promise.all([loadDownscaledPhoto(file, 1100), getRealGPS()]);
        let ex = {};
        let aiFailed = false;
        try {
          ex = await extractWithVision(photo, FIDO_RESULT_PROMPT);
        } catch (e) { aiFailed = true; }
        setPending((p) => p ? {
          ...p,
          gps: p.gps && p.gps.lat !== "Unknown" ? p.gps : (gps || p.gps),
          resultShot: { photo, ...ex },
          resultSessionId: ex.sessionId || "",
        } : p);
        if (aiFailed) setCaptureError("Couldn't read that graph — enter the session ID manually to confirm the match.");
      } else if (type === "asset") {
        // Asset mapping — photo, GPS, AI-read serial and service date where visible
        const [photo, gps] = await Promise.all([loadDownscaledPhoto(file, 1100), getRealGPS()]);
        let serial = "";
        let lastService = "";
        try {
          const extracted = await extractWithVision(photo, `This is a photograph of an estate infrastructure asset (a meter, valve, kiosk, hydrant, fire extinguisher, hose reel, chamber, telemetry device or similar). Respond with ONLY a JSON object: {"serial": string|null, "lastServiceDate": string|null}. Return the serial or device number if one is clearly legible on the asset, otherwise null. If a service label, inspection tag or certificate showing a service date is clearly legible, return that date as YYYY-MM-DD, otherwise null. Do not guess either value.`);
          serial = extracted.serial || "";
          lastService = extracted.lastServiceDate || "";
        } catch (visionErr) {
          // silent — serial is optional and can be typed
        }
        setPending({
          id: Date.now(),
          type: "asset",
          position: position.trim(),
          reading: "",
          serial: "",
          confidence: 0,
          gps: gps || survey.gps || { lat: "Unknown", lng: "Unknown" },
          timestamp: nowStamp(),
          status: "needs_review",
          tech: survey.tech,
          sentAt: null,
          photo,
          sensor: null,
          consumption: null,
          fido: null,
          oldMeter: null,
          newMeter: null,
          assetCategory: null,
          assetType: null,
          assetSerial: serial,
          assetDescription: "",
          assetZone: "",
          assetErf: "",
          assetAccessNotes: "",
          assetLastService: lastService,
          assetLastCartridge: "",
          assetLastFlush: "",
          assetCartridgeCondition: "",
          assetFilterSpec: "",
        });
      } else if (type === "amrAsset") {
        // AMR asset photo (MUC or repeater) — AI reads serial if legible, GPS at capture
        const [photo, gps] = await Promise.all([loadDownscaledPhoto(file, 1100), getRealGPS()]);
        let serial = "";
        try {
          const extracted = await extractWithVision(photo, `This is a photograph of AMR telemetry equipment (a MUC, concentrator, or signal repeater) mounted in the field. Respond with ONLY a JSON object: {"serial": string|null}. Return the device serial number if one is clearly legible on a label or the device body, otherwise null. Do not guess.`);
          serial = extracted.serial || "";
        } catch (visionErr) {
          // silent — serial can be typed manually
        }
        setPending((p) => {
          const base = p || {
            id: Date.now(),
            type: "amr",
            position: position.trim(),
            reading: "",
            serial: "",
            confidence: 0,
            gps: gps || survey.gps || { lat: "Unknown", lng: "Unknown" },
            timestamp: nowStamp(),
            status: "needs_review",
            tech: survey.tech,
            sentAt: null,
            photo: null,
            sensor: null,
            consumption: null,
            fido: null,
            oldMeter: null,
            newMeter: null,
            amrAssetType: pendingRegAsset?.assetType === "REPEATER" ? "Repeater" : "MUC",
            amrSerial: pendingRegAsset?.serial || "",
            registerGps: pendingRegAsset?.gps || null,
            noMetersAllocated: false,
            noMetersNote: "",
            amrShots: [],
          };
          return {
            ...base,
            photo,
            gps: base.gps && base.gps.lat !== "Unknown" ? base.gps : (gps || base.gps),
            amrSerial: base.amrSerial || pendingRegAsset?.serial || serial,
          };
        });
        if (!serial) {
          setCaptureError("Serial not readable from the photo — please type it in below.");
        }
      } else if (type === "amrShot") {
        // Handheld reader screenshot — AI extracts the meter list
        const photo = await loadDownscaledPhoto(file, 1200);
        let extracted = { meters: [], concentratorId: null };
        let aiFailed = false;
        try {
          extracted = await extractWithVision(photo, AMR_SCREENSHOT_PROMPT);
        } catch (visionErr) {
          aiFailed = true;
        }
        setPending((p) => {
          if (!p) return p;
          const shots = [...(p.amrShots || []), {
            photo,
            meters: extracted.meters || [],
            concentratorId: extracted.concentratorId || null,
            readingDate: extracted.readingDate || null,
          }];
          return { ...p, amrShots: shots };
        });
        if (aiFailed) {
          setCaptureError("Couldn't read that screenshot — try a clearer capture.");
        }
      } else if (type === "oldMeter" || type === "newMeter") {
        // Meter replacement — old meter out, new meter in, sharing one GPS at the position
        const [photo, gps] = await Promise.all([loadDownscaledPhoto(file, 1100), getRealGPS()]);
        let reading = "";
        let serial = "";
        let confidence = 0;
        let aiFailed = false;
        try {
          const extracted = await extractWithVision(photo, METER_PROMPT);
          reading = extracted.reading || "";
          serial = extracted.serial || "";
          confidence = Math.max(0, Math.min(100, Number(extracted.confidence) || 0));
        } catch (visionErr) {
          aiFailed = true;
        }
        const side = { photo, reading, serial, confidence };
        setPending((p) => {
          const base = p || {
            id: Date.now(),
            type: "replacement",
            position: position.trim(),
            reading: "",
            serial: "",
            confidence: 0,
            gps: gps || survey.gps || { lat: "Unknown", lng: "Unknown" },
            timestamp: nowStamp(),
            status: "needs_review",
            tech: survey.tech,
            sentAt: null,
            photo: null,
            sensor: null,
            consumption: null,
            fido: null,
            oldMeter: null,
            newMeter: null,
            meterType: "",
            conditionFlags: [],
            conditionNote: "",
          };
          return {
            ...base,
            // keep the first GPS captured at this position for both meters
            gps: base.gps && base.gps.lat !== "Unknown" ? base.gps : (gps || base.gps),
            [type]: side,
          };
        });
        if (aiFailed) {
          setCaptureError("AI couldn't read this meter — type the serial and reading manually.");
        } else if (!reading || !serial) {
          setCaptureError("Some details weren't clearly legible — please check and complete them.");
        }
      } else if (type === "amrAsset") {
        // AMR asset photo (MUC or repeater) — AI reads serial if legible, GPS at capture
        const [photo, gps] = await Promise.all([loadDownscaledPhoto(file, 1100), getRealGPS()]);
        let serial = "";
        try {
          const extracted = await extractWithVision(photo, `This is a photograph of AMR telemetry equipment (a MUC, concentrator, or signal repeater) mounted in the field. Respond with ONLY a JSON object: {"serial": string|null}. Return the device serial number if one is clearly legible on a label or the device body, otherwise null. Do not guess.`);
          serial = extracted.serial || "";
        } catch (visionErr) {
          // silent — serial can be typed manually
        }
        setPending((p) => {
          const base = p || {
            id: Date.now(),
            type: "amr",
            position: position.trim(),
            reading: "",
            serial: "",
            confidence: 0,
            gps: gps || survey.gps || { lat: "Unknown", lng: "Unknown" },
            timestamp: nowStamp(),
            status: "needs_review",
            tech: survey.tech,
            sentAt: null,
            photo: null,
            sensor: null,
            consumption: null,
            fido: null,
            oldMeter: null,
            newMeter: null,
            amrAssetType: "MUC",
            amrSerial: "",
            amrShots: [],
          };
          return {
            ...base,
            photo,
            gps: base.gps && base.gps.lat !== "Unknown" ? base.gps : (gps || base.gps),
            amrSerial: base.amrSerial || serial,
          };
        });
        if (!serial) {
          setCaptureError("Serial not readable from the photo — please type it in below.");
        }
      } else if (type === "amrShot") {
        // Handheld reader screenshot — AI extracts the meter list
        const photo = await loadDownscaledPhoto(file, 1200);
        let extracted = { meters: [], concentratorId: null };
        let aiFailed = false;
        try {
          extracted = await extractWithVision(photo, AMR_SCREENSHOT_PROMPT);
        } catch (visionErr) {
          aiFailed = true;
        }
        setPending((p) => {
          if (!p) return p;
          const shots = [...(p.amrShots || []), {
            photo,
            meters: extracted.meters || [],
            concentratorId: extracted.concentratorId || null,
            readingDate: extracted.readingDate || null,
          }];
          return { ...p, amrShots: shots };
        });
        if (aiFailed) {
          setCaptureError("Couldn't read that screenshot — try a clearer capture.");
        }
      } else if (type === "oldMeter" || type === "newMeter") {
        // Meter replacement — old meter out, new meter in, sharing one GPS at the position
        const [photo, gps] = await Promise.all([loadDownscaledPhoto(file, 1100), getRealGPS()]);
        let reading = "";
        let serial = "";
        let confidence = 0;
        let aiFailed = false;
        try {
          const extracted = await extractWithVision(photo, METER_PROMPT);
          reading = extracted.reading || "";
          serial = extracted.serial || "";
          confidence = Math.max(0, Math.min(100, Number(extracted.confidence) || 0));
        } catch (visionErr) {
          aiFailed = true;
        }
        const side = { photo, reading, serial, confidence };
        setPending((p) => {
          const base = p || {
            id: Date.now(),
            type: "replacement",
            position: position.trim(),
            reading: "",
            serial: "",
            confidence: 0,
            gps: gps || survey.gps || { lat: "Unknown", lng: "Unknown" },
            timestamp: nowStamp(),
            status: "needs_review",
            tech: survey.tech,
            sentAt: null,
            photo: null,
            sensor: null,
            consumption: null,
            fido: null,
            oldMeter: null,
            newMeter: null,
          };
          return {
            ...base,
            // keep the first GPS captured at this position for both meters
            gps: base.gps && base.gps.lat !== "Unknown" ? base.gps : (gps || base.gps),
            [type]: side,
          };
        });
        if (aiFailed) {
          setCaptureError("AI couldn't read this meter — type the serial and reading manually.");
        } else if (!reading || !serial) {
          setCaptureError("Some details weren't clearly legible — please check and complete them.");
        }
      } else if (type === "sensor") {
        // Photo of where the sensor is deployed — no AI needed on the installation photo itself
        const [photo, gps] = await Promise.all([loadDownscaledPhoto(file, 1100), getRealGPS()]);
        setPending({
          id: Date.now(),
          type: "sensor",
          position: position.trim(),
          reading: "",
          serial: "",
          confidence: 0,
          gps: gps || survey.gps || { lat: "Unknown", lng: "Unknown" },
          timestamp: nowStamp(),
          status: "captured",
          tech: survey.tech,
          sentAt: null,
          photo,
          sensor: null,
          consumption: null,
        });
      } else {
        // FIDO feedback — a screenshot from the gallery; GPS captured at upload, session type chosen next
        const [photo, gps] = await Promise.all([loadDownscaledPhoto(file, 1000), getRealGPS()]);
        setPending({
          id: Date.now(),
          type: "fido",
          position: position.trim(),
          reading: "",
          serial: "",
          confidence: 0,
          gps: gps || survey.gps || { lat: "Unknown", lng: "Unknown" },
          timestamp: nowStamp(),
          status: "captured",
          tech: survey.tech,
          sentAt: null,
          photo: null,
          sensor: null,
          consumption: null,
          fido: { photo, gps, sessionType: null },
        });
      }
      setStage("result");
    } catch (err) {
      setCaptureError("Couldn't read that photo. Please try again.");
      setStage(pending ? "result" : "idle");
    }
  };

  const save = () => {
    // Freeze the register comparison with the capture so review and exports can use it
    if (pending.type === "amr" && survey.amrRegister) {
      const cmp = compareToRegister(survey.amrRegister, pending.amrSerial, (pending.amrShots || []).flatMap((s) => s.meters || []));
      if (cmp && !cmp.unknownAsset) {
        pending.expectedCount = cmp.expected.length;
        pending.matchedCount = cmp.matched.length;
        pending.missingMeters = cmp.missing;
        pending.extraMeters = cmp.extra;
      }
    }
    const flagged = (pending.conditionFlags || []).length > 0;
    let record = flagged && pending.status !== "flagged"
      ? { ...pending, status: "needs_review" }
      : pending;
    // A confirmed leak, or a session ID that doesn't match, always goes to review
    if (record.type === "amr" && record.noMetersAllocated) {
      record = { ...record, status: "needs_review" };
    }
    if (record.type === "amr" && (record.missingMeters || []).length > 0) {
      record = { ...record, status: "needs_review" };
    }
    if (record.type === "fido2_waypoint" && record.outcome === "LEAK CONFIRMED") {
      record = { ...record, status: "needs_review" };
    }
    if (record.type === "fido2_retrieve") {
      const a = (record.sessionId || "").trim().toLowerCase();
      const g = (record.resultSessionId || "").trim().toLowerCase();
      const mismatch = a && g && a !== g;
      if (mismatch || record.outcome === "CONFIRMED LEAK" || record.outcome === "SUSPECTED LEAK") {
        record = { ...record, status: "needs_review" };
      }
    }
    setCaptures((c) => [record, ...c]);
    setPending(null);
    setPosition("");
    setCaptureError("");
    setStage("idle");
    setPendingBug("");
    setPendingRegAsset(null);
    setJustSaved(true);
    setTimeout(() => setJustSaved(false), 2200);
    setTimeout(() => inputRef.current?.focus(), 50);
  };

  const editField = (key, val) => setPending((p) => ({ ...p, [key]: val }));

  return (
    <div style={{ display: "flex", justifyContent: "center", padding: "20px 4px" }}>
      {zoomShot && (
        <div
          onClick={() => setZoomShot(null)}
          style={{
            position: "fixed", inset: 0, background: "rgba(20,22,24,0.94)", zIndex: 200,
            display: "flex", alignItems: "center", justifyContent: "center", padding: 14, cursor: "zoom-out"
          }}>
          <img src={zoomShot} alt="Graph full size" style={{ maxWidth: "100%", maxHeight: "100%", objectFit: "contain", borderRadius: 8 }} />
        </div>
      )}
      <div style={{
        width: 340, borderRadius: 28, border: `8px solid ${C.charcoal}`, background: "#fff",
        boxShadow: "0 20px 40px -18px rgba(43,47,51,0.35)", overflow: "hidden"
      }}>
        {/* phone header */}
        <div style={{ background: C.charcoal, padding: "10px 16px 14px" }}>
          <Logo dark />
          <div style={{ marginTop: 8 }}>
            <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 13, fontWeight: 600, color: "#fff" }}>
              {survey.siteName || "Untitled Site"}
            </div>
            <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 10.5, color: "rgba(255,255,255,0.6)" }}>
              {survey.surveyName} · {survey.tech}
            </div>
          </div>
        </div>

        <input
          ref={fileInputRef}
          type="file"
          accept="image/*"
          capture="environment"
          onChange={handleFileSelected}
          style={{ display: "none" }}
        />
        <input
          ref={galleryInputRef}
          type="file"
          accept="image/*"
          onChange={handleFileSelected}
          style={{ display: "none" }}
        />

        <div style={{ padding: 16, minHeight: 380, position: "relative" }}>
          {showException && (
            <div
              onClick={() => setShowException(false)}
              style={{
                position: "absolute", inset: 0, zIndex: 6, background: "rgba(43,47,51,0.45)",
                display: "flex", alignItems: "flex-end"
              }}>
              <div onClick={(e) => e.stopPropagation()} style={{
                width: "100%", background: "#fff", borderRadius: "16px 16px 0 0", padding: "14px 16px 18px",
                maxHeight: "100%", overflowY: "auto"
              }}>
                <div style={{ width: 34, height: 4, borderRadius: 99, background: C.line, margin: "0 auto 12px" }} />
                <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 12.5, fontWeight: 700, color: C.charcoal, marginBottom: 3, textAlign: "center" }}>
                  Why can't this meter be read?
                </div>
                <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, color: C.charcoalSoft, marginBottom: 12, textAlign: "center" }}>
                  {position.trim()} will be flagged for a revisit.
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                  {EXCEPTION_REASONS.map((r) => (
                    <button key={r} onClick={() => saveException(r)} style={{
                      padding: "11px 12px", borderRadius: 9, cursor: "pointer", textAlign: "left",
                      border: `1.5px solid ${C.line}`, background: "#fff",
                      fontFamily: "'Inter',sans-serif", fontSize: 12, fontWeight: 600, color: C.charcoal
                    }}>
                      {r}
                    </button>
                  ))}
                </div>
                <button onClick={() => setShowException(false)} style={{
                  width: "100%", padding: "10px", border: "none", background: "none", cursor: "pointer",
                  color: C.charcoalSoft, fontFamily: "'Inter',sans-serif", fontWeight: 600, fontSize: 12, marginTop: 6
                }}>
                  Cancel
                </button>
              </div>
            </div>
          )}
          {chooseFor && (
            <div
              onClick={() => setChooseFor(null)}
              style={{
                position: "absolute", inset: 0, zIndex: 5, background: "rgba(43,47,51,0.45)",
                display: "flex", alignItems: "flex-end"
              }}>
              <div onClick={(e) => e.stopPropagation()} style={{
                width: "100%", background: "#fff", borderRadius: "16px 16px 0 0", padding: "14px 16px 18px"
              }}>
                <div style={{ width: 34, height: 4, borderRadius: 99, background: C.line, margin: "0 auto 12px" }} />
                <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 12.5, fontWeight: 700, color: C.charcoal, marginBottom: 10, textAlign: "center" }}>
                  {chooseFor === "meter" ? "Meter photo" : chooseFor === "asset" ? "Asset photo" : "Installation photo"}
                </div>
                <button onClick={() => chooseSource("camera")} style={{
                  width: "100%", padding: "13px", borderRadius: 10, border: "none", cursor: "pointer",
                  background: C.primary, color: "#fff", fontFamily: "'Inter',sans-serif", fontWeight: 600, fontSize: 13.5,
                  display: "flex", alignItems: "center", justifyContent: "center", gap: 8, marginBottom: 8
                }}>
                  <Camera size={16} /> Take Photo
                </button>
                <button onClick={() => chooseSource("gallery")} style={{
                  width: "100%", padding: "13px", borderRadius: 10, border: `1.5px solid ${C.line}`, cursor: "pointer",
                  background: "#fff", color: C.charcoal, fontFamily: "'Inter',sans-serif", fontWeight: 600, fontSize: 13.5,
                  display: "flex", alignItems: "center", justifyContent: "center", gap: 8
                }}>
                  <History size={16} /> Upload from Gallery
                </button>
                <button onClick={() => setChooseFor(null)} style={{
                  width: "100%", padding: "10px", border: "none", background: "none", cursor: "pointer",
                  color: C.charcoalSoft, fontFamily: "'Inter',sans-serif", fontWeight: 600, fontSize: 12, marginTop: 4
                }}>
                  Cancel
                </button>
              </div>
            </div>
          )}
          {stage !== "result" && (
            <>
              <label style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, fontWeight: 600, color: C.charcoalSoft, display: "block", marginBottom: 6 }}>
                Position / Unit Number
              </label>
              <input
                ref={inputRef}
                autoFocus
                value={position}
                onChange={(e) => setPosition(e.target.value)}
                placeholder="e.g. 14B"
                style={{
                  width: "100%", boxSizing: "border-box", padding: "10px 12px", borderRadius: 9,
                  border: `1.5px solid ${C.line}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 15,
                  color: C.charcoal, outline: "none", marginBottom: 14
                }}
              />

              <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, fontWeight: 600, color: C.charcoalSoft, marginBottom: 8 }}>
                What are you capturing at this point?
              </div>

              {task === "fido2" ? (() => {
                const bugStatus = buildBugStatus(captures);
                const deployed = Object.values(bugStatus).filter((b) => b.state === "DEPLOYED");
                const available = Object.values(bugStatus).filter((b) => b.state !== "DEPLOYED");
                return (
                  <>
                    <div style={{ display: "flex", gap: 6, background: C.paperDeep, padding: 4, borderRadius: 10, marginBottom: 12 }}>
                      {[
                        { id: "deploy", label: `Deploy (${available.length})` },
                        { id: "retrieve", label: `Retrieve (${deployed.length})` },
                        { id: "waypoints", label: `Waypoints (${captures.filter((c) => c.type === "fido2_waypoint").length})` },
                        { id: "board", label: "Board" },
                      ].map((m) => (
                        <button key={m.id} onClick={() => setFidoMode(m.id)} style={{
                          flex: 1, border: "none", cursor: "pointer", padding: "8px 6px", borderRadius: 7,
                          fontFamily: "'Inter',sans-serif", fontSize: 9.5, fontWeight: 700,
                          background: fidoMode === m.id ? "#fff" : "transparent",
                          color: fidoMode === m.id ? C.charcoal : C.charcoalSoft,
                          boxShadow: fidoMode === m.id ? "0 1px 3px rgba(43,47,51,0.12)" : "none"
                        }}>{m.label}</button>
                      ))}
                    </div>

                    {fidoMode === "waypoints" && (() => {
                      const waypoints = captures.filter((c) => c.type === "fido2_waypoint");
                      const flagged = captures.filter((c) =>
                        c.type === "fido2_retrieve" &&
                        (c.outcome === "SUSPECTED LEAK" || c.outcome === "CONFIRMED LEAK"));
                      const investigated = new Set(waypoints.map((w) => w.linkedBug).filter(Boolean));
                      const awaiting = flagged.filter((f) => !investigated.has(f.bugSerial));
                      return (
                        <>
                          {awaiting.length > 0 && (
                            <>
                              <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.review, marginBottom: 6 }}>
                                FLAGGED \u2014 NEEDS INVESTIGATION
                              </div>
                              <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 14 }}>
                                {awaiting.map((f) => (
                                  <button key={f.id}
                                    onClick={() => {
                                      setPending({
                                        id: Date.now(),
                                        type: "fido2_waypoint",
                                        waypointRef: nextWaypointRef(captures),
                                        position: f.position,
                                        reading: "", serial: "", confidence: 0,
                                        gps: f.gps,
                                        timestamp: nowStamp(),
                                        status: "needs_review",
                                        tech: survey.tech,
                                        sentAt: null,
                                        photo: null,
                                        linkedBug: f.bugSerial,
                                        triggerOutcome: f.outcome,
                                        pipeMaterial: f.pipeMaterial || "",
                                        tests: [],
                                        pinpointPhoto: null,
                                        pinpointGps: null,
                                        outcome: "ONGOING",
                                        outcomeNote: "",
                                      });
                                      setStage("result");
                                    }}
                                    style={{
                                      display: "flex", alignItems: "center", justifyContent: "space-between",
                                      padding: "12px 13px", borderRadius: 10, cursor: "pointer",
                                      border: `1.5px solid ${C.review}`, background: C.reviewSoft, textAlign: "left"
                                    }}>
                                    <div>
                                      <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 12.5, fontWeight: 700, color: C.charcoal }}>
                                        {f.bugSerial}
                                      </div>
                                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, color: C.charcoalSoft, marginTop: 2 }}>
                                        {f.outcome} \u00b7 {f.position}
                                      </div>
                                    </div>
                                    <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, fontWeight: 700, color: C.review }}>
                                      OPEN WAYPOINT
                                    </span>
                                  </button>
                                ))}
                              </div>
                            </>
                          )}

                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                            <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.charcoalSoft }}>
                              WAYPOINTS
                            </span>
                            <button
                              onClick={() => {
                                if (!position.trim()) { setShake(true); inputRef.current?.focus(); setTimeout(() => setShake(false), 420); return; }
                                getRealGPS().then((gps) => {
                                  setPending({
                                    id: Date.now(),
                                    type: "fido2_waypoint",
                                    waypointRef: nextWaypointRef(captures),
                                    position: position.trim(),
                                    reading: "", serial: "", confidence: 0,
                                    gps: gps || survey.gps || { lat: "Unknown", lng: "Unknown" },
                                    timestamp: nowStamp(),
                                    status: "needs_review",
                                    tech: survey.tech,
                                    sentAt: null,
                                    photo: null,
                                    linkedBug: "",
                                    triggerOutcome: "",
                                    pipeMaterial: "",
                                    tests: [],
                                    pinpointPhoto: null,
                                    pinpointGps: null,
                                    outcome: "ONGOING",
                                    outcomeNote: "",
                                  });
                                  setStage("result");
                                });
                              }}
                              style={{
                                border: "none", background: "none", color: C.primary, cursor: "pointer",
                                fontFamily: "'Inter',sans-serif", fontWeight: 700, fontSize: 11
                              }}>
                              + New waypoint
                            </button>
                          </div>

                          {waypoints.length === 0 ? (
                            <p style={{ fontFamily: "'Inter',sans-serif", fontSize: 11, color: C.charcoalSoft, textAlign: "center", padding: "16px 0" }}>
                              No waypoints yet.
                            </p>
                          ) : (
                            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                              {waypoints.map((w) => {
                                const st = WAYPOINT_OUTCOMES[w.outcome] || WAYPOINT_OUTCOMES.ONGOING;
                                return (
                                  <button key={w.id}
                                    onClick={() => {
                                      setCaptures((cs) => cs.filter((c) => c.id !== w.id));
                                      setPending(w);
                                      setStage("result");
                                    }}
                                    style={{
                                      display: "flex", alignItems: "center", justifyContent: "space-between",
                                      padding: "11px 13px", borderRadius: 10, cursor: "pointer",
                                      border: `1.5px solid ${C.line}`, background: "#fff", textAlign: "left"
                                    }}>
                                    <div>
                                      <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 12.5, fontWeight: 700, color: C.charcoal }}>
                                        {w.waypointRef}
                                      </div>
                                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, color: C.charcoalSoft, marginTop: 2 }}>
                                        {w.position} \u00b7 {(w.tests || []).length} test{(w.tests || []).length === 1 ? "" : "s"}
                                        {w.linkedBug ? ` \u00b7 from ${w.linkedBug}` : ""}
                                      </div>
                                    </div>
                                    <span style={{
                                      fontFamily: "'Inter',sans-serif", fontSize: 9, fontWeight: 700,
                                      color: st.colour, background: st.soft, padding: "3px 9px", borderRadius: 999
                                    }}>{w.outcome}</span>
                                  </button>
                                );
                              })}
                            </div>
                          )}
                        </>
                      );
                    })()}

                    {fidoMode === "board" && (
                      <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
                        {Object.values(bugStatus).map((b) => {
                          const tone = b.state === "DEPLOYED" ? C.review : b.state === "RETURNED" ? C.approve : C.charcoalSoft;
                          const soft = b.state === "DEPLOYED" ? C.reviewSoft : b.state === "RETURNED" ? C.approveSoft : C.paper;
                          return (
                            <div key={b.serial} style={{
                              display: "flex", alignItems: "center", justifyContent: "space-between",
                              padding: "9px 11px", borderRadius: 8, background: soft
                            }}>
                              <div>
                                <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 12, fontWeight: 700, color: C.charcoal }}>
                                  {b.serial}
                                </div>
                                {b.deployment && b.state === "DEPLOYED" && (
                                  <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9.5, color: C.charcoalSoft, marginTop: 1 }}>
                                    {b.deployment.deployAsset || "\u2014"} \u00b7 {b.deployment.position} \u00b7 {daysSince(b.deployment.id)}d out
                                  </div>
                                )}
                                {b.retrieval && b.state === "RETURNED" && (
                                  <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9.5, color: C.charcoalSoft, marginTop: 1 }}>
                                    {b.retrieval.outcome || "no outcome"}
                                  </div>
                                )}
                              </div>
                              <span style={{
                                fontFamily: "'Inter',sans-serif", fontSize: 9, fontWeight: 700,
                                color: tone, padding: "3px 9px", borderRadius: 999, background: "#fff"
                              }}>{b.state}</span>
                            </div>
                          );
                        })}
                      </div>
                    )}

                    {fidoMode === "deploy" && (
                      <>
                        <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.charcoalSoft, marginBottom: 6 }}>
                          SELECT A SENSOR
                        </div>
                        <div style={{ display: "flex", flexWrap: "wrap", gap: 5, marginBottom: 4 }}>
                          {available.map((b) => {
                            const active = pendingBug === b.serial;
                            return (
                              <button key={b.serial} onClick={() => setPendingBug(active ? "" : b.serial)} style={{
                                padding: "8px 11px", borderRadius: 8, cursor: "pointer",
                                border: `1.5px solid ${active ? C.primary : C.line}`,
                                background: active ? "#E7F2FE" : "#fff",
                                fontFamily: "'IBM Plex Mono',monospace", fontSize: 11, fontWeight: active ? 700 : 600,
                                color: active ? C.primary : C.charcoal
                              }}>{b.serial}</button>
                            );
                          })}
                        </div>
                        {available.length === 0 && (
                          <p style={{ fontFamily: "'Inter',sans-serif", fontSize: 11, color: C.review, textAlign: "center" }}>
                            All sensors are currently deployed.
                          </p>
                        )}
                        <button
                          onClick={() => {
                            if (!pendingBug) return;
                            setPending((p) => ({ ...(p || {}), bugSerial: pendingBug }));
                            openPicker("fido2Deploy");
                          }}
                          disabled={!pendingBug || !position.trim() || stage === "scanning"}
                          style={{
                            width: "100%", marginTop: 10, padding: "14px", borderRadius: 12,
                            border: "none", cursor: pendingBug && position.trim() ? "pointer" : "not-allowed",
                            background: pendingBug && position.trim() ? C.primary : C.line, color: "#fff",
                            fontFamily: "'Inter',sans-serif", fontWeight: 700, fontSize: 13,
                            display: "flex", alignItems: "center", justifyContent: "center", gap: 8
                          }}>
                          {stage === "scanning" ? <Loader2 size={16} className="spin" /> : <Camera size={16} />}
                          Photograph Deployment Point
                        </button>
                      </>
                    )}

                    {fidoMode === "retrieve" && (
                      <>
                        <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.charcoalSoft, marginBottom: 6 }}>
                          WHICH SENSOR ARE YOU COLLECTING?
                        </div>
                        {deployed.length === 0 ? (
                          <p style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, color: C.charcoalSoft, textAlign: "center", padding: "20px 0" }}>
                            No sensors are currently deployed.
                          </p>
                        ) : (
                          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                            {deployed.map((b) => (
                              <button key={b.serial}
                                onClick={() => {
                                  const d = b.deployment;
                                  setPending({
                                    id: Date.now(),
                                    type: "fido2_retrieve",
                                    position: d.position,
                                    reading: "", serial: "", confidence: 0,
                                    gps: d.gps,
                                    timestamp: nowStamp(),
                                    status: "needs_review",
                                    tech: survey.tech,
                                    sentAt: null,
                                    photo: null,
                                    bugSerial: b.serial,
                                    deployAsset: d.deployAsset || "",
                                    fidoMode: d.fidoMode || "",
                                    sessionId: d.sessionId || "",
                                    pipeMaterial: d.pipeMaterial || "",
                                    cond_diameter: d.cond_diameter || "",
                                    cond_pressure: d.cond_pressure || "",
                                    cond_backfill: d.cond_backfill || "",
                                    cond_pipecondition: d.cond_pipecondition || "",
                                    cond_background: d.cond_background || "",
                                    cond_consumption: d.cond_consumption || "",
                                    deployedDays: daysSince(d.id),
                                    resultShot: null,
                                    resultSessionId: "",
                                    outcome: "",
                                    outcomeNote: "",
                                  });
                                  setStage("result");
                                }}
                                style={{
                                  display: "flex", alignItems: "center", justifyContent: "space-between",
                                  padding: "12px 13px", borderRadius: 10, cursor: "pointer",
                                  border: `1.5px solid ${C.line}`, background: "#fff", textAlign: "left"
                                }}>
                                <div>
                                  <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 12.5, fontWeight: 700, color: C.charcoal }}>
                                    {b.serial}
                                  </div>
                                  <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, color: C.charcoalSoft, marginTop: 2 }}>
                                    {b.deployment.deployAsset || "\u2014"} \u00b7 {b.deployment.position} \u00b7 out {daysSince(b.deployment.id)} day{daysSince(b.deployment.id) === 1 ? "" : "s"}
                                  </div>
                                </div>
                                <ChevronRight size={16} color={C.charcoalSoft} />
                              </button>
                            ))}
                          </div>
                        )}
                      </>
                    )}
                  </>
                );
              })() : task === "assets" ? (
                <div
                  onClick={() => openPicker("asset")}
                  style={{
                    height: 130, borderRadius: 14, cursor: stage === "idle" ? "pointer" : "default",
                    border: `2px dashed ${position.trim() ? C.primary : C.line}`,
                    background: C.paper, display: "flex", flexDirection: "column",
                    alignItems: "center", justifyContent: "center", gap: 8,
                    animation: shake ? "shake 0.4s" : "none"
                  }}>
                  {stage === "scanning" ? (
                    <>
                      <Loader2 size={26} color={C.primary} className="spin" />
                      <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 12.5, color: C.charcoalSoft }}>Reading photo…</span>
                    </>
                  ) : (
                    <>
                      <LayoutGrid size={26} color={position.trim() ? C.primary : C.charcoalSoft} />
                      <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 12.5, fontWeight: 600, color: position.trim() ? C.primary : C.charcoalSoft }}>
                        Photograph Asset
                      </span>
                    </>
                  )}
                </div>
              ) : task === "amr" ? (
                <>
                {(() => {
                  const prog = registerProgress(survey.amrRegister, captures);
                  if (!prog) return null;
                  return (
                    <div style={{ marginBottom: 12 }}>
                      <div style={{
                        display: "flex", justifyContent: "space-between", alignItems: "center",
                        padding: "10px 12px", borderRadius: 9,
                        background: prog.done === prog.total ? C.approveSoft : "#E7F2FE", marginBottom: 8
                      }}>
                        <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 11, fontWeight: 700, color: prog.done === prog.total ? C.approve : C.primaryDeep }}>
                          {prog.done} of {prog.total} SURVEYED
                        </span>
                        <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 11, fontWeight: 700, color: prog.done === prog.total ? C.approve : C.primaryDeep }}>
                          {prog.outstanding.length} left
                        </span>
                      </div>

                      {prog.outstanding.length > 0 && !pendingRegAsset && (
                        <>
                          <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, fontWeight: 700, color: C.charcoalSoft, marginBottom: 5 }}>
                            STILL TO SURVEY \u2014 TAP TO SELECT
                          </div>
                          <div style={{ display: "flex", flexDirection: "column", gap: 4, maxHeight: 210, overflowY: "auto" }}>
                            {prog.outstanding.map((a) => (
                              <button key={a.serial}
                                onClick={() => { setPendingRegAsset(a); setPosition(a.description || a.serial); }}
                                style={{
                                  display: "flex", justifyContent: "space-between", alignItems: "center",
                                  padding: "9px 11px", borderRadius: 8, cursor: "pointer",
                                  border: `1.5px solid ${C.line}`, background: "#fff", textAlign: "left"
                                }}>
                                <div>
                                  <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, fontWeight: 600, color: C.charcoal }}>
                                    {a.description || a.serial}
                                  </div>
                                  <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 9.5, color: C.charcoalSoft, marginTop: 1 }}>
                                    {a.serial}{a.assetType ? ` \u00b7 ${a.assetType}` : ""}
                                  </div>
                                </div>
                                <ChevronRight size={14} color={C.charcoalSoft} />
                              </button>
                            ))}
                          </div>
                        </>
                      )}

                      {pendingRegAsset && (
                        <div style={{ padding: "11px 13px", borderRadius: 9, background: "#E7F2FE", border: `1.5px solid ${C.primary}` }}>
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                            <div>
                              <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 12.5, fontWeight: 700, color: C.primaryDeep }}>
                                {pendingRegAsset.description || pendingRegAsset.serial}
                              </div>
                              <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 10.5, color: C.charcoalSoft, marginTop: 2 }}>
                                {pendingRegAsset.serial}{pendingRegAsset.assetType ? ` \u00b7 ${pendingRegAsset.assetType}` : ""}
                              </div>
                              {pendingRegAsset.gps && (
                                <div style={{ display: "flex", alignItems: "center", gap: 4, marginTop: 4 }}>
                                  <MapPin size={10} color={C.charcoalSoft} />
                                  <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 9.5, color: C.charcoalSoft }}>
                                    on record: {pendingRegAsset.gps.lat}, {pendingRegAsset.gps.lng}
                                  </span>
                                </div>
                              )}
                            </div>
                            <button onClick={() => { setPendingRegAsset(null); setPosition(""); }}
                              style={{ border: "none", background: "none", cursor: "pointer", padding: 2 }}>
                              <X size={14} color={C.charcoalSoft} />
                            </button>
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })()}

                <div
                  onClick={() => openPicker("amrAsset")}
                  style={{
                    height: 130, borderRadius: 14, cursor: stage === "idle" ? "pointer" : "default",
                    border: `2px dashed ${position.trim() ? C.primary : C.line}`,
                    background: C.paper, display: "flex", flexDirection: "column",
                    alignItems: "center", justifyContent: "center", gap: 8,
                    animation: shake ? "shake 0.4s" : "none"
                  }}>
                  {stage === "scanning" ? (
                    <>
                      <Loader2 size={26} color={C.primary} className="spin" />
                      <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 12.5, color: C.charcoalSoft }}>Reading photo…</span>
                    </>
                  ) : (
                    <>
                      <Radio size={26} color={position.trim() ? C.primary : C.charcoalSoft} />
                      <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 12.5, fontWeight: 600, color: position.trim() ? C.primary : C.charcoalSoft }}>
                        Photograph MUC / Repeater
                      </span>
                    </>
                  )}
                </div>
                </>
              ) : (task === "meterwork" && survey.meterMode === "replace") ? (
                <div style={{ display: "flex", flexDirection: "column", gap: 9 }}>
                  <button
                    onClick={() => openPicker("oldMeter")}
                    disabled={stage === "scanning"}
                    style={{
                      padding: "18px 12px", borderRadius: 12, cursor: "pointer",
                      border: `2px dashed ${pending?.oldMeter ? C.approve : position.trim() ? C.primary : C.line}`,
                      background: pending?.oldMeter ? C.approveSoft : C.paper,
                      display: "flex", alignItems: "center", justifyContent: "center", gap: 8,
                      fontFamily: "'Inter',sans-serif", fontSize: 13, fontWeight: 600,
                      color: pending?.oldMeter ? C.approve : position.trim() ? C.primary : C.charcoalSoft,
                      animation: shake ? "shake 0.4s" : "none"
                    }}>
                    {pending?.oldMeter ? <CheckCircle2 size={18} /> : <Camera size={18} />}
                    {pending?.oldMeter ? "Old Meter Captured — retake" : "1 · Photograph OLD Meter"}
                  </button>
                  <button
                    onClick={() => openPicker("newMeter")}
                    disabled={stage === "scanning"}
                    style={{
                      padding: "18px 12px", borderRadius: 12, cursor: "pointer",
                      border: `2px dashed ${pending?.newMeter ? C.approve : position.trim() ? C.primary : C.line}`,
                      background: pending?.newMeter ? C.approveSoft : C.paper,
                      display: "flex", alignItems: "center", justifyContent: "center", gap: 8,
                      fontFamily: "'Inter',sans-serif", fontSize: 13, fontWeight: 600,
                      color: pending?.newMeter ? C.approve : position.trim() ? C.primary : C.charcoalSoft
                    }}>
                    {pending?.newMeter ? <CheckCircle2 size={18} /> : <Camera size={18} />}
                    {pending?.newMeter ? "New Meter Captured — retake" : "2 · Photograph NEW Meter"}
                  </button>
                  {stage === "scanning" && (
                    <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 8, marginTop: 4 }}>
                      <Loader2 size={16} color={C.primary} className="spin" />
                      <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 12, color: C.charcoalSoft }}>Reading meter…</span>
                    </div>
                  )}
                </div>
              ) : (
              <div
                onClick={() => openPicker(task === "meterwork" ? "meter" : "sensor")}
                style={{
                  height: 130, borderRadius: 14, cursor: stage === "idle" ? "pointer" : "default",
                  border: `2px dashed ${position.trim() ? C.primary : C.line}`,
                  background: C.paper, display: "flex", flexDirection: "column",
                  alignItems: "center", justifyContent: "center", gap: 8, transition: "border-color 0.2s",
                  animation: shake ? "shake 0.4s" : "none"
                }}>
                {stage === "scanning" ? (
                  <>
                    <Loader2 size={26} color={C.primary} className="spin" />
                    <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 12.5, color: C.charcoalSoft }}>
                      Reading photo…
                    </span>
                  </>
                ) : (
                  <>
                    {task === "fido" ? <Radio size={26} color={position.trim() ? C.primary : C.charcoalSoft} /> : <Camera size={26} color={position.trim() ? C.primary : C.charcoalSoft} />}
                    <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 12.5, fontWeight: 600, color: position.trim() ? C.primary : C.charcoalSoft }}>
                      {task === "meterwork" ? "Capture Meter" : "Sensor Deployment"}
                    </span>
                  </>
                )}
              </div>
              )}

              {task === "fido" && (
                <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
                  <button
                    onClick={() => openPicker("fido")}
                    disabled={stage === "scanning"}
                    style={{
                      flex: 1, padding: "11px 8px", borderRadius: 11, cursor: "pointer",
                      border: `1.5px dashed ${position.trim() ? C.primary : C.line}`, background: C.paper,
                      display: "flex", alignItems: "center", justifyContent: "center", gap: 6,
                      fontFamily: "'Inter',sans-serif", fontSize: 11.5, fontWeight: 600,
                      color: position.trim() ? C.primary : C.charcoalSoft
                    }}>
                    <Activity size={14} /> FIDO Session Feedback
                  </button>
                </div>
              )}

              {task === "meterwork" && previousForPosition && (
                <div style={{
                  marginTop: 10, padding: "9px 11px", borderRadius: 9, background: "#E7F2FE",
                  display: "flex", alignItems: "center", justifyContent: "space-between"
                }}>
                  <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 11, fontWeight: 700, color: C.primaryDeep }}>
                    PREVIOUS READING
                  </span>
                  <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 12.5, fontWeight: 700, color: C.primaryDeep }}>
                    {previousForPosition.reading} m\u00b3{previousForPosition.date ? ` \u00b7 ${previousForPosition.date}` : ""}
                  </span>
                </div>
              )}

              <p style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, color: shake || captureError ? C.flag : justSaved ? C.approve : C.charcoalSoft, textAlign: "center", marginTop: 10, fontWeight: justSaved ? 600 : 400 }}>
                {captureError || (shake ? "Enter a position first." : justSaved ? "\u2713 Saved \u2014 ready for the next point" : "Pick what you're capturing \u2014 you can add the others afterwards. GPS and timestamp are recorded automatically.")}
              </p>

              {task === "meterwork" && (
                <button
                  onClick={openException}
                  style={{
                    width: "100%", marginTop: 6, padding: "10px", borderRadius: 9,
                    border: `1.5px solid ${C.line}`, background: "#fff", cursor: "pointer",
                    fontFamily: "'Inter',sans-serif", fontSize: 12, fontWeight: 600, color: C.review,
                    display: "flex", alignItems: "center", justifyContent: "center", gap: 6
                  }}>
                  <AlertTriangle size={14} /> Can't read this meter
                </button>
              )}
            </>
          )}

          {stage === "result" && pending && (
            <>
              {/* type banner so it's always clear which kind of capture this is */}
              <div style={{
                display: "flex", alignItems: "center", gap: 6, marginBottom: 10,
                fontFamily: "'Inter',sans-serif", fontSize: 12, fontWeight: 700, color: C.primary
              }}>
                {pending.type === "meter" && <><Camera size={14} /> Meter Reading</>}
                {pending.type === "sensor" && <><Radio size={14} /> Sensor Deployment</>}
                {pending.type === "fido" && <><Activity size={14} /> FIDO Feedback</>}
                {pending.type === "replacement" && <><RotateCcw size={14} /> Meter Replacement</>}
                {pending.type === "amr" && <><Radio size={14} /> {pending.amrAssetType || "AMR"} Survey</>}
                {pending.type === "asset" && <><LayoutGrid size={14} /> {pending.assetType || "Asset Mapping"}</>}
                {pending.type === "fido2_deploy" && <><Radio size={14} /> Deploy {pending.bugSerial}</>}
                {pending.type === "fido2_retrieve" && <><Radio size={14} /> Retrieve {pending.bugSerial}</>}
                {pending.type === "fido2_waypoint" && <><MapPin size={14} /> {pending.waypointRef}</>}
              </div>

              {captureError && (
                <div style={{
                  marginBottom: 10, padding: "8px 11px", borderRadius: 8, background: C.reviewSoft,
                  fontFamily: "'Inter',sans-serif", fontSize: 11.5, color: C.review, fontWeight: 600
                }}>
                  {captureError}
                </div>
              )}

              {/* ---- METER: photo + reading/serial + confidence ---- */}
              {pending.type === "meter" && (
                <>
                  <PhotoWell position={pending.position} timestamp={pending.timestamp} gps={pending.gps} photo={pending.photo} />
                  <div style={{
                    display: "flex", alignItems: "center", justifyContent: "center", gap: 5, marginTop: 8,
                    padding: "7px 10px", borderRadius: 8, background: C.paper
                  }}>
                    <MapPin size={12} color={pending.gps?.lat !== "Unknown" ? C.approve : C.review} />
                    <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 11.5, color: C.charcoal, fontWeight: 600 }}>
                      {pending.gps ? `${pending.gps.lat}, ${pending.gps.lng}` : "GPS unavailable"}
                    </span>
                  </div>
                  <div style={{ display: "flex", gap: 10, marginTop: 14 }}>
                    <div style={{ flex: 1 }}>
                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 600, color: C.charcoalSoft, marginBottom: 4 }}>READING (m³)</div>
                      <input value={pending.reading} onChange={(e) => editField("reading", e.target.value)}
                        style={{ width: "100%", boxSizing: "border-box", padding: "8px 10px", borderRadius: 8, border: `1.5px solid ${C.line}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 14, fontWeight: 600, color: C.charcoal }} />
                    </div>
                    <div style={{ flex: 1 }}>
                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 600, color: C.charcoalSoft, marginBottom: 4 }}>SERIAL</div>
                      <input value={pending.serial} onChange={(e) => editField("serial", e.target.value)}
                        style={{ width: "100%", boxSizing: "border-box", padding: "8px 10px", borderRadius: 8, border: `1.5px solid ${C.line}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 14, fontWeight: 600, color: C.charcoal }} />
                    </div>
                  </div>
                  <div style={{
                    marginTop: 12, display: "flex", alignItems: "center", justifyContent: "space-between",
                    padding: "10px 12px", borderRadius: 9,
                    background: pending.confidence < 85 ? C.reviewSoft : C.approveSoft
                  }}>
                    <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 12.5, fontWeight: 600, color: pending.confidence < 85 ? C.review : C.approve }}>
                      {pending.confidence < 85 ? "Low confidence — flagged for review" : "AI extraction confident"}
                    </span>
                    <MiniGauge value={pending.confidence} color={pending.confidence < 85 ? C.review : C.approve} />
                  </div>

                  {(() => {
                    const prev = lookupPrevious(survey.previousReadings, pending.position, pending.serial);
                    if (!prev) return null;
                    const now = parseFloat(pending.reading);
                    const before = parseFloat(prev.reading);
                    const valid = !isNaN(now) && !isNaN(before);
                    const used = valid ? now - before : null;
                    const odd = valid && (used < 0 || used > before * 3);
                    return (
                      <div style={{
                        marginTop: 10, padding: "10px 12px", borderRadius: 9,
                        background: odd ? C.reviewSoft : "#E7F2FE"
                      }}>
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                          <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: odd ? C.review : C.primaryDeep }}>
                            PREVIOUS {prev.date ? `\u00b7 ${prev.date}` : ""}
                          </span>
                          <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 12, fontWeight: 700, color: odd ? C.review : C.primaryDeep }}>
                            {prev.reading} m\u00b3
                          </span>
                        </div>
                        {valid && (
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 4 }}>
                            <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: odd ? C.review : C.primaryDeep }}>
                              CONSUMPTION
                            </span>
                            <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 12, fontWeight: 700, color: odd ? C.review : C.primaryDeep }}>
                              {used.toFixed(2)} m\u00b3
                            </span>
                          </div>
                        )}
                        {odd && (
                          <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, fontWeight: 600, color: C.review, marginTop: 5 }}>
                            {used < 0 ? "READING IS LOWER THAN LAST TIME \u2014 CHECK IT" : "UNUSUALLY HIGH CONSUMPTION \u2014 CHECK IT"}
                          </div>
                        )}
                      </div>
                    );
                  })()}
                </>
              )}

              {/* ---- SENSOR: installation photo + session screenshot ---- */}
              {pending.type === "sensor" && (
                <>
                  <PhotoWell position={pending.position} timestamp={pending.timestamp} gps={pending.gps} photo={pending.photo} />
                  <div style={{
                    display: "flex", alignItems: "center", justifyContent: "center", gap: 5, marginTop: 8,
                    padding: "7px 10px", borderRadius: 8, background: C.paper
                  }}>
                    <MapPin size={12} color={pending.gps?.lat !== "Unknown" ? C.approve : C.review} />
                    <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 11.5, color: C.charcoal, fontWeight: 600 }}>
                      {pending.gps ? `${pending.gps.lat}, ${pending.gps.lng}` : "GPS unavailable"}
                    </span>
                  </div>
                  <p style={{ fontFamily: "'Inter',sans-serif", fontSize: 11, color: C.charcoalSoft, margin: "8px 0 0", textAlign: "center" }}>
                    Installation photo. Now add the session screenshot from your gallery:
                  </p>
                  <EvidenceSlot
                    title="Session Screenshot"
                    icon={Radio}
                    promptText={SENSOR_PROMPT}
                    value={pending.sensor}
                    onChange={(v) => editField("sensor", v)}
                    fieldsConfig={[
                      { key: "sessionId", label: "SESSION ID" },
                      { key: "deviceId", label: "DEVICE ID" },
                      { key: "signalStrength", label: "SIGNAL" },
                      { key: "batteryVoltage", label: "BATTERY" },
                    ]}
                  />
                </>
              )}

              {/* ---- FIDO 2 WAYPOINT ---- */}
              {pending.type === "fido2_waypoint" && (
                <>
                  <div style={{ padding: "12px 14px", borderRadius: 10, background: "#E7F2FE", marginBottom: 12 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 16, fontWeight: 700, color: C.primaryDeep }}>
                        {pending.waypointRef}
                      </span>
                      <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 10, color: C.charcoalSoft }}>
                        {pending.gps ? `${pending.gps.lat}, ${pending.gps.lng}` : "GPS unavailable"}
                      </span>
                    </div>
                    <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, color: C.charcoalSoft, marginTop: 3 }}>
                      {pending.position}{pending.linkedBug ? ` \u00b7 raised by ${pending.linkedBug} (${pending.triggerOutcome})` : ""}
                    </div>
                  </div>

                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                    <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.charcoalSoft }}>
                      TESTS AT THIS WAYPOINT ({(pending.tests || []).length})
                    </span>
                  </div>

                  <div style={{ display: "flex", flexWrap: "wrap", gap: 5, marginBottom: 10 }}>
                    {Object.keys(WAYPOINT_TESTS).map((m) => (
                      <button key={m}
                        onClick={() => { setPendingTestMethod(m); openPicker("waypointTest"); }}
                        disabled={stage === "scanning"}
                        style={{
                          padding: "8px 11px", borderRadius: 8, cursor: "pointer",
                          border: `1.5px solid ${C.line}`, background: "#fff",
                          fontFamily: "'Inter',sans-serif", fontSize: 10, fontWeight: 600, color: C.charcoal,
                          display: "flex", alignItems: "center", gap: 5
                        }}>
                        {stage === "scanning" && pendingTestMethod === m ? <Loader2 size={11} className="spin" /> : <Camera size={11} />}
                        {m}
                      </button>
                    ))}
                  </div>

                  {(pending.tests || []).length > 0 && (
                    <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 12 }}>
                      {pending.tests.map((t, i) => (
                        <div key={i} style={{ padding: 10, borderRadius: 9, background: C.paper }}>
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 7 }}>
                            <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.charcoal }}>
                              {t.method}
                            </span>
                            <button onClick={() => editField("tests", pending.tests.filter((_, j) => j !== i))}
                              style={{ border: "none", background: "none", cursor: "pointer", padding: 2 }}>
                              <X size={12} color={C.charcoalSoft} />
                            </button>
                          </div>
                          <img src={t.photo} alt={t.method}
                            onClick={() => setZoomShot(t.photo)}
                            style={{ width: "100%", maxHeight: 190, objectFit: "contain", borderRadius: 7, background: "#fff", cursor: "zoom-in" }} />
                          {WAYPOINT_TESTS[t.method]?.needsDistance && (
                            <div style={{ display: "flex", gap: 7, marginTop: 8 }}>
                              <div style={{ flex: 1 }}>
                                <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 8.5, fontWeight: 700, color: C.charcoalSoft }}>LEAK DISTANCE</div>
                                <input value={t.distance || ""}
                                  onChange={(e) => editField("tests", pending.tests.map((x, j) => j === i ? { ...x, distance: e.target.value } : x))}
                                  placeholder="\u2014"
                                  style={{ width: "100%", boxSizing: "border-box", padding: "5px 7px", borderRadius: 6, border: `1.5px solid ${C.line}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 11.5, fontWeight: 700, color: C.charcoal, background: "#fff" }} />
                              </div>
                              <div style={{ flex: 1 }}>
                                <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 8.5, fontWeight: 700, color: C.charcoalSoft }}>SENSOR SPACING</div>
                                <input value={t.sensorSpacing || ""}
                                  onChange={(e) => editField("tests", pending.tests.map((x, j) => j === i ? { ...x, sensorSpacing: e.target.value } : x))}
                                  placeholder="\u2014"
                                  style={{ width: "100%", boxSizing: "border-box", padding: "5px 7px", borderRadius: 6, border: `1.5px solid ${C.line}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 11.5, color: C.charcoal, background: "#fff" }} />
                              </div>
                            </div>
                          )}
                          <input value={t.note || ""}
                            onChange={(e) => editField("tests", pending.tests.map((x, j) => j === i ? { ...x, note: e.target.value } : x))}
                            placeholder="What did this test show?"
                            style={{ width: "100%", boxSizing: "border-box", marginTop: 7, padding: "7px 9px", borderRadius: 7, border: `1.5px solid ${C.line}`, fontFamily: "'Inter',sans-serif", fontSize: 11.5, color: C.charcoal, background: "#fff" }} />
                        </div>
                      ))}
                    </div>
                  )}

                  <div style={{
                    padding: 11, borderRadius: 11, marginBottom: 12,
                    border: `1.5px dashed ${pending.pinpointPhoto ? C.approve : C.line}`,
                    background: pending.pinpointPhoto ? C.approveSoft : C.paper
                  }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, fontWeight: 700, color: pending.pinpointPhoto ? C.approve : C.charcoal }}>
                        PINPOINT / BEACON
                      </span>
                      <button onClick={() => openPicker("waypointPinpoint")} disabled={stage === "scanning"} style={{
                        border: "none", background: "none", color: C.primary, cursor: "pointer",
                        fontFamily: "'Inter',sans-serif", fontWeight: 600, fontSize: 11.5,
                        display: "flex", alignItems: "center", gap: 4
                      }}>
                        <Camera size={13} /> {pending.pinpointPhoto ? "Retake" : "Add"}
                      </button>
                    </div>
                    {pending.pinpointPhoto ? (
                      <>
                        <img src={pending.pinpointPhoto} alt="pinpoint"
                          style={{ width: "100%", maxHeight: 180, objectFit: "cover", borderRadius: 7, marginTop: 9 }} />
                        <div style={{ display: "flex", alignItems: "center", gap: 5, marginTop: 6 }}>
                          <MapPin size={11} color={C.approve} />
                          <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 11, fontWeight: 700, color: C.charcoal }}>
                            {pending.pinpointGps ? `${pending.pinpointGps.lat}, ${pending.pinpointGps.lng}` : "GPS unavailable"}
                          </span>
                        </div>
                      </>
                    ) : (
                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, color: C.charcoalSoft, marginTop: 5 }}>
                        Place a beacon at the suspected leak and photograph it. GPS is recorded at that moment.
                      </div>
                    )}
                  </div>

                  <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.charcoalSoft, marginBottom: 6 }}>
                    INVESTIGATION RESULT
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}>
                    {Object.entries(WAYPOINT_OUTCOMES).map(([o, st]) => {
                      const on = pending.outcome === o;
                      return (
                        <button key={o} onClick={() => editField("outcome", o)} style={{
                          padding: "11px 6px", borderRadius: 9, cursor: "pointer",
                          border: `1.5px solid ${on ? st.colour : C.line}`,
                          background: on ? st.soft : "#fff",
                          fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700,
                          color: on ? st.colour : C.charcoalSoft
                        }}>{o}</button>
                      );
                    })}
                  </div>
                  <input value={pending.outcomeNote || ""}
                    onChange={(e) => editField("outcomeNote", e.target.value)}
                    placeholder="Conclusion \u2014 what was found and what should happen next?"
                    style={{
                      width: "100%", boxSizing: "border-box", marginTop: 9, padding: "9px 11px",
                      borderRadius: 8, border: `1.5px solid ${C.line}`,
                      fontFamily: "'Inter',sans-serif", fontSize: 12, color: C.charcoal
                    }} />
                </>
              )}

              {/* ---- FIDO 2 DEPLOYMENT ---- */}
              {pending.type === "fido2_deploy" && (
                <>
                  <PhotoWell position={pending.position} timestamp={pending.timestamp} gps={pending.gps} photo={pending.photo} />
                  <div style={{
                    display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, marginTop: 8,
                    padding: "8px 11px", borderRadius: 8, background: "#E7F2FE"
                  }}>
                    <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 12.5, fontWeight: 700, color: C.primaryDeep }}>
                      {pending.bugSerial}
                    </span>
                    <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 10.5, color: C.charcoalSoft }}>
                      {pending.gps ? `${pending.gps.lat}, ${pending.gps.lng}` : "GPS unavailable"}
                    </span>
                  </div>

                  <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.charcoalSoft, margin: "13px 0 5px" }}>
                    DEPLOYED ON
                  </div>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>
                    {FIDO_DEPLOY_ASSETS.map((a) => {
                      const on = pending.deployAsset === a.v;
                      return (
                        <button key={a.v} onClick={() => editField("deployAsset", a.v)} style={{
                          padding: "7px 10px", borderRadius: 999, cursor: "pointer",
                          border: `1.5px solid ${on ? C.primary : C.line}`,
                          background: on ? "#E7F2FE" : "#fff",
                          fontFamily: "'Inter',sans-serif", fontSize: 10, fontWeight: on ? 700 : 600,
                          color: on ? C.primary : C.charcoalSoft
                        }}>{a.v}</button>
                      );
                    })}
                  </div>
                  {(() => {
                    const a = FIDO_DEPLOY_ASSETS.find((x) => x.v === pending.deployAsset);
                    if (!a || !a.note) return null;
                    const col = a.coupling > 1 ? C.approve : C.review;
                    return (
                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9.5, color: col, marginTop: 5, fontWeight: 600 }}>
                        {a.note}
                      </div>
                    );
                  })()}

                  <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.charcoalSoft, margin: "13px 0 5px" }}>
                    PIPE MATERIAL
                  </div>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>
                    {Object.entries(PIPE_MATERIALS).map(([k, m]) => {
                      const on = pending.pipeMaterial === k;
                      return (
                        <button key={k} onClick={() => editField("pipeMaterial", k)} style={{
                          padding: "7px 10px", borderRadius: 999, cursor: "pointer",
                          border: `1.5px solid ${on ? m.tone : C.line}`,
                          background: on ? `${m.tone}18` : "#fff",
                          fontFamily: "'Inter',sans-serif", fontSize: 10, fontWeight: on ? 700 : 600,
                          color: on ? m.tone : C.charcoalSoft
                        }}>{m.label}</button>
                      );
                    })}
                  </div>
                  {pending.pipeMaterial && (
                    <div style={{ marginTop: 9 }}>
                      <PipeAcousticGuide material={pending.pipeMaterial} />
                    </div>
                  )}

                  <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.charcoalSoft, margin: "13px 0 5px" }}>
                    SESSION MODE
                  </div>
                  <div style={{ display: "flex", gap: 5 }}>
                    {FIDO_MODES.map((m) => {
                      const on = pending.fidoMode === m;
                      return (
                        <button key={m} onClick={() => editField("fidoMode", m)} style={{
                          flex: 1, padding: "9px 5px", borderRadius: 8, cursor: "pointer",
                          border: `1.5px solid ${on ? C.primary : C.line}`,
                          background: on ? "#E7F2FE" : "#fff",
                          fontFamily: "'Inter',sans-serif", fontSize: 9.5, fontWeight: on ? 700 : 600,
                          color: on ? C.primary : C.charcoalSoft
                        }}>{m}</button>
                      );
                    })}
                  </div>

                  <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.charcoalSoft, margin: "14px 0 6px" }}>
                    LISTENING CONDITIONS
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 9 }}>
                    {Object.entries(DEPLOY_CONDITIONS).map(([key, cfg]) => (
                      <div key={key}>
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 4 }}>
                          <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 9, fontWeight: 700, color: C.charcoalSoft }}>{cfg.label}</span>
                          <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 8.5, color: C.charcoalSoft, opacity: 0.8 }}>{cfg.hint}</span>
                        </div>
                        <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
                          {cfg.options.map((o) => {
                            const on = pending[`cond_${key}`] === o.v;
                            const col = o.na ? C.charcoalSoft : o.good > 0 ? C.approve : o.good === 0 ? C.review : C.flag;
                            return (
                              <button key={o.v} onClick={() => editField(`cond_${key}`, on ? "" : o.v)} style={{
                                flex: "1 1 auto", minWidth: 62, padding: "7px 4px", borderRadius: 7, cursor: "pointer",
                                border: `1.5px solid ${on ? col : C.line}`,
                                background: on ? `${col}18` : "#fff",
                                fontFamily: "'Inter',sans-serif", fontSize: 8.5, fontWeight: on ? 700 : 600,
                                color: on ? col : C.charcoalSoft
                              }}>{o.v}</button>
                            );
                          })}
                        </div>
                      </div>
                    ))}
                  </div>

                  {(() => {
                    const s = conditionScore(pending);
                    if (!s) return null;
                    return (
                      <div style={{
                        marginTop: 10, padding: "10px 12px", borderRadius: 9,
                        background: s.soft, border: `1px solid ${s.colour}55`
                      }}>
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                          <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: s.colour }}>
                            {s.label}
                          </span>
                          <span style={{ fontFamily: "'Space Grotesk',sans-serif", fontSize: 15, fontWeight: 700, color: s.colour }}>
                            {s.pct}%
                          </span>
                        </div>
                        <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, color: C.charcoalSoft, marginTop: 4, lineHeight: 1.4 }}>
                          {s.advice}
                        </div>
                      </div>
                    );
                  })()}

                  <div style={{
                    marginTop: 13, padding: 11, borderRadius: 11,
                    border: `1.5px dashed ${pending.sessionShot ? C.approve : C.line}`,
                    background: pending.sessionShot ? C.approveSoft : C.paper
                  }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, fontWeight: 700, color: pending.sessionShot ? C.approve : C.charcoal }}>
                        SESSION SCREENSHOT
                      </span>
                      <button onClick={() => openPicker("fido2Session")} disabled={stage === "scanning"} style={{
                        border: "none", background: "none", color: C.primary, cursor: "pointer",
                        fontFamily: "'Inter',sans-serif", fontWeight: 600, fontSize: 11.5,
                        display: "flex", alignItems: "center", gap: 4
                      }}>
                        {stage === "scanning" ? <Loader2 size={13} className="spin" /> : <Camera size={13} />}
                        {pending.sessionShot ? "Retake" : "Add"}
                      </button>
                    </div>
                    {pending.sessionShot && (
                      <div style={{ display: "flex", gap: 9, marginTop: 9 }}>
                        <img src={pending.sessionShot.photo} alt="session" style={{ width: 52, height: 68, objectFit: "cover", borderRadius: 6, flexShrink: 0 }} />
                        <div style={{ flex: 1, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}>
                          <div style={{ gridColumn: "1 / -1" }}>
                            <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9, fontWeight: 600, color: C.charcoalSoft }}>SESSION ID</div>
                            <input value={pending.sessionId || ""} onChange={(e) => editField("sessionId", e.target.value)} placeholder="\u2014"
                              style={{ width: "100%", boxSizing: "border-box", padding: "5px 7px", borderRadius: 6, border: `1.5px solid ${pending.sessionId ? C.line : C.review}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 12, fontWeight: 700, color: C.charcoal, background: "#fff" }} />
                          </div>
                          <div>
                            <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9, fontWeight: 600, color: C.charcoalSoft }}>SIGNAL</div>
                            <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 11, color: C.charcoal }}>{pending.sessionShot.signalStrength || "\u2014"}</div>
                          </div>
                          <div>
                            <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9, fontWeight: 600, color: C.charcoalSoft }}>BATTERY</div>
                            <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 11, color: C.charcoal }}>{pending.sessionShot.batteryVoltage || "\u2014"}</div>
                          </div>
                        </div>
                      </div>
                    )}
                  </div>

                  {(!pending.deployAsset || !pending.pipeMaterial || !pending.sessionId) && (
                    <p style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, color: C.review, textAlign: "center", marginTop: 9 }}>
                      {!pending.deployAsset ? "Choose what it's deployed on." : !pending.pipeMaterial ? "Choose the pipe material." : "A session ID is needed to link the retrieval later."}
                    </p>
                  )}
                </>
              )}

              {/* ---- FIDO 2 RETRIEVAL ---- */}
              {pending.type === "fido2_retrieve" && (
                <>
                  <div style={{
                    padding: "11px 13px", borderRadius: 10, background: "#E7F2FE", marginBottom: 12
                  }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 13, fontWeight: 700, color: C.primaryDeep }}>
                        {pending.bugSerial}
                      </span>
                      <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, fontWeight: 600, color: C.primaryDeep }}>
                        OUT {pending.deployedDays} DAY{pending.deployedDays === 1 ? "" : "S"}
                      </span>
                    </div>
                    <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, color: C.charcoalSoft, marginTop: 3 }}>
                      {pending.deployAsset} \u00b7 {pending.position} \u00b7 session {pending.sessionId || "\u2014"}
                    </div>
                  </div>

                  {pending.pipeMaterial && (
                    <div style={{ marginBottom: 11 }}>
                      <PipeAcousticGuide material={pending.pipeMaterial} compact />
                    </div>
                  )}

                  {(() => {
                    const s = conditionScore(pending);
                    if (!s) return null;
                    return (
                      <div style={{
                        marginBottom: 11, padding: "10px 12px", borderRadius: 9,
                        background: s.soft, border: `1px solid ${s.colour}55`
                      }}>
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                          <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: s.colour }}>
                            {s.label}
                          </span>
                          <span style={{ fontFamily: "'Space Grotesk',sans-serif", fontSize: 15, fontWeight: 700, color: s.colour }}>
                            {s.pct}%
                          </span>
                        </div>
                        <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9.5, color: C.charcoalSoft, marginTop: 4, lineHeight: 1.4 }}>
                          {[
                            pending.deployAsset && `on ${pending.deployAsset.toLowerCase()}`,
                            PIPE_MATERIALS[pending.pipeMaterial]?.label,
                            pending.cond_diameter && `\u2300 ${pending.cond_diameter}`,
                            pending.cond_pressure && `${pending.cond_pressure.toLowerCase()} pressure`,
                            pending.cond_backfill && `${pending.cond_backfill.toLowerCase()} backfill`,
                            pending.cond_pipecondition && `${pending.cond_pipecondition.toLowerCase()}`,
                            pending.cond_background && `${pending.cond_background.toLowerCase()}`,
                            pending.cond_consumption && `${pending.cond_consumption.toLowerCase()} demand`,
                          ].filter(Boolean).join(" \u00b7 ")}
                        </div>
                        {s.pct < 45 && (
                          <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9.5, fontWeight: 600, color: s.colour, marginTop: 5 }}>
                            {s.advice}
                          </div>
                        )}
                      </div>
                    );
                  })()}

                  <div style={{
                    padding: 11, borderRadius: 11,
                    border: `1.5px dashed ${pending.resultShot ? C.approve : C.line}`,
                    background: pending.resultShot ? C.approveSoft : C.paper
                  }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, fontWeight: 700, color: pending.resultShot ? C.approve : C.charcoal }}>
                        RESULTS GRAPH
                      </span>
                      <button onClick={() => openPicker("fido2Result")} disabled={stage === "scanning"} style={{
                        border: "none", background: "none", color: C.primary, cursor: "pointer",
                        fontFamily: "'Inter',sans-serif", fontWeight: 600, fontSize: 11.5,
                        display: "flex", alignItems: "center", gap: 4
                      }}>
                        {stage === "scanning" ? <Loader2 size={13} className="spin" /> : <Camera size={13} />}
                        {pending.resultShot ? "Retake" : "Add"}
                      </button>
                    </div>
                    {pending.resultShot && (
                      <>
                        <img src={pending.resultShot.photo} alt="results"
                          onClick={() => setZoomShot(pending.resultShot.photo)}
                          style={{ width: "100%", maxHeight: 240, objectFit: "contain", borderRadius: 7, marginTop: 9, background: "#fff", cursor: "zoom-in" }} />
                        <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9.5, color: C.charcoalSoft, textAlign: "center", marginTop: 4 }}>
                          Tap the graph to view it full size
                        </div>
                        {pending.resultShot.dateRange && (
                          <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 10.5, color: C.charcoalSoft, marginTop: 6 }}>
                            {pending.resultShot.dateRange}
                          </div>
                        )}
                        {pending.resultShot.notes && (
                          <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 11, color: C.charcoalSoft, marginTop: 3 }}>
                            {pending.resultShot.notes}
                          </div>
                        )}
                      </>
                    )}
                  </div>

                  {pending.resultShot && (() => {
                    const a = (pending.sessionId || "").trim().toLowerCase();
                    const b = (pending.resultSessionId || "").trim().toLowerCase();
                    const known = a && b;
                    const match = known && a === b;
                    return (
                      <div style={{
                        marginTop: 10, padding: "9px 11px", borderRadius: 9,
                        background: !known ? C.paperDeep : match ? C.approveSoft : C.flagSoft
                      }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                          {known && (match ? <CheckCircle2 size={14} color={C.approve} /> : <AlertTriangle size={14} color={C.flag} />)}
                          <span style={{
                            fontFamily: "'Inter',sans-serif", fontSize: 11, fontWeight: 700,
                            color: !known ? C.charcoalSoft : match ? C.approve : C.flag
                          }}>
                            {!known ? "SESSION ID NOT CONFIRMED" : match ? "SESSION ID MATCHES DEPLOYMENT" : "SESSION ID DOES NOT MATCH"}
                          </span>
                        </div>
                        <div style={{ display: "flex", gap: 8, marginTop: 7 }}>
                          <div style={{ flex: 1 }}>
                            <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9, fontWeight: 600, color: C.charcoalSoft }}>AT DEPLOYMENT</div>
                            <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 12, fontWeight: 700, color: C.charcoal }}>{pending.sessionId || "\u2014"}</div>
                          </div>
                          <div style={{ flex: 1 }}>
                            <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9, fontWeight: 600, color: C.charcoalSoft }}>ON GRAPH</div>
                            <input value={pending.resultSessionId || ""} onChange={(e) => editField("resultSessionId", e.target.value)} placeholder="\u2014"
                              style={{ width: "100%", boxSizing: "border-box", padding: "3px 6px", borderRadius: 5, border: `1px solid ${C.line}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 12, fontWeight: 700, color: C.charcoal, background: "#fff" }} />
                          </div>
                        </div>
                      </div>
                    );
                  })()}

                  <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.charcoalSoft, margin: "14px 0 6px" }}>
                    FINDING
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}>
                    {Object.entries(FIDO_OUTCOMES).map(([o, style]) => {
                      const on = pending.outcome === o;
                      return (
                        <button key={o} onClick={() => editField("outcome", o)} style={{
                          padding: "11px 6px", borderRadius: 9, cursor: "pointer",
                          border: `1.5px solid ${on ? style.colour : C.line}`,
                          background: on ? style.soft : "#fff",
                          fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700,
                          color: on ? style.colour : C.charcoalSoft
                        }}>{o}</button>
                      );
                    })}
                  </div>
                  <input
                    value={pending.outcomeNote || ""}
                    onChange={(e) => editField("outcomeNote", e.target.value)}
                    placeholder="Note — what did you find?"
                    style={{
                      width: "100%", boxSizing: "border-box", marginTop: 9, padding: "9px 11px",
                      borderRadius: 8, border: `1.5px solid ${C.line}`,
                      fontFamily: "'Inter',sans-serif", fontSize: 12, color: C.charcoal
                    }}
                  />
                  {!pending.outcome && (
                    <p style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, color: C.review, textAlign: "center", marginTop: 8 }}>
                      Record a finding before saving.
                    </p>
                  )}
                </>
              )}

              {/* ---- ASSET MAPPING: photo, GPS, category/type, details ---- */}
              {pending.type === "asset" && (
                <>
                  <PhotoWell position={pending.position} timestamp={pending.timestamp} gps={pending.gps} photo={pending.photo} />
                  <div style={{
                    display: "flex", alignItems: "center", justifyContent: "center", gap: 5, marginTop: 8,
                    padding: "7px 10px", borderRadius: 8, background: C.paper
                  }}>
                    <MapPin size={12} color={pending.gps?.lat !== "Unknown" ? C.approve : C.review} />
                    <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 11.5, color: C.charcoal, fontWeight: 600 }}>
                      {pending.gps ? `${pending.gps.lat}, ${pending.gps.lng}` : "GPS unavailable"}
                    </span>
                  </div>

                  <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.charcoalSoft, margin: "14px 0 6px" }}>
                    CATEGORY
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}>
                    {Object.keys(ASSET_CATEGORIES).map((cat) => {
                      const active = pending.assetCategory === cat;
                      const colour = CATEGORY_COLOURS[cat];
                      return (
                        <button key={cat}
                          onClick={() => setPending((p) => ({ ...p, assetCategory: cat, assetType: null }))}
                          style={{
                            padding: "10px 6px", borderRadius: 9, cursor: "pointer",
                            border: `1.5px solid ${active ? colour : C.line}`,
                            background: active ? `${colour}18` : "#fff",
                            fontFamily: "'Inter',sans-serif", fontSize: 11, fontWeight: 700,
                            color: active ? colour : C.charcoalSoft
                          }}>
                          {cat}
                        </button>
                      );
                    })}
                  </div>

                  {pending.assetCategory && (
                    <>
                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.charcoalSoft, margin: "14px 0 6px" }}>
                        ASSET TYPE
                      </div>
                      <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
                        {ASSET_CATEGORIES[pending.assetCategory].map((t) => {
                          const active = pending.assetType === t;
                          const colour = CATEGORY_COLOURS[pending.assetCategory];
                          return (
                            <button key={t}
                              onClick={() => editField("assetType", t)}
                              style={{
                                padding: "9px 11px", borderRadius: 8, cursor: "pointer", textAlign: "left",
                                border: `1.5px solid ${active ? colour : C.line}`,
                                background: active ? `${colour}18` : "#fff",
                                fontFamily: "'Inter',sans-serif", fontSize: 11.5, fontWeight: active ? 700 : 500,
                                color: active ? colour : C.charcoal
                              }}>
                              {t}
                            </button>
                          );
                        })}
                      </div>
                    </>
                  )}

                  {pending.assetType && (
                    <div style={{ marginTop: 14, display: "flex", flexDirection: "column", gap: 10 }}>
                      <div>
                        <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.charcoalSoft, marginBottom: 4 }}>
                          DESCRIPTION — WHERE IS IT? <span style={{ color: C.review }}>REQUIRED</span>
                        </div>
                        <textarea
                          value={pending.assetDescription || ""}
                          onChange={(e) => editField("assetDescription", e.target.value.toUpperCase())}
                          placeholder="E.G. CUSSONIA WAY KIOSK 1 - INSIDE KIOSK ON LEFT"
                          rows={2}
                          style={{
                            width: "100%", boxSizing: "border-box", padding: "8px 10px", borderRadius: 8,
                            border: `1.5px solid ${pending.assetDescription ? C.line : C.review}`,
                            fontFamily: "'Inter',sans-serif", fontSize: 12, color: C.charcoal, resize: "vertical"
                          }}
                        />
                      </div>
                      <div style={{ display: "flex", gap: 8 }}>
                        <div style={{ flex: 1 }}>
                          <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.charcoalSoft, marginBottom: 4 }}>SERIAL</div>
                          <input value={pending.assetSerial || ""} onChange={(e) => editField("assetSerial", e.target.value.toUpperCase())} placeholder="IF VISIBLE"
                            style={{ width: "100%", boxSizing: "border-box", padding: "8px 10px", borderRadius: 8, border: `1.5px solid ${C.line}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 12, color: C.charcoal }} />
                        </div>
                        <div style={{ flex: 1 }}>
                          <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.charcoalSoft, marginBottom: 4 }}>ERF / UNIT</div>
                          <input value={pending.assetErf || ""} onChange={(e) => editField("assetErf", e.target.value.toUpperCase())} placeholder="IF APPLICABLE"
                            style={{ width: "100%", boxSizing: "border-box", padding: "8px 10px", borderRadius: 8, border: `1.5px solid ${C.line}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 12, color: C.charcoal }} />
                        </div>
                      </div>
                      <div>
                        <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.charcoalSoft, marginBottom: 4 }}>ZONE / STREET</div>
                        <input value={pending.assetZone || ""} onChange={(e) => editField("assetZone", e.target.value.toUpperCase())} placeholder="E.G. CUSSONIA WAY"
                          style={{ width: "100%", boxSizing: "border-box", padding: "8px 10px", borderRadius: 8, border: `1.5px solid ${C.line}`, fontFamily: "'Inter',sans-serif", fontSize: 12, color: C.charcoal }} />
                      </div>
                      <div>
                        <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.charcoalSoft, marginBottom: 4 }}>ACCESS NOTES</div>
                        <input value={pending.assetAccessNotes || ""} onChange={(e) => editField("assetAccessNotes", e.target.value.toUpperCase())} placeholder="GATE CODES, KEYS, DOGS, WHO TO CALL"
                          style={{ width: "100%", boxSizing: "border-box", padding: "8px 10px", borderRadius: 8, border: `1.5px solid ${C.line}`, fontFamily: "'Inter',sans-serif", fontSize: 12, color: C.charcoal }} />
                      </div>

                      {FILTRATION_TYPES.has(pending.assetType) && (
                        <>
                          {[
                            { key: "assetLastCartridge", label: "DATE OF LAST CARTRIDGE REPLACEMENT", months: CARTRIDGE_INTERVAL_MONTHS },
                            { key: "assetLastFlush", label: "DATE OF LAST FLUSH / BACKWASH", months: FLUSH_INTERVAL_MONTHS },
                          ].map((f) => {
                            const st = dueStatus(pending[f.key], f.months);
                            return (
                              <div key={f.key}>
                                <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.charcoalSoft, marginBottom: 4 }}>
                                  {f.label}
                                </div>
                                <input
                                  type="date"
                                  value={pending[f.key] || ""}
                                  onChange={(e) => editField(f.key, e.target.value)}
                                  style={{
                                    width: "100%", boxSizing: "border-box", padding: "8px 10px", borderRadius: 8,
                                    border: `1.5px solid ${st?.overdue ? C.flag : C.line}`,
                                    fontFamily: "'IBM Plex Mono',monospace", fontSize: 12.5, color: C.charcoal
                                  }}
                                />
                                {st && (
                                  <div style={{
                                    marginTop: 5, padding: "6px 9px", borderRadius: 7,
                                    background: st.overdue ? C.flagSoft : C.approveSoft,
                                    fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 600,
                                    color: st.overdue ? C.flag : C.approve
                                  }}>
                                    {st.overdue
                                      ? `OVERDUE BY ${st.monthsOverdue} MONTH${st.monthsOverdue === 1 ? "" : "S"} \u2014 was due ${st.dueDate}`
                                      : `Next due ${st.dueDate}`}
                                  </div>
                                )}
                              </div>
                            );
                          })}

                          <div>
                            <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.charcoalSoft, marginBottom: 4 }}>
                              CARTRIDGE CONDITION
                            </div>
                            <div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>
                              {CARTRIDGE_CONDITIONS.map((cond) => {
                                const active = pending.assetCartridgeCondition === cond;
                                const bad = cond === "FOULED" || cond === "NEEDS REPLACEMENT";
                                const col = bad ? C.flag : cond === "DISCOLOURED" ? C.review : C.approve;
                                return (
                                  <button key={cond} onClick={() => editField("assetCartridgeCondition", cond)} style={{
                                    padding: "7px 11px", borderRadius: 999, cursor: "pointer",
                                    border: `1.5px solid ${active ? col : C.line}`,
                                    background: active ? `${col}18` : "#fff",
                                    fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: active ? 700 : 600,
                                    color: active ? col : C.charcoalSoft
                                  }}>{cond}</button>
                                );
                              })}
                            </div>
                          </div>

                          <div>
                            <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.charcoalSoft, marginBottom: 4 }}>
                              FILTER SIZE / SPEC
                            </div>
                            <input value={pending.assetFilterSpec || ""} onChange={(e) => editField("assetFilterSpec", e.target.value.toUpperCase())}
                              placeholder='E.G. 10" 5 MICRON'
                              style={{ width: "100%", boxSizing: "border-box", padding: "8px 10px", borderRadius: 8, border: `1.5px solid ${C.line}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 12, color: C.charcoal }} />
                          </div>
                        </>
                      )}

                      {!FILTRATION_TYPES.has(pending.assetType) && SERVICEABLE_TYPES.has(pending.assetType) && (() => {
                        const st = serviceStatus(pending.assetType, pending.assetLastService);
                        return (
                          <div>
                            <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.charcoalSoft, marginBottom: 4 }}>
                              DATE OF LAST SERVICE
                            </div>
                            <input
                              type="date"
                              value={pending.assetLastService || ""}
                              onChange={(e) => editField("assetLastService", e.target.value)}
                              style={{
                                width: "100%", boxSizing: "border-box", padding: "8px 10px", borderRadius: 8,
                                border: `1.5px solid ${st?.overdue ? C.flag : C.line}`,
                                fontFamily: "'IBM Plex Mono',monospace", fontSize: 12.5, color: C.charcoal
                              }}
                            />
                            {st ? (
                              <div style={{
                                marginTop: 5, padding: "6px 9px", borderRadius: 7,
                                background: st.overdue ? C.flagSoft : C.approveSoft,
                                fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 600,
                                color: st.overdue ? C.flag : C.approve
                              }}>
                                {st.overdue
                                  ? `OVERDUE BY ${st.monthsOverdue} MONTH${st.monthsOverdue === 1 ? "" : "S"} — was due ${st.dueDate}`
                                  : `Next service due ${st.dueDate}`}
                              </div>
                            ) : (
                              <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, color: C.charcoalSoft, marginTop: 4 }}>
                                From the service label or certificate. Leave blank if unknown.
                              </div>
                            )}
                          </div>
                        );
                      })()}
                    </div>
                  )}

                  {(!pending.assetType || !pending.assetDescription) && (
                    <p style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, color: C.review, textAlign: "center", marginTop: 10 }}>
                      {!pending.assetCategory ? "Choose a category." : !pending.assetType ? "Choose an asset type." : "A description is required — it's how someone finds this later."}
                    </p>
                  )}
                </>
              )}

              {/* ---- AMR SURVEY: asset photo, serial, and reader screenshots ---- */}
              {pending.type === "amr" && (
                <>
                  <PhotoWell position={pending.position} timestamp={pending.timestamp} gps={pending.gps} photo={pending.photo} />
                  <div style={{
                    display: "flex", alignItems: "center", justifyContent: "center", gap: 5, marginTop: 8,
                    padding: "7px 10px", borderRadius: 8, background: C.paper
                  }}>
                    <MapPin size={12} color={pending.gps?.lat !== "Unknown" ? C.approve : C.review} />
                    <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 11.5, color: C.charcoal, fontWeight: 600 }}>
                      {pending.gps ? `${pending.gps.lat}, ${pending.gps.lng}` : "GPS unavailable"}
                    </span>
                  </div>

                  <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
                    <div style={{ flex: "0 0 40%" }}>
                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 600, color: C.charcoalSoft, marginBottom: 4 }}>ASSET TYPE</div>
                      <div style={{ display: "flex", gap: 5 }}>
                        {AMR_ASSET_TYPES.map((t) => {
                          const active = pending.amrAssetType === t;
                          return (
                            <button key={t} onClick={() => editField("amrAssetType", t)} style={{
                              flex: 1, padding: "8px 4px", borderRadius: 8, cursor: "pointer",
                              border: `1.5px solid ${active ? C.primary : C.line}`,
                              background: active ? "#E7F2FE" : "#fff",
                              fontFamily: "'Inter',sans-serif", fontSize: 11, fontWeight: 600,
                              color: active ? C.primary : C.charcoalSoft
                            }}>{t}</button>
                          );
                        })}
                      </div>
                    </div>
                    <div style={{ flex: 1 }}>
                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 600, color: C.charcoalSoft, marginBottom: 4 }}>SERIAL NUMBER</div>
                      <input
                        value={pending.amrSerial || ""}
                        onChange={(e) => editField("amrSerial", e.target.value)}
                        placeholder="Type if not on photo"
                        style={{ width: "100%", boxSizing: "border-box", padding: "8px 10px", borderRadius: 8, border: `1.5px solid ${C.line}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 13, fontWeight: 600, color: C.charcoal }}
                      />
                    </div>
                  </div>

                  <div style={{ marginTop: 14 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 7 }}>
                      <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, fontWeight: 700, color: C.charcoal }}>
                        Reader screenshots ({(pending.amrShots || []).length}/4)
                      </span>
                      {(pending.amrShots || []).length < 4 && (
                        <button onClick={() => openPicker("amrShot")} disabled={stage === "scanning"} style={{
                          border: "none", background: "none", color: C.primary, cursor: "pointer",
                          fontFamily: "'Inter',sans-serif", fontWeight: 600, fontSize: 11.5,
                          display: "flex", alignItems: "center", gap: 4
                        }}>
                          {stage === "scanning" ? <Loader2 size={13} className="spin" /> : <Camera size={13} />} Add
                        </button>
                      )}
                    </div>
                    {(pending.amrShots || []).length === 0 ? (
                      <>
                        <div style={{
                          padding: "14px 10px", borderRadius: 10, border: `1.5px dashed ${C.line}`, background: C.paper,
                          textAlign: "center", fontFamily: "'Inter',sans-serif", fontSize: 11, color: C.charcoalSoft
                        }}>
                          Add up to 4 screenshots from the handheld reader
                        </div>
                        <button
                          onClick={() => editField("noMetersAllocated", !pending.noMetersAllocated)}
                          style={{
                            width: "100%", marginTop: 8, padding: "11px", borderRadius: 9, cursor: "pointer",
                            border: `1.5px solid ${pending.noMetersAllocated ? C.review : C.line}`,
                            background: pending.noMetersAllocated ? C.reviewSoft : "#fff",
                            fontFamily: "'Inter',sans-serif", fontSize: 11.5, fontWeight: 700,
                            color: pending.noMetersAllocated ? C.review : C.charcoalSoft,
                            display: "flex", alignItems: "center", justifyContent: "center", gap: 7
                          }}>
                          {pending.noMetersAllocated ? <Check size={14} /> : <AlertTriangle size={14} />}
                          NO METERS ALLOCATED
                        </button>
                        {pending.noMetersAllocated && (
                          <input
                            value={pending.noMetersNote || ""}
                            onChange={(e) => editField("noMetersNote", e.target.value)}
                            placeholder="Why? Faulty, out of range, newly installed\u2026"
                            style={{
                              width: "100%", boxSizing: "border-box", marginTop: 7, padding: "9px 11px",
                              borderRadius: 8, border: `1.5px solid ${C.line}`,
                              fontFamily: "'Inter',sans-serif", fontSize: 12, color: C.charcoal
                            }}
                          />
                        )}
                      </>
                    ) : (
                      <div style={{ display: "flex", flexDirection: "column", gap: 7 }}>
                        {pending.amrShots.map((shot, i) => (
                          <div key={i} style={{ padding: 9, borderRadius: 9, background: C.paper }}>
                            <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 8 }}>
                              <img src={shot.photo} alt={`Screenshot ${i + 1}`} style={{ width: 34, height: 44, objectFit: "cover", borderRadius: 5, flexShrink: 0 }} />
                              <div style={{ flex: 1 }}>
                                <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, fontWeight: 600, color: C.charcoal }}>
                                  Screenshot {i + 1} · {(shot.meters || []).length} meters
                                </div>
                                {shot.concentratorId && (
                                  <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 10, color: C.charcoalSoft }}>
                                    ID: {shot.concentratorId}
                                  </div>
                                )}
                              </div>
                              <button onClick={() => editField("amrShots", pending.amrShots.filter((_, j) => j !== i))} style={{
                                border: "none", background: "none", cursor: "pointer", padding: 4
                              }}>
                                <X size={13} color={C.charcoalSoft} />
                              </button>
                            </div>

                            {(shot.meters || []).length > 0 && (
                              <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                                <div style={{ display: "grid", gridTemplateColumns: "1.3fr 1fr 0.8fr 20px", gap: 4, fontFamily: "'Inter',sans-serif", fontSize: 8.5, fontWeight: 700, color: C.charcoalSoft, padding: "0 2px" }}>
                                  <span>SERIAL</span><span>MODEL</span><span>SIGNAL</span><span />
                                </div>
                                {shot.meters.map((m, mi) => {
                                  const updateMeter = (patch) => {
                                    const shots = pending.amrShots.map((s, si) =>
                                      si !== i ? s : { ...s, meters: s.meters.map((mm, mmi) => (mmi === mi ? { ...mm, ...patch } : mm)) }
                                    );
                                    editField("amrShots", shots);
                                  };
                                  const removeMeter = () => {
                                    const shots = pending.amrShots.map((s, si) =>
                                      si !== i ? s : { ...s, meters: s.meters.filter((_, mmi) => mmi !== mi) }
                                    );
                                    editField("amrShots", shots);
                                  };
                                  return (
                                    <div key={mi} style={{ display: "grid", gridTemplateColumns: "1.3fr 1fr 0.8fr 20px", gap: 4, alignItems: "center" }}>
                                      <input value={m.serial || ""} onChange={(e) => updateMeter({ serial: e.target.value })}
                                        style={{ width: "100%", boxSizing: "border-box", padding: "4px 5px", borderRadius: 5, border: `1px solid ${C.line}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 10, color: C.charcoal, background: "#fff" }} />
                                      <input value={m.model || ""} onChange={(e) => updateMeter({ model: e.target.value })} placeholder="—"
                                        style={{ width: "100%", boxSizing: "border-box", padding: "4px 5px", borderRadius: 5, border: `1px solid ${C.line}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 10, color: C.charcoalSoft, background: "#fff" }} />
                                      {(() => {
                                        const band = signalBand(m.signal);
                                        return (
                                          <input value={m.signal || ""} onChange={(e) => updateMeter({ signal: e.target.value })} placeholder="—"
                                            title={band ? band.label : ""}
                                            style={{
                                              width: "100%", boxSizing: "border-box", padding: "4px 5px", borderRadius: 5,
                                              border: `1px solid ${band ? band.colour : C.line}`,
                                              background: band ? band.soft : "#fff",
                                              fontFamily: "'IBM Plex Mono',monospace", fontSize: 10,
                                              color: band ? band.colour : C.charcoalSoft, fontWeight: band ? 700 : 400
                                            }} />
                                        );
                                      })()}
                                      <button onClick={removeMeter} style={{ border: "none", background: "none", cursor: "pointer", padding: 0, display: "flex", justifyContent: "center" }}>
                                        <X size={10} color={C.charcoalSoft} />
                                      </button>
                                    </div>
                                  );
                                })}
                                <button
                                  onClick={() => {
                                    const shots = pending.amrShots.map((s, si) =>
                                      si !== i ? s : { ...s, meters: [...(s.meters || []), { serial: "", model: "", signal: "" }] }
                                    );
                                    editField("amrShots", shots);
                                  }}
                                  style={{
                                    border: "none", background: "none", color: C.primary, cursor: "pointer",
                                    fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 600, padding: "3px 0", textAlign: "left"
                                  }}>
                                  + Add missing meter
                                </button>
                              </div>
                            )}
                          </div>
                        ))}
                        {(() => {
                          const all = pending.amrShots.flatMap((s) => s.meters || []);
                          const bands = { GOOD: 0, FAIR: 0, WEAK: 0 };
                          all.forEach((m) => { const b = signalBand(m.signal); if (b) bands[b.label]++; });
                          return (
                            <div style={{ padding: "9px 11px", borderRadius: 8, background: C.paperDeep }}>
                              <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, fontWeight: 700, color: C.charcoal, textAlign: "center", marginBottom: 7 }}>
                                {all.length} meters linked to this {pending.amrAssetType?.toLowerCase() || "asset"}
                              </div>
                              <div style={{ display: "flex", gap: 5, justifyContent: "center" }}>
                                {[["GOOD", C.approve, C.approveSoft], ["FAIR", C.review, C.reviewSoft], ["WEAK", C.flag, C.flagSoft]].map(([k, col, soft]) => (
                                  <span key={k} style={{
                                    padding: "3px 10px", borderRadius: 999, background: soft,
                                    fontFamily: "'Inter',sans-serif", fontSize: 10, fontWeight: 700, color: col
                                  }}>{bands[k]} {k}</span>
                                ))}
                              </div>
                            </div>
                          );
                        })()}

                        {(() => {
                          const cmp = compareToRegister(
                            survey.amrRegister,
                            pending.amrSerial,
                            pending.amrShots.flatMap((s) => s.meters || [])
                          );
                          if (!cmp) return null;
                          if (cmp.unknownAsset) {
                            return (
                              <div style={{
                                padding: "9px 11px", borderRadius: 8, background: C.reviewSoft,
                                fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 600, color: C.review, textAlign: "center"
                              }}>
                                {pending.amrSerial ? `${pending.amrSerial} is not on the register` : "Enter the asset serial to check against the register"}
                              </div>
                            );
                          }
                          const ok = cmp.missing.length === 0;
                          return (
                            <div style={{
                              padding: "10px 12px", borderRadius: 9,
                              background: ok ? C.approveSoft : C.flagSoft,
                              border: `1px solid ${ok ? C.approve : C.flag}55`
                            }}>
                              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                                <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: ok ? C.approve : C.flag }}>
                                  {ok ? "ALL EXPECTED METERS FOUND" : `${cmp.missing.length} EXPECTED METER${cmp.missing.length === 1 ? "" : "S"} MISSING`}
                                </span>
                                <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 11, fontWeight: 700, color: ok ? C.approve : C.flag }}>
                                  {cmp.matched.length}/{cmp.expected.length}
                                </span>
                              </div>
                              {cmp.missing.length > 0 && (
                                <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 10, color: C.charcoal, marginTop: 5, lineHeight: 1.5 }}>
                                  {cmp.missing.join(", ")}
                                </div>
                              )}
                              {cmp.extra.length > 0 && (
                                <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9.5, color: C.charcoalSoft, marginTop: 5 }}>
                                  {cmp.extra.length} not on the register: {cmp.extra.join(", ")}
                                </div>
                              )}
                            </div>
                          );
                        })()}
                      </div>
                    )}
                  </div>
                </>
              )}

              {/* ---- METER REPLACEMENT: old and new side by side ---- */}
              {pending.type === "replacement" && (
                <>
                  {[
                    { key: "oldMeter", label: "OLD METER (removed)", tone: C.review, toneSoft: C.reviewSoft },
                    { key: "newMeter", label: "NEW METER (installed)", tone: C.approve, toneSoft: C.approveSoft },
                  ].map((side) => {
                    const data = pending[side.key];
                    return (
                      <div key={side.key} style={{
                        marginBottom: 12, padding: 11, borderRadius: 12,
                        border: `1.5px solid ${data ? side.tone : C.line}`,
                        background: data ? side.toneSoft : C.paper
                      }}>
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: data ? 9 : 0 }}>
                          <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 11, fontWeight: 700, color: side.tone, letterSpacing: 0.3 }}>
                            {side.label}
                          </span>
                          <button
                            onClick={() => openPicker(side.key)}
                            style={{
                              border: "none", background: "none", color: C.primary, cursor: "pointer",
                              fontFamily: "'Inter',sans-serif", fontWeight: 600, fontSize: 11,
                              display: "flex", alignItems: "center", gap: 4
                            }}>
                            <Camera size={12} /> {data ? "Retake" : "Add"}
                          </button>
                        </div>
                        {data && (
                          <div style={{ display: "flex", gap: 9 }}>
                            <img src={data.photo} alt={side.label} style={{ width: 68, height: 52, objectFit: "cover", borderRadius: 7, flexShrink: 0 }} />
                            <div style={{ flex: 1, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}>
                              <div>
                                <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9, fontWeight: 600, color: C.charcoalSoft }}>SERIAL</div>
                                <input
                                  value={data.serial || ""}
                                  onChange={(e) => editField(side.key, { ...data, serial: e.target.value })}
                                  placeholder="—"
                                  style={{ width: "100%", boxSizing: "border-box", padding: "5px 7px", borderRadius: 6, border: `1.5px solid ${C.line}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 11.5, color: C.charcoal, background: "#fff" }}
                                />
                              </div>
                              <div>
                                <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9, fontWeight: 600, color: C.charcoalSoft }}>READING (m³)</div>
                                <input
                                  value={data.reading || ""}
                                  onChange={(e) => editField(side.key, { ...data, reading: e.target.value })}
                                  placeholder="—"
                                  style={{ width: "100%", boxSizing: "border-box", padding: "5px 7px", borderRadius: 6, border: `1.5px solid ${C.line}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 11.5, color: C.charcoal, background: "#fff" }}
                                />
                              </div>
                            </div>
                          </div>
                        )}
                      </div>
                    );
                  })}
                  <div style={{
                    display: "flex", alignItems: "center", justifyContent: "center", gap: 5,
                    padding: "8px 10px", borderRadius: 8, background: C.paper
                  }}>
                    <MapPin size={12} color={pending.gps?.lat !== "Unknown" ? C.approve : C.review} />
                    <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 11.5, color: C.charcoal, fontWeight: 600 }}>
                      {pending.gps ? `${pending.gps.lat}, ${pending.gps.lng}` : "GPS unavailable"}
                    </span>
                  </div>
                  {(!pending.oldMeter || !pending.newMeter) && (
                    <p style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, color: C.review, textAlign: "center", marginTop: 8 }}>
                      {!pending.oldMeter ? "Old meter photo still needed." : "New meter photo still needed."}
                    </p>
                  )}
                </>
              )}

              {/* ---- FIDO FEEDBACK: the screenshot + session type selection ---- */}
              {pending.type === "fido" && pending.fido && (
                <>
                  <div style={{
                    position: "relative", width: "100%", height: 170, borderRadius: 12, overflow: "hidden", background: "#fff",
                    border: `1px solid ${C.line}`
                  }}>
                    <img src={pending.fido.photo} alt="FIDO feedback"
                      style={{ width: "100%", height: "100%", objectFit: "contain", display: "block" }} />
                  </div>
                  <div style={{
                    display: "flex", alignItems: "center", justifyContent: "center", gap: 5, marginTop: 8,
                    padding: "7px 10px", borderRadius: 8, background: C.paper
                  }}>
                    <MapPin size={12} color={pending.fido.gps ? C.approve : C.review} />
                    <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 11.5, color: C.charcoal, fontWeight: 600 }}>
                      {pending.fido.gps ? `${pending.fido.gps.lat}, ${pending.fido.gps.lng}` : "GPS unavailable"}
                    </span>
                  </div>
                  <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, fontWeight: 600, color: C.charcoal, margin: "12px 0 7px" }}>
                    What kind of session was this?
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 7 }}>
                    {FIDO_SESSION_TYPES.map((t) => {
                      const active = pending.fido.sessionType === t;
                      return (
                        <button
                          key={t}
                          onClick={() => editField("fido", { ...pending.fido, sessionType: t })}
                          style={{
                            padding: "10px 8px", borderRadius: 9, cursor: "pointer",
                            border: `1.5px solid ${active ? C.primary : C.line}`,
                            background: active ? "#E7F2FE" : "#fff",
                            fontFamily: "'Inter',sans-serif", fontSize: 11, fontWeight: 600,
                            color: active ? C.primary : C.charcoalSoft, textAlign: "center"
                          }}>
                          {t}
                        </button>
                      );
                    })}
                  </div>
                  {!pending.fido.sessionType && (
                    <p style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, color: C.review, textAlign: "center", marginTop: 8 }}>
                      Select the session type before saving.
                    </p>
                  )}
                </>
              )}

              {(pending.type === "meter" || pending.type === "replacement") && (
                <>
                  <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.charcoalSoft, margin: "14px 0 6px" }}>
                    METER TYPE
                  </div>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>
                    {METER_TYPES.map((t) => {
                      const active = pending.meterType === t;
                      return (
                        <button key={t} onClick={() => editField("meterType", t)} style={{
                          padding: "7px 11px", borderRadius: 999, cursor: "pointer",
                          border: `1.5px solid ${active ? C.primary : C.line}`,
                          background: active ? "#E7F2FE" : "#fff",
                          fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: active ? 700 : 600,
                          color: active ? C.primary : C.charcoalSoft
                        }}>{t}</button>
                      );
                    })}
                  </div>

                  <div style={{
                    marginTop: 14, borderRadius: 10, border: `1.5px solid ${(pending.conditionFlags || []).length ? C.review : C.line}`,
                    background: (pending.conditionFlags || []).length ? C.reviewSoft : C.paper, padding: 11
                  }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 8 }}>
                      <AlertTriangle size={13} color={(pending.conditionFlags || []).length ? C.review : C.charcoalSoft} />
                      <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 11, fontWeight: 700, color: (pending.conditionFlags || []).length ? C.review : C.charcoal }}>
                        ANY PROBLEM WITH THIS METER?
                      </span>
                    </div>
                    <div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>
                      {CONDITION_FLAGS.map((f) => {
                        const on = (pending.conditionFlags || []).includes(f);
                        return (
                          <button key={f}
                            onClick={() => {
                              const cur = pending.conditionFlags || [];
                              editField("conditionFlags", on ? cur.filter((x) => x !== f) : [...cur, f]);
                            }}
                            style={{
                              padding: "6px 10px", borderRadius: 999, cursor: "pointer",
                              border: `1.5px solid ${on ? C.flag : C.line}`,
                              background: on ? C.flagSoft : "#fff",
                              fontFamily: "'Inter',sans-serif", fontSize: 10, fontWeight: on ? 700 : 600,
                              color: on ? C.flag : C.charcoalSoft
                            }}>{f}</button>
                        );
                      })}
                    </div>
                    {(pending.conditionFlags || []).length > 0 && (
                      <input
                        value={pending.conditionNote || ""}
                        onChange={(e) => editField("conditionNote", e.target.value)}
                        placeholder="Add detail — what exactly is wrong?"
                        style={{
                          width: "100%", boxSizing: "border-box", marginTop: 8, padding: "8px 10px",
                          borderRadius: 8, border: `1.5px solid ${C.line}`,
                          fontFamily: "'Inter',sans-serif", fontSize: 12, color: C.charcoal
                        }}
                      />
                    )}
                  </div>
                </>
              )}

              <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
                <button
                  onClick={() => {
                    if (!confirmDiscard) {
                      setConfirmDiscard(true);
                      setTimeout(() => setConfirmDiscard(false), 3500);
                      return;
                    }
                    setConfirmDiscard(false);
                    setStage("idle");
                    setPending(null);
                    setCaptureError("");
                  }}
                  style={{
                    flex: confirmDiscard ? 1 : "0 0 44px", height: 44, borderRadius: 10,
                    border: `1.5px solid ${confirmDiscard ? C.flag : C.line}`,
                    background: confirmDiscard ? C.flagSoft : "#fff",
                    cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 6,
                    transition: "all 0.15s"
                  }}>
                  <RotateCcw size={16} color={confirmDiscard ? C.flag : C.charcoalSoft} />
                  {confirmDiscard && (
                    <span style={{ fontFamily: "'Inter',sans-serif", fontWeight: 600, fontSize: 12, color: C.flag }}>
                      Tap again to discard
                    </span>
                  )}
                </button>
                <button onClick={save} style={{
                  flex: 1, height: 44, borderRadius: 10, border: "none", cursor: "pointer",
                  background: C.primary, color: "#fff", fontFamily: "'Inter',sans-serif", fontWeight: 600, fontSize: 14.5,
                  display: "flex", alignItems: "center", justifyContent: "center", gap: 6
                }}>
                  Save Point <ArrowRight size={15} />
                </button>
              </div>
            </>
          )}
        </div>

        <div style={{
          borderTop: `1px solid ${C.line}`, padding: "10px 16px", display: "flex",
          alignItems: "center", justifyContent: "space-between"
        }}>
          <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 11.5, color: C.charcoalSoft }}>
            {captures.length} captured this run
          </span>
          <button onClick={() => setScreen("office")} style={{
            border: "none", background: "none", cursor: "pointer", color: C.primary,
            fontFamily: "'Inter',sans-serif", fontSize: 12.5, fontWeight: 600,
            display: "flex", alignItems: "center", gap: 3
          }}>
            Finish survey <ChevronRight size={14} />
          </button>
        </div>
      </div>
    </div>
  );
}

/* ---------- Screen: Office Review ---------- */
function OfficeScreen({ survey, captures, setCaptures, onDeleteSurvey }) {
  const [postExport, setPostExport] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [selected, setSelected] = useState(null);
  const [filter, setFilter] = useState("all");
  const [reviewer, setReviewer] = useState("");
  const unsentCount = captures.filter((c) => !c.sentAt).length;
  const [exportScope, setExportScope] = useState(unsentCount > 0 ? "new" : "all");

  const counts = {
    total: captures.length,
    approved: captures.filter((c) => c.status === "approved").length,
    review: captures.filter((c) => c.status === "needs_review").length,
    flagged: captures.filter((c) => c.status === "flagged").length,
  };
  const pct = counts.total ? Math.round((counts.approved / counts.total) * 100) : 0;

  const [editMode, setEditMode] = useState(false);
  const [zoomPhoto, setZoomPhoto] = useState(null);
  const duplicateMeters = findDuplicateMeters(captures);
  const [toast, setToast] = useState(null);

  const flashToast = (text, tone) => {
    setToast({ text, tone });
    setTimeout(() => setToast(null), 2200);
  };
  const updateCapture = (id, patch) => {
    setCaptures((cs) => cs.map((c) => (c.id === id ? { ...c, ...patch } : c)));
    setSelected((s) => (s && s.id === id ? { ...s, ...patch } : s));
  };

  const filtered = captures.filter((c) => filter === "all" ? true : c.status === filter);

  const StatChip = ({ label, value, color, id }) => (
    <button onClick={() => setFilter(id)} style={{
      border: "none", background: "none", cursor: "pointer", textAlign: "left", padding: 0
    }}>
      <div style={{ fontFamily: "'Space Grotesk',sans-serif", fontSize: 24, fontWeight: 700, color }}>{value}</div>
      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, color: C.charcoalSoft, fontWeight: 600 }}>{label}</div>
    </button>
  );

  if (!captures.length) {
    return (
      <div style={{ maxWidth: 420, margin: "60px auto", textAlign: "center" }}>
        <LayoutGrid size={34} color={C.line} style={{ marginBottom: 10 }} />
        <p style={{ fontFamily: "'Inter',sans-serif", color: C.charcoalSoft, fontSize: 14 }}>
          Nothing captured yet. Head to Field Capture to record your first meter.
        </p>
      </div>
    );
  }

  return (
    <div style={{ maxWidth: 960, margin: "0 auto", padding: "24px 4px" }}>
      <div style={{
        background: "#fff", border: `1px solid ${C.line}`, borderRadius: 18, padding: "22px 26px",
        display: "flex", alignItems: "center", gap: 30, marginBottom: 22, flexWrap: "wrap"
      }}>
        <MeterDial value={pct} />
        <div style={{ flex: 1, minWidth: 220 }}>
          <h2 style={{ fontFamily: "'Space Grotesk',sans-serif", fontSize: 19, fontWeight: 700, color: C.charcoal, margin: 0 }}>
            {survey.siteName || "Untitled Site"}
          </h2>
          <p style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 12, color: C.charcoalSoft, margin: "3px 0 14px" }}>
            {survey.surveyName}
          </p>
          <div style={{ display: "flex", gap: 26 }}>
            <StatChip label="Captured" value={counts.total} color={C.charcoal} id="all" />
            <StatChip label="Approved" value={counts.approved} color={C.approve} id="approved" />
            <StatChip label="Review" value={counts.review} color={C.review} id="needs_review" />
            <StatChip label="Flagged" value={counts.flagged} color={C.flag} id="flagged" />
          </div>
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 8, minWidth: 210 }}>
          <label style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 600, color: C.charcoalSoft }}>
            Reviewed by
          </label>
          <input
            value={reviewer}
            onChange={(e) => setReviewer(e.target.value)}
            placeholder="e.g. Rijn Geyser"
            style={{
              padding: "8px 10px", borderRadius: 8, border: `1.5px solid ${C.line}`,
              fontFamily: "'Inter',sans-serif", fontSize: 13, color: C.charcoal, outline: "none"
            }}
          />

          <div style={{ display: "flex", gap: 4, background: C.paperDeep, padding: 3, borderRadius: 8 }}>
            {[
              { id: "new", label: `New (${unsentCount})` },
              { id: "all", label: `All (${captures.length})` },
            ].map((opt) => (
              <button key={opt.id} onClick={() => setExportScope(opt.id)} style={{
                flex: 1, border: "none", cursor: "pointer", padding: "6px 8px", borderRadius: 6,
                fontFamily: "'Inter',sans-serif", fontSize: 11, fontWeight: 600,
                background: exportScope === opt.id ? "#fff" : "transparent",
                color: exportScope === opt.id ? C.charcoal : C.charcoalSoft,
                boxShadow: exportScope === opt.id ? "0 1px 3px rgba(43,47,51,0.12)" : "none"
              }}>
                {opt.label}
              </button>
            ))}
          </div>

          {(() => {
            const scoped = exportScope === "new" ? captures.filter((c) => !c.sentAt) : captures;
            const markSent = () => {
              if (exportScope !== "new") return;
              const ids = new Set(scoped.map((c) => c.id));
              setCaptures((cs) => cs.map((c) => (ids.has(c.id) ? { ...c, sentAt: Date.now() } : c)));
            };
            const disabled = scoped.length === 0;
            return (
              <>
                <button
                  disabled={disabled}
                  onClick={() => { exportExcel(survey, scoped, reviewer); markSent(); setPostExport(true); }}
                  style={{
                    display: "flex", alignItems: "center", justifyContent: "center", gap: 7,
                    padding: "9px 12px", borderRadius: 9, border: `1.5px solid ${C.line}`,
                    background: disabled ? C.paperDeep : "#fff", cursor: disabled ? "not-allowed" : "pointer",
                    color: disabled ? C.charcoalSoft : C.charcoal, fontFamily: "'Inter',sans-serif", fontWeight: 600, fontSize: 12.5
                  }}>
                  <FileSpreadsheet size={14} color={disabled ? C.charcoalSoft : C.approve} /> Download Excel
                </button>
                {survey.taskType === "assets" && (
                  <button
                    disabled={disabled}
                    onClick={() => { exportAssetCSV(survey, scoped); markSent(); setPostExport(true); }}
                    style={{
                      display: "flex", alignItems: "center", justifyContent: "center", gap: 7,
                      padding: "9px 12px", borderRadius: 9, border: `1.5px solid ${C.line}`,
                      background: disabled ? C.paperDeep : "#fff", cursor: disabled ? "not-allowed" : "pointer",
                      color: disabled ? C.charcoalSoft : C.charcoal, fontFamily: "'Inter',sans-serif", fontWeight: 600, fontSize: 12.5
                    }}>
                    <FileSpreadsheet size={14} color={disabled ? C.charcoalSoft : C.primary} /> Download CSV
                  </button>
                )}
                <button
                  disabled={disabled}
                  onClick={() => {
                    if (survey.taskType === "assets") {
                      printAssetRegister(survey, scoped, reviewer);
                    } else if (survey.taskType === "fido2") {
                      printFidoReport(survey, scoped, reviewer);
                    } else {
                      printReport(survey, scoped, reviewer, counts, pct);
                    }
                    markSent();
                    setPostExport(true);
                  }}
                  style={{
                    display: "flex", alignItems: "center", justifyContent: "center", gap: 7,
                    padding: "9px 12px", borderRadius: 9, border: "none",
                    background: disabled ? C.line : C.primary, cursor: disabled ? "not-allowed" : "pointer",
                    color: "#fff", fontFamily: "'Inter',sans-serif", fontWeight: 600, fontSize: 12.5
                  }}>
                  <Send size={14} /> {survey.taskType === "assets" ? "Asset Register PDF" : survey.taskType === "fido2" ? "Leak Analysis PDF" : "Export PDF"}
                </button>
                {disabled && exportScope === "new" && (
                  <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, color: C.charcoalSoft, textAlign: "center" }}>
                    Nothing new to send since last export.
                  </span>
                )}
              </>
            );
          })()}
        </div>
      </div>

      {postExport && (
        <div style={{
          background: "#fff", border: `1px solid ${C.line}`, borderLeft: `3px solid ${C.primary}`,
          borderRadius: 12, padding: "14px 18px", marginBottom: 22,
          display: "flex", alignItems: "center", justifyContent: "space-between", gap: 14, flexWrap: "wrap"
        }}>
          <div>
            <div style={{ fontFamily: "'Inter',sans-serif", fontWeight: 600, fontSize: 13.5, color: C.charcoal }}>
              Report exported. All done with this site?
            </div>
            <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, color: C.charcoalSoft, marginTop: 2 }}>
              Deleting removes it from the shared list for everyone. Your downloaded report is unaffected.
            </div>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <button
              onClick={() => { setPostExport(false); setConfirmDelete(false); }}
              style={{
                padding: "9px 14px", borderRadius: 9, border: `1.5px solid ${C.line}`, background: "#fff",
                color: C.charcoal, fontFamily: "'Inter',sans-serif", fontWeight: 600, fontSize: 12.5, cursor: "pointer"
              }}>
              Keep site open
            </button>
            <button
              onClick={() => {
                if (!confirmDelete) {
                  setConfirmDelete(true);
                  setTimeout(() => setConfirmDelete(false), 3500);
                  return;
                }
                onDeleteSurvey();
              }}
              style={{
                padding: "9px 14px", borderRadius: 9, border: "none",
                background: C.flag, color: "#fff",
                fontFamily: "'Inter',sans-serif", fontWeight: 600, fontSize: 12.5, cursor: "pointer",
                display: "flex", alignItems: "center", gap: 6
              }}>
              <Trash2 size={13} /> {confirmDelete ? "Tap again to confirm" : "Delete site"}
            </button>
          </div>
        </div>
      )}

      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
        <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 13, fontWeight: 600, color: C.charcoal }}>
          {filter === "all" ? "All meters" : STATUS_META[filter].label} · {filtered.length}
        </span>
        {filter !== "all" && (
          <button onClick={() => setFilter("all")} style={{
            border: "none", background: "none", color: C.primary, cursor: "pointer",
            fontFamily: "'Inter',sans-serif", fontSize: 12.5, fontWeight: 600
          }}>Clear filter</button>
        )}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(150px, 1fr))", gap: 12 }}>
        {filtered.map((c) => {
          const meta = STATUS_META[c.status];
          const Icon = meta.icon;
          return (
            <div key={c.id} onClick={() => { setSelected(c); setEditMode(false); }} style={{
              cursor: "pointer", borderRadius: 13, border: `1px solid ${C.line}`, overflow: "hidden", background: "#fff"
            }}>
              <PhotoWell size="small" photo={c.photo || c.resultShot?.photo || c.sessionShot?.photo || c.newMeter?.photo || c.oldMeter?.photo || c.fido?.photo || c.consumption?.photo || c.sensor?.photo} />
              <div style={{ padding: "9px 10px" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontWeight: 600, fontSize: 13, color: C.charcoal }}>{c.position}</span>
                  <Icon size={13} color={meta.color} />
                </div>
                <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 11, color: C.charcoalSoft, marginTop: 2, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                  <span style={{ display: "flex", alignItems: "center", gap: 4 }}>
                    {(!c.type || c.type === "meter") && <>{c.reading ? `${c.reading} m³` : "no reading"}</>}
                    {c.type === "sensor" && <><Radio size={10} color={C.primary} /> Sensor</>}
                    {c.type === "fido" && <><Activity size={10} color={C.primary} /> FIDO</>}
                    {c.type === "replacement" && <><RotateCcw size={10} color={C.primary} /> Replaced</>}
                    {c.type === "exception" && <><AlertTriangle size={10} color={C.review} /> Could not read</>}
                    {c.type === "fido2_deploy" && <><Radio size={10} color={C.primary} /> {c.bugSerial} out</>}
                    {c.type === "fido2_waypoint" && (() => {
                      const st = WAYPOINT_OUTCOMES[c.outcome];
                      return <><MapPin size={10} color={st ? st.colour : C.charcoalSoft} /> {c.waypointRef}</>;
                    })()}
                    {c.type === "fido2_retrieve" && (() => {
                      const st = FIDO_OUTCOMES[c.outcome];
                      return <><Radio size={10} color={st ? st.colour : C.charcoalSoft} /> {c.outcome || "no finding"}</>;
                    })()}
                    {c.type === "amr" && (c.noMetersAllocated
                      ? <><AlertTriangle size={10} color={C.review} /> No meters</>
                      : <><Radio size={10} color={C.primary} /> {(c.amrShots || []).reduce((n, s) => n + (s.meters || []).length, 0)} meters</>)}
                    {c.type === "asset" && <><LayoutGrid size={10} color={CATEGORY_COLOURS[c.assetCategory] || C.primary} /> {c.assetType || "Asset"}</>}
                    {c.type === "consumption" && <><Activity size={10} color={C.primary} /> Profile</>}
                  </span>
                  {c.sentAt && <Send size={10} color={C.charcoalSoft} />}
                </div>
                {c.sensor?.sessionId && (
                  <div style={{ display: "flex", alignItems: "center", gap: 3, marginTop: 3 }}>
                    <Radio size={10} color={C.primary} />
                    <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 10, color: C.primary }}>{c.sensor.sessionId}</span>
                  </div>
                )}
              </div>
            </div>
          );
        })}
      </div>

      {selected && (
        <div onClick={() => setSelected(null)} style={{
          position: "fixed", inset: 0, background: "rgba(43,47,51,0.5)", display: "flex",
          alignItems: "center", justifyContent: "center", padding: 20, zIndex: 50
        }}>
          <div onClick={(e) => e.stopPropagation()} style={{
            background: "#fff", borderRadius: 18, width: selected.type === "amr" ? 640 : 420, maxWidth: "100%", padding: 22,
            maxHeight: "88vh", overflowY: "auto"
          }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 12 }}>
              <div>
                <div style={{ fontFamily: "'Space Grotesk',sans-serif", fontWeight: 700, fontSize: 18, color: C.charcoal }}>
                  {selected.position}
                </div>
                <div style={{
                  display: "inline-flex", alignItems: "center", gap: 5, marginTop: 5, padding: "3px 9px",
                  borderRadius: 999, background: STATUS_META[selected.status].bg
                }}>
                  {(() => { const I = STATUS_META[selected.status].icon; return <I size={11} color={STATUS_META[selected.status].color} />; })()}
                  <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 11, fontWeight: 600, color: STATUS_META[selected.status].color }}>
                    {STATUS_META[selected.status].label}
                  </span>
                </div>
              </div>
              <button onClick={() => setSelected(null)} style={{ border: "none", background: "none", cursor: "pointer" }}>
                <X size={20} color={C.charcoalSoft} />
              </button>
            </div>

            <PhotoWell position={selected.position} timestamp={selected.timestamp} gps={selected.gps} photo={selected.photo || selected.fido?.photo || selected.consumption?.photo} />

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, margin: "14px 0" }}>
              {/* editable fields (edit mode) or static display */}
              {(selected.type === "meter" || selected.type === "replacement" || !selected.type) && (
                <div style={{ background: editMode ? "#fff" : C.paper, border: editMode ? `1.5px solid ${C.primary}` : "none", borderRadius: 9, padding: "8px 10px" }}>
                  <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, color: C.charcoalSoft, fontWeight: 600 }}>Meter Type</div>
                  {editMode ? (
                    <select value={selected.meterType || ""} onChange={(e) => updateCapture(selected.id, { meterType: e.target.value })}
                      style={{ width: "100%", border: "none", outline: "none", fontFamily: "'Inter',sans-serif", fontSize: 12, color: C.charcoal, fontWeight: 600, background: "transparent" }}>
                      <option value="">\u2014</option>
                      {METER_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
                    </select>
                  ) : (
                    <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 12.5, color: selected.meterType ? C.charcoal : C.charcoalSoft, fontWeight: 600 }}>
                      {selected.meterType || "\u2014"}
                    </div>
                  )}
                </div>
              )}
              {(!selected.type || selected.type === "meter") && (
                editMode ? (
                  <>
                    <div style={{ background: "#fff", border: `1.5px solid ${C.primary}`, borderRadius: 9, padding: "8px 10px" }}>
                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, color: C.charcoalSoft, fontWeight: 600 }}>Reading (m³)</div>
                      <input value={selected.reading} onChange={(e) => updateCapture(selected.id, { reading: e.target.value })}
                        style={{ width: "100%", border: "none", outline: "none", fontFamily: "'IBM Plex Mono',monospace", fontSize: 13, color: C.charcoal, fontWeight: 600, padding: 0, background: "transparent" }} />
                    </div>
                    <div style={{ background: "#fff", border: `1.5px solid ${C.primary}`, borderRadius: 9, padding: "8px 10px" }}>
                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, color: C.charcoalSoft, fontWeight: 600 }}>Serial</div>
                      <input value={selected.serial} onChange={(e) => updateCapture(selected.id, { serial: e.target.value })}
                        style={{ width: "100%", border: "none", outline: "none", fontFamily: "'IBM Plex Mono',monospace", fontSize: 13, color: C.charcoal, fontWeight: 600, padding: 0, background: "transparent" }} />
                    </div>
                  </>
                ) : (
                  <>
                    <div style={{ background: C.paper, borderRadius: 9, padding: "8px 10px" }}>
                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, color: C.charcoalSoft, fontWeight: 600 }}>Reading</div>
                      <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 13, color: C.charcoal, fontWeight: 600 }}>{selected.reading ? `${selected.reading} m³` : "—"}</div>
                    </div>
                    <div style={{ background: C.paper, borderRadius: 9, padding: "8px 10px" }}>
                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, color: C.charcoalSoft, fontWeight: 600 }}>Serial</div>
                      <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 13, color: C.charcoal, fontWeight: 600 }}>{selected.serial || "—"}</div>
                    </div>
                  </>
                )
              )}
              {selected.type === "fido" && (
                <>
                  <div style={{ background: C.paper, borderRadius: 9, padding: "8px 10px" }}>
                    <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, color: C.charcoalSoft, fontWeight: 600 }}>Type</div>
                    <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 13, color: C.charcoal, fontWeight: 600 }}>FIDO Feedback</div>
                  </div>
                  <div style={{ background: editMode ? "#fff" : C.paper, border: editMode ? `1.5px solid ${C.primary}` : "none", borderRadius: 9, padding: "8px 10px" }}>
                    <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, color: C.charcoalSoft, fontWeight: 600 }}>Session</div>
                    {editMode ? (
                      <select
                        value={selected.fido?.sessionType || ""}
                        onChange={(e) => updateCapture(selected.id, { fido: { ...selected.fido, sessionType: e.target.value } })}
                        style={{ width: "100%", border: "none", outline: "none", fontFamily: "'IBM Plex Mono',monospace", fontSize: 12, color: C.charcoal, fontWeight: 600, background: "transparent" }}>
                        <option value="">—</option>
                        {FIDO_SESSION_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
                      </select>
                    ) : (
                      <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 13, color: C.charcoal, fontWeight: 600 }}>{selected.fido?.sessionType || "—"}</div>
                    )}
                  </div>
                </>
              )}
              {selected.type === "asset" && (
                <>
                  <div style={{ background: C.paper, borderRadius: 9, padding: "8px 10px" }}>
                    <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, color: C.charcoalSoft, fontWeight: 600 }}>Category</div>
                    <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 12.5, fontWeight: 700, color: CATEGORY_COLOURS[selected.assetCategory] || C.charcoal }}>
                      {selected.assetCategory || "—"}
                    </div>
                  </div>
                  <div style={{ background: editMode ? "#fff" : C.paper, border: editMode ? `1.5px solid ${C.primary}` : "none", borderRadius: 9, padding: "8px 10px" }}>
                    <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, color: C.charcoalSoft, fontWeight: 600 }}>Asset Type</div>
                    {editMode && selected.assetCategory ? (
                      <select value={selected.assetType || ""} onChange={(e) => updateCapture(selected.id, { assetType: e.target.value })}
                        style={{ width: "100%", border: "none", outline: "none", fontFamily: "'Inter',sans-serif", fontSize: 12, color: C.charcoal, fontWeight: 600, background: "transparent" }}>
                        <option value="">—</option>
                        {(ASSET_CATEGORIES[selected.assetCategory] || []).map((t) => <option key={t} value={t}>{t}</option>)}
                      </select>
                    ) : (
                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 12.5, color: C.charcoal, fontWeight: 600 }}>{selected.assetType || "—"}</div>
                    )}
                  </div>
                </>
              )}
              {selected.type === "amr" && (
                <>
                  <div style={{ background: C.paper, borderRadius: 9, padding: "8px 10px" }}>
                    <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, color: C.charcoalSoft, fontWeight: 600 }}>Asset Type</div>
                    <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 13, color: C.charcoal, fontWeight: 600 }}>{selected.amrAssetType || "—"}</div>
                  </div>
                  <div style={{ background: editMode ? "#fff" : C.paper, border: editMode ? `1.5px solid ${C.primary}` : "none", borderRadius: 9, padding: "8px 10px" }}>
                    <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, color: C.charcoalSoft, fontWeight: 600 }}>Serial</div>
                    {editMode ? (
                      <input value={selected.amrSerial || ""} onChange={(e) => updateCapture(selected.id, { amrSerial: e.target.value })}
                        style={{ width: "100%", border: "none", outline: "none", fontFamily: "'IBM Plex Mono',monospace", fontSize: 13, color: C.charcoal, fontWeight: 600, padding: 0, background: "transparent" }} />
                    ) : (
                      <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 13, color: C.charcoal, fontWeight: 600 }}>{selected.amrSerial || "—"}</div>
                    )}
                  </div>
                </>
              )}
              {selected.type === "replacement" && (
                <>
                  <div style={{ background: C.paper, borderRadius: 9, padding: "8px 10px" }}>
                    <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, color: C.charcoalSoft, fontWeight: 600 }}>Type</div>
                    <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 13, color: C.charcoal, fontWeight: 600 }}>Meter Replacement</div>
                  </div>
                  <div style={{ background: C.paper, borderRadius: 9, padding: "8px 10px" }}>
                    <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, color: C.charcoalSoft, fontWeight: 600 }}>Position</div>
                    <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 13, color: C.charcoal, fontWeight: 600 }}>{selected.position}</div>
                  </div>
                </>
              )}
              {selected.type === "sensor" && (
                <>
                  <div style={{ background: C.paper, borderRadius: 9, padding: "8px 10px" }}>
                    <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, color: C.charcoalSoft, fontWeight: 600 }}>Type</div>
                    <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 13, color: C.charcoal, fontWeight: 600 }}>Sensor Deployment</div>
                  </div>
                  <div style={{ background: C.paper, borderRadius: 9, padding: "8px 10px" }}>
                    <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, color: C.charcoalSoft, fontWeight: 600 }}>Session</div>
                    <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 13, color: C.charcoal, fontWeight: 600 }}>{selected.sensor?.sessionId || "—"}</div>
                  </div>
                </>
              )}
              {[
                ["GPS", selected.gps ? `${selected.gps.lat}, ${selected.gps.lng}` : "—"],
                ["Date / Time", `${selected.timestamp.date}, ${selected.timestamp.time}`],
                ["Technician", selected.tech],
              ].map(([label, val]) => (
                <div key={label} style={{ background: C.paper, borderRadius: 9, padding: "8px 10px" }}>
                  <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, color: C.charcoalSoft, fontWeight: 600 }}>{label}</div>
                  <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 13, color: C.charcoal, fontWeight: 600 }}>{val}</div>
                </div>
              ))}
            </div>

            {(!selected.type || selected.type === "meter") && (
              <div style={{
                display: "flex", alignItems: "center", justifyContent: "space-between", padding: "9px 12px",
                borderRadius: 9, background: C.paperDeep, marginBottom: 16
              }}>
                <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 12, fontWeight: 600, color: C.charcoalSoft }}>AI Confidence</span>
                <MiniGauge value={selected.confidence} color={selected.confidence < 85 ? C.review : C.approve} />
              </div>
            )}

            {selected.type === "fido2_waypoint" && (
              <div style={{ marginBottom: 16 }}>
                {(() => {
                  const st = WAYPOINT_OUTCOMES[selected.outcome] || WAYPOINT_OUTCOMES.ONGOING;
                  return (
                    <div style={{
                      padding: "12px 14px", borderRadius: 10, background: st.soft,
                      border: `1px solid ${st.colour}55`, marginBottom: 11
                    }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                        <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 15, fontWeight: 700, color: C.charcoal }}>
                          {selected.waypointRef}
                        </span>
                        <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, fontWeight: 700, color: st.colour }}>
                          {selected.outcome}
                        </span>
                      </div>
                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, color: C.charcoalSoft, marginTop: 3 }}>
                        {selected.position}{selected.linkedBug ? ` \u00b7 raised by ${selected.linkedBug}` : ""}
                      </div>
                    </div>
                  );
                })()}

                {(selected.tests || []).map((t, i) => (
                  <div key={i} style={{ marginBottom: 11, padding: 10, borderRadius: 9, background: C.paper }}>
                    <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.charcoal, marginBottom: 6 }}>
                      {t.method}{t.distance ? ` \u00b7 leak at ${t.distance}` : ""}
                    </div>
                    <img src={t.photo} alt={t.method} onClick={() => setZoomPhoto(t.photo)}
                      style={{ width: "100%", maxHeight: 300, objectFit: "contain", borderRadius: 7, background: "#fff", cursor: "zoom-in" }} />
                    {t.note && (
                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, color: C.charcoalSoft, marginTop: 6 }}>{t.note}</div>
                    )}
                  </div>
                ))}

                {selected.pinpointPhoto && (
                  <div style={{ padding: 11, borderRadius: 9, background: C.flagSoft, border: `1px solid ${C.flag}55`, marginBottom: 11 }}>
                    <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, fontWeight: 700, color: C.flag, marginBottom: 7 }}>
                      LEAK POSITION MARKED
                    </div>
                    <img src={selected.pinpointPhoto} alt="pinpoint" onClick={() => setZoomPhoto(selected.pinpointPhoto)}
                      style={{ width: "100%", maxHeight: 220, objectFit: "cover", borderRadius: 7, cursor: "zoom-in" }} />
                    <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 12, fontWeight: 700, color: C.charcoal, marginTop: 7 }}>
                      {selected.pinpointGps ? `${selected.pinpointGps.lat}, ${selected.pinpointGps.lng}` : "\u2014"}
                    </div>
                  </div>
                )}

                <div>
                  <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9.5, fontWeight: 700, color: C.charcoalSoft, marginBottom: 3 }}>CONCLUSION</div>
                  {editMode ? (
                    <input value={selected.outcomeNote || ""} onChange={(e) => updateCapture(selected.id, { outcomeNote: e.target.value })}
                      style={{ width: "100%", boxSizing: "border-box", padding: "7px 9px", borderRadius: 7, border: `1.5px solid ${C.primary}`, fontFamily: "'Inter',sans-serif", fontSize: 11.5, color: C.charcoal }} />
                  ) : (
                    <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 12, color: selected.outcomeNote ? C.charcoal : C.charcoalSoft }}>
                      {selected.outcomeNote || "\u2014"}
                    </div>
                  )}
                </div>
              </div>
            )}

            {(selected.type === "fido2_deploy" || selected.type === "fido2_retrieve") && (
              <div style={{ marginBottom: 16 }}>
                <div style={{
                  padding: "11px 13px", borderRadius: 10, background: "#E7F2FE", marginBottom: 10,
                  display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 6
                }}>
                  <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 13, fontWeight: 700, color: C.primaryDeep }}>
                    {selected.bugSerial}
                  </span>
                  <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, color: C.charcoalSoft }}>
                    {selected.deployAsset || "\u2014"}{selected.fidoMode ? ` \u00b7 ${selected.fidoMode}` : ""}
                  </span>
                </div>

                {selected.pipeMaterial && (
                  <div style={{ marginBottom: 10 }}>
                    <PipeAcousticGuide material={selected.pipeMaterial} compact />
                  </div>
                )}

                {(() => {
                  const s = conditionScore(selected);
                  if (!s) return null;
                  return (
                    <div style={{
                      marginBottom: 10, padding: "10px 12px", borderRadius: 9,
                      background: s.soft, border: `1px solid ${s.colour}55`
                    }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                        <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: s.colour }}>
                          {s.label}
                        </span>
                        <span style={{ fontFamily: "'Space Grotesk',sans-serif", fontSize: 15, fontWeight: 700, color: s.colour }}>{s.pct}%</span>
                      </div>
                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9.5, color: C.charcoalSoft, marginTop: 4 }}>
                        {[
                          selected.deployAsset && `on ${selected.deployAsset.toLowerCase()}`,
                          PIPE_MATERIALS[selected.pipeMaterial]?.label,
                          selected.cond_diameter && `\u2300 ${selected.cond_diameter}`,
                          selected.cond_pressure && `${selected.cond_pressure.toLowerCase()} pressure`,
                          selected.cond_backfill && `${selected.cond_backfill.toLowerCase()} backfill`,
                          selected.cond_pipecondition && selected.cond_pipecondition.toLowerCase(),
                          selected.cond_background && selected.cond_background.toLowerCase(),
                          selected.cond_consumption && `${selected.cond_consumption.toLowerCase()} demand`,
                        ].filter(Boolean).join(" \u00b7 ")}
                      </div>
                      {selected.outcome === "NO LEAK" && s.pct < 45 && (
                        <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9.5, fontWeight: 700, color: s.colour, marginTop: 5 }}>
                          A no-leak result under these conditions is not conclusive \u2014 consider a repeat session.
                        </div>
                      )}
                    </div>
                  );
                })()}

                {selected.type === "fido2_deploy" && selected.sessionShot && (
                  <div style={{ display: "flex", gap: 10, marginBottom: 10 }}>
                    <img src={selected.sessionShot.photo} alt="session" onClick={() => setZoomPhoto(selected.sessionShot.photo)}
                      style={{ width: 70, height: 92, objectFit: "cover", borderRadius: 7, cursor: "zoom-in" }} />
                    <div style={{ flex: 1 }}>
                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9.5, fontWeight: 700, color: C.charcoalSoft }}>SESSION ID</div>
                      {editMode ? (
                        <input value={selected.sessionId || ""} onChange={(e) => updateCapture(selected.id, { sessionId: e.target.value })}
                          style={{ width: "100%", boxSizing: "border-box", padding: "5px 7px", borderRadius: 6, border: `1.5px solid ${C.primary}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 13, fontWeight: 700, color: C.charcoal }} />
                      ) : (
                        <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 14, fontWeight: 700, color: C.charcoal }}>{selected.sessionId || "\u2014"}</div>
                      )}
                      <div style={{ display: "flex", gap: 12, marginTop: 7 }}>
                        <div>
                          <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9, color: C.charcoalSoft }}>SIGNAL</div>
                          <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 11, color: C.charcoal }}>{selected.sessionShot.signalStrength || "\u2014"}</div>
                        </div>
                        <div>
                          <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9, color: C.charcoalSoft }}>BATTERY</div>
                          <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 11, color: C.charcoal }}>{selected.sessionShot.batteryVoltage || "\u2014"}</div>
                        </div>
                      </div>
                    </div>
                  </div>
                )}

                {selected.type === "fido2_retrieve" && (
                  <>
                    {selected.resultShot && (
                      <img src={selected.resultShot.photo} alt="results" onClick={() => setZoomPhoto(selected.resultShot.photo)}
                        style={{ width: "100%", maxHeight: 280, objectFit: "contain", borderRadius: 8, background: "#fff", border: `1px solid ${C.line}`, cursor: "zoom-in", marginBottom: 10 }} />
                    )}
                    {(() => {
                      const a = (selected.sessionId || "").trim().toLowerCase();
                      const b = (selected.resultSessionId || "").trim().toLowerCase();
                      const known = a && b;
                      const match = known && a === b;
                      return (
                        <div style={{
                          padding: "9px 11px", borderRadius: 9, marginBottom: 10,
                          background: !known ? C.paperDeep : match ? C.approveSoft : C.flagSoft,
                          display: "flex", alignItems: "center", gap: 7
                        }}>
                          {known && (match ? <CheckCircle2 size={14} color={C.approve} /> : <AlertTriangle size={14} color={C.flag} />)}
                          <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 11, fontWeight: 700, color: !known ? C.charcoalSoft : match ? C.approve : C.flag }}>
                            {!known ? "SESSION ID NOT CONFIRMED" : match ? `SESSION ${selected.sessionId} CONFIRMED` : `MISMATCH \u2014 ${selected.sessionId} vs ${selected.resultSessionId}`}
                          </span>
                        </div>
                      );
                    })()}
                    {(() => {
                      const st = FIDO_OUTCOMES[selected.outcome];
                      return (
                        <div style={{ marginBottom: 8 }}>
                          <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9.5, fontWeight: 700, color: C.charcoalSoft, marginBottom: 4 }}>FINDING</div>
                          {editMode ? (
                            <select value={selected.outcome || ""} onChange={(e) => updateCapture(selected.id, { outcome: e.target.value })}
                              style={{ width: "100%", padding: "7px 9px", borderRadius: 7, border: `1.5px solid ${C.primary}`, fontFamily: "'Inter',sans-serif", fontSize: 12, fontWeight: 700, color: C.charcoal }}>
                              <option value="">\u2014</option>
                              {Object.keys(FIDO_OUTCOMES).map((o) => <option key={o} value={o}>{o}</option>)}
                            </select>
                          ) : (
                            <span style={{
                              display: "inline-block", padding: "6px 14px", borderRadius: 999,
                              background: st ? st.soft : C.paperDeep,
                              fontFamily: "'Inter',sans-serif", fontSize: 12, fontWeight: 700,
                              color: st ? st.colour : C.charcoalSoft
                            }}>{selected.outcome || "NO FINDING RECORDED"}</span>
                          )}
                        </div>
                      );
                    })()}
                    <div>
                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9.5, fontWeight: 700, color: C.charcoalSoft, marginBottom: 3 }}>NOTE</div>
                      {editMode ? (
                        <input value={selected.outcomeNote || ""} onChange={(e) => updateCapture(selected.id, { outcomeNote: e.target.value })}
                          style={{ width: "100%", boxSizing: "border-box", padding: "6px 8px", borderRadius: 6, border: `1.5px solid ${C.primary}`, fontFamily: "'Inter',sans-serif", fontSize: 11.5, color: C.charcoal }} />
                      ) : (
                        <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 12, color: selected.outcomeNote ? C.charcoal : C.charcoalSoft }}>
                          {selected.outcomeNote || "\u2014"}
                        </div>
                      )}
                    </div>
                  </>
                )}
              </div>
            )}

            {selected.type === "exception" && (
              <div style={{
                marginBottom: 16, padding: 12, borderRadius: 10,
                background: C.reviewSoft, border: `1.5px solid ${C.review}`
              }}>
                <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 6 }}>
                  <AlertTriangle size={15} color={C.review} />
                  <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 12, fontWeight: 700, color: C.review }}>
                    COULD NOT BE READ
                  </span>
                </div>
                <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 13, fontWeight: 600, color: C.charcoal }}>
                  {selected.exceptionReason || "No reason recorded"}
                </div>
                <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 11, color: C.charcoalSoft, marginTop: 5 }}>
                  Flagged for a revisit. Arrange access or schedule a return visit.
                </div>
              </div>
            )}

            {(selected.conditionFlags || []).length > 0 && (
              <div style={{
                marginBottom: 16, padding: 12, borderRadius: 10,
                background: C.flagSoft, border: `1.5px solid ${C.flag}`
              }}>
                <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 8 }}>
                  <AlertTriangle size={15} color={C.flag} />
                  <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 12, fontWeight: 700, color: C.flag }}>
                    PROBLEM REPORTED
                  </span>
                </div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>
                  {selected.conditionFlags.map((f) => (
                    <span key={f} style={{
                      padding: "4px 10px", borderRadius: 999, background: "#fff",
                      border: `1px solid ${C.flag}`, fontFamily: "'Inter',sans-serif",
                      fontSize: 10, fontWeight: 700, color: C.flag
                    }}>{f}</span>
                  ))}
                </div>
                {editMode ? (
                  <input value={selected.conditionNote || ""} onChange={(e) => updateCapture(selected.id, { conditionNote: e.target.value })}
                    placeholder="Detail"
                    style={{ width: "100%", boxSizing: "border-box", marginTop: 8, padding: "7px 9px", borderRadius: 7, border: `1.5px solid ${C.primary}`, fontFamily: "'Inter',sans-serif", fontSize: 11.5, color: C.charcoal }} />
                ) : selected.conditionNote ? (
                  <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 12, color: C.charcoal, marginTop: 8 }}>
                    {selected.conditionNote}
                  </div>
                ) : null}
              </div>
            )}

            {selected.type === "asset" && (
              <div style={{ marginBottom: 16, padding: 11, borderRadius: 10, background: C.paper }}>
                {[
                  { key: "assetDescription", label: "DESCRIPTION", multiline: true },
                  { key: "assetSerial", label: "SERIAL", mono: true },
                  { key: "assetZone", label: "ZONE / STREET" },
                  { key: "assetErf", label: "ERF / UNIT", mono: true },
                  { key: "assetAccessNotes", label: "ACCESS NOTES" },
                ].map((f) => (
                  <div key={f.key} style={{ marginBottom: 9 }}>
                    <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9.5, fontWeight: 700, color: C.charcoalSoft, marginBottom: 3 }}>
                      {f.label}
                    </div>
                    {editMode ? (
                      f.multiline ? (
                        <textarea value={selected[f.key] || ""} onChange={(e) => updateCapture(selected.id, { [f.key]: e.target.value.toUpperCase() })} rows={2}
                          style={{ width: "100%", boxSizing: "border-box", padding: "6px 8px", borderRadius: 6, border: `1.5px solid ${C.primary}`, fontFamily: "'Inter',sans-serif", fontSize: 11.5, color: C.charcoal, resize: "vertical" }} />
                      ) : (
                        <input value={selected[f.key] || ""} onChange={(e) => updateCapture(selected.id, { [f.key]: e.target.value.toUpperCase() })}
                          style={{ width: "100%", boxSizing: "border-box", padding: "6px 8px", borderRadius: 6, border: `1.5px solid ${C.primary}`, fontFamily: f.mono ? "'IBM Plex Mono',monospace" : "'Inter',sans-serif", fontSize: 11.5, color: C.charcoal }} />
                      )
                    ) : (
                      <div style={{ fontFamily: f.mono ? "'IBM Plex Mono',monospace" : "'Inter',sans-serif", fontSize: 12, color: selected[f.key] ? C.charcoal : C.charcoalSoft }}>
                        {selected[f.key] || "—"}
                      </div>
                    )}
                  </div>
                ))}

                {FILTRATION_TYPES.has(selected.assetType) && (
                  <>
                    {[
                      { key: "assetLastCartridge", label: "LAST CARTRIDGE REPLACEMENT", months: CARTRIDGE_INTERVAL_MONTHS },
                      { key: "assetLastFlush", label: "LAST FLUSH / BACKWASH", months: FLUSH_INTERVAL_MONTHS },
                    ].map((f) => {
                      const st = dueStatus(selected[f.key], f.months);
                      return (
                        <div key={f.key} style={{ marginBottom: 9 }}>
                          <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9.5, fontWeight: 700, color: C.charcoalSoft, marginBottom: 3 }}>{f.label}</div>
                          {editMode ? (
                            <input type="date" value={selected[f.key] || ""} onChange={(e) => updateCapture(selected.id, { [f.key]: e.target.value })}
                              style={{ width: "100%", boxSizing: "border-box", padding: "6px 8px", borderRadius: 6, border: `1.5px solid ${C.primary}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 11.5, color: C.charcoal }} />
                          ) : (
                            <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 12, color: selected[f.key] ? C.charcoal : C.charcoalSoft }}>
                              {selected[f.key] || "\u2014"}
                            </div>
                          )}
                          {st && (
                            <div style={{
                              marginTop: 4, padding: "4px 8px", borderRadius: 6, display: "inline-block",
                              background: st.overdue ? C.flagSoft : C.approveSoft,
                              fontFamily: "'Inter',sans-serif", fontSize: 9.5, fontWeight: 700,
                              color: st.overdue ? C.flag : C.approve
                            }}>
                              {st.overdue ? `OVERDUE BY ${st.monthsOverdue} MONTH${st.monthsOverdue === 1 ? "" : "S"}` : `NEXT DUE ${st.dueDate}`}
                            </div>
                          )}
                        </div>
                      );
                    })}
                    <div style={{ marginBottom: 9 }}>
                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9.5, fontWeight: 700, color: C.charcoalSoft, marginBottom: 3 }}>CARTRIDGE CONDITION</div>
                      {editMode ? (
                        <select value={selected.assetCartridgeCondition || ""} onChange={(e) => updateCapture(selected.id, { assetCartridgeCondition: e.target.value })}
                          style={{ width: "100%", padding: "6px 8px", borderRadius: 6, border: `1.5px solid ${C.primary}`, fontFamily: "'Inter',sans-serif", fontSize: 11.5, color: C.charcoal }}>
                          <option value="">\u2014</option>
                          {CARTRIDGE_CONDITIONS.map((c2) => <option key={c2} value={c2}>{c2}</option>)}
                        </select>
                      ) : (
                        <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 12, fontWeight: 600, color: (selected.assetCartridgeCondition === "FOULED" || selected.assetCartridgeCondition === "NEEDS REPLACEMENT") ? C.flag : C.charcoal }}>
                          {selected.assetCartridgeCondition || "\u2014"}
                        </div>
                      )}
                    </div>
                    <div style={{ marginBottom: 9 }}>
                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9.5, fontWeight: 700, color: C.charcoalSoft, marginBottom: 3 }}>FILTER SIZE / SPEC</div>
                      {editMode ? (
                        <input value={selected.assetFilterSpec || ""} onChange={(e) => updateCapture(selected.id, { assetFilterSpec: e.target.value.toUpperCase() })}
                          style={{ width: "100%", boxSizing: "border-box", padding: "6px 8px", borderRadius: 6, border: `1.5px solid ${C.primary}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 11.5, color: C.charcoal }} />
                      ) : (
                        <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 12, color: selected.assetFilterSpec ? C.charcoal : C.charcoalSoft }}>
                          {selected.assetFilterSpec || "\u2014"}
                        </div>
                      )}
                    </div>
                  </>
                )}

                {!FILTRATION_TYPES.has(selected.assetType) && SERVICEABLE_TYPES.has(selected.assetType) && (() => {
                  const st = serviceStatus(selected.assetType, selected.assetLastService);
                  return (
                    <div>
                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9.5, fontWeight: 700, color: C.charcoalSoft, marginBottom: 3 }}>
                        DATE OF LAST SERVICE
                      </div>
                      {editMode ? (
                        <input type="date" value={selected.assetLastService || ""}
                          onChange={(e) => updateCapture(selected.id, { assetLastService: e.target.value })}
                          style={{ width: "100%", boxSizing: "border-box", padding: "6px 8px", borderRadius: 6, border: `1.5px solid ${C.primary}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 11.5, color: C.charcoal }} />
                      ) : (
                        <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 12, color: selected.assetLastService ? C.charcoal : C.charcoalSoft }}>
                          {selected.assetLastService || "—"}
                        </div>
                      )}
                      {st && (
                        <div style={{
                          marginTop: 5, padding: "5px 8px", borderRadius: 6, display: "inline-block",
                          background: st.overdue ? C.flagSoft : C.approveSoft,
                          fontFamily: "'Inter',sans-serif", fontSize: 10, fontWeight: 700,
                          color: st.overdue ? C.flag : C.approve
                        }}>
                          {st.overdue
                            ? `SERVICE OVERDUE BY ${st.monthsOverdue} MONTH${st.monthsOverdue === 1 ? "" : "S"}`
                            : `NEXT DUE ${st.dueDate}`}
                        </div>
                      )}
                    </div>
                  );
                })()}
              </div>
            )}

            {selected.type === "amr" && (selected.amrShots || []).length > 0 && (
              <div style={{ marginBottom: 16 }}>
                <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, fontWeight: 700, color: C.charcoal, marginBottom: 4 }}>
                  Meters linked to this {selected.amrAssetType?.toLowerCase() || "asset"} ({selected.amrShots.reduce((n, s) => n + (s.meters || []).length, 0)})
                </div>
                <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, color: C.charcoalSoft, marginBottom: 10 }}>
                  Tap a screenshot to view it full size and check the readings.
                </div>

                {selected.noMetersAllocated && (
                  <div style={{
                    marginBottom: 11, padding: "11px 13px", borderRadius: 9,
                    background: C.reviewSoft, border: `1px solid ${C.review}55`
                  }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 7 }}>
                      <AlertTriangle size={14} color={C.review} />
                      <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, fontWeight: 700, color: C.review }}>
                        NO METERS ALLOCATED TO THIS {selected.amrAssetType?.toUpperCase() || "ASSET"}
                      </span>
                    </div>
                    {editMode ? (
                      <input value={selected.noMetersNote || ""} onChange={(e) => updateCapture(selected.id, { noMetersNote: e.target.value })}
                        placeholder="Reason"
                        style={{ width: "100%", boxSizing: "border-box", marginTop: 7, padding: "7px 9px", borderRadius: 7, border: `1.5px solid ${C.primary}`, fontFamily: "'Inter',sans-serif", fontSize: 11.5, color: C.charcoal }} />
                    ) : selected.noMetersNote ? (
                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, color: C.charcoal, marginTop: 6 }}>
                        {selected.noMetersNote}
                      </div>
                    ) : null}
                  </div>
                )}

                {(selected.missingMeters || []).length > 0 && (
                  <div style={{
                    marginBottom: 11, padding: "10px 12px", borderRadius: 9,
                    background: C.flagSoft, border: `1px solid ${C.flag}55`
                  }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.flag }}>
                        {selected.missingMeters.length} EXPECTED METER{selected.missingMeters.length === 1 ? "" : "S"} NOT REPORTING
                      </span>
                      <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 11, fontWeight: 700, color: C.flag }}>
                        {selected.matchedCount}/{selected.expectedCount}
                      </span>
                    </div>
                    <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 10.5, color: C.charcoal, marginTop: 6, lineHeight: 1.6 }}>
                      {selected.missingMeters.join(", ")}
                    </div>
                    {(selected.extraMeters || []).length > 0 && (
                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, color: C.charcoalSoft, marginTop: 6 }}>
                        Not on the register: {selected.extraMeters.join(", ")}
                      </div>
                    )}
                  </div>
                )}

                {(() => {
                  const mine = (selected.amrShots || []).flatMap((s) => s.meters || []);
                  const bands = { GOOD: 0, FAIR: 0, WEAK: 0 };
                  mine.forEach((m) => { const b = signalBand(m.signal); if (b) bands[b.label]++; });
                  const dupHere = mine
                    .map((m) => (m.serial || "").trim().toUpperCase())
                    .filter((s) => s && duplicateMeters[s]);
                  const uniqueDupes = [...new Set(dupHere)];
                  const losing = uniqueDupes.filter((s) => {
                    const d = duplicateMeters[s];
                    return d.best && d.best.captureId !== selected.id;
                  });
                  return (
                    <div style={{ marginBottom: 12 }}>
                      <div style={{ display: "flex", gap: 5, marginBottom: uniqueDupes.length ? 8 : 0 }}>
                        {[["GOOD", C.approve, C.approveSoft], ["FAIR", C.review, C.reviewSoft], ["WEAK", C.flag, C.flagSoft]].map(([k, col, soft]) => (
                          <span key={k} style={{
                            padding: "4px 11px", borderRadius: 999, background: soft,
                            fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: col
                          }}>{bands[k]} {k}</span>
                        ))}
                      </div>
                      {uniqueDupes.length > 0 && (
                        <div style={{
                          padding: "9px 11px", borderRadius: 9, background: C.reviewSoft,
                          border: `1px solid ${C.review}`
                        }}>
                          <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.review, marginBottom: 5 }}>
                            {uniqueDupes.length} METER{uniqueDupes.length === 1 ? "" : "S"} ALSO SEEN BY ANOTHER ASSET
                          </div>
                          {uniqueDupes.map((s) => {
                            const d = duplicateMeters[s];
                            const isBest = d.best && d.best.captureId === selected.id;
                            return (
                              <div key={s} style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 10.5, color: C.charcoal, marginBottom: 2 }}>
                                {s} \u2014 {isBest
                                  ? "strongest here, this asset should own it"
                                  : `stronger on ${d.best ? `${d.best.assetType} ${d.best.position} (${d.best.signal})` : "another asset"}`}
                              </div>
                            );
                          })}
                          {losing.length > 0 && (
                            <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, color: C.charcoalSoft, marginTop: 5 }}>
                              Consider removing the weaker allocations so each meter reports to one device.
                            </div>
                          )}
                        </div>
                      )}
                    </div>
                  );
                })()}

                {selected.amrShots.map((shot, si) => (
                  <div key={si} style={{ marginBottom: 16, borderRadius: 10, border: `1px solid ${C.line}`, overflow: "hidden" }}>
                    <div style={{
                      padding: "7px 10px", background: C.paper, display: "flex",
                      justifyContent: "space-between", alignItems: "center"
                    }}>
                      <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: C.charcoal }}>
                        SCREENSHOT {si + 1} · {(shot.meters || []).length} METERS
                      </span>
                      {shot.readingDate && (
                        <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 10, color: C.charcoalSoft }}>
                          {shot.readingDate}
                        </span>
                      )}
                    </div>

                    <div
                      onClick={() => setZoomPhoto(shot.photo)}
                      style={{
                        position: "relative", cursor: "zoom-in", background: "#fff",
                        borderBottom: `1px solid ${C.line}`
                      }}>
                      <img src={shot.photo} alt={`Screenshot ${si + 1}`}
                        style={{ width: "100%", maxHeight: 420, objectFit: "contain", display: "block" }} />
                      <div style={{
                        position: "absolute", right: 8, bottom: 8, background: "rgba(43,47,51,0.75)",
                        color: "#fff", borderRadius: 6, padding: "3px 8px",
                        fontFamily: "'Inter',sans-serif", fontSize: 10, fontWeight: 600
                      }}>
                        Tap to enlarge
                      </div>
                    </div>

                    <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: "'IBM Plex Mono',monospace", fontSize: 11 }}>
                      <thead>
                        <tr style={{ background: C.paper }}>
                          <th style={{ textAlign: "left", padding: "6px 8px", fontFamily: "'Inter',sans-serif", fontSize: 9.5, color: C.charcoalSoft }}>SERIAL</th>
                          <th style={{ textAlign: "left", padding: "6px 8px", fontFamily: "'Inter',sans-serif", fontSize: 9.5, color: C.charcoalSoft }}>MODEL</th>
                          <th style={{ textAlign: "right", padding: "6px 8px", fontFamily: "'Inter',sans-serif", fontSize: 9.5, color: C.charcoalSoft }}>SIGNAL</th>
                          {editMode && <th style={{ width: 22 }} />}
                        </tr>
                      </thead>
                      <tbody>
                        {(shot.meters || []).map((m, mi) => {
                          const updateMeter = (patch) => {
                            const shots = selected.amrShots.map((s, i) =>
                              i !== si ? s : { ...s, meters: s.meters.map((mm, j) => (j === mi ? { ...mm, ...patch } : mm)) }
                            );
                            updateCapture(selected.id, { amrShots: shots });
                          };
                          const removeMeter = () => {
                            const shots = selected.amrShots.map((s, i) =>
                              i !== si ? s : { ...s, meters: s.meters.filter((_, j) => j !== mi) }
                            );
                            updateCapture(selected.id, { amrShots: shots });
                          };
                          return (
                            <tr key={mi} style={{ borderTop: `1px solid ${C.line}` }}>
                              <td style={{ padding: "4px 6px" }}>
                                {editMode ? (
                                  <input value={m.serial || ""} onChange={(e) => updateMeter({ serial: e.target.value })}
                                    style={{ width: "100%", boxSizing: "border-box", padding: "3px 5px", borderRadius: 4, border: `1.5px solid ${C.primary}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 10.5, color: C.charcoal }} />
                                ) : (() => {
                                  const key = (m.serial || "").trim().toUpperCase();
                                  const dup = duplicateMeters[key];
                                  const isBest = dup && dup.best && dup.best.captureId === selected.id;
                                  return (
                                    <span style={{ display: "flex", alignItems: "center", gap: 5 }}>
                                      <span style={{ color: C.charcoal, fontWeight: 600 }}>{m.serial}</span>
                                      {dup && (
                                        <span title={isBest ? "Strongest signal — this asset should own it" : "Also seen by another asset with a stronger signal"}
                                          style={{
                                            padding: "1px 6px", borderRadius: 999, fontFamily: "'Inter',sans-serif",
                                            fontSize: 8.5, fontWeight: 700,
                                            background: isBest ? C.approveSoft : C.reviewSoft,
                                            color: isBest ? C.approve : C.review
                                          }}>
                                          {isBest ? "BEST" : "DUP"}
                                        </span>
                                      )}
                                    </span>
                                  );
                                })()}
                              </td>
                              <td style={{ padding: "4px 6px" }}>
                                {editMode ? (
                                  <input value={m.model || ""} onChange={(e) => updateMeter({ model: e.target.value })}
                                    style={{ width: "100%", boxSizing: "border-box", padding: "3px 5px", borderRadius: 4, border: `1.5px solid ${C.primary}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 10.5, color: C.charcoal }} />
                                ) : (
                                  <span style={{ color: C.charcoalSoft }}>{m.model || "\u2014"}</span>
                                )}
                              </td>
                              <td style={{ padding: "4px 6px", textAlign: "right" }}>
                                {editMode ? (
                                  <input value={m.signal || ""} onChange={(e) => updateMeter({ signal: e.target.value })}
                                    style={{ width: "100%", boxSizing: "border-box", padding: "3px 5px", borderRadius: 4, border: `1.5px solid ${C.primary}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 10.5, color: C.charcoal, textAlign: "right" }} />
                                ) : (() => {
                                  const band = signalBand(m.signal);
                                  return (
                                    <span style={{
                                      display: "inline-block", padding: band ? "2px 7px" : 0, borderRadius: 999,
                                      background: band ? band.soft : "transparent",
                                      color: band ? band.colour : C.charcoalSoft, fontWeight: band ? 700 : 400
                                    }}>
                                      {m.signal || "\u2014"}
                                    </span>
                                  );
                                })()}
                              </td>
                              {editMode && (
                                <td style={{ padding: "4px 4px", textAlign: "center" }}>
                                  <button onClick={removeMeter} style={{ border: "none", background: "none", cursor: "pointer", padding: 0 }}>
                                    <X size={11} color={C.charcoalSoft} />
                                  </button>
                                </td>
                              )}
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>

                    {editMode && (
                      <button
                        onClick={() => {
                          const shots = selected.amrShots.map((s, i) =>
                            i !== si ? s : { ...s, meters: [...(s.meters || []), { serial: "", model: "", signal: "" }] }
                          );
                          updateCapture(selected.id, { amrShots: shots });
                        }}
                        style={{
                          width: "100%", border: "none", borderTop: `1px solid ${C.line}`, background: "#fff",
                          color: C.primary, cursor: "pointer", padding: "7px",
                          fontFamily: "'Inter',sans-serif", fontSize: 11, fontWeight: 600
                        }}>
                        + Add missing meter
                      </button>
                    )}
                  </div>
                ))}
              </div>
            )}

            {selected.type === "replacement" && (
              <div style={{ marginBottom: 16 }}>
                {[
                  { key: "oldMeter", label: "OLD METER (removed)", tone: C.review, toneSoft: C.reviewSoft },
                  { key: "newMeter", label: "NEW METER (installed)", tone: C.approve, toneSoft: C.approveSoft },
                ].map((side) => {
                  const data = selected[side.key];
                  if (!data) return (
                    <div key={side.key} style={{ marginBottom: 10, padding: 10, borderRadius: 9, background: C.paper, fontFamily: "'Inter',sans-serif", fontSize: 11.5, color: C.review }}>
                      {side.label} — not captured
                    </div>
                  );
                  return (
                    <div key={side.key} style={{ marginBottom: 10, padding: 10, borderRadius: 9, background: side.toneSoft, border: `1px solid ${side.tone}` }}>
                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, color: side.tone, marginBottom: 8, letterSpacing: 0.3 }}>
                        {side.label}
                      </div>
                      <div style={{ display: "flex", gap: 9 }}>
                        <img src={data.photo} alt={side.label} style={{ width: 80, height: 60, objectFit: "cover", borderRadius: 7, flexShrink: 0 }} />
                        <div style={{ flex: 1, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}>
                          <div>
                            <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9, fontWeight: 600, color: C.charcoalSoft }}>SERIAL</div>
                            {editMode ? (
                              <input value={data.serial || ""} onChange={(e) => updateCapture(selected.id, { [side.key]: { ...data, serial: e.target.value } })}
                                style={{ width: "100%", boxSizing: "border-box", padding: "4px 6px", borderRadius: 5, border: `1.5px solid ${C.primary}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 11.5, color: C.charcoal }} />
                            ) : (
                              <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 12, fontWeight: 600, color: C.charcoal }}>{data.serial || "—"}</div>
                            )}
                          </div>
                          <div>
                            <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 9, fontWeight: 600, color: C.charcoalSoft }}>READING</div>
                            {editMode ? (
                              <input value={data.reading || ""} onChange={(e) => updateCapture(selected.id, { [side.key]: { ...data, reading: e.target.value } })}
                                style={{ width: "100%", boxSizing: "border-box", padding: "4px 6px", borderRadius: 5, border: `1.5px solid ${C.primary}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 11.5, color: C.charcoal }} />
                            ) : (
                              <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 12, fontWeight: 600, color: C.charcoal }}>{data.reading ? `${data.reading} m³` : "—"}</div>
                            )}
                          </div>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}

            {selected.sensor?.photo && (
              <div style={{ marginBottom: 12, padding: 10, borderRadius: 9, background: C.paper }}>
                <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 8 }}>
                  <Radio size={13} color={C.charcoalSoft} />
                  <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, fontWeight: 600, color: C.charcoal }}>Sensor Deployment</span>
                </div>
                <div style={{ display: "flex", gap: 8 }}>
                  <img src={selected.sensor.photo} alt="sensor" style={{ width: 58, height: 44, objectFit: "cover", borderRadius: 6 }} />
                  {editMode ? (
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 5, flex: 1 }}>
                      {[
                        ["sessionId", "SESSION"],
                        ["deviceId", "DEVICE"],
                        ["signalStrength", "SIGNAL"],
                        ["batteryVoltage", "BATTERY"],
                      ].map(([k, label]) => (
                        <div key={k}>
                          <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 8.5, color: C.charcoalSoft, fontWeight: 600 }}>{label}</div>
                          <input
                            value={selected.sensor[k] || ""}
                            onChange={(e) => updateCapture(selected.id, { sensor: { ...selected.sensor, [k]: e.target.value } })}
                            placeholder="—"
                            style={{ width: "100%", boxSizing: "border-box", padding: "3px 6px", borderRadius: 5, border: `1.5px solid ${C.primary}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 11, color: C.charcoal }}
                          />
                        </div>
                      ))}
                    </div>
                  ) : (
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 4, flex: 1, fontFamily: "'IBM Plex Mono',monospace", fontSize: 11, color: C.charcoal }}>
                      <div>Session: <b>{selected.sensor.sessionId || "—"}</b></div>
                      <div>Device: <b>{selected.sensor.deviceId || "—"}</b></div>
                      <div>Signal: <b>{selected.sensor.signalStrength || "—"}</b></div>
                      <div>Battery: <b>{selected.sensor.batteryVoltage || "—"}</b></div>
                    </div>
                  )}
                </div>
              </div>
            )}

            {selected.consumption?.photo && (
              <div style={{ marginBottom: 16, padding: 10, borderRadius: 9, background: C.paper }}>
                <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 8 }}>
                  <Activity size={13} color={C.charcoalSoft} />
                  <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, fontWeight: 600, color: C.charcoal }}>Consumption Profile</span>
                </div>
                <div style={{ display: "flex", gap: 8 }}>
                  <img src={selected.consumption.photo} alt="consumption" style={{ width: 58, height: 44, objectFit: "cover", borderRadius: 6 }} />
                  <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 11, color: C.charcoal }}>
                    {selected.consumption.dateRange && <div>{selected.consumption.dateRange}</div>}
                    {selected.consumption.notes && <div style={{ color: C.charcoalSoft, fontFamily: "'Inter',sans-serif", marginTop: 3 }}>{selected.consumption.notes}</div>}
                  </div>
                </div>
              </div>
            )}

            {selected.status === "approved" ? (
              <div style={{
                display: "flex", alignItems: "center", justifyContent: "center", gap: 8, padding: "11px",
                background: C.approveSoft, borderRadius: 10
              }}>
                <Lock size={14} color={C.approve} />
                <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 13, fontWeight: 600, color: C.approve }}>
                  Approved & locked
                </span>
              </div>
            ) : (
              <div style={{ display: "flex", gap: 8 }}>
                <button
                  onClick={() => {
                    setEditMode(false);
                    setCaptures((cs) => cs.map((c) => (c.id === selected.id ? { ...c, status: "flagged" } : c)));
                    flashToast(`${selected.position} flagged`, C.flag);
                    setZoomPhoto(null);
                    setSelected(null);
                  }}
                  style={{
                    flex: 1, padding: "11px", borderRadius: 10, border: `1.5px solid ${C.flag}`, background: "#fff",
                    color: C.flag, fontFamily: "'Inter',sans-serif", fontWeight: 600, fontSize: 13, cursor: "pointer",
                    display: "flex", alignItems: "center", justifyContent: "center", gap: 6
                  }}>
                  <Flag size={14} /> Flag
                </button>
                <button onClick={() => setEditMode((m) => !m)} style={{
                  flex: 1, padding: "11px", borderRadius: 10,
                  border: `1.5px solid ${editMode ? C.primary : C.line}`,
                  background: editMode ? "#E7F2FE" : "#fff",
                  color: editMode ? C.primary : C.charcoal,
                  fontFamily: "'Inter',sans-serif", fontWeight: 600, fontSize: 13, cursor: "pointer",
                  display: "flex", alignItems: "center", justifyContent: "center", gap: 6
                }}>
                  {editMode ? <Check size={14} /> : <Edit3 size={14} />} {editMode ? "Done" : "Edit"}
                </button>
                <button
                  onClick={() => {
                    setEditMode(false);
                    setCaptures((cs) => cs.map((c) => (c.id === selected.id ? { ...c, status: "approved" } : c)));
                    flashToast(`${selected.position} approved`, C.approve);
                    setZoomPhoto(null);
                    setSelected(null);
                  }}
                  style={{
                    flex: 1, padding: "11px", borderRadius: 10, border: "none", background: C.approve,
                    color: "#fff", fontFamily: "'Inter',sans-serif", fontWeight: 600, fontSize: 13, cursor: "pointer",
                    display: "flex", alignItems: "center", justifyContent: "center", gap: 6
                  }}>
                  <Check size={14} /> Approve
                </button>
              </div>
            )}
          </div>
        </div>
      )}

      {toast && (
        <div style={{
          position: "fixed", bottom: 24, left: "50%", transform: "translateX(-50%)",
          background: C.charcoal, color: "#fff", borderRadius: 999, padding: "10px 20px",
          fontFamily: "'Inter',sans-serif", fontSize: 13, fontWeight: 600, zIndex: 200,
          display: "flex", alignItems: "center", gap: 8,
          boxShadow: "0 6px 20px -6px rgba(43,47,51,0.5)"
        }}>
          <Check size={14} color={toast.tone} /> {toast.text}
        </div>
      )}

      {zoomPhoto && (
        <div
          onClick={() => setZoomPhoto(null)}
          style={{
            position: "fixed", inset: 0, background: "rgba(20,22,24,0.94)", zIndex: 100,
            display: "flex", alignItems: "center", justifyContent: "center", padding: 16,
            cursor: "zoom-out"
          }}>
          <img src={zoomPhoto} alt="Screenshot full size"
            style={{ maxWidth: "100%", maxHeight: "100%", objectFit: "contain", borderRadius: 8 }} />
          <button
            onClick={(e) => { e.stopPropagation(); setZoomPhoto(null); }}
            style={{
              position: "absolute", top: 16, right: 16, width: 38, height: 38, borderRadius: 999,
              border: "none", background: "rgba(255,255,255,0.15)", cursor: "pointer",
              display: "flex", alignItems: "center", justifyContent: "center"
            }}>
            <X size={20} color="#fff" />
          </button>
        </div>
      )}
    </div>
  );
}

/* ---------- sync status indicator ---------- */
function SyncBadge({ status }) {
  const map = {
    idle: null,
    syncing: { icon: CloudUpload, color: C.charcoalSoft, label: "Saving…" },
    synced: { icon: Cloud, color: C.approve, label: "Saved" },
    offline: { icon: CloudOff, color: C.review, label: "Offline — will save when back online" },
  };
  const m = map[status];
  if (!m) return null;
  const Icon = m.icon;
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 5 }} title={m.label}>
      <Icon size={13} color={m.color} className={status === "syncing" ? "spin" : ""} />
      <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 11, color: m.color, fontWeight: 600 }}>
        {m.label}
      </span>
    </div>
  );
}

/* ---------- unlock screen (team access key) ---------- */
function UnlockScreen({ onUnlocked }) {
  const [key, setKey] = useState("");
  const [checking, setChecking] = useState(false);
  const [error, setError] = useState("");

  const tryUnlock = async () => {
    if (!key.trim() || checking) return;
    setChecking(true);
    setError("");
    try {
      const result = await verifyAppKey(key.trim());
      if (result) {
        setAppKey(key.trim());
        saveSession(key.trim(), result.role, result.username);
        onUnlocked(result.role, result.username);
      } else {
        setError("That key isn't right — check with your team and try again.");
      }
    } catch {
      setError("Couldn't reach the server — check your connection and try again.");
    } finally {
      setChecking(false);
    }
  };

  return (
    <div style={{
      minHeight: "70vh", display: "flex", flexDirection: "column",
      alignItems: "center", justifyContent: "center", padding: 20
    }}>
      <div style={{
        width: "100%", maxWidth: 340, background: "#fff", border: `1px solid ${C.line}`,
        borderRadius: 18, padding: "28px 24px", textAlign: "center"
      }}>
        <div style={{ display: "flex", justifyContent: "center", marginBottom: 14 }}>
          <div style={{
            width: 46, height: 46, borderRadius: 13,
            background: `linear-gradient(135deg, ${C.primary}, ${C.primaryDeep})`,
            display: "flex", alignItems: "center", justifyContent: "center"
          }}>
            <Lock size={20} color="#fff" />
          </div>
        </div>
        <h2 style={{ fontFamily: "'Space Grotesk',sans-serif", fontSize: 18, fontWeight: 700, color: C.charcoal, margin: "0 0 4px" }}>
          Team access key
        </h2>
        <p style={{ fontFamily: "'Inter',sans-serif", fontSize: 12.5, color: C.charcoalSoft, margin: "0 0 18px" }}>
          Enter your field or office key to open the app.
        </p>
        <input
          type="password"
          value={key}
          onChange={(e) => setKey(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && tryUnlock()}
          placeholder="Access key"
          autoFocus
          style={{
            width: "100%", boxSizing: "border-box", padding: "11px 13px", borderRadius: 9,
            border: `1.5px solid ${error ? C.flag : C.line}`, fontFamily: "'IBM Plex Mono',monospace",
            fontSize: 14, color: C.charcoal, outline: "none", textAlign: "center", marginBottom: 10
          }}
        />
        {error && (
          <p style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, color: C.flag, margin: "0 0 10px" }}>{error}</p>
        )}
        <button
          onClick={tryUnlock}
          disabled={checking || !key.trim()}
          style={{
            width: "100%", padding: "12px", borderRadius: 10, border: "none",
            cursor: checking || !key.trim() ? "not-allowed" : "pointer",
            background: checking || !key.trim() ? C.line : C.primary, color: "#fff",
            fontFamily: "'Inter',sans-serif", fontWeight: 600, fontSize: 14,
            display: "flex", alignItems: "center", justifyContent: "center", gap: 8
          }}>
          {checking ? <Loader2 size={16} className="spin" /> : "Unlock"}
        </button>
      </div>
    </div>
  );
}

/* ---------- task hub (choose what you're doing today) ---------- */
function TaskHub({ onChoose, role }) {
  return (
    <div style={{ maxWidth: 440, margin: "0 auto", padding: "36px 16px" }}>
      <h1 style={{ fontFamily: "'Space Grotesk',sans-serif", fontSize: 22, fontWeight: 700, color: C.charcoal, margin: "0 0 4px", textAlign: "center" }}>
        What are you doing today?
      </h1>
      <p style={{ fontFamily: "'Inter',sans-serif", fontSize: 13, color: C.charcoalSoft, margin: "0 0 24px", textAlign: "center" }}>
        {role === "office" ? "Choose a task to review its site reports." : "Choose a task to start or continue a site."}
      </p>
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        {Object.entries(TASKS).map(([id, t]) => {
          const Icon = t.icon;
          return (
            <button key={id} onClick={() => onChoose(id)} style={{
              display: "flex", alignItems: "center", gap: 14, textAlign: "left",
              padding: "18px 18px", borderRadius: 16, border: `1.5px solid ${C.line}`,
              background: "#fff", cursor: "pointer", transition: "border-color 0.15s"
            }}
              onMouseEnter={(e) => (e.currentTarget.style.borderColor = C.primary)}
              onMouseLeave={(e) => (e.currentTarget.style.borderColor = C.line)}
            >
              <div style={{
                width: 46, height: 46, borderRadius: 13, flexShrink: 0,
                background: `linear-gradient(135deg, ${C.primary}, ${C.primaryDeep})`,
                display: "flex", alignItems: "center", justifyContent: "center"
              }}>
                <Icon size={21} color="#fff" />
              </div>
              <div style={{ flex: 1 }}>
                <div style={{ fontFamily: "'Space Grotesk',sans-serif", fontWeight: 700, fontSize: 15.5, color: C.charcoal }}>
                  {t.label}
                </div>
                <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, color: C.charcoalSoft, marginTop: 2 }}>
                  {t.blurb}
                </div>
              </div>
              <ChevronRight size={18} color={C.charcoalSoft} />
            </button>
          );
        })}
      </div>
    </div>
  );
}

/* ---------- App ---------- */
export default function App() {
  // Restore a saved session so a download or reload doesn't force a fresh login
  const restored = loadSession();
  if (restored) setAppKey(restored.key);

  const [role, setRole] = useState(restored ? restored.role : null);
  const [username, setUsername] = useState(restored ? restored.username : "");
  const unlocked = !!role;
  const [task, setTask] = useState(null); // null (hub) | "meterwork" | "fido" | "amr" | "assets"
  const [screen, setScreen] = useState("setup");
  const [survey, setSurvey] = useState({ id: null, siteName: "", address: "", surveyName: "", tech: "", gps: null });
  const [captures, setCaptures] = useState([]);
  const [syncStatus, setSyncStatus] = useState("idle");

  // A restored session is trusted immediately so work isn't interrupted, but the key
  // is re-checked in the background in case it was changed or removed on the server.
  useEffect(() => {
    if (!restored) return;
    verifyAppKey(restored.key)
      .then((r) => {
        if (!r) {
          clearSession();
          setRole(null);
          setUsername("");
        }
      })
      .catch(() => { /* offline — keep working with the saved session */ });
    // eslint-disable-next-line
  }, []);

  const stateRef = useRef({ survey, captures });
  stateRef.current = { survey, captures };

  const syncNow = () => {
    const { survey: s, captures: c } = stateRef.current;
    if (!s.id) return;
    setSyncStatus("syncing");
    saveSurveyRecord(s.id, s, c)
      .then(() => setSyncStatus("synced"))
      .catch(() => setSyncStatus("offline"));
  };
  const syncNowRef = useRef(syncNow);
  syncNowRef.current = syncNow;

  // autosave whenever survey or captures change
  useEffect(() => {
    if (!survey.id) return;
    syncNowRef.current();
  }, [survey, captures]);

  // retry in the background while offline, and immediately when connectivity returns
  useEffect(() => {
    const onOnline = () => syncNowRef.current();
    window.addEventListener("online", onOnline);
    const interval = setInterval(() => {
      if (syncStatus === "offline") syncNowRef.current();
    }, 8000);
    return () => {
      window.removeEventListener("online", onOnline);
      clearInterval(interval);
    };
  }, [syncStatus]);

  const startNewSurvey = () => {
    setSurvey((s) => ({ ...s, id: genSurveyId(), taskType: task }));
    setCaptures([]);
    setScreen("capture");
  };

  const logOut = () => {
    clearSession();
    setRole(null);
    setUsername("");
    setTask(null);
    setSurvey({ id: null, siteName: "", address: "", surveyName: "", tech: "", gps: null, taskType: null, meterMode: null });
    setCaptures([]);
    setSyncStatus("idle");
    setScreen("setup");
  };

  const switchTask = () => {
    setTask(null);
    setSurvey({ id: null, siteName: "", address: "", surveyName: "", tech: "", gps: null, taskType: null, meterMode: null });
    setCaptures([]);
    setSyncStatus("idle");
    setScreen("setup");
  };

  const [resuming, setResuming] = useState(false);
  const resumeSurvey = async (summary) => {
    setResuming(true);
    try {
      const record = await fetchSurveyRecord(summary.key);
      const s = record.survey || {};
      // Older surveys used separate task types — map them onto the merged Meter Work task
      const legacy = s.taskType === "meter" || s.taskType === "replacement";
      setSurvey({
        id: summary.key,
        ...s,
        taskType: legacy ? "meterwork" : (s.taskType || task),
        meterMode: s.meterMode || (s.taskType === "replacement" ? "replace" : s.taskType === "meter" ? "read" : undefined),
      });
      setCaptures(record.captures || []);
      setScreen("capture");
    } catch (err) {
      alert("Couldn't load that survey — check your connection and try again.");
    } finally {
      setResuming(false);
    }
  };

  const deleteCurrentSurvey = async () => {
    try {
      if (survey.id) await deleteSurveyRecord(survey.id);
    } catch (err) {
      // even if the server delete fails, clear locally; it can be re-deleted from the list
    }
    setSurvey({ id: null, siteName: "", address: "", surveyName: "", tech: "", gps: null });
    setCaptures([]);
    setSyncStatus("idle");
    setScreen("setup");
  };

  return (
    <div style={{ minHeight: "100vh", background: C.paper, fontFamily: "'Inter',sans-serif" }}>
      <style>{FONTS}{`
        .spin { animation: spin 1s linear infinite; }
        @keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
        @keyframes shake { 0%,100% { transform: translateX(0); } 25% { transform: translateX(-6px); } 75% { transform: translateX(6px); } }
        * { box-sizing: border-box; }
      `}</style>

      <div style={{
        position: "sticky", top: 0, zIndex: 10, background: "rgba(244,247,249,0.9)", backdropFilter: "blur(6px)",
        borderBottom: `1px solid ${C.line}`, padding: "12px 20px", display: "flex",
        alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 10
      }}>
        <Logo />
        {unlocked && (
          <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
            <button
              onClick={logOut}
              title="Log out"
              style={{
                fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, letterSpacing: 0.5,
                color: role === "office" ? C.primaryDeep : C.approve,
                background: role === "office" ? "#E7F2FE" : C.approveSoft,
                padding: "3px 9px", borderRadius: 999, textTransform: "capitalize",
                border: "none", cursor: "pointer", display: "flex", alignItems: "center", gap: 5
              }}>
              {username} · {role === "office" ? "Office" : "Field"}
              <Lock size={10} />
            </button>
            {task && (
              <button onClick={switchTask} title="Switch task" style={{
                display: "flex", alignItems: "center", gap: 5, border: `1.5px solid ${C.line}`,
                background: "#fff", padding: "4px 10px", borderRadius: 999, cursor: "pointer",
                fontFamily: "'Inter',sans-serif", fontSize: 11, fontWeight: 600, color: C.charcoal
              }}>
                {TASKS[task]?.label}{task === "meterwork" && survey.meterMode ? ` · ${survey.meterMode === "replace" ? "Replace" : "Read"}` : ""} <X size={11} color={C.charcoalSoft} />
              </button>
            )}
            {task && <SyncBadge status={syncStatus} />}
            {task && <NavPills screen={screen} setScreen={setScreen} captures={captures} role={role} />}
          </div>
        )}
      </div>

      {!unlocked && <UnlockScreen onUnlocked={(r, u) => { setRole(r); setUsername(u); }} />}

      {unlocked && !task && (
        <TaskHub role={role} onChoose={(t) => { setTask(t); setScreen("setup"); }} />
      )}

      {unlocked && task && screen === "setup" && (
        <SetupScreen survey={survey} setSurvey={setSurvey} onStart={startNewSurvey} onResume={resumeSurvey} resuming={resuming} role={role} task={task} username={username} />
      )}
      {unlocked && task === "routetrace" && screen === "capture" && (
        <RouteTrace survey={survey} captures={captures} setCaptures={setCaptures} setScreen={setScreen} />
      )}
      {unlocked && task && task !== "routetrace" && screen === "capture" && (
        <CaptureScreen survey={survey} captures={captures} setCaptures={setCaptures} setScreen={setScreen} task={task} />
      )}
      {unlocked && task === "routetrace" && screen === "office" && (
        <RouteReview survey={survey} captures={captures} setCaptures={setCaptures} role={role} />
      )}
      {unlocked && task && task !== "routetrace" && screen === "office" && role !== "field" && (
        <OfficeScreen survey={survey} captures={captures} setCaptures={setCaptures} onDeleteSurvey={deleteCurrentSurvey} />
      )}
    </div>
  );
}
