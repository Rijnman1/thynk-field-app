import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
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

/* Team access key — entered once when the app opens, held for the session. */
let APP_KEY = "";
const setAppKey = (k) => { APP_KEY = k; };
const keyHeaders = (extra = {}) => ({ "X-App-Key": APP_KEY, ...extra });

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

/* ---------- field task definitions (the hub) ---------- */
const TASKS = {
  meter: {
    label: "Meter Survey",
    blurb: "Photograph existing meters and capture readings and serials.",
    icon: Camera,
  },
  fido: {
    label: "FIDO Leak Analysis",
    blurb: "Ultrasonic AI leak detection — sensor deployments and FIDO session feedback.",
    icon: Radio,
  },
  sat: {
    label: "SAT Survey",
    blurb: "Site survey — photos, locations, and notes.",
    icon: MapPin,
  },
  replacement: {
    label: "Meter Replacement",
    blurb: "Record old and new meter photos, serials, and readings at each position.",
    icon: RotateCcw,
  },
  amr: {
    label: "AMR Survey",
    blurb: "Map MUCs and repeaters, and record which meters each one can see.",
    icon: Radio,
  },
  assets: {
    label: "Asset Mapping",
    blurb: "Map estate infrastructure — water, electrical, AMR and general assets with coordinates.",
    icon: LayoutGrid,
  },
};

/* Asset taxonomy — matches the estate asset intake template exactly */
const ASSET_CATEGORIES = {
  WATER: [
    "BULK WATER METER",
    "ZONE METER",
    "SECTIONAL TITLE METER",
    "FREEHOLD WATER METER",
    "COMMON PROPERTY METER",
    "SECTIONAL TITLE INTERNAL METER",
    "ISOLATION VALVE",
    "PRESSURE CONTROL VALVE",
    "BOREHOLE",
    "WATER TANK",
  ],
  ELECTRICAL: [
    "FREEHOLD ELECTRICAL METER",
    "SECTIONAL TITLE ELECTRICAL METER",
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
  GENERAL: [
    "FIRE HYDRANT",
    "DRAIN COVER",
    "MANHOLE",
    "STORMWATER INLET",
    "FIBRE CHAMBER",
    "STREET LIGHT",
    "IRRIGATION POINT",
    "SIGNAGE",
  ],
};

const CATEGORY_COLOURS = {
  WATER: "#0D86F3",
  ELECTRICAL: "#D98A22",
  "AMR EQUIPMENT": "#7B5BD6",
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
  const task = survey.taskType || "meter";
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
      ...trailing(c),
    }));
    widths = [{ wch: 18 }, { wch: 20 }, { wch: 16 }, { wch: 12 }, { wch: 18 }, { wch: 13 }, { wch: 12 }, { wch: 9 }, { wch: 24 }, { wch: 14 }, { wch: 14 }, { wch: 15 }];
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
      date_captured: c.timestamp.date,
      captured_by: c.tech || "",
      verified_by: c.status === "approved" ? (reviewer || "") : "",
      status: STATUS_EXPORT_LABEL[c.status] || c.status,
    }));
    widths = [{ wch: 16 }, { wch: 15 }, { wch: 30 }, { wch: 46 }, { wch: 16 }, { wch: 13 }, { wch: 13 }, { wch: 20 }, { wch: 14 }, { wch: 34 }, { wch: 13 }, { wch: 14 }, { wch: 14 }, { wch: 15 }];
  } else if (task === "replacement") {
    sheetName = "Meter Replacements";
    fileSuffix = "meter_replacements";
    rows = captures.map((c) => ({
      ...common(c),
      "Old Meter Serial": c.oldMeter?.serial || "",
      "Old Meter Reading (m\u00b3)": c.oldMeter?.reading || "",
      "New Meter Serial": c.newMeter?.serial || "",
      "New Meter Reading (m\u00b3)": c.newMeter?.reading || "",
      ...trailing(c),
    }));
    widths = [{ wch: 18 }, { wch: 20 }, { wch: 16 }, { wch: 18 }, { wch: 20 }, { wch: 18 }, { wch: 20 }, { wch: 12 }, { wch: 9 }, { wch: 24 }, { wch: 14 }, { wch: 14 }, { wch: 15 }];
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
  } else if (task === "sat") {
    sheetName = "SAT Survey Points";
    fileSuffix = "sat_survey";
    rows = captures.map((c) => ({
      ...common(c),
      Notes: c.satNotes || "",
      ...trailing(c),
    }));
    widths = [{ wch: 18 }, { wch: 20 }, { wch: 16 }, { wch: 40 }, { wch: 12 }, { wch: 9 }, { wch: 24 }, { wch: 14 }, { wch: 14 }, { wch: 15 }];
  } else {
    sheetName = "Meter Readings";
    fileSuffix = "meter_survey";
    rows = captures.map((c) => ({
      ...common(c),
      "Serial Number": c.serial || "",
      "Meter Reading (m\u00b3)": c.reading || "",
      "AI Confidence %": c.confidence,
      ...trailing(c),
    }));
    widths = [{ wch: 18 }, { wch: 20 }, { wch: 16 }, { wch: 16 }, { wch: 18 }, { wch: 15 }, { wch: 12 }, { wch: 9 }, { wch: 24 }, { wch: 14 }, { wch: 14 }, { wch: 15 }];
  }

  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = widths;
  XLSX.utils.book_append_sheet(wb, ws, sheetName);

  // AMR surveys get a second sheet: one row per linked meter
  if (task === "amr") {
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
        { wch: 16 }, { wch: 16 }, { wch: 14 }, { wch: 14 }, { wch: 24 }, { wch: 14 },
      ];
      XLSX.utils.book_append_sheet(wb, amrWs, "Linked Meters");
    }
  }

  XLSX.writeFile(wb, `${safeFilename(survey.surveyName)}_${fileSuffix}.xlsx`);
}

/* CSV export in the estate asset intake template format */
function exportAssetCSV(survey, captures) {
  const headers = [
    "asset_id", "category", "asset_type", "description", "serial",
    "latitude", "longitude", "zone_or_street", "erf_or_unit", "access_notes",
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

/* Asset register PDF — one card per asset so a person can go and find it */
function printAssetRegister(survey, captures, reviewer) {
  const win = window.open("", "_blank", "width=900,height=1000");
  if (!win) return;

  const byCategory = {};
  captures.forEach((c) => {
    const cat = c.assetCategory || "UNCATEGORISED";
    (byCategory[cat] = byCategory[cat] || []).push(c);
  });

  const catColour = { WATER: "#0D86F3", ELECTRICAL: "#D98A22", "AMR EQUIPMENT": "#7B5BD6", GENERAL: "#5B6570", UNCATEGORISED: "#5B6570" };

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
  const TYPE_SHORT = { meter: "Meter", sensor: "Sensor", fido: "FIDO", sat: "SAT", replacement: "Replacement", amr: "AMR", consumption: "Profile" };
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

  const amrAssets = captures.filter((c) => c.type === "amr" && (c.amrShots || []).length);
  const amrAppendix = amrAssets.length ? `
      <h2 style="font-size:15px; margin:30px 0 6px;">Appendix — AMR Linked Meters</h2>
      ${amrAssets.map((c) => {
        const meters = (c.amrShots || []).flatMap((s) => s.meters || []);
        return `
        <h3 style="font-size:12.5px; margin:16px 0 4px;">${c.amrAssetType || "AMR"} · ${c.position}${c.amrSerial ? ` · ${c.amrSerial}` : ""} <span style="font-weight:normal;color:#5B6570;">(${meters.length} meters)</span></h3>
        <table>
          <thead><tr><th>Meter Serial</th><th>Model</th><th>Signal</th></tr></thead>
          <tbody>
            ${meters.map((m) => `<tr><td>${m.serial || "—"}</td><td>${m.model || "—"}</td><td>${m.signal || "—"}</td></tr>`).join("")}
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
          // only sites for the chosen task; older records without a task count as meter surveys
          setPrevious(recs.filter((r) => (r.taskType || "meter") === task));
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
    // FIDO feedback and AMR reader screenshots come from the gallery.
    if (type === "fido" || type === "amrShot") {
      pickTypeRef.current = type;
      galleryInputRef.current?.click();
      return;
    }
    // Meters, sensor installs, and SAT site photos: choose camera or gallery.
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
      } else if (type === "asset") {
        // Asset mapping — photo, GPS, and AI-read serial where visible
        const [photo, gps] = await Promise.all([loadDownscaledPhoto(file, 1100), getRealGPS()]);
        let serial = "";
        try {
          const extracted = await extractWithVision(photo, `This is a photograph of an estate infrastructure asset (a meter, valve, kiosk, hydrant, chamber, telemetry device or similar). Respond with ONLY a JSON object: {"serial": string|null}. Return the serial or device number if one is clearly legible on the asset, otherwise null. Do not guess.`);
          serial = extracted.serial || "";
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
      } else if (type === "sat") {
        // SAT survey — site photo with GPS and notes
        const [photo, gps] = await Promise.all([loadDownscaledPhoto(file, 1100), getRealGPS()]);
        setPending({
          id: Date.now(),
          type: "sat",
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
          fido: null,
          satNotes: "",
        });
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
    setCaptures((c) => [pending, ...c]);
    setPending(null);
    setPosition("");
    setCaptureError("");
    setStage("idle");
    setJustSaved(true);
    setTimeout(() => setJustSaved(false), 2200);
    setTimeout(() => inputRef.current?.focus(), 50);
  };

  const editField = (key, val) => setPending((p) => ({ ...p, [key]: val }));

  return (
    <div style={{ display: "flex", justifyContent: "center", padding: "20px 4px" }}>
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
                  {chooseFor === "meter" ? "Meter photo" : chooseFor === "sat" ? "Site photo" : "Installation photo"}
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

              {task === "assets" ? (
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
              ) : task === "replacement" ? (
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
                onClick={() => openPicker(task === "meter" ? "meter" : task === "sat" ? "sat" : "sensor")}
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
                      {task === "meter" ? "Capture Meter" : task === "sat" ? "Site Photo" : "Sensor Deployment"}
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

              <p style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, color: shake || captureError ? C.flag : justSaved ? C.approve : C.charcoalSoft, textAlign: "center", marginTop: 10, fontWeight: justSaved ? 600 : 400 }}>
                {captureError || (shake ? "Enter a position first." : justSaved ? "✓ Saved — ready for the next point" : "Pick what you're capturing — you can add the others afterwards. GPS and timestamp are recorded automatically.")}
              </p>
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
                {pending.type === "sat" && <><MapPin size={14} /> SAT Survey Point</>}
                {pending.type === "replacement" && <><RotateCcw size={14} /> Meter Replacement</>}
                {pending.type === "amr" && <><Radio size={14} /> {pending.amrAssetType || "AMR"} Survey</>}
                {pending.type === "asset" && <><LayoutGrid size={14} /> {pending.assetType || "Asset Mapping"}</>}
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
                      <div style={{
                        padding: "14px 10px", borderRadius: 10, border: `1.5px dashed ${C.line}`, background: C.paper,
                        textAlign: "center", fontFamily: "'Inter',sans-serif", fontSize: 11, color: C.charcoalSoft
                      }}>
                        Add up to 4 screenshots from the handheld reader
                      </div>
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
                                      <input value={m.signal || ""} onChange={(e) => updateMeter({ signal: e.target.value })} placeholder="—"
                                        style={{ width: "100%", boxSizing: "border-box", padding: "4px 5px", borderRadius: 5, border: `1px solid ${C.line}`, fontFamily: "'IBM Plex Mono',monospace", fontSize: 10, color: C.charcoalSoft, background: "#fff" }} />
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
                        <div style={{
                          padding: "8px 11px", borderRadius: 8, background: C.approveSoft,
                          fontFamily: "'Inter',sans-serif", fontSize: 11.5, fontWeight: 600, color: C.approve, textAlign: "center"
                        }}>
                          {pending.amrShots.reduce((n, s) => n + (s.meters || []).length, 0)} meters linked to this {pending.amrAssetType?.toLowerCase() || "asset"}
                        </div>
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

              {/* ---- SAT SURVEY: site photo + GPS + notes ---- */}
              {pending.type === "sat" && (
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
                  <div style={{ marginTop: 12 }}>
                    <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 600, color: C.charcoalSoft, marginBottom: 4 }}>NOTES</div>
                    <textarea
                      value={pending.satNotes || ""}
                      onChange={(e) => editField("satNotes", e.target.value)}
                      placeholder="What did you find at this point?"
                      rows={3}
                      style={{
                        width: "100%", boxSizing: "border-box", padding: "9px 11px", borderRadius: 9,
                        border: `1.5px solid ${C.line}`, fontFamily: "'Inter',sans-serif", fontSize: 13,
                        color: C.charcoal, outline: "none", resize: "vertical"
                      }}
                    />
                  </div>
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

  const updateStatus = (id, status) => {
    setCaptures((cs) => cs.map((c) => (c.id === id ? { ...c, status } : c)));
    setSelected((s) => (s ? { ...s, status } : s));
  };

  const [editMode, setEditMode] = useState(false);
  const [zoomPhoto, setZoomPhoto] = useState(null);
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
                  <Send size={14} /> {survey.taskType === "assets" ? "Asset Register PDF" : "Export PDF"}
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
              <PhotoWell size="small" photo={c.photo || c.newMeter?.photo || c.oldMeter?.photo || c.fido?.photo || c.consumption?.photo || c.sensor?.photo} />
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
                    {c.type === "sat" && <><MapPin size={10} color={C.primary} /> SAT</>}
                    {c.type === "replacement" && <><RotateCcw size={10} color={C.primary} /> Replaced</>}
                    {c.type === "amr" && <><Radio size={10} color={C.primary} /> {c.amrAssetType || "AMR"}</>}
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
              {selected.type === "sat" && (
                <>
                  <div style={{ background: C.paper, borderRadius: 9, padding: "8px 10px" }}>
                    <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, color: C.charcoalSoft, fontWeight: 600 }}>Type</div>
                    <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 13, color: C.charcoal, fontWeight: 600 }}>SAT Survey</div>
                  </div>
                  <div style={{ background: editMode ? "#fff" : C.paper, border: editMode ? `1.5px solid ${C.primary}` : "none", borderRadius: 9, padding: "8px 10px" }}>
                    <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10, color: C.charcoalSoft, fontWeight: 600 }}>Notes</div>
                    {editMode ? (
                      <input value={selected.satNotes || ""} onChange={(e) => updateCapture(selected.id, { satNotes: e.target.value })}
                        style={{ width: "100%", border: "none", outline: "none", fontFamily: "'Inter',sans-serif", fontSize: 12.5, color: C.charcoal, padding: 0, background: "transparent" }} />
                    ) : (
                      <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 12.5, color: C.charcoal }}>{selected.satNotes || "—"}</div>
                    )}
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
                                ) : (
                                  <span style={{ color: C.charcoal, fontWeight: 600 }}>{m.serial}</span>
                                )}
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
                                ) : (
                                  <span style={{ color: C.charcoalSoft }}>{m.signal || "\u2014"}</span>
                                )}
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
                <button onClick={() => updateStatus(selected.id, "flagged")} style={{
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
                <button onClick={() => { setEditMode(false); updateStatus(selected.id, "approved"); }} style={{
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
  const [role, setRole] = useState(null);
  const [username, setUsername] = useState("");
  const unlocked = !!role;
  const [task, setTask] = useState(null); // null (hub) | "meter" | "fido" | "sat"
  const [screen, setScreen] = useState("setup");
  const [survey, setSurvey] = useState({ id: null, siteName: "", address: "", surveyName: "", tech: "", gps: null });
  const [captures, setCaptures] = useState([]);
  const [syncStatus, setSyncStatus] = useState("idle");

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

  const switchTask = () => {
    setTask(null);
    setSurvey({ id: null, siteName: "", address: "", surveyName: "", tech: "", gps: null, taskType: null });
    setCaptures([]);
    setSyncStatus("idle");
    setScreen("setup");
  };

  const [resuming, setResuming] = useState(false);
  const resumeSurvey = async (summary) => {
    setResuming(true);
    try {
      const record = await fetchSurveyRecord(summary.key);
      setSurvey({ id: summary.key, ...record.survey });
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
            <span style={{
              fontFamily: "'Inter',sans-serif", fontSize: 10.5, fontWeight: 700, letterSpacing: 0.5,
              color: role === "office" ? C.primaryDeep : C.approve,
              background: role === "office" ? "#E7F2FE" : C.approveSoft,
              padding: "3px 9px", borderRadius: 999, textTransform: "capitalize"
            }}>
              {username} · {role === "office" ? "Office" : "Field"}
            </span>
            {task && (
              <button onClick={switchTask} title="Switch task" style={{
                display: "flex", alignItems: "center", gap: 5, border: `1.5px solid ${C.line}`,
                background: "#fff", padding: "4px 10px", borderRadius: 999, cursor: "pointer",
                fontFamily: "'Inter',sans-serif", fontSize: 11, fontWeight: 600, color: C.charcoal
              }}>
                {TASKS[task]?.label} <X size={11} color={C.charcoalSoft} />
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
      {unlocked && task && screen === "capture" && (
        <CaptureScreen survey={survey} captures={captures} setCaptures={setCaptures} setScreen={setScreen} task={task} />
      )}
      {unlocked && task && screen === "office" && role !== "field" && (
        <OfficeScreen survey={survey} captures={captures} setCaptures={setCaptures} onDeleteSurvey={deleteCurrentSurvey} />
      )}
    </div>
  );
}
