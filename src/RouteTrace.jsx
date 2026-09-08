import { useState, useEffect, useRef } from "react";
import {
  MapPin, Check, X, Plus, Flag, Loader2, AlertTriangle,
  Ruler, ArrowRight, Play, Square,
} from "lucide-react";
import {
  createRoute, createNode, createSegment, makeVertex, appendVertex,
  validateSegment, shouldDropVertex, formatDepth, accuracyState, haversine,
  UTILITY_CLASS, MATERIAL, NODE_TYPE, REGISTRY_NODE_TYPES,
  DEPTH_METHOD, SELECTABLE_DEPTH_METHODS, CAPTURE_TYPE,
  INSTALLATION, DEPTH_APPLIES, prettyInstall,
  CAPTURE_CONFIG, ACCURACY_GATE_M,
} from "./lib/routeModel.js";

/* ---------- design tokens (mirrors App.jsx) ---------- */
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

const ACC_COLOR = { GREEN: C.approve, AMBER: C.review, RED: C.flag, UNKNOWN: C.charcoalSoft };
const CONF_COLOR = { HIGH: C.approve, MEDIUM: C.review, LOW: C.flag };

/* Fitting types first (most common while walking), then registry assets. */
const NODE_GROUPS = [
  {
    label: "FITTINGS",
    hint: "geometry only — not billed",
    types: [
      NODE_TYPE.TEE, NODE_TYPE.ELBOW, NODE_TYPE.BEND, NODE_TYPE.SADDLE,
      NODE_TYPE.REDUCER, NODE_TYPE.COUPLING, NODE_TYPE.CABLE_JOINT,
      NODE_TYPE.TERMINATION, NODE_TYPE.CONNECTION, NODE_TYPE.END_CAP,
    ],
  },
  {
    label: "REGISTRY ASSETS",
    hint: "registered and billable",
    types: [
      NODE_TYPE.VALVE, NODE_TYPE.PRV, NODE_TYPE.HYDRANT, NODE_TYPE.MANHOLE,
      NODE_TYPE.CHAMBER, NODE_TYPE.WATER_METER, NODE_TYPE.ELECTRICAL_METER,
      NODE_TYPE.DB_KIOSK, NODE_TYPE.MINI_SUB, NODE_TYPE.DRAW_PIT,
      NODE_TYPE.DISTRIBUTION_BOX,
    ],
  },
];

const pretty = (s) => String(s).replace(/_/g, " ");

/* ---------- shared styles ---------- */
const card = {
  background: "#fff", border: `1px solid ${C.line}`, borderRadius: 12,
  padding: 14, marginBottom: 12,
};
const label = {
  fontFamily: "'Inter',sans-serif", fontSize: 10, fontWeight: 700,
  letterSpacing: 0.6, color: C.charcoalSoft, marginBottom: 6, display: "block",
};
const input = {
  width: "100%", padding: "10px 12px", borderRadius: 9,
  border: `1px solid ${C.line}`, fontFamily: "'Inter',sans-serif",
  fontSize: 14, color: C.charcoal, background: "#fff", boxSizing: "border-box",
};
const btn = (bg, fg = "#fff") => ({
  padding: "12px 16px", borderRadius: 10, border: "none", background: bg,
  color: fg, fontFamily: "'Inter',sans-serif", fontSize: 13, fontWeight: 600,
  cursor: "pointer", display: "inline-flex", alignItems: "center",
  justifyContent: "center", gap: 7,
});
const chip = (active, tint) => ({
  padding: "8px 11px", borderRadius: 8, cursor: "pointer",
  border: `1px solid ${active ? tint : C.line}`,
  background: active ? tint : "#fff",
  color: active ? "#fff" : C.charcoal,
  fontFamily: "'Inter',sans-serif", fontSize: 12, fontWeight: 600,
});
const mono = { fontFamily: "'IBM Plex Mono',monospace", fontWeight: 600 };

/* ============================================================ */

export default function RouteTrace({ survey, captures, setCaptures, setScreen }) {
  const [stage, setStage] = useState("setup"); // setup | walking | node | done
  const [route, setRoute] = useState(null);
  const [nodes, setNodes] = useState([]);
  const [segments, setSegments] = useState([]);

  // live GPS
  const [fix, setFix] = useState(null);        // {lat, lon, accuracy_m}
  const [gpsError, setGpsError] = useState(null);
  const [tracking, setTracking] = useState(false);

  // the segment currently being walked
  const [pending, setPending] = useState(null); // {startNode, vertices[]}
  const pendingRef = useRef(null);
  useEffect(() => { pendingRef.current = pending; }, [pending]);

  const watchRef = useRef(null);

  /* ---------- GPS watch. Never blocks anything. ---------- */
  useEffect(() => {
    if (!tracking) return;
    if (!navigator.geolocation) { setGpsError("No GPS on this device."); return; }

    watchRef.current = navigator.geolocation.watchPosition(
      (pos) => {
        const v = {
          lat: pos.coords.latitude,
          lon: pos.coords.longitude,
          accuracy_m: pos.coords.accuracy,
        };
        setFix(v);
        setGpsError(null);

        const p = pendingRef.current;
        if (p) {
          const last = p.vertices[p.vertices.length - 1];
          if (shouldDropVertex(last, v)) {
            setPending((cur) => cur
              ? { ...cur, vertices: [...cur.vertices, makeVertex(v.lat, v.lon, v.accuracy_m)] }
              : cur);
          }
        }
      },
      (err) => setGpsError(err.message || "GPS unavailable"),
      { enableHighAccuracy: true, maximumAge: 0, timeout: 15000 }
    );

    return () => {
      if (watchRef.current != null) navigator.geolocation.clearWatch(watchRef.current);
      watchRef.current = null;
    };
  }, [tracking]);

  /* ---------- actions ---------- */

  function beginRoute(meta) {
    setRoute(createRoute({
      estate_id: survey?.siteName || survey?.surveyName || "UNSPECIFIED",
      name: meta.name,
      utility_class: meta.utility_class,
      capture_type: meta.capture_type,
      captured_by: survey?.username || "field",
    }));
    setTracking(true);
    setStage("walking");
  }

  function dropManualVertex() {
    if (!fix || !pending) return;
    const last = pending.vertices[pending.vertices.length - 1];
    if (last && shouldDropVertex(last, fix, CAPTURE_CONFIG.min_vertex_interval_m) === false) return;
    setPending({ ...pending, vertices: [...pending.vertices, makeVertex(fix.lat, fix.lon, fix.accuracy_m)] });
  }

  function placeNode(payload) {
    const n = createNode({
      route_id: route.route_id,
      node_type: payload.node_type,
      lat: payload.lat,
      lon: payload.lon,
      accuracy_m: payload.accuracy_m,
      chainage_m: payload.chainage_m,
      serial_number: payload.serial_number || null,
      depth_m: payload.depth_m,
      depth_method: payload.depth_method,
      notes: payload.notes,
      captured_by: survey?.username || "field",
    });

    // Close the open segment ON the node, not at the last auto-dropped vertex.
    // Without this the line stops short and leaves a gap at every node.
    if (pending) {
      const nodeVertex = makeVertex(payload.lat, payload.lon, payload.accuracy_m);
      const walked = pending.vertices;
      const last = walked[walked.length - 1];
      const verts = last && haversine(last, nodeVertex) < 0.5
        ? walked                       // already standing on it
        : [...walked, nodeVertex];     // extend the line to the node
      const seg = createSegment({
        route_id: route.route_id,
        startNode: pending.startNode,
        endNode: n,
        vertices: verts,
        utility_class: route.utility_class,
        capture_type: route.capture_type,
        material: payload.material,
        diameter_mm: payload.diameter_mm,
        installation: payload.installation,
        observed_depth_m: route.capture_type === CAPTURE_TYPE.NEW_BUILD && DEPTH_APPLIES.has(payload.installation) ? payload.depth_m : null,
        observed_method: route.capture_type === CAPTURE_TYPE.NEW_BUILD && DEPTH_APPLIES.has(payload.installation) ? payload.depth_method : null,
        captured_by: survey?.username || "field",
      });
      setSegments((s) => [...s, seg]);
    }

    setNodes((ns) => [...ns, n]);
    setPending({ startNode: n, vertices: [makeVertex(payload.lat, payload.lon, payload.accuracy_m)] });
    setStage("walking");
  }

  function finishRoute() {
    setTracking(false);
    setPending(null);
    setStage("done");
  }

  function saveRoute() {
    const record = {
      id: route.route_id,
      kind: "ROUTE_TRACE",
      route,
      nodes,
      segments,
      capturedAt: new Date().toISOString(),
    };
    setCaptures([...(captures || []), record]);
    setScreen("capture");
  }

  const runningLength = pending
    ? segments.reduce((t, s) => t + s.length_m, 0)
    : segments.reduce((t, s) => t + s.length_m, 0);

  /* ---------- render ---------- */

  return (
    <div style={{ padding: 14, maxWidth: 620, margin: "0 auto" }}>
      <GpsBar fix={fix} error={gpsError} tracking={tracking} />

      {stage === "setup" && <SetupCard onBegin={beginRoute} />}

      {stage === "walking" && route && (
        <WalkingCard
          route={route}
          fix={fix}
          pending={pending}
          nodes={nodes}
          segments={segments}
          runningLength={runningLength}
          onDrop={dropManualVertex}
          onNode={() => setStage("node")}
          onFinish={finishRoute}
        />
      )}

      {stage === "node" && (
        <NodeCard
          route={route}
          fix={fix}
          firstNode={nodes.length === 0}
          onCancel={() => setStage("walking")}
          onPlace={placeNode}
        />
      )}

      {stage === "done" && route && (
        <DoneCard
          route={route}
          nodes={nodes}
          segments={segments}
          onSave={saveRoute}
          onBack={() => { setTracking(true); setStage("walking"); }}
        />
      )}
    </div>
  );
}

/* ============================================================ */

function GpsBar({ fix, error, tracking }) {
  const state = fix ? accuracyState(fix.accuracy_m) : "UNKNOWN";
  const tint = ACC_COLOR[state];
  return (
    <div style={{
      ...card, padding: "10px 13px", marginBottom: 12,
      display: "flex", alignItems: "center", gap: 10,
      borderLeft: `4px solid ${tint}`,
    }}>
      {tracking && !fix
        ? <Loader2 size={15} color={tint} />
        : <MapPin size={15} color={tint} />}
      <div style={{ flex: 1 }}>
        <div style={{ ...mono, fontSize: 12, color: C.charcoal }}>
          {fix ? `${fix.lat.toFixed(6)}, ${fix.lon.toFixed(6)}` : "Acquiring position…"}
        </div>
        <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 10.5, color: C.charcoalSoft }}>
          {fix ? `±${fix.accuracy_m.toFixed(2)} m · ${state}` : (error || "waiting for GPS")}
          {state === "RED" && " · saves allowed, confidence lowered"}
        </div>
      </div>
      {state === "GREEN" && (
        <span style={{ ...mono, fontSize: 9.5, color: C.approve }}>
          ≤{ACCURACY_GATE_M} m
        </span>
      )}
    </div>
  );
}

/* ---------- 1. route setup ---------- */

function SetupCard({ onBegin }) {
  const [name, setName] = useState("");
  const [uc, setUc] = useState(UTILITY_CLASS.WATER);
  const [ct, setCt] = useState(CAPTURE_TYPE.EXISTING);

  return (
    <div style={card}>
      <div style={{
        fontFamily: "'Space Grotesk',sans-serif", fontSize: 16, fontWeight: 700,
        color: C.charcoal, marginBottom: 12,
      }}>
        New route
      </div>

      <label style={label}>ROUTE NAME</label>
      <input
        style={{ ...input, marginBottom: 14 }}
        value={name}
        onChange={(e) => setName(e.target.value)}
        placeholder="110mm uPVC main — Gate 2 to Reservoir"
      />

      <label style={label}>UTILITY</label>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 14 }}>
        {Object.values(UTILITY_CLASS).map((v) => (
          <button key={v} style={chip(uc === v, C.primary)} onClick={() => setUc(v)}>
            {pretty(v)}
          </button>
        ))}
      </div>

      <label style={label}>CAPTURE TYPE</label>
      <div style={{ display: "flex", gap: 6, marginBottom: 8 }}>
        <button
          style={{ ...chip(ct === CAPTURE_TYPE.NEW_BUILD, C.approve), flex: 1, textAlign: "left" }}
          onClick={() => setCt(CAPTURE_TYPE.NEW_BUILD)}
        >
          New build
        </button>
        <button
          style={{ ...chip(ct === CAPTURE_TYPE.EXISTING, C.review), flex: 1, textAlign: "left" }}
          onClick={() => setCt(CAPTURE_TYPE.EXISTING)}
        >
          Existing
        </button>
      </div>
      <div style={{
        fontFamily: "'Inter',sans-serif", fontSize: 11, color: C.charcoalSoft,
        lineHeight: 1.5, marginBottom: 16,
      }}>
        {ct === CAPTURE_TYPE.NEW_BUILD
          ? "Trench open, pipe visible. Measure depth on each run — a change in depth splits the segment."
          : "Depth is only known where you can reach the pipe. Record it at each valve or chamber; the runs between are interpolated."}
      </div>

      <button
        style={{ ...btn(name.trim() ? C.primary : C.line, name.trim() ? "#fff" : C.charcoalSoft), width: "100%" }}
        disabled={!name.trim()}
        onClick={() => onBegin({ name: name.trim(), utility_class: uc, capture_type: ct })}
      >
        <Play size={14} /> Start walking
      </button>
    </div>
  );
}

/* ---------- 2. walking ---------- */

function WalkingCard({ route, fix, pending, nodes, segments, runningLength, onDrop, onNode, onFinish }) {
  const vcount = pending ? pending.vertices.length : 0;
  return (
    <>
      <div style={card}>
        <div style={{
          fontFamily: "'Space Grotesk',sans-serif", fontSize: 14, fontWeight: 700,
          color: C.charcoal, marginBottom: 2,
        }}>
          {route.name}
        </div>
        <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 11, color: C.charcoalSoft, marginBottom: 14 }}>
          {pretty(route.utility_class)} · {route.capture_type === CAPTURE_TYPE.NEW_BUILD ? "New build" : "Existing"}
        </div>

        <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
          <Stat label="NODES" value={nodes.length} />
          <Stat label="SEGMENTS" value={segments.length} />
          <Stat label="LENGTH" value={`${runningLength.toFixed(0)} m`} />
          <Stat label="POINTS" value={vcount} />
        </div>

        {!pending && (
          <div style={{
            background: C.reviewSoft, border: `1px solid ${C.review}`, borderRadius: 9,
            padding: 11, marginBottom: 12, display: "flex", gap: 8, alignItems: "flex-start",
          }}>
            <Flag size={14} color={C.review} style={{ flexShrink: 0, marginTop: 1 }} />
            <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, color: C.charcoal, lineHeight: 1.5 }}>
              Place the start node to begin. Stand at the first valve, chamber or fitting.
            </div>
          </div>
        )}

        {pending && (
          <div style={{
            fontFamily: "'Inter',sans-serif", fontSize: 11.5, color: C.charcoalSoft,
            lineHeight: 1.5, marginBottom: 12,
          }}>
            Walking from <strong style={{ color: C.charcoal }}>{pretty(pending.startNode.node_type)}</strong>.
            A point drops automatically every {CAPTURE_CONFIG.vertex_interval_m} m — add one by hand at corners.
          </div>
        )}

        <div style={{ display: "flex", gap: 8 }}>
          {pending && (
            <button style={{ ...btn(C.paperDeep, C.charcoal), flex: 1 }} onClick={onDrop} disabled={!fix}>
              <Plus size={14} /> Drop point
            </button>
          )}
          <button style={{ ...btn(C.primary), flex: 1.4 }} onClick={onNode} disabled={!fix}>
            <MapPin size={14} /> {pending ? "Place node" : "Start node"}
          </button>
        </div>
      </div>

      {segments.length > 0 && (
        <div style={card}>
          <div style={{ ...label, marginBottom: 10 }}>CAPTURED SO FAR</div>
          {segments.map((s, i) => (
            <SegmentRow key={s.segment_id} seg={s} index={i} />
          ))}
        </div>
      )}

      {segments.length > 0 && (
        <button style={{ ...btn(C.approve), width: "100%" }} onClick={onFinish}>
          <Square size={13} /> Finish route
        </button>
      )}
    </>
  );
}

function Stat({ label: l, value }) {
  return (
    <div style={{
      flex: 1, background: C.paper, borderRadius: 9, padding: "8px 6px", textAlign: "center",
    }}>
      <div style={{ ...mono, fontSize: 15, color: C.charcoal }}>{value}</div>
      <div style={{
        fontFamily: "'Inter',sans-serif", fontSize: 8.5, fontWeight: 700,
        letterSpacing: 0.5, color: C.charcoalSoft, marginTop: 1,
      }}>{l}</div>
    </div>
  );
}

function SegmentRow({ seg, index }) {
  const v = validateSegment(seg);
  return (
    <div style={{
      padding: "9px 0", borderTop: index === 0 ? "none" : `1px solid ${C.line}`,
    }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
        <span style={{
          ...mono, fontSize: 9.5, padding: "2px 6px", borderRadius: 5,
          background: CONF_COLOR[seg.confidence], color: "#fff",
        }}>{seg.confidence}</span>
        <span style={{ ...mono, fontSize: 12, color: C.charcoal }}>{seg.length_m.toFixed(1)} m</span>
        <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 11, color: C.charcoalSoft }}>
          {pretty(seg.material)}{seg.diameter_mm ? ` · ${seg.diameter_mm}mm` : ""}
        </span>
      </div>
      <div style={{
        fontFamily: "'Inter',sans-serif", fontSize: 11, color: C.charcoalSoft, marginTop: 3,
      }}>
        {formatDepth(seg)}
      </div>
      {v.warnings.map((w, i) => (
        <div key={i} style={{
          display: "flex", gap: 5, alignItems: "flex-start", marginTop: 4,
          fontFamily: "'Inter',sans-serif", fontSize: 10.5, color: C.review, lineHeight: 1.4,
        }}>
          <AlertTriangle size={11} style={{ flexShrink: 0, marginTop: 1 }} /> {w}
        </div>
      ))}
    </div>
  );
}

/* ---------- 3. node placement ---------- */

function NodeCard({ route, fix, firstNode, onCancel, onPlace }) {
  const [type, setType] = useState(null);
  const [depth, setDepth] = useState("");
  const [method, setMethod] = useState(DEPTH_METHOD.UNKNOWN);
  const [material, setMaterial] = useState(MATERIAL.UNKNOWN);
  const [dia, setDia] = useState("");
  const [serial, setSerial] = useState("");
  const [install, setInstall] = useState(INSTALLATION.BURIED);
  const [notes, setNotes] = useState("");

  const isRegistry = type ? REGISTRY_NODE_TYPES.has(type) : false;
  const newBuild = route.capture_type === CAPTURE_TYPE.NEW_BUILD;
  const buried = DEPTH_APPLIES.has(install);

  function submit() {
    onPlace({
      node_type: type,
      lat: fix.lat,
      lon: fix.lon,
      accuracy_m: fix.accuracy_m,
      chainage_m: null,
      serial_number: serial.trim() || null,
      installation: install,
      depth_m: !buried || depth === "" ? null : Number(depth),
      depth_method: !buried ? DEPTH_METHOD.NOT_APPLICABLE : (depth === "" ? DEPTH_METHOD.UNKNOWN : method),
      material,
      diameter_mm: dia === "" ? null : Number(dia),
      notes: notes.trim(),
    });
  }

  return (
    <div style={card}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
        <div style={{ fontFamily: "'Space Grotesk',sans-serif", fontSize: 15, fontWeight: 700, color: C.charcoal }}>
          {firstNode ? "Start node" : "Place node"}
        </div>
        <button onClick={onCancel} style={{ background: "none", border: "none", cursor: "pointer", padding: 4 }}>
          <X size={16} color={C.charcoalSoft} />
        </button>
      </div>

      {NODE_GROUPS.map((g) => (
        <div key={g.label} style={{ marginBottom: 12 }}>
          <div style={{ ...label, marginBottom: 6 }}>
            {g.label} <span style={{ fontWeight: 500, letterSpacing: 0 }}>· {g.hint}</span>
          </div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>
            {g.types.map((t) => (
              <button
                key={t}
                style={{
                  ...chip(type === t, g.label === "FITTINGS" ? C.charcoalSoft : C.primary),
                  fontSize: 11, padding: "6px 9px",
                }}
                onClick={() => setType(t)}
              >
                {pretty(t)}
              </button>
            ))}
          </div>
        </div>
      ))}

      {isRegistry && (
        <>
          <label style={label}>SERIAL / TAG (optional)</label>
          <input style={{ ...input, marginBottom: 12 }} value={serial}
            onChange={(e) => setSerial(e.target.value)} placeholder="cross-check against the tag" />
        </>
      )}

      <div style={{ height: 1, background: C.line, margin: "4px 0 14px" }} />

      <div style={{ ...label, marginBottom: 8 }}>HOW DOES THIS RUN SIT?</div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 5, marginBottom: 12 }}>
        {Object.values(INSTALLATION).map((v) => (
          <button
            key={v}
            style={{
              ...chip(install === v, DEPTH_APPLIES.has(v) ? C.charcoalSoft : C.approve),
              fontSize: 11, padding: "6px 9px",
            }}
            onClick={() => setInstall(v)}
          >
            {prettyInstall(v)}
          </button>
        ))}
      </div>

      {!buried && (
        <div style={{
          background: C.approveSoft, border: `1px solid ${C.approve}`, borderRadius: 8,
          padding: "9px 11px", marginBottom: 14,
          fontFamily: "'Inter',sans-serif", fontSize: 11, color: C.charcoal, lineHeight: 1.45,
        }}>
          Visible run — no depth needed. Position is directly observed, so this segment can reach
          high confidence even on an existing-property trace.
        </div>
      )}

      {buried && (
        <>
          <div style={{ ...label, marginBottom: 8 }}>
            {newBuild ? "DEPTH ON THIS RUN" : "DEPTH AT THIS POINT"}
          </div>
          <div style={{
            fontFamily: "'Inter',sans-serif", fontSize: 11, color: C.charcoalSoft,
            lineHeight: 1.5, marginBottom: 10,
          }}>
            {newBuild
              ? "Measure in the open trench. If the depth changes further along, place a node there and start a new segment."
              : "Only record what you can actually reach. Leave blank if the pipe isn't accessible here."}
          </div>

          <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
            <div style={{ flex: 1 }}>
              <input style={input} type="number" step="0.05" value={depth}
                onChange={(e) => setDepth(e.target.value)} placeholder="0.90" />
            </div>
            <div style={{ flex: 1.6 }}>
              <select style={input} value={method} onChange={(e) => setMethod(e.target.value)}>
                {SELECTABLE_DEPTH_METHODS.map((m) => (
                  <option key={m} value={m}>{pretty(m)}</option>
                ))}
              </select>
            </div>
          </div>

          {depth !== "" && method === DEPTH_METHOD.UNKNOWN && (
            <Warn text="A depth value needs a method. Pick how it was determined." />
          )}
          {depth === "" && !newBuild && (
            <Warn text="No depth here — the run will be interpolated from the nearest access points." tone="soft" />
          )}
        </>
      )}

      <div style={{ height: 1, background: C.line, margin: "4px 0 14px" }} />

      <div style={{ ...label, marginBottom: 8 }}>SEGMENT JUST WALKED</div>
      <label style={label}>MATERIAL</label>
      <select style={{ ...input, marginBottom: 12 }} value={material} onChange={(e) => setMaterial(e.target.value)}>
        {Object.entries(MATERIAL).map(([k, v]) => (
          <option key={k} value={v}>{v}</option>
        ))}
      </select>

      <label style={label}>DIAMETER (mm)</label>
      <input style={{ ...input, marginBottom: 12 }} type="number" value={dia}
        onChange={(e) => setDia(e.target.value)} placeholder="110" />

      <label style={label}>NOTES</label>
      <input style={{ ...input, marginBottom: 16 }} value={notes}
        onChange={(e) => setNotes(e.target.value)} placeholder="optional" />

      <button
        style={{ ...btn(type && fix ? C.primary : C.line, type && fix ? "#fff" : C.charcoalSoft), width: "100%" }}
        disabled={!type || !fix}
        onClick={submit}
      >
        <Check size={14} /> {firstNode ? "Set start node" : "Place node & close segment"}
      </button>
    </div>
  );
}

function Warn({ text, tone = "hard" }) {
  const t = tone === "hard" ? C.flag : C.review;
  const bg = tone === "hard" ? C.flagSoft : C.reviewSoft;
  return (
    <div style={{
      background: bg, border: `1px solid ${t}`, borderRadius: 8, padding: "8px 10px",
      marginBottom: 12, display: "flex", gap: 7, alignItems: "flex-start",
    }}>
      <AlertTriangle size={12} color={t} style={{ flexShrink: 0, marginTop: 1 }} />
      <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 11, color: C.charcoal, lineHeight: 1.45 }}>
        {text}
      </span>
    </div>
  );
}

/* ---------- 4. done ---------- */

function DoneCard({ route, nodes, segments, onSave, onBack }) {
  const total = segments.reduce((t, s) => t + s.length_m, 0);
  const billable = nodes.filter((n) => n.is_registry_asset).length + segments.length;
  const weak = segments.filter((s) => s.confidence === "LOW").length;

  return (
    <>
      <div style={card}>
        <div style={{
          fontFamily: "'Space Grotesk',sans-serif", fontSize: 16, fontWeight: 700,
          color: C.charcoal, marginBottom: 2,
        }}>
          {route.name}
        </div>
        <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 11, color: C.charcoalSoft, marginBottom: 14 }}>
          Ready to sync — office verification will confirm and register it.
        </div>

        <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
          <Stat label="LENGTH" value={`${total.toFixed(0)} m`} />
          <Stat label="NODES" value={nodes.length} />
          <Stat label="SEGMENTS" value={segments.length} />
          <Stat label="BILLABLE" value={billable} />
        </div>

        {weak > 0 && (
          <Warn tone="soft" text={`${weak} of ${segments.length} segments are low confidence. They will export marked as such.`} />
        )}

        <div style={{ display: "flex", gap: 8 }}>
          <button style={{ ...btn(C.paperDeep, C.charcoal), flex: 1 }} onClick={onBack}>
            <ArrowRight size={13} style={{ transform: "rotate(180deg)" }} /> Keep walking
          </button>
          <button style={{ ...btn(C.approve), flex: 1.4 }} onClick={onSave}>
            <Check size={14} /> Save route
          </button>
        </div>
      </div>

      <div style={card}>
        <div style={{ ...label, marginBottom: 10 }}>SEGMENTS</div>
        {segments.map((s, i) => <SegmentRow key={s.segment_id} seg={s} index={i} />)}
      </div>

      <div style={{
        fontFamily: "'Inter',sans-serif", fontSize: 10.5, color: C.charcoalSoft,
        lineHeight: 1.5, padding: "0 4px 20px", display: "flex", gap: 7,
      }}>
        <Ruler size={12} style={{ flexShrink: 0, marginTop: 1 }} />
        Indicative record only. Positions and depths must be verified on site before any excavation.
      </div>
    </>
  );
}
/**
 * THYNK-H2O — Route Trace data model (Phase A)
 * v1.0
 *
 * Route / Node / Segment. Shared by the field app (React) and the worker.
 * Pure functions, no I/O, no framework. Safe to import in either place.
 *
 * Rules enforced here:
 *   - Depth_Method is mandatory on every segment. UNKNOWN is allowed; blank is not.
 *   - Confidence is derived. It can be overridden DOWN, never UP.
 *   - Length is derived from vertices. Never hand-entered.
 *   - GPS accuracy never blocks a save. Poor accuracy lowers confidence, that is all.
 *   - Lifecycle_State is the billing trigger. Operational state is separate and never billed.
 */

// ---------------------------------------------------------------------------
// Pick-lists (seed data — admin-editable in production, not hard-coded)
// ---------------------------------------------------------------------------

export const UTILITY_CLASS = {
  WATER: 'WATER',
  SEWER: 'SEWER',
  STORMWATER: 'STORMWATER',
  ELECTRICAL: 'ELECTRICAL',
  FIBRE: 'FIBRE',
  GAS: 'GAS',
  IRRIGATION: 'IRRIGATION',
  UNKNOWN: 'UNKNOWN',
};

export const MATERIAL = {
  UPVC: 'uPVC',
  HDPE: 'HDPE',
  PE: 'PE',
  AC: 'Asbestos cement',
  STEEL: 'Steel',
  CAST_IRON: 'Cast iron',
  COPPER: 'Copper',
  CLAY: 'Clay',
  CONCRETE: 'Concrete',
  CONDUIT: 'Conduit / sleeve',
  DIRECT_BURIED: 'Direct buried cable',
  UNKNOWN: 'Unknown',
};

/** Node types that sit ON a route. Fitting types are new; asset types reuse the registry. */
export const NODE_TYPE = {
  // fittings — route-only, no separate asset record
  TEE: 'TEE',
  ELBOW: 'ELBOW',
  BEND: 'BEND',
  SADDLE: 'SADDLE',
  REDUCER: 'REDUCER',
  COUPLING: 'COUPLING',
  CABLE_JOINT: 'CABLE_JOINT',
  TERMINATION: 'TERMINATION',
  CHAMBER: 'CHAMBER',
  CONNECTION: 'CONNECTION', // joins another route
  END_CAP: 'END_CAP',
  // registry assets — these carry a full Asset_ID and fire REGISTRATION
  VALVE: 'VALVE',
  PRV: 'PRV',
  HYDRANT: 'HYDRANT',
  MANHOLE: 'MANHOLE',
  WATER_METER: 'WATER_METER',
  ELECTRICAL_METER: 'ELECTRICAL_METER',
  DB_KIOSK: 'DB_KIOSK',
  MINI_SUB: 'MINI_SUB',
  DRAW_PIT: 'DRAW_PIT',
  DISTRIBUTION_BOX: 'DISTRIBUTION_BOX',
};

/** Node types that create a registry asset record (and therefore a billable REGISTRATION). */
export const REGISTRY_NODE_TYPES = new Set([
  NODE_TYPE.VALVE,
  NODE_TYPE.PRV,
  NODE_TYPE.HYDRANT,
  NODE_TYPE.MANHOLE,
  NODE_TYPE.WATER_METER,
  NODE_TYPE.ELECTRICAL_METER,
  NODE_TYPE.DB_KIOSK,
  NODE_TYPE.MINI_SUB,
  NODE_TYPE.DRAW_PIT,
  NODE_TYPE.DISTRIBUTION_BOX,
]);

/**
 * Where the service physically runs. Not everything is buried — fire mains,
 * plant-room pipework and cable trays are often visible, and a visible service
 * is the strongest position record there is.
 */
export const INSTALLATION = {
  BURIED: 'BURIED',           // in the ground — depth applies
  IN_DUCT: 'IN_DUCT',         // buried sleeve or duct — depth applies
  SURFACE: 'SURFACE',         // laid on the ground / against a wall
  ELEVATED: 'ELEVATED',       // on brackets, piers, gantry, overhead
  IN_CHAMBER: 'IN_CHAMBER',   // inside a chamber or valve pit
  IN_BUILDING: 'IN_BUILDING', // plant room, riser, ceiling void
};

/** Only these need a depth. Everything else you can see. */
export const DEPTH_APPLIES = new Set([INSTALLATION.BURIED, INSTALLATION.IN_DUCT]);

export const DEPTH_METHOD = {
  MEASURED: 'MEASURED',         // open trench or chamber — physically measured
  LOCATED: 'LOCATED',           // pipe/cable locator with depth function
  ESTIMATED: 'ESTIMATED',       // judged from surface features or local knowledge
  INTERPOLATED: 'INTERPOLATED', // derived between two observed points — never hand-entered
  NOT_APPLICABLE: 'NOT_APPLICABLE', // above ground — nothing to measure
  UNKNOWN: 'UNKNOWN',           // buried but not determined — honest blank
};

/** Methods a technician may select. The other two are system-derived only. */
export const SELECTABLE_DEPTH_METHODS = [
  DEPTH_METHOD.MEASURED,
  DEPTH_METHOD.LOCATED,
  DEPTH_METHOD.ESTIMATED,
  DEPTH_METHOD.UNKNOWN,
];

export const CONFIDENCE = { HIGH: 'HIGH', MEDIUM: 'MEDIUM', LOW: 'LOW' };
const CONFIDENCE_RANK = { LOW: 0, MEDIUM: 1, HIGH: 2 };

export const LIFECYCLE_STATE = {
  PROVISIONAL: 'PROVISIONAL',
  CONFIRMED: 'CONFIRMED',
  REMOVED: 'REMOVED',
};

export const OPERATIONAL_STATE = {
  ACTIVE: 'ACTIVE',
  INACTIVE: 'INACTIVE',
  FAULTY: 'FAULTY',
};

export const CAPTURE_TYPE = {
  NEW_BUILD: 'NEW_BUILD',   // trench open, before backfill
  EXISTING: 'EXISTING',     // recovery trace
};

// ---------------------------------------------------------------------------
// Accuracy gate
// ---------------------------------------------------------------------------

/** Route capture is tighter than the 0.5 m asset-confirmation gate. Tune after RTK testing. */
export const ACCURACY_GATE_M = 0.30;
export const ACCURACY_DEGRADED_M = 2.00;

/** Never blocks a save. Returns a display state only. */
export function accuracyState(accuracy_m) {
  if (accuracy_m == null) return 'UNKNOWN';
  if (accuracy_m <= ACCURACY_GATE_M) return 'GREEN';
  if (accuracy_m <= ACCURACY_DEGRADED_M) return 'AMBER';
  return 'RED';
}

// ---------------------------------------------------------------------------
// Capture configuration
// ---------------------------------------------------------------------------

/**
 * Distance-based, not time-based — standing still must not spam vertices.
 * 10 m is the launch setting. Tune per site if needed; corners and features
 * are captured by manual drops and nodes, not by the auto interval.
 */
export const CAPTURE_CONFIG = {
  vertex_interval_m: 10,
  min_vertex_interval_m: 2,   // manual drops closer than this are ignored as noise
  photo_per_vertex: false,    // photos are on request only
};

/** True when the walker has moved far enough to warrant a new auto vertex. */
export function shouldDropVertex(lastVertex, current, interval_m = CAPTURE_CONFIG.vertex_interval_m) {
  if (!lastVertex) return true;
  return haversine(lastVertex, current) >= interval_m;
}

// ---------------------------------------------------------------------------
// Geometry
// ---------------------------------------------------------------------------

const R_EARTH_M = 6371000;

export function haversine(a, b) {
  const toRad = d => (d * Math.PI) / 180;
  const dLat = toRad(b.lat - a.lat);
  const dLon = toRad(b.lon - a.lon);
  const lat1 = toRad(a.lat);
  const lat2 = toRad(b.lat);
  const h =
    Math.sin(dLat / 2) ** 2 +
    Math.cos(lat1) * Math.cos(lat2) * Math.sin(dLon / 2) ** 2;
  return 2 * R_EARTH_M * Math.asin(Math.sqrt(h));
}

/** Derived — never hand-entered. */
export function segmentLength(vertices = []) {
  let total = 0;
  for (let i = 1; i < vertices.length; i++) total += haversine(vertices[i - 1], vertices[i]);
  return Math.round(total * 100) / 100;
}

/** Worst accuracy on the run governs the segment. One bad vertex taints the line. */
export function worstAccuracy(vertices = []) {
  const vals = vertices.map(v => v.accuracy_m).filter(v => v != null);
  return vals.length ? Math.max(...vals) : null;
}

// ---------------------------------------------------------------------------
// Confidence — derived, overridable DOWN only
// ---------------------------------------------------------------------------

export function deriveConfidence({ depth_method, vertices = [], capture_type, installation = INSTALLATION.BURIED }) {
  const acc = worstAccuracy(vertices);
  const accState = accuracyState(acc);
  const visible = !DEPTH_APPLIES.has(installation);

  let byDepth;
  if (visible) byDepth = CONFIDENCE.HIGH; // nothing to measure and nothing hidden
  else if (depth_method === DEPTH_METHOD.MEASURED) byDepth = CONFIDENCE.HIGH;
  else if (depth_method === DEPTH_METHOD.LOCATED) byDepth = CONFIDENCE.MEDIUM;
  else byDepth = CONFIDENCE.LOW; // ESTIMATED, INTERPOLATED or UNKNOWN

  let byAcc;
  if (accState === 'GREEN') byAcc = CONFIDENCE.HIGH;
  else if (accState === 'AMBER') byAcc = CONFIDENCE.MEDIUM;
  else byAcc = CONFIDENCE.LOW; // RED or UNKNOWN

  // Weakest link wins.
  let result = CONFIDENCE_RANK[byDepth] <= CONFIDENCE_RANK[byAcc] ? byDepth : byAcc;

  // A recovery trace is never HIGH — we did not see it in the ground.
  // Above-ground services are the exception: we did see it, so the cap lifts.
  if (capture_type === CAPTURE_TYPE.EXISTING && result === CONFIDENCE.HIGH && !visible) {
    result = CONFIDENCE.MEDIUM;
  }
  return result;
}

/** Applies a manual override, silently ignoring any attempt to raise confidence. */
export function applyConfidenceOverride(derived, override) {
  if (!override) return derived;
  return CONFIDENCE_RANK[override] < CONFIDENCE_RANK[derived] ? override : derived;
}

// ---------------------------------------------------------------------------
// Depth resolution
// ---------------------------------------------------------------------------
//
// Depth belongs to the point where it was observed, not to the run of pipe.
//
//   NEW_BUILD  — the trench is open, the pipe is visible. Depth is observed
//                continuously, so a change in depth is a real observation and
//                forces a segment break. One segment, one measured depth.
//
//   EXISTING   — depth is only knowable where you can reach the pipe: an
//                isolation valve, a chamber, a meter box. Between two access
//                points the depth is unknown and undetectable. Forcing a break
//                would mean inventing one. So the segment carries the depths
//                observed at each end and is marked INTERPOLATED.
//
// An interpolated depth is never presented as a single number.

/** Span beyond which interpolation between two observed points is not credible. */
export const INTERPOLATION_SPAN_WARN_M = 100;

/**
 * Works out a segment's depth from its bounding nodes.
 * Returns { depth_m, depth_min_m, depth_max_m, depth_method, depth_source }.
 */
export function resolveSegmentDepth({ capture_type, startNode, endNode, observed_depth_m = null, observed_method = null, length_m = 0, installation = INSTALLATION.BURIED }) {
  // Above ground — there is no depth, and that is a complete answer.
  if (!DEPTH_APPLIES.has(installation)) {
    return {
      depth_m: null,
      depth_min_m: null,
      depth_max_m: null,
      depth_method: DEPTH_METHOD.NOT_APPLICABLE,
      depth_source: 'ABOVE_GROUND',
    };
  }

  // New build: the technician measured this run directly.
  if (capture_type === CAPTURE_TYPE.NEW_BUILD && observed_depth_m != null) {
    return {
      depth_m: observed_depth_m,
      depth_min_m: observed_depth_m,
      depth_max_m: observed_depth_m,
      depth_method: observed_method || DEPTH_METHOD.MEASURED,
      depth_source: 'OBSERVED_ON_SEGMENT',
    };
  }

  const a = startNode && startNode.depth_m != null && startNode.depth_method !== DEPTH_METHOD.UNKNOWN ? startNode : null;
  const b = endNode && endNode.depth_m != null && endNode.depth_method !== DEPTH_METHOD.UNKNOWN ? endNode : null;

  // Both ends known — interpolate, and keep the range visible.
  if (a && b) {
    const lo = Math.min(a.depth_m, b.depth_m);
    const hi = Math.max(a.depth_m, b.depth_m);
    const same = lo === hi;
    return {
      depth_m: Math.round(((a.depth_m + b.depth_m) / 2) * 100) / 100,
      depth_min_m: lo,
      depth_max_m: hi,
      // Equal depths at both ends measured the same way is still only an
      // assumption about the middle — it stays INTERPOLATED.
      depth_method: DEPTH_METHOD.INTERPOLATED,
      depth_source: same ? 'INTERPOLATED_EQUAL_ENDS' : 'INTERPOLATED_BETWEEN_NODES',
      span_m: length_m,
    };
  }

  // One end known — carry it forward but do not pretend it holds along the run.
  const only = a || b;
  if (only) {
    return {
      depth_m: only.depth_m,
      depth_min_m: only.depth_m,
      depth_max_m: null,
      depth_method: DEPTH_METHOD.INTERPOLATED,
      depth_source: 'SINGLE_ACCESS_POINT',
      span_m: length_m,
    };
  }

  // Nothing known. Say so.
  return {
    depth_m: null,
    depth_min_m: null,
    depth_max_m: null,
    depth_method: DEPTH_METHOD.UNKNOWN,
    depth_source: 'NONE',
  };
}

const INSTALL_LABEL = {
  BURIED: 'buried',
  IN_DUCT: 'in duct',
  SURFACE: 'on surface',
  ELEVATED: 'elevated',
  IN_CHAMBER: 'in chamber',
  IN_BUILDING: 'in building',
};
export function prettyInstall(i) { return INSTALL_LABEL[i] || 'unspecified'; }

/** Human-readable depth for the map label and the PDF plan set. */
export function formatDepth(segment) {
  const { depth_m, depth_min_m, depth_max_m, depth_method, installation } = segment;
  if (depth_method === DEPTH_METHOD.NOT_APPLICABLE) {
    return `Above ground · ${prettyInstall(installation)}`;
  }
  if (depth_m == null) return 'Depth unknown';
  if (depth_method === DEPTH_METHOD.INTERPOLATED) {
    if (depth_max_m == null) return `~${depth_min_m} m at access point only`;
    if (depth_min_m === depth_max_m) return `~${depth_min_m} m (interpolated)`;
    return `${depth_min_m}–${depth_max_m} m (interpolated)`;
  }
  return `${depth_m} m (${depth_method.toLowerCase()})`;
}

// ---------------------------------------------------------------------------
// IDs
// ---------------------------------------------------------------------------

let _seq = 0;
function newId(prefix) {
  _seq = (_seq + 1) % 1000;
  const t = Date.now().toString(36).toUpperCase();
  const r = Math.floor(Math.random() * 1296).toString(36).toUpperCase().padStart(2, '0');
  return `${prefix}-${t}${r}${String(_seq).padStart(3, '0')}`;
}

// ---------------------------------------------------------------------------
// Factories
// ---------------------------------------------------------------------------

export function createRoute({
  estate_id,
  name,
  utility_class = UTILITY_CLASS.UNKNOWN,
  capture_type = CAPTURE_TYPE.EXISTING,
  ownership = null,
  captured_by = null,
}) {
  return {
    route_id: newId('RTE'),
    estate_id,
    name,
    utility_class,
    capture_type,
    ownership,
    lifecycle_state: LIFECYCLE_STATE.PROVISIONAL,
    segment_ids: [],
    node_ids: [],
    captured_by,
    captured_at: new Date().toISOString(),
    verified_by: null,
    verified_at: null,
  };
}

export function createNode({
  route_id,
  node_type,
  lat,
  lon,
  accuracy_m = null,
  chainage_m = null,
  serial_number = null,
  depth_m = null,
  depth_method = DEPTH_METHOD.UNKNOWN,
  photo_links = [],
  notes = '',
  captured_by = null,
}) {
  const isRegistryAsset = REGISTRY_NODE_TYPES.has(node_type);
  return {
    node_id: newId('NOD'),
    // Registry assets get a full Asset_ID; fittings do not.
    asset_id: isRegistryAsset ? newId('AST') : null,
    is_registry_asset: isRegistryAsset,
    route_id,
    node_type,
    lat,
    lon,
    accuracy_m,
    accuracy_state: accuracyState(accuracy_m),
    chainage_m,
    serial_number,
    // Depth observed AT this point — on a recovery trace this is the only
    // place depth is ever truly known. Segments inherit or interpolate from here.
    depth_m,
    depth_method,
    lifecycle_state: LIFECYCLE_STATE.PROVISIONAL,
    operational_state: OPERATIONAL_STATE.ACTIVE,
    photo_links,
    notes,
    captured_by,
    captured_at: new Date().toISOString(),
  };
}

export function createSegment({
  route_id,
  startNode,
  endNode,
  vertices = [],
  utility_class,
  capture_type = CAPTURE_TYPE.EXISTING,
  material = MATERIAL.UNKNOWN,
  diameter_mm = null,
  installation = INSTALLATION.BURIED,
  // New build only: depth measured along this open trench.
  observed_depth_m = null,
  observed_method = null,
  confidence_override = null,
  photo_links = [],
  notes = '',
  captured_by = null,
}) {
  const length_m = segmentLength(vertices);
  const depth = resolveSegmentDepth({
    capture_type, startNode, endNode, observed_depth_m, observed_method, length_m, installation,
  });
  const derived = deriveConfidence({ depth_method: depth.depth_method, vertices, capture_type, installation });
  return {
    segment_id: newId('SEG'),
    route_id,
    start_node_id: startNode ? startNode.node_id : null,
    end_node_id: endNode ? endNode.node_id : null,
    vertices,
    utility_class,
    capture_type,
    material,
    diameter_mm,
    installation,
    depth_m: depth.depth_m,
    depth_min_m: depth.depth_min_m,
    depth_max_m: depth.depth_max_m,
    depth_method: depth.depth_method,
    depth_source: depth.depth_source,
    confidence: applyConfidenceOverride(derived, confidence_override),
    confidence_derived: derived,
    confidence_override,
    length_m,
    worst_accuracy_m: worstAccuracy(vertices),
    lifecycle_state: LIFECYCLE_STATE.PROVISIONAL,
    photo_links,
    notes,
    captured_by,
    captured_at: new Date().toISOString(),
  };
}

export function makeVertex(lat, lon, accuracy_m = null) {
  return { lat, lon, accuracy_m, t: new Date().toISOString() };
}

/** Appending a vertex re-derives everything downstream. */
export function appendVertex(segment, vertex) {
  const vertices = [...segment.vertices, vertex];
  const derived = deriveConfidence({
    depth_method: segment.depth_method,
    vertices,
    capture_type: segment.capture_type,
    installation: segment.installation,
  });
  return {
    ...segment,
    vertices,
    length_m: segmentLength(vertices),
    worst_accuracy_m: worstAccuracy(vertices),
    confidence_derived: derived,
    confidence: applyConfidenceOverride(derived, segment.confidence_override),
  };
}

// ---------------------------------------------------------------------------
// Validation — warnings never block a save
// ---------------------------------------------------------------------------

export function validateSegment(segment) {
  const errors = [];
  const warnings = [];

  if (!segment.route_id) errors.push('Segment must belong to a route.');
  if (!segment.start_node_id || !segment.end_node_id) {
    errors.push('Segment must have a start and end node.');
  }
  const buriedRun = DEPTH_APPLIES.has(segment.installation);

  if (!Object.values(DEPTH_METHOD).includes(segment.depth_method)) {
    errors.push('Depth method is mandatory. Use UNKNOWN if not determined.');
  }
  if (segment.depth_m != null && segment.depth_method === DEPTH_METHOD.UNKNOWN) {
    errors.push('A depth value requires a depth method.');
  }
  if (segment.vertices.length < 2) {
    errors.push('Segment needs at least two vertices.');
  }

  // New build: the pipe was visible, so one segment carries one measured depth.
  // A range here means the depth changed along an open trench and the segment
  // should have been split at the point of change.
  if (buriedRun && segment.capture_type === CAPTURE_TYPE.NEW_BUILD) {
    if (segment.depth_min_m != null && segment.depth_max_m != null &&
        segment.depth_min_m !== segment.depth_max_m) {
      errors.push('Depth changes along this run. Split the segment at the point of change.');
    }
    if (segment.depth_method === DEPTH_METHOD.INTERPOLATED) {
      warnings.push('New build depth was not measured on this run — measure it while the trench is open.');
    }
  }

  // Existing trace: interpolation is expected, but say how far it is being stretched.
  if (buriedRun && segment.capture_type === CAPTURE_TYPE.EXISTING) {
    if (segment.depth_source === 'SINGLE_ACCESS_POINT') {
      warnings.push('Depth known at one end only — the rest of this run is assumed.');
    }
    if (segment.depth_method === DEPTH_METHOD.INTERPOLATED &&
        segment.length_m > INTERPOLATION_SPAN_WARN_M) {
      warnings.push(
        `Depth interpolated over ${Math.round(segment.length_m)} m with no access point between. ` +
        'Consider an intermediate reading.'
      );
    }
  }

  if (buriedRun && (segment.depth_m === 0 || segment.depth_min_m === 0)) {
    warnings.push('A depth of 0 m was recorded at an access point. Zero cover on a buried service is unlikely — check the entry.');
  }
  if (buriedRun && segment.depth_m == null) {
    warnings.push('No depth recorded — segment will show as low confidence.');
  }
  if (segment.material === MATERIAL.UNKNOWN) warnings.push('Material not identified.');
  if (segment.diameter_mm == null) warnings.push('Diameter not recorded.');
  if (accuracyState(segment.worst_accuracy_m) === 'RED') {
    warnings.push('GPS accuracy poor on this run — saved, but flagged low confidence.');
  }
  if (!segment.photo_links.length) warnings.push('No photo attached.');

  return { ok: errors.length === 0, errors, warnings };
}

// ---------------------------------------------------------------------------
// Lifecycle / billing
// ---------------------------------------------------------------------------

export const BILLABLE_EVENT = {
  REGISTRATION: 'REGISTRATION',
  ESTATE_RECORD: 'ESTATE_RECORD',
};

/**
 * Office verification is the billing trigger, not field capture.
 * Fires one REGISTRATION per confirmed segment and per confirmed registry-asset node.
 * Fittings (tee, elbow, saddle) do not bill — they are geometry, not assets.
 */
export function confirmRoute(route, nodes, segments, verified_by) {
  const now = new Date().toISOString();
  const events = [];

  /* Verification confirms the record. Billing is a separate question:
     fittings are geometry and raise no event, but they are still verified. */
  const confirm = (rec, kind, billable = true) => {
    if (rec.lifecycle_state === LIFECYCLE_STATE.CONFIRMED) return rec; // idempotent
    if (!billable) return { ...rec, lifecycle_state: LIFECYCLE_STATE.CONFIRMED };
    events.push({
      event_type: BILLABLE_EVENT.REGISTRATION,
      subject_kind: kind,
      subject_id: rec.segment_id || rec.node_id,
      asset_id: rec.asset_id || null,
      route_id: route.route_id,
      estate_id: route.estate_id,
      verified_by,
      occurred_at: now,
      idempotency_key: `REG:${rec.segment_id || rec.node_id}`,
    });
    return { ...rec, lifecycle_state: LIFECYCLE_STATE.CONFIRMED };
  };

  const nextNodes = nodes.map(n => confirm(n, 'NODE', n.is_registry_asset));
  const nextSegments = segments.map(s => confirm(s, 'SEGMENT'));
  const nextRoute = {
    ...route,
    lifecycle_state: LIFECYCLE_STATE.CONFIRMED,
    verified_by,
    verified_at: now,
  };

  return { route: nextRoute, nodes: nextNodes, segments: nextSegments, events };
}

// ---------------------------------------------------------------------------
// GeoJSON export — the client deliverable
// ---------------------------------------------------------------------------

const DISCLAIMER =
  'Indicative record only. Positions and depths must be verified on site before any excavation.';

export function toGeoJSON(route, nodes, segments) {
  const features = [];

  for (const s of segments) {
    // A LineString needs two points. A one-vertex segment is invalid GeoJSON
    // and will make strict readers reject the whole file, so it exports as a point.
    const coords = s.vertices.map(v => [v.lon, v.lat]);
    const geometry = coords.length >= 2
      ? { type: 'LineString', coordinates: coords }
      : { type: 'Point', coordinates: coords[0] || [0, 0] };
    features.push({
      type: 'Feature',
      geometry,
      properties: {
        kind: 'SEGMENT',
        segment_id: s.segment_id,
        route_id: s.route_id,
        route_name: route.name,
        utility_class: s.utility_class,
        material: s.material,
        diameter_mm: s.diameter_mm,
        installation: s.installation,
        depth_m: s.depth_m,
        depth_min_m: s.depth_min_m,
        depth_max_m: s.depth_max_m,
        depth_method: s.depth_method,
        depth_source: s.depth_source,
        depth_label: formatDepth(s),
        confidence: s.confidence,
        length_m: s.length_m,
        accuracy_m: s.worst_accuracy_m,
        capture_type: s.capture_type,
        lifecycle_state: s.lifecycle_state,
        captured_at: s.captured_at,
        disclaimer: DISCLAIMER,
      },
    });
  }

  for (const n of nodes) {
    features.push({
      type: 'Feature',
      geometry: { type: 'Point', coordinates: [n.lon, n.lat] },
      properties: {
        kind: 'NODE',
        node_id: n.node_id,
        asset_id: n.asset_id,
        node_type: n.node_type,
        route_id: n.route_id,
        chainage_m: n.chainage_m,
        serial_number: n.serial_number,
        depth_m: n.depth_m,
        depth_method: n.depth_method,
        accuracy_m: n.accuracy_m,
        lifecycle_state: n.lifecycle_state,
        operational_state: n.operational_state,
        captured_at: n.captured_at,
        disclaimer: DISCLAIMER,
      },
    });
  }

  return {
    type: 'FeatureCollection',
    name: route.name,
    metadata: {
      route_id: route.route_id,
      estate_id: route.estate_id,
      utility_class: route.utility_class,
      capture_type: route.capture_type,
      exported_at: new Date().toISOString(),
      disclaimer: DISCLAIMER,
    },
    features,
  };
}

// ---------------------------------------------------------------------------
// KV key structure (Cloudflare)
// ---------------------------------------------------------------------------

export const KV = {
  route: (estate, id) => `route:${estate}:${id}`,
  node: (estate, id) => `node:${estate}:${id}`,
  segment: (estate, id) => `segment:${estate}:${id}`,
  routeIndex: estate => `routeidx:${estate}`,
};
