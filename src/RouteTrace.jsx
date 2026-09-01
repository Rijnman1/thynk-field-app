import { useState, useEffect, useRef } from "react";
import {
  MapPin, Check, X, Plus, Flag, Loader2, AlertTriangle,
  Ruler, ArrowRight, Play, Square,
} from "lucide-react";
import {
  createRoute, createNode, createSegment, makeVertex, appendVertex,
  validateSegment, shouldDropVertex, formatDepth, accuracyState,
  UTILITY_CLASS, MATERIAL, NODE_TYPE, REGISTRY_NODE_TYPES,
  DEPTH_METHOD, SELECTABLE_DEPTH_METHODS, CAPTURE_TYPE,
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
      estate_id: survey?.estate || survey?.site || "UNSPECIFIED",
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

    // Close the open segment against this node.
    if (pending) {
      const verts = pending.vertices.length
        ? pending.vertices
        : [makeVertex(payload.lat, payload.lon, payload.accuracy_m)];
      const seg = createSegment({
        route_id: route.route_id,
        startNode: pending.startNode,
        endNode: n,
        vertices: verts,
        utility_class: route.utility_class,
        capture_type: route.capture_type,
        material: payload.material,
        diameter_mm: payload.diameter_mm,
        observed_depth_m: route.capture_type === CAPTURE_TYPE.NEW_BUILD ? payload.depth_m : null,
        observed_method: route.capture_type === CAPTURE_TYPE.NEW_BUILD ? payload.depth_method : null,
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
  const [notes, setNotes] = useState("");

  const isRegistry = type ? REGISTRY_NODE_TYPES.has(type) : false;
  const newBuild = route.capture_type === CAPTURE_TYPE.NEW_BUILD;

  function submit() {
    onPlace({
      node_type: type,
      lat: fix.lat,
      lon: fix.lon,
      accuracy_m: fix.accuracy_m,
      chainage_m: null,
      serial_number: serial.trim() || null,
      depth_m: depth === "" ? null : Number(depth),
      depth_method: depth === "" ? DEPTH_METHOD.UNKNOWN : method,
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
