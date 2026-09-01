import { useState } from "react";
import {
  Check, AlertTriangle, MapPin, Download, ChevronRight, Ruler, X,
} from "lucide-react";
import {
  validateSegment, confirmRoute, toGeoJSON, formatDepth,
  LIFECYCLE_STATE, prettyInstall,
} from "./lib/routeModel.js";

const C = {
  primary: "#0D86F3", charcoal: "#2B2F33", charcoalSoft: "#5B6570",
  paper: "#F4F7F9", paperDeep: "#E8EDF1", line: "#DCE3E8",
  approve: "#1B9C6E", approveSoft: "#E4F5EE",
  review: "#D98A22", reviewSoft: "#FBF0DE",
  flag: "#D6485A", flagSoft: "#FBE6E9",
};
const CONF_COLOR = { HIGH: C.approve, MEDIUM: C.review, LOW: C.flag };
const pretty = (s) => String(s).replace(/_/g, " ");

const card = {
  background: "#fff", border: `1px solid ${C.line}`, borderRadius: 12,
  padding: 14, marginBottom: 12,
};
const label = {
  fontFamily: "'Inter',sans-serif", fontSize: 10, fontWeight: 700,
  letterSpacing: 0.6, color: C.charcoalSoft, marginBottom: 6, display: "block",
};
const btn = (bg, fg = "#fff") => ({
  padding: "11px 15px", borderRadius: 10, border: "none", background: bg,
  color: fg, fontFamily: "'Inter',sans-serif", fontSize: 13, fontWeight: 600,
  cursor: "pointer", display: "inline-flex", alignItems: "center",
  justifyContent: "center", gap: 7,
});
const mono = { fontFamily: "'IBM Plex Mono',monospace", fontWeight: 600 };

export default function RouteReview({ survey, captures, setCaptures, role }) {
  const [open, setOpen] = useState(null);
  const [reviewer, setReviewer] = useState("");
  const [toast, setToast] = useState(null);

  const routes = (captures || []).filter((c) => c.kind === "ROUTE_TRACE");

  function flash(t) { setToast(t); setTimeout(() => setToast(null), 2400); }

  function download(rec) {
    const gj = toGeoJSON(rec.route, rec.nodes, rec.segments);
    const blob = new Blob([JSON.stringify(gj, null, 2)], { type: "application/geo+json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${rec.route.name.replace(/[^\w-]+/g, "_")}.geojson`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function verify(rec) {
    const who = reviewer.trim();
    if (!who) { flash("Enter your name before verifying."); return; }
    const out = confirmRoute(rec.route, rec.nodes, rec.segments, who);
    const updated = {
      ...rec,
      route: out.route,
      nodes: out.nodes,
      segments: out.segments,
      events: [...(rec.events || []), ...out.events],
      verifiedAt: new Date().toISOString(),
    };
    setCaptures((cs) => cs.map((c) => (c.id === rec.id ? updated : c)));
    setOpen(updated);
    flash(`Verified — ${out.events.length} registration events raised.`);
  }

  if (routes.length === 0) {
    return (
      <div style={{ padding: 20, maxWidth: 620, margin: "0 auto" }}>
        <div style={{ ...card, textAlign: "center", padding: 30 }}>
          <MapPin size={22} color={C.charcoalSoft} />
          <div style={{
            fontFamily: "'Space Grotesk',sans-serif", fontSize: 15, fontWeight: 700,
            color: C.charcoal, marginTop: 10, marginBottom: 5,
          }}>
            No routes captured yet
          </div>
          <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 12, color: C.charcoalSoft, lineHeight: 1.6 }}>
            Walk a route on the capture screen and save it. It will appear here for review, verification and export.
          </div>
        </div>
      </div>
    );
  }

  if (open) {
    return (
      <RouteDetail
        rec={open}
        reviewer={reviewer}
        setReviewer={setReviewer}
        onBack={() => setOpen(null)}
        onVerify={() => verify(open)}
        onDownload={() => download(open)}
        toast={toast}
        canVerify={role !== "field"}
      />
    );
  }

  return (
    <div style={{ padding: 14, maxWidth: 620, margin: "0 auto" }}>
      <div style={{ ...label, marginBottom: 10 }}>CAPTURED ROUTES · {routes.length}</div>
      {routes.map((r) => {
        const total = r.segments.reduce((t, s) => t + s.length_m, 0);
        const done = r.route.lifecycle_state === LIFECYCLE_STATE.CONFIRMED;
        const weak = r.segments.filter((s) => s.confidence === "LOW").length;
        return (
          <button
            key={r.id}
            onClick={() => setOpen(r)}
            style={{ ...card, width: "100%", textAlign: "left", cursor: "pointer", display: "block" }}
          >
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
              <span style={{
                ...mono, fontSize: 9.5, padding: "2px 7px", borderRadius: 5,
                background: done ? C.approveSoft : C.reviewSoft,
                color: done ? C.approve : C.review,
              }}>
                {done ? "VERIFIED" : "AWAITING REVIEW"}
              </span>
              <ChevronRight size={14} color={C.charcoalSoft} style={{ marginLeft: "auto" }} />
            </div>
            <div style={{
              fontFamily: "'Space Grotesk',sans-serif", fontSize: 14, fontWeight: 700,
              color: C.charcoal, marginBottom: 3,
            }}>
              {r.route.name}
            </div>
            <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, color: C.charcoalSoft }}>
              {total.toFixed(0)} m · {r.segments.length} segments · {r.nodes.length} nodes
              {weak > 0 && ` · ${weak} low confidence`}
            </div>
          </button>
        );
      })}
    </div>
  );
}

function RouteDetail({ rec, reviewer, setReviewer, onBack, onVerify, onDownload, toast, canVerify }) {
  const { route, nodes, segments } = rec;
  const total = segments.reduce((t, s) => t + s.length_m, 0);
  const billable = nodes.filter((n) => n.is_registry_asset).length + segments.length;
  const done = route.lifecycle_state === LIFECYCLE_STATE.CONFIRMED;

  return (
    <div style={{ padding: 14, maxWidth: 620, margin: "0 auto" }}>
      <button onClick={onBack} style={{
        ...btn(C.paperDeep, C.charcoal), marginBottom: 12, padding: "8px 12px", fontSize: 12,
      }}>
        <X size={13} /> Back to routes
      </button>

      {toast && (
        <div style={{
          ...card, background: C.approveSoft, borderColor: C.approve, padding: "10px 13px",
          fontFamily: "'Inter',sans-serif", fontSize: 12, color: C.charcoal,
        }}>
          {toast}
        </div>
      )}

      <div style={card}>
        <div style={{
          fontFamily: "'Space Grotesk',sans-serif", fontSize: 16, fontWeight: 700,
          color: C.charcoal, marginBottom: 2,
        }}>
          {route.name}
        </div>
        <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 11, color: C.charcoalSoft, marginBottom: 14 }}>
          {pretty(route.utility_class)} · {pretty(route.capture_type)} · captured by {route.captured_by || "field"}
        </div>

        <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
          <Stat label="LENGTH" value={`${total.toFixed(0)} m`} />
          <Stat label="SEGMENTS" value={segments.length} />
          <Stat label="NODES" value={nodes.length} />
          <Stat label="BILLABLE" value={billable} />
        </div>

        {!done && canVerify && (
          <>
            <label style={label}>VERIFIED BY</label>
            <input
              value={reviewer}
              onChange={(e) => setReviewer(e.target.value)}
              placeholder="your name"
              style={{
                width: "100%", padding: "10px 12px", borderRadius: 9,
                border: `1px solid ${C.line}`, fontFamily: "'Inter',sans-serif",
                fontSize: 14, marginBottom: 10, boxSizing: "border-box",
              }}
            />
            <div style={{
              fontFamily: "'Inter',sans-serif", fontSize: 11, color: C.charcoalSoft,
              lineHeight: 1.5, marginBottom: 12,
            }}>
              Verifying confirms every segment and registry asset on this route and raises {billable} registration
              events. Fittings are geometry and are not billed.
            </div>
          </>
        )}

        <div style={{ display: "flex", gap: 8 }}>
          <button style={{ ...btn(C.paperDeep, C.charcoal), flex: 1 }} onClick={onDownload}>
            <Download size={14} /> GeoJSON
          </button>
          {!done && canVerify && (
            <button style={{ ...btn(C.approve), flex: 1.3 }} onClick={onVerify}>
              <Check size={14} /> Verify route
            </button>
          )}
          {done && (
            <div style={{
              ...btn(C.approveSoft, C.approve), flex: 1.3, cursor: "default",
            }}>
              <Check size={14} /> Verified
            </div>
          )}
        </div>
      </div>

      <div style={card}>
        <div style={{ ...label, marginBottom: 10 }}>SEGMENTS</div>
        {segments.map((s, i) => {
          const v = validateSegment(s);
          return (
            <div key={s.segment_id} style={{
              padding: "10px 0", borderTop: i === 0 ? "none" : `1px solid ${C.line}`,
            }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 3 }}>
                <span style={{
                  ...mono, fontSize: 9.5, padding: "2px 6px", borderRadius: 5,
                  background: CONF_COLOR[s.confidence], color: "#fff",
                }}>{s.confidence}</span>
                <span style={{ ...mono, fontSize: 12.5, color: C.charcoal }}>{s.length_m.toFixed(1)} m</span>
                <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, color: C.charcoalSoft }}>
                  {pretty(s.material)}{s.diameter_mm ? ` · ${s.diameter_mm}mm` : ""}
                </span>
              </div>
              <div style={{ fontFamily: "'Inter',sans-serif", fontSize: 11.5, color: C.charcoalSoft }}>
                {formatDepth(s)} · {prettyInstall(s.installation)}
              </div>
              {v.warnings.map((w, k) => (
                <div key={k} style={{
                  display: "flex", gap: 5, alignItems: "flex-start", marginTop: 4,
                  fontFamily: "'Inter',sans-serif", fontSize: 10.5, color: C.review, lineHeight: 1.4,
                }}>
                  <AlertTriangle size={11} style={{ flexShrink: 0, marginTop: 1 }} /> {w}
                </div>
              ))}
            </div>
          );
        })}
      </div>

      <div style={card}>
        <div style={{ ...label, marginBottom: 10 }}>NODES</div>
        {nodes.map((n, i) => (
          <div key={n.node_id} style={{
            padding: "8px 0", borderTop: i === 0 ? "none" : `1px solid ${C.line}`,
            display: "flex", alignItems: "center", gap: 8,
          }}>
            <span style={{
              ...mono, fontSize: 9, padding: "2px 6px", borderRadius: 5,
              background: n.is_registry_asset ? C.approveSoft : C.paper,
              color: n.is_registry_asset ? C.approve : C.charcoalSoft,
            }}>
              {n.is_registry_asset ? "ASSET" : "FITTING"}
            </span>
            <span style={{ fontFamily: "'Inter',sans-serif", fontSize: 12.5, color: C.charcoal, fontWeight: 600 }}>
              {pretty(n.node_type)}
            </span>
            <span style={{ ...mono, fontSize: 10.5, color: C.charcoalSoft, marginLeft: "auto" }}>
              ±{n.accuracy_m != null ? n.accuracy_m.toFixed(2) : "?"} m
            </span>
          </div>
        ))}
      </div>

      <div style={{
        fontFamily: "'Inter',sans-serif", fontSize: 10.5, color: C.charcoalSoft,
        lineHeight: 1.5, padding: "0 4px 24px", display: "flex", gap: 7,
      }}>
        <Ruler size={12} style={{ flexShrink: 0, marginTop: 1 }} />
        Indicative record only. Positions and depths must be verified on site before any excavation.
      </div>
    </div>
  );
}

function Stat({ label: l, value }) {
  return (
    <div style={{ flex: 1, background: C.paper, borderRadius: 9, padding: "8px 6px", textAlign: "center" }}>
      <div style={{ ...mono, fontSize: 15, color: C.charcoal }}>{value}</div>
      <div style={{
        fontFamily: "'Inter',sans-serif", fontSize: 8.5, fontWeight: 700,
        letterSpacing: 0.5, color: C.charcoalSoft, marginTop: 1,
      }}>{l}</div>
    </div>
  );
}

