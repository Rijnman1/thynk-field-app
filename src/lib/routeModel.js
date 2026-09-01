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

  const confirm = (rec, kind) => {
    if (rec.lifecycle_state === LIFECYCLE_STATE.CONFIRMED) return rec; // idempotent
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

  const nextNodes = nodes.map(n => (n.is_registry_asset ? confirm(n, 'NODE') : n));
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
    features.push({
      type: 'Feature',
      geometry: {
        type: 'LineString',
        coordinates: s.vertices.map(v => [v.lon, v.lat]),
      },
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
