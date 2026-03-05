import { initializeBlock } from "@airtable/blocks/interface/ui";
import React, { useMemo, useState, useRef, useEffect } from "react";
import pptxgen from "pptxgenjs";
import type {
  Initiative,
  InitiativeStatus,
  Priority,
  Project,
  Capability,
  SegmentBreakdown,
  StatKey,
  ViewMode,
  StatusConfig,
} from "./types";
import { getCapabilities } from "./types";
import { usePortfolioData } from "./usePortfolioData";
import "./style.css";

// --- Brand Tokens (single source of truth for all inline styles) ---
const C = {
  bgPage: "#f2f4f7", bgPanel: "#ffffff", bgHover: "#f0f4ff", bgActive: "#e8f0fe",
  border: "#e5e4e4", borderLight: "#f0f1f2",
  textPrimary: "#333333", textSecondary: "#6b7280", textTertiary: "#9ca3af",
  blue: "#2d7ff9", blueHover: "#1a6ce8", blueSoft: "#cfdfff",
  green: "#20c933", greenSoft: "#d1f7c4", greenDark: "#338a17",
  amber: "#fcb400", amberSoft: "#ffeab6", amberDark: "#b87503",
  red: "#f82b60", redSoft: "#ffdce5", redDark: "#ba1e45",
  purple: "#8b46ff", purpleSoft: "#ede2fe", purpleDark: "#6b1cb0",
  walmartTrueBlue: "#0053E2", walmartBentonvilleBlue: "#001E60",
  walmartSparkYellow: "#FFC220", // LOGO ONLY — never UI chrome
  shadowCard: "none", shadowCardHover: "0 2px 8px rgba(0,0,0,0.07)",
  shadowDropdown: "0 4px 16px rgba(0,0,0,0.10), 0 1px 4px rgba(0,0,0,0.06)",
  shadowModal: "0 8px 32px rgba(0,0,0,0.14)", shadowSheet: "-4px 0 24px rgba(0,0,0,0.12)",
  radiusCheckbox: 3, radiusChip: 4, radiusButton: 6, radiusCard: 8, radiusHero: 12, radiusPill: 100,
  font: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
} as const;

// --- SVG Icon Primitives (fill="none", stroke, strokeWidth=1.75, round caps) ---
const SVG_DEFAULTS = { fill: "none", strokeWidth: "1.75", strokeLinecap: "round" as const, strokeLinejoin: "round" as const };

function IconX({ size = 16, color = "currentColor" }: { size?: number; color?: string }) {
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" {...SVG_DEFAULTS} stroke={color} aria-hidden>
      <line x1="4" y1="4" x2="12" y2="12" /><line x1="12" y1="4" x2="4" y2="12" />
    </svg>
  );
}

function IconArrowLeft({ size = 14, color = "currentColor" }: { size?: number; color?: string }) {
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" {...SVG_DEFAULTS} stroke={color} aria-hidden>
      <path d="M10 3L5 8l5 5" />
    </svg>
  );
}

function IconChevronDown({ size = 12, color = "currentColor" }: { size?: number; color?: string }) {
  return (
    <svg width={size} height={size} viewBox="0 0 12 12" {...SVG_DEFAULTS} stroke={color} aria-hidden>
      <path d="M2 4l4 4 4-4" />
    </svg>
  );
}

function IconChevronUp({ size = 12, color = "currentColor" }: { size?: number; color?: string }) {
  return (
    <svg width={size} height={size} viewBox="0 0 12 12" {...SVG_DEFAULTS} stroke={color} aria-hidden>
      <path d="M2 8l4-4 4 4" />
    </svg>
  );
}

function IconChevronRight({ size = 12, color = "currentColor" }: { size?: number; color?: string }) {
  return (
    <svg width={size} height={size} viewBox="0 0 12 12" {...SVG_DEFAULTS} stroke={color} aria-hidden>
      <path d="M4 2l4 4-4 4" />
    </svg>
  );
}

function IconAlertTriangle({ size = 16, color = "currentColor" }: { size?: number; color?: string }) {
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" {...SVG_DEFAULTS} stroke={color} aria-hidden>
      <path d="M8 2L1.5 13h13L8 2z" /><line x1="8" y1="7" x2="8" y2="10" /><circle cx="8" cy="12" r="0.5" fill={color} stroke="none" />
    </svg>
  );
}

function IconSparkle({ size = 12, color = "currentColor" }: { size?: number; color?: string }) {
  return (
    <svg width={size} height={size} viewBox="0 0 12 12" {...SVG_DEFAULTS} stroke={color} aria-hidden>
      <path d="M6 1v2M6 9v2M1 6h2M9 6h2M2.5 2.5l1.4 1.4M8.1 8.1l1.4 1.4M9.5 2.5L8.1 3.9M3.9 8.1L2.5 9.5" />
    </svg>
  );
}

function IconExternalLink({ size = 12, color = "currentColor" }: { size?: number; color?: string }) {
  return (
    <svg width={size} height={size} viewBox="0 0 12 12" {...SVG_DEFAULTS} stroke={color} aria-hidden>
      <path d="M7 2h3v3M10 2L5.5 6.5M6 3H3a1 1 0 0 0-1 1v5a1 1 0 0 0 1 1h5a1 1 0 0 0 1-1V7" />
    </svg>
  );
}

function IconLink({ size = 12, color = "currentColor" }: { size?: number; color?: string }) {
  return (
    <svg width={size} height={size} viewBox="0 0 12 12" {...SVG_DEFAULTS} stroke={color} aria-hidden>
      <path d="M5 6.5a2.5 2.5 0 0 0 3.5.3l1.5-1.5A2.5 2.5 0 0 0 6.5 2L5.8 2.7" />
      <path d="M7 5.5a2.5 2.5 0 0 0-3.5-.3L2 6.7A2.5 2.5 0 0 0 5.5 10l.7-.7" />
    </svg>
  );
}

function StatusDot({ color, size = 6 }: { color: string; size?: number }) {
  return (
    <span style={{ display: "inline-block", width: size, height: size, borderRadius: "50%", background: color, flexShrink: 0 }} />
  );
}

type EntityType = 'goal' | 'initiative' | 'project' | 'capability' | 'epic';

const ENTITY_COLORS: Record<EntityType, { color: string; bg: string }> = {
  goal:        { color: C.textSecondary, bg: C.bgPanel },
  initiative:  { color: C.textSecondary, bg: C.bgPanel },
  project:     { color: C.textSecondary, bg: C.bgPanel },
  capability:  { color: C.textSecondary, bg: C.bgPanel },
  epic:        { color: C.textSecondary, bg: C.bgPanel },
};

function TypeBadge({ type }: { type: EntityType }) {
  const ec = ENTITY_COLORS[type];
  return (
    <span style={{
      fontSize: 10, fontWeight: 600, fontFamily: C.font,
      color: ec.color, backgroundColor: ec.bg,
      border: `1px solid ${C.border}`,
      padding: '1px 6px', borderRadius: C.radiusChip,
      textTransform: 'uppercase' as const, letterSpacing: '0.08em',
      whiteSpace: 'nowrap' as const, flexShrink: 0, lineHeight: '16px',
      display: 'inline-flex', alignItems: 'center',
      alignSelf: 'flex-start',
    }}>
      {type}
    </span>
  );
}

// --- Constants ---
const FIELD_COLORS: Record<string, { bg: string; fg: string }> = {
  blue:   { bg: C.blueSoft,    fg: C.blue },
  cyan:   { bg: '#D0F0FD',     fg: '#0B76B7' },
  teal:   { bg: '#c2f5e9',     fg: '#06846a' },
  green:  { bg: C.greenSoft,   fg: C.greenDark },
  yellow: { bg: C.amberSoft,   fg: C.amberDark },
  orange: { bg: '#fee2d5',     fg: '#d14d00' },
  red:    { bg: C.redSoft,     fg: C.redDark },
  purple: { bg: C.purpleSoft,  fg: C.purpleDark },
  gray:   { bg: '#f3f4f6',     fg: C.textSecondary },
};
const FIELD_COLOR_KEYS = Object.keys(FIELD_COLORS);

function getFieldColor(s: string): { bg: string; fg: string } {
  return FIELD_COLORS[FIELD_COLOR_KEYS[Math.abs(hashToIndex(s, FIELD_COLOR_KEYS.length))]];
}

const STATUS: Record<InitiativeStatus, StatusConfig> = {
  "on-track": {
    color: C.greenDark,
    bg: C.greenSoft,
    dot: C.green,
    label: "On Track",
  },
  "at-risk": {
    color: C.amberDark,
    bg: C.amberSoft,
    dot: C.amber,
    label: "At Risk",
  },
  "off-track": {
    color: C.redDark,
    bg: C.redSoft,
    dot: C.red,
    label: "Off Track",
  },
};

function hashToIndex(s: string, max: number): number {
  let h = 0;
  for (let i = 0; i < s.length; i++) h = ((h << 5) - h + s.charCodeAt(i)) | 0;
  return Math.abs(h) % max;
}

/** Escape a value for CSV: wrap in double quotes and escape internal quotes. */
function csvEscape(value: string): string {
  const safe = (value ?? "").replace(/"/g, '""');
  return `"${safe}"`;
}

/** Build a downloadable CSV from filtered initiatives. */
function buildInitiativesCSV(initiatives: Initiative[]): string {
  const headers = [
    "Initiative Name",
    "Priority",
    "GPA",
    "Segments",
    "User Status",
    "AI Status",
    "Capabilities Count",
    "Product Lead",
  ];
  const rows = initiatives.map((init) => {
    const segments = (init.segments ?? []).join("; ");
    const capsCount = getCapabilities(init).length;
    const aiStatus = ""; // Initiative-level AI status not in current model
    return [
      csvEscape(init.name ?? ""),
      csvEscape(init.goalAlignment ?? ""),
      csvEscape(init.gpa ?? ""),
      csvEscape(segments),
      csvEscape(init.trueStatus ?? ""),
      csvEscape(aiStatus),
      String(capsCount),
      csvEscape(init.productLead ?? ""),
    ];
  });
  return [headers.join(","), ...rows.map((r) => r.join(","))].join("\n");
}

/** Trigger download of initiatives as CSV. */
function downloadInitiativesCSV(initiatives: Initiative[]): void {
  const csvContent = buildInitiativesCSV(initiatives);
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.setAttribute("href", url);
  link.setAttribute("download", `Portfolio_Report_${new Date().toISOString().split("T")[0]}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}


/** Format GMV for display (B/M or raw). */
const formatGMV = (val: number): string =>
  val >= 1_000_000_000
    ? `$${(val / 1_000_000_000).toFixed(1)}B`
    : val >= 1_000_000
      ? `$${(val / 1_000_000).toFixed(1)}M`
      : `$${val.toLocaleString()}`;

/** Build priority groupings from initiatives by goalAlignment. */
function buildPriorities(initiatives: Initiative[]): Priority[] {
  const byGoal = new Map<string, Initiative[]>();
  for (const i of initiatives) {
    const key = i.goalAlignment?.trim() || "(No goal alignment)";
    if (!byGoal.has(key)) byGoal.set(key, []);
    byGoal.get(key)!.push(i);
  }
  const result: Priority[] = [];
  for (const [name, inits] of byGoal) {
    const onTrack = inits.filter((i) => i.trueStatus === "on-track").length;
    const atRisk = inits.filter((i) => i.trueStatus === "at-risk").length;
    const offTrack = inits.filter((i) => i.trueStatus === "off-track").length;
    const capabilitiesTotal = inits.reduce((s, i) => s + getCapabilities(i).length, 0);
    const segments: SegmentBreakdown = {};
    inits.forEach((i) => i.segments.forEach((seg) => { segments[seg] = (segments[seg] ?? 0) + 1; }));
    const gpaSet = [...new Set(inits.map((i) => i.gpa).filter(Boolean))];
    const overall: InitiativeStatus =
      offTrack > inits.length * 0.25 ? "off-track" : atRisk > inits.length * 0.25 ? "at-risk" : "on-track";
    result.push({
      name,
      initiatives: inits,
      count: inits.length,
      capabilitiesTotal,
      onTrack,
      atRisk,
      offTrack,
      segments,
      gpa: gpaSet,
      overallStatus: overall,
      keyMetrics: [],
    });
  }
  return result;
}

// --- UI primitives ---
interface DotProps {
  status: InitiativeStatus;
  size?: number;
}
function Dot({ status, size = 8 }: DotProps) {
  return (
    <span
      style={{
        display: "inline-block",
        width: size,
        height: size,
        borderRadius: "50%",
        background: STATUS[status]?.dot ?? C.border,
        flexShrink: 0,
      }}
    />
  );
}

function Badge({ status }: { status: InitiativeStatus }) {
  const s = STATUS[status] ?? STATUS["on-track"];
  const variant =
    status === "on-track" ? "badge-success" : status === "at-risk" ? "badge-warning" : "badge-danger";
  return (
    <span className={`ld-badge ${variant}`} style={{ display: "inline-flex", alignItems: "center", gap: 4 }}>
      <StatusDot color={s.dot} size={5} />
      {s.label}
    </span>
  );
}

function Bar({
  onTrack,
  atRisk,
  offTrack,
}: {
  onTrack: number;
  atRisk: number;
  offTrack: number;
}) {
  const totalStatus = (onTrack ?? 0) + (atRisk ?? 0) + (offTrack ?? 0);
  const safeTotal = totalStatus > 0 ? totalStatus : 1;
  const greenPct = totalStatus === 0 ? 0 : ((onTrack ?? 0) / safeTotal) * 100;
  const yellowPct = totalStatus === 0 ? 0 : ((atRisk ?? 0) / safeTotal) * 100;
  const redPct = totalStatus === 0 ? 0 : ((offTrack ?? 0) / safeTotal) * 100;
  return (
    <div
      className="progress-bar-container"
      style={{
        display: "flex",
        height: "4px",
        borderRadius: C.radiusChip,
        overflow: "hidden",
        backgroundColor: C.borderLight,
      }}
    >
      {totalStatus === 0 ? (
        <div
          style={{ width: "100%", backgroundColor: C.border }}
          title="No status data available"
        />
      ) : (
        <>
          <div style={{ width: `${greenPct}%`, backgroundColor: C.green }} />
          <div style={{ width: `${yellowPct}%`, backgroundColor: C.amber }} />
          <div style={{ width: `${redPct}%`, backgroundColor: C.red }} />
        </>
      )}
    </div>
  );
}

function getComplexityLabel(depScore: number): "Low" | "Medium" | "High" {
  if (depScore >= 67) return "High";
  if (depScore >= 34) return "Medium";
  return "Low";
}

function DependencyTags({ depCount, depScore }: { depCount: number; depScore: number }) {
  const complexity = getComplexityLabel(depScore);
  const complexityClass =
    complexity === "Low"
      ? "dep-tag--complexity-low"
      : complexity === "Medium"
        ? "dep-tag--complexity-medium"
        : "dep-tag--complexity-high";
  return (
    <div className="dep-tags">
      <span className="dep-tag dep-tag--count">{depCount} GPA Dependencies</span>
      <span className={`dep-tag dep-tag--complexity ${complexityClass}`}>{complexity} Complexity</span>
    </div>
  );
}

/** Compact dependency pill: "N deps · High" */
function DependencyPill({ depCount, depScore }: { depCount: number; depScore: number }) {
  const complexity = getComplexityLabel(depScore);
  const variant =
    complexity === "High" ? "badge-danger" : complexity === "Medium" ? "badge-warning" : "badge-neutral";
  return (
    <span className={`ld-badge ${variant}`} style={{ display: "inline-flex", alignItems: "center", gap: 4 }}>
      <IconLink size={11} color="currentColor" /> {depCount} dep{depCount !== 1 ? "s" : ""} · {complexity}
    </span>
  );
}

function epicStatusColor(status: string | null | undefined): string {
  const s = (status ?? "").toLowerCase();
  return s === "red" ? C.red : s === "yellow" ? C.amber : C.green;
}

function capStatusClass(status: string | null | undefined): string {
  const s = (status ?? "").toLowerCase().replace(/\s+/g, "-");
  return `ld-badge badge-${s || "neutral"}`;
}

interface CapabilityRowProps {
  cap: {
    id: string;
    name: string;
    status?: string | null;
    aiStatus?: string | null;
    startQ?: string | null;
    launchQ?: string | null;
    size?: string | null;
    depCount: number;
    depScore: number;
    statusNotes?: string | null;
    epics?: Array<{ id: string; name: string; status?: string | null }>;
    metrics?: Array<{ id: string; name: string; baseline: number | null; target: number | null; actual: number | null; status: InitiativeStatus }>;
  };
  isLast?: boolean;
}

function CapabilityRow({ cap, isLast }: CapabilityRowProps) {
  const statusBorderColor =
    (cap.status ?? "").toLowerCase() === "red" ? C.red
    : (cap.status ?? "").toLowerCase() === "yellow" ? C.amber
    : C.green;

  return (
    <div style={{
      paddingTop: 12, paddingBottom: 12,
      borderBottom: isLast ? "none" : `1px solid ${C.borderLight}`,
    }}>
      {/* Badge above name */}
      <TypeBadge type="capability" />
      {/* Name + status row */}
      <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 8, marginTop: 4, marginBottom: 6 }}>
        <span style={{ fontSize: 13, fontWeight: 600, color: C.textPrimary, lineHeight: 1.4, flex: 1, minWidth: 0 }}>{cap.name}</span>
        <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
          {cap.aiStatus && cap.aiStatus !== cap.status && (
            <span className="ld-badge badge-ai" title="AI detects a discrepancy" style={{ display: "inline-flex", alignItems: "center", gap: 3 }}>
              <IconSparkle size={10} color="currentColor" /> AI
            </span>
          )}
          <span className={capStatusClass(cap.status)}>{cap.status || "Unassigned"}</span>
        </div>
      </div>
      {/* Meta row */}
      <div style={{ display: "flex", alignItems: "center", flexWrap: "wrap", gap: 8, fontSize: 12, color: C.textSecondary }}>
        {cap.startQ && cap.launchQ && (
          <span>{cap.startQ} → {cap.launchQ}</span>
        )}
        {cap.size && <span className="ld-badge badge-neutral" style={{ fontSize: 11 }}>{cap.size}</span>}
        <DependencyPill depCount={cap.depCount} depScore={cap.depScore} />
      </div>
      {/* Status notes */}
      {cap.statusNotes && (
        <div style={{
          marginTop: 8, padding: "8px 12px",
          backgroundColor: C.bgPage,
          borderLeft: `3px solid ${statusBorderColor}`,
          fontSize: 12, color: C.textSecondary, lineHeight: 1.5,
        }}>
          {cap.statusNotes}
        </div>
      )}
      {/* Epics */}
      {cap.epics && cap.epics.length > 0 && (
        <div style={{ marginTop: 8, paddingLeft: 12, borderLeft: `2px solid ${C.borderLight}` }}>
          {cap.epics.map((epic, i) => (
            <div key={epic.id} style={{
              display: "flex", alignItems: "center", gap: 8,
              padding: "5px 0",
              borderTop: i > 0 ? `1px solid ${C.borderLight}` : "none",
              fontSize: 12,
            }}>
              <StatusDot color={epicStatusColor(epic.status)} size={6} />
              <span style={{ color: C.textPrimary, fontWeight: 500, flex: 1, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{epic.name}</span>
              <span style={{ fontSize: 11, color: C.textSecondary, flexShrink: 0 }}>{epic.status}</span>
            </div>
          ))}
        </div>
      )}
      {/* Metrics table */}
      {cap.metrics && cap.metrics.length > 0 && (
        <div style={{ marginTop: 10 }}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 64px 64px 64px 90px", gap: 8, fontSize: 10, fontWeight: 600, color: C.textTertiary, textTransform: "uppercase", letterSpacing: "0.06em", paddingBottom: 4, borderBottom: `1px solid ${C.border}` }}>
            <span>Metric</span><span style={{ textAlign: "right" }}>Baseline</span><span style={{ textAlign: "right" }}>Target</span><span style={{ textAlign: "right" }}>Actual</span><span>Status</span>
          </div>
          {cap.metrics.map((metric) => (
            <div key={metric.id} style={{ display: "grid", gridTemplateColumns: "1fr 64px 64px 64px 90px", gap: 8, fontSize: 12, alignItems: "center", padding: "6px 0", borderBottom: `1px solid ${C.borderLight}` }}>
              <span style={{ color: C.textPrimary, fontWeight: 500, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{metric.name}</span>
              <span style={{ textAlign: "right", color: C.textSecondary }}>{metric.baseline !== null ? metric.baseline.toLocaleString() : "—"}</span>
              <span style={{ textAlign: "right", color: C.textSecondary }}>{metric.target !== null ? metric.target.toLocaleString() : "—"}</span>
              <span style={{ textAlign: "right", color: C.blue, fontWeight: 600 }}>{metric.actual !== null ? metric.actual.toLocaleString() : "—"}</span>
              <Badge status={metric.status} />
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function DualStatusBadge({
  userStatus,
  aiStatus,
  variant = "default",
}: {
  userStatus: InitiativeStatus;
  aiStatus?: InitiativeStatus | null;
  variant?: "default" | "card";
}) {
  const hasDiscrepancy = aiStatus != null && aiStatus !== userStatus;
  if (!hasDiscrepancy) return <Badge status={userStatus} />;
  const aiLabel = STATUS[aiStatus]?.label ?? "AI";
  if (variant === "card") {
    return (
      <span className="dual-status dual-status--card">
        <Badge status={userStatus} />
        <span className="dual-status__card-warning" title={`AI recommends: ${aiLabel}`} style={{ display: "inline-flex", alignItems: "center", gap: 4 }}>
          <IconAlertTriangle size={12} color="currentColor" /> AI flags as {aiLabel}
        </span>
      </span>
    );
  }
  return (
    <span className="dual-status">
      <Badge status={userStatus} />
      <span className="dual-status__warning" title={`AI recommends: ${aiLabel}`}>!</span>
      <span className="dual-status__ai-badge">AI: {aiLabel}</span>
    </span>
  );
}

interface StatRowData {
  count: number;
  capabilitiesTotal: number;
  onTrack: number;
  atRisk: number;
  offTrack: number;
  /** When set with totalProjects, StatRow uses 7-column layout (Groups, Initiatives, Projects, Capabilities + 3 dots). */
  totalGroups?: number;
  totalProjects?: number;
}
const statRowHeaderLabelStyleSmall: React.CSSProperties = {
  fontSize: 10,
  fontWeight: 600,
  color: C.textSecondary,
  textTransform: "uppercase",
  letterSpacing: "0.03em",
};
const statRowHeaderLabelStyle: React.CSSProperties = {
  fontSize: 10,
  fontWeight: 600,
  color: C.textSecondary,
  textTransform: "uppercase",
  letterSpacing: "0.5px",
};
function StatRow({
  data,
  onStat,
}: {
  data: StatRowData;
  onStat?: (key: StatKey) => void;
}) {
  const onTrackCount = data.onTrack ?? 0;
  const atRiskCount = data.atRisk ?? 0;
  const offTrackCount = data.offTrack ?? 0;
  const useExtendedLayout = data.totalGroups != null && data.totalProjects != null;
  const headerStyle = useExtendedLayout ? statRowHeaderLabelStyleSmall : statRowHeaderLabelStyle;
  const gridCols = useExtendedLayout
    ? "repeat(4, minmax(50px, 1fr)) repeat(3, minmax(40px, 1fr))"
    : "minmax(75px, 1.5fr) minmax(85px, 1.5fr) 1fr 1fr 1fr";
  const statBoxBase = {
    borderRight: `1px solid ${C.borderLight}`,
    padding: useExtendedLayout ? "8px 12px" : "16px 20px",
    minWidth: 0,
  };
  const statBoxLast = useExtendedLayout
    ? { padding: "8px 12px", minWidth: 0 }
    : { padding: "16px 20px", minWidth: 0 };
  const dotCol = {
    cursor: onStat ? "pointer" : "default",
    display: "flex",
    alignItems: "center",
    gap: 4,
    ...statBoxBase,
  };
  if (useExtendedLayout) {
    return (
      <div
        className="stat-row-container"
        style={{
          display: "grid",
          gridTemplateColumns: gridCols,
          borderBottom: `1px solid ${C.borderLight}`,
          background: C.bgPanel,
        }}
      >
        <div className="stat-box" style={statBoxBase}>
          <div className="stat-label" style={headerStyle}>Groups</div>
          <div className="stat-value-row" style={{ fontSize: 20, fontWeight: 700, color: C.textPrimary }}>{data.totalGroups}</div>
        </div>
        <div className="stat-box" style={statBoxBase}>
          <div className="stat-label" style={headerStyle}>Initiatives</div>
          <div className="stat-value-row" style={{ fontSize: 20, fontWeight: 700, color: C.textPrimary }}>{data.count ?? 0}</div>
        </div>
        <div className="stat-box" style={statBoxBase}>
          <div className="stat-label" style={headerStyle}>Projects</div>
          <div className="stat-value-row" style={{ fontSize: 20, fontWeight: 700, color: C.textPrimary }}>{data.totalProjects}</div>
        </div>
        <div className="stat-box" style={statBoxBase}>
          <div className="stat-label" style={headerStyle}>Capabilities</div>
          <div className="stat-value-row" style={{ fontSize: 20, fontWeight: 700, color: C.textPrimary }}>{data.capabilitiesTotal ?? 0}</div>
        </div>
        <div onClick={() => onStat?.("on-track")} className="stat-box" style={dotCol}>
          <StatusDot color={C.green} size={7} />
          <span style={{ fontSize: 20, fontWeight: 700, color: C.textPrimary }}>{onTrackCount}</span>
        </div>
        <div onClick={() => onStat?.("at-risk")} className="stat-box" style={dotCol}>
          <StatusDot color={C.amber} size={7} />
          <span style={{ fontSize: 20, fontWeight: 700, color: C.textPrimary }}>{atRiskCount}</span>
        </div>
        <div
          onClick={() => onStat?.("off-track")}
          className="stat-box"
          style={{ ...statBoxLast, ...dotCol, borderRight: "none" }}
        >
          <StatusDot color={C.red} size={7} />
          <span style={{ fontSize: 20, fontWeight: 700, color: C.textPrimary }}>{offTrackCount}</span>
        </div>
      </div>
    );
  }
  return (
    <div
      className="stat-row-container"
      style={{
        display: "grid",
        gridTemplateColumns: gridCols,
        borderBottom: `1px solid ${C.borderLight}`,
        background: C.bgPanel,
      }}
    >
      <div className="stat-box" style={statBoxBase}>
        <div className="stat-label" style={headerStyle}>Initiatives</div>
        <div className="stat-value-row" style={{ fontSize: 20, fontWeight: 700, color: C.textPrimary }}>{data.count ?? 0}</div>
      </div>
      <div className="stat-box" style={statBoxBase}>
        <div className="stat-label" style={headerStyle}>Capabilities</div>
        <div className="stat-value-row" style={{ fontSize: 20, fontWeight: 700, color: C.textPrimary }}>{data.capabilitiesTotal ?? 0}</div>
      </div>
      <div onClick={() => onStat?.("on-track")} className="stat-box" style={dotCol}>
        <StatusDot color={C.green} size={7} />
        <span style={{ fontSize: 20, fontWeight: 700, color: C.textPrimary }}>{onTrackCount}</span>
      </div>
      <div onClick={() => onStat?.("at-risk")} className="stat-box" style={dotCol}>
        <StatusDot color={C.amber} size={7} />
        <span style={{ fontSize: 20, fontWeight: 700, color: C.textPrimary }}>{atRiskCount}</span>
      </div>
      <div onClick={() => onStat?.("off-track")} className="stat-box" style={statBoxLast}>
        <StatusDot color={C.red} size={7} />
        <span style={{ fontSize: 20, fontWeight: 700, color: C.textPrimary }}>{offTrackCount}</span>
      </div>
    </div>
  );
}

// --- SideSheet: initiative list or initiative detail (Initiative -> Projects -> Capabilities) ---
interface SideSheetProps {
  priority: Priority | null;
  initiative: Initiative | null;
  statKey: StatKey | null;
  contextLabel?: string;
  onClose: () => void;
  onSelectInitiative: (init: Initiative | null) => void;
  onStatKey: (key: StatKey | null) => void;
}

function SideSheet({
  priority,
  initiative,
  statKey,
  contextLabel,
  onClose,
  onSelectInitiative,
  onStatKey,
}: SideSheetProps) {
  const [localFilter, setLocalFilter] = useState<"all" | "attention">("all");
  if (!priority) return null;

  const contextTitle = priority.name;

  const inits = priority.initiatives ?? [];
  const filtered: Initiative[] =
    statKey === "on-track"
      ? inits.filter((i) => i.trueStatus === "on-track")
      : statKey === "at-risk"
        ? inits.filter((i) => i.trueStatus === "at-risk")
        : statKey === "off-track"
          ? inits.filter((i) => i.trueStatus === "off-track")
          : inits;

  const titles: Record<StatKey, string> = {
    initiatives: `All Initiatives (${inits.length})`,
    capabilities: `All Capabilities (${priority.capabilitiesTotal})`,
    "on-track": `On Track Initiatives (${priority.onTrack})`,
    "at-risk": `At Risk Initiatives (${priority.atRisk})`,
    "off-track": `Off Track Initiatives (${priority.offTrack})`,
  };

  const offTrackCount = inits.filter((i) => i.trueStatus === "off-track").length;
  const atRiskCount = inits.filter((i) => i.trueStatus === "at-risk").length;
  const displayedInitiatives = inits.filter((i) => {
    if (localFilter === "all") return true;
    const s = (i.trueStatus ?? "").toLowerCase();
    return s === "off-track" || s === "at-risk";
  });
  const overallStatus: InitiativeStatus = offTrackCount > 0 ? "off-track" : atRiskCount > 0 ? "at-risk" : "on-track";
  const headerBg = STATUS[overallStatus].dot;
  const headerBgLabel = STATUS[overallStatus].label;
  const cleanTitle = priority.name.replace(/^Enterprise Priority:\s*/i, "");

  // Detail view: single initiative (Initiative -> Projects -> Capabilities)
  if (initiative) {
    const isMissing = initiative.financeProjection === "MISSING" || !initiative.financeProjection.trim();
    const detailHeaderBg = STATUS[initiative.trueStatus]?.dot ?? C.green;
    return (
      <div
        className="side-sheet-overlay side-sheet-open"
        onClick={onClose}
        role="presentation"
      >
        <div className="side-sheet-panel" onClick={(e) => e.stopPropagation()}>
          <header className="side-sheet-header">
            <div className="side-sheet-header__title-group">
              <button
                type="button"
                onClick={() => onSelectInitiative(null)}
                className="side-sheet-back-link"
                style={{
                  background: "none", border: "none", fontSize: 11, cursor: "pointer",
                  padding: 0, marginBottom: 8, textAlign: "left",
                  color: C.textSecondary, fontWeight: 500, letterSpacing: "0.02em",
                  fontFamily: "inherit", display: "flex", alignItems: "center", gap: 4,
                }}
              >
                <IconArrowLeft size={12} color={C.textSecondary} /> Back to Priority
              </button>
              <div style={{ display: "flex", flexDirection: "column", gap: 4, marginTop: 4 }}>
                <TypeBadge type="initiative" />
                <div className="side-sheet-header__title" style={{ margin: 0 }}>{initiative.name}</div>
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginTop: 8 }}>
                <span style={{
                  display: "inline-flex", alignItems: "center", gap: 5,
                  fontSize: 10, fontWeight: 700, padding: "2px 8px", borderRadius: C.radiusChip,
                  background: detailHeaderBg + "18", color: detailHeaderBg,
                  border: `1px solid ${detailHeaderBg}40`,
                  textTransform: "uppercase", letterSpacing: "0.08em",
                }}>
                  <StatusDot color={detailHeaderBg} size={5} />
                  {initiative.trueStatus?.replace(/-/g, " ")}
                </span>
                {[initiative.gpa, ...initiative.segments].filter(Boolean).map((g) => (
                  <span
                    key={g}
                    style={{
                      fontSize: 11, background: C.bgPage, color: C.textSecondary,
                      padding: "2px 8px", borderRadius: C.radiusChip, fontWeight: 500,
                      border: `1px solid ${C.border}`,
                    }}
                  >
                    {g}
                  </span>
                ))}
              </div>
            </div>
            <button type="button" className="side-sheet-header__close" onClick={onClose} aria-label="Close">
              <IconX size={14} color={C.textSecondary} />
            </button>
          </header>
          {contextLabel != null && contextLabel !== "" && (
            <div className="side-sheet-context">
              <span className="context-label">{contextLabel}</span>
              <h2 className="context-title">{contextTitle}</h2>
            </div>
          )}
          <div className="side-sheet-body">
            <div style={{ padding: "12px 20px", borderBottom: `1px solid ${C.border}`, display: "flex", flexWrap: "wrap", gap: 16, fontSize: 12, color: C.textSecondary, alignItems: "center" }}>
              <span><span style={{ fontWeight: 600, color: C.textTertiary, marginRight: 4 }}>Lead</span>{initiative.productLead?.trim() || "Unassigned"}</span>
              <span style={{ display: "inline-flex", alignItems: "center", gap: 4 }}><span style={{ fontWeight: 600, color: C.textTertiary, marginRight: 4 }}>Status</span><Badge status={initiative.trueStatus} /></span>
              <span><span style={{ fontWeight: 600, color: C.textTertiary, marginRight: 4 }}>Risk</span>{initiative.risks?.trim() || "None"}</span>
              <span style={{ display: "inline-flex", alignItems: "center", gap: 4 }}><span style={{ fontWeight: 600, color: C.textTertiary, marginRight: 4 }}>Finance</span>{initiative.financeProjection?.trim() && initiative.financeProjection.trim() !== "MISSING" ? initiative.financeProjection.trim() : <span style={{ color: C.amberDark, display: "inline-flex", alignItems: "center", gap: 4 }}><IconAlertTriangle size={12} color={C.amberDark} />Not set</span>}</span>
            </div>
            <div className="hierarchy-container" style={{ display: "flex", flexDirection: "column", gap: "24px", marginTop: "24px" }}>
              {initiative.projects?.map((project) => (
                <div
                  key={project.id}
                  className="project-card"
                  style={{
                    backgroundColor: C.bgPanel,
                    border: `1px solid ${C.border}`,
                    borderRadius: C.radiusCard,
                    padding: "16px",
                  }}
                >
                  {/* Level 2: Project Header */}
                  <div style={{ marginBottom: 16, paddingBottom: 12, borderBottom: `1px solid ${C.borderLight}` }}>
                    <TypeBadge type="project" />
                    <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 4 }}>
                      <h3 style={{ margin: 0, fontSize: 15, color: C.textPrimary, fontWeight: 600, flex: 1, minWidth: 0 }}>
                        {project.name}
                      </h3>
                      <span className={`ld-badge badge-${project.status?.toLowerCase().replace(/\s+/g, "-") || "neutral"}`}>
                        {project.status || "Unassigned"}
                      </span>
                    </div>
                  </div>
                  {/* Level 3: Capabilities List */}
                  <div
                    className="capabilities-list"
                    style={{
                      display: "flex",
                      flexDirection: "column",
                      paddingLeft: 16,
                      borderLeft: `2px solid ${C.borderLight}`,
                      marginLeft: 8,
                    }}
                  >
                    {project.capabilities && project.capabilities.length > 0 ? (
                      project.capabilities.map((cap, idx) => (
                        <CapabilityRow key={cap.id} cap={cap} isLast={idx === project.capabilities!.length - 1} />
                      ))
                    ) : (
                      <div style={{ fontSize: "12px", color: C.textSecondary, fontStyle: "italic", padding: "8px 0" }}>
                        No capabilities linked to this project.
                      </div>
                    )}
                  </div>
                </div>
              ))}
            </div>
            {/* Orphaned Capabilities (Directly Linked to Initiative) */}
            {initiative.orphanedCapabilities && initiative.orphanedCapabilities.length > 0 && (
              <div
                className="hierarchy-container orphaned-capabilities"
                style={{ display: "flex", flexDirection: "column", gap: "24px", marginTop: "24px" }}
              >
                <h3
                  style={{
                    margin: "0 0 12px 0",
                    fontSize: 13,
                    color: C.textSecondary,
                    textTransform: "uppercase",
                    letterSpacing: "0.5px",
                    fontWeight: 700,
                  }}
                >
                  Uncategorized Capabilities (Directly Linked to Initiative)
                </h3>
                <div
                  className="capabilities-list"
                  style={{
                    display: "flex",
                    flexDirection: "column",
                    paddingLeft: 16,
                    borderLeft: `2px solid ${C.borderLight}`,
                    marginLeft: 8,
                  }}
                >
                  {initiative.orphanedCapabilities.map((cap, idx) => (
                    <CapabilityRow key={cap.id} cap={cap} isLast={idx === initiative.orphanedCapabilities!.length - 1} />
                  ))}
                </div>
              </div>
            )}
          </div>
        </div>
      </div>
    );
  }

  // List view: flood-fill header by status, then stats grid, then initiative list
  return (
    <div className="side-sheet-overlay side-sheet-open" onClick={onClose} role="presentation">
      <div className="side-sheet-panel" onClick={(e) => e.stopPropagation()}>
        <div style={{ display: "flex", flexDirection: "column", width: "100%" }}>
          <div style={{ background: C.bgPanel, borderBottom: `1px solid ${C.border}`, padding: "20px 24px", position: "relative" }}>
            <button
              type="button"
              onClick={onClose}
              aria-label="Close"
              className="side-sheet-header__close"
              style={{ position: "absolute", top: 16, right: 16 }}
            >
              <IconX size={14} color={C.textSecondary} />
            </button>
            <div style={{ marginBottom: 8 }}>
              <span style={{
                display: "inline-flex", alignItems: "center", gap: 5,
                fontSize: 10, fontWeight: 700, padding: "2px 8px", borderRadius: C.radiusChip,
                background: headerBg + "18", color: headerBg,
                border: `1px solid ${headerBg}40`,
                textTransform: "uppercase", letterSpacing: "0.08em",
              }}>
                <StatusDot color={headerBg} size={5} />
                {headerBgLabel}
              </span>
            </div>
            <h2 style={{ margin: "0 0 2px 0", fontSize: 20, fontWeight: 700, color: C.textPrimary, lineHeight: 1.2, paddingRight: 48, letterSpacing: "-0.01em" }}>
              {cleanTitle}
            </h2>
            <p style={{ margin: 0, fontSize: 12, color: C.textSecondary, fontWeight: 400 }}>
              Executive Priority Drill-Down
            </p>
          </div>
          {/* Stats Grid - Admin Hub StatCard pattern */}
          <div style={{
            display: "grid",
            gridTemplateColumns: "repeat(3, 1fr)",
            borderBottom: `1px solid ${C.border}`,
          }}>
            <div
              onClick={() => setLocalFilter("all")}
              style={{
                padding: "16px 20px",
                cursor: "pointer",
                borderRight: `1px solid ${C.border}`,
                background: localFilter === "all" ? C.bgPage : C.bgPanel,
                transition: "background 0.12s",
              }}
            >
              <div style={{ fontSize: 11, fontWeight: 600, color: C.textTertiary, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 8 }}>
                Initiatives
              </div>
              <div style={{ fontSize: 20, fontWeight: 700, color: C.textPrimary, lineHeight: 1 }}>
                {inits.length}
              </div>
              <div style={{ fontSize: 11, color: C.textTertiary, marginTop: 4 }}>
                {localFilter === "all" ? "Showing all" : "Click to show all"}
              </div>
            </div>
            <div
              onClick={() => setLocalFilter("attention")}
              style={{
                padding: "16px 20px",
                cursor: "pointer",
                borderRight: `1px solid ${C.border}`,
                background: localFilter === "attention" ? C.bgPage : C.bgPanel,
                transition: "background 0.12s",
              }}
            >
              <div style={{ fontSize: 11, fontWeight: 600, color: C.textTertiary, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 8 }}>
                Needs Attention
              </div>
              <div style={{ fontSize: 20, fontWeight: 700, color: offTrackCount + atRiskCount > 0 ? C.redDark : C.textPrimary, lineHeight: 1 }}>
                {offTrackCount + atRiskCount}
              </div>
              <div style={{ fontSize: 11, color: C.textTertiary, marginTop: 4 }}>
                At risk or off track
              </div>
            </div>
            <div style={{
              padding: "16px 20px",
              background: C.bgPanel,
            }}>
              <div style={{ fontSize: 11, fontWeight: 600, color: C.textTertiary, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 8 }}>
                Overall Status
              </div>
              <div style={{ fontSize: 13, fontWeight: 700, lineHeight: 1, marginBottom: 4 }}>
                <Badge status={overallStatus} />
              </div>
              <div style={{ fontSize: 11, color: C.textTertiary, marginTop: 8 }}>
                Portfolio health
              </div>
            </div>
          </div>
        </div>
        <div className="side-sheet-body" style={{ padding: 0, gap: 0 }}>
            {displayedInitiatives.map((init) => {
              const hasAiDiscrepancy = init.aiStatus && init.aiStatus.toLowerCase().replace(/\s+/g, "-") !== init.trueStatus;
              return (
                <button
                  key={init.id}
                  type="button"
                  onClick={() => onSelectInitiative(init)}
                  className="side-sheet-initiative-row"
                >
                  <div style={{ minWidth: 0, flex: 1 }}>
                    {/* Top row: name + status badge */}
                    <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 12 }}>
                      <div style={{ display: "flex", flexDirection: "column", gap: 4, minWidth: 0, flex: 1 }}>
                        <TypeBadge type="initiative" />
                        <div style={{ fontWeight: 600, fontSize: 13, color: C.textPrimary, lineHeight: 1.4 }}>{init.name}</div>
                      </div>
                      <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
                        <Badge status={init.trueStatus} />
                        {hasAiDiscrepancy && (
                          <span className="ld-badge badge-ai" title={`AI flags as ${init.aiStatus}`} style={{ display: "inline-flex", alignItems: "center", gap: 3 }}><IconSparkle size={10} color="currentColor" /> AI</span>
                        )}
                      </div>
                    </div>
                    {/* Bottom row: meta */}
                    <div style={{ marginTop: 6, display: "flex", alignItems: "center", gap: 12, fontSize: 12, color: C.textSecondary, flexWrap: "wrap" }}>
                      <span>{init.gpa}</span>
                      {init.segments.length > 0 && <><span style={{ color: C.border }}>·</span><span>{init.segments.join(", ")}</span></>}
                      {init.productLead?.trim() && <><span style={{ color: C.border }}>·</span><span>{init.productLead.trim()}</span></>}
                      {init.rawFinance && <><span style={{ color: C.border }}>·</span><span style={{ fontWeight: 500, color: C.textPrimary }}>{init.rawFinance}</span></>}
                    </div>
                  </div>
                  <IconChevronRight size={14} color={C.textSecondary} />
                </button>
              );
            })}
        </div>
      </div>
    </div>
  );
}

// --- PCard (priority card): click opens SideSheet ---
function PCard({
  priority,
  onClick,
  isSelected,
}: {
  priority: Priority;
  onClick: (p: Priority) => void;
  isSelected: boolean;
}) {
  return (
    <div
      onClick={() => onClick(priority)}
      className={`ld-card pcard ${isSelected ? "selected" : ""}`}
    >
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 }}>
        <h3 className="ld-heading-s" style={{ margin: 0, lineHeight: 1.3 }}>
          {priority.name}
        </h3>
        <Badge status={priority.overallStatus} />
      </div>
      <div className="pcard-stats" style={{ display: "flex", gap: 18 }}>
        {[
          ["Initiatives", priority.count],
          ["Capabilities", priority.capabilitiesTotal],
        ].map(([l, v], i) => (
          <div key={i}>
            <div className="ld-caption" style={{ fontSize: 10, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.06em" }}>
              {l}
            </div>
            <div className="ld-heading-m" style={{ fontSize: 20, fontWeight: 700 }}>{String(v)}</div>
          </div>
        ))}
      </div>
      <div>
        <Bar
          onTrack={priority.onTrack}
          atRisk={priority.atRisk}
          offTrack={priority.offTrack}
        />
        <div className="ld-caption stat-legend" style={{ display: "flex", gap: 10, marginTop: 5, fontWeight: 600 }}>
          <span className="text-success">{priority.onTrack} on track</span>
          <span className="text-warning">{priority.atRisk} at risk</span>
          <span className="text-danger">{priority.offTrack} off track</span>
        </div>
      </div>
    </div>
  );
}

// --- AggCard (GPA / Segment / Priority view) ---
function AggCard({
  title,
  priorities,
  initCount,
  capCount,
  onTrack,
  atRisk,
  offTrack,
  segBreakdown,
  scopedInits,
  onOpenSheet,
}: {
  title: string;
  priorities: Priority[];
  initCount: number;
  capCount: number;
  onTrack: number;
  atRisk: number;
  offTrack: number;
  segBreakdown: SegmentBreakdown;
  scopedInits?: Initiative[];
  onOpenSheet: (priority: Priority) => void;
}) {
  const total = onTrack + atRisk + offTrack;
  const overallStatus: InitiativeStatus =
    total > 0 && offTrack > total * 0.25
      ? "off-track"
      : total > 0 && atRisk > total * 0.25
        ? "at-risk"
        : "on-track";
  const initiatives = scopedInits ?? priorities.flatMap((p) => p.initiatives);
  const totalGroups = new Set(initiatives.flatMap((i) => i.initiativeGroups || [])).size;
  const totalProjects = new Set(initiatives.flatMap((i) => i.projectIds || (i.projects || []).map((p) => p.id))).size;
  const totalCapabilities = capCount;
  const financeTotals = initiatives.reduce((acc, init) => {
    if (init.targetGmv > 0 && init.financeType) {
      acc[init.financeType] = (acc[init.financeType] ?? 0) + init.targetGmv;
    }
    return acc;
  }, {} as Record<string, number>);
  const financeStrings = Object.entries(financeTotals).map(([type, val]) => `${formatGMV(val)} ${type}`);
  const displayFinance = financeStrings.length > 0 ? financeStrings.join(" | ") : "No projections";

  const offTrackCount = initiatives.filter((i) => i.trueStatus === "off-track").length;
  const atRiskCount = initiatives.filter((i) => i.trueStatus === "at-risk").length;
  let statusText = "ALL ON TRACK";
  let borderColor: string = C.green;
  if (offTrackCount > 0) {
    statusText = `${offTrackCount} OFF TRACK`;
    borderColor = C.red;
  } else if (atRiskCount > 0) {
    statusText = `${atRiskCount} AT RISK`;
    borderColor = C.amber;
  } else if (initiatives.length === 0) {
    statusText = "NO DATA";
    borderColor = C.border;
  }

  const cleanTitle = title.replace(/^Enterprise Priority:\s*/i, "");
  const [isHovered, setIsHovered] = useState(false);

  const virtualP: Priority = {
    name: title,
    initiatives,
    count: initCount,
    capabilitiesTotal: capCount,
    onTrack,
    atRisk,
    offTrack,
    overallStatus,
    keyMetrics: [],
    gpa: [...new Set(priorities.flatMap((p) => p.gpa))],
    segments: segBreakdown ?? {},
  };
  return (
    <div
      className="ld-card aggcard"
      role="button"
      tabIndex={0}
      onClick={() => onOpenSheet(virtualP)}
      onKeyDown={(e) => {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          onOpenSheet(virtualP);
        }
      }}
      onMouseEnter={() => setIsHovered(true)}
      onMouseLeave={() => setIsHovered(false)}
      style={{
        display: "flex",
        flexDirection: "column",
        height: "100%",
        padding: "16px 20px",
        backgroundColor: C.bgPanel,
        borderRadius: C.radiusCard,
        border: `1px solid ${C.border}`,
        cursor: "pointer",
        boxShadow: isHovered ? C.shadowCardHover : C.shadowCard,
        transition: "box-shadow 0.15s, border-color 0.12s",
        opacity: initiatives.length === 0 ? 0.6 : 1,
      }}
    >
      {/* Top: Clean Title + Status Badge */}
      <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 8, marginBottom: 16 }}>
        <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-start", gap: 8, minWidth: 0, flex: 1 }}>
          <TypeBadge type="goal" />
          <div style={{ fontSize: 15, fontWeight: 600, color: C.textPrimary, lineHeight: 1.4 }}>
            {cleanTitle}
          </div>
        </div>
        <div
          style={{
            display: "flex", alignItems: "center", gap: 5,
            fontSize: 10, fontWeight: 700, color: borderColor,
            textTransform: "uppercase", letterSpacing: "0.07em",
            padding: "2px 8px", background: borderColor + "14",
            border: `1px solid ${borderColor}35`, borderRadius: C.radiusChip,
            whiteSpace: "nowrap", flexShrink: 0,
          }}
        >
          <StatusDot color={borderColor} size={5} />
          {statusText}
        </div>
      </div>

      {/* Initiative count */}
      <div style={{ display: "flex", alignItems: "baseline", gap: 8, marginBottom: 20 }}>
        <div style={{ fontSize: 20, fontWeight: 700, color: C.textPrimary, lineHeight: 1 }}>
          {initiatives.length}
        </div>
        <div style={{ fontSize: 12, fontWeight: 600, color: C.textSecondary, textTransform: "uppercase", letterSpacing: "0.5px" }}>
          Initiatives
        </div>
      </div>

      {/* Status progress bar */}
      <div style={{ display: "flex", height: 4, borderRadius: 2, overflow: "hidden", backgroundColor: C.borderLight, marginBottom: 12 }}>
        {total > 0 ? (
          <>
            <div style={{ width: `${(onTrack / total) * 100}%`, backgroundColor: C.green, transition: "width 0.2s" }} />
            <div style={{ width: `${(atRisk / total) * 100}%`, backgroundColor: C.amber, transition: "width 0.2s" }} />
            <div style={{ width: `${(offTrack / total) * 100}%`, backgroundColor: C.red, transition: "width 0.2s" }} />
          </>
        ) : (
          <div style={{ width: "100%", backgroundColor: C.borderLight }} />
        )}
      </div>
      {/* Status legend */}
      <div style={{ display: "flex", gap: 12, marginBottom: 20, fontSize: 11, fontWeight: 600 }}>
        <span style={{ display: "flex", alignItems: "center", gap: 4, color: C.greenDark }}><StatusDot color={C.green} size={5} />{onTrack}</span>
        <span style={{ display: "flex", alignItems: "center", gap: 4, color: C.amberDark }}><StatusDot color={C.amber} size={5} />{atRisk}</span>
        <span style={{ display: "flex", alignItems: "center", gap: 4, color: C.redDark }}><StatusDot color={C.red} size={5} />{offTrack}</span>
      </div>

      {/* Secondary Metrics: Clean borderless row */}
      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(3, 1fr)",
          width: "100%",
          borderTop: `1px solid ${C.borderLight}`,
          borderBottom: `1px solid ${C.borderLight}`,
          marginBottom: "20px",
          padding: "10px 0",
        }}
      >
        {[
          { label: "Groups", value: totalGroups },
          { label: "Projects", value: totalProjects },
          { label: "Capabilities", value: totalCapabilities },
        ].map((item, i) => (
          <div
            key={item.label}
            style={{
              display: "flex",
              flexDirection: "column",
              textAlign: "center",
              borderLeft: i > 0 ? `1px solid ${C.borderLight}` : "none",
            }}
          >
            <span style={{ fontSize: 10, letterSpacing: "0.06em", textTransform: "uppercase", color: C.textSecondary, fontWeight: 600 }}>
              {item.label}
            </span>
            <span style={{ fontSize: 15, fontWeight: 700, color: C.textPrimary, marginTop: 4, lineHeight: 1 }}>
              {item.value}
            </span>
          </div>
        ))}
      </div>

      {financeStrings.length > 0 && (
        <div style={{ marginTop: "auto", paddingTop: 16, borderTop: `1px solid ${C.borderLight}`, display: "flex", justifyContent: "space-between", fontSize: 11, fontWeight: 600, color: C.textSecondary, textTransform: "uppercase", letterSpacing: "0.5px" }}>
          <span>TARGET: {displayFinance}</span>
          <span>ACTUAL: TBD</span>
        </div>
      )}
    </div>
  );
}

function GPAView({
  initiatives,
  onOpenSheet,
}: {
  initiatives: Initiative[];
  onOpenSheet: (p: Priority) => void;
}) {
  const byGpa = new Map<string, Initiative[]>();
  for (const init of initiatives) {
    const g = init.gpa?.trim() || "Unassigned";
    if (!byGpa.has(g)) byGpa.set(g, []);
    byGpa.get(g)!.push(init);
  }
  const items = Array.from(byGpa.entries()).map(([gpa, inits]) => {
    const onT = inits.filter((i) => i.trueStatus === "on-track").length;
    const atR = inits.filter((i) => i.trueStatus === "at-risk").length;
    const offT = inits.filter((i) => i.trueStatus === "off-track").length;
    const caps = inits.reduce((s, i) => s + getCapabilities(i).length, 0);
    const segC: SegmentBreakdown = {};
    inits.forEach((i) => i.segments.forEach((s) => { segC[s] = (segC[s] ?? 0) + 1; }));
    const virtualP: Priority = {
      name: gpa,
      initiatives: inits,
      count: inits.length,
      capabilitiesTotal: caps,
      onTrack: onT,
      atRisk: atR,
      offTrack: offT,
      segments: segC,
      gpa: [gpa],
      overallStatus: offT > inits.length * 0.25 ? "off-track" : atR > inits.length * 0.25 ? "at-risk" : "on-track",
      keyMetrics: [],
    };
    return { gpa, inits, virtualP, onT, atR, offT, caps, segC };
  });
  const sortedItems = [...items].sort((a, b) => b.inits.length - a.inits.length);
  return (
    <div className="card-grid">
      {sortedItems.map(({ gpa, inits, virtualP, onT, atR, offT, caps, segC }) => (
        <AggCard
          key={gpa}
          title={gpa}
          priorities={[virtualP]}
          initCount={inits.length}
          capCount={caps}
          onTrack={onT}
          atRisk={atR}
          offTrack={offT}
          segBreakdown={segC}
          scopedInits={inits}
          onOpenSheet={onOpenSheet}
        />
      ))}
    </div>
  );
}

function SegmentView({
  initiatives,
  onOpenSheet,
}: {
  initiatives: Initiative[];
  onOpenSheet: (p: Priority) => void;
}) {
  const segs = ["Walmart US", "International", "Sam's Club"];
  const items = segs.map((seg) => {
    const inits = initiatives.filter((i) => i.segments?.includes(seg));
    const onT = inits.filter((i) => i.trueStatus === "on-track").length;
    const atR = inits.filter((i) => i.trueStatus === "at-risk").length;
    const offT = inits.filter((i) => i.trueStatus === "off-track").length;
    const caps = inits.reduce((s, i) => s + getCapabilities(i).length, 0);
    const gpaC: SegmentBreakdown = {};
    inits.forEach((i) => {
      gpaC[i.gpa] = (gpaC[i.gpa] ?? 0) + 1;
    });
    const virtualP: Priority = {
      name: seg,
      initiatives: inits,
      count: inits.length,
      capabilitiesTotal: caps,
      onTrack: onT,
      atRisk: atR,
      offTrack: offT,
      segments: { [seg]: inits.length },
      gpa: [...new Set(inits.map((i) => i.gpa))],
      overallStatus: offT > inits.length * 0.25 ? "off-track" : atR > inits.length * 0.25 ? "at-risk" : "on-track",
      keyMetrics: [],
    };
    return { seg, inits, virtualP, onT, atR, offT, caps, gpaC };
  });
  const sortedItems = [...items].sort((a, b) => b.inits.length - a.inits.length);
  return (
    <div className="card-grid">
      {sortedItems.map(({ seg, inits, virtualP, onT, atR, offT, caps, gpaC }) => (
        <AggCard
          key={seg}
          title={seg}
          priorities={[virtualP]}
          initCount={inits.length}
          capCount={caps}
          onTrack={onT}
          atRisk={atR}
          offTrack={offT}
          segBreakdown={gpaC}
          scopedInits={inits}
          onOpenSheet={onOpenSheet}
        />
      ))}
    </div>
  );
}

function Summary({
  priorities,
  seg,
  fy,
  market,
  globalOnTrack,
  globalAtRisk,
  globalOffTrack,
  statusFilter,
  onStatusFilter,
}: {
  priorities: Priority[];
  seg: string;
  fy: string;
  market: string;
  globalOnTrack: number;
  globalAtRisk: number;
  globalOffTrack: number;
  statusFilter: string | null;
  onStatusFilter: (status: string | null) => void;
}) {
  const t = priorities.reduce(
    (a, p) => ({
      i: a.i + p.count,
      c: a.c + p.capabilitiesTotal,
      ot: a.ot + p.onTrack,
      ar: a.ar + p.atRisk,
      of: a.of + p.offTrack,
    }),
    { i: 0, c: 0, ot: 0, ar: 0, of: 0 }
  );
  const tot = t.ot + t.ar + t.of;
  const safeTot = tot > 0 ? tot : 1;
  const greenPct = tot === 0 ? 0 : (t.ot / safeTot) * 100;
  const yellowPct = tot === 0 ? 0 : (t.ar / safeTot) * 100;
  const redPct = tot === 0 ? 0 : (t.of / safeTot) * 100;
  const summaryStats = [
    { label: "Priorities", value: priorities.length, dot: null, status: null as string | null },
    { label: "Initiatives", value: t.i, dot: null, status: null as string | null },
    { label: "Capabilities", value: t.c, dot: null, status: null as string | null },
    { label: "On Track", value: globalOnTrack, dot: C.green, status: "on-track" as string | null },
    { label: "At Risk", value: globalAtRisk, dot: C.amber, status: "at-risk" as string | null },
    { label: "Off Track", value: globalOffTrack, dot: C.red, status: "off-track" as string | null },
  ];
  return (
    <div className="summary-banner ld-card" style={{ padding: 0, overflow: "hidden" }}>
      {/* Header row */}
      <div style={{ padding: "12px 24px", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "baseline", justifyContent: "space-between" }}>
        <h2 style={{ margin: 0, fontSize: 11, fontWeight: 700, letterSpacing: "0.07em", textTransform: "uppercase", color: C.textSecondary }}>
          Portfolio Health
        </h2>
        <p style={{ margin: 0, fontSize: 12, color: C.textTertiary }}>
          {seg === "All" ? "All Segments" : seg} · {fy === "All" ? "All Years" : fy} · {market}
        </p>
      </div>
      {/* Stat row */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(6, 1fr)" }}>
        {summaryStats.map((stat, i) => (
          <div
            key={stat.label}
            onClick={() => stat.status && onStatusFilter(statusFilter === stat.status ? null : stat.status)}
            style={{
              padding: "20px 0",
              textAlign: "center",
              borderRight: i < 5 ? `1px solid ${C.border}` : "none",
              borderLeft: i === 3 ? `1px solid ${C.borderLight}` : "none",
              cursor: stat.status ? "pointer" : "default",
              opacity: statusFilter && statusFilter !== stat.status ? 0.4 : 1,
              transition: "opacity 0.15s, background 0.12s",
              background: stat.status && statusFilter === stat.status ? C.bgPage : "transparent",
            }}
          >
            <div style={{ fontSize: 10, fontWeight: 600, color: stat.dot ?? C.textSecondary, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 10, display: "flex", alignItems: "center", justifyContent: "center", gap: 5 }}>
              {stat.dot && <span style={{ display: "inline-block", width: 6, height: 6, borderRadius: "50%", background: stat.dot, flexShrink: 0 }} />}
              {stat.label}
            </div>
            <div style={{ fontSize: 20, fontWeight: 700, color: C.textPrimary, lineHeight: 1 }}>
              {stat.value}
            </div>
          </div>
        ))}
      </div>
      {/* Progress bar */}
      <div style={{ height: 4, display: "flex", backgroundColor: C.borderLight }}>
        {tot > 0 && (
          <>
            <div style={{ width: `${greenPct}%`, background: C.green, transition: "width 0.2s" }} />
            <div style={{ width: `${yellowPct}%`, background: C.amber, transition: "width 0.2s" }} />
            <div style={{ width: `${redPct}%`, background: C.red, transition: "width 0.2s" }} />
          </>
        )}
      </div>
    </div>
  );
}

function MarketMultiSelect({
  selectedMarkets,
  onMarketsChange,
  marketOptions,
}: {
  selectedMarkets: string[];
  onMarketsChange: (markets: string[]) => void;
  marketOptions: string[];
}) {
  const [isMarketDropdownOpen, setIsMarketDropdownOpen] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const close = (e: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setIsMarketDropdownOpen(false);
      }
    };
    document.addEventListener("click", close);
    return () => document.removeEventListener("click", close);
  }, []);

  const toggleMarket = (market: string) => {
    onMarketsChange(
      selectedMarkets.includes(market)
        ? selectedMarkets.filter((m) => m !== market)
        : [...selectedMarkets, market]
    );
  };

  const uniqueMarkets = marketOptions;
  const chevronSvg = `url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="%23666" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"></polyline></svg>')`;
  return (
    <div ref={containerRef} style={{ position: "relative", display: "inline-block", minWidth: "180px" }}>
      <div
        onClick={(e) => {
          e.stopPropagation();
          setIsMarketDropdownOpen(!isMarketDropdownOpen);
        }}
        style={{
          backgroundColor: C.bgPanel,
          border: `1px solid ${C.border}`,
          borderRadius: C.radiusButton,
          padding: "6px 32px 6px 12px",
          fontSize: "13px",
          fontWeight: 500,
          color: C.textPrimary,
          cursor: "pointer",
          outline: "none",
          boxShadow: "0 1px 2px rgba(0,0,0,0.04)",
          backgroundImage: chevronSvg,
          backgroundRepeat: "no-repeat",
          backgroundPosition: "right 10px center",
          height: "32px",
          minWidth: "110px",
          display: "flex",
          alignItems: "center",
        }}
      >
        <span style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
          {selectedMarkets.length === 0 ? "All Markets" : `${selectedMarkets.length} selected`}
        </span>
      </div>

        {isMarketDropdownOpen && (
          <div
            style={{
              position: "absolute",
              top: "calc(100% + 4px)",
              left: 0,
              right: 0,
              backgroundColor: C.bgPanel,
              border: `1px solid ${C.border}`,
              borderRadius: C.radiusCard,
              boxShadow: "0 4px 16px rgba(0,0,0,0.10), 0 1px 4px rgba(0,0,0,0.06)",
              zIndex: 1000,
              maxHeight: "250px",
              overflowY: "auto",
            }}
          >
            <div
              onClick={() => {
                onMarketsChange([]);
              }}
              style={{
                padding: "8px 12px",
                borderBottom: `1px solid ${C.borderLight}`,
                cursor: "pointer",
                fontSize: "13px",
                fontWeight: selectedMarkets.length === 0 ? 600 : 400,
                backgroundColor: selectedMarkets.length === 0 ? C.bgPage : C.bgPanel,
              }}
            >
              All Markets
            </div>
            {uniqueMarkets.map((market) => (
              <label
                key={market}
                style={{
                  display: "flex",
                  alignItems: "center",
                  padding: "8px 12px",
                  cursor: "pointer",
                  fontSize: "13px",
                  margin: 0,
                }}
                className="market-dropdown-option"
              >
                <input
                  type="checkbox"
                  checked={selectedMarkets.includes(market)}
                  onChange={() => toggleMarket(market)}
                  style={{ marginRight: "8px", cursor: "pointer" }}
                />
                {market || "(Blank)"}
              </label>
            ))}
          </div>
        )}
    </div>
  );
}

function Filters({
  seg,
  onSeg,
  fy,
  onFy,
  selectedMarkets,
  onMarketsChange,
  fyOptions,
  marketOptions,
  view,
  onView,
  onExportCSV,
  onExportPDF,
  onExportPPT,
}: {
  seg: string;
  onSeg: (s: string) => void;
  fy: string;
  onFy: (s: string) => void;
  selectedMarkets: string[];
  onMarketsChange: (markets: string[]) => void;
  fyOptions: string[];
  marketOptions: string[];
  view: ViewMode;
  onView: (v: ViewMode) => void;
  onExportCSV?: () => void;
  onExportPDF?: () => void;
  onExportPPT?: () => void;
}) {
  const segs = ["All", "Walmart US", "International", "Sam's Club"];
  const views: Array<{ id: ViewMode; l: string }> = [
    { id: "priority", l: "By Priority" },
    { id: "gpa", l: "By GPA" },
    { id: "segment", l: "By Segment" },
  ];
  return (
    <div style={{ display: "flex", flexDirection: "column", width: "100%" }}>
      {/* Row 1: Navigation & Export Bar */}
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "flex-end",
          borderBottom: `1px solid ${C.border}`,
          marginBottom: "24px",
          paddingBottom: 0,
        }}
      >
        {/* View Tabs */}
        <div style={{ display: "flex", gap: "32px", paddingLeft: "8px" }}>
          {views.map((v) => (
            <div
              key={v.id}
              onClick={() => onView(v.id)}
              role="button"
              tabIndex={0}
              onKeyDown={(e) => {
                if (e.key === "Enter" || e.key === " ") {
                  e.preventDefault();
                  onView(v.id);
                }
              }}
              style={{
                padding: "12px 0",
                cursor: "pointer",
                fontSize: "13px",
                fontWeight: view === v.id ? 700 : 500,
                color: view === v.id ? C.textPrimary : C.textSecondary,
                borderBottom: view === v.id ? `3px solid ${C.blue}` : "3px solid transparent",
                transition: "background 0.12s, box-shadow 0.15s",
                marginBottom: "-1px",
              }}
            >
              {v.l}
            </div>
          ))}
        </div>
        {/* Actions (Export) */}
        <div style={{ paddingBottom: "12px" }}>
          <select
            className="btn btn-outline"
            style={{
              appearance: "none",
              WebkitAppearance: "none",
              background: "linear-gradient(135deg, #2d7ff9 0%, #1a6fe8 100%)",
              border: "none",
              borderRadius: C.radiusButton,
              padding: "6px 32px 6px 12px",
              color: "#fff",
              boxShadow: "0 1px 4px rgba(45,127,249,0.3)",
              cursor: "pointer",
              fontFamily: C.font,
              fontSize: "13px",
              fontWeight: 600,
              backgroundImage: `url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="%23ffffff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"></polyline></svg>')`,
              backgroundRepeat: "no-repeat",
              backgroundPosition: "right 10px center",
              height: "32px",
              minWidth: "110px",
            }}
            onChange={(e) => {
              const val = e.target.value;
              if (val === "csv") onExportCSV?.();
              if (val === "pdf") onExportPDF?.();
              if (val === "ppt") onExportPPT?.();
              e.target.value = "";
            }}
          >
            <option value="" disabled>
              Export...
            </option>
            <option value="csv">Export CSV</option>
            <option value="pdf">Export PDF</option>
            <option value="ppt">Export PPT</option>
          </select>
        </div>
      </div>
      {/* Row 2: Deep Filter Bar */}
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          marginBottom: "32px",
        }}
      >
        {/* Mac-Style Segmented Control */}
        <div
          style={{
            display: "flex",
            backgroundColor: C.bgPage,
            padding: "4px",
            borderRadius: C.radiusCard,
            border: `1px solid ${C.borderLight}`,
          }}
        >
          {segs.map((s) => (
            <div
              key={s}
              onClick={() => onSeg(s)}
              role="button"
              tabIndex={0}
              onKeyDown={(e) => {
                if (e.key === "Enter" || e.key === " ") {
                  e.preventDefault();
                  onSeg(s);
                }
              }}
              style={{
                padding: "8px 16px",
                cursor: "pointer",
                fontSize: "13px",
                fontWeight: seg === s ? 600 : 500,
                color: seg === s ? C.textPrimary : C.textSecondary,
                backgroundColor: seg === s ? C.bgPanel : "transparent",
                borderRadius: C.radiusButton,
                boxShadow: seg === s ? "0 2px 4px rgba(0,0,0,0.06)" : "none",
                transition: "background 0.12s, box-shadow 0.15s",
              }}
            >
              {s}
            </div>
          ))}
        </div>
        {/* Dropdown Filters */}
        <div style={{ display: "flex", alignItems: "center", gap: "24px" }}>
          <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
            <span
              style={{
                fontSize: "11px",
                fontWeight: 700,
                color: C.textSecondary,
                textTransform: "uppercase",
                letterSpacing: "0.5px",
              }}
            >
              Fiscal Year
            </span>
            <select
              value={fy}
              onChange={(e) => onFy(e.target.value)}
              style={{
                appearance: "none",
                WebkitAppearance: "none",
                backgroundColor: C.bgPanel,
                border: `1px solid ${C.border}`,
                borderRadius: C.radiusButton,
                padding: "6px 32px 6px 12px",
                fontSize: "13px",
                fontWeight: 500,
                color: C.textPrimary,
                cursor: "pointer",
                outline: "none",
                boxShadow: "0 1px 2px rgba(0,0,0,0.04)",
                backgroundImage: `url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="%23666" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"></polyline></svg>')`,
                backgroundRepeat: "no-repeat",
                backgroundPosition: "right 10px center",
                height: "32px",
                minWidth: "110px",
              }}
            >
              <option value="All">All</option>
              {fyOptions.map((f) => (
                <option key={f} value={f}>
                  {f || "(Blank)"}
                </option>
              ))}
            </select>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
            <span
              style={{
                fontSize: "11px",
                fontWeight: 700,
                color: C.textSecondary,
                textTransform: "uppercase",
                letterSpacing: "0.5px",
              }}
            >
              Markets
            </span>
            <MarketMultiSelect
              selectedMarkets={selectedMarkets}
              onMarketsChange={onMarketsChange}
              marketOptions={marketOptions}
            />
          </div>
        </div>
      </div>
    </div>
  );
}

function App() {
  const { initiatives, enterprisePriorities, otherPriorities, isLoading, error } = usePortfolioData();
  const [seg, setSeg] = useState<string>("All");
  const [fy, setFy] = useState<string>("All");
  const [selectedMarkets, setSelectedMarkets] = useState<string[]>([]);
  const [statusFilter, setStatusFilter] = useState<string | null>(null);
  const [view, setView] = useState<ViewMode>("priority");
  const [selectedPriority, setSelectedPriority] = useState<Priority | null>(null);
  const [sideSheetInitiative, setSideSheetInitiative] = useState<Initiative | null>(null);
  const [sideSheetStatKey, setSideSheetStatKey] = useState<StatKey | null>(null);
  const [sideSheetContextLabel, setSideSheetContextLabel] = useState<string>("");
  const [isBannerExpanded, setIsBannerExpanded] = useState(false);

  const fyOptions = useMemo(
    () => [...new Set(initiatives.map((i) => i.fy).filter(Boolean))].sort(),
    [initiatives]
  );
  const marketOptions = useMemo(() => {
    const allMarkets = new Set<string>();
    initiatives.forEach((init) => {
      if (init.market && Array.isArray(init.market)) {
        init.market.forEach((m) => allMarkets.add(m));
      }
    });
    return Array.from(allMarkets).sort();
  }, [initiatives]);

  /** Base filter: Segment + FY + Market only (no statusFilter). Used for summary stats. */
  const filteredBase = useMemo(
    () =>
      initiatives.filter((init) => {
        const matchSeg = seg === "All" || (init.segments && init.segments.includes(seg));
        const matchFy = fy === "All" || init.fy === fy;
        const matchMarket =
          selectedMarkets.length === 0 ||
          (init.market && selectedMarkets.some((m) => init.market!.includes(m)));
        return matchSeg && matchFy && matchMarket;
      }),
    [initiatives, seg, fy, selectedMarkets]
  );

  /** Full filter: base + statusFilter. Used for the card grid. */
  const filteredInitiatives = useMemo(
    () =>
      statusFilter
        ? filteredBase.filter((i) => i.trueStatus === statusFilter)
        : filteredBase,
    [filteredBase, statusFilter]
  );

  const priorities = useMemo(() => buildPriorities(initiatives), [initiatives]);
  const filtered = useMemo(
    () => buildPriorities(filteredInitiatives),
    [filteredInitiatives]
  );

  /** Summary uses base filter so status counts don't collapse when filtering. */
  const summaryPriorities = useMemo(() => buildPriorities(filteredBase), [filteredBase]);

  /** Global status counts drawn from base filter (unaffected by statusFilter). */
  const globalOnTrack = filteredBase.filter((i) => i.trueStatus === "on-track").length;
  const globalAtRisk = filteredBase.filter((i) => i.trueStatus === "at-risk").length;
  const globalOffTrack = filteredBase.filter((i) => i.trueStatus === "off-track").length;

  /** Initiatives that need attention (at-risk or off-track) for the banner expanded list. */
  const attentionItems = useMemo(
    () =>
      filteredBase.filter(
        (i) => i.trueStatus === "at-risk" || i.trueStatus === "off-track"
      ),
    [filteredBase]
  );

  /** Split filtered priorities into Enterprise vs Other (exclude Unassigned). */
  const filteredEnterprisePriorities = useMemo(
    () =>
      filtered.filter(
        (p) => p.name !== "Unassigned" && p.name.startsWith("Enterprise Priority")
      ),
    [filtered]
  );
  const filteredOtherPriorities = useMemo(
    () =>
      filtered.filter(
        (p) => p.name !== "Unassigned" && p.name.startsWith("Other")
      ),
    [filtered]
  );

  const openSheet = (priority: Priority, statKey?: StatKey | null, contextLabel?: string) => {
    setSelectedPriority(priority);
    setSideSheetInitiative(null);
    setSideSheetStatKey(statKey ?? null);
    setSideSheetContextLabel(contextLabel ?? "");
  };

  const handleExportCSV = () => downloadInitiativesCSV(filteredInitiatives);
  const handleExportPDF = () => {
    window.print();
  };
  const handleExportPPT = () => {
    try {
      const pres = new pptxgen();
      const slide = pres.addSlide();
      slide.addText("Executive Portfolio Report", { x: 1, y: 1, fontSize: 24, bold: true, color: "001E60" });
      slide.addText(`Generated on: ${new Date().toLocaleDateString()}`, { x: 1, y: 1.5, fontSize: 14, color: "6b7280" });

      let yPos = 2.5;
      enterprisePriorities.forEach((p) => {
        if (yPos < 6) {
          slide.addText(`${p.name} - ${p.initiatives.length} Initiatives`, { x: 1, y: yPos, fontSize: 12 });
          yPos += 0.4;
        }
      });

      pres.writeFile({ fileName: `Portfolio_Report_${new Date().toISOString().split("T")[0]}.pptx` });
    } catch (err) {
      console.error("PPTX generation failed. Ensure pptxgenjs is installed.", err);
      alert("PowerPoint export library not initialized.");
    }
  };

  if (isLoading) {
    return (
      <div
        className="initiative-report"
        style={{
          minHeight: "100vh",
          padding: 24,
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          flexDirection: "column",
          gap: 12,
        }}
      >
        <div
          style={{
            width: 36,
            height: 36,
            border: `3px solid ${C.border}`,
            borderTopColor: C.blue,
            borderRadius: "50%",
            animation: "spin 0.8s linear infinite",
          }}
        />
        <div style={{ fontSize: 13, color: C.textSecondary, fontWeight: 600 }}>Loading portfolio data…</div>
      </div>
    );
  }

  if (error) {
    return (
      <div
        className="initiative-report"
        style={{
          minHeight: "100vh",
          padding: 24,
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          flexDirection: "column",
          gap: 12,
        }}
      >
        <div
          style={{
            padding: "16px 20px",
            background: C.redSoft,
            border: `1px solid ${C.red}`,
            borderRadius: C.radiusHero,
            color: C.red,
            fontSize: 13,
            fontWeight: 600,
            maxWidth: 400,
          }}
        >
          {error.message}
        </div>
      </div>
    );
  }

  return (
    <div
      className="initiative-report"
      style={{
        minHeight: "100vh",
        padding: "24px 28px",
      }}
    >
      <div
        className="global-app-header"
        style={{
          background: C.walmartBentonvilleBlue,
          borderBottom: "none",
          padding: "16px 24px",
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          gap: "12px",
          margin: "-24px -28px 24px -28px",
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          {/* Official Walmart Spark — WMT-Spark-SparkYellow-RGB.svg */}
          <svg viewBox="0 0 532.262 600" width={22} height={22} xmlns="http://www.w3.org/2000/svg" style={{ flexShrink: 0 }} aria-hidden>
            <path fill="#FFC220" d="M375.663,273.363c12.505-2.575,123.146-53.269,133.021-58.97c22.547-13.017,30.271-41.847,17.254-64.393s-41.847-30.271-64.393-17.254c-9.876,5.702-109.099,76.172-117.581,85.715c-9.721,10.937-11.402,26.579-4.211,39.033C346.945,269.949,361.331,276.314,375.663,273.363z"/>
            <path fill="#FFC220" d="M508.685,385.607c-9.876-5.702-120.516-56.396-133.021-58.97c-14.332-2.951-28.719,3.415-35.909,15.87c-7.191,12.455-5.51,28.097,4.211,39.033c8.482,9.542,107.705,80.013,117.581,85.715c22.546,13.017,51.376,5.292,64.393-17.254S531.231,398.624,508.685,385.607z"/>
            <path fill="#FFC220" d="M266.131,385.012c-14.382,0-27.088,9.276-31.698,23.164c-4.023,12.117-15.441,133.282-15.441,144.685c0,26.034,21.105,47.139,47.139,47.139c26.034,0,47.139-21.105,47.139-47.139c0-11.403-11.418-132.568-15.441-144.685C293.219,394.288,280.513,385.012,266.131,385.012z"/>
            <path fill="#FFC220" d="M156.599,326.637c-12.505,2.575-123.146,53.269-133.021,58.97C1.031,398.624-6.694,427.454,6.323,450c13.017,22.546,41.847,30.271,64.393,17.254c9.876-5.702,109.098-76.172,117.58-85.715c9.722-10.937,11.402-26.579,4.211-39.033S170.931,323.686,156.599,326.637z"/>
            <path fill="#FFC220" d="M70.717,132.746C48.171,119.729,19.341,127.454,6.323,150c-13.017,22.546-5.292,51.376,17.254,64.393c9.876,5.702,120.517,56.396,133.021,58.97c14.332,2.951,28.719-3.415,35.91-15.87c7.191-12.455,5.51-28.096-4.211-39.033C179.815,208.918,80.592,138.447,70.717,132.746z"/>
            <path fill="#FFC220" d="M266.131,0c-26.035,0-47.139,21.105-47.139,47.139c0,11.403,11.418,132.568,15.441,144.685c4.611,13.888,17.317,23.164,31.698,23.164s27.088-9.276,31.698-23.164c4.023-12.117,15.441-133.282,15.441-144.685C313.27,21.105,292.165,0,266.131,0z"/>
          </svg>
          {/* Official Walmart Wordmark — white on dark bg */}
          <svg viewBox="0 0 200.818 36.441" height={16} xmlns="http://www.w3.org/2000/svg" style={{ flexShrink: 0 }} aria-label="walmart">
            <polygon fill="#ffffff" points="38.104,0 33.448,23.328 28.222,0 19.147,0 13.921,23.328 9.265,0 0,0 7.554,35.634 18.339,35.634 23.613,11.973 28.887,35.634 39.435,35.634 46.941,0"/>
            <path fill="#ffffff" d="M59.698,6.557c-5.749,0-9.787,1.948-11.26,3.326v7.602c1.71-1.52,5.321-3.753,10.072-3.753c2.946,0,4.038,0.808,4.038,2.471c0,1.425-1.52,1.995-5.749,2.898c-6.414,1.33-10.12,3.658-10.12,9.217c0,5.131,3.373,8.124,8.267,8.124c4.1,0,6.546-1.904,7.887-4.461v3.653h8.837V18.149C71.671,10.12,67.49,6.557,59.698,6.557z M58.463,30.692c-2.09,0-3.231-1.283-3.231-3.041c0-2.281,1.805-3.183,4.086-3.991c1.189-0.446,2.378-0.911,3.183-1.615v3.943C62.501,28.982,60.934,30.692,58.463,30.692z"/>
            <rect x="76.09" fill="#ffffff" width="9.17" height="35.634"/>
            <path fill="#ffffff" d="M123.174,6.652c-4.456,0-7.331,2.683-8.833,6.252c-0.803-3.822-3.478-6.252-7.226-6.252c-4.243,0-7.009,2.475-8.41,5.953V7.459h-9.027v28.174h9.17V20.002c0-3.848,1.283-6.034,4.038-6.034c2.233,0,2.993,1.52,2.993,3.896v17.769h9.17V20.002c0-3.848,1.283-6.034,4.038-6.034c2.233,0,2.993,1.52,2.993,3.896v17.769h9.17V16.582C131.251,10.643,128.448,6.652,123.174,6.652z"/>
            <path fill="#ffffff" d="M147.31,6.557c-5.749,0-9.787,1.948-11.26,3.326v7.602c1.71-1.52,5.321-3.753,10.072-3.753c2.946,0,4.038,0.808,4.038,2.471c0,1.425-1.52,1.995-5.749,2.898c-6.414,1.33-10.12,3.658-10.12,9.217c0,5.131,3.373,8.124,8.267,8.124c4.1,0,6.546-1.904,7.887-4.461v3.653h8.837V18.149C159.283,10.12,155.102,6.557,147.31,6.557z M146.074,30.692c-2.091,0-3.231-1.283-3.231-3.041c0-2.281,1.805-3.183,4.086-3.991c1.189-0.446,2.378-0.911,3.183-1.615v3.943C150.113,28.982,148.545,30.692,146.074,30.692z"/>
            <path fill="#ffffff" d="M172.728,15.379v-7.92h-9.027v28.174h9.217V23.661c0-5.511,3.611-6.984,6.889-6.984c1.093,0,2.138,0.143,2.613,0.285V7.269C177.062,7.021,173.931,10.299,172.728,15.379z"/>
            <path fill="#ffffff" d="M200.818,14.586V7.459h-5.796V1.995h-9.17v24.231c0,6.794,3.801,9.977,9.93,9.977c2.851,0,4.371-0.57,5.036-0.998v-7.079c-0.523,0.38-1.378,0.665-2.471,0.665c-1.995,0.048-3.326-0.855-3.326-3.848V14.586H200.818z"/>
          </svg>
          <div style={{ width: 1, height: 18, background: "rgba(255,255,255,0.3)", flexShrink: 0 }} />
          <div style={{ fontSize: 15, fontWeight: 700, color: "#ffffff", letterSpacing: "-0.01em", fontFamily: "'EverydaySans', 'Inter', sans-serif" }}>Product Hub</div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12, color: "rgba(255,255,255,0.7)" }}>
          <StatusDot color={C.green} size={7} />
          Live · {fy === "All" ? "All Years" : fy}
        </div>
      </div>

      <div style={{ marginBottom: 16 }}>
        <Summary
          priorities={summaryPriorities}
          seg={seg}
          fy={fy}
          market={selectedMarkets.length === 0 ? "All Markets" : selectedMarkets.join(", ")}
          globalOnTrack={globalOnTrack}
          globalAtRisk={globalAtRisk}
          globalOffTrack={globalOffTrack}
          statusFilter={statusFilter}
          onStatusFilter={(s) => setStatusFilter((prev) => (prev === s ? null : s))}
        />
      </div>
      {(() => {
        const attentionCount = globalAtRisk + globalOffTrack;
        if (attentionCount <= 0) return null;
        return (
          <div style={{ marginBottom: "24px" }}>
            <div
              role="alert"
              onClick={() => setIsBannerExpanded(!isBannerExpanded)}
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                width: "100%",
                padding: "16px 24px",
                background: C.amberSoft,
                border: `1px solid ${C.amberDark}`,
                borderLeftWidth: "4px",
                borderLeftColor: C.amberDark,
                borderRadius: isBannerExpanded ? `${C.radiusCard}px ${C.radiusCard}px 0 0` : C.radiusCard,
                cursor: "pointer",
              }}
            >
              <span style={{ fontWeight: 600, color: C.amberDark, fontSize: "13px", display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ display: "inline-block", width: 8, height: 8, borderRadius: "50%", background: C.amber, flexShrink: 0 }} />
                {attentionCount} {attentionCount === 1 ? 'item needs' : 'items need'} attention
              </span>
              <span style={{ fontSize: "13px", fontWeight: 600, color: C.amberDark, display: "flex", alignItems: "center", gap: 4 }}>
                View Details {isBannerExpanded ? <IconChevronUp size={12} color="currentColor" /> : <IconChevronDown size={12} color="currentColor" />}
              </span>
            </div>
            {isBannerExpanded && attentionItems.length > 0 && (
              <div
                style={{
                  backgroundColor: C.bgPanel,
                  border: `1px solid ${C.border}`,
                  borderTop: "none",
                  borderBottomLeftRadius: C.radiusCard,
                  borderBottomRightRadius: C.radiusCard,
                  padding: 0,
                  maxHeight: "320px",
                  overflowY: "auto",
                  boxShadow: "0 4px 12px rgba(0,0,0,0.05)",
                }}
              >
                <table style={{ width: "100%", borderCollapse: "collapse", textAlign: "left" }}>
                  <tbody>
                    {attentionItems.map((item, index) => {
                      const priorityForItem = summaryPriorities.find((p) =>
                        p.initiatives.some((i) => i.id === item.id)
                      );
                      return (
                      <tr
                        key={item.id}
                        className="attention-table-row"
                        onClick={() => {
                          if (priorityForItem) {
                            setSelectedPriority(priorityForItem);
                            setSideSheetInitiative(item);
                            setIsBannerExpanded(false);
                          }
                        }}
                        style={{
                          borderBottom:
                            index === attentionItems.length - 1
                              ? "none"
                              : `1px solid ${C.borderLight}`,
                          cursor: "pointer",
                        }}
                      >
                        <td
                          style={{
                            padding: "12px 24px",
                            fontSize: 13,
                            fontWeight: 600,
                            color: C.textPrimary,
                            width: "50%",
                          }}
                        >
                          {item.name}
                        </td>
                        <td
                          style={{
                            padding: "12px 24px",
                            fontSize: "12px",
                            color: C.textSecondary,
                            width: "30%",
                          }}
                        >
                          {item.goalAlignment || "Uncategorized"}
                        </td>
                        <td
                          style={{
                            padding: "12px 24px",
                            width: "20%",
                            textAlign: "right",
                          }}
                        >
                          <Badge status={item.trueStatus} />
                        </td>
                      </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        );
      })()}
      <div className="global-filters" style={{ marginBottom: 16 }}>
        <Filters
          seg={seg}
          onSeg={(s) => {
            setSeg(s);
            setSelectedPriority(null);
          }}
          fy={fy}
          onFy={(f) => {
            setFy(f);
            setSelectedPriority(null);
          }}
          selectedMarkets={selectedMarkets}
          onMarketsChange={(markets) => {
            setSelectedMarkets(markets);
            setSelectedPriority(null);
          }}
          fyOptions={fyOptions}
          marketOptions={marketOptions}
          view={view}
          onView={(v) => {
            setView(v);
            setSelectedPriority(null);
          }}
          onExportCSV={handleExportCSV}
          onExportPDF={handleExportPDF}
          onExportPPT={handleExportPPT}
        />
      </div>

      {view === "priority" && (
        <>
          <div className="priority-section">
            <h2 className="section-title" style={{ fontSize: 15, fontWeight: 600, margin: "24px 0 16px" }}>
              Enterprise Priorities
            </h2>
            <div className="card-grid">
              {[...filteredEnterprisePriorities]
                .sort((a, b) => b.initiatives.length - a.initiatives.length)
                .map((p) => (
                <AggCard
                  key={p.name}
                  title={p.name}
                  priorities={[p]}
                  initCount={p.count}
                  capCount={p.capabilitiesTotal}
                  onTrack={p.onTrack}
                  atRisk={p.atRisk}
                  offTrack={p.offTrack}
                  segBreakdown={p.segments}
                  scopedInits={p.initiatives}
                  onOpenSheet={(priority) => {
                    if (selectedPriority?.name === priority.name) {
                      setSelectedPriority(null);
                      setSideSheetInitiative(null);
                      setSideSheetStatKey(null);
                    } else {
                      openSheet(priority, undefined, "ENTERPRISE PRIORITY");
                    }
                  }}
                />
              ))}
            </div>
          </div>
          <div className="priority-section">
            <h2 className="section-title" style={{ fontSize: 15, fontWeight: 600, margin: "32px 0 16px" }}>
              Other Priorities
            </h2>
            <div className="card-grid">
              {[...filteredOtherPriorities]
                .sort((a, b) => b.initiatives.length - a.initiatives.length)
                .map((p) => (
                <AggCard
                  key={p.name}
                  title={p.name}
                  priorities={[p]}
                  initCount={p.count}
                  capCount={p.capabilitiesTotal}
                  onTrack={p.onTrack}
                  atRisk={p.atRisk}
                  offTrack={p.offTrack}
                  segBreakdown={p.segments}
                  scopedInits={p.initiatives}
                  onOpenSheet={(priority) => {
                    if (selectedPriority?.name === priority.name) {
                      setSelectedPriority(null);
                      setSideSheetInitiative(null);
                      setSideSheetStatKey(null);
                    } else {
                      openSheet(priority, undefined, "OTHER PRIORITY");
                    }
                  }}
                />
              ))}
            </div>
          </div>
        </>
      )}
      {view === "gpa" && (
        <GPAView
          initiatives={filteredInitiatives}
          onOpenSheet={(p) => openSheet(p, undefined, "GPA")}
        />
      )}
      {view === "segment" && (
        <SegmentView
          initiatives={filteredInitiatives}
          onOpenSheet={(p) => openSheet(p, undefined, "SEGMENT")}
        />
      )}

      {selectedPriority && (
        <SideSheet
          priority={selectedPriority}
          initiative={sideSheetInitiative}
          statKey={sideSheetStatKey}
          contextLabel={sideSheetContextLabel || undefined}
          onClose={() => {
            setSelectedPriority(null);
            setSideSheetInitiative(null);
            setSideSheetStatKey(null);
            setSideSheetContextLabel("");
          }}
          onSelectInitiative={setSideSheetInitiative}
          onStatKey={setSideSheetStatKey}
        />
      )}

      <div
        style={{
          marginTop: 22,
          display: "flex",
          gap: 16,
          alignItems: "center",
          justifyContent: "flex-end",
          flexWrap: "wrap",
        }}
      >
        <span style={{ fontSize: 11, color: C.textSecondary }}>Legend:</span>
        {(["on-track", "at-risk", "off-track"] as const).map((s) => (
          <span
            key={s}
            style={{
              display: "flex",
              alignItems: "center",
              gap: 4,
              fontSize: 11,
              color: STATUS[s].color,
              fontWeight: 600,
            }}
          >
            <Dot status={s} size={7} />
            {STATUS[s].label}
          </span>
        ))}
        <span style={{ fontSize: 11, color: C.textSecondary, marginLeft: 8 }}>PM · AI Status</span>
        <span style={{ fontSize: 11, color: C.blue, fontWeight: 600, display: "inline-flex", alignItems: "center", gap: 4 }}>
          <IconExternalLink size={10} color="currentColor" /> Click a card to open side sheet
        </span>
      </div>
    </div>
  );
}

initializeBlock({ interface: () => <App /> });
