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

// --- Constants ---
const PALETTE: readonly string[] = [
  "#90D6F9", "#AACF9A", "#ACC8FB", "#A3DBE9", "#FFEDBC", "#D0C2D8",
  "#FED1B3", "#F4BDD3", "#F9BDB8", "#FFFBB3",
];

const STATUS: Record<InitiativeStatus, StatusConfig> = {
  "on-track": {
    color: "var(--color-status-success-text)",
    bg: "var(--color-status-success-bg)",
    dot: "var(--color-status-success)",
    label: "On Track",
  },
  "at-risk": {
    color: "var(--color-status-warning-text)",
    bg: "var(--color-status-warning-bg)",
    dot: "var(--color-status-warning)",
    label: "At Risk",
  },
  "off-track": {
    color: "var(--color-status-danger-text)",
    bg: "var(--color-status-danger-bg)",
    dot: "var(--color-status-danger)",
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

function getPriorityColor(priorityName: string): string {
  return PALETTE[hashToIndex(priorityName, PALETTE.length)];
}

function getGPAColor(gpa: string): string {
  return PALETTE[hashToIndex(gpa, PALETTE.length)];
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
        background: STATUS[status]?.dot ?? "var(--color-border)",
        flexShrink: 0,
      }}
    />
  );
}

function Badge({ status }: { status: InitiativeStatus }) {
  const s = STATUS[status] ?? STATUS["on-track"];
  const variant =
    status === "on-track" ? "badge-success" : status === "at-risk" ? "badge-warning" : "badge-danger";
  return <span className={`ld-badge ${variant}`}>{s.label}</span>;
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
        height: "8px",
        borderRadius: "4px",
        overflow: "hidden",
        backgroundColor: "var(--color-bg-app)",
      }}
    >
      {totalStatus === 0 ? (
        <div
          style={{ width: "100%", backgroundColor: "var(--color-border)" }}
          title="No status data available"
        />
      ) : (
        <>
          <div style={{ width: `${greenPct}%`, backgroundColor: "var(--color-status-success)" }} />
          <div style={{ width: `${yellowPct}%`, backgroundColor: "var(--color-status-warning)" }} />
          <div style={{ width: `${redPct}%`, backgroundColor: "var(--color-status-danger)" }} />
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

/** Single pill for SideSheet capability cards: "🔗 N GPA Dependencies (High Complexity)" with dynamic color. */
function DependencyPill({ depCount, depScore }: { depCount: number; depScore: number }) {
  const complexity = getComplexityLabel(depScore);
  const variant =
    complexity === "High" ? "badge-danger" : complexity === "Medium" ? "badge-warning" : "badge-neutral";
  return (
    <span className={`ld-badge ${variant}`}>
      🔗 {depCount} GPA Dependencies ({complexity} Complexity)
    </span>
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
        <span className="dual-status__card-warning" title={`AI recommends: ${aiLabel}`}>
          ⚠️ AI flags as {aiLabel}
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
  fontSize: 9,
  fontWeight: 600,
  color: "var(--color-text-muted)",
  textTransform: "uppercase",
  letterSpacing: "0.03em",
};
const statRowHeaderLabelStyle: React.CSSProperties = {
  fontSize: 10,
  fontWeight: 600,
  color: "var(--color-text-muted)",
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
    borderRight: "1px solid var(--color-border-light)",
    padding: useExtendedLayout ? "8px 10px" : "10px 18px",
    minWidth: 0,
  };
  const statBoxLast = useExtendedLayout
    ? { padding: "8px 10px", minWidth: 0 }
    : { padding: "10px 18px", minWidth: 0 };
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
          borderBottom: "1px solid var(--color-border-light)",
          background: "var(--color-bg-card)",
        }}
      >
        <div className="stat-box" style={statBoxBase}>
          <div className="stat-label" style={headerStyle}>Groups</div>
          <div className="stat-value-row" style={{ fontSize: 20, fontWeight: 700, color: "var(--color-text-main)" }}>{data.totalGroups}</div>
        </div>
        <div className="stat-box" style={statBoxBase}>
          <div className="stat-label" style={headerStyle}>Initiatives</div>
          <div className="stat-value-row" style={{ fontSize: 20, fontWeight: 700, color: "var(--color-text-main)" }}>{data.count ?? 0}</div>
        </div>
        <div className="stat-box" style={statBoxBase}>
          <div className="stat-label" style={headerStyle}>Projects</div>
          <div className="stat-value-row" style={{ fontSize: 20, fontWeight: 700, color: "var(--color-text-main)" }}>{data.totalProjects}</div>
        </div>
        <div className="stat-box" style={statBoxBase}>
          <div className="stat-label" style={headerStyle}>Capabilities</div>
          <div className="stat-value-row" style={{ fontSize: 20, fontWeight: 700, color: "var(--color-text-main)" }}>{data.capabilitiesTotal ?? 0}</div>
        </div>
        <div onClick={() => onStat?.("on-track")} className="stat-box" style={dotCol}>
          <span style={{ color: "var(--color-status-success)", fontSize: 14 }}>●</span>
          <span style={{ fontSize: 20, fontWeight: 700, color: "var(--color-text-main)" }}>{onTrackCount}</span>
        </div>
        <div onClick={() => onStat?.("at-risk")} className="stat-box" style={dotCol}>
          <span style={{ color: "var(--color-status-warning)", fontSize: 14 }}>●</span>
          <span style={{ fontSize: 20, fontWeight: 700, color: "var(--color-text-main)" }}>{atRiskCount}</span>
        </div>
        <div
          onClick={() => onStat?.("off-track")}
          className="stat-box"
          style={{ ...statBoxLast, ...dotCol, borderRight: "none" }}
        >
          <span style={{ color: "var(--color-status-danger)", fontSize: 14 }}>●</span>
          <span style={{ fontSize: 20, fontWeight: 700, color: "var(--color-text-main)" }}>{offTrackCount}</span>
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
        borderBottom: "1px solid var(--color-border-light)",
        background: "var(--color-border-light)",
      }}
    >
      <div className="stat-box" style={statBoxBase}>
        <div className="stat-label" style={headerStyle}>Initiatives</div>
        <div className="stat-value-row" style={{ fontSize: 24, fontWeight: 700, color: "var(--color-text-main)" }}>{data.count ?? 0}</div>
      </div>
      <div className="stat-box" style={statBoxBase}>
        <div className="stat-label" style={headerStyle}>Capabilities</div>
        <div className="stat-value-row" style={{ fontSize: 24, fontWeight: 700, color: "var(--color-text-main)" }}>{data.capabilitiesTotal ?? 0}</div>
      </div>
      <div onClick={() => onStat?.("on-track")} className="stat-box" style={dotCol}>
        <span style={{ color: "var(--color-status-success)", fontSize: 14 }}>●</span>
        <span style={{ fontSize: 24, fontWeight: 700, color: "var(--color-text-main)" }}>{onTrackCount}</span>
      </div>
      <div onClick={() => onStat?.("at-risk")} className="stat-box" style={dotCol}>
        <span style={{ color: "var(--color-status-warning)", fontSize: 14 }}>●</span>
        <span style={{ fontSize: 24, fontWeight: 700, color: "var(--color-text-main)" }}>{atRiskCount}</span>
      </div>
      <div onClick={() => onStat?.("off-track")} className="stat-box" style={statBoxLast}>
        <span style={{ color: "var(--color-status-danger)", fontSize: 14 }}>●</span>
        <span style={{ fontSize: 24, fontWeight: 700, color: "var(--color-text-main)" }}>{offTrackCount}</span>
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
  let headerBg = "#338a17";
  let overallStatus = "On Track";
  let statusColor = "var(--color-status-success-text)";
  if (offTrackCount > 0) {
    headerBg = "#ba1e45";
    overallStatus = "Off Track";
    statusColor = "var(--color-status-danger-text)";
  } else if (atRiskCount > 0) {
    headerBg = "#b87503";
    overallStatus = "At Risk";
    statusColor = "var(--color-status-warning-text)";
  }
  const cleanTitle = priority.name.replace(/^Enterprise Priority:\s*/i, "");

  // Detail view: single initiative (Initiative -> Projects -> Capabilities)
  if (initiative) {
    const isMissing = initiative.financeProjection === "MISSING" || !initiative.financeProjection.trim();
    let detailHeaderBg = "#338a17";
    const detailStatus = (initiative.trueStatus ?? "").toLowerCase();
    if (detailStatus === "off-track" || detailStatus === "red") {
      detailHeaderBg = "#ba1e45";
    } else if (detailStatus === "at-risk" || detailStatus === "yellow") {
      detailHeaderBg = "#b87503";
    }
    return (
      <div
        className="side-sheet-overlay side-sheet-open"
        onClick={onClose}
        role="presentation"
      >
        <div className="side-sheet-panel" onClick={(e) => e.stopPropagation()}>
          <header
            className="side-sheet-header"
            style={{
              background: "linear-gradient(135deg, #1a1f2e 0%, #0f1623 100%)",
              color: "#fff",
            }}
          >
            <div className="side-sheet-header__title-group">
              <button
                type="button"
                onClick={() => onSelectInitiative(null)}
                className="side-sheet-back-link"
                style={{
                  background: "none",
                  border: "none",
                  fontSize: 11,
                  cursor: "pointer",
                  padding: 0,
                  marginBottom: 8,
                  textAlign: "left",
                  color: "rgba(255,255,255,0.5)",
                  fontWeight: 500,
                  letterSpacing: "0.02em",
                  fontFamily: "inherit",
                }}
              >
                ← Back to Priority
              </button>
              <div className="side-sheet-header__title">{initiative.name}</div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginTop: 8 }}>
                {/* Status badge */}
                <span style={{
                  fontSize: 10, fontWeight: 700, padding: "2px 8px", borderRadius: 4,
                  background: detailHeaderBg + "30", color: detailHeaderBg,
                  border: `1px solid ${detailHeaderBg}50`,
                  textTransform: "uppercase", letterSpacing: "0.08em",
                }}>
                  ● {initiative.trueStatus?.replace(/-/g, " ")}
                </span>
                {[initiative.gpa, ...initiative.segments].filter(Boolean).map((g) => (
                  <span
                    key={g}
                    style={{
                      fontSize: 11,
                      background: "rgba(255,255,255,0.08)",
                      color: "rgba(255,255,255,0.7)",
                      padding: "2px 8px",
                      borderRadius: 4,
                      fontWeight: 500,
                      border: "1px solid rgba(255,255,255,0.12)",
                    }}
                  >
                    {g}
                  </span>
                ))}
              </div>
            </div>
            <button
              type="button"
              className="side-sheet-header__close"
              onClick={onClose}
              aria-label="Close"
            >
              ×
            </button>
          </header>
          {contextLabel != null && contextLabel !== "" && (
            <div className="side-sheet-context">
              <span className="context-label">{contextLabel}</span>
              <h2 className="context-title">{contextTitle}</h2>
            </div>
          )}
          <div className={`side-sheet-finance-banner ${isMissing ? "side-sheet-finance-banner--missing" : ""}`}>
            <svg className="side-sheet-finance-banner__icon" viewBox="0 0 16 16" fill="currentColor" aria-hidden style={{ width: 14, height: 14, flexShrink: 0 }}>
              <path d="M8 1.333A6.667 6.667 0 1 0 8 14.667 6.667 6.667 0 0 0 8 1.333zm.667 10H7.333V7.333h1.334V11.333zm0-5.333H7.333V4.667h1.334V6z"/>
            </svg>
            <span>Target GMV/MRR: {initiative.financeProjection?.trim() || "MISSING"}</span>
          </div>
          <div className="side-sheet-body">
            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(3, 1fr)",
                gap: "16px",
                padding: "16px 24px",
                borderBottom: "1px solid var(--color-border)",
                marginBottom: 8,
              }}
            >
              {[
                { label: "Product Lead", value: initiative.productLead?.trim() || "Unassigned", color: "var(--color-text-main)" as string | undefined },
                { label: "Status", value: initiative.trueStatus?.replace(/-/g, " ") || "Unknown", color: detailHeaderBg },
                { label: "Risks", value: initiative.risks?.trim() || "None", color: initiative.risks?.trim() ? "var(--color-status-danger-text)" : "var(--color-text-muted)" },
              ].map((item) => (
                <div key={item.label} style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                  <span style={{ fontSize: 10, textTransform: "uppercase", color: "var(--color-text-muted)", fontWeight: 600, letterSpacing: "0.06em" }}>
                    {item.label}
                  </span>
                  <span style={{ fontSize: 13, fontWeight: 600, color: item.color ?? "var(--color-text-main)", textTransform: item.label === "Status" ? ("capitalize" as const) : ("none" as const) }}>
                    {item.value}
                  </span>
                </div>
              ))}
            </div>
            <div className="hierarchy-container" style={{ display: "flex", flexDirection: "column", gap: "24px", marginTop: "24px" }}>
              {initiative.projects?.map((project) => (
                <div
                  key={project.id}
                  className="project-card"
                  style={{
                    backgroundColor: "var(--color-bg-card)",
                    border: "1px solid var(--color-border)",
                    borderRadius: "8px",
                    padding: "16px",
                  }}
                >
                  {/* Level 2: Project Header */}
                  <div
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: "12px",
                      marginBottom: "16px",
                      paddingBottom: "12px",
                      borderBottom: "1px solid var(--color-border-subtle)",
                    }}
                  >
                    <span
                      style={{
                        fontSize: "11px",
                        fontWeight: 700,
                        color: "var(--color-text-muted)",
                        textTransform: "uppercase",
                        letterSpacing: "0.5px",
                      }}
                    >
                      PROJECT
                    </span>
                    <h3 style={{ margin: 0, fontSize: "16px", color: "var(--color-text-main)", fontWeight: 600 }}>
                      {project.name}
                    </h3>
                    <span className={`ld-badge badge-${project.status?.toLowerCase().replace(/\s+/g, "-") || "neutral"}`}>
                      {project.status || "Unassigned"}
                    </span>
                  </div>
                  {/* Level 3: Capabilities List */}
                  <div
                    className="capabilities-list"
                    style={{
                      display: "flex",
                      flexDirection: "column",
                      gap: "16px",
                      paddingLeft: "12px",
                      borderLeft: "3px solid var(--color-bg-app)",
                    }}
                  >
                    {project.capabilities && project.capabilities.length > 0 ? (
                      project.capabilities.map((cap) => (
                        <div
                          key={cap.id}
                          className="capability-card"
                          style={{ backgroundColor: "var(--color-bg-app)", padding: "12px", borderRadius: "6px" }}
                        >
                          {/* Capability Header & Dual RAG */}
                          <div style={{ display: "flex", alignItems: "center", gap: "12px", marginBottom: "8px" }}>
                            <h4 style={{ margin: 0, fontSize: "14px", color: "var(--color-text-main)", fontWeight: 600 }}>
                              {cap.name}
                            </h4>
                            <span className={`ld-badge badge-${cap.status?.toLowerCase().replace(/\s+/g, "-") || "neutral"}`}>
                              {cap.status || "Unassigned"}
                            </span>
                            {cap.aiStatus && cap.aiStatus !== cap.status && (
                              <span className="ld-badge badge-ai" title="AI detects a discrepancy">
                                ✦ AI flags as {cap.aiStatus}
                              </span>
                            )}
                          </div>
                          <div className="capability-card__main">
                            {(cap.startQ && cap.launchQ) ? (
                              <div className="capability-card__meta">{cap.startQ} to {cap.launchQ}</div>
                            ) : null}
                            <div className="capability-card__dependency">
                              <DependencyPill depCount={cap.depCount} depScore={cap.depScore} />
                            </div>
                          </div>
                          <div className="capability-card__actions">
                            <span className="ld-badge badge-neutral">{cap.size}</span>
                            <div
                              style={{
                                fontSize: "12px",
                                fontWeight: 600,
                                color: "var(--color-brand-primary)",
                                cursor: "pointer",
                                display: "flex",
                                alignItems: "center",
                                gap: "4px",
                              }}
                              onClick={() => alert(`Navigating to Roadmap for ${cap.name}`)}
                            >
                              View Roadmap <span style={{ fontSize: "14px" }}>↗</span>
                            </div>
                          </div>
                          {/* Status Notes */}
                          {cap.statusNotes && (
                            <div
                              className="status-notes-box"
                              style={{
                                marginTop: "8px",
                                padding: "10px",
                                backgroundColor: "var(--color-bg-card)",
                                borderLeft: `3px solid var(--color-status-${cap.status?.toLowerCase() === "red" ? "danger" : cap.status?.toLowerCase() === "yellow" ? "warning" : "success"})`,
                                fontSize: "12px",
                                color: "var(--color-text-muted)",
                              }}
                            >
                              <strong>Status Notes:</strong> {cap.statusNotes}
                            </div>
                          )}
                          {/* Level 4: Epics */}
                          {cap.epics && cap.epics.length > 0 && (
                            <div
                              className="epic-list-container"
                              style={{
                                marginTop: "12px",
                                paddingLeft: "16px",
                                borderLeft: "2px dashed var(--color-border)",
                              }}
                            >
                              <h5 style={{ fontSize: "11px", textTransform: "uppercase", color: "var(--color-text-muted)", marginBottom: "8px" }}>
                                Epics ({cap.epics.length})
                              </h5>
                              <div style={{ display: "flex", flexDirection: "column", gap: "6px" }}>
                                {cap.epics.map((epic) => (
                                  <div
                                    key={epic.id}
                                    className="epic-row"
                                    style={{
                                      display: "flex",
                                      alignItems: "center",
                                      gap: "8px",
                                      fontSize: "13px",
                                      backgroundColor: "var(--color-bg-card)",
                                      padding: "6px 12px",
                                      borderRadius: "4px",
                                      border: "1px solid var(--color-border-subtle)",
                                    }}
                                  >
                                    <span className={`status-dot ${epic.status?.toLowerCase().replace(/\s+/g, "-")}`}>●</span>
                                    <span style={{ color: "var(--color-text-main)", fontWeight: 500 }}>{epic.name}</span>
                                    <span style={{ marginLeft: "auto", fontSize: "11px", color: "var(--color-text-muted)" }}>{epic.status}</span>
                                  </div>
                                ))}
                              </div>
                            </div>
                          )}
                          {cap.metrics && cap.metrics.length > 0 && (
                            <div className="capability-card__metrics" style={{ marginTop: 12 }}>
                              <div className="capability-card__metrics-title">Metrics Health</div>
                              <div className="metric-grid metric-grid-header">
                                <span className="metric-name capability-metric-name">Metric</span>
                                <span className="capability-metric-label metric-col-right">Baseline</span>
                                <span className="capability-metric-label metric-col-right">Target</span>
                                <span className="capability-metric-label metric-col-right">Actual</span>
                                <span className="capability-metric-label">Status</span>
                              </div>
                              <div className="metric-grid">
                                {cap.metrics.map((metric) => (
                                  <React.Fragment key={metric.id}>
                                    <span className="metric-name capability-metric-name">{metric.name}</span>
                                    <span className="capability-metric-cell metric-col-right">
                                      <span className="capability-metric-value">
                                        {metric.baseline !== null ? metric.baseline.toLocaleString() : "—"}
                                      </span>
                                    </span>
                                    <span className="capability-metric-cell metric-col-right">
                                      <span className="capability-metric-value">
                                        {metric.target !== null ? metric.target.toLocaleString() : "—"}
                                      </span>
                                    </span>
                                    <span className="capability-metric-cell metric-col-right">
                                      <span className="capability-metric-value capability-metric-value--actual">
                                        {metric.actual !== null ? metric.actual.toLocaleString() : "—"}
                                      </span>
                                    </span>
                                    <Badge status={metric.status} />
                                  </React.Fragment>
                                ))}
                              </div>
                            </div>
                          )}
                        </div>
                      ))
                    ) : (
                      <div style={{ fontSize: "12px", color: "var(--color-text-muted)", fontStyle: "italic" }}>
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
                    fontSize: "14px",
                    color: "var(--color-text-muted)",
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
                    gap: "16px",
                    paddingLeft: "12px",
                    borderLeft: "3px solid var(--color-bg-app)",
                  }}
                >
                  {initiative.orphanedCapabilities.map((cap) => (
                    <div
                      key={cap.id}
                      className="capability-card"
                      style={{ backgroundColor: "var(--color-bg-app)", padding: "12px", borderRadius: "6px" }}
                    >
                      <div style={{ display: "flex", alignItems: "center", gap: "12px", marginBottom: "8px" }}>
                        <h4 style={{ margin: 0, fontSize: "14px", color: "var(--color-text-main)", fontWeight: 600 }}>
                          {cap.name}
                        </h4>
                        <span className={`ld-badge badge-${cap.status?.toLowerCase().replace(/\s+/g, "-") || "neutral"}`}>
                          {cap.status || "Unassigned"}
                        </span>
                        {cap.aiStatus && cap.aiStatus !== cap.status && (
                          <span className="ld-badge badge-ai" title="AI detects a discrepancy">
                            ✦ AI flags as {cap.aiStatus}
                          </span>
                        )}
                      </div>
                      <div className="capability-card__main">
                        {(cap.startQ && cap.launchQ) ? (
                          <div className="capability-card__meta">{cap.startQ} to {cap.launchQ}</div>
                        ) : null}
                        <div className="capability-card__dependency">
                          <DependencyPill depCount={cap.depCount} depScore={cap.depScore} />
                        </div>
                      </div>
                      <div className="capability-card__actions">
                        <span className="ld-badge badge-neutral">{cap.size}</span>
                        <button
                          type="button"
                          className="btn btn-outline"
                          onClick={() => alert(`Navigating to Roadmap for ${cap.name}`)}
                        >
                          View on Roadmap
                        </button>
                      </div>
                      {cap.statusNotes && (
                        <div
                          className="status-notes-box"
                          style={{
                            marginTop: "8px",
                            padding: "10px",
                            backgroundColor: "var(--color-bg-card)",
                            borderLeft: `3px solid var(--color-status-${cap.status?.toLowerCase() === "red" ? "danger" : cap.status?.toLowerCase() === "yellow" ? "warning" : "success"})`,
                            fontSize: "12px",
                            color: "var(--color-text-muted)",
                          }}
                        >
                          <strong>Status Notes:</strong> {cap.statusNotes}
                        </div>
                      )}
                      {cap.epics && cap.epics.length > 0 && (
                        <div
                          className="epic-list-container"
                          style={{
                            marginTop: "12px",
                            paddingLeft: "16px",
                            borderLeft: "2px dashed var(--color-border)",
                          }}
                        >
                          <h5 style={{ fontSize: "11px", textTransform: "uppercase", color: "var(--color-text-muted)", marginBottom: "8px" }}>
                            Epics ({cap.epics.length})
                          </h5>
                          <div style={{ display: "flex", flexDirection: "column", gap: "6px" }}>
                            {cap.epics.map((epic) => (
                              <div
                                key={epic.id}
                                className="epic-row"
                                style={{
                                  display: "flex",
                                  alignItems: "center",
                                  gap: "8px",
                                  fontSize: "13px",
                                  backgroundColor: "var(--color-bg-card)",
                                  padding: "6px 12px",
                                  borderRadius: "4px",
                                  border: "1px solid var(--color-border-subtle)",
                                }}
                              >
                                <span className={`status-dot ${epic.status?.toLowerCase().replace(/\s+/g, "-")}`}>●</span>
                                <span style={{ color: "var(--color-text-main)", fontWeight: 500 }}>{epic.name}</span>
                                <span style={{ marginLeft: "auto", fontSize: "11px", color: "var(--color-text-muted)" }}>{epic.status}</span>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}
                      {cap.metrics && cap.metrics.length > 0 && (
                        <div className="capability-card__metrics" style={{ marginTop: 12 }}>
                          <div className="capability-card__metrics-title">Metrics Health</div>
                          <div className="metric-grid metric-grid-header">
                            <span className="metric-name capability-metric-name">Metric</span>
                            <span className="capability-metric-label metric-col-right">Baseline</span>
                            <span className="capability-metric-label metric-col-right">Target</span>
                            <span className="capability-metric-label metric-col-right">Actual</span>
                            <span className="capability-metric-label">Status</span>
                          </div>
                          <div className="metric-grid">
                            {cap.metrics.map((metric) => (
                              <React.Fragment key={metric.id}>
                                <span className="metric-name capability-metric-name">{metric.name}</span>
                                <span className="capability-metric-cell metric-col-right">
                                  <span className="capability-metric-value">
                                    {metric.baseline !== null ? metric.baseline.toLocaleString() : "—"}
                                  </span>
                                </span>
                                <span className="capability-metric-cell metric-col-right">
                                  <span className="capability-metric-value">
                                    {metric.target !== null ? metric.target.toLocaleString() : "—"}
                                  </span>
                                </span>
                                <span className="capability-metric-cell metric-col-right">
                                  <span className="capability-metric-value capability-metric-value--actual">
                                    {metric.actual !== null ? metric.actual.toLocaleString() : "—"}
                                  </span>
                                </span>
                                <Badge status={metric.status} />
                              </React.Fragment>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>
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
          {/* Dark Gradient Header - sophisticated, not childish color flood */}
          <div style={{ background: "linear-gradient(135deg, #1a1f2e 0%, #0f1623 100%)", padding: "24px 28px", color: "#fff", position: "relative" }}>
            <button
              type="button"
              onClick={onClose}
              aria-label="Close"
              style={{
                position: "absolute",
                top: "16px",
                right: "16px",
                background: "rgba(255,255,255,0.08)",
                border: "1px solid rgba(255,255,255,0.12)",
                borderRadius: "6px",
                width: "28px",
                height: "28px",
                color: "rgba(255,255,255,0.6)",
                fontSize: "14px",
                cursor: "pointer",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                transition: "background 0.12s",
              }}
              onMouseEnter={(e) => { e.currentTarget.style.background = "rgba(255,255,255,0.15)"; }}
              onMouseLeave={(e) => { e.currentTarget.style.background = "rgba(255,255,255,0.08)"; }}
            >
              ✕
            </button>
            {/* Status badge in header */}
            <div style={{ marginBottom: 10 }}>
              <span style={{
                fontSize: 10,
                fontWeight: 700,
                padding: "2px 8px",
                borderRadius: 4,
                background: headerBg + "30",
                color: headerBg,
                border: `1px solid ${headerBg}50`,
                textTransform: "uppercase",
                letterSpacing: "0.08em",
              }}>
                ● {overallStatus}
              </span>
            </div>
            <h2 style={{ margin: "0 0 4px 0", fontSize: 20, fontWeight: 700, lineHeight: 1.2, paddingRight: "40px", letterSpacing: "-0.02em" }}>
              {cleanTitle}
            </h2>
            <p style={{ margin: 0, fontSize: 12, color: "rgba(255,255,255,0.45)", fontWeight: 400 }}>
              Executive Priority Drill-Down
            </p>
          </div>
          {/* Stats Grid - Admin Hub StatCard pattern */}
          <div style={{
            display: "grid",
            gridTemplateColumns: "repeat(3, 1fr)",
            borderBottom: "1px solid var(--color-border)",
          }}>
            <div
              onClick={() => setLocalFilter("all")}
              style={{
                padding: "18px 20px",
                cursor: "pointer",
                borderRight: "1px solid var(--color-border)",
                background: localFilter === "all" ? "var(--color-bg-app)" : "var(--color-bg-card)",
                transition: "background 0.12s",
              }}
            >
              <div style={{ fontSize: 11, fontWeight: 600, color: "var(--color-text-tertiary)", textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 8 }}>
                Initiatives
              </div>
              <div style={{ fontSize: 26, fontWeight: 700, color: "var(--color-text-main)", lineHeight: 1 }}>
                {inits.length}
              </div>
              <div style={{ fontSize: 11, color: "var(--color-text-tertiary)", marginTop: 4 }}>
                {localFilter === "all" ? "Showing all" : "Click to show all"}
              </div>
            </div>
            <div
              onClick={() => setLocalFilter("attention")}
              style={{
                padding: "18px 20px",
                cursor: "pointer",
                borderRight: "1px solid var(--color-border)",
                background: localFilter === "attention" ? "var(--color-bg-app)" : "var(--color-bg-card)",
                transition: "background 0.12s",
              }}
            >
              <div style={{ fontSize: 11, fontWeight: 600, color: "var(--color-text-tertiary)", textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 8 }}>
                Needs Attention
              </div>
              <div style={{ fontSize: 26, fontWeight: 700, color: offTrackCount + atRiskCount > 0 ? "var(--color-status-danger-text)" : "var(--color-text-main)", lineHeight: 1 }}>
                {offTrackCount + atRiskCount}
              </div>
              <div style={{ fontSize: 11, color: "var(--color-text-tertiary)", marginTop: 4 }}>
                At risk or off track
              </div>
            </div>
            <div style={{
              padding: "18px 20px",
              background: "var(--color-bg-card)",
            }}>
              <div style={{ fontSize: 11, fontWeight: 600, color: "var(--color-text-tertiary)", textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 8 }}>
                Overall Status
              </div>
              <div style={{ fontSize: 14, fontWeight: 700, lineHeight: 1, marginBottom: 4 }}>
                <span style={{
                  padding: "3px 10px", borderRadius: 4, fontSize: 12,
                  background: headerBg + "18", color: headerBg,
                  border: `1px solid ${headerBg}40`, fontWeight: 600,
                }}>
                  ● {overallStatus}
                </span>
              </div>
              <div style={{ fontSize: 11, color: "var(--color-text-tertiary)", marginTop: 8 }}>
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
                      <div className="ld-body-s" style={{ fontWeight: 600, fontSize: 13, color: "var(--color-text-main)", lineHeight: 1.4 }}>{init.name}</div>
                      <div style={{ display: "flex", alignItems: "center", gap: 6, flexShrink: 0 }}>
                        <Badge status={init.trueStatus} />
                        {hasAiDiscrepancy && (
                          <span className="ld-badge badge-ai" title={`AI flags as ${init.aiStatus}`}>✦ AI</span>
                        )}
                      </div>
                    </div>
                    {/* Bottom row: meta */}
                    <div style={{ marginTop: 6, display: "flex", alignItems: "center", gap: 12, fontSize: 12, color: "var(--color-text-muted)", flexWrap: "wrap" }}>
                      <span>{init.gpa}</span>
                      {init.segments.length > 0 && <><span style={{ color: "var(--color-border)" }}>·</span><span>{init.segments.join(", ")}</span></>}
                      {init.productLead?.trim() && <><span style={{ color: "var(--color-border)" }}>·</span><span>{init.productLead.trim()}</span></>}
                      {init.rawFinance && <><span style={{ color: "var(--color-border)" }}>·</span><span style={{ fontWeight: 500, color: "var(--color-text-main)" }}>{init.rawFinance}</span></>}
                    </div>
                  </div>
                  <span style={{ fontSize: 16, color: "var(--color-text-muted)", flexShrink: 0 }}>›</span>
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
            <div className="ld-heading-m" style={{ fontSize: 22, fontWeight: 700 }}>{String(v)}</div>
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
  let borderColor = "var(--color-status-success)";
  if (offTrackCount > 0) {
    statusText = `${offTrackCount} OFF TRACK`;
    borderColor = "var(--color-status-danger)";
  } else if (atRiskCount > 0) {
    statusText = `${atRiskCount} AT RISK`;
    borderColor = "var(--color-status-warning)";
  } else if (initiatives.length === 0) {
    statusText = "NO DATA";
    borderColor = "var(--color-border)";
  }

  const cleanTitle = title.replace(/^Enterprise Priority:\s*/i, "");

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
      onMouseEnter={(e) => {
        e.currentTarget.style.transform = "translateY(-2px)";
        e.currentTarget.style.boxShadow = `0 8px 40px rgba(0,0,0,0.11), 0 0 0 4px ${borderColor}18`;
      }}
      onMouseLeave={(e) => {
        e.currentTarget.style.transform = "translateY(0)";
        e.currentTarget.style.boxShadow = "0 2px 8px rgba(0,0,0,0.04)";
      }}
      style={{
        display: "flex",
        flexDirection: "column",
        height: "100%",
        padding: "24px",
        backgroundColor: "var(--color-bg-card)",
        borderRadius: "12px",
        border: "1px solid var(--color-border)",
        borderLeft: `6px solid ${borderColor}`,
        cursor: "pointer",
        boxShadow: "0 2px 8px rgba(0,0,0,0.04)",
        transition: "transform 0.2s ease, box-shadow 0.2s ease",
        opacity: initiatives.length === 0 ? 0.6 : 1,
      }}
    >
      {/* Top: Clean Title (Prefix Removed) */}
      <div
        style={{
          fontSize: "15px",
          fontWeight: 600,
          color: "var(--color-text-main)",
          marginBottom: "16px",
          lineHeight: "1.4",
          minHeight: "42px",
        }}
      >
        {cleanTitle}
      </div>

      {/* Hero Metric: Spread out */}
      <div
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          width: "100%",
          marginBottom: "20px",
        }}
      >
        <div style={{ display: "flex", alignItems: "baseline", gap: "8px" }}>
          <div
            style={{
              fontSize: "28px",
              fontWeight: 700,
              color: "var(--color-text-main)",
              lineHeight: "1",
              letterSpacing: "-1px",
            }}
          >
            {initiatives.length}
          </div>
          <div
            style={{
              fontSize: "12px",
              fontWeight: 700,
              color: "var(--color-text-muted)",
              textTransform: "uppercase",
              letterSpacing: "0.5px",
            }}
          >
            Initiatives
          </div>
        </div>
        {/* Status Badge - tinted by status color */}
        <div
          style={{
            fontSize: 10,
            fontWeight: 700,
            color: borderColor,
            textTransform: "uppercase",
            letterSpacing: "0.07em",
            padding: "3px 8px",
            background: borderColor + "14",
            border: `1px solid ${borderColor}35`,
            borderRadius: 4,
            whiteSpace: "nowrap",
          }}
        >
          ● {statusText}
        </div>
      </div>

      {/* Secondary Metrics: Clean borderless row */}
      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(3, 1fr)",
          width: "100%",
          borderTop: "1px solid var(--color-border-subtle)",
          borderBottom: "1px solid var(--color-border-subtle)",
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
              borderLeft: i > 0 ? "1px solid var(--color-border-subtle)" : "none",
            }}
          >
            <span style={{ fontSize: 10, letterSpacing: "0.06em", textTransform: "uppercase", color: "var(--color-text-muted)", fontWeight: 600 }}>
              {item.label}
            </span>
            <span style={{ fontSize: 15, fontWeight: 700, color: "var(--color-text-main)", marginTop: 4, lineHeight: 1 }}>
              {item.value}
            </span>
          </div>
        ))}
      </div>

      {/* Bottom: Subdued Financial Targets */}
      <div
        style={{
          marginTop: "auto",
          paddingTop: "16px",
          borderTop: "1px solid var(--color-border-subtle)",
          display: "flex",
          justifyContent: "space-between",
          fontSize: "11px",
          fontWeight: 600,
          color: "var(--color-text-muted)",
          textTransform: "uppercase",
          letterSpacing: "0.5px",
        }}
      >
        <span>TARGET: {displayFinance || "NOT SET"}</span>
        <span>ACTUAL: TBD</span>
      </div>
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
    { label: "On Track", value: globalOnTrack, dot: "var(--color-status-success)", status: "on-track" as string | null },
    { label: "At Risk", value: globalAtRisk, dot: "var(--color-status-warning)", status: "at-risk" as string | null },
    { label: "Off Track", value: globalOffTrack, dot: "var(--color-status-danger)", status: "off-track" as string | null },
  ];
  return (
    <div className="summary-banner ld-card" style={{ padding: 0, overflow: "hidden" }}>
      {/* Header row */}
      <div style={{ padding: "14px 24px 12px", borderBottom: "1px solid var(--color-border)", display: "flex", alignItems: "baseline", justifyContent: "space-between" }}>
        <h2 style={{ margin: 0, fontSize: 11, fontWeight: 700, letterSpacing: "0.07em", textTransform: "uppercase", color: "var(--color-text-muted)" }}>
          Portfolio Health
        </h2>
        <p style={{ margin: 0, fontSize: 12, color: "var(--color-text-tertiary)" }}>
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
              borderRight: i < 5 ? "1px solid var(--color-border)" : "none",
              borderLeft: i === 3 ? "1px solid var(--color-border-subtle)" : "none",
              cursor: stat.status ? "pointer" : "default",
              opacity: statusFilter && statusFilter !== stat.status ? 0.4 : 1,
              transition: "opacity 0.15s, background 0.15s",
              background: stat.status && statusFilter === stat.status ? "var(--color-bg-app)" : "transparent",
            }}
          >
            <div style={{ fontSize: 10, fontWeight: 600, color: stat.dot ?? "var(--color-text-muted)", textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 10, display: "flex", alignItems: "center", justifyContent: "center", gap: 5 }}>
              {stat.dot && <span style={{ display: "inline-block", width: 6, height: 6, borderRadius: "50%", background: stat.dot, flexShrink: 0 }} />}
              {stat.label}
            </div>
            <div style={{ fontSize: 26, fontWeight: 700, color: "var(--color-text-main)", lineHeight: 1 }}>
              {stat.value}
            </div>
          </div>
        ))}
      </div>
      {/* Progress bar */}
      <div style={{ height: 4, display: "flex", backgroundColor: "var(--color-border-light)" }}>
        {tot > 0 && (
          <>
            <div style={{ width: `${greenPct}%`, background: "var(--color-status-success)", transition: "width 0.3s" }} />
            <div style={{ width: `${yellowPct}%`, background: "var(--color-status-warning)", transition: "width 0.3s" }} />
            <div style={{ width: `${redPct}%`, background: "var(--color-status-danger)", transition: "width 0.3s" }} />
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
          backgroundColor: "var(--color-bg-card)",
          border: "1px solid var(--color-border)",
          borderRadius: "6px",
          padding: "6px 32px 6px 12px",
          fontSize: "13px",
          fontWeight: 500,
          color: "var(--color-text-main)",
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
              backgroundColor: "var(--color-bg-card)",
              border: "1px solid var(--color-border)",
              borderRadius: "8px",
              boxShadow: "0 4px 16px rgba(0,0,0,0.10), 0 1px 4px rgba(0,0,0,0.06)",
              zIndex: 1000,
              maxHeight: "250px",
              overflowY: "auto",
              overflow: "hidden",
            }}
          >
            <div
              onClick={() => {
                onMarketsChange([]);
              }}
              style={{
                padding: "8px 12px",
                borderBottom: "1px solid var(--color-border-subtle)",
                cursor: "pointer",
                fontSize: "13px",
                fontWeight: selectedMarkets.length === 0 ? 600 : 400,
                backgroundColor: selectedMarkets.length === 0 ? "var(--color-bg-app)" : "var(--color-bg-card)",
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
          borderBottom: "1px solid var(--color-border)",
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
                fontSize: "14px",
                fontWeight: view === v.id ? 700 : 500,
                color: view === v.id ? "var(--color-text-main)" : "var(--color-text-muted)",
                borderBottom: view === v.id ? "3px solid var(--color-brand-primary)" : "3px solid transparent",
                transition: "all 0.2s ease",
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
              borderRadius: "6px",
              padding: "6px 32px 6px 12px",
              color: "#fff",
              boxShadow: "0 1px 4px rgba(45,127,249,0.3)",
              cursor: "pointer",
              fontFamily: "var(--font-family-base)",
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
            backgroundColor: "var(--color-bg-app)",
            padding: "4px",
            borderRadius: "8px",
            border: "1px solid var(--color-border-subtle)",
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
                padding: "6px 16px",
                cursor: "pointer",
                fontSize: "13px",
                fontWeight: seg === s ? 600 : 500,
                color: seg === s ? "var(--color-text-main)" : "var(--color-text-muted)",
                backgroundColor: seg === s ? "#FFFFFF" : "transparent",
                borderRadius: "6px",
                boxShadow: seg === s ? "0 2px 4px rgba(0,0,0,0.06)" : "none",
                transition: "all 0.2s ease",
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
                color: "var(--color-text-muted)",
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
                backgroundColor: "var(--color-bg-card)",
                border: "1px solid var(--color-border)",
                borderRadius: "6px",
                padding: "6px 32px 6px 12px",
                fontSize: "13px",
                fontWeight: 500,
                color: "var(--color-text-main)",
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
                color: "var(--color-text-muted)",
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

  /** Compound filter: Segment AND Fiscal Year AND Market (intersection). */
  const filteredInitiatives = useMemo(
    () =>
      initiatives.filter((init) => {
        const matchSeg = seg === "All" || (init.segments && init.segments.includes(seg));
        const matchFy = fy === "All" || init.fy === fy;
        const matchMarket =
          selectedMarkets.length === 0 ||
          (init.market && selectedMarkets.some((m) => init.market!.includes(m)));
        const matchStatus = !statusFilter || init.trueStatus === statusFilter;
        return matchSeg && matchFy && matchMarket && matchStatus;
      }),
    [initiatives, seg, fy, selectedMarkets, statusFilter]
  );

  const priorities = useMemo(() => buildPriorities(initiatives), [initiatives]);
  const filtered = useMemo(
    () => buildPriorities(filteredInitiatives),
    [filteredInitiatives]
  );

  /** Summary and cards both use the same compound-filtered priorities. */
  const summaryPriorities = filtered;

  /** Global status counts for Portfolio Health legend (dot status). */
  const globalOnTrack = filteredInitiatives.filter((i) => i.trueStatus === "on-track").length;
  const globalAtRisk = filteredInitiatives.filter((i) => i.trueStatus === "at-risk").length;
  const globalOffTrack = filteredInitiatives.filter((i) => i.trueStatus === "off-track").length;

  /** Initiatives that need attention (at-risk or off-track) for the banner expanded list. */
  const attentionItems = useMemo(
    () =>
      filteredInitiatives.filter(
        (i) => i.trueStatus === "at-risk" || i.trueStatus === "off-track"
      ),
    [filteredInitiatives]
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
            border: "3px solid var(--color-border)",
            borderTopColor: "var(--color-brand-primary)",
            borderRadius: "50%",
            animation: "spin 0.8s linear infinite",
          }}
        />
        <div style={{ fontSize: 14, color: "var(--color-text-muted)", fontWeight: 600 }}>Loading portfolio data…</div>
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
            background: "var(--color-status-danger-bg)",
            border: "1px solid var(--color-status-danger)",
            borderRadius: 12,
            color: "var(--color-status-danger)",
            fontSize: 14,
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
        padding: "24px 32px",
      }}
    >
      <div
        className="global-app-header"
        style={{
          background: "linear-gradient(135deg, #1a1f2e 0%, #0f1623 100%)",
          padding: "16px 24px",
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          gap: "12px",
          color: "#fff",
          margin: "-24px -32px 24px -32px",
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
          <svg viewBox="0 0 100 100" width={32} height={32} xmlns="http://www.w3.org/2000/svg" style={{ flexShrink: 0, marginRight: "4px" }} aria-hidden>
            <g fill="#FFC220">
              <rect x="42.5" y="5" width="15" height="28" rx="7.5" />
              <rect x="42.5" y="5" width="15" height="28" rx="7.5" transform="rotate(60 50 50)" />
              <rect x="42.5" y="5" width="15" height="28" rx="7.5" transform="rotate(120 50 50)" />
              <rect x="42.5" y="5" width="15" height="28" rx="7.5" transform="rotate(180 50 50)" />
              <rect x="42.5" y="5" width="15" height="28" rx="7.5" transform="rotate(240 50 50)" />
              <rect x="42.5" y="5" width="15" height="28" rx="7.5" transform="rotate(300 50 50)" />
            </g>
          </svg>
          <div>
            <div style={{ fontSize: 15, fontWeight: 700, color: "#fff", letterSpacing: "-0.01em" }}>Product Hub</div>
            <div style={{ fontSize: 11, color: "rgba(255,255,255,0.5)", fontWeight: 500, marginTop: 1 }}>Executive Strategic Portfolio</div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: "rgba(255,255,255,0.9)" }}>
          <div style={{ width: 7, height: 7, borderRadius: "50%", background: "#FFC220" }} />
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
          onStatusFilter={setStatusFilter}
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
                background: "var(--color-status-warning-bg)",
                border: "1px solid var(--color-status-warning-text)",
                borderLeftWidth: "4px",
                borderLeftColor: "var(--color-status-warning-text)",
                borderRadius: isBannerExpanded ? "8px 8px 0 0" : "8px",
                cursor: "pointer",
              }}
            >
              <span style={{ fontWeight: 600, color: "var(--color-status-warning-text)", fontSize: "14px", display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ display: "inline-block", width: 8, height: 8, borderRadius: "50%", background: "var(--color-status-warning)", flexShrink: 0 }} />
                {attentionCount} items need attention
              </span>
              <span style={{ fontSize: "13px", fontWeight: 600, color: "var(--color-status-warning-text)" }}>
                View Details {isBannerExpanded ? "▲" : "▼"}
              </span>
            </div>
            {isBannerExpanded && attentionItems.length > 0 && (
              <div
                style={{
                  backgroundColor: "var(--color-bg-card)",
                  border: "1px solid var(--color-border)",
                  borderTop: "none",
                  borderBottomLeftRadius: "8px",
                  borderBottomRightRadius: "8px",
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
                              : "1px solid var(--color-border-subtle)",
                          cursor: "pointer",
                        }}
                      >
                        <td
                          style={{
                            padding: "12px 24px",
                            fontSize: "14px",
                            fontWeight: 600,
                            color: "var(--color-text-main)",
                            width: "50%",
                          }}
                        >
                          {item.name}
                        </td>
                        <td
                          style={{
                            padding: "12px 24px",
                            fontSize: "12px",
                            color: "var(--color-text-muted)",
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
            <h2 className="section-title" style={{ fontSize: 18, fontWeight: 600, margin: "24px 0 16px" }}>
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
            <h2 className="section-title" style={{ fontSize: 18, fontWeight: 600, margin: "32px 0 16px" }}>
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
        <span style={{ fontSize: 11, color: "var(--color-text-muted)" }}>Legend:</span>
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
        <span style={{ fontSize: 11, color: "var(--color-text-muted)", marginLeft: 8 }}>PM · AI Status</span>
        <span style={{ fontSize: 11, color: "var(--color-brand-primary)", fontWeight: 600 }}>
          ↗ Click a card to open side sheet
        </span>
      </div>
    </div>
  );
}

initializeBlock({ interface: () => <App /> });
