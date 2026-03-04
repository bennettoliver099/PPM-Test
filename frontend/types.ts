/**
 * Data types for the Executive Initiative Reporting extension.
 * Hierarchy: Initiative -> Project -> Capability.
 * Status unions align with Airtable single-select choices.
 */

/** Initiative/portfolio status (TRUE_STATUS, user-facing) */
export type InitiativeStatus = "on-track" | "at-risk" | "off-track";

/** Capability/PM status (USER_STATUS) */
export type UserStatus = "on-track" | "at-risk" | "off-track";

/** AI / GenAI recommended status (AI_STATUS) - same choice set */
export type AIStatus = "on-track" | "at-risk" | "off-track";

/** Legacy/UI status color for badges and dots */
export type StatusColor = "green" | "yellow" | "red";

/** Capability size (SIZE field) */
export type CapabilitySize = "S" | "M" | "L" | "XL";

/** Segment / GPA single-select display (id + name from Airtable) */
export type SelectOption = { id: string; name: string; color?: string };

/** Linked record reference from Airtable cell value */
export type LinkedRecordRef = { id: string; name?: string };

// --- Metric (linked to Capability) ---

export interface CapabilityMetric {
  id: string;
  name: string;
  baseline: number | null;
  target: number | null;
  actual: number | null;
  status: InitiativeStatus;
}

// --- Epic (under Capability) ---

export interface Epic {
  id: string;
  name: string;
  status: string;
}

// --- Capability (leaf) ---

export interface Capability {
  id: string;
  name: string;
  size: CapabilitySize;
  /** Raw status from Airtable (e.g. "Red", "Yellow", "Green") for display and badge class. */
  status: string;
  /** Parsed status for logic and DualStatusBadge when needed. */
  userStatus: UserStatus;
  /** Raw AI-recommended status from Airtable for display. */
  aiStatus: string;
  depCount: number;
  depScore: number;
  /** Metrics linked to this capability (from METRICS table). */
  metrics: CapabilityMetric[];
  /** Epics linked to this capability (from Epics table). */
  epics?: Epic[];
  /** Optional display fields if present in Airtable (e.g. startQ, launchQ, pct, mStatus) */
  startQ?: string;
  launchQ?: string;
  pct?: number;
  mStatus?: StatusColor;
  statusNotes?: string;
}

// --- Project (Sub-Initiative) ---

export interface Project {
  id: string;
  name: string;
  status: string;
  capabilities: Capability[];
}

// --- Initiative (root) ---

export interface Initiative {
  id: string;
  name: string;
  goalAlignment: string;
  gpa: string;
  segments: string[];
  trueStatus: InitiativeStatus;
  aiStatus?: string;
  productLead: string;
  risks: string;
  financeProjection: string;
  pillar: string;
  fy: string;
  market?: string[];
  targetGmv: number;
  financeType?: string;
  rawFinance?: string;
  projects: Project[];
  /** Capabilities linked to this initiative but not linked to any project. */
  orphanedCapabilities?: Capability[];
  /** Linked record IDs for counting (e.g. in AggCard). */
  projectIds?: string[];
  /** Initiative Groups linked record IDs from Airtable. */
  initiativeGroups?: string[];
}

// --- UI / aggregation types (used by index.tsx) ---

export interface StatusConfig {
  color: string;
  bg: string;
  dot: string;
  label: string;
}

export interface Metric {
  label: string;
  baseline: string;
  target: string;
  actual: string;
  status: InitiativeStatus;
}

/** Segment breakdown: segment name -> count */
export type SegmentBreakdown = Record<string, number>;

export interface Priority {
  name: string;
  initiatives: Initiative[];
  count: number;
  capabilitiesTotal: number;
  onTrack: number;
  atRisk: number;
  offTrack: number;
  segments: SegmentBreakdown;
  gpa: string[];
  overallStatus: InitiativeStatus;
  keyMetrics: Metric[];
}

export type ViewMode = "priority" | "gpa" | "segment";
export type StatKey = "initiatives" | "capabilities" | "on-track" | "at-risk" | "off-track";

/** Flatten capabilities from Initiative -> Project -> Capability for UI that expects initiative.capabilities. */
export function getCapabilities(initiative: Initiative): Capability[] {
  return initiative.projects.flatMap((p) => p.capabilities);
}
