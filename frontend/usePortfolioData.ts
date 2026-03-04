import { useMemo } from "react";
import { useBase, useRecords } from "@airtable/blocks/interface/ui";
import type { Record } from "@airtable/blocks/interface/models";
import { TABLES, FIELDS } from "./schema";
import type { Initiative, Project, Capability, CapabilityMetric, Epic, InitiativeStatus, UserStatus, AIStatus, CapabilitySize, Priority, SegmentBreakdown } from "./types";
import { getCapabilities } from "./types";

const DEFAULT_FINANCE = "MISSING";
const DEFAULT_STRING = "";
const DEFAULT_STATUS: InitiativeStatus = "on-track";
const DEFAULT_USER_STATUS: UserStatus = "on-track";
const DEFAULT_AI_STATUS: AIStatus = "on-track";
const DEFAULT_SIZE: CapabilitySize = "M";

/** Normalize single-select value to string (use name for display). */
function singleSelectName(cell: unknown): string {
  if (cell === null || cell === undefined) return DEFAULT_STRING;
  if (typeof cell === "object" && "name" in cell && typeof (cell as { name: unknown }).name === "string")
    return (cell as { name: string }).name;
  if (typeof cell === "string") return cell;
  return DEFAULT_STRING;
}

/** Normalize multi-select or linked records to array of ids. */
function linkedRecordIds(cell: unknown): string[] {
  if (cell === null || cell === undefined) return [];
  if (!Array.isArray(cell)) return [];
  return cell
    .filter((item): item is { id: string } => typeof item === "object" && item !== null && "id" in item && typeof (item as { id: unknown }).id === "string")
    .map((item) => item.id);
}

/** Parse initiative status from cell (single-select name or string). */
function parseInitiativeStatus(cell: unknown): InitiativeStatus {
  const s = singleSelectName(cell).toLowerCase();
  if (s === "at-risk" || s === "at risk") return "at-risk";
  if (s === "off-track" || s === "off track") return "off-track";
  if (s === "on-track" || s === "on track") return "on-track";
  return DEFAULT_STATUS;
}

/** Parse user/AI status (green/yellow/red or on-track/at-risk/off-track). */
function parseUserOrAIStatus(cell: unknown): UserStatus {
  const raw = typeof cell === "string" ? cell : singleSelectName(cell);
  const s = raw.toLowerCase();
  if (s === "yellow" || s === "at-risk" || s === "at risk") return "at-risk";
  if (s === "red" || s === "off-track" || s === "off track") return "off-track";
  if (s === "green" || s === "on-track" || s === "on track") return "on-track";
  return DEFAULT_USER_STATUS;
}

/** Parse capability size. */
function parseSize(cell: unknown): CapabilitySize {
  const raw = typeof cell === "string" ? cell : singleSelectName(cell);
  const u = raw.toUpperCase();
  if (u === "S" || u === "M" || u === "L" || u === "XL") return u as CapabilitySize;
  return DEFAULT_SIZE;
}

/** Safe number from cell. */
function num(cell: unknown, fallback: number): number {
  if (typeof cell === "number" && !Number.isNaN(cell)) return cell;
  if (typeof cell === "string") {
    const n = Number(cell);
    if (!Number.isNaN(n)) return n;
  }
  return fallback;
}

/** Metric numeric field: return number or null if empty. */
function metricNum(cell: unknown): number | null {
  if (cell === null || cell === undefined || cell === "") return null;
  if (typeof cell === "number" && !Number.isNaN(cell)) return cell;
  if (typeof cell === "string") {
    const trimmed = cell.trim();
    if (trimmed === "") return null;
    const n = Number(trimmed);
    if (!Number.isNaN(n)) return n;
  }
  return null;
}

/** Safe string from cell; default for blank. */
function str(cell: unknown, fallback: string): string {
  if (cell === null || cell === undefined) return fallback;
  const s = String(cell).trim();
  return s === "" ? fallback : s;
}

const UNASSIGNED = "Unassigned";

/** Parse FINANCE_PROJECTION text: extract numeric value and label (e.g. GMV vs savings). */
function parseFinance(val: string): { value: number; type: string } {
  if (!val) return { value: 0, type: "" };
  const match = val.match(/\$?\d+(?:\.\d+)?[MBK]?/i);
  if (!match) return { value: 0, type: val.trim() };
  const rawNum = match[0].toUpperCase().replace("$", "");
  const type = val.replace(match[0], "").trim();

  let multiplier = 1;
  if (rawNum.includes("B")) multiplier = 1_000_000_000;
  if (rawNum.includes("M")) multiplier = 1_000_000;
  if (rawNum.includes("K")) multiplier = 1_000;

  const num = parseFloat(rawNum.replace(/[MBK]/g, ""));
  return { value: Number.isNaN(num) ? 0 : num * multiplier, type };
}

export interface UsePortfolioDataResult {
  initiatives: Initiative[];
  enterprisePriorities: Priority[];
  otherPriorities: Priority[];
  isLoading: boolean;
  error: Error | null;
}

/** Build priority groupings from initiatives by goalAlignment (used for enterprise/other split). */
function buildPrioritiesFromInitiatives(initiatives: Initiative[]): Priority[] {
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

/**
 * Fetches Initiatives, Projects, and Capabilities from Airtable and stitches
 * them into Initiative -> Project -> Capability hierarchy.
 * Uses useRecords on all three tables at top level (no N+1).
 */
export function usePortfolioData(): UsePortfolioDataResult {
  const base = useBase();

  const initiativesTable = base.getTableByIdIfExists(TABLES.INITIATIVES);
  const projectsTable = base.getTableByIdIfExists(TABLES.PROJECTS);
  const capabilitiesTable = base.getTableByIdIfExists(TABLES.CAPABILITIES);
  const metricsTable = base.getTableByIdIfExists(TABLES.METRICS);
  const epicsTable = base.getTableByIdIfExists(TABLES.EPICS) ?? base.getTableByNameIfExists("Epics");

  const fallbackTable = base.tables[0];
  const initiativesRecords = useRecords(initiativesTable ?? fallbackTable);
  const projectsRecords = useRecords(projectsTable ?? fallbackTable);
  const capabilitiesRecords = useRecords(capabilitiesTable ?? fallbackTable);
  const metricsRecords = useRecords(metricsTable ?? fallbackTable);
  const epicsRecords = useRecords(epicsTable ?? fallbackTable);

  const { initiatives, enterprisePriorities, otherPriorities, error } = useMemo(() => {
    const err: Error | null =
      !initiativesTable || !projectsTable || !capabilitiesTable
        ? new Error("One or more required tables are missing from this base.")
        : null;

    if (err) {
      return {
        initiatives: [] as Initiative[],
        enterprisePriorities: [],
        otherPriorities: [],
        error: err,
      };
    }

    const I = initiativesTable!;
    const P = projectsTable!;
    const C = capabilitiesTable!;

    const nameFieldI = I.getFieldIfExists(FIELDS.INITIATIVE.NAME);
    const goalFieldI = I.getFieldIfExists(FIELDS.INITIATIVE.GOAL_ALIGNMENT);
    const gpaFieldI = I.getFieldIfExists(FIELDS.INITIATIVE.GPA);
    const segmentsFieldI = I.getFieldIfExists(FIELDS.INITIATIVE.SEGMENTS);
    const productLeadFieldI = I.getFieldIfExists(FIELDS.INITIATIVE.PRODUCT_LEAD);
    const risksFieldI = I.getFieldIfExists(FIELDS.INITIATIVE.RISKS);
    const financeFieldI = I.getFieldIfExists(FIELDS.INITIATIVE.FINANCE_PROJECTION);
    const pillarFieldI = I.getFieldIfExists(FIELDS.INITIATIVE.PILLAR);
    const fyFieldI = I.getFieldIfExists(FIELDS.INITIATIVE.FY);
    const marketFieldI = I.getFieldIfExists(FIELDS.INITIATIVE.MARKET);
    const initiativeGroupsFieldI = I.getFieldIfExists(FIELDS.INITIATIVE.INITIATIVE_GROUPS);

    const nameFieldP = P.getFieldIfExists(FIELDS.PROJECT.NAME);
    const initiativeLinkFieldP = P.getFieldIfExists(FIELDS.PROJECT.INITIATIVE_LINK);
    const statusFieldP = P.getFieldIfExists(FIELDS.PROJECT.STATUS);

    const nameFieldC = C.getFieldIfExists(FIELDS.CAPABILITY.NAME);
    const sizeFieldC = C.getFieldIfExists(FIELDS.CAPABILITY.SIZE);
    const statusFieldC = C.getFieldIfExists(FIELDS.CAPABILITY.STATUS);
    const aiStatusFieldC = C.getFieldIfExists(FIELDS.CAPABILITY.AI_STATUS);
    const depCountFieldC = C.getFieldIfExists(FIELDS.CAPABILITY.DEP_COUNT);
    const depScoreFieldC = C.getFieldIfExists(FIELDS.CAPABILITY.DEP_SCORE);
    const projectLinkFieldC = C.getFieldIfExists(FIELDS.CAPABILITY.PROJECT_LINK);
    const statusNotesFieldC = C.getFieldIfExists(FIELDS.CAPABILITY.STATUS_NOTES);

    const initiativesRaw =
      I.id === TABLES.INITIATIVES ? initiativesRecords : ([] as Record[]);
    const projectsRaw =
      P.id === TABLES.PROJECTS ? projectsRecords : ([] as Record[]);
    const capabilitiesRaw =
      C.id === TABLES.CAPABILITIES ? capabilitiesRecords : ([] as Record[]);

    const capabilityById = new Map<string, Capability>();
    const capIdToLinkedProjectIds = new Map<string, string[]>();

    const initiativesById = new Map<string, Initiative>();
    for (const rec of initiativesRaw) {
      const name = nameFieldI ? singleSelectName(rec.getCellValue(nameFieldI)) : rec.name;
      const goalAlignment = goalFieldI
        ? (rec.getCellValueAsString(goalFieldI) ?? UNASSIGNED).trim() || UNASSIGNED
        : UNASSIGNED;
      const gpa = gpaFieldI
        ? (rec.getCellValueAsString(gpaFieldI) ?? UNASSIGNED).trim() || UNASSIGNED
        : UNASSIGNED;
      const segmentStr = segmentsFieldI
        ? (rec.getCellValueAsString(segmentsFieldI) ?? "").trim() || UNASSIGNED
        : UNASSIGNED;
      const segments: string[] = segmentStr ? [segmentStr] : [];
      const rawLeads = productLeadFieldI
        ? (rec.getCellValue(productLeadFieldI) as Array<{ name: string }> | null)
        : null;
      const productLead =
        rawLeads && rawLeads.length > 0
          ? rawLeads.map((l) => l.name).join(", ")
          : "None provided";
      const risks = risksFieldI ? str(rec.getCellValue(risksFieldI), DEFAULT_STRING) : DEFAULT_STRING;
      const rawFinance = financeFieldI
        ? (rec.getCellValueAsString(financeFieldI) ?? "").trim()
        : "";
      const { value: targetGmv, type: financeType } = parseFinance(rawFinance);
      const financeProjection = rawFinance || DEFAULT_FINANCE;
      const pillar = pillarFieldI
        ? (rec.getCellValueAsString(pillarFieldI) ?? "").trim() || UNASSIGNED
        : UNASSIGNED;
      const fy = fyFieldI
        ? (rec.getCellValueAsString(fyFieldI) ?? "").trim() || UNASSIGNED
        : UNASSIGNED;
      const rawMarket = marketFieldI ? rec.getCellValue(marketFieldI) : null;
      const market: string[] = Array.isArray(rawMarket)
        ? rawMarket
            .map((m: { name?: string } | string) => (typeof m === "object" && m && "name" in m ? m.name : String(m ?? "")))
            .filter((s): s is string => Boolean(s))
        : [];
      const initiativeGroups = initiativeGroupsFieldI
        ? linkedRecordIds(rec.getCellValue(initiativeGroupsFieldI))
        : undefined;

      initiativesById.set(rec.id, {
        id: rec.id,
        name: name || DEFAULT_STRING,
        goalAlignment,
        gpa,
        segments,
        trueStatus: DEFAULT_STATUS,
        productLead,
        risks,
        financeProjection,
        pillar,
        fy,
        market,
        targetGmv,
        financeType: financeType || undefined,
        rawFinance: rawFinance || undefined,
        projects: [],
        orphanedCapabilities: [],
        initiativeGroups: initiativeGroups?.length ? initiativeGroups : undefined,
      });
    }

    // --- Step A & B: Map Capabilities (upward link: PROJECT_LINK). Epics assigned later. ---
    for (const rec of capabilitiesRaw) {
      const name = nameFieldC ? singleSelectName(rec.getCellValue(nameFieldC)) : rec.name;
      const size = sizeFieldC ? parseSize(rec.getCellValue(sizeFieldC)) : DEFAULT_SIZE;
      const status = statusFieldC ? (rec.getCellValueAsString(statusFieldC) ?? "").trim() || "Unassigned" : "Unassigned";
      const aiStatus = aiStatusFieldC ? (rec.getCellValueAsString(aiStatusFieldC) ?? "").trim() || "" : "";
      const userStatus = statusFieldC ? parseUserOrAIStatus(rec.getCellValue(statusFieldC)) : DEFAULT_USER_STATUS;
      const depCount = depCountFieldC ? num(rec.getCellValue(depCountFieldC), 0) : 0;
      const depScore = depScoreFieldC ? num(rec.getCellValue(depScoreFieldC), 0) : 0;
      const statusNotes = statusNotesFieldC
        ? (rec.getCellValueAsString(statusNotesFieldC) ?? "").trim()
        : undefined;

      const cap: Capability = {
        id: rec.id,
        name: name || DEFAULT_STRING,
        size,
        status,
        userStatus,
        aiStatus,
        depCount,
        depScore,
        metrics: [],
        statusNotes: statusNotes || undefined,
        epics: [],
      };

      const rawProjectLinks = projectLinkFieldC
        ? (rec.getCellValue(projectLinkFieldC) as Array<{ id: string }> | null)
        : null;
      const linkedProjectIds = rawProjectLinks ? rawProjectLinks.map((x) => x.id) : [];
      capabilityById.set(rec.id, cap);
      capIdToLinkedProjectIds.set(rec.id, linkedProjectIds);
    }

    // --- Step C: Map Projects (TABLES.PROJECTS). Extract INITIATIVE_LINK; assign capabilities by link. ---
    const projectById = new Map<string, Project>();
    const projIdToLinkedInitIds = new Map<string, string[]>();
    const allCapabilities = Array.from(capabilityById.values());

    for (const rec of projectsRaw) {
      const name = nameFieldP ? singleSelectName(rec.getCellValue(nameFieldP)) : rec.name;
      const status = statusFieldP ? (rec.getCellValueAsString(statusFieldP) ?? "").trim() || "" : "";

      const rawInitLinks = initiativeLinkFieldP
        ? (rec.getCellValue(initiativeLinkFieldP) as Array<{ id: string }> | null)
        : null;
      const linkedInitIds = rawInitLinks ? rawInitLinks.map((x) => x.id) : [];

      const capabilities = allCapabilities.filter((cap) =>
        capIdToLinkedProjectIds.get(cap.id)?.includes(rec.id)
      );

      projectById.set(rec.id, {
        id: rec.id,
        name: name || DEFAULT_STRING,
        status,
        capabilities,
      });
      projIdToLinkedInitIds.set(rec.id, linkedInitIds);
    }

    const allProjects = Array.from(projectById.values());

    // --- Step D: Map Initiatives. Assign projects by link; orphanedCapabilities; derive trueStatus from children. ---
    for (const init of initiativesById.values()) {
      init.projects = allProjects.filter((proj) =>
        projIdToLinkedInitIds.get(proj.id)?.includes(init.id)
      );
      init.projectIds = init.projects.map((p) => p.id);
      init.orphanedCapabilities = init.orphanedCapabilities ?? [];

      const initProjects = init.projects;
      const orphanedCaps = init.orphanedCapabilities ?? [];
      const childStatuses: string[] = [];
      initProjects.forEach((p) => {
        if (p.status) childStatuses.push(p.status.toLowerCase().replace(/\s+/g, "-"));
        p.capabilities?.forEach((c) => {
          if (c.status) childStatuses.push(c.status.toLowerCase().replace(/\s+/g, "-"));
        });
      });
      orphanedCaps.forEach((c) => {
        if (c.status) childStatuses.push(c.status.toLowerCase().replace(/\s+/g, "-"));
      });
      let derivedStatus: InitiativeStatus = "on-track";
      if (childStatuses.some((s) => s === "red" || s === "off-track")) {
        derivedStatus = "off-track";
      } else if (childStatuses.some((s) => s === "yellow" || s === "at-risk")) {
        derivedStatus = "at-risk";
      } else if (childStatuses.length === 0) {
        derivedStatus = "on-track";
      }
      init.trueStatus = derivedStatus;
    }

    // --- Epics: Map Epics; assign cap.epics by EPIC.CAPABILITY_LINK (upward link). ---
    try {
      if (epicsTable && FIELDS.EPIC) {
        const nameFieldE = epicsTable.getFieldIfExists(FIELDS.EPIC.NAME);
        const statusFieldE = epicsTable.getFieldIfExists(FIELDS.EPIC.STATUS);
        const capabilityLinkFieldE = epicsTable.getFieldIfExists(FIELDS.EPIC.CAPABILITY_LINK);
        const epicsRaw = epicsTable.id === TABLES.EPICS || epicsTable.name === "Epics" ? epicsRecords : ([] as Record[]);

        for (const rec of epicsRaw) {
          const epicName = nameFieldE ? (rec.getCellValueAsString(nameFieldE) ?? "").trim() || rec.name : rec.name;
          const epicStatus = statusFieldE ? (rec.getCellValueAsString(statusFieldE) ?? "").trim() || "" : "";
          const epic: Epic = { id: rec.id, name: epicName, status: epicStatus };

          const linkedCapIds = capabilityLinkFieldE
            ? (rec.getCellValue(capabilityLinkFieldE) as Array<{ id: string }> | null)
            : null;
          const capIds = linkedCapIds ? linkedCapIds.map((x) => x.id) : [];

          for (const capId of capIds) {
            const cap = capabilityById.get(capId);
            if (cap && cap.epics) cap.epics.push(epic);
          }
        }
      }
    } catch {
      // Epics table or fields not configured; leave cap.epics as []
    }

    if (metricsTable && metricsTable.id === TABLES.METRICS) {
      const nameFieldM = metricsTable.getFieldIfExists(FIELDS.METRIC.NAME);
      const baselineFieldM = metricsTable.getFieldIfExists(FIELDS.METRIC.BASELINE);
      const targetFieldM = metricsTable.getFieldIfExists(FIELDS.METRIC.TARGET);
      const actualFieldM = metricsTable.getFieldIfExists(FIELDS.METRIC.ACTUAL);
      const statusFieldM = metricsTable.getFieldIfExists(FIELDS.METRIC.STATUS);
      const capabilityLinkFieldM = metricsTable.getFieldIfExists(FIELDS.METRIC.CAPABILITY_LINK);

      const metricsRaw =
        metricsTable.id === TABLES.METRICS ? metricsRecords : ([] as Record[]);

      for (const rec of metricsRaw) {
        const name = nameFieldM
          ? (rec.getCellValueAsString(nameFieldM) ?? "").trim() || DEFAULT_STRING
          : rec.name;
        const baseline = baselineFieldM ? metricNum(rec.getCellValue(baselineFieldM)) : null;
        const target = targetFieldM ? metricNum(rec.getCellValue(targetFieldM)) : null;
        const actual = actualFieldM ? metricNum(rec.getCellValue(actualFieldM)) : null;
        const statusStr = statusFieldM
          ? (rec.getCellValueAsString(statusFieldM) ?? "").trim()
          : "";
        const status = parseInitiativeStatus(statusStr || null);

        const metric: CapabilityMetric = {
          id: rec.id,
          name,
          baseline,
          target,
          actual,
          status,
        };

        const capabilityIds = capabilityLinkFieldM
          ? linkedRecordIds(rec.getCellValue(capabilityLinkFieldM))
          : [];
        for (const capId of capabilityIds) {
          const cap = capabilityById.get(capId);
          if (cap) {
            cap.metrics.push(metric);
          }
        }
      }
    }

    const initiatives: Initiative[] = Array.from(initiativesById.values());
    const allPriorities = buildPrioritiesFromInitiatives(initiatives);
    const enterprisePriorities = allPriorities.filter(
      (p) => p.name !== UNASSIGNED && p.name.startsWith("Enterprise Priority")
    );
    const otherPriorities = allPriorities.filter(
      (p) => p.name !== UNASSIGNED && p.name.startsWith("Other")
    );
    console.log("Mapped Initiatives:", initiatives);
    return { initiatives, enterprisePriorities, otherPriorities, error: null };
  }, [
    initiativesTable,
    projectsTable,
    capabilitiesTable,
    metricsTable,
    epicsTable,
    initiativesRecords,
    projectsRecords,
    capabilitiesRecords,
    metricsRecords,
    epicsRecords,
  ]);

  const isLoading = false;

  return {
    initiatives,
    enterprisePriorities,
    otherPriorities,
    isLoading,
    error,
  };
}
