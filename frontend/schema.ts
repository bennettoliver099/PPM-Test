/**
 * Centralized Airtable table and field IDs for the Executive Initiative Reporting extension.
 * Use strict IDs to avoid breaking when field or table names change in Airtable.
 * Use with base.getTableById(), table.getFieldByIdIfExists(), and record.getCellValue(field).
 */

/** Airtable table IDs */
export const TABLES = {
  INITIATIVES: "tblKDKvi05iMh36yw",
  PROJECTS: "tblr8pI9Zkl2U60PI", // Sub-Initiatives
  CAPABILITIES: "tblkEBxlCGBZp8SJp",
  METRICS: "tbl7Jic9zv4OWHiHr",
  // TODO: Replace with exact Airtable table ID for Epics (or use base.getTableByNameIfExists("Epics")).
  EPICS: "tblMPtjMbg2ZHHUXq",
} as const;

export type TableId = (typeof TABLES)[keyof typeof TABLES];

/** Field IDs by table. Use with getFieldByIdIfExists(table, fieldId). */
export const FIELDS = {
  INITIATIVE: {
    NAME: "fld0smLWG5mnrMRlZ",
    GOAL_ALIGNMENT: "fldNpgS6JEmnK4EuX", 
    GPA: "fld4LDrQ0MWaxBXZD", 
    SEGMENTS: "fldv0YncrH1K15MTS", // CORRECTED: "Segment AOP Owner" from Initiatives table
    MARKET: "fldWV6jb5BfDFvvQr",
    FY: "fldWltc8evaQG5Rky", 
    INITIATIVE_GROUPS: "fldQR724zF77u6gaQ",
    PRODUCT_LEAD: "fldL6T1KQiRodNiGU",
    RISKS: "fld5SbLYXisJJ6oSu",
    FINANCE_PROJECTION: "fldEweq0c8kocTRHB",
    PILLAR: "fldwN1TSj0pWJcNYv", 
    
    // Note: Initiatives DO NOT have a "True Status" field in your JSON. 
    // They roll up status from projects/capabilities. Remove TRUE_STATUS and AI_STATUS from INITIATIVE.
  },
  PROJECT: {
    NAME: "fldCLYvf3pcFUZuUY",
    INITIATIVE_LINK: "fldMVqYa7FAzfIWp8",
    STATUS: "fldB7wQvWZn7OZdWk",
  },
  CAPABILITY: {
    NAME: "fld3131kLQr6irkJ9",
    SIZE: "fldj11Z326CrOIBqD",
    STATUS: "fldC6LdgmeH6OJZEo", // CORRECT: "True Status" is here
    USER_STATUS: "fldBbTYgnL6Gm9YJp",
    AI_STATUS: "fldohfaPoY1MmcPiN", // CORRECT: "Recommended Status" is here
    DEP_COUNT: "fldnlLBsZt7jOZoZp",
    DEP_SCORE: "fldM5z50B1WkQmUX4",
    SEGMENTS: "fldp33fryZXn4aGyI", // CORRECT: Capability-level Segments
    PROJECT_LINK: "fldW3PH965BrFM3Hf",
    STATUS_NOTES: "fldAB1zLEkTIkUaVY",
  },
  METRIC: {
    NAME: "fldEzqwwqEbSP02eW",
    BASELINE: "fldM1Fkla4402XSuo",
    TARGET: "fldbaVy3OwrnrDBlS",
    ACTUAL: "fldzKytURj8X2zj81",
    STATUS: "fldJbL6fmuHtN5njT",
    CAPABILITY_LINK: "fldD6pDy00aMUWpaN",
  },
  // TODO: Replace with actual Airtable Field IDs for the Epics table.
  EPIC: {
    NAME: "fldpMXzdBGgaIfePO",
    STATUS: "fldj7kMjIoAsz0mVD",
    CAPABILITY_LINK: "fldzZK2DBEdKtsF6L",
  },
} as const;

export type InitiativeFieldId = (typeof FIELDS.INITIATIVE)[keyof typeof FIELDS.INITIATIVE];
export type ProjectFieldId = (typeof FIELDS.PROJECT)[keyof typeof FIELDS.PROJECT];
export type CapabilityFieldId = (typeof FIELDS.CAPABILITY)[keyof typeof FIELDS.CAPABILITY];
export type MetricFieldId = (typeof FIELDS.METRIC)[keyof typeof FIELDS.METRIC];
export type EpicFieldId = (typeof FIELDS.EPIC)[keyof typeof FIELDS.EPIC];
