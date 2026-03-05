# Product Hub — Session Log

**Project:** PPM Executive Dashboard (Airtable Interface Extension)
**Repo:** https://github.com/bennettoliver099/PPM-Test
**Block ID:** `blkMvLVaFtkU0Jcyk`
**Stack:** React 19 + TypeScript, `@airtable/blocks` interface SDK, no external UI libraries

---

## Core Goal

Build a consumer-grade executive portfolio dashboard embedded inside an Airtable Interface page. The dashboard surfaces initiative health, project and capability hierarchies, and brand-compliant Walmart design language for use in stakeholder reporting.

---

## Final Solution

### Architecture

```
frontend/
  index.tsx          # ~2,800 lines — all UI, components, data transform
  style.css          # Design tokens (CSS vars) + component classes
  usePortfolioData.ts # Airtable SDK data fetching (not modified this session)
  types.ts           # TypeScript interfaces (not modified this session)
  assets/
    wmt-spark.svg
    wmt-wordmark-trueblue.svg
    fonts/           # EverydaySans woff/woff2 (Regular, Medium, Bold)
```

### Brand Token System

All inline styles reference a single `C` constant object — no hardcoded hex values anywhere in component code:

```tsx
const C = {
  bgPage, bgPanel, bgHover, bgActive,
  border, borderLight,
  textPrimary, textSecondary, textTertiary,
  blue, blueHover, blueSoft,
  green, greenSoft, greenDark,
  amber, amberSoft, amberDark,
  red, redSoft, redDark,
  purple, purpleSoft, purpleDark,
  walmartTrueBlue, walmartBentonvilleBlue, walmartSparkYellow,
  shadowCard, shadowCardHover, shadowDropdown, shadowModal, shadowSheet,
  radiusCheckbox, radiusChip, radiusButton, radiusCard, radiusHero, radiusPill,
  font,
} as const;
```

### Key Components

| Component | Purpose |
|---|---|
| `TypeBadge` | Entity-type label (goal/initiative/project/capability/epic). White bg, gray border, subtle. Always `flex-start` aligned above record names. |
| `Badge` | Initiative status chip (On Track / At Risk / Off Track) with colored `StatusDot` inside. White bg, gray border, colored text. |
| `DependencyPill` | Compact dep count badge: `"3 deps · High"` |
| `CapabilityRow` | Enterprise-grade capability display: TypeBadge → name+status row → meta row (quarters, size, deps) → status notes → epics → metrics |
| `AggCard` | Priority/GPA/Segment card with progress bar, status legend, conditional finance footer |
| `Summary` (Portfolio Health) | 6-stat banner with clickable single-select status filter |
| `SideSheet` | Slide-in panel: initiative list → initiative detail (projects → capabilities → epics/metrics) |
| `Filters` | Tab navigation (By Priority / By GPA / By Segment) + segment control + FY dropdown + market multi-select |

### Filter Architecture

Two filtered sets are maintained separately to prevent the status filter from contaminating the summary stats:

```tsx
// Base: seg + FY + market only — feeds Summary stats, attention banner
const filteredBase = useMemo(() => initiatives.filter(seg + fy + market), [...]);

// Full: base + statusFilter — feeds the card grid
const filteredInitiatives = useMemo(
  () => statusFilter ? filteredBase.filter(i => i.trueStatus === statusFilter) : filteredBase,
  [filteredBase, statusFilter]
);
```

Status filter is single-select toggle: click active filter to deselect, click a different one to switch.

### Global Header

Bentonville Blue (`#001E60`) background with:
- Official Walmart Spark SVG (SparkYellow, unchanged)
- Official Walmart Wordmark SVG (paths changed from TrueBlue → white for dark bg)
- EverydaySans "Product Hub" label in white
- Padding: `16px 24px` (20% taller than original)

### Badge System

All badges (`ld-badge` CSS class) use a neutral pattern:
- White background (`#ffffff`)
- 1px solid gray border (`var(--color-border)`)
- Colored text only (green/amber/red for status, gray for neutral)
- 10px font, `1px 6px` padding, 4px border-radius
- Never stretches — TypeBadge has `alignSelf: flex-start`

---

## Key Decisions

### 1. All styles inline via `C` tokens (no Tailwind in components)
Airtable's interface extension sandbox strips some CSS. Inline styles with a central token object (`C`) are the safest pattern — no build-time surprises, full TypeScript autocomplete.

### 2. Interface extension vs. base extension SDK
`@airtable/blocks/interface/ui` is required (not `@airtable/blocks/ui`). `initializeBlock({ interface: () => <App /> })` form is required. `baseId: "NONE"` in `.block/remote.json`. Package must be installed explicitly at `0.0.0-experimental-801a212b8-20260206` — the `interface-alpha` dist-tag resolves to an incompatible version.

### 3. `CapabilityRow` extracted as a standalone component
The old pattern rendered capability cards inline in two places (project capabilities + orphaned capabilities). Extracting `CapabilityRow` eliminates ~200 lines of duplication and ensures both sections are visually identical.

### 4. Status filter decoupled from summary stats
When `statusFilter` was applied to `filteredInitiatives` and then `filteredInitiatives` fed the Summary banner, clicking "Off Track" caused On Track/At Risk counts to drop to 0 — confusing and misleading. Splitting into `filteredBase` (stats source) vs. `filteredInitiatives` (cards source) fixed this.

### 5. TypeBadge: neutral over colorful
Initially tried colored entity badges (blue for goal, purple for initiative, etc.). User feedback: too visually loud. Reverted to white bg + gray border + gray text — subtle enough to provide context without competing with status signals.

### 6. Market dropdown scroll fix
`overflow: "hidden"` was declared after `overflowY: "auto"` in the same style object, making the dropdown non-scrollable. Fixed by removing the redundant `overflow` override.

---

## Constraints & Edge Cases

| Constraint | Handling |
|---|---|
| Airtable sandbox — no CDN fonts | EverydaySans served via `@font-face` from local `assets/fonts/` |
| Airtable sandbox — no SVG `<img>` src | Official Walmart logos inlined as JSX paths |
| `applicationId = "NONE"` for interface extensions | `baseId: "NONE"` in `.block/remote.json` — otherwise "wrong base" error |
| `textTransform: 'uppercase'` TypeScript cast | Must use `as const` to satisfy React's CSSProperties type |
| TypeBadge inside flexbox containers | `alignSelf: 'flex-start'` prevents it from stretching to fill row height |
| `DependencyPill` text too long | Shortened from `"3 GPA Dependencies (High Complexity)"` → `"3 deps · High"` |
| `statusFilter` poisoning summary counts | Decoupled `filteredBase` (no status filter) from `filteredInitiatives` (with status filter) |
| Cap status raw values (`"green"`, `"red"`) | `capStatusClass()` maps raw values to CSS class names via `toLowerCase().replace(/\s+/g, "-")` |

---

## What Could Have Been Asked More Clearly

### 1. Brand spec upfront, not retroactively
Many changes were applied in multiple passes (colors → tokens → brand assets → badge style) because the full brand constraints weren't specified at the start. Providing a complete design brief — including badge style, spacing scale, font rules, and Walmart brand assets — before any code is written would eliminate 3–4 revision cycles.

**Better prompt:** *"Build this using Walmart brand guidelines: TrueBlue + BentonvilleBlue for headers, EverydaySans for display text, SparkYellow only for the logo. Badges should be white bg with 1px gray border, not colored. Spacing: 4/8/12/16/20/24px only. Font sizes: 10/11/12/13/15/20px only."*

### 2. "Fix the filter" needs a behavior description
"The filter objects at the top are buggy" was ambiguous — it could mean visual glitches, broken state, wrong counts, or non-functional clicks. Each interpretation leads to different fixes.

**Better prompt:** *"When I click 'Off Track' in the Portfolio Health banner, the On Track and At Risk counts go to 0. They should always show the true totals regardless of which status I'm filtering."*

### 3. "Enterprise grade" needs a reference point
"Make it enterprise grade" was interpreted as: consistent layout, flat hierarchy indicators, border separators instead of cards-in-cards. But it could also mean: density, accessibility, keyboard navigation, or specific patterns from tools like Jira/Salesforce.

**Better prompt:** *"Capability rows should look like Notion database rows or Linear issue lists — flat, dense, separator lines only, no nested card boxes."*

### 4. TypeBadge placement was specified gradually
"Top left above record names" was specified after "make them discrete" and after the initial placement was wrong. The final layout (badge above name, name below) should have been described in the first request.

**Better prompt:** *"Each record should have its entity type label (goal/initiative/project/etc.) as a small uppercase chip positioned above the record name, not inline with it."*

### 5. Single-select vs. multi-select was flip-flopped
Multi-select was built, then immediately reverted to single-select. A clearer statement of the intended interaction model would avoid the round-trip.

**Better prompt:** *"Clicking a status filter should activate it and dim the others. Clicking the same one again deselects it. Only one can be active at a time."*

### 6. "20% thicker" is relative without a baseline
The header padding change needed interpretation. Explicitly stating the target padding or referencing a visual example would be cleaner.

**Better prompt:** *"Increase the top navigation bar height so it's about 54px tall instead of ~46px."*
