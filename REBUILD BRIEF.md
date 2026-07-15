# Karjat Protection Portal — Rebuild Brief

**Standing directive (Ashish, saved):** the current portal was built incrementally across many
sessions and he is not convinced by it. When he finishes specifying his needs, Claude is to
**rebuild the whole website from scratch** — a clean rewrite, not more patches. Ashish supports
by deploying `code.gs` and hosting on GitHub.

This document accumulates the requirements so the rewrite starts from a complete brief.

---

## Context (authoritative — confirmed by Ashish, do not assume from generic patterns)

- **Substation:** 400/220 kV, 1002 MVA, Karjat (MSETCL / MahaTransco). S/S Code **J966**.
  SLD doc date 11-JAN-2022. 62 relays, 45,593 settings, 26 real bays + 3 bus-zone pseudo-bays.
- **400 kV — one-and-a-half breaker scheme**, 5 diameters:
  - Dia1: 401 Bus Reactor 125MVAr (top) – 402 Tie Bay1 (tie) – 403 ICT1 HV 501MVA (bottom)
  - Dia2: 404 Girawali Ckt1 – 405 Tie Bay2 – 406 ICT2 HV
  - Dia3: 407 Girawali Ckt2 – 408 Tie Bay3 – 409 ICT3 HV (no files)
  - Dia4: 410 Lonikand Ckt1 – 411 Tie Bay4 – 412 FutBay2 (no files)
  - Dia5: 413 Lonikand Ckt2 – 414 Tie Bay5 – 415 FutBay3 (no files)
  - **Spare ICT is on BUS-2, not an auxiliary bus.**
- **220 kV — double main bus + transfer bus**, bays 203–212:
  - 203 ICT2 IV, 204 ICT1 IV, 205 Karjat-Ahilyanagar, 206 Karjat-Bhose, 207 Bus Coupler,
    208 TBC, 209 Karjat-Shirsuphal, 210 Karjat-Bhigwan, 211 Karjat-Jeur Ckt1, 212 Karjat-Jeur Ckt2
  - **Bus-1 / Bus-A:** 203, 205, 207, 209, 211
  - **Bus-2 / Bus-B:** 204, 206, 208, 210, 212
- **Relay families:** NR Electric PCS-9xx (sheets: Index | Setting Title | Value | Unit | Step | Range)
  and MiCOM P-series (Name | Address | Value; row 0 may be "Model Number").
- **Theme:** MSETCL navy/gold, Cinzel + Inter. **Must not be bright** — muted, low-glare,
  easy on the eyes for long sessions. Natural background wallpaper that **changes daily**.

## Architecture (settled)

- **Storage = Google Drive/Sheets** via Apps Script Web App (`code.gs`). Word "**Storage**" in all UI.
- Database in sheets: `_Bays`, `_Relays`, `_Settings`, `_Meta` inside "Karjat Protection Portal Data".
- **Overview = one SEPARATE spreadsheet per entity** (relay / bay / substation), tracked in
  `_OverviewRegistry`, filed under an "Overviews" folder. Each can have many real tabs.
- Portal is a thin client; **full build keeps embedded fallback data** so it always opens.
- `SS_ID` must be pasted into code.gs after first sync — biggest performance win.
- Future: merge into a larger LTS webpage. No login page at this stage.

## Requirements log

### Done in the current (incremental) build
- Immutable baseline v1 + draft/applied revisions, audit trail, correction ("typo") workflow
  with date/reason/user/attachment and reverse-chronological history.
- Upload settings file as a version at relay level; zip upload at bay/substation with
  filename→relay matching. Asks **older (v1.1-style)** vs **newer/revised (v2.1-style)**.
- **Delete any non-baseline version** (extra confirm if it's the applied one).
- Cards: merged main+custom, sections, pin-to-bay, corner "＋ Card" button, auto width,
  search anywhere across title/group/value/unit/desc/vm with meaning snippets in results.
- Hover-to-reveal meanings when meanings are hidden (title cell, value cell, cards).
- Users: register → admin approval, roles, view-only default, **viewer preview mode**.
- Freeform Overview mirroring the sheet **with formatting** (bg/font colours, bold/italic,
  size, alignment, wrap, merges, column widths), all tabs shown, drag-to-resize columns
  persisted back to the sheet, bulk Excel upload at all three levels.
- **"Database (auto)" first tab** in every overview spreadsheet — a filtered view of the master
  database for that entity (relay: Group/Setting/Value/Unit/Step/Range/meanings; bay: + Relay,
  Relay Type; substation: + Bay). Written by Apps Script, so user tabs in the SAME spreadsheet
  reference it with plain FILTER/VLOOKUP/QUERY — **no IMPORTRANGE, no auth prompts**.
  Rebuilt on overview creation and via "⟳ Sync database tab".
- Manuals page (official NR Electric / Schneider refs only), interruption history,
  Excel export, backup/restore, data-freshness badge, sidebar collapsed by default,
  bay page shows relays as clickable cards.

### Still outstanding (carry into the rewrite)
- **SLD rework** — separate clean 400 kV drawing with ICT symbols (capacity + names),
  spare ICT on Bus-2; separate 220 kV with Bus-A/Bus-B assignments. Ashish will supply a
  refined SLD with isolator + CB positions for tripping analysis.
  *Preferred hand-off format: a structured table (bay → isolators → bus association → CB),
  which drives tripping logic directly; a dimensioned PDF/image is the fallback.*
- **Formatted Excel exports** — light-grey **centred** headers, rich tables, wrapped text in
  description/text fields, auto row height to tallest cell, per-page remark
  "(specify name) some columns are user generated".
- **Drive-style file-manager page** — download/upload/open/delete/edit, with two systems:
  1. *User-generated*: formatted Excel with abstract + comparison details **fetched from the
     Overview sheets** — substation level (overview + zip of all bays' sheets), bay level
     (all-relay-data-with-abstract-comparisons + zip of relay files), relay level.
  2. *System-generated*: the originally uploaded files, auto-segregated (substation zip /
     bay zip / relay level).
- **Cards**: drag-resizable width (remembered), drag cards between sections.
- **Meanings data gap**: only 1 of 62 relays (Girawali-1 Main-1 Distance) has meanings.
  Bulk generation exists but is slow (~3,950 AI calls). Needs a better ingestion story.

## Known traps (learned the hard way)
- Apps Script **must be redeployed as a New Version**; saving alone keeps serving old code.
  Symptoms of stale deployment: `relayId required`, `blockId required`.
  Access must be **"Anyone"**, else Google returns a sign-in page instead of JSON.
- Claude's sandbox **cannot reach script.google.com** — all Drive logic must be verified by
  Node simulation, never live.
- localStorage ceiling ~5–10 MB; `simpleHash_` is not real security (documented, pending LTS auth).
- Overview data shape changed twice; always migrate old shapes defensively.
