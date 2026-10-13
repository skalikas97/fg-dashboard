# STATE - AL-IDRISI drydock 2027 (dd2027)

Read first, update last. Keep about 2 KB. Pack: `pack/00`-`03b`. Last update: 13/10/2026.

## Goal
Phase 1: EMPTY PREDEFINED template, built 1:1 on the TS references (construction only, no job content). Deviations (D1-D8, A3 fixes, skill build-ins) OFF unless TS says yes. Phase 2 (content) later.

## Tab board
| Tab | Stage | Note |
|---|---|---|
| X1 Daily SOW | Stage 1 proposed 13/10 | Waiting S1-S5 |
| Other tabs | Stage 1 open | |
| Yard list | Content draft | `AI_DD2027_Yard_list_Rev02.xlsx` (not in repo) |

## Open decisions
Daily SOW S1-S5 (13/10). Planning start date + days. SOW D1-D8 only if wanted.

## Decisions log (TS-confirmed only)
- 29/09/2026 method optimise -> construct -> content, sign-off per stage.
- 13/10/2026 template fidelity first, deviations OFF by default (instead of the 29/09 redesign proposal D1-D8).
- 13/10/2026 build continues in Claude Code; repo folder `dd2027/`.
- 13/10/2026 T1 all reference tabs, TS removes (instead of a pre-selection). T2 two workbooks: X = execution/preparation (AA), C = commercial/during DD (SO) (instead of one). T3 categories = yard letters A-X (SO), AI extra jobs added later from AI files (instead of AA JDN groups). T4 15 Spares / 5 elsewhere / 10 Rev rows. T5 placeholders everywhere + one workbook-wide auto "last change" date shown on every tab (new; method S2). T6 35 days from one start-date cell. T7 Planner scroll bar reproduced as a real form control (instead of a plain input cell). T8 AA style governs both workbooks (instead of each tab its own). T9 SANTIAGO DOCKING colours D99694 / B9CDE5 (instead of COSTFILE F2AA84 / 83CBEB). T10 WL item = vessel work-list item raised by Master / Chief Engineer. Cost code [TS 13/10] "0080 (docking)-8038 (al idrisi) - XXX (category of cost)"; references write it 0080/8578/411 (Budget) and carrier 8579 (AXI) -> AI: 0080/8038/XXX, carrier 8038.

## Inputs on hand
- 4 reference .xlsx in `reference/` (13/10/2026, read only, local only - not pushed). Sheet names, freeze splits and package parts match pack 03 A0/A5/A7.

## Missing inputs
- `AI_DD2027_Yard_list_Rev02.xlsx`.
- Full JDN code tree (300, 600); docking plan B0184000; predecessor masterfile.

## Hard dates
ISM/ISPS + SMC/SSC 09/03/2027; Centre Tailshaft + CSL 134-138 10/03/2027; ILO 152 periodical 11/03/2027; handover due 25/03/2027 [HO]; BV renewals 26/03/2027 [BV].
