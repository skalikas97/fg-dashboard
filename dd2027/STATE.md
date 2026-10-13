# STATE - AL-IDRISI drydock 2027 (dd2027)

Read first, update last. Keep about 2 KB. Pack: `pack/00`-`03b`. Last update: 13/10/2026.

## Goal
Phase 1: EMPTY PREDEFINED template, built 1:1 on the TS references (construction only, no job content). Deviations (D1-D8, A3 fixes, skill build-ins) OFF unless TS says yes. Phase 2 (content) later.

## Tab board
| File | Tabs | State |
|---|---|---|
| `out/AI_DD2027_Execution_Rev01.xlsx` | Daily SOW (master), Spares, Subcontractors-attendees, Planning | Built 13/10, waiting TS review |
| `out/AI_DD2027_Commercial_Rev00.xlsx` | Costfile, AXI, Scaffolding, 3rd party Planner, 3rd parties, SOW final, SOW initial (master), Budget file | Built 13/10, waiting TS review |
| Yard list | Content draft `AI_DD2027_Yard_list_Rev02.xlsx` (not in repo) | Phase 2 |

## Open decisions
Review of both files (TS removes tabs he does not want). Macro paste once per workbook (S1). Planning/Planner start dates are input cells. SOW D1-D8 only if wanted.

## Decisions log (TS-confirmed only)
- 29/09/2026 method optimise -> construct -> content, sign-off per stage.
- 13/10/2026 template fidelity first, deviations OFF by default (instead of the 29/09 redesign proposal D1-D8).
- 13/10/2026 build continues in Claude Code; repo folder `dd2027/`.
- 13/10/2026 T1 all reference tabs, TS removes (instead of a pre-selection). T2 two workbooks: X = execution/preparation (AA), C = commercial/during DD (SO) (instead of one). T3 categories = yard letters A-X (SO), AI extra jobs added later from AI files (instead of AA JDN groups). T4 15 Spares / 5 elsewhere / 10 Rev rows. T5 placeholders everywhere + one workbook-wide auto "last change" date shown on every tab (new; method S2). T6 35 days from one start-date cell. T7 Planner scroll bar reproduced as a real form control (instead of a plain input cell). T8 AA style governs both workbooks (instead of each tab its own). T9 SANTIAGO DOCKING colours D99694 / B9CDE5 (instead of COSTFILE F2AA84 / 83CBEB). T10 WL item = vessel work-list item raised by Master / Chief Engineer. Cost code [TS 13/10] "0080 (docking)-8038 (al idrisi) - XXX (category of cost)"; references write it 0080/8578/411 (Budget) and carrier 8579 (AXI) -> AI: 0080/8038/XXX, carrier 8038.

- 13/10/2026 Daily SOW: S1 last-change by macro, .xlsm, TS pastes once per workbook (instead of no auto date). S2 letters A-X on all execution tabs, keys A.1 / A.1.1 (instead of JDN groups). S3 header "WL item" (instead of AA "WO item"). S4 Dept D = Deck dept, E = Engine/Elec dept, T = Technical Superintendent. S5 title B2:E13, 10 Rev rows + Last change row 13, header row 15 (instead of B2:E8 / row 10).

- 13/10/2026 TS: build all remaining tabs at once (instead of one tab per turn). P1 Daily SOW print fit to one page wide (default taken). Commercial master tab = SOW initial (proposal). Changes flagged to TS: SOW header "Remarks" (ref "Remakrs"), Costfile P6 "Final", "Amount With discount" uses the discount cell (ref 0.95 / 0.965 fixed), AXI linked inside the workbook (ref external link), IFERROR on empty-template divisions, Planner WORK DAYS = NETWORKDAYS, Planning letter colours = legend colours, Planning start date must be a Monday.

## Inputs on hand
- 4 reference .xlsx in `reference/` and their extracted specs in `build/spec/` (13/10/2026, read only, local only - never pushed). Sheet names, freeze splits and package parts match pack 03 A0/A5/A7.

## Missing inputs
- `AI_DD2027_Yard_list_Rev02.xlsx`.
- Full JDN code tree (300, 600); docking plan B0184000; predecessor masterfile.

## Hard dates
ISM/ISPS + SMC/SSC 09/03/2027; Centre Tailshaft + CSL 134-138 10/03/2027; ILO 152 periodical 11/03/2027; handover due 25/03/2027 [HO]; BV renewals 26/03/2027 [BV].
