# STATE - AL-IDRISI drydock 2027 (dd2027)

Read first, update last. Keep about 2 KB. Pack: `pack/00`-`03b`. Last update: 13/10/2026.

## Goal
Phase 1: EMPTY PREDEFINED template, built 1:1 on the TS references (construction only, no job content). Deviations (D1-D8, A3 fixes, skill build-ins) OFF unless TS says yes. Phase 2 (content) later.

## Tab board
| Tab | Stage | Note |
|---|---|---|
| All tabs | Stage 1 open | Waiting T1-T10 |
| Yard list | Content draft | `AI_DD2027_Yard_list_Rev02.xlsx` (not in repo) |

## Open decisions (put 13/10/2026)
T1 tabs and source file; T2 one or two workbooks; T3 category lists; T4 buffer rows; T5 placeholders; T6 Planning grid start/days; T7 Planner scroll bar; T8 AA/SO style mix; T9 money-tab colour base; T10 Dept "T", "WL item", cost carrier. SOW D1-D8 (only if a deviation is wanted).

## Decisions log (TS-confirmed only)
- 29/09/2026 method optimise -> construct -> content, sign-off per stage.
- 13/10/2026 template fidelity first, deviations OFF by default (instead of the 29/09 redesign proposal D1-D8).
- 13/10/2026 build continues in Claude Code; repo folder `dd2027/`.

## Missing inputs
- The 4 reference .xlsx (read only) -> `reference/`: AA_DAILY_SOW Afloat Repairs final, SO - SOW, SANTIAGO DOCKING, SO - COSTFILE - AXI SPLITUP. Needed for render comparison at construct.
- `AI_DD2027_Yard_list_Rev02.xlsx`.
- Full JDN code tree (300, 600); docking plan B0184000; predecessor masterfile.

## Hard dates
ISM/ISPS + SMC/SSC 09/03/2027; Centre Tailshaft + CSL 134-138 10/03/2027; ILO 152 periodical 11/03/2027; handover due 25/03/2027 [HO]; BV renewals 26/03/2027 [BV].
